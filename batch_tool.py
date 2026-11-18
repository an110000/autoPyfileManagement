"""本地收件整理工具。运行 python batch_tool.py --help 查看用法。"""

import argparse
import csv
import json
import os
import re
import shutil
import stat
import sys
import tempfile
import time
import zipfile
from collections import Counter
from datetime import datetime
from fnmatch import fnmatchcase
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font, PatternFill
from PIL import Image, ImageFilter, ImageOps
import xlrd

from sheet_headers import MAX_HEADER_DEPTH, read_header, worksheet_merges


SOURCE_COLUMNS = ["__来源文件", "__工作表", "__原行号"]
EXCEL_INPUTS = {".xls", ".xlsx"}
TABLES = EXCEL_INPUTS | {".csv"}
IMAGES = {".jpg", ".jpeg", ".png", ".bmp", ".webp", ".tif", ".tiff", ".gif"}
DOCUMENTS = {".doc", ".docx", ".pdf", ".ppt", ".pptx", ".txt", ".md", ".rtf"}
EXCEL_ROWS, EXCEL_COLUMNS = 1_048_576, 16_384
INVALID_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]")


def is_directory_link(path):
    """识别符号链接和 Windows 目录联接，供递归扫描避开循环。"""
    return path.is_symlink() or (
        os.name == "nt" and path.lstat().st_reparse_tag == stat.IO_REPARSE_TAG_MOUNT_POINT
    )


def input_files(root, output):
    """按稳定顺序产生输入文件路径，跳过链接和本批次输出目录。"""
    def scan_error(error):
        """将目录访问失败交给调用方，避免默默遗漏文件。"""
        raise error

    for folder, dirs, files in os.walk(root, onerror=scan_error):
        base = Path(folder)
        # 在递归前剪枝，输出放在输入之内时，重跑也不会读到历史结果。
        dirs[:] = sorted(
            d for d in dirs if base / d != output and not is_directory_link(base / d)
        )
        for name in sorted(files):
            path = base / name
            if not path.is_symlink() and path.is_file():
                yield path


def clean_value(value):
    """清除文本首尾空白并校验 Excel 限制；数字、日期保持原类型。"""
    if not isinstance(value, str):
        return value
    value = value.strip()
    if len(value) > 32_767 or INVALID_XML.search(value):
        raise ValueError("文本超长或包含 Excel 无法保存的字符")
    return value or None


def numbered_rows(rows, sheet):
    """逐行清理实际值，去掉右侧空格式单元格，保留原始行号及中间空位。"""
    for number, row in enumerate(rows, 1):
        try:
            values = [clean_value(v) for v in row]
            while values and values[-1] is None:
                values.pop()
        except ValueError as error:
            raise ValueError(f"{sheet or 'CSV'} 第 {number} 行：{error}") from error
        yield number, values


def table_frame(rows, source, sheet, merges=(), header_row=0, header_depth=1):
    """从物理行构造数据框；标题、合并表头和后续新增列均按各自的位置处理。"""
    names, span, data = read_header(numbered_rows(rows, sheet), merges, header_row, header_depth)
    named = [name for name in names if name is not None]
    used = set(named)
    if len(used) != len(named) or used.intersection(SOURCE_COLUMNS):
        raise ValueError(f"{sheet or 'CSV'} 第 {span[0]} 行：列名重复或占用了来源列名，请检查表头")
    generated = []

    def unnamed(position):
        """按原始列号给无标题列命名，同时避开用户实际使用的同名业务列。"""
        name = f"未命名列_{position}"
        while name in used:
            name += "_"
        used.add(name)
        generated.append(name)
        return name

    headers = [name or unnamed(i) for i, name in enumerate(names, 1)]
    records = []
    for number, values in data:
        if not values:
            continue
        headers.extend(unnamed(i) for i in range(len(headers) + 1, len(values) + 1))
        values.extend([None] * (len(headers) - len(values)))
        records.append([source, sheet, number, *values])
    frame = pd.DataFrame(records, columns=SOURCE_COLUMNS + headers, dtype=object)
    frame.attrs.update(sheet=sheet, header_span=span, unnamed=generated)
    return frame


def xls_values(sheet, datemode):
    """逐行还原 BIFF 单元格；日期使用工作簿的日期系统，公式读取其缓存结果。"""
    def value(cell):
        """将带类型的旧版单元格转换为与 XLSX 读取器一致的 Python 标量。"""
        if cell.ctype == xlrd.XL_CELL_DATE:
            date = xlrd.xldate_as_datetime(cell.value, datemode)
            return date.time() if 0 <= cell.value < 1 else date
        if cell.ctype == xlrd.XL_CELL_BOOLEAN:
            return bool(cell.value)
        if cell.ctype == xlrd.XL_CELL_ERROR:
            return xlrd.error_text_from_code.get(cell.value, "#ERROR!")
        if cell.ctype == xlrd.XL_CELL_NUMBER and cell.value.is_integer():
            return int(cell.value)
        return cell.value

    for i in range(sheet.nrows):
        yield (value(cell) for cell in sheet.row(i))


def matches_sheet(title, pattern):
    """空规则选择全部，否则用不区分大小写的通配符匹配完整工作表名。"""
    return not pattern or fnmatchcase(title.casefold(), pattern.casefold())


def sheet_names(path):
    """仅读取工作簿目录，返回可处理的数据工作表名称（按原顺序，排除图表页）。"""
    if path.suffix.lower() == ".xls":
        with xlrd.open_workbook(str(path), on_demand=True) as book:
            return book.sheet_names()
    book = load_workbook(path, read_only=True)
    try:
        return [sheet.title for sheet in book.worksheets]
    finally:
        book.close()


def read_tables(path, source, header_row=0, header_depth=1, sheet_pattern=""):
    """按扩展名解码，多工作表共用同一套表头识别；整本成功后才交付数据。"""
    if path.suffix.lower() == ".csv":
        if sheet_pattern:
            return []
        def read_csv(encoding):
            """以指定编码按 CSV 记录读取，支持带引号的逗号和换行。"""
            with path.open(encoding=encoding, newline="") as stream:
                return table_frame(csv.reader(stream, strict=True), source, "", (), header_row, header_depth)

        try:
            frame = read_csv("utf-8-sig")
        except UnicodeDecodeError:
            frame = read_csv("gb18030")
        return [frame]

    if path.suffix.lower() == ".xls":
        frames = []
        with xlrd.open_workbook(str(path), formatting_info=True, on_demand=True, ragged_rows=True) as book:
            for i, title in enumerate(book.sheet_names()):
                if not matches_sheet(title, sheet_pattern):
                    continue
                sheet = book.sheet_by_index(i)
                # xlrd 使用从零起、右端不包含的区间；在此统一为 Excel 坐标。
                merges = [(r + 1, end, col + 1, right) for r, end, col, right in sheet.merged_cells]
                frames.append(table_frame(xls_values(sheet, book.datemode), source, sheet.name,
                                          merges, header_row, header_depth))
                frames[-1].attrs["sheet_number"] = i + 1
                book.unload_sheet(i)
        return frames

    # XLSX 保留公式原文，输出时作为文字写入。
    book = load_workbook(path, read_only=True, data_only=False)
    try:
        frames = []
        for i, sheet in enumerate(book.worksheets, 1):
            if not matches_sheet(sheet.title, sheet_pattern):
                continue
            sheet.reset_dimensions()
            limit = header_row + MAX_HEADER_DEPTH if header_row else EXCEL_ROWS
            merges = worksheet_merges(sheet, limit)
            frames.append(table_frame(sheet.values, source, sheet.title, merges, header_row, header_depth))
            frames[-1].attrs["sheet_number"] = i
        return frames
    finally:
        book.close()


def excel_cell(sheet, value):
    """构造只写单元格：缺失值留空，等号开头的字符串也按文字写。"""
    if pd.isna(value):
        return None
    cell = WriteOnlyCell(sheet, value=value)
    if isinstance(value, str):
        cell.data_type = "s"  # 包括以等号开头的编号、正文和公式原文。
    return cell


def write_workbook(frames, target):
    """按首次出现的列名对齐数据块，写出单页 Excel 并返回数据行数。"""
    columns = list(dict.fromkeys(c for frame in frames for c in frame.columns)) if frames else SOURCE_COLUMNS
    count = sum(len(frame) for frame in frames)
    if count + 1 > EXCEL_ROWS or len(columns) > EXCEL_COLUMNS:
        raise ValueError("合并结果超出单张 Excel 工作表限制，请将输入分批处理")
    book = Workbook(write_only=True)
    sheet = book.create_sheet("合并数据")
    sheet.freeze_panes = "D2" if columns[:3] == SOURCE_COLUMNS else "A2"
    header = [excel_cell(sheet, column) for column in columns]
    for cell in header:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="245573")
    sheet.append(header)
    # 每次只对齐一个文件/工作表的数据块，避免再复制一份完整合并表。
    for frame in frames:
        aligned = frame.reindex(columns=columns)
        for row in aligned.itertuples(index=False, name=None):
            sheet.append([excel_cell(sheet, v) for v in row])
    book.save(target)
    return count


def convert_image(path, target, sharpen, quality=90):
    """按目标后缀转换单帧图片；先校正方向，按需锐化，再处理透明度。"""
    with Image.open(path) as image:
        if getattr(image, "n_frames", 1) != 1:
            raise ValueError("多帧图片需先选定帧，已保留原文件")
        oriented = ImageOps.exif_transpose(image)
        rgba = oriented.convert("RGBA")
        alpha = rgba.getchannel("A")
        rgb = rgba.convert("RGB")
        if sharpen:
            rgb = rgb.filter(ImageFilter.UnsharpMask(radius=1.2, percent=110, threshold=3))
        if target.suffix in {".jpg", ".bmp"}:
            canvas = Image.new("RGB", rgb.size, "white")
            canvas.paste(rgb, mask=alpha)
            if target.suffix == ".jpg":
                canvas.save(target, quality=quality, subsampling=0)
            else:
                canvas.save(target)
        else:
            rgb.putalpha(alpha)
            rgb.save(target, **({"quality": quality} if target.suffix == ".webp" else {}))


def process_files(root, output, stage, image_format, sharpen, progress=None, selected=None):
    """逐文件完成一键归档，将单文件失败写入清单，返回计数及文件记录。"""
    frames, entries = [], []
    for number, path in enumerate(selected if selected is not None else input_files(root, output), 1):
        source = path.relative_to(root).as_posix()
        entry = {"source": source, "status": "ok"}
        entries.append(entry)
        suffix = path.suffix.lower()
        target = None
        try:
            if path.name.startswith("~$"):
                entry.update(status="skipped", reason="Office 临时文件")
            elif suffix in TABLES:
                tables = read_tables(path, source)
                frames.extend(tables)  # 整个文件读完才纳入汇总。
                entry.update(output="合并数据.xlsx", rows=sum(len(t) for t in tables))
            elif suffix in IMAGES or suffix in DOCUMENTS:
                stem = f"{number:04d}_{path.stem[:48]}"
                if suffix in IMAGES:
                    ext = ".jpg" if image_format == "jpeg" else ".png"
                    target = stage / "图片" / (stem + ext)
                else:
                    target = stage / "文档" / (stem + suffix)
                target.parent.mkdir(exist_ok=True)
                if suffix in IMAGES:
                    convert_image(path, target, sharpen)
                else:
                    shutil.copy2(path, target)
                entry["output"] = target.relative_to(stage).as_posix()
            else:
                entry.update(status="skipped", reason="未配置此后缀")
        except Exception as error:
            # 个别文件损坏时继续处理其余输入。删除本次未写完的副本。
            if target is not None:
                target.unlink(missing_ok=True)
            entry.update(status="error", reason=f"{type(error).__name__}: {error}")
        if progress is not None:
            progress("处理文件", entry)
    if progress is not None:
        progress("写入汇总表", None)
    rows = write_workbook(frames, stage / "合并数据.xlsx")
    counts = Counter(e["status"] for e in entries)
    return {
        "summary": {**{s: counts[s] for s in ("ok", "error", "skipped")}, "rows": rows},
        "options": {"image_format": image_format, "sharpen": sharpen},
        "files": entries,
    }


def publish(root, output, process, progress=None):
    """在临时目录执行 process，写清单和 ZIP，整批就绪后再发布。"""
    root, output = root.absolute().resolve(), output.absolute().resolve()
    if not root.is_dir():
        raise ValueError(f"输入目录不存在：{root}")
    if root == output:
        raise ValueError("输入和输出目录相同，请为结果指定单独的目录")
    output.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(dir=output, prefix=".batch-") as temp:
        stage = Path(temp) / "result"
        stage.mkdir()
        report = process(stage)
        if progress is not None:
            progress("生成清单与压缩包", None)
        (stage / "处理清单.json").write_text(
            json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        members = sorted(p for p in stage.rglob("*") if p.is_file())
        with zipfile.ZipFile(stage / "整理结果.zip", "w", zipfile.ZIP_DEFLATED) as archive:
            for path in members:
                archive.write(path, path.relative_to(stage).as_posix())
        # 临时目录和成品位于同一文件系统，整批写完后才发布。
        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        batch_id = Path(temp).name.removeprefix(".batch-")
        result = output / f"批次-{timestamp}-{batch_id}"
        # Windows 索引器偶尔会短暂占用刚写完的文件。只重试同一个发布动作；
        # 若目标已出现或权限错误持续存在，保留错误，不能覆盖别的批次。
        for delay in (0.08, 0.2, 0.5, 1.0, None):
            try:
                stage.rename(result)
                break
            except PermissionError as error:
                if delay is None or getattr(error, "winerror", 0) not in (5, 32) or result.exists():
                    raise
                time.sleep(delay)
    return result, report


def run(root, output, image_format, sharpen, progress=None, selected=None):
    """命令行与桌面共用的一键归档入口，返回批次路径和数量汇总。"""
    root, output = root.absolute().resolve(), output.absolute().resolve()

    def archive(stage):
        """在发布器提供的临时位置写入表格、图片和文档副本。"""
        return process_files(root, output, stage, image_format, sharpen, progress, selected)

    result, report = publish(root, output, archive, progress)
    return result, report["summary"]


def main():
    """解析命令行参数；成功、部分文件失败、整批失败分别返回 0、1、2。"""
    parser = argparse.ArgumentParser(description="按列名汇总表格、转换图片、复制文档并打包。")
    parser.add_argument("input", type=Path, help="输入文件夹，包含其子目录")
    parser.add_argument("output", type=Path, help="输出根目录，每次运行创建新批次")
    parser.add_argument("--format", choices=["jpeg", "png"], default="jpeg", help="图片输出格式（默认 jpeg）")
    parser.add_argument("--sharpen", action="store_true", help="启用轻度锐化")
    args = parser.parse_args()
    try:
        result, summary = run(args.input, args.output, args.format, args.sharpen)
    except Exception as error:
        print(f"批次未完成：{error}", file=sys.stderr)
        return 2
    print(f"结果目录：{result}")
    print("成功 {ok} 个，失败 {error} 个，跳过 {skipped} 个，合并 {rows} 行。".format(**summary))
    return 1 if summary["error"] else 0


if __name__ == "__main__":
    sys.exit(main())
