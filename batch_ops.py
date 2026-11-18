"""桌面批处理规则。预览和执行共用本模块；stage 为 None 表示只计算结果。"""

import codecs
import hashlib
import os
import re
import shutil
import string
from collections import Counter, defaultdict
from datetime import datetime
from decimal import Decimal, InvalidOperation
from itertools import islice
from pathlib import Path

import pandas as pd

from batch_tool import DOCUMENTS, EXCEL_INPUTS, IMAGES, SOURCE_COLUMNS, TABLES, archive_batch, convert_image, input_files, publish
from batch_tool import matches_sheet, read_tables, sheet_names, write_workbook


TEXT = {".txt", ".md", ".log", ".csv", ".tsv", ".json", ".xml", ".html", ".ini", ".yaml", ".yml"}
IMAGE_OUTPUTS = {"JPEG": ".jpg", "PNG": ".png", "WEBP": ".webp", "BMP": ".bmp", "TIFF": ".tiff"}
INVENTORY_COLUMNS = ["相对路径", "文件名", "类别", "扩展名", "大小（字节）", "修改时间", "创建时间"]
FILE_CATEGORIES = {
    "演示文稿": {".ppt", ".pptx", ".pps", ".ppsx"},
    "视频": {".mp4", ".avi", ".mov", ".wmv", ".flv", ".mkv", ".webm"},
    "音频": {".mp3", ".wav", ".flac", ".aac", ".m4a", ".ogg"},
    "压缩包": {".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz"},
    "程序": {".exe", ".msi", ".bat", ".cmd", ".sh", ".py", ".js", ".html", ".css"},
    "字体": {".ttf", ".otf", ".woff", ".woff2"},
    "数据": {".json", ".xml", ".sql", ".db", ".sqlite", ".yaml", ".yml"},
}
INVALID_NAME = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def valid_name(name):
    """验证一个 Windows 文件名分量；拒绝路径、设备名和会被系统截断的尾字符。"""
    if (not name or name in {".", ".."} or INVALID_NAME.search(name)
            or name.endswith((" ", ".")) or Path(name).is_reserved()
            or len(name.encode("utf-16-le")) > 510):
        raise ValueError(f"不能作为文件名：{name!r}")
    return name


def slug(value):
    """将工作表名或分组值变成便于阅读的文件名片段；序号负责区分同名值。"""
    text = "" if pd.isna(value) else str(value)
    return INVALID_NAME.sub("_", text).strip(" .")[:60] or "空值"


def compile_match(options):
    """根据普通文本/正则和大小写选项编译规则；预览前即报告非法表达式。"""
    search = options["find"]
    if not search:
        raise ValueError("请填写查找内容。")
    regex = options.get("regex") or options.get("rule") == "正则改名"
    pattern = re.compile(search if regex else re.escape(search),
                         0 if options.get("case", True) else re.IGNORECASE)
    if regex:
        # 在批次启动前检查替换串中的组号，即使没有任何文件匹配也能报告配置错误。
        pattern.sub(options["replacement"], "")
    return pattern


def replace_text(path, options, pattern, samples=None):
    """返回新字节、次数和说明；可同时记录前五次替换的原始行号及前后文字。"""
    raw = path.read_bytes()
    encoding = options["encoding"]
    if encoding == "自动":
        if raw.startswith((codecs.BOM_UTF32_LE, codecs.BOM_UTF32_BE)):
            raise ValueError("暂不支持 UTF-32 文本")
        if raw.startswith(codecs.BOM_UTF16_LE):
            encoding = "utf-16-le"
        elif raw.startswith(codecs.BOM_UTF16_BE):
            encoding = "utf-16-be"
        elif raw.startswith(codecs.BOM_UTF8):
            encoding = "utf-8-sig"
        else:
            try:
                raw.decode("utf-8")
                encoding = "utf-8"
            except UnicodeDecodeError:
                encoding = "gb18030"
    # 单独保存 UTF-16 BOM；带 ^ 的正则应从真正的第一个字符开始匹配。
    bom = b""
    if encoding in ("utf-16-le", "utf-16-be"):
        marker = codecs.BOM_UTF16_LE if encoding == "utf-16-le" else codecs.BOM_UTF16_BE
        if raw.startswith(marker):
            bom, raw = marker, raw[len(marker):]
    text = raw.decode(encoding)
    if "\0" in text:
        raise ValueError("内容包含空字符，请检查文本编码")
    first = pattern.search(text)
    if first is None:
        return bom + raw, 0, "未匹配"
    replacement = options["replacement"]
    # 普通模式允许反斜线原样出现；正则模式支持 \1 等捕获组引用。
    replace = replacement if options["regex"] else lambda _match: replacement
    changed, count = pattern.subn(replace, text)
    example = first.expand(replacement) if options["regex"] else replacement
    if samples is not None:
        line = 1
        breaks = re.finditer(r"\r\n|[\r\n]", text)
        next_break = next(breaks, None)
        for match in islice(pattern.finditer(text), 5):
            # 计算原文中的行号，CRLF 算一次换行；多行匹配记其起始行。
            while next_break is not None and next_break.end() <= match.start():
                line += 1
                next_break = next(breaks, None)
            after = match.expand(replacement) if options["regex"] else replacement
            samples.append([str(line), match.group()[:160], after[:160]])
    detail = f"{count} 处 · {first.group()[:40]!r} → {example[:40]!r} · {encoding}"
    return bom + changed.encode(encoding), count, detail


def rename_to(path, number, options, pattern=None):
    """在原扩展名前应用模板、查找替换或前后缀；返回经校验的新文件名。"""
    mode = options["rule"]
    if mode == "模板序号":
        template = options["template"]
        # 只开放文件名、序号及日期，避免模板访问 Python 对象属性。
        for _, field, fmt, conversion in string.Formatter().parse(template):
            if field is not None and (field not in {"name", "n", "date"} or conversion
                                      or (fmt and not (field == "n" and re.fullmatch(r"0?[1-9]d", fmt)))):
                raise ValueError("模板仅支持 {name}、{date}、{n} 或 {n:03d} 等序号格式")
        stem = template.format(name=path.stem, n=number, date=datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y%m%d"))
    elif mode == "正则改名":
        pattern = pattern or compile_match(options)
        stem = pattern.sub(options["replacement"], path.stem)
    elif mode == "查找替换":
        if not options["find"]:
            raise ValueError("请输入文件名中要替换的文字")
        stem = path.stem.replace(options["find"], options["replacement"])
    else:
        stem = options["prefix"] + path.stem + options["suffix"]
    if not stem:
        raise ValueError("生成的文件主名为空")
    return valid_name(stem + path.suffix)


def digest(path):
    """分块计算 SHA-256；大文件无需整体载入内存。"""
    result = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            result.update(block)
    return result.hexdigest()


def move_name(source, target):
    """在同一目录内改名且不覆盖现有文件。Windows 使用系统的拒绝覆盖语义。"""
    if source.parent != target.parent or target.exists() or target.is_symlink():
        raise FileExistsError(f"目标名称已占用或不在原目录：{target}")
    if os.name == "nt":
        source.rename(target)
    else:
        # POSIX 的 rename 默认会覆盖。先独占建立硬链接，再移除旧名字。
        os.link(source, target)
        try:
            source.unlink()
        except OSError:
            target.unlink()
            raise


def fingerprint(path):
    """记录文件身份、大小和修改时间，用于撤销前核对文件仍是原件。"""
    info = path.stat()
    return [info.st_dev, info.st_ino, info.st_size, info.st_mtime_ns]


def undo_renames(root, entries):
    """逆序恢复本程序记录的原位改名。跳过已经变动的文件，返回逐项失败原因。"""
    errors = []
    for entry in reversed(entries):
        if "identity" not in entry:
            continue
        original, current = root / entry["source"], root / entry["output"]
        try:
            if current.is_symlink() or fingerprint(current) != entry["identity"]:
                raise ValueError("文件已发生变化")
            move_name(current, original)
        except (OSError, ValueError) as error:
            errors.append(f"{entry['source']}：{error}")
        else:
            entry.pop("identity")
    return errors


def rename_originals(root, output, options, selected, preview, progress, pattern=None):
    """先登记完整变更计划，再原位改名。若批次清单无法发布，按记录回滚。"""
    changes = []

    def process(stage):
        """构造源名与目标名的映射，任何名称冲突均记录失败，不覆盖原件。"""
        entries = []
        reserved = set()
        extensions = {s.lower().lstrip(".") for s in columns(options.get("extensions", ""))}
        seq = options.get("start", 1)
        paths = list(selected if selected is not None else input_files(root, output))
        for path in paths:
            item = {"source": path.relative_to(root).as_posix(), "status": "ok"}
            entries.append(item)
            try:
                if path.name.startswith("~$") or (extensions and path.suffix.lower().lstrip(".") not in extensions):
                    item.update(status="skipped", reason="不符合文件筛选条件")
                else:
                    target = path.with_name(rename_to(path, seq, options, pattern))
                    seq += 1
                    if path.name == target.name:
                        item.update(status="skipped", reason="名称没有变化")
                    elif target in reserved or target.exists() or target.is_symlink():
                        raise FileExistsError(f"名称已占用：{target.name}")
                    else:
                        reserved.add(target)
                        item.update(output=target.relative_to(root).as_posix(), detail="原位改名")
            except (OSError, ValueError) as error:
                item.update(status="error", reason=str(error))
        for item in entries:
            if stage and item["status"] == "ok":
                src, dest = root / item["source"], root / item["output"]
                try:
                    identity = fingerprint(src)
                    move_name(src, dest)
                except OSError as error:
                    item.update(status="error", reason=str(error))
                    item.pop("detail", None)
                    item.pop("output", None)
                else:
                    item["identity"] = identity
                    changes.append(item)
            if progress:
                progress("预览改名" if preview else "原位改名", item)
        counts = Counter(e["status"] for e in entries)
        summary = {k: counts[k] for k in ("ok", "skipped", "error")}
        return {"files": entries, "options": options, "summary": dict(summary, rows=0)}

    if preview:
        return None, process(None)
    try:
        return publish(root, output, process, progress)
    except Exception as error:
        faults = undo_renames(root, changes)
        if faults:
            raise RuntimeError(f"批次未发布，部分名字未还原：{'；'.join(faults)}") from error
        raise


def category(path):
    """按后缀给出分类子目录，不能识别的文件放入“其他”。"""
    ext = path.suffix.lower()
    if ext in TABLES or ext == ".tsv":
        return "表格"
    if ext in IMAGES or ext == ".svg":
        return "图片"
    for group, suffixes in FILE_CATEGORIES.items():
        if ext in suffixes:
            return group
    if ext in DOCUMENTS or ext in TEXT:
        return "文档"
    return "其他"


def inventory_row(path, relative):
    """读取一次 stat，返回可导出的目录记录；系统未提供创建时间时留空。"""
    info = path.stat()
    created = getattr(info, "st_birthtime", info.st_ctime if os.name == "nt" else None)
    modified = datetime.fromtimestamp(info.st_mtime).isoformat(sep=" ", timespec="seconds")
    birth = datetime.fromtimestamp(created).isoformat(sep=" ", timespec="seconds") if created is not None else ""
    return [relative.as_posix(), path.name, category(path), path.suffix.lower(), info.st_size, modified, birth]


class Outputs:
    """记录本批次的相对路径。预览和正式运行使用同一套同名避让规则。"""

    def __init__(self, stage):
        """stage 是临时发布目录；为 None 时仅分配名称，不建立文件夹。"""
        self.stage = stage
        self.names = set()

    def allocate(self, relative):
        """为输出预留文件名，重名追加 (2)、(3)；返回相对名与实际写入路径。"""
        candidate = relative
        index = 1
        while candidate.as_posix().casefold() in self.names:
            index += 1
            candidate = relative.with_name(f"{relative.stem} ({index}){relative.suffix}")
        self.names.add(candidate.as_posix().casefold())
        target = self.stage / candidate if self.stage else None
        if target:
            target.parent.mkdir(parents=True, exist_ok=True)
        return candidate.as_posix(), target


def columns(text):
    """解析用英文或中文逗号分隔的列名，去掉两侧空格，保留顺序。"""
    return list(dict.fromkeys(c.strip() for c in re.split("[,，]", text) if c.strip()))


def table_product(frame, name):
    """将产物的前五行映射到窗口样本；文件写入和计数始终使用完整数据。"""
    business = [col for col in frame.columns if col not in SOURCE_COLUMNS][:24]
    rows = [["" if pd.isna(v) else str(v)[:160] for v in row]
            for row in frame[business].head(5).itertuples(index=False, name=None)]
    return {"output": name, "rows": len(frame), "sample_columns": business, "sample_rows": rows}


def filter_rows(data, field, options):
    """在合并后的指定列上筛选；缺失与空串同为空值，旧规则默认精确比较。"""
    series = data[field].astype("string").fillna("")
    value = options.get("filter_value", "")
    operator = options.get("filter_op", "精确等于")
    if operator == "精确等于":
        mask = series == value
    elif operator == "包含文字":
        mask = series.str.contains(value, regex=False)
    elif operator == "正则匹配":
        try:
            pattern = re.compile(value)
        except re.error as error:
            raise ValueError(f"筛选正则无效：{error}") from error
        mask = series.map(lambda text: pattern.search(text) is not None)
    elif operator in {"为空", "非空"}:
        mask = (series == "") if operator == "为空" else (series != "")
    else:
        raise ValueError(f"未知筛选条件：{operator}")
    return data.loc[mask.astype(bool)]


def describe_tables(tables):
    """列出各工作表的实际表头行与补齐的列名，供文件明细及导出清单核对。"""
    descriptions = []
    for table in tables:
        title = table.attrs["sheet"] or "CSV"
        span = table.attrs["header_span"]
        if span is None:
            descriptions.append(f"{title} 空表")
            continue
        first, last = span
        location = str(first) if first == last else f"{first}–{last}"
        note = f"{title} 表头第 {location} 行"
        if table.attrs["unnamed"]:
            note += "，补列 " + "、".join(table.attrs["unnamed"])
        descriptions.append(note)
    return "；".join(descriptions)


def convert_tables(tables, relative, destination):
    """转换单份输入的全部选中工作表。仅在整份成功后返回样本，避免残留部分结果。"""
    products = []
    for table in tables:
        frame = table.drop(columns=SOURCE_COLUMNS)
        if relative.suffix.lower() == ".csv":
            name, target = destination(Path("表格") / relative.with_suffix(".xlsx"))
            if target:
                write_workbook([frame], target)
        else:
            title = slug(table.attrs["sheet"])
            number = table.attrs["sheet_number"]
            name, target = destination(Path("表格") / relative / f"{number:03d}_{title}.csv")
            if target:
                frame.to_csv(target, index=False, encoding="utf-8-sig")
        products.append(table_product(frame, name))
    return products


def column_profile(data, selected):
    """按筛选后的数据统计各列。计频按文本比较，众数并列时保留先出现的值。"""
    headings = ["列名", "总行数", "非空数", "空值数", "不同非空值数", "最常见值", "出现次数"]
    rows = []
    for name in selected or [c for c in data.columns if c not in SOURCE_COLUMNS]:
        values = data[name].dropna().astype(str)
        frequencies = Counter(v for v in values if v != "")
        filled = sum(frequencies.values())
        common, count = frequencies.most_common(1)[0] if frequencies else ("", 0)
        rows.append([name, len(data), filled, len(data) - filled, len(frequencies), common, count])
    return pd.DataFrame(rows, columns=headings)


def transform_tables(frames, options, outputs):
    """全批次按名对齐后筛选、去重、分组或拆分，返回产物列表及输出行数。"""
    if not frames:
        return [], 0
    data = pd.concat(frames, ignore_index=True, sort=False)
    action = options["rule"]
    keys = columns(options.get("keys", ""))
    keep = columns(options.get("columns", ""))
    field = options.get("filter_column", "").strip()
    amount = options.get("amount", "").strip()
    required = [*keep, *([field] if field else [])]
    if action in {"按键去重", "分组汇总", "按值拆分"}:
        if not keys:
            raise ValueError("请填写分组列或去重键")
        required.extend(keys)
    if action == "分组汇总":
        if not amount or amount in keys:
            raise ValueError("数值列不能为空，也不能与分组列相同")
        required.append(amount)
    missing = [k for k in required if k not in data.columns]
    if missing:
        raise ValueError(f"找不到列 {', '.join(missing)}；可用列：{', '.join(data.columns)}")
    if field:
        data = filter_rows(data, field, options)
    if action == "列质量报告":
        data = column_profile(data, keep)
    if action == "按键去重":
        data = data.drop_duplicates(subset=keys, keep="first")
    if action == "分组汇总":
        # 金额用十进制解析。无法识别的值报告来源行，避免静默改变总额。
        numbers = []
        for _, record in data.iterrows():
            value = record[amount]
            try:
                number = Decimal(0) if pd.isna(value) else Decimal(str(value))
                if not number.is_finite():
                    raise InvalidOperation
                numbers.append(number)
            except InvalidOperation as error:
                raise ValueError(f"{record[SOURCE_COLUMNS[0]]} 第 {record[SOURCE_COLUMNS[2]]} 行的 {amount} 不是有效数值：{value}") from error
        if "记录数" in keys or "合计" in keys:
            raise ValueError("分组列与输出列“记录数”或“合计”重名")
        data = data.copy()
        data[amount] = numbers
        grouped = data.groupby(keys, dropna=False, sort=False)[amount]
        data = grouped.agg(记录数="size", 合计="sum").reset_index()
    products = []

    def emit(frame, name):
        """按输出列顺序写一个工作簿，同时给预览提供少量真实样例。"""
        if keep and action not in {"分组汇总", "列质量报告"}:
            frame = frame.loc[:, list(dict.fromkeys(SOURCE_COLUMNS + keep))]
        rel, target = outputs.allocate(Path("表格") / name)
        if target:
            write_workbook([frame], target)
        products.append(table_product(frame, rel))

    if action == "按值拆分":
        group_key = keys[0] if len(keys) == 1 else keys
        for i, (value, frame) in enumerate(data.groupby(group_key, dropna=False, sort=False), 1):
            emit(frame, f"{i:04d}_{slug(str(value) if isinstance(value, tuple) else value)}.xlsx")
    else:
        emit(data, action + ".xlsx")
    return products, sum(p["rows"] for p in products)


def duplicate_report(root, paths, outputs, options, progress):
    """先按字节数分组；只读取同大小的候选文件，以完整 SHA-256 确认重复。"""
    entries = []
    sizes = defaultdict(list)
    extensions = {v.lower().lstrip(".") for v in columns(options.get("extensions", ""))}
    if progress:
        progress("按大小筛选候选文件", None)
    for path in paths:
        item = {"source": path.relative_to(root).as_posix(), "status": "ok"}
        entries.append(item)
        if path.name.startswith("~$") or (extensions and path.suffix.lower().lstrip(".") not in extensions):
            item.update(status="skipped", reason="不符合当前文件筛选条件")
            if progress:
                progress("筛选输入文件", item)
            continue
        try:
            identity = fingerprint(path)
        except OSError as error:
            item.update(status="error", reason=str(error))
            if progress:
                progress("读取文件信息", item)
            continue
        sizes[identity[2]].append((path, identity, item))

    rows = []
    group_id = copies = 0
    for size, candidates in sizes.items():
        if len(candidates) == 1:
            _, _, item = candidates[0]
            item["detail"] = "大小唯一"
            if progress:
                progress("按大小排除唯一文件", item)
            continue
        matches = defaultdict(list)
        for path, identity, item in candidates:
            try:
                checksum = digest(path)
                if fingerprint(path) != identity:
                    raise ValueError("文件在扫描期间发生变化，请重试")
            except (OSError, ValueError) as error:
                item.update(status="error", reason=str(error))
            else:
                matches[checksum].append(item)
            if progress:
                progress("校验候选文件内容", item)
        for checksum, group in matches.items():
            group[0]["detail"] = "首次出现"
            if len(group) < 2:
                continue
            group_id += 1
            copies += len(group) - 1
            for item in group:
                if item is not group[0]:
                    item["detail"] = f"与 {group[0]['source']} 内容相同"
                rows.append([group_id, item["source"], size, checksum])
    counts = Counter(e["status"] for e in entries)
    summary = {key: counts[key] for key in ("ok", "skipped", "error")}
    frame = pd.DataFrame(rows, columns=["重复组", "相对路径", "大小（字节）", "SHA-256"])
    name, dest = outputs.allocate(Path("重复文件.csv"))
    if dest:
        frame.to_csv(dest, index=False, encoding="utf-8-sig")
    product = table_product(frame, name)
    product["detail"] = f"{group_id} 组重复内容 · {copies} 个额外副本"
    return {"files": entries, "products": [product], "summary": dict(summary, rows=len(rows)), "options": options}


def perform(root, output, options, selected=None, preview=False, progress=None):
    """执行一个工具标签；返回 (批次目录或 None, 清单)，预览不会写磁盘。"""
    root, output = root.absolute().resolve(), output.absolute().resolve()
    if not root.is_dir() or root == output:
        raise ValueError("请选择存在的输入目录，并指定不同的输出目录")
    mode = options["mode"]
    if mode == "一键归档":
        return archive_batch(root, output, options["format"].lower(), options["sharpen"], progress, selected, preview)
    regex_rename = mode == "批量重命名" and options.get("rule") == "正则改名"
    pattern = compile_match(options) if regex_rename or mode == "文本替换" else None
    if mode == "批量重命名" and options.get("in_place"):
        return rename_originals(root, output, options, selected, preview, progress, pattern)
    if mode not in {"批量重命名", "文本替换", "图片转换", "文件整理", "表格处理"}:
        raise ValueError(f"未知工具：{mode}")
    allowed = {"图片转换": IMAGES, "文本替换": TEXT, "表格处理": TABLES}.get(mode)
    if mode == "表格处理":
        allowed = {"CSV 转 Excel": {".csv"}, "Excel 转 CSV": EXCEL_INPUTS, "工作表目录": EXCEL_INPUTS}.get(options["rule"], TABLES)
    sheet_pattern = options.get("sheet_pattern", "").strip() if mode == "表格处理" else ""

    def process(stage):
        """遍历选定文件，逐项形成记录；需要跨文件的数据处理放到遍历结束后。"""
        outputs = Outputs(stage)
        entries, frames, products = [], [], []
        inventory = []
        paths = selected if selected is not None else input_files(root, output)
        if mode == "文件整理" and options["rule"] == "查找重复":
            return duplicate_report(root, paths, outputs, options, progress)
        seq = options.get("start", 1)
        extensions = {v.lower().lstrip(".") for v in columns(options.get("extensions", ""))}
        total_rows = 0
        for path in paths:
            rel = path.relative_to(root)
            item = {"source": rel.as_posix(), "status": "ok"}
            entries.append(item)
            created = []

            def destination(name):
                """登记本文件的产物，若后续写入失败，可只撤回本文件的副本。"""
                name, target = outputs.allocate(name)
                if target:
                    created.append(target)
                return name, target

            try:
                ext = path.suffix.lower()
                if path.name.startswith("~$") or (allowed is not None and ext not in allowed) or (extensions and ext.lstrip(".") not in extensions):
                    item.update(status="skipped", reason="不符合当前文件筛选条件")
                elif mode == "批量重命名":
                    name = rename_to(path, seq, options, pattern)
                    seq += 1
                    if regex_rename and name == path.name:
                        item.update(status="skipped", reason="名称没有变化")
                    else:
                        item["output"], dest = destination(Path("重命名") / rel.with_name(name))
                        if dest:
                            shutil.copy2(path, dest)
                elif mode == "文本替换":
                    samples = []
                    raw, count, detail = replace_text(path, options, pattern, samples)
                    item["detail"] = detail
                    if count:
                        item["output"], dest = destination(Path("文本") / rel)
                        if dest:
                            dest.write_bytes(raw)
                        item.update(sample_columns=["原文行号", "匹配文字", "替换结果"], sample_rows=samples)
                    else:
                        item.update(status="skipped", reason=detail)
                elif mode == "图片转换":
                    new_ext = IMAGE_OUTPUTS[options["format"]]
                    item["output"], dest = destination(Path("图片") / rel.with_suffix(new_ext))
                    fmt, before, after = convert_image(path, dest or Path(item["output"]), options["sharpen"],
                                                       options.get("quality", 90), options.get("max_edge", 0),
                                                       preview=preview)
                    item["detail"] = f"{fmt} {before[0]}×{before[1]} → {options['format']} {after[0]}×{after[1]}"
                elif mode == "文件整理":
                    rule = options["rule"]
                    if rule == "文件清单":
                        record = inventory_row(path, rel)
                        inventory.append(record)
                        item["detail"] = f"{record[2]} · {record[4]} 字节 · {record[5]}"
                    else:
                        folder = category(path) if rule == "按类型" else datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m")
                        item["output"], dest = destination(Path("分类") / folder / path.name)
                        if dest:
                            shutil.copy2(path, dest)
                elif mode == "表格处理" and options["rule"] == "工作表目录":
                    # 目录直接读取工作表元信息，含有无效列名的工作表也可以先被找到。
                    names = [(i, title) for i, title in enumerate(sheet_names(path), 1)
                             if matches_sheet(title, sheet_pattern)]
                    if names:
                        inventory.extend((rel.as_posix(), i, title) for i, title in names)
                        item.update(rows=len(names), detail="、".join(title for _, title in names))
                    else:
                        item.update(status="skipped", reason="没有匹配的工作表")
                elif mode == "表格处理":
                    tables = read_tables(path, rel.as_posix(), options.get("header_row", 0),
                                         options.get("header_depth", 1), sheet_pattern)
                    if not tables:
                        item.update(status="skipped", reason="没有匹配的工作表")
                    else:
                        item.update(rows=sum(len(t) for t in tables), detail=describe_tables(tables))
                        if options["rule"] in {"CSV 转 Excel", "Excel 转 CSV"}:
                            converted = convert_tables(tables, rel, destination)
                            products.extend(converted)
                            item["output"] = "、".join(p["output"] for p in converted)
                            total_rows += item["rows"]
                        else:
                            frames.extend(tables)
                else:
                    raise ValueError(f"未知工具：{mode}")
            except Exception as error:
                for dest in created:
                    dest.unlink(missing_ok=True)
                item.pop("output", None)
                item.pop("rows", None)
                item.pop("detail", None)
                item.update(status="error", reason=str(error))
            if progress:
                progress("正在预览" if preview else mode, item)
        if frames:
            products, total_rows = transform_tables(frames, options, outputs)
        if (mode == "文件整理" and options["rule"] == "文件清单") or (mode == "表格处理" and options["rule"] == "工作表目录"):
            headers = INVENTORY_COLUMNS if mode == "文件整理" else ["来源文件", "工作表序号", "工作表名"]
            frame = pd.DataFrame(inventory, columns=headers)
            name, dest = outputs.allocate(Path(options["rule"] + ".csv"))
            if dest:
                frame.to_csv(dest, index=False, encoding="utf-8-sig")
            products.append(table_product(frame, name))
            total_rows = len(frame)
        count = Counter(entry["status"] for entry in entries)
        summary = {k: count[k] for k in ("ok", "skipped", "error")}
        return {"files": entries, "products": products, "summary": dict(summary, rows=total_rows), "options": options}

    return (None, process(None)) if preview else publish(root, output, process, progress)
