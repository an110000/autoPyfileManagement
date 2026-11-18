"""回归样例均由测试生成，不读取个人文件。运行 python -m unittest discover -s tests。"""

import codecs
import csv
import json
import re
import shutil
import tempfile
import unittest
from unittest.mock import patch
import zipfile
from pathlib import Path

from openpyxl import Workbook, load_workbook
from PIL import Image

from batch_ops import digest, perform, undo_renames


class ToolTests(unittest.TestCase):
    def setUp(self):
        """建立隔离的收件目录，同时保留一个不存在的输出目录用于验证预览。"""
        self.temp = tempfile.TemporaryDirectory(dir=Path.cwd())
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name) / "输入"
        self.root.mkdir()
        self.out = self.root / "输出"

    def execute(self, mode, preview=False, **options):
        """使用和桌面工作线程相同的业务入口，返回批次路径与处理清单。"""
        return perform(self.root, self.out, dict(options, mode=mode), preview=preview)

    def orders(self):
        """建立两份列序不同的订单，其中有跨文件重复编号、金额和缺失值。"""
        (self.root / "a.csv").write_text("编号,地区,金额\n001,东,1.1\n002,西,2.2\n", encoding="utf-8-sig")
        book = Workbook()
        book.active.title = "明细"
        book.active.append(["金额", "地区", "编号"])
        book.active.append([3.3, "东", "001"])
        book.create_sheet("空表").append(["金额", "地区", "编号"])
        book.save(self.root / "b.xlsx")

    def workbook_rows(self, result, relative):
        """重新打开真实工作簿，并关闭文件句柄以便 Windows 释放临时目录。"""
        book = load_workbook(result / relative, read_only=False)
        try:
            return list(book.active.values)
        finally:
            book.close()

    def test_rename_modes_and_collisions(self):
        """验证三种命名方式、序号、同名避让，以及预览零写入。"""
        for name in ["a.txt", "b.txt"]:
            (self.root / name).write_text(name)
        cfg = dict(rule="模板序号", template="结果", start=7, extensions=".txt")
        _, preview = self.execute("批量重命名", True, **cfg)
        self.assertFalse(self.out.exists())
        result, report = self.execute("批量重命名", **cfg)
        names = [r["output"] for r in report["files"]]
        self.assertEqual(names, [r["output"] for r in preview["files"]])
        self.assertEqual((result / names[1]).read_text(), "b.txt")
        self.assertIn("(2)", names[1])
        for rule, expected, args in [
            ("查找替换", "x.txt", dict(find="a", replacement="x")),
            ("前后缀", "XaZ.txt", dict(prefix="X", suffix="Z")),
            ("模板序号", "007_a.txt", dict(template="{n:03d}_{name}", start=7)),
        ]:
            _, rep = self.execute("批量重命名", True, rule=rule, **args)
            self.assertTrue(rep["files"][0]["output"].endswith(expected))
        _, bad = self.execute("批量重命名", True, rule="模板序号", template="../外部")
        self.assertEqual(bad["summary"]["error"], 2)

    def test_text_encoding_regex_and_line_endings(self):
        """替换时保留 UTF-8 BOM、GBK 和 CRLF，普通替换允许反斜线原样写回。"""
        original = codecs.BOM_UTF8 + b"Code-01\r\nCODE-02\r\n"
        (self.root / "a.txt").write_bytes(original)
        (self.root / "b.txt").write_bytes("项目 Code-03".encode("gbk"))
        cfg = dict(find=r"code-(\d+)", replacement=r"ID:\1", encoding="自动", regex=True, case=False)
        result, rep = self.execute("文本替换", **cfg)
        self.assertEqual(rep["summary"]["ok"], 2)
        self.assertEqual((result / "文本/a.txt").read_bytes(), codecs.BOM_UTF8 + b"ID:01\r\nID:02\r\n")
        self.assertEqual((result / "文本/b.txt").read_bytes().decode("gbk"), "项目 ID:03")
        self.assertEqual((self.root / "a.txt").read_bytes(), original)
        result, _ = self.execute("文本替换", **dict(cfg, find="Code", replacement=r"\1", regex=False))
        self.assertIn(b"\\1-01", (result / "文本/a.txt").read_bytes())

    def test_original_names_and_undo(self):
        """预览不改名，原位执行可撤销；原名被占用时撤销必须保留两个文件。"""
        src = self.root / "订单.txt"
        src.write_text("receipt")
        cfg = dict(rule="前后缀", prefix="新_", suffix="", in_place=True)
        _, preview = self.execute("批量重命名", True, **cfg)
        self.assertTrue(src.exists())
        self.assertEqual(preview["files"][0]["output"], "新_订单.txt")
        _, report = self.execute("批量重命名", **cfg)
        self.assertFalse(src.exists())
        src.write_text("occupied")
        self.assertEqual(len(undo_renames(self.root, report["files"])), 1)
        self.assertEqual(src.read_text(), "occupied")
        src.unlink()
        self.assertEqual(undo_renames(self.root, report["files"]), [])
        self.assertEqual(src.read_text(), "receipt")

    def test_images_organize_and_duplicates(self):
        """验证五种图片输出、复制分类及内容查重，重复文件保留在原位置。"""
        Image.new("RGBA", (12, 9), (250, 0, 0, 0)).save(self.root / "img.png")
        for fmt in ["JPEG", "PNG", "WEBP", "BMP", "TIFF"]:
            result, rep = self.execute("图片转换", format=fmt, quality=85, sharpen=True)
            with Image.open(result / rep["files"][0]["output"]) as img:
                self.assertEqual((img.format, img.size), (fmt, (12, 9)))
        for name in ["a.txt", "b.txt"]:
            (self.root / name).write_text("same")
        result, rep = self.execute("文件整理", rule="查找重复")
        self.assertIn("1 组", rep["products"][0]["detail"])
        self.assertIn("b.txt", (result / "重复文件.csv").read_text(encoding="utf-8-sig"))
        for rule in ["按类型", "按修改月份"]:
            result, rep = self.execute("文件整理", rule=rule)
            self.assertTrue(all((result / e["output"]).exists() for e in rep["files"]))

    def test_table_rules(self):
        """核对跨表去重、筛选投影、十进制汇总、分组拆分及错误列名。"""
        self.orders()
        result, rep = self.execute("表格处理", rule="按键去重", keys="编号", columns="编号,金额")
        data = self.workbook_rows(result, rep["products"][0]["output"])
        self.assertEqual([r[-2:] for r in data], [("编号","金额"), ("001","1.1"), ("002","2.2")])
        result, rep = self.execute("表格处理", rule="分组汇总", keys="地区", amount="金额")
        self.assertEqual(self.workbook_rows(result,rep["products"][0]["output"]),
                         [("地区","记录数","合计"), ("东",2,4.4), ("西",1,2.2)])
        _, rep = self.execute("表格处理", True, rule="按列合并", filter_column="地区", filter_value="东")
        self.assertEqual(rep["summary"]["rows"], 2)
        result, rep = self.execute("表格处理", rule="按值拆分", keys="地区")
        self.assertEqual([p["rows"] for p in rep["products"]], [2, 1])
        with zipfile.ZipFile(result / "整理结果.zip") as archive:
            self.assertIsNone(archive.testzip())
            self.assertEqual(json.loads(archive.read("处理清单.json")), rep)
        with self.assertRaisesRegex(ValueError, "找不到列"):
            self.execute("表格处理", rule="分组汇总", keys="不存在", amount="金额")

    def test_table_format_round_trip(self):
        """转换应覆盖工作簿内所有表，包括只有表头的空表。"""
        self.orders()
        result, _ = self.execute("表格处理", rule="Excel 转 CSV")
        self.assertEqual((result / "表格/b.xlsx/002_空表.csv").read_text(encoding="utf-8-sig").strip(), "金额,地区,编号")
        result, _ = self.execute("表格处理", rule="CSV 转 Excel")
        rows = self.workbook_rows(result, "表格/a.xlsx")
        self.assertEqual(rows[1], ("001","东","1.1"))

    def test_titles_grouped_headers_and_late_columns(self):
        """同一本表内分别识别横幅、两层表头和单列表头，迟到的额外值也要留存。"""
        book = Workbook()
        ws = book.active
        ws.title = "分组"
        ws.append(["月度台账"])
        ws.merge_cells("A1:C1")
        ws.append(["货号", "订单"])
        ws.merge_cells("A2:A3")
        ws.merge_cells("B2:C2")
        ws.append([None, "数量", "金额"])
        ws.append(["001", 2, 10])
        ws.append(["002", 3, 12, "附注"])
        flat = book.create_sheet("单列表头")
        flat.append(["姓名"])
        flat.append(["甲", "乙"])
        book.save(self.root / "新表.xlsx")
        cfg = dict(rule="Excel 转 CSV")
        _, preview = self.execute("表格处理", True, **cfg)
        self.assertFalse(self.out.exists())
        result, report = self.execute("表格处理", **cfg)
        self.assertEqual(preview, report)
        self.assertIn("表头第 2–3 行", report["files"][0]["detail"])
        self.assertEqual(report["summary"]["rows"], 3)
        product, single = report["products"]
        self.assertEqual(product["sample_columns"], ["货号", "数量", "金额", "未命名列_4"])
        with (result / product["output"]).open(encoding="utf-8-sig", newline="") as stream:
            self.assertEqual(list(csv.reader(stream))[1:], [["001","2","10",""], ["002","3","12","附注"]])
        self.assertEqual(single["sample_rows"], [["甲", "乙"]])
        result, merged = self.execute("表格处理", rule="按列合并")
        self.assertEqual([row[2] for row in self.workbook_rows(result, merged["products"][0]["output"])[1:]],
                         [4, 5, 2])

    def test_csv_manual_header_and_unnamed_collision(self):
        """手选 CSV 的原始记录号；缺标题列与真实同名列必须分别保存。"""
        (self.root / "订单.csv").write_text("说明\n\n货号,,未命名列_2\n001,东,7,extra\n", encoding="utf-8")
        cfg = dict(rule="按列合并", header_row=3)
        _, report = self.execute("表格处理", True, **cfg)
        p = report["products"][0]
        self.assertEqual(p["sample_columns"], ["货号", "未命名列_2_", "未命名列_2", "未命名列_4"])
        self.assertEqual(p["sample_rows"], [["001","东","7","extra"]])
        _, bad = self.execute("表格处理", True, rule="按列合并", header_row=10)
        self.assertIn("第 10 行", bad["files"][0]["reason"])
        (self.root / "订单.csv").write_text("编号,编号\n1,2", encoding="utf-8")
        _, duplicate = self.execute("表格处理", True, rule="按列合并")
        self.assertEqual(duplicate["summary"]["error"], 1)

    def test_legacy_excel_sheet_conversion(self):
        """真实的合成 BIFF 工作簿含前导零、日期、布尔和跨两层的重复子标题。"""
        shutil.copyfile(Path(__file__).parent / "fixtures/legacy_orders.xls", self.root / "历史.xls")
        result, report = self.execute("表格处理", rule="Excel 转 CSV")
        actual, empty = report["products"]
        self.assertEqual(actual["sample_columns"], ["货号", "订货 / 数量", "日期", "交付 / 数量", "完成"])
        self.assertEqual(actual["sample_rows"], [["0012", "4", "2025-01-07 00:00:00", "2", "True"]])
        self.assertEqual(empty["sample_rows"], [])
        self.assertTrue((result / empty["output"]).is_file())
        _, totals = self.execute("表格处理", True, rule="分组汇总", keys="货号", amount="交付 / 数量")
        self.assertEqual(totals["products"][0]["sample_rows"], [["0012","1","2"]])

    def test_inventory_and_categories(self):
        """分类、清单与预览应一致，空文件保留零字节，相同名字按原目录区分。"""
        categories = {"a.PPTX":"演示文稿", "a.mp4":"视频", "a.WOFF2":"字体", "a.wav":"音频",
                      "a.JSON":"数据", "a.py":"程序", "a.svg":"图片", "a.xls":"表格", "a.bin":"其他"}
        for name in categories:
            (self.root / name).write_bytes(b"")
        (self.root / "sub").mkdir()
        (self.root / "sub/a.wav").write_bytes(b"abc")
        _, preview = self.execute("文件整理", True, rule="文件清单")
        self.assertFalse(self.out.exists())
        result, report = self.execute("文件整理", rule="文件清单")
        self.assertEqual(report, preview)
        with (result / "文件清单.csv").open(encoding="utf-8-sig", newline="") as stream:
            rows = {r["相对路径"]: r for r in csv.DictReader(stream)}
        self.assertEqual(report["summary"]["rows"], 10)
        self.assertEqual(rows["sub/a.wav"]["大小（字节）"], "3")
        for name, category in categories.items():
            self.assertEqual(rows[name]["类别"], category)
            self.assertEqual(rows[name]["大小（字节）"], "0")
        _, sorted_files = self.execute("文件整理", True, rule="按类型")
        self.assertEqual(sorted_files["files"][-1]["output"], "分类/音频/a (2).wav")

    def test_catalog_and_sheet_selection_before_reading(self):
        """列出所有表名时无需验证内容；只选中第二张表转换时保留原始序号。"""
        b = Workbook()
        b.active.title = "说明"
        b.active.append(["重复", "重复"])
        sheet = b.create_sheet("一月岗位")
        sheet.append(["编号", "备注"])
        sheet.append(["001", "东区"])
        b.save(self.root / "部门.xlsx")
        shutil.copyfile(Path(__file__).parent / "fixtures/legacy_orders.xls", self.root / "历史.xls")
        (self.root / "旁注.csv").write_text("编号\n002\n", encoding="utf-8")
        _, catalog = self.execute("表格处理", True, rule="工作表目录")
        rows = catalog["products"][0]["sample_rows"]
        self.assertIn(["部门.xlsx", "2", "一月岗位"], rows)
        self.assertIn(["历史.xls", "2", "空表"], rows)
        self.assertEqual(catalog["summary"], dict(ok=2, skipped=1, error=0, rows=4))
        self.assertFalse(self.out.exists())
        result, report = self.execute("表格处理", rule="Excel 转 CSV", sheet_pattern="*岗位")
        self.assertEqual(report["summary"], dict(ok=1, skipped=2, error=0, rows=1))
        self.assertTrue((result / "表格/部门.xlsx/002_一月岗位.csv").is_file())
        _, old = self.execute("表格处理", True, rule="按列合并", sheet_pattern="订单")
        self.assertEqual(old["summary"], dict(ok=1, skipped=2, error=0, rows=1))

    def test_table_filters(self):
        """普通包含将标点视为文字；正则和空值按行筛选，写出的结果与预览相同。"""
        (self.root / "项目.csv").write_text("编号,备注\n001,A.1\n002,A21\n003,\n",encoding="utf-8")
        cases = [("包含文字", ".", ["001"]), ("正则匹配", "^A[0-9]", ["002"]),
                 ("为空", "", ["003"]), ("非空", "", ["001","002"]), ("精确等于", "A.1", ["001"])]
        for op, value, expected in cases:
            with self.subTest(op=op):
                cfg = dict(rule="按列合并", filter_column="备注", filter_op=op, filter_value=value)
                _, preview = self.execute("表格处理", True, **cfg)
                result, executed = self.execute("表格处理", **cfg)
                self.assertEqual(executed, preview)
                rows = self.workbook_rows(result, executed["products"][0]["output"])
                self.assertEqual([r[3] for r in rows[1:]], expected)
        with self.assertRaisesRegex(ValueError, "筛选正则无效"):
            self.execute("表格处理", True, rule="按列合并", filter_column="备注",
                         filter_op="正则匹配", filter_value="[")

    def test_regex_rename_copy_original_and_undo(self):
        """两种执行方式共用捕获组和扩展名规则；无匹配时跳过，原位变更能够撤销。"""
        old = self.root / "INV_002.TXT"
        old.write_text("original")
        (self.root / "notes.txt").touch()
        cfg = dict(rule="正则改名", find=r"^inv_(\d+)$", replacement=r"订单-\1", case=False)
        _, preview = self.execute("批量重命名", True, **cfg)
        self.assertFalse(self.out.exists())
        base, report = self.execute("批量重命名", **cfg)
        self.assertEqual(report, preview)
        self.assertEqual(report["summary"]["skipped"], 1)
        self.assertEqual((base / "重命名/订单-002.TXT").read_text(), "original")
        _, changed = self.execute("批量重命名", **cfg, in_place=True)
        self.assertFalse(old.exists())
        self.assertEqual(undo_renames(self.root, changed["files"]), [])
        self.assertEqual(old.read_text(), "original")
        with self.assertRaisesRegex(re.error, "group"):
            self.execute("批量重命名", True, **dict(cfg, replacement=r"\8"))

    def test_text_sample_numbers_encoding_and_limits(self):
        """保留 UTF-16 两种字节序与 BOM；只展示前五处，写入仍覆盖全部匹配。"""
        cfg = dict(find=r"X(\d)", replacement=r"单\1", case=True, regex=True, encoding="自动")
        content = "X1\r\nX2\rX3 X4\nX5\nX6"
        for enc, bom in [("utf-16-le", codecs.BOM_UTF16_LE), ("utf-16-be", codecs.BOM_UTF16_BE)]:
            path = self.root / (enc + ".txt")
            path.write_bytes(bom + content.encode(enc))
        _, preview = self.execute("文本替换", True, **cfg)
        base, actual = self.execute("文本替换", **cfg)
        self.assertEqual(preview, actual)
        for entry in actual["files"]:
            encoding = Path(entry["source"]).stem
            samples = entry["sample_rows"]
            self.assertEqual([row[0] for row in samples], ["1", "2", "3", "3", "4"])
            self.assertEqual(samples[-1][1:], ["X5", "单5"])
            out = (base / entry["output"]).read_bytes()
            self.assertEqual(out[:2], (self.root / entry["source"]).read_bytes()[:2])
            self.assertEqual(out[2:].decode(encoding), content.replace("X", "单"))

    def test_oriented_image_dimensions_and_broken_preview(self):
        """旋转后再缩小；小图保留原尺寸，截断图片在预览时也会报告失败。"""
        picture = Image.new("RGB", (80, 40), "red")
        exif = picture.getexif()
        exif[274] = 6
        picture.save(self.root / "rotate.jpg", exif=exif)
        Image.new("RGBA", (4, 5), (10, 20, 30, 0)).save(self.root / "small.png")
        broken = self.root / "truncated.bmp"
        picture.save(broken)
        broken.write_bytes(broken.read_bytes()[:100])
        opts = dict(format="PNG", sharpen=True, max_edge=20)
        _, preview = self.execute("图片转换", True, **opts)
        self.assertFalse(self.out.exists())
        result, report = self.execute("图片转换", **opts)
        self.assertEqual(preview["files"], report["files"])
        self.assertEqual(report["summary"]["error"], 1)
        with Image.open(result / "图片/rotate.png") as img:
            self.assertEqual(img.size, (10, 20))
        with Image.open(result / "图片/small.png") as img:
            self.assertEqual(img.size, (4, 5))
            self.assertEqual(img.getpixel((0, 0))[3], 0)

    def test_image_width_height_preview_matches_output(self):
        """用带方向标记的横图验证单边、双边、等比边框、放大及透明通道。"""
        image = Image.new("RGBA", (8, 4), (20, 50, 80, 90))
        exif = image.getexif()
        exif[274] = 6
        image.save(self.root / "turn.png", exif=exif)
        cases = [
            ("指定宽高", 0, 0, (4, 8)), ("指定宽高", 7, 0, (7, 14)),
            ("指定宽高", 0, 6, (3, 6)), ("指定宽高", 9, 3, (9, 3)),
            ("等比适应宽高", 30, 20, (10, 20)), ("等比适应宽高", 6, 18, (6, 12)),
            ("等比适应宽高", 0, 1, (1, 1)),
        ]
        for mode, width, height, expected in cases:
            with self.subTest(mode=mode, width=width, height=height):
                opts = dict(format="PNG", sharpen=False, resize_mode=mode, width=width, height=height)
                _, preview = self.execute("图片转换", True, **opts)
                result, actual = self.execute("图片转换", **opts)
                self.assertEqual(actual, preview)
                self.assertEqual(actual["files"][0]["detail"], f"PNG 4×8 → PNG {expected[0]}×{expected[1]}")
                with Image.open(result / actual["files"][0]["output"]) as converted:
                    self.assertEqual(converted.size, expected)
                    self.assertEqual(converted.getpixel((0, 0))[3], 90)

    def test_reject_unencodable_resize_during_preview(self):
        """过大像素数、超出编码上限、非法 JSON 尺寸均在预览阶段返回失败。"""
        Image.new("RGB", (1, 2)).save(self.root / "small.png")
        cases = [("PNG", 65535, 65535), ("PNG", 65535, 0), ("PNG", -3, 8),
                 ("PNG", 4.5, 12), ("WEBP", 16384, 1), ("JPEG", 1, 65501)]
        for fmt, width, height in cases:
            with self.subTest(fmt=fmt, width=width, height=height):
                opts = dict(format=fmt, resize_mode="指定宽高", width=width, height=height, sharpen=False)
                _, report = self.execute("图片转换", True, **opts)
                self.assertEqual(report["summary"]["error"], 1)
                self.assertNotIn("output", report["files"][0])
                self.assertFalse(self.out.exists())

    def test_size_first_duplicate_report(self):
        """只对同大小候选求完整散列，排除同尺寸不同内容，并保留零字节重复项。"""
        values = {"a": b"x", "b": b"x", "c": b"y", "d": b"unique", "e": b"", "f": b""}
        for name, raw in values.items():
            (self.root / name).write_bytes(raw)
        with patch("batch_ops.digest", wraps=digest) as hashed:
            _, preview = self.execute("文件整理", True, rule="查找重复")
        self.assertEqual({call.args[0].name for call in hashed.call_args_list}, {"a","b","c","e","f"})
        result, report = self.execute("文件整理", rule="查找重复")
        self.assertEqual(preview, report)
        with (result / "重复文件.csv").open(encoding="utf-8-sig", newline="") as f:
            rows = list(csv.DictReader(f))
        self.assertEqual([r["相对路径"] for r in rows], ["a", "b", "e", "f"])
        self.assertEqual(report["summary"]["rows"], 4)

    def test_filtered_column_profile(self):
        """在行筛选之后统计，编号、空值、并列众数及零行输入都可对照工作簿。"""
        (self.root / "data.csv").write_text("区域,编号,备注\n东,002,\n东,001,NA\n西,001,NA", encoding="utf-8")
        cfg = dict(rule="列质量报告", columns="编号,备注", filter_column="区域", filter_value="东")
        _, preview = self.execute("表格处理", True, **cfg)
        result, actual = self.execute("表格处理", **cfg)
        self.assertEqual(preview, actual)
        rows = self.workbook_rows(result, actual["products"][0]["output"])
        self.assertEqual(rows[1:], [("编号",2,2,0,2,"002",1), ("备注",2,1,1,1,"NA",1)])
        _, empty = self.execute("表格处理", True, **dict(cfg, filter_value="北"))
        self.assertEqual(empty["products"][0]["sample_rows"][0], ["编号","0","0","0","0","","0"])

    def test_archive_plan_matches_published_zip(self):
        """组合任务预览不写盘，多个同名图片有各自去向；损坏项不会进入 ZIP。"""
        self.orders()
        (self.root / "note.txt").write_text("hello")
        (self.root / "bad.png").write_bytes(b"broken")
        for folder in ["东", "西"]:
            dest = self.root / folder
            dest.mkdir()
            Image.new("RGB", (3, 2)).save(dest / "img.png")
        opts = dict(format="PNG", sharpen=False)
        _, plan = self.execute("一键归档", True, **opts)
        self.assertFalse(self.out.exists())
        finished, actual = self.execute("一键归档", **opts)
        self.assertEqual(plan, actual)
        self.assertEqual(actual["summary"], dict(ok=5, error=1, skipped=0, rows=3))
        with zipfile.ZipFile(finished / "整理结果.zip") as archive:
            self.assertEqual(sorted(archive.namelist()), plan["archive_members"])
            self.assertIsNone(archive.testzip())
