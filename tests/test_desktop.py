"""通过真实 Qt 控件运行预览和发布，保护工作线程、表单和数据表之间的连接。"""

import importlib.util
import os
import tempfile
import time
import unittest
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
if importlib.util.find_spec("PyQt6") is None:
    raise unittest.SkipTest("当前环境未安装 PyQt6 界面")

from PyQt6.QtTest import QTest
from PyQt6.QtWidgets import QApplication
from openpyxl import Workbook
from PIL import Image

from batch_ui import Workbench
from rule_forms import MODES


class DesktopTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        """整个测试进程只建立一个 Qt 事件循环。"""
        cls.app = QApplication.instance() or QApplication([])

    def finish_task(self, window):
        """驱动 Qt 事件队列，直到收到工作线程结束信号；超时视为界面故障。"""
        deadline = time.monotonic() + 12
        while window.worker is not None and time.monotonic() < deadline:
            QTest.qWait(15)
        self.assertIsNone(window.worker, "后台线程没有正常结束")
        self.assertTrue(window.settings.isEnabled())
        self.assertIsNotNone(window.report, window.empty.text())

    def test_group_preview_then_publish(self):
        """从六个真实标签进入分组汇总，检查预览零写入及执行后的输出位置。"""
        with tempfile.TemporaryDirectory(dir=Path.cwd()) as temp:
            root = Path(temp)
            (root / "a.csv").write_text("货号,数量\n001,2\n001,4", encoding="utf-8-sig")
            window = Workbench()
            window.show()
            try:
                self.assertEqual(window.tabs.count(), len(MODES))
                for i, mode in enumerate(MODES):
                    window.tabs.setCurrentIndex(i)
                    self.app.processEvents()
                    self.assertEqual(window.forms.currentWidget().options()["mode"], mode)
                window.source.setText(str(root))
                window.tabs.setCurrentIndex(MODES.index("表格处理"))
                form = window.forms.currentWidget()
                form.fields["rule"].setCurrentText("分组汇总")
                form.fields["keys"].setText("货号")
                form.fields["amount"].setText("数量")
                window.preview.click()
                self.finish_task(window)
                self.assertEqual(window.sample.item(0, 2).text(), "6")
                window.resize(1220, 770)
                self.app.processEvents()
                self.assertLess(window.tabs.y(), window.columns.y())
                self.assertLess(window.banner.width(), window.stack.width())
                self.assertGreater(window.sample.height(), 135)
                before = window.sample.height()
                window.result_split.setSizes([180, 400])
                self.app.processEvents()
                self.assertGreater(window.sample.height(), before)
                self.assertFalse((root / "整理结果").exists())
                window.start.click()
                self.finish_task(window)
                self.assertTrue((window.result_dir / "表格/分组汇总.xlsx").is_file())
            finally:
                window.close()

    def test_workbook_header_controls_and_csv_sample(self):
        """实际点击 Excel 转 CSV：验证自动识别、手选表头和原始数据在结果表中的对应。"""
        with tempfile.TemporaryDirectory(dir=Path.cwd()) as temp:
            root = Path(temp)
            book = Workbook()
            sheet = book.active
            sheet.append(["季度清单"])
            sheet.merge_cells("A1:B1")
            sheet.append(["货号", "数量"])
            sheet.append(["001", 8, "新增明细"])
            book.save(root / "季度.xlsx")
            window = Workbench()
            window.show()
            try:
                window.source.setText(str(root))
                window.tabs.setCurrentIndex(MODES.index("表格处理"))
                form = window.forms.currentWidget()
                form.fields["rule"].setCurrentText("Excel 转 CSV")
                self.assertNotIn("header_depth", form.options())
                window.preview.click()
                self.finish_task(window)
                self.assertFalse((root / "整理结果").exists())
                self.assertEqual(window.sample.horizontalHeaderItem(2).text(), "未命名列_3")
                self.assertEqual(window.sample.item(0, 2).text(), "新增明细")
                form.fields["header_row"].setValue(2)
                self.assertEqual(form.options()["header_depth"], 1)
                window.start.click()
                self.finish_task(window)
                product = window.report["products"][0]
                self.assertTrue((window.result_dir / product["output"]).is_file())
                self.assertEqual(window.sample.item(0, 0).text(), "001")
                form.fields["rule"].setCurrentText("工作表目录")
                self.assertNotIn("header_row", form.options())
                window.preview.click()
                self.finish_task(window)
                self.assertEqual(window.sample.item(0, 2).text(), "Sheet")
                form.fields["rule"].setCurrentText("按列合并")
                form.fields["sheet_pattern"].setText("Sh*")
                form.fields["filter_column"].setText("货号")
                form.fields["filter_op"].setCurrentText("为空")
                self.assertNotIn("filter_value", form.options())
                window.preview.click()
                self.finish_task(window)
                self.assertEqual(window.report["summary"]["rows"], 0)
            finally:
                window.close()

    def test_six_tool_previews_and_publication(self):
        """逐个从标签设置新功能，验证按钮、后台线程、结果控件和正式产物相互对应。"""
        with tempfile.TemporaryDirectory(dir=Path.cwd()) as temp:
            source = Path(temp) / "收件"
            source.mkdir()
            (source / "id12.txt").write_bytes(b"one\ncode12\n")
            (source / "copy.txt").write_bytes(b"one\ncode12\n")
            (source / "orders.csv").write_text("sku,qty\n001,2\n001,", encoding="utf-8-sig")
            Image.new("RGB", (60, 30)).save(source / "photo.png")
            window = Workbench()
            window.show()
            try:
                window.source.setText(str(source))
                for index, mode in enumerate(MODES):
                    with self.subTest(mode=mode):
                        window.tabs.setCurrentIndex(index)
                        target = Path(temp) / mode
                        window.output.setText(str(target))
                        fields = window.forms.currentWidget().fields
                        if mode == "批量重命名":
                            fields["rule"].setCurrentText("正则改名")
                            fields["find"].setText(r"id(\d+)")
                            fields["replacement"].setText(r"订单\1")
                        elif mode == "文本替换":
                            fields["find"].setText("code12")
                            fields["replacement"].setText("已核对")
                        elif mode == "图片转换":
                            fields["resize_mode"].setCurrentText("限制最长边")
                            fields["max_edge"].setValue(12)
                            self.assertNotIn("width", window.forms.currentWidget().options())
                            fields["resize_mode"].setCurrentText("指定宽高")
                            fields["width"].setValue(19)
                            fields["height"].setValue(11)
                            self.assertNotIn("max_edge", window.forms.currentWidget().options())
                        elif mode == "文件整理":
                            fields["rule"].setCurrentText("查找重复")
                        elif mode == "表格处理":
                            fields["rule"].setCurrentText("列质量报告")
                        self.assertTrue(window.preview.isEnabled())
                        window.preview.click()
                        self.finish_task(window)
                        preview = window.report
                        self.assertFalse(target.exists())
                        if mode == "文本替换":
                            self.assertEqual(window.sample.item(0, 0).text(), "2")
                        elif mode == "一键归档":
                            self.assertIn("合并数据.xlsx", preview["archive_members"])
                            window.filter.setCurrentIndex(window.filter.findData("members"))
                            self.assertEqual([r["source"] for r in window.model.visible], preview["archive_members"])
                        window.start.click()
                        self.finish_task(window)
                        self.assertEqual(window.report, preview)
                        self.assertEqual(window.report["summary"]["error"], 0)
                        self.assertTrue((window.result_dir / "整理结果.zip").is_file())
                        if mode == "批量重命名":
                            self.assertTrue((window.result_dir / "重命名/订单12.txt").is_file())
                        elif mode == "图片转换":
                            with Image.open(window.result_dir / "图片/photo.jpg") as image:
                                self.assertEqual(image.size, (19, 11))
                        elif mode == "文件整理":
                            self.assertEqual(window.sample.rowCount(), 2)
                        elif mode == "表格处理":
                            self.assertEqual(window.sample.horizontalHeaderItem(3).text(), "空值数")
            finally:
                window.close()
