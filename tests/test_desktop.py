"""通过真实 Qt 控件运行预览和发布，保护工作线程、表单和数据表之间的连接。"""

import os
import tempfile
import time
import unittest
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PyQt6.QtTest import QTest
from PyQt6.QtWidgets import QApplication

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
                self.assertFalse((root / "整理结果").exists())
                window.start.click()
                self.finish_task(window)
                self.assertTrue((window.result_dir / "表格/分组汇总.xlsx").is_file())
            finally:
                window.close()
