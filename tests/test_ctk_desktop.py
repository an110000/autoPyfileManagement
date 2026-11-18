"""用 CustomTkinter 的实际控件和事件循环核对六个标签及可拖动布局。"""

import importlib.util
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch

from PIL import Image

if importlib.util.find_spec("customtkinter") is None:
    raise unittest.SkipTest("当前环境未安装 CustomTkinter 界面")

from ctk_ui import MODES, Workbench


class TkDesktopTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        """在同一 Tk 根窗口中驱动按钮，后台任务也经过正式的消息队列。"""
        cls.w = Workbench()
        cls.w.geometry("1220x770")
        cls.w.update()

    @classmethod
    def tearDownClass(cls):
        """测试结束后关闭根窗口，释放 Tcl 和窗口资源。"""
        cls.w.destroy()

    def finish(self):
        """推动 Tk 事件循环，等待线程交付清单，并检查按钮已恢复。"""
        deadline = time.monotonic() + 15
        while self.w.worker is not None and time.monotonic() < deadline:
            self.w.update()
            time.sleep(.01)
        self.w.update()
        self.assertIsNone(self.w.worker)
        self.assertEqual(self.w.preview.cget("state"), "normal")

    def test_preview_and_publish_each_tool(self):
        """逐标签验证预览零写入、正式 ZIP、样本行、归档清单与窗口分隔。"""
        w = self.w
        with tempfile.TemporaryDirectory(dir=Path.cwd()) as tmp:
            source = Path(tmp) / "in"
            source.mkdir()
            (source / "notes.txt").write_text("hello", encoding="utf-8")
            (source / "orders.csv").write_text("sku,qty\n0012,2\n0012,4", encoding="utf-8-sig")
            Image.new("RGB", (24, 12)).save(source / "img.png")
            w.source.set(str(source))
            for mode in MODES:
                with self.subTest(mode=mode):
                    dest = Path(tmp) / mode
                    w.output.set(str(dest))
                    w.select_mode(mode)
                    w.update()
                    fields = w.forms[mode].vars
                    if mode == "文本替换":
                        fields["find"].set("hello")
                        fields["replacement"].set("done")
                    elif mode == "图片转换":
                        fields["resize_mode"].set("等比适应宽高")
                        fields["width"].set("30")
                        fields["height"].set("20")
                        self.assertNotIn("max_edge", w.forms[mode].options())
                    elif mode == "表格处理":
                        fields["rule"].set("分组汇总")
                        fields["keys"].set("sku")
                        fields["amount"].set("qty")
                    self.assertEqual(len(w.split.panes()), 1)
                    w.preview.invoke()
                    self.finish()
                    self.assertIsNotNone(w.report, w.status.get())
                    self.assertFalse(dest.exists())
                    self.assertTrue(w.details.tree.get_children())
                    preview = w.report
                    if mode == "表格处理":
                        first = w.sample.tree.get_children()[0]
                        self.assertEqual(w.sample.tree.item(first, "values"), ("0012", "2", "6"))
                        self.assertLess(w.tabs.winfo_rooty(), w.banner.winfo_rooty())
                        self.assertLess(w.banner.winfo_width(), w.split.winfo_width())
                        self.assertGreater(w.sample.tree.winfo_height(), 135)
                        initial = w.sample.tree.winfo_height()
                        w.split.sash_place(0, 0, w.split.sash_coord(0)[1] - 65)
                        w.update()
                        self.assertGreater(w.sample.tree.winfo_height(), initial)
                    if mode == "一键归档":
                        w.filter_rows("ZIP 成员")
                        self.assertEqual(len(w.details.tree.get_children()), len(preview["archive_members"]))
                    w.start.invoke()
                    self.finish()
                    self.assertEqual(w.report, preview)
                    self.assertEqual(w.report["summary"]["error"], 0)
                    self.assertTrue((w.result_dir / "整理结果.zip").is_file())
                    if mode == "图片转换":
                        with Image.open(w.result_dir / "图片/img.jpg") as converted:
                            self.assertEqual(converted.size, (30, 15))

    def test_selection_wide_sample_and_failed_task(self):
        """只选一个文件，检查最后一列可横向访问，后台错误后仍能再次预览。"""
        w = self.w
        with tempfile.TemporaryDirectory(dir=Path.cwd()) as tmp:
            source = Path(tmp)
            csv = source / "wide.csv"
            csv.write_text(",".join(f"字段{i}" for i in range(18)) + "\n" + ",".join(["值"]*18),
                           encoding="utf-8-sig")
            (source / "other.txt").write_text("other", encoding="utf-8")
            w.select_mode("表格处理")
            w.forms[w.mode].vars["rule"].set("按列合并")
            with patch("ctk_ui.filedialog.askopenfilenames", return_value=(str(csv),)):
                w.pick_files()
            w.output.set("")
            w.preview.invoke()
            self.finish()
            self.assertEqual(len(w.report["files"]), 1)
            self.assertLess(w.sample.tree.xview()[1], 1)
            w.sample.tree.xview_moveto(1)
            self.assertEqual(w.sample.tree.xview()[1], 1)
            w.source.set(str(source))
            self.assertIsNone(w.selected)
            w.select_mode("文本替换")
            w.forms[w.mode].vars["find"].set("")
            w.preview.invoke()
            self.finish()
            self.assertIsNone(w.report)
            self.assertIn("批次未完成", w.status.get())
            w.forms[w.mode].vars["find"].set("other")
            w.preview.invoke()
            self.finish()
            self.assertIsNotNone(w.report)
