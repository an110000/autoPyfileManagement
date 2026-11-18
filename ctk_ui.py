"""CustomTkinter 桌面工作台。界面收集参数，共享引擎负责读写文件。"""

import json
import os
import queue
import threading
import time
import tkinter as tk
import webbrowser
from pathlib import Path
from tkinter import filedialog, font, messagebox, ttk

import customtkinter as ctk
from PIL import Image, ImageTk

from batch_ops import perform, undo_renames
from workbench_view import (
    HEADER_IMAGE, METRICS, MODES, STATUS, batch_paths, build_fields,
    filter_records, record_detail, report_view, task_options, visible_fields,
)

BG, PAPER, INK, GREEN = "#f2f5ef", "#ffffff", "#203c36", "#165d51"
SOFT, LINE = "#eaf1e8", "#d3dfd5"


def button(parent, title, command, primary=False, **options):
    """统一按钮颜色和焦点区域，主要动作采用深绿色。"""
    return ctk.CTkButton(parent, text=title, command=command, height=32,
                         fg_color=GREEN if primary else SOFT, text_color="white" if primary else INK,
                         hover_color="#104a42" if primary else "#d9e7da", **options)


class RuleForm(ctk.CTkFrame):
    """把共享规则声明转成可编辑的控件，每个标签单独保存输入。"""

    def __init__(self, parent, mode):
        """建好全部字段，再按当前规则隐藏无关字段。"""
        super().__init__(parent, fg_color="transparent")
        self.mode = mode
        self.vars, self.rows, self.inputs, self.bounds = {}, {}, {}, {}
        self.columnconfigure(0, weight=1)
        build_fields(self, mode)
        for key in ("rule", "format", "header_row", "filter_column", "filter_op"):
            if key in self.vars:
                self.vars[key].trace_add("write", self.refresh)
        self.refresh()

    def field(self, key, title, variable):
        """登记变量，生成字段的容器；隐藏容器会同时隐藏标签及输入框。"""
        row = ctk.CTkFrame(self, fg_color="transparent")
        row.grid(row=len(self.rows), column=0, sticky="ew", pady=(2, 4))
        row.columnconfigure(0, weight=1)
        self.rows[key], self.vars[key] = row, variable
        if title:
            ctk.CTkLabel(row, text=title, text_color=INK, anchor="w", height=20).grid(sticky="ew")
        return row

    def line(self, key, title, value="", hint=""):
        """创建文本输入；有示例的字段在输入框下显示填写提示。"""
        var = tk.StringVar(self, value)
        row = self.field(key, title, var)
        self.inputs[key] = ctk.CTkEntry(row, textvariable=var)
        self.inputs[key].grid(sticky="ew")
        if hint:
            ctk.CTkLabel(row, text=hint, text_color="#536b60", anchor="w",
                         height=16, font=ctk.CTkFont(size=11)).grid(sticky="ew")

    def choice(self, key, title, choices):
        """用只读菜单提供引擎接受的枚举值。"""
        var = tk.StringVar(self, choices[0])
        row = self.field(key, title, var)
        self.inputs[key] = ctk.CTkOptionMenu(row, variable=var, values=choices,
                                             fg_color=GREEN, button_color=GREEN)
        self.inputs[key].grid(sticky="ew")

    def number(self, key, title, minimum, maximum, value, zero_text=""):
        """显示带上下箭头的整数框；提交时还会校验键盘输入的范围。"""
        var = tk.StringVar(self, str(value))
        row = self.field(key, title + (f"（0 = {zero_text}）" if zero_text else ""), var)
        self.bounds[key] = title, minimum, maximum
        self.inputs[key] = ttk.Spinbox(row, from_=minimum, to=maximum, textvariable=var)
        self.inputs[key].grid(sticky="ew", ipady=3)

    def check(self, key, title, value=False):
        """建立布尔开关，把勾选状态直接交给引擎。"""
        var = tk.BooleanVar(self, value)
        row = self.field(key, "", var)
        self.inputs[key] = ctk.CTkCheckBox(row, text=title, variable=var, fg_color=GREEN)
        self.inputs[key].grid(sticky="w")

    def refresh(self, *_):
        """响应规则、表头行、筛选列等变化，在原来的位置恢复可见控件。"""
        active = visible_fields(self.mode, {k: v.get() for k, v in self.vars.items()})
        for key, row in self.rows.items():
            if key in active:
                row.grid()
            else:
                row.grid_remove()

    def options(self):
        """生成一次任务的参数，仅校验和提交所选规则用到的字段。"""
        values = {k: v.get() for k, v in self.vars.items()}
        result = task_options(self.mode, values)
        for key, value in result.items():
            if key in self.bounds:
                title, low, high = self.bounds[key]
                try:
                    value = int(value)
                    if not low <= value <= high:
                        raise ValueError
                except ValueError:
                    raise ValueError(f"{title}请输入 {low} 至 {high} 的整数。") from None
                result[key] = value
        return result


class DataTable(ttk.Frame):
    """带双向滚动条的数据表；用列序号作内部标识，标题可重复或为空。"""

    def __init__(self, parent):
        """把 Treeview 和滚动条绑定，窗口缩小时仍能访问右侧列。"""
        super().__init__(parent)
        self.tree = ttk.Treeview(self, show="headings", height=3, style="Workbench.Treeview")
        vertical = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.horizontal = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vertical.set, xscrollcommand=self.update_horizontal)
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        self.horizontal.grid(row=1, column=0, sticky="ew")
        self.tree.tag_configure("odd", background="#f4f8f3")

    def update_horizontal(self, first, last):
        """仅当列宽溢出时显示横向滚动条，把短表节省出的高度留给数据行。"""
        self.horizontal.set(first, last)
        if float(first) > 0 or float(last) < 1:
            self.horizontal.grid()
        else:
            self.horizontal.grid_remove()

    def fill(self, headers, rows, widths=None):
        """重建列并填入内容；样本列按标题和数据实测宽度设置滚动范围。"""
        self.tree.delete(*self.tree.get_children())
        ids = tuple(f"col{i}" for i in range(len(headers)))
        self.tree.configure(columns=ids)
        measure = font.nametofont("TkDefaultFont").measure
        for i, (key, title) in enumerate(zip(ids, headers)):
            width = widths[i] if widths else min(480, max(100, measure(title) + 40,
                max((measure(str(row[i])) + 28 for row in rows if len(row) > i), default=0)))
            self.tree.heading(key, text=title, anchor="w")
            self.tree.column(key, width=width, minwidth=width, stretch=not widths or i != 1, anchor="w")
        for i, values in enumerate(rows):
            self.tree.insert("", "end", values=values, tags=("odd",) if i % 2 else ())


def run_batch(events, root, dest, options, selected, preview):
    """后台线程只接触普通 Python 对象，通过队列向 Tk 主线程交付进度和清单。"""
    counts, last_update = dict.fromkeys(METRICS, 0), 0

    def progress(phase, item):
        """逐项计数，把密集的进度通知合并到约每秒八次。"""
        nonlocal last_update
        if item:
            counts[item["status"]] += 1
            counts["rows"] += item.get("rows", 0)
            if time.monotonic() - last_update < .12:
                return
        last_update = time.monotonic()
        events.put(("progress", (phase, item["source"] if item else "", counts.copy())))

    try:
        result = perform(root, dest, options, selected, preview, progress)
    except Exception as exc:
        events.put(("failed", str(exc)))
    else:
        events.put(("done", result))


class Workbench(ctk.CTk):
    """六标签桌面窗口，顶部选工具，左侧设规则，右侧核对明细与样本。"""

    def __init__(self):
        """先构建固定导航与两个可拖动面板，再接入窗口关闭事件。"""
        ctk.set_appearance_mode("light")
        super().__init__(fg_color=BG)
        self.title("归档台 · 文件整理工作台 · CustomTkinter")
        # Tk 给出物理像素，CTk 的 geometry 接受缩放前的尺寸。在高 DPI 屏上先换算，
        # 否则默认窗口可能超出屏幕，底部的预览和执行按钮也会落到任务栏下面。
        scale = self._get_window_scaling()
        width = min(1360, int(self.winfo_screenwidth() / scale) - 30)
        height = min(860, int(self.winfo_screenheight() / scale) - 85)
        self.geometry(f"{width}x{height}")
        self.minsize(760, 495)
        self.selected = self.worker = self.undo_state = self.result_dir = self.report = None
        self.close_after_run = False
        self.events, self.controls, self.records, self.samples = queue.Queue(), [], [], {}
        self.mode, self.sample_fraction = MODES[0], .53
        self.source, self.output, self.status = (tk.StringVar(self) for _ in range(3))
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        self.tabs = ctk.CTkSegmentedButton(self, values=list(MODES), command=self.select_mode,
            height=35, dynamic_resizing=False, fg_color=LINE, selected_color="#a4d2b8",
            selected_hover_color="#a4d2b8", unselected_color=SOFT, unselected_hover_color="#d6e5d9",
            text_color=INK)
        self.tabs.grid(row=0, column=0, sticky="ew", padx=12, pady=(8, 6))
        self.tabs.set(self.mode)
        self.columns = tk.PanedWindow(self, orient="horizontal", sashwidth=7, bg=LINE, bd=0)
        self.columns.grid(row=1, column=0, sticky="nsew", padx=12)
        left = ctk.CTkFrame(self.columns, fg_color=BG, corner_radius=0)
        right = ctk.CTkFrame(self.columns, fg_color=PAPER, corner_radius=0)
        self.columns.add(left, minsize=272, width=325, stretch="never")
        self.columns.add(right, minsize=350, stretch="always")
        self.build_sidebar(left)
        self.build_results(right)
        ctk.CTkLabel(self, textvariable=self.status, text_color=INK, anchor="w", height=22
                     ).grid(row=2, column=0, sticky="ew", padx=15, pady=3)
        self.select_mode(self.mode)
        self.protocol("WM_DELETE_WINDOW", self.request_close)

    def build_sidebar(self, left):
        """将标题叠在虚化底图上，规则独立滚动，让两个操作按钮始终可见。"""
        left.columnconfigure(0, weight=1)
        left.rowconfigure(1, weight=1)
        self.banner = tk.Canvas(left, height=60, width=1, bg=BG, highlightthickness=0)
        self.banner.grid(sticky="ew", pady=(0, 6))
        with Image.open(HEADER_IMAGE) as image:
            self.banner_image = image.convert("RGBA")
        self.art_id = self.banner.create_image(0, 0, anchor="nw")
        self.banner.create_text(16, 30, text="文件整理工作台", fill="white",
                                font=("Microsoft YaHei UI", 16, "bold"), anchor="w")
        self.banner.bind("<Configure>", self.resize_banner)
        scroll = ctk.CTkScrollableFrame(left, fg_color=PAPER)
        scroll.grid(row=1, column=0, sticky="nsew")
        scroll.columnconfigure(0, weight=1)
        ctk.CTkLabel(scroll, text="输入与规则", text_color=INK, font=ctk.CTkFont(size=18, weight="bold"),
                     anchor="w").grid(row=0, sticky="ew", padx=7)
        self.path_field(scroll, 1, "输入文件夹", self.source)
        scope = ctk.CTkFrame(scroll, fg_color="transparent")
        scope.grid(row=3, sticky="ew", padx=7)
        pick = button(scope, "指定文件…", self.pick_files, width=84)
        pick.pack(side="left")
        self.controls.append(pick)
        self.scope = ctk.CTkLabel(scope, text="扫描目录及子目录", text_color=INK)
        self.scope.pack(side="left", padx=6)
        self.source.trace_add("write", self.clear_selection)
        self.path_field(scroll, 4, "输出文件夹（可留空）", self.output)
        self.forms = {mode: RuleForm(scroll, mode) for mode in MODES}
        for form in self.forms.values():
            form.grid(row=6, column=0, sticky="ew", padx=7, pady=(3, 8))
        actions = ctk.CTkFrame(left, fg_color="transparent")
        actions.grid(row=2, sticky="ew", pady=(6, 0))
        actions.columnconfigure((0, 1), weight=1, uniform="actions")
        self.preview = button(actions, "预览结果", lambda: self.start_batch(True), width=85)
        self.start = button(actions, "执行处理", lambda: self.start_batch(False), primary=True, width=85)
        self.preview.grid(row=0, column=0, sticky="ew", padx=(0, 3))
        self.start.grid(row=0, column=1, sticky="ew", padx=(3, 0))

    def resize_banner(self, event):
        """跟随左栏宽度缩放底图，文字始终由前景画布清晰绘制。"""
        image = self.banner_image.resize((max(1, event.width), max(1, event.height)),
                                         getattr(Image, "Resampling", Image).LANCZOS)
        self.banner_photo = ImageTk.PhotoImage(image, master=self)
        self.banner.itemconfigure(self.art_id, image=self.banner_photo)

    def path_field(self, parent, row, title, var):
        """创建可直接粘贴路径的输入框和目录浏览按钮。"""
        ctk.CTkLabel(parent, text=title, text_color=INK, anchor="w", height=23
                     ).grid(row=row, sticky="ew", padx=7)
        line = ctk.CTkFrame(parent, fg_color="transparent")
        line.grid(row=row+1, sticky="ew", padx=7, pady=(0, 5))
        line.columnconfigure(0, weight=1)
        entry = ctk.CTkEntry(line, textvariable=var, width=80)
        entry.grid(row=0, column=0, sticky="ew", padx=(0, 4))
        browse = button(line, "浏览", lambda: self.browse(var), width=53)
        browse.grid(row=0, column=1)
        self.controls.extend((entry, browse))

    def browse(self, var):
        """把所选文件夹写入相应路径，取消时保留原输入。"""
        path = filedialog.askdirectory(parent=self, initialdir=var.get() or None)
        if path:
            var.set(path)

    def clear_selection(self, *_):
        """修改输入路径后恢复递归扫描，丢弃旧目录的文件选择。"""
        self.selected = None
        self.scope.configure(text="扫描目录及子目录")

    def pick_files(self):
        """选择一组文件，取共同父目录供引擎记录相对路径。"""
        paths = filedialog.askopenfilenames(parent=self, initialdir=self.source.get() or None)
        if paths:
            files = sorted({Path(p).resolve() for p in paths})
            try:
                self.source.set(os.path.commonpath([p.parent for p in files]))
            except ValueError:
                self.status.set("请选择同一磁盘中的文件。")
                return
            self.selected = files
            self.scope.configure(text=f"已选 {len(files)} 个文件")

    def build_results(self, right):
        """以紧凑统计栏、纵向分隔条和底部工具栏包围两块结果表。"""
        right.columnconfigure(0, weight=1)
        right.rowconfigure(3, weight=1)
        style = ttk.Style(self)
        style.configure("Workbench.Treeview", rowheight=27, background=PAPER, foreground=INK,
                        font=("Microsoft YaHei UI", 10))
        style.configure("Workbench.Treeview.Heading", foreground=INK, font=("Microsoft YaHei UI", 10, "bold"))
        ctk.CTkLabel(right, text="处理结果", text_color=INK, anchor="w",
                     font=ctk.CTkFont(size=19, weight="bold")).grid(row=0, sticky="ew", padx=10, pady=(5, 2))
        counts = ctk.CTkFrame(right, fg_color=SOFT)
        counts.grid(row=1, sticky="ew", padx=10)
        self.metrics = {}
        for i, title in enumerate(METRICS.values()):
            counts.columnconfigure(i, weight=1, uniform="metrics")
            label = ctk.CTkLabel(counts, text=title + "   —", text_color=INK)
            label.grid(row=0, column=i, sticky="ew", pady=4)
            self.metrics[title] = label
        toolbar = ctk.CTkFrame(right, fg_color="transparent")
        toolbar.grid(row=2, sticky="ew", padx=10, pady=5)
        ctk.CTkLabel(toolbar, text="文件明细", text_color=INK).pack(side="left")
        self.filters = {"全部状态": "", **{title: key for key, title in STATUS.items()}, "ZIP 成员": "members"}
        self.filter = ctk.CTkOptionMenu(toolbar, values=list(self.filters), command=self.filter_rows,
                                         width=110, fg_color=GREEN, button_color=GREEN)
        self.filter.pack(side="right")
        self.split = tk.PanedWindow(right, orient="vertical", sashwidth=7, bg=LINE, bd=0)
        self.split.grid(row=3, sticky="nsew", padx=10)
        self.details = DataTable(self.split)
        self.split.add(self.details, minsize=75, stretch="always")
        self.sample_panel = ctk.CTkFrame(self.split, fg_color=PAPER, corner_radius=0)
        self.sample_panel.rowconfigure(1, weight=1)
        self.sample_panel.columnconfigure(0, weight=1)
        self.picker = ctk.CTkOptionMenu(self.sample_panel, values=["数据样本"], command=self.show_sample,
                                         anchor="w", fg_color=GREEN, button_color=GREEN)
        self.picker.grid(row=0, sticky="ew", pady=(4, 3))
        self.sample = DataTable(self.sample_panel)
        self.sample.grid(row=1, sticky="nsew")
        foot = ctk.CTkFrame(right, fg_color="transparent")
        foot.grid(row=4, sticky="ew", padx=10, pady=6)
        self.export = button(foot, "导出清单…", self.export_report, width=84)
        self.export.pack(side="left")
        self.undo = button(foot, "撤销原位改名", self.restore_names, width=95)
        self.open_result = button(foot, "打开结果目录", self.open_folder, width=100)
        self.open_result.pack(side="right")

    def select_mode(self, mode):
        """只显示所选规则，清除上一批结果并更新归档专用筛选项。"""
        if self.worker is not None:
            return
        self.mode = mode
        self.tabs.set(mode)
        for name, form in self.forms.items():
            if name == mode:
                form.grid()
            else:
                form.grid_remove()
        self.filter.configure(values=[key for key in self.filters if key != "ZIP 成员" or mode == "一键归档"])
        self.clear_results()
        self.status.set("选择输入文件夹或指定文件，然后预览或执行。")

    def clear_results(self):
        """释放旧清单和样本，记住纵向比例，无样本时让明细占满结果区。"""
        self.report = self.result_dir = None
        self.records, self.samples = [], {}
        if len(self.split.panes()) == 2:
            self.sample_fraction = self.split.sash_coord(0)[1] / max(1, self.split.winfo_height())
            self.split.forget(self.sample_panel)
        self.filter.set("全部状态")
        self.filter_rows("全部状态")
        self.filter.configure(state="disabled")
        self.open_result.configure(state="disabled")
        self.export.configure(state="disabled")
        for title, label in self.metrics.items():
            label.configure(text=title + "   —")

    def filter_rows(self, title):
        """按文件状态或 ZIP 成员刷新明细；来源和处理原因沿用共享清单。"""
        rows = filter_records(self.records, self.filters[title], (self.report or {}).get("archive_members", ()))
        self.details.fill(("源文件", "状态", "输出位置 / 原因"),
            [(r["source"], r.get("label") or STATUS[r["status"]], record_detail(r)) for r in rows],
            (140, 65, 180))

    def show_sample(self, title):
        """根据选择的产物重建预览表，保留清单给出的列顺序与文本值。"""
        item = self.samples[title]
        self.sample.fill(item["sample_columns"], item["sample_rows"])

    def set_busy(self, busy):
        """运行期间锁住任务输入，避免一个批次的结果出现在另一个标签。"""
        state = "disabled" if busy else "normal"
        for widget in [self.tabs, self.preview, self.start, self.undo, *self.controls,
                       *self.forms[self.mode].inputs.values()]:
            widget.configure(state=state)

    def start_batch(self, preview):
        """校验目录和规则，固定本次文件范围，再启动唯一的工作线程。"""
        if self.worker is not None:
            return
        try:
            root, dest = batch_paths(self.source.get(), self.output.get())
            options = self.forms[self.mode].options()
        except (ValueError, OSError, RuntimeError) as exc:
            self.status.set(str(exc))
            return
        if not preview and options.get("in_place") and not messagebox.askyesno(
                "原位改名", "本次将修改输入目录中的原文件名。继续执行？", parent=self, default="no"):
            return
        self.clear_results()
        self.previewing, self.batch_root = preview, root.resolve()
        self.set_busy(True)
        self.update_progress("正在扫描目录", "", dict.fromkeys(METRICS, 0))
        files = tuple(self.selected) if self.selected is not None else None
        self.worker = threading.Thread(target=run_batch, args=(self.events, root, dest, options, files, preview))
        self.worker.start()
        self.after(35, self.poll)

    def update_progress(self, phase, source, counts):
        """按预览或执行显示文件计数，并在状态行显示最近处理的文件。"""
        for key, title in METRICS.items():
            caption = "可处理文件" if key == "ok" and self.previewing else title
            self.metrics[title].configure(text=f"{caption}   {counts[key]:,}")
        self.status.set(f"{phase} · {source}" if source else phase)

    def poll(self):
        """主线程定时取消息；成功、失败两条路径均会恢复按钮并处理关闭请求。"""
        while True:
            try:
                kind, data = self.events.get_nowait()
            except queue.Empty:
                self.after(35, self.poll)
                return
            if kind == "progress":
                self.update_progress(*data)
                continue
            try:
                if kind == "done":
                    self.completed(*data)
                else:
                    self.status.set("批次未完成 · " + data)
            finally:
                self.worker = None
                self.set_busy(False)
                if self.close_after_run:
                    self.destroy()
            return

    def completed(self, result, report):
        """展示文件与产物；有样本时还原第二块面板，正式执行后开放输出目录。"""
        self.result_dir, self.report = result, report
        self.records, items = report_view(report, self.previewing)
        self.samples = dict(items)
        self.filter_rows(self.filter.get())
        self.filter.configure(state="normal" if self.records else "disabled")
        if items:
            self.picker.configure(values=list(self.samples))
            self.picker.set(items[0][0])
            self.show_sample(items[0][0])
            self.split.add(self.sample_panel, minsize=75, stretch="always")
            self.split.update_idletasks()
            self.split.sash_place(0, 0, int(self.split.winfo_height() * self.sample_fraction))
        if result:
            self.open_result.configure(state="normal")
            if report["options"].get("in_place") and any("identity" in e for e in report["files"]):
                self.undo_state = self.batch_root, report["files"]
                self.undo.pack(side="left", padx=4)
        self.export.configure(state="normal")
        title = "预览完成，尚未写入文件" if self.previewing else "处理完成"
        if report["summary"]["error"]:
            title += "，部分文件失败"
        self.update_progress(title, str(result or ""), report["summary"])

    def restore_names(self):
        """按引擎记录的文件身份撤销最近一批原位改名。"""
        if self.undo_state and self.worker is None:
            errors = undo_renames(*self.undo_state)
            self.status.set("部分名称未还原：" + "；".join(errors) if errors else "已还原文件名。")
            if not errors:
                self.undo_state = None
                self.undo.pack_forget()

    def export_report(self):
        """把当前清单保存为可供命令行重放的 UTF-8 JSON。"""
        path = filedialog.asksaveasfilename(parent=self, defaultextension=".json",
            initialfile="处理清单.json", filetypes=[("JSON", "*.json")])
        if path and self.report is not None:
            try:
                Path(path).write_text(json.dumps(self.report, ensure_ascii=False, indent=2), encoding="utf-8")
            except OSError as exc:
                self.status.set(str(exc))
            else:
                self.status.set(f"已导出 · {path}")

    def open_folder(self):
        """使用操作系统的文件管理器打开已发布的结果目录。"""
        if self.result_dir:
            try:
                if os.name == "nt":
                    os.startfile(self.result_dir)
                else:
                    webbrowser.open(self.result_dir.resolve().as_uri())
            except OSError as exc:
                self.status.set(str(exc))

    def request_close(self):
        """在后台任务结束后关闭，确保正式写入的批次能够完成。"""
        if self.worker is None:
            self.destroy()
        else:
            self.close_after_run = True
            self.status.set("批次完成后关闭窗口。")


def main():
    """启动 CustomTkinter 的主事件循环并返回正常退出码。"""
    Workbench().mainloop()
    return 0
