"""归档台，PyQt6 桌面入口。运行 python batch_ui.py。"""

import json
import os
import sys
import time
from pathlib import Path

from PyQt6.QtCore import QAbstractTableModel, QModelIndex, QPointF, QRectF, Qt, QThread, QUrl, pyqtSignal
from PyQt6.QtGui import QColor, QDesktopServices, QFont, QIcon, QPainter, QPixmap, QPolygonF
from PyQt6.QtWidgets import (
    QApplication, QComboBox, QFileDialog, QFrame, QGridLayout, QHBoxLayout,
    QHeaderView, QLabel, QLineEdit, QMainWindow, QProgressBar, QPushButton,
    QSizePolicy, QSplitter, QStackedWidget, QTableView, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget, QTabBar, QScrollArea, QMessageBox,
)

from batch_ops import perform, undo_renames
from rule_forms import RuleForm
from workbench_view import HEADER_IMAGE, METRICS, MODES, STATUS, batch_paths, filter_records, record_detail, report_view


class BatchWorker(QThread):
    progress = pyqtSignal(str, str, dict)
    completed = pyqtSignal(object, dict)
    failed = pyqtSignal(str)

    def __init__(self, root, output, options, selected, preview, parent):
        """保存主线程拍下的参数快照，避免批次中途读取界面控件。"""
        super().__init__(parent)
        self.root, self.output, self.options = root, output, options
        self.selected, self.preview = selected, preview
        self.counts = dict.fromkeys(METRICS, 0)
        self.last_update = 0

    def update_progress(self, phase, entry):
        """逐项累计文件和行数，以不超过约每秒 8 次的频率向界面发信号。"""
        if entry is not None:
            self.counts[entry["status"]] += 1
            self.counts["rows"] += entry.get("rows", 0)
            # 处理许多小文件时，限制跨线程刷新频率，计数仍逐个更新。
            if time.monotonic() - self.last_update < 0.12:
                return
        self.last_update = time.monotonic()
        self.progress.emit(phase, entry["source"] if entry else "", self.counts.copy())

    def run(self):
        """在工作线程运行预览或正式批次，向主线程交付清单或错误原因。"""
        try:
            result, report = perform(self.root, self.output, self.options, self.selected,
                                     self.preview, self.update_progress)
        except Exception as error:
            self.failed.emit(str(error))
        else:
            self.completed.emit(result, report)


class FileModel(QAbstractTableModel):
    """仅为可见行提供显示内容，切换筛选时保留清单中的顺序。"""

    def __init__(self, parent):
        """分别保存完整清单和当前筛选后的行。"""
        super().__init__(parent)
        self.records = []
        self.visible = []

    def select_status(self, status, members=()):
        """显示选定状态的源记录，或显示 ZIP 中每个成员的完整相对路径。"""
        self.beginResetModel()
        self.visible = filter_records(self.records, status, members)
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()):
        """返回当前可见行数；这是平面表，没有子节点。"""
        return 0 if parent.isValid() else len(self.visible)

    def columnCount(self, parent=QModelIndex()):
        """提供来源、状态和处理说明三列。"""
        return 0 if parent.isValid() else 3

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        """为横向表头提供中文名称。"""
        if orientation == Qt.Orientation.Horizontal and role == Qt.ItemDataRole.DisplayRole:
            return ("源文件", "状态", "输出位置 / 原因")[section]
        return None

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        """按显示、颜色或提示角色返回单元格内容，长路径可悬停阅读。"""
        if not index.isValid():
            return None
        item = self.visible[index.row()]
        column = index.column()
        if role == Qt.ItemDataRole.ForegroundRole and column == 1:
            return QColor({"ok": "#236c50", "skipped": "#796537", "error": "#ab4a34"}[item["status"]])
        if role not in (Qt.ItemDataRole.DisplayRole, Qt.ItemDataRole.ToolTipRole):
            return None
        if column == 0:
            return item["source"]
        if column == 1:
            return item.get("label") or STATUS[item["status"]]
        return record_detail(item)


class FolderArt(QWidget):
    """以 Qt 绘图绘制文件夹和两张纸，可随屏幕缩放。"""

    def __init__(self):
        """固定插画的逻辑尺寸，由 Qt 根据屏幕倍率缩放。"""
        super().__init__()
        self.setFixedSize(156, 97)

    def paintEvent(self, event):
        """使用 QPainter 绘制文件夹、表格纸张及照片，无外部图标依赖。"""
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setPen(Qt.PenStyle.NoPen)
        p.translate(78, 50)
        p.rotate(-7)

        def rect(x, y, w, h, color, radius=0):
            """以指定颜色填充一个局部坐标矩形。"""
            p.setBrush(QColor(color))
            p.drawRoundedRect(QRectF(x, y, w, h), radius, radius)

        rect(-62, -21, 123, 63, "#819f66", 6)
        rect(-62, -27, 46, 15, "#819f66", 5)
        p.save()
        p.rotate(-9)
        rect(-40, -40, 54, 64, "#f2f4dc", 2)
        rect(-34, -32, 22, 4, "#6f977e")
        for y in (-22, -15, -8):
            rect(-34, y, 41, 2, "#b4cbb2")
        p.restore()
        p.save()
        p.rotate(12)
        rect(8, -41, 47, 48, "#f2f3dd", 2)
        rect(12, -37, 39, 34, "#abcab2")
        rect(39, -33, 7, 7, "#fff3a7", 3.5)
        p.setBrush(QColor("#55856c"))
        p.drawPolygon(QPolygonF([QPointF(12,-6), QPointF(25,-23), QPointF(32,-15),
                                 QPointF(39,-25), QPointF(51,-11), QPointF(51,-3), QPointF(12,-3)]))
        p.restore()
        p.shear(-0.09, 0)
        rect(-62, -6, 127, 50, "#d9e58c", 6)
        rect(34, 3, 20, 5, "#f5f2c3", 2)
        p.end()


def label(text, name=""):
    """创建纯文本标签；对象名用于应用统一样式。"""
    widget = QLabel(text)
    widget.setTextFormat(Qt.TextFormat.PlainText)
    widget.setObjectName(name)
    return widget


def card(title):
    """创建带标题的面板，同时返回其布局以供调用方添加控件。"""
    frame = QFrame()
    frame.setObjectName("card")
    layout = QVBoxLayout(frame)
    layout.setContentsMargins(14, 12, 14, 12)
    layout.setSpacing(7)
    layout.addWidget(label(title, "cardTitle"))
    return frame, layout


class Workbench(QMainWindow):
    def __init__(self):
        """建立功能标签、共享的输入区域、预览表和后台任务状态栏。"""
        super().__init__()
        self.setWindowTitle("归档台 · 文件整理工作台")
        screen = self.screen().availableGeometry()
        self.resize(min(1360, screen.width() - 24), min(860, screen.height() - 54))
        self.worker = None
        self.selected = None
        self.report = None
        self.result_dir = None
        self.undo_state = None
        self.close_after_run = False
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(14, 10, 14, 8)
        layout.setSpacing(8)
        self.setWindowIcon(QIcon(FolderArt().grab()))
        self.tabs = QTabBar()
        self.tabs.setDrawBase(False)
        self.tabs.setExpanding(True)
        for mode in MODES:
            self.tabs.addTab(mode)
        layout.addWidget(self.tabs)

        self.columns = QSplitter(Qt.Orientation.Horizontal)
        self.columns.setChildrenCollapsible(False)
        left = QWidget()
        left_box = QVBoxLayout(left)
        left_box.setContentsMargins(0, 0, 0, 0)
        left_box.setSpacing(8)
        self.banner = QLabel()
        self.banner.setFixedHeight(64)
        self.banner.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
        self.banner.setPixmap(QPixmap(str(HEADER_IMAGE)))
        self.banner.setScaledContents(True)
        head = QHBoxLayout(self.banner)
        head.setContentsMargins(18, 0, 0, 0)
        head.addWidget(label("文件整理工作台", "title"))
        left_box.addWidget(self.banner)
        self.settings, setup = card("输入与规则")
        self.settings.setMinimumWidth(270)
        self.source = self.path_field(setup, "输入文件夹", "选择或粘贴待整理文件夹的路径")
        self.source.textChanged.connect(self.clear_selection)
        select_row = QHBoxLayout()
        select = QPushButton("选择指定文件…")
        select.clicked.connect(self.pick_files)
        select_row.addWidget(select)
        self.scope = label("扫描目录及子目录", "caption")
        select_row.addWidget(self.scope, 1)
        setup.addLayout(select_row)
        self.output = self.path_field(setup, "输出文件夹", "留空时保存到输入目录内的「整理结果」")
        self.forms = QStackedWidget()
        for mode in MODES:
            self.forms.addWidget(RuleForm(mode))
        setup.addWidget(self.forms)
        setup.addStretch()
        actions = QHBoxLayout()
        self.preview = QPushButton("预览结果")
        self.preview.clicked.connect(lambda: self.start_batch(True))
        actions.addWidget(self.preview)
        self.start = QPushButton("执行处理")
        self.start.setObjectName("primary")
        self.start.setMinimumHeight(42)
        self.start.clicked.connect(lambda: self.start_batch(False))
        actions.addWidget(self.start, 1)
        scroll = QScrollArea()
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setWidgetResizable(True)
        scroll.setWidget(self.settings)
        scroll.setMinimumWidth(290)
        left_box.addWidget(scroll, 1)
        left_box.addLayout(actions)
        self.columns.addWidget(left)

        result, right = card("处理结果")
        counts = QGridLayout()
        counts.setSpacing(6)
        self.numbers = {}
        self.captions = {}
        for col, (key, title) in enumerate(METRICS.items()):
            tile = QFrame()
            tile.setObjectName("metric")
            inner = QHBoxLayout(tile)
            inner.setContentsMargins(8, 4, 8, 4)
            inner.setSpacing(4)
            self.captions[key] = label(title, "caption")
            inner.addWidget(self.captions[key], 1)
            self.numbers[key] = label("—", "number")
            inner.addWidget(self.numbers[key])
            counts.addWidget(tile, 0, col)
            counts.setColumnStretch(col, 1)
        right.addLayout(counts)
        toolbar = QHBoxLayout()
        toolbar.addWidget(label("文件明细", "field"))
        toolbar.addStretch()
        self.filter = QComboBox()
        self.filter.setAccessibleName("按处理状态筛选")
        self.filter.addItem("全部状态", "")
        for key, title in STATUS.items():
            self.filter.addItem(title, key)
        self.filter.setEnabled(False)
        self.filter.currentIndexChanged.connect(self.filter_rows)
        toolbar.addWidget(self.filter)
        right.addLayout(toolbar)
        self.stack = QStackedWidget()
        self.stack.setMinimumHeight(100)
        self.empty = label("选择输入文件夹或指定文件，然后预览或执行。", "empty")
        self.empty.setWordWrap(True)
        self.empty.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.stack.addWidget(self.empty)
        self.model = FileModel(self)
        self.table = QTableView()
        self.table.setModel(self.model)
        self.table.verticalHeader().hide()
        self.table.verticalHeader().setDefaultSectionSize(32)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.table.setShowGrid(False)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        self.stack.addWidget(self.table)
        self.result_split = QSplitter(Qt.Orientation.Vertical)
        self.result_split.setChildrenCollapsible(False)
        self.result_split.addWidget(self.stack)
        self.sample_panel = QWidget()
        sample_box = QVBoxLayout(self.sample_panel)
        sample_box.setContentsMargins(0, 0, 0, 0)
        sample_box.setSpacing(4)
        self.sample_picker = QComboBox()
        self.sample_picker.setAccessibleName("选择预览明细")
        self.sample_picker.currentIndexChanged.connect(self.show_sample)
        sample_box.addWidget(self.sample_picker)
        self.sample = QTableWidget()
        self.sample.setEditTriggers(QTableView.EditTrigger.NoEditTriggers)
        self.sample.setMinimumHeight(80)
        self.sample.verticalHeader().setDefaultSectionSize(24)
        sample_box.addWidget(self.sample, 1)
        self.result_split.addWidget(self.sample_panel)
        self.result_split.setStretchFactor(0, 1)
        self.result_split.setStretchFactor(1, 1)
        self.result_split.setSizes([330, 295])
        self.sample_panel.hide()
        right.addWidget(self.result_split, 1)
        self.open_result = QPushButton("打开结果目录")
        self.open_result.setEnabled(False)
        self.open_result.setToolTip("查看工作簿、图片、文档、清单和 ZIP")
        self.open_result.clicked.connect(self.open_folder)
        log_row = QHBoxLayout()
        self.export_log = QPushButton("导出清单…")
        self.export_log.setEnabled(False)
        self.export_log.clicked.connect(self.save_report)
        log_row.addWidget(self.export_log)
        self.undo_button = QPushButton("撤销上次原位改名")
        self.undo_button.clicked.connect(self.restore_names)
        self.undo_button.hide()
        log_row.addWidget(self.undo_button)
        log_row.addStretch()
        log_row.addWidget(self.open_result)
        right.addLayout(log_row)
        self.columns.addWidget(result)
        self.columns.setStretchFactor(0, 0)
        self.columns.setStretchFactor(1, 1)
        self.columns.setSizes([330, 960])
        layout.addWidget(self.columns, 1)

        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedHeight(5)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.hide()
        layout.addWidget(self.progress_bar)
        self.status = label("就绪", "status")
        self.status.setWordWrap(True)
        self.status.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        self.status.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        layout.addWidget(self.status)
        self.tabs.currentChanged.connect(self.mode_changed)

    def clear_selection(self):
        """手工修改目录或重新选目录时，恢复递归扫描全部文件。"""
        self.selected = None
        self.scope.setText("扫描目录及子目录")

    def pick_files(self):
        """将系统对话框选中的文件作为本次范围，共同父目录用于生成相对路径。"""
        paths, _ = QFileDialog.getOpenFileNames(self, "选择待处理文件", self.source.text())
        if paths:
            files = sorted(set(Path(p).resolve() for p in paths))
            try:
                root = os.path.commonpath([str(p.parent) for p in files])
            except ValueError:
                self.status.setText("请选择同一磁盘中的文件。")
                return
            self.source.setText(root)
            self.selected = files
            self.scope.setText(f"仅处理选中的 {len(files)} 个文件")

    def mode_changed(self, index):
        """切换独立规则表单，清空上一工具的结果，防止误读旧预览。"""
        self.forms.setCurrentIndex(index)
        self.report = self.result_dir = None
        self.model.records = []
        self.model.select_status("")
        self.filter.setCurrentIndex(0)
        member_index = self.filter.findData("members")
        if MODES[index] == "一键归档" and member_index < 0:
            self.filter.addItem("ZIP 成员", "members")
        elif MODES[index] != "一键归档" and member_index >= 0:
            self.filter.removeItem(member_index)
        self.filter.setEnabled(False)
        self.open_result.setEnabled(False)
        self.export_log.setEnabled(False)
        self.sample_panel.hide()
        self.set_empty("当前工具：" + MODES[index] + "\n\n填写规则后执行处理。")
        for number in self.numbers.values():
            number.setText("—")
        self.status.setText("就绪")

    def save_report(self):
        """按用户选择的位置导出完整 JSON 清单（含参数、来源及输出样例）。"""
        path, _ = QFileDialog.getSaveFileName(self, "导出处理清单", "处理清单.json", "JSON (*.json)")
        if path and self.report is not None:
            try:
                Path(path).write_text(json.dumps(self.report, ensure_ascii=False, indent=2), encoding="utf-8")
            except OSError as error:
                self.status.setText(str(error))
            else:
                self.status.setText(f"已导出：{path}")

    def restore_names(self):
        """逆序还原最近一次原位改名，仅对身份与修改时间均匹配的文件操作。"""
        if not self.undo_state or self.worker is not None:
            return
        root, entries = self.undo_state
        errors = undo_renames(root, entries)
        self.status.setText("已还原文件名。" if not errors else "部分名称未还原：" + "；".join(errors))
        self.undo_button.setVisible(bool(errors))
        if not errors:
            self.undo_state = None

    def path_field(self, layout, title, placeholder):
        """添加路径输入及系统目录选择按钮，允许直接粘贴路径。"""
        edit = QLineEdit()
        edit.setPlaceholderText(placeholder)
        edit.setAccessibleName(title)
        heading = label(title, "field")
        heading.setBuddy(edit)
        layout.addWidget(heading)
        row = QHBoxLayout()
        row.addWidget(edit, 1)
        browse = QPushButton("浏览…")
        browse.setAccessibleName("浏览" + title)

        def choose():
            """浏览目录；用户取消时保留当前输入。"""
            chosen = QFileDialog.getExistingDirectory(self, "选择" + title, edit.text())
            if chosen:
                edit.setText(chosen)
                if edit is self.source:
                    self.clear_selection()

        browse.clicked.connect(choose)
        row.addWidget(browse)
        layout.addLayout(row)
        return edit

    def set_empty(self, text):
        """以状态提示代替结果表，显示等待、无匹配或错误原因。"""
        self.empty.setText(text)
        self.stack.setCurrentWidget(self.empty)

    def filter_rows(self):
        """按下拉框选项更新明细；没有匹配时显示可读提示。"""
        self.model.select_status(self.filter.currentData(), (self.report or {}).get("archive_members", ()))
        if self.model.visible:
            self.stack.setCurrentWidget(self.table)
        elif self.model.records:
            self.set_empty("此状态下没有文件。")
        else:
            self.set_empty("输入文件夹中没有可列出的文件。")

    def start_batch(self, preview=False):
        """核对目录，固定本次规则和文件范围，然后启动唯一的后台任务。"""
        if self.worker is not None:
            return
        try:
            root, dest = batch_paths(self.source.text(), self.output.text())
        except (ValueError, OSError, RuntimeError) as error:
            self.status.setText(str(error))
            return
        options = self.forms.currentWidget().options()
        if not preview and options.get("in_place"):
            answer = QMessageBox.question(self, "原位改名", "本次将修改输入目录中的原文件名。继续执行？",
                                          QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                          QMessageBox.StandardButton.No)
            if answer != QMessageBox.StandardButton.Yes:
                return
        self.settings.setEnabled(False)
        self.start.setEnabled(False)
        self.preview.setEnabled(False)
        self.captions["ok"].setText("可处理文件" if preview else "成功文件")
        self.tabs.setEnabled(False)
        self.export_log.setEnabled(False)
        self.sample_panel.hide()
        self.report = self.result_dir = None
        self.open_result.setEnabled(False)
        self.filter.setEnabled(False)
        self.model.records = []
        self.filter.setCurrentIndex(0)
        self.model.select_status("")
        self.set_empty("正在处理文件…\n\n完成后可在这里查看每个文件的去向。")
        self.update_progress("正在扫描目录", "", dict.fromkeys(METRICS, 0))
        self.progress_bar.setRange(0, 0)
        self.progress_bar.show()
        self.start.setText("正在预览…" if preview else "正在处理…")
        self.worker = BatchWorker(root, dest, options,
                                  self.selected, preview, self)
        self.worker.progress.connect(self.update_progress)
        self.worker.completed.connect(self.completed)
        self.worker.failed.connect(self.failed)
        self.worker.finished.connect(self.finished)
        self.worker.start()

    def update_progress(self, phase, source, counts):
        """在主线程刷新四项计数和最近处理的相对路径。"""
        self.status.setText(f"{phase} · {source}" if source else phase)
        for key, value in counts.items():
            self.numbers[key].setText(f"{value:,}")

    def completed(self, result, report):
        """展示输入文件与生成产物的记录；仅实际执行后允许打开结果目录。"""
        self.result_dir = result
        self.report = report
        if result and report["options"].get("in_place") and any("identity" in e for e in report["files"]):
            self.undo_state = (self.worker.root.resolve(), [dict(e) for e in report["files"]])
            self.undo_button.show()
        summary = report["summary"]
        records, samples = report_view(report, result is None)
        self.sample_picker.clear()
        for title, item in samples:
            self.sample_picker.addItem(title, item)
        self.sample_panel.setVisible(bool(samples))
        self.model.records = records
        self.filter_rows()
        self.filter.setEnabled(bool(records))
        self.open_result.setEnabled(result is not None)
        self.export_log.setEnabled(True)
        title = "预览完成，尚未写入文件" if result is None else "处理完成"
        if summary["error"]:
            title += "，部分文件失败"
        self.update_progress(title, str(result) if result else "", summary)

    def show_sample(self, index):
        """按真实列展示所选产物的前五行；切换分组可分别核对拆分结果。"""
        product = self.sample_picker.itemData(index)
        if not product:
            return
        cols, rows = product["sample_columns"], product["sample_rows"]
        self.sample.clear()
        self.sample.setColumnCount(len(cols))
        self.sample.setHorizontalHeaderLabels(cols)
        self.sample.setRowCount(len(rows))
        for r, row in enumerate(rows):
            for c, value in enumerate(row):
                self.sample.setItem(r, c, QTableWidgetItem(value))
        self.sample.resizeColumnsToContents()
        if 0 < len(cols) <= 5:
            self.sample.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        else:
            self.sample.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)

    def failed(self, reason):
        """展示未能完成整批处理的原因，留待用户调整输入或规则。"""
        self.set_empty("批次未完成\n\n" + reason)
        self.status.setText("批次未完成")

    def finished(self):
        """回收工作线程，恢复按钮，并处理运行中收到的关闭请求。"""
        self.worker.deleteLater()
        self.worker = None
        self.settings.setEnabled(True)
        self.start.setEnabled(True)
        self.preview.setEnabled(True)
        self.tabs.setEnabled(True)
        self.progress_bar.hide()
        self.start.setText("执行处理")
        if self.close_after_run:
            self.close()

    def open_folder(self):
        """使用系统文件管理器打开最后成功发布的批次。"""
        if self.result_dir and not QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.result_dir))):
            self.status.setText(f"系统未能打开结果目录：{self.result_dir}")

    def closeEvent(self, event):
        """若仍在写入则延迟关闭，避免留下不完整成品或提前销毁线程。"""
        if self.worker is not None:
            self.close_after_run = True
            self.setWindowTitle("批次完成后关闭 · 归档台")
            event.ignore()
        else:
            super().closeEvent(event)


STYLE = """
QMainWindow { background:#f2f5ef; }
QTabBar::tab { background:#e5ede1; color:#476452; padding:9px 8px; margin-right:3px; border-radius:5px; }
QTabBar::tab:selected { background:#165d51; color:#fff; font-weight:700; }
QWidget { color:#203c36; }
QLabel#title { color:#f4f7e6; font-size:24px; font-weight:700; }
QSplitter::handle { background:#d2ddd5; }
QSplitter::handle:hover { background:#73a38c; }
QSplitter::handle:horizontal { width:6px; }
QSplitter::handle:vertical { height:6px; }
QFrame#card { background:white; border:1px solid #d6e0d6; border-radius:12px; }
QLabel#cardTitle { font-size:20px; font-weight:700; padding-bottom:5px; }
QLabel#field { font-weight:600; }
QLineEdit, QComboBox { background:#fcfdfa; border:1px solid #cdd9ce; border-radius:5px; padding:6px; }
QLineEdit:focus, QComboBox:focus, QPushButton:focus { border:1px solid #4b846a; }
QPushButton { background:#edf3e9; border:1px solid #d3dfd1; border-radius:6px; padding:6px 11px; }
QPushButton:hover { background:#e1ecdc; }
QPushButton:disabled { background:#f1f3ef; color:#78857c; }
QPushButton#primary { background:#165d51; color:white; font-size:15px; font-weight:700; border:0; }
QPushButton#primary:hover { background:#104a42; }
QPushButton#primary:disabled { background:#647c73; }
QRadioButton, QCheckBox { spacing:6px; min-height:23px; }
QFrame#metric { background:#f1f5ec; border-radius:7px; }
QLabel#caption, QLabel#status { color:#546a5c; font-size:12px; }
QLabel#number { color:#244c3d; font-size:22px; font-weight:700; }
QLabel#empty { background:#f9fbf6; border:1px dashed #cbd8c7; border-radius:8px; color:#5b7361; padding:12px; }
QTableView { background:white; alternate-background-color:#f6f8f3; border:1px solid #dce4d9; selection-background-color:#d0e5d4; selection-color:#16392a; }
QHeaderView::section { border:0; border-bottom:1px solid #d6dfd2; background:#edf2e9; padding:6px; color:#41584a; }
QProgressBar { background:#dae5d8; border:0; border-radius:2px; }
QProgressBar::chunk { background:#4c896a; }
"""


def main():
    """启动 Qt 事件循环，显示六个功能标签的主窗口，返回进程退出码。"""
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setFont(QFont("Microsoft YaHei UI", 10))
    app.setStyleSheet(STYLE)
    window = Workbench()
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
