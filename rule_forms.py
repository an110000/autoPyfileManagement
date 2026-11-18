"""工作台的规则表单：选项名称、可见条件和业务参数在同一处维护。"""

from PyQt6.QtWidgets import QCheckBox, QComboBox, QFormLayout, QLineEdit, QSpinBox, QWidget


from workbench_view import MODES, build_fields, task_options, visible_fields


class RuleForm(QWidget):
    def __init__(self, mode):
        """为一个标签创建独立控件；切换标签时保留已填写的规则。"""
        super().__init__()
        self.mode = mode
        self.fields = {}
        self.form = QFormLayout(self)
        self.form.setContentsMargins(0, 0, 0, 0)
        self.form.setVerticalSpacing(8)
        build_fields(self, mode)
        if mode == "表格处理":
            self.fields["header_row"].valueChanged.connect(self.refresh)
            self.fields["filter_column"].textChanged.connect(self.refresh)
            self.fields["filter_op"].currentTextChanged.connect(self.refresh)
        if "rule" in self.fields:
            self.fields["rule"].currentTextChanged.connect(self.refresh)
        if mode == "图片转换":
            self.fields["format"].currentTextChanged.connect(self.refresh)
        self.refresh()

    def add(self, key, title, widget):
        """将控件同时登记在布局和参数映射内，便于按 key 读取或隐藏。"""
        widget.setAccessibleName(title)
        self.fields[key] = widget
        self.form.addRow(title, widget)
        return widget

    def line(self, key, title, value="", hint=""):
        """添加单行文本输入，提示文本不会作为实际参数提交。"""
        box = self.add(key, title, QLineEdit(value))
        box.setPlaceholderText(hint)

    def choice(self, key, title, values):
        """添加只能在给定值中选择的下拉框。"""
        self.add(key, title, QComboBox()).addItems(values)

    def number(self, key, title, low, high, value, zero_text=""):
        """添加有上下界的整数输入，用于限制序号或图片编码质量。"""
        box = self.add(key, title, QSpinBox())
        box.setRange(low, high)
        box.setValue(value)
        if zero_text:
            box.setSpecialValueText(zero_text)

    def check(self, key, title, value=False):
        """添加布尔选项，并保留可读的无障碍名称。"""
        box = QCheckBox(title)
        box.setChecked(value)
        self.fields[key] = box
        self.form.addRow(box)

    def refresh(self):
        """根据当前规则隐藏无关参数，避免旧输入干扰新规则。"""
        values = self.values()
        visible = visible_fields(self.mode, values)
        if self.mode == "批量重命名":
            hint = r"例如：订单_\1" if values["rule"] == "正则改名" else ""
            self.fields["replacement"].setPlaceholderText(hint)
        for key, widget in self.fields.items():
            self.form.setRowVisible(widget, key in visible)

    def values(self):
        """从所有控件读取原始值，供条件显示与参数快照共用。"""
        result = {}
        for key, widget in self.fields.items():
            if isinstance(widget, QComboBox):
                result[key] = widget.currentText()
            elif isinstance(widget, QCheckBox):
                result[key] = widget.isChecked()
            elif isinstance(widget, QSpinBox):
                result[key] = widget.value()
            else:
                result[key] = widget.text()
        return result

    def options(self):
        """只提交当前可见的字段，兼容旧图片任务的默认质量。"""
        return task_options(self.mode, self.values())
