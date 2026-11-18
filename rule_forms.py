"""工作台的规则表单：选项名称、可见条件和业务参数在同一处维护。"""

from PyQt6.QtWidgets import QCheckBox, QComboBox, QFormLayout, QLineEdit, QSpinBox, QWidget


MODES = ("批量重命名", "文本替换", "图片转换", "文件整理", "表格处理", "一键归档")


class RuleForm(QWidget):
    def __init__(self, mode):
        """为一个标签创建独立控件；切换标签时保留已填写的规则。"""
        super().__init__()
        self.mode = mode
        self.fields = {}
        self.form = QFormLayout(self)
        self.form.setContentsMargins(0, 0, 0, 0)
        self.form.setVerticalSpacing(8)
        if mode == "批量重命名":
            self.choice("rule", "命名方式", ["模板序号", "查找替换", "前后缀"])
            self.line("template", "文件名模板", "{name}_{n:03d}")
            self.number("start", "起始序号", 1, 999999, 1)
            self.line("find", "查找文字")
            self.line("replacement", "替换为")
            self.line("prefix", "前缀")
            self.line("suffix", "后缀")
            self.line("extensions", "后缀筛选", hint="如 jpg,png；留空表示全部")
            self.check("in_place", "直接更改原文件名（可撤销本次）")
        elif mode == "文本替换":
            self.line("find", "查找内容")
            self.line("replacement", "替换为", hint="留空可删除匹配内容")
            self.choice("encoding", "文本编码", ["自动", "utf-8", "utf-8-sig", "gb18030", "gbk", "gb2312", "ascii"])
            self.check("regex", "使用正则表达式")
            self.check("case", "区分大小写", True)
        elif mode in ("图片转换", "一键归档"):
            self.choice("format", "图片格式", ["JPEG", "PNG", "WEBP", "BMP", "TIFF"] if mode == "图片转换" else ["JPEG", "PNG"])
            if mode == "图片转换":
                self.number("quality", "编码质量", 1, 100, 90)
            self.check("sharpen", "轻度锐化图片边缘")
        elif mode == "文件整理":
            self.choice("rule", "整理方式", ["按类型", "按修改月份", "查找重复", "文件清单"])
        elif mode == "表格处理":
            self.choice("rule", "操作", ["按列合并", "按键去重", "分组汇总", "按值拆分", "CSV 转 Excel", "Excel 转 CSV", "工作表目录"])
            self.line("sheet_pattern", "工作表名", hint="留空处理全部；如 *岗位*，仅选 Excel 表")
            self.number("header_row", "表头起始行", 0, 1048576, 0)
            self.fields["header_row"].setSpecialValueText("自动识别")
            self.number("header_depth", "表头占用行数", 1, 8, 1)
            self.fields["header_row"].valueChanged.connect(self.refresh)
            self.line("keys", "分组 / 去重列", hint="如 客户,订单号")
            self.line("amount", "数值列", hint="如 金额")
            self.line("filter_column", "筛选列", hint="可选，如 状态")
            self.choice("filter_op", "筛选条件", ["精确等于", "包含文字", "正则匹配", "为空", "非空"])
            self.line("filter_value", "匹配值", hint="输入文字或正则表达式")
            self.fields["filter_column"].textChanged.connect(self.refresh)
            self.fields["filter_op"].currentTextChanged.connect(self.refresh)
            self.line("columns", "保留业务列", hint="用逗号分隔，留空保留全部")
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

    def number(self, key, title, low, high, value):
        """添加有上下界的整数输入，用于限制序号或图片编码质量。"""
        box = self.add(key, title, QSpinBox())
        box.setRange(low, high)
        box.setValue(value)

    def check(self, key, title, value=False):
        """添加布尔选项，并保留可读的无障碍名称。"""
        box = QCheckBox(title)
        box.setChecked(value)
        self.fields[key] = box
        self.form.addRow(box)

    def refresh(self):
        """根据当前规则隐藏无关参数，避免旧输入干扰新规则。"""
        if self.mode == "批量重命名":
            rule = self.fields["rule"].currentText()
            show = {"模板序号": ["template", "start"], "查找替换": ["find", "replacement"], "前后缀": ["prefix", "suffix"]}
            visible = {"rule", "extensions", "in_place", *show[rule]}
        elif self.mode == "表格处理":
            rule = self.fields["rule"].currentText()
            visible = {"rule"}
            if rule != "CSV 转 Excel":
                visible.add("sheet_pattern")
            if rule != "工作表目录":
                visible.add("header_row")
                if self.fields["header_row"].value():
                    visible.add("header_depth")
            if rule not in ("CSV 转 Excel", "Excel 转 CSV", "工作表目录"):
                visible.add("filter_column")
                if self.fields["filter_column"].text().strip():
                    visible.add("filter_op")
                    if self.fields["filter_op"].currentText() not in ("为空", "非空"):
                        visible.add("filter_value")
                visible.add("amount" if rule == "分组汇总" else "columns")
                if rule != "按列合并":
                    visible.add("keys")
        elif self.mode == "图片转换":
            visible = {"format", "sharpen"}
            if self.fields["format"].currentText() in ("JPEG", "WEBP"):
                visible.add("quality")
        else:
            return
        for key, widget in self.fields.items():
            self.form.setRowVisible(widget, key in visible)

    def options(self):
        """将当前可用参数转换为普通 Python 字典，工作线程不再访问控件。"""
        result = {"mode": self.mode}
        for key, widget in self.fields.items():
            if not self.form.isRowVisible(widget):
                continue
            if isinstance(widget, QComboBox):
                result[key] = widget.currentText()
            elif isinstance(widget, QCheckBox):
                result[key] = widget.isChecked()
            elif isinstance(widget, QSpinBox):
                result[key] = widget.value()
            else:
                result[key] = widget.text()
        if self.mode == "图片转换":
            result.setdefault("quality", 90)
        return result
