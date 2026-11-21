"""工作台的规则字段、路径校验和结果呈现逻辑。"""

from pathlib import Path


MODES = ("批量重命名", "文本替换", "图片转换", "文件整理", "表格处理", "一键归档")
STATUS = {"ok": "成功", "skipped": "跳过", "error": "失败"}
METRICS = {"ok": "成功文件", "skipped": "跳过文件", "error": "失败文件", "rows": "输出行数"}
HEADER_IMAGE = Path(__file__).resolve().parent / "assets" / "workbench-header.png"


def build_fields(form, mode):
    """调用表单的控件方法创建规则，集中声明字段名、范围和默认值。"""
    if mode == "批量重命名":
        form.choice("rule", "命名方式", ["模板序号", "查找替换", "前后缀", "正则改名"])
        form.line("template", "文件名模板", "{name}_{n:03d}")
        form.number("start", "起始序号", 1, 999999, 1)
        form.line("find", "查找文字")
        form.line("replacement", "替换为")
        form.check("case", "区分大小写", True)
        form.line("prefix", "前缀")
        form.line("suffix", "后缀")
        form.line("extensions", "后缀筛选", hint="如 jpg,png；留空表示全部")
        form.check("in_place", "直接更改原文件名（可撤销本次）")
    elif mode == "文本替换":
        form.line("find", "查找内容")
        form.line("replacement", "替换为", hint="留空可删除匹配内容")
        form.choice("encoding", "文本编码", ["自动", "utf-8", "utf-8-sig", "utf-16-le", "utf-16-be",
                                         "gb18030", "gbk", "gb2312", "ascii"])
        form.check("regex", "使用正则表达式")
        form.check("case", "区分大小写", True)
    elif mode in ("图片转换", "一键归档"):
        form.choice("format", "图片格式", ["JPEG", "PNG", "WEBP", "BMP", "TIFF"] if mode == "图片转换" else ["JPEG", "PNG"])
        if mode == "图片转换":
            form.number("quality", "编码质量", 1, 100, 90)
            form.choice("resize_mode", "尺寸方式", ["指定宽高", "等比适应宽高", "限制最长边"])
            form.number("width", "目标宽度（像素）", 0, 65535, 0, "自动")
            form.number("height", "目标高度（像素）", 0, 65535, 0, "自动")
            form.number("max_edge", "最长边（像素）", 0, 65535, 0, "保留原尺寸")
        form.check("sharpen", "轻度锐化图片边缘")
    elif mode == "文件整理":
        form.choice("rule", "整理方式", ["按类型", "按修改月份", "查找重复", "文件清单"])
    elif mode == "表格处理":
        form.choice("rule", "操作", ["按列合并", "按键去重", "分组汇总", "按值拆分", "CSV 转 Excel", "Excel 转 CSV", "工作表目录", "列质量报告"])
        form.line("sheet_pattern", "工作表名", hint="留空处理全部；如 *岗位*，仅选 Excel 表")
        form.number("header_row", "表头起始行", 0, 1048576, 0, "自动识别")
        form.number("header_depth", "表头占用行数", 1, 8, 1)
        form.line("keys", "分组 / 去重列", hint="如 客户,订单号")
        form.line("amount", "数值列", hint="如 金额")
        form.line("filter_column", "筛选列", hint="可选，如 状态")
        form.choice("filter_op", "筛选条件", ["精确等于", "包含文字", "正则匹配", "为空", "非空"])
        form.line("filter_value", "匹配值", hint="输入文字或正则表达式")
        form.line("columns", "保留业务列", hint="用逗号分隔，留空保留全部")


def visible_fields(mode, values):
    """决定参与任务的字段；隐藏字段在编辑器中保留原值。"""
    if mode == "批量重命名":
        rule = values["rule"]
        show = {"模板序号": ["template", "start"], "查找替换": ["find", "replacement"],
                "前后缀": ["prefix", "suffix"], "正则改名": ["find", "replacement", "case"]}
        visible = {"rule", "extensions", "in_place", *show[rule]}
    elif mode == "表格处理":
        rule = values["rule"]
        visible = {"rule"}
        if rule != "CSV 转 Excel":
            visible.add("sheet_pattern")
        if rule != "工作表目录":
            visible.add("header_row")
            if str(values["header_row"]).strip() not in ("", "0"):
                visible.add("header_depth")
        if rule not in ("CSV 转 Excel", "Excel 转 CSV", "工作表目录"):
            visible.add("filter_column")
            if values["filter_column"].strip():
                visible.add("filter_op")
                if values["filter_op"] not in ("为空", "非空"):
                    visible.add("filter_value")
            visible.add("amount" if rule == "分组汇总" else "columns")
            if rule in ("按键去重", "按值拆分", "分组汇总"):
                visible.add("keys")
    elif mode == "图片转换":
        visible = {"format", "sharpen", "resize_mode"}
        if values["resize_mode"] == "限制最长边":
            visible.add("max_edge")
        else:
            visible.update(("width", "height"))
        if values["format"] in ("JPEG", "WEBP"):
            visible.add("quality")
    else:
        return set(values)
    return visible


def task_options(mode, values):
    """只提交可见字段，并为旧图片任务保留共用的默认编码质量。"""
    result = {"mode": mode, **{key: values[key] for key in visible_fields(mode, values)}}
    if mode == "图片转换":
        result.setdefault("quality", 90)
    return result


def batch_paths(source, output):
    """验证输入输出，并解析用户留空时采用的默认批次目录。"""
    source, output = source.strip(), output.strip()
    root = Path(source).expanduser()
    dest = Path(output).expanduser() if output else root / "整理结果"
    if not source or not root.is_dir():
        raise ValueError("请选择存在的输入文件夹。")
    if dest.resolve() == root.resolve() or (dest.exists() and not dest.is_dir()):
        raise ValueError("请将输出位置设为另一个文件夹。")
    return root, dest


def report_view(report, preview):
    """构造明细行和可选样本；界面标签只写入副本，导出的原始清单保持可重放。"""
    rows, samples = [], []
    for item in report["files"]:
        row = dict(item)
        if preview and row["status"] == "ok":
            row["label"] = "可处理"
        rows.append(row)
        if "sample_columns" in item:
            samples.append(("替换明细 · " + item["source"], item))
    for product in report.get("products", []):
        detail = product.get("detail") or f"{product['rows']} 行"
        rows.append(dict(source="输出预览" if preview else "批次产物", status="ok",
                         label="产物", output=product["output"], detail=detail))
        if "sample_columns" in product:
            samples.append(("数据样本 · " + product["output"], product))
    return rows, samples


def filter_records(records, status, members=()):
    """保留原顺序筛选文件，归档模式也允许单独查看 ZIP 的所有成员。"""
    if status == "members":
        return [dict(source=name, status="ok", label="成员", output="整理结果.zip") for name in members]
    return [item for item in records if item["status"] == status] if status else records


def record_detail(item):
    """生成第三列的说明，兼容含输出文件、行数、失败原因的不同清单记录。"""
    if item.get("detail"):
        return " · ".join(v for v in (item.get("output"), item["detail"]) if v)
    if "rows" in item:
        return f"{item['output']} · {item['rows']} 行"
    return item.get("output") or item.get("reason", "")
