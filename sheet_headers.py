"""表头识别。合并区域使用 Excel 的行列号（从 1 起，两端均包含）。"""

from collections import Counter
from itertools import chain, islice
from xml.etree.ElementTree import iterparse

from openpyxl.utils.cell import range_boundaries


HEADER_LOOKAHEAD = 32
MAX_HEADER_DEPTH = 8


def worksheet_merges(sheet, limit):
    """流式读取表头附近的合并坐标；openpyxl 的只读工作表不提供 merged_cells。"""
    ranges = []
    # 复用只读工作表自己的 XML 流，可处理重命名、调整顺序的工作表。
    with sheet._get_source() as stream:
        for _, node in iterparse(stream, events=("end",)):
            if node.tag.rsplit("}", 1)[-1] == "mergeCell":
                left, top, right, bottom = range_boundaries(node.attrib["ref"])
                if top <= limit:
                    ranges.append((top, bottom, left, right))
            node.clear()
    return ranges


def header_span(lines, merges):
    """在有限的前导行内定位表头；只在有合并坐标佐证时略过标题、展开多行。"""
    # read_header 已跳过开头的空行，探测不到横幅时第 0 项就是表头。
    start = 0
    for i, (number, values) in enumerate(lines):
        cells = [v for v in values if v is not None]
        if not cells:
            continue
        if len(cells) >= 2 and all(isinstance(v, str) for v in cells):
            start = i
            break
        banner = len(cells) == 1 and any(
            top == number and right > left and left <= len(values) and values[left - 1] is not None
            for top, _, left, right in merges
        )
        if not banner:
            break

    end = start
    # 纵向合并明确给出表头深度；横向分组下面至少要有两个文字子标题。
    while end + 1 < len(lines) and end - start + 1 < MAX_HEADER_DEPTH:
        number, following = lines[end + 1]
        active = [r for r in merges if lines[start][0] <= r[0] < number]
        vertical = any(bottom >= number and bottom < lines[start][0] + MAX_HEADER_DEPTH
                       for _, bottom, _, _ in active)
        filled = [v for v in following if v is not None]
        subheaders = filled and all(isinstance(v, str) for v in filled) and any(
            bottom == number - 1 and right > left
            and sum(v is not None for v in following[left - 1:right]) >= 2
            for _, bottom, left, right in active
        )
        if not (vertical or subheaders):
            break
        end += 1
    return start, end


def column_names(lines, merges):
    """展开表头合并格，优先用最下级列名。重名时用上级名称区分两列。"""
    first, last = lines[0][0], lines[-1][0]
    grid = {n: {i: str(v) for i, v in enumerate(values, 1) if v is not None}
            for n, values in lines}
    for top, bottom, left, right in merges:
        if first <= top <= last and left in grid[top]:
            label = grid[top][left]
            for n in range(top, min(last, bottom) + 1):
                for col in range(left, right + 1):
                    grid[n][col] = label
    width = max((max(row, default=0) for row in grid.values()), default=0)
    paths = []
    for col in range(1, width + 1):
        path = []
        for row in grid.values():
            value = row.get(col)
            if value is not None and (not path or path[-1] != value):
                path.append(value)
        paths.append(path)
    counts = Counter(path[-1] for path in paths if path)
    names = []
    for path in paths:
        if not path:
            names.append(None)
        elif counts[path[-1]] > 1:
            names.append(" / ".join(path))
        else:
            names.append(path[-1])
    return names


def read_header(rows, merges, first, depth):
    """返回列名、实际表头起止行和其余数据流；手选行号使用原表编号。"""
    if not 0 <= first <= 1_048_576 or not 1 <= depth <= MAX_HEADER_DEPTH:
        raise ValueError("表头起始行或层数超出范围")
    if first:
        block = list(islice(rows, first - 1, first - 1 + depth))
        if len(block) != depth or not any(values for _, values in block):
            raise ValueError(f"指定的第 {first} 行起、共 {depth} 行表头不存在或为空")
        return column_names(block, merges), (first, block[-1][0]), rows
    # 前置空行不占探测窗口，普通单列文件也保留第一条非空记录作表头。
    head = next(((n, values) for n, values in rows if values), None)
    if head is None:
        return [], None, rows
    lines = [head, *islice(rows, HEADER_LOOKAHEAD - 1)]
    start, end = header_span(lines, merges)
    block = lines[start:end + 1]
    return column_names(block, merges), (block[0][0], block[-1][0]), chain(lines[end + 1:], rows)
