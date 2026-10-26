"""
markdown_to_blocks.py —— 把 Markdown 文本转成飞书 docx 块(blocks)结构
====================================================================
用于「写」飞书文档：解析 Markdown，产出可直接喂给飞书
`/docx/v1/documents/{doc}/blocks/{block}/descendant` 接口的块结构。

为什么用 descendant 接口而不是 children 接口：
  飞书 children 接口不支持嵌套块，而表格是 table→table_cell→text 三层嵌套，
  必须用 descendant 接口一次性提交带层级的块。

输出设计（便于分批）：
  convert_markdown(md) -> List[Unit]
  每个 Unit = {"root_id": str, "blocks": [block, ...]}
    - 简单块（标题/正文/列表/代码/引用/分割线）的 Unit 只含 1 个 block；
    - 表格 Unit 含 1 个 table + N 个 cell + N 个 cell 内 text 块。
  每个 block 形如：
    {"block_id": "b3", "block_type": 31, "table": {...}, "children": ["b4","b5"]}
  顶层顺序由 Unit 顺序决定；调用方按 50 块上限分批提交。

支持的 Markdown 语法：
  # ~ #########        -> heading1 ~ heading9 (block_type 3~11)
  普通段落             -> text (block_type 2)，含行内 **粗体** *斜体* ~~删除~~ `代码` [文字](url)
  - / * / +           -> bullet 无序列表 (block_type 12)
  1. 2. 3.            -> ordered 有序列表 (block_type 13)
  ```lang ... ```     -> code 代码块 (block_type 14)
  >                   -> quote 引用 (block_type 15)
  --- / *** / ___     -> divider 分割线 (block_type 22)
  | a | b |           -> table 表格 (block_type 31 + 32)

本期不处理：图片（需 Drive 上传权限）、todo、callout、嵌套列表层级。
"""
import re
from itertools import count

# ---------------- 飞书 block_type 枚举 ----------------
BT_TEXT = 2
BT_HEADING_BASE = 2  # heading{n} 的 block_type = 2 + n（heading1=3 ... heading9=11）
BT_BULLET = 12
BT_ORDERED = 13
BT_CODE = 14
BT_QUOTE = 15
BT_CALLOUT = 19
BT_DIVIDER = 22
BT_TABLE = 31
BT_TABLE_CELL = 32

# heading{n} 的字段名
HEADING_KEY = {n: f"heading{n}" for n in range(1, 10)}

# 代码语言 -> 飞书 language 枚举（未命中用 1=PlainText，保证不报错）
CODE_LANG = {
    "plaintext": 1, "text": 1, "abap": 2, "bash": 4, "shell": 4, "sh": 4,
    "csharp": 5, "cs": 5, "cpp": 6, "c++": 6, "c": 7, "css": 8, "go": 22,
    "golang": 22, "html": 24, "java": 27, "javascript": 28, "js": 28,
    "json": 31, "kotlin": 32, "lua": 33, "markdown": 36, "md": 36,
    "php": 42, "python": 49, "py": 49, "ruby": 51, "rust": 52, "scala": 53,
    "sql": 54, "swift": 55, "typescript": 58, "ts": 58, "xml": 60, "yaml": 61,
    "yml": 61,
}


def _lang_code(lang: str) -> int:
    return CODE_LANG.get((lang or "").strip().lower(), 1)


# ---------------- 行内样式解析 ----------------
# 顺序很重要：先匹配三/双标记，再单标记。链接和行内代码优先级最高。
# <u>下划线</u> 支持读取端生成的下划线标签（写入时退化为普通文本，飞书写入不支持下划线）。
_INLINE_RE = re.compile(
    r"(?P<code>`[^`]+`)"
    r"|(?P<link>\[[^\]]*\]\([^)]+\))"
    r"|(?P<underline><u>(?P<u_content>[^<]*)</u>)"
    r"|(?P<bold>\*\*[^*]+\*\*|__[^_]+__)"
    r"|(?P<strike>~~[^~]+~~)"
    r"|(?P<italic>\*[^*]+\*|_[^_]+_)"
)


def parse_inline(text: str) -> list:
    """把一段文字解析成 text_run elements 列表，识别行内样式。"""
    elements = []
    pos = 0
    for m in _INLINE_RE.finditer(text):
        if m.start() > pos:
            elements.append(_text_run(text[pos:m.start()]))
        kind = m.lastgroup
        raw = m.group()
        if kind == "code":
            elements.append(_text_run(raw[1:-1], inline_code=True))
        elif kind == "link":
            lm = re.match(r"\[([^\]]*)\]\(([^)]+)\)", raw)
            label, url = lm.group(1), lm.group(2)
            elements.append(_text_run(label or url, link=url))
        elif kind == "underline":
            # 下划线标签：飞书写入不支持下划线，退化为普通文本
            elements.append(_text_run(m.group("u_content")))
        elif kind == "bold":
            elements.append(_text_run(raw[2:-2], bold=True))
        elif kind == "strike":
            elements.append(_text_run(raw[2:-2], strikethrough=True))
        elif kind == "italic":
            elements.append(_text_run(raw[1:-1], italic=True))
        pos = m.end()
    if pos < len(text):
        elements.append(_text_run(text[pos:]))
    if not elements:
        elements.append(_text_run(""))
    return elements


def _text_run(content: str, bold=False, italic=False, strikethrough=False,
              inline_code=False, link: str = None) -> dict:
    style = {}
    if bold:
        style["bold"] = True
    if italic:
        style["italic"] = True
    if strikethrough:
        style["strikethrough"] = True
    if inline_code:
        style["inline_code"] = True
    run = {"content": content}
    if style:
        run["text_element_style"] = style
    if link:
        # 飞书要求 link.url 做 URL 编码
        from urllib.parse import quote
        run.setdefault("text_element_style", {})["link"] = {"url": quote(link, safe="")}
    return {"text_run": run}


# ---------------- 块构造 ----------------
class _IdGen:
    def __init__(self):
        self._c = count(1)

    def next(self) -> str:
        return f"b{next(self._c)}"


def _simple_unit(idg, block_type, field, elements, extra_field=None):
    """构造只含 1 个块的 Unit（标题/正文/列表/引用等）。"""
    bid = idg.next()
    block = {"block_id": bid, "block_type": block_type}
    payload = {"elements": elements}
    if extra_field:
        payload.update(extra_field)
    block[field] = payload
    return {"root_id": bid, "blocks": [block]}


def _divider_unit(idg):
    bid = idg.next()
    return {"root_id": bid, "blocks": [{"block_id": bid, "block_type": BT_DIVIDER, "divider": {}}]}


def _callout_unit(idg, child_units, bg_color=3, border_color=3, emoji_id="star"):
    """构造 Callout 容器块（block_type=19）+ 子块。
    child_units: 内部内容对应的 Unit 列表（每个 Unit 的块都作为 callout 的后代）。
    结构：callout_block → children=[child_root_ids]，descendants 包含所有后代块。"""
    callout_id = idg.next()
    child_root_ids = []
    all_descendants = []
    for unit in child_units:
        child_root_ids.append(unit["root_id"])
        all_descendants.extend(unit["blocks"])
    callout_block = {
        "block_id": callout_id,
        "block_type": BT_CALLOUT,
        "callout": {
            "background_color": bg_color,
            "border_color": border_color,
            "emoji_id": emoji_id,
        },
        "children": child_root_ids,
    }
    return {"root_id": callout_id, "blocks": [callout_block] + all_descendants}


def _rebase_units(idg, units):
    """把一组 Unit 中的所有 block_id 用当前 idg 重新分配，保持结构关系不变。
    用于递归解析后的 callout 子内容，避免 id 与外层冲突。"""
    rebased = []
    for unit in units:
        old_to_new = {}
        # 第一遍：为所有 block 分配新 id
        for block in unit["blocks"]:
            old_id = block["block_id"]
            new_id = idg.next()
            old_to_new[old_id] = new_id
        # 第二遍：替换 id 和 children 引用
        new_blocks = []
        for block in unit["blocks"]:
            new_block = dict(block)
            new_block["block_id"] = old_to_new[block["block_id"]]
            if "children" in new_block:
                new_block["children"] = [old_to_new.get(c, c) for c in new_block["children"]]
            new_blocks.append(new_block)
        new_root_id = old_to_new[unit["root_id"]]
        rebased.append({"root_id": new_root_id, "blocks": new_blocks})
    return rebased


def _code_unit(idg, code_text, lang):
    bid = idg.next()
    block = {
        "block_id": bid,
        "block_type": BT_CODE,
        "code": {
            "elements": [_text_run(code_text)],
            "style": {"language": _lang_code(lang), "wrap": True},
        },
    }
    return {"root_id": bid, "blocks": [block]}


def _cell_display_width(text: str) -> int:
    """估算单元格文字显示宽度：中文/全角算 2，其余算 1。取各行最长一行。
    行的切分同时支持 \\n 与 <br> 标记（单元格多行用 <br>）。"""
    width = 0
    lines = _CELL_BR_RE.split(text or "")
    for chunk in lines:
        for line in chunk.split("\n"):
            w = 0
            for ch in line:
                w += 2 if ord(ch) > 0x2E80 else 1
            width = max(width, w)
    return width


def _calc_column_widths(rows, n_cols, total=720, col_min=80, col_max=480):
    """按每列内容的最大显示宽度，按比例分配列宽(px)，并限制单列上下限。
    total 为表格目标总宽（飞书正文区约 720px）。返回 List[int]，长度=n_cols。"""
    # 每列取该列所有单元格的最大显示宽度作为"内容权重"
    col_weight = [1] * n_cols
    for r in rows:
        for ci in range(n_cols):
            cell = r[ci] if ci < len(r) else ""
            col_weight[ci] = max(col_weight[ci], _cell_display_width(cell))

    # 先按权重比例分配，再夹到 [col_min, col_max]
    wsum = sum(col_weight) or 1
    widths = [max(col_min, min(col_max, round(total * w / wsum))) for w in col_weight]
    return widths


# 单元格内换行标记：<br> <br/> <br /> 各种大小写
_CELL_BR_RE = re.compile(r"<br\s*/?>", re.IGNORECASE)


def _split_cell_lines(text: str) -> list:
    """把单元格文字按 <br> 系列标记切分为多行。无标记时返回单元素列表。
    去掉切分后产生的空行两端空白；全空则回退为 [""] 以保证单元格至少有一个 text 子块。"""
    parts = [seg.strip() for seg in _CELL_BR_RE.split(text or "")]
    parts = [p for p in parts if p != ""]
    return parts or [""]


def _cell_seg_block(idg, seg: str):
    """判断单元格内一段文字的块类型：无序列表 / 有序列表 / 普通文本。
    返回 (block_id, block_dict)。"""
    bid = idg.next()
    # 无序列表：以 - 或 * 或 + 开头（后跟空格）
    m_bullet = _BULLET_RE.match(seg)
    if m_bullet:
        return bid, {
            "block_id": bid,
            "block_type": BT_BULLET,
            "bullet": {"elements": parse_inline(m_bullet.group(1))},
        }
    # 有序列表：以 数字. 或 数字) 开头（后跟空格）
    m_ordered = _ORDERED_RE.match(seg)
    if m_ordered:
        return bid, {
            "block_id": bid,
            "block_type": BT_ORDERED,
            "ordered": {"elements": parse_inline(m_ordered.group(1))},
        }
    # 普通文本
    return bid, {
        "block_id": bid,
        "block_type": BT_TEXT,
        "text": {"elements": parse_inline(seg)},
    }


def _table_unit(idg, rows, has_header):
    """rows: List[List[str]]，每个单元格是原始 markdown 文字（会做行内解析）。
    返回的 Unit 附带 col_widths（List[int]），供写入后用 PATCH 接口逐列设置列宽
    （飞书创建表格时会忽略 column_width，必须事后 PATCH）。
    单元格内支持列表语法：- / * / + 开头识别为无序列表，数字. 开头识别为有序列表。"""
    n_rows = len(rows)
    n_cols = max(len(r) for r in rows)
    col_widths = _calc_column_widths(rows, n_cols)
    table_id = idg.next()
    all_blocks = []
    cell_ids = []
    for r in range(n_rows):
        for col_index in range(n_cols):
            cell_id = idg.next()
            raw = rows[r][col_index] if col_index < len(rows[r]) else ""
            # 单元格内换行：按 <br> 切分成多行，每行按语法识别块类型。
            segments = _split_cell_lines(raw.strip())
            child_ids = []
            cell_child_blocks = []
            for seg in segments:
                bid, block = _cell_seg_block(idg, seg)
                child_ids.append(bid)
                cell_child_blocks.append(block)
            cell_block = {
                "block_id": cell_id,
                "block_type": BT_TABLE_CELL,
                "table_cell": {},
                "children": child_ids,
            }
            all_blocks.append(cell_block)
            all_blocks.extend(cell_child_blocks)
            cell_ids.append(cell_id)
    table_block = {
        "block_id": table_id,
        "block_type": BT_TABLE,
        "table": {
            "property": {
                "row_size": n_rows,
                "column_size": n_cols,
                "header_row": bool(has_header),
            }
        },
        "children": cell_ids,
    }
    # descendant 要求：顶层块在前，后代随后；这里 table 在最前
    return {"root_id": table_id, "blocks": [table_block] + all_blocks,
            "is_table": True, "col_widths": col_widths}


# ---------------- 主解析 ----------------
_HEADING_RE = re.compile(r"^(#{1,9})\s+(.*)$")
_BULLET_RE = re.compile(r"^[-*+]\s+(.*)$")
_ORDERED_RE = re.compile(r"^\d+[.)]\s+(.*)$")
_QUOTE_RE = re.compile(r"^>\s?(.*)$")
_DIVIDER_RE = re.compile(r"^(-{3,}|\*{3,}|_{3,})$")
_FENCE_RE = re.compile(r"^```(.*)$")
_TABLE_SEP_RE = re.compile(r"^\s*\|?\s*:?-{1,}:?\s*(\|\s*:?-{1,}:?\s*)+\|?\s*$")

# Callout 块语法：:::callout{color=yellow emoji=bulb} ... :::
_CALLOUT_OPEN_RE = re.compile(
    r"^:::callout(?:\{([^}]*)\})?\s*$"
)
_CALLOUT_CLOSE_RE = re.compile(r"^:::\s*$")

CALLOUT_COLORS = {
    "red": 1, "orange": 2, "yellow": 3, "green": 4,
    "blue": 5, "purple": 6, "grey": 7, "gray": 7,
}


def _split_table_row(line: str) -> list:
    line = line.strip()
    if line.startswith("|"):
        line = line[1:]
    if line.endswith("|"):
        line = line[:-1]
    return [c.strip() for c in line.split("|")]


def convert_markdown(md: str) -> list:
    """把 Markdown 文本转成 Unit 列表（见模块 docstring）。
    会自动去除 <!-- [不可还原: ...] --> 形式的 HTML 注释备注（读取端生成的不可还原标记）。"""
    # 预处理：去除不可还原注释标记
    md = re.sub(r'<!--\s*\[不可还原:[^\]]*\]\s*-->', '', md)
    idg = _IdGen()
    units = []
    lines = md.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    i = 0
    n = len(lines)
    while i < n:
        line = lines[i]
        stripped = line.strip()

        # 跳过纯空行
        if not stripped:
            i += 1
            continue

        # Callout 块：:::callout{color=yellow emoji=bulb} ... :::
        m_callout = _CALLOUT_OPEN_RE.match(stripped)
        if m_callout:
            # 解析参数
            params_str = m_callout.group(1) or ""
            bg_color = 3  # 默认 yellow
            border_color = 3
            emoji_id = "star"
            for part in params_str.split():
                if "=" in part:
                    k, v = part.split("=", 1)
                    k = k.strip().lower()
                    v = v.strip().strip("'\"")
                    if k == "color":
                        bg_color = CALLOUT_COLORS.get(v, 3)
                        border_color = bg_color
                    elif k == "border":
                        border_color = CALLOUT_COLORS.get(v, bg_color)
                    elif k == "emoji":
                        emoji_id = v
            # 收集内部行直到 :::
            i += 1
            inner_lines = []
            while i < n and not _CALLOUT_CLOSE_RE.match(lines[i].strip()):
                inner_lines.append(lines[i])
                i += 1
            i += 1  # 跳过结束 :::
            # 递归解析内部内容为子 Unit
            inner_md = "\n".join(inner_lines)
            child_units = convert_markdown(inner_md)
            # 子 Unit 的 block_id 需要重新分配以避免冲突
            # 由于递归调用使用独立的 _IdGen，id 可能重复；改用当前 idg 重新生成
            child_units_rebased = _rebase_units(idg, child_units)
            units.append(_callout_unit(idg, child_units_rebased, bg_color, border_color, emoji_id))
            continue

        # 代码块（围栏）
        m = _FENCE_RE.match(stripped)
        if m:
            lang = m.group(1).strip()
            buf = []
            i += 1
            while i < n and not _FENCE_RE.match(lines[i].strip()):
                buf.append(lines[i])
                i += 1
            i += 1  # 跳过结束围栏
            units.append(_code_unit(idg, "\n".join(buf), lang))
            continue

        # 表格：当前行像表格行，且下一行是分隔行
        if "|" in line and i + 1 < n and _TABLE_SEP_RE.match(lines[i + 1]):
            header = _split_table_row(line)
            rows = [header]
            i += 2  # 跳过表头与分隔行
            while i < n and "|" in lines[i] and lines[i].strip():
                rows.append(_split_table_row(lines[i]))
                i += 1
            units.append(_table_unit(idg, rows, has_header=True))
            continue

        # 标题
        m = _HEADING_RE.match(stripped)
        if m:
            level = len(m.group(1))
            units.append(_simple_unit(
                idg, BT_HEADING_BASE + level, HEADING_KEY[level],
                parse_inline(m.group(2).strip())))
            i += 1
            continue

        # 分割线
        if _DIVIDER_RE.match(stripped):
            units.append(_divider_unit(idg))
            i += 1
            continue

        # 引用
        m = _QUOTE_RE.match(stripped)
        if m:
            units.append(_simple_unit(idg, BT_QUOTE, "quote", parse_inline(m.group(1).strip())))
            i += 1
            continue

        # 无序列表
        m = _BULLET_RE.match(stripped)
        if m:
            units.append(_simple_unit(idg, BT_BULLET, "bullet", parse_inline(m.group(1).strip())))
            i += 1
            continue

        # 有序列表
        m = _ORDERED_RE.match(stripped)
        if m:
            units.append(_simple_unit(idg, BT_ORDERED, "ordered", parse_inline(m.group(1).strip())))
            i += 1
            continue

        # 普通段落：合并连续非空、非特殊行
        para = [stripped]
        i += 1
        while i < n:
            nxt = lines[i].strip()
            if (not nxt or _HEADING_RE.match(nxt) or _BULLET_RE.match(nxt)
                    or _ORDERED_RE.match(nxt) or _QUOTE_RE.match(nxt)
                    or _DIVIDER_RE.match(nxt) or _FENCE_RE.match(nxt)
                    or ("|" in lines[i] and i + 1 < n and _TABLE_SEP_RE.match(lines[i + 1]))):
                break
            para.append(nxt)
            i += 1
        units.append(_simple_unit(idg, BT_TEXT, "text", parse_inline(" ".join(para))))

    return units


def batch_units(units: list, limit: int = 50) -> list:
    """把 Unit 列表按「每批总块数 ≤ limit」分批，返回 List[(children_id, descendants)]。
    单个 Unit 块数若超过 limit（极端大表格），单独成批（接口可能仍报错，由调用方兜底）。"""
    batches = []
    cur_children = []
    cur_desc = []
    cur_count = 0
    for u in units:
        size = len(u["blocks"])
        if cur_count + size > limit and cur_children:
            batches.append((cur_children, cur_desc))
            cur_children, cur_desc, cur_count = [], [], 0
        cur_children.append(u["root_id"])
        cur_desc.extend(u["blocks"])
        cur_count += size
    if cur_children:
        batches.append((cur_children, cur_desc))
    return batches


def collect_table_widths(units: list) -> list:
    """按文档中表格出现的先后顺序，返回每个表格的列宽列表 List[List[int]]。
    供写入后用 PATCH 接口逐表逐列设置列宽（飞书创建时忽略 column_width）。"""
    return [u["col_widths"] for u in units if u.get("is_table")]
