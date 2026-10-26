"""
飞书 docx 块(blocks) -> Markdown 文本
====================================
尽可能完整还原飞书文档的格式到 Markdown，包括行内样式（加粗/斜体/删除线/
行内代码/高亮/文字颜色）、块级结构（callout/todo/嵌套列表）以及不可还原元素
的注释标注。

飞书 docx 的块结构要点：
- 每个 block 有 block_id、parent_id、block_type、children(子块id列表)。
- 文本类块（正文/标题/引用/代码/列表项等）内容在对应字段的 elements[] 里，
  每个 element 是 {text_run:{content:"..."}} 之类，需把 content 拼起来。
- 表格 table 块本身只有结构信息(行列数)，单元格是它的 children(table_cell)，
  单元格里再挂文本块 -> 需要按 block_id 建索引递归取文本。

block_type 枚举（常用）：
  1 page(根/文档标题) 2 text 3~11 heading1~9 12 bullet 13 ordered 14 code
  15 quote 17 todo 19 callout 22 divider 27 image 28 file 30 iframe
  31 table 32 table_cell ...

无法还原的内容用 HTML 注释标注：<!-- [不可还原: 描述] -->
写入飞书时 markdown_to_blocks.py 会自动去除这些注释。
"""

# 标题级别映射：block_type -> markdown 井号数
HEADING_TYPES = {3: 1, 4: 2, 5: 3, 6: 4, 7: 5, 8: 6, 9: 6, 10: 6, 11: 6}

# block_type -> 该类型存放 text 内容的字段名
TEXT_FIELD_BY_TYPE = {
    2: "text",
    3: "heading1", 4: "heading2", 5: "heading3",
    6: "heading4", 7: "heading5", 8: "heading6",
    9: "heading7", 10: "heading8", 11: "heading9",
    12: "bullet",
    13: "ordered",
    14: "code",
    15: "quote",
    17: "todo",
    19: "callout",
}

# 飞书代码块语言枚举 -> 语言名（反向映射，用于读取时还原语言标记）
_CODE_LANG_REVERSE = {
    1: "", 2: "abap", 4: "bash", 5: "csharp", 6: "cpp", 7: "c", 8: "css",
    22: "go", 24: "html", 27: "java", 28: "javascript", 31: "json",
    32: "kotlin", 33: "lua", 36: "markdown", 42: "php", 49: "python",
    51: "ruby", 52: "rust", 53: "scala", 54: "sql", 55: "swift",
    58: "typescript", 60: "xml", 61: "yaml",
}

# 飞书文字颜色枚举 -> 颜色名（用于行内备注）
_FONT_COLOR_NAMES = {
    1: "红", 2: "橙", 3: "黄", 4: "绿", 5: "蓝", 6: "紫", 7: "灰",
}

# 飞书高亮背景色枚举 -> 颜色名
_BG_COLOR_NAMES = {
    1: "浅红底", 2: "浅橙底", 3: "浅黄底", 4: "浅绿底",
    5: "浅蓝底", 6: "浅紫底", 7: "浅灰底",
}

# Callout 背景色枚举 -> 颜色名（与 markdown_to_blocks 的 CALLOUT_COLORS 反向）
_CALLOUT_COLOR_NAMES = {
    1: "red", 2: "orange", 3: "yellow", 4: "green",
    5: "blue", 6: "purple", 7: "grey",
}


def _decode_link_url(url):
    """飞书把链接里的 & 等转义为 \\u0026/百分号编码，这里还原成可点击的正常 URL。"""
    if not url:
        return ""
    try:
        from urllib.parse import unquote
        return unquote(url)
    except Exception:
        return url


def _extract_text(block: dict, field: str, links: list = None, user_names: dict = None) -> str:
    """从某个文本类字段里把 elements 的文本拼出来，完整还原行内样式。
    支持：加粗(**) 斜体(*) 删除线(~~) 行内代码(`) 超链接 @用户 内联文档
    高亮背景色/文字颜色用 HTML 注释备注（Markdown 无原生语法）。
    遇到 @用户(mention_user)，若 user_names 提供了 open_id->姓名 映射则渲染成 @姓名。"""
    node = block.get(field) or {}
    elements = node.get("elements", []) or []
    parts = []
    for el in elements:
        tr = el.get("text_run")
        if tr and tr.get("content") is not None:
            content = tr["content"]
            style = tr.get("text_element_style") or {}

            # 超链接优先
            link = (style.get("link") or {}).get("url")
            url = _decode_link_url(link) if link else ""

            # 行内代码：直接用反引号包裹，不叠加其他样式
            if style.get("inline_code"):
                segment = f"`{content}`"
            else:
                segment = content
                # 加粗
                if style.get("bold"):
                    segment = f"**{segment}**"
                # 斜体
                if style.get("italic"):
                    segment = f"*{segment}*"
                # 删除线
                if style.get("strikethrough"):
                    segment = f"~~{segment}~~"
                # 下划线 —— Markdown 无原生语法，用 HTML <u> 标签
                if style.get("underline"):
                    segment = f"<u>{segment}</u>"

            # 超链接包裹
            if url:
                segment = f"[{segment}]({url})"
                if links is not None:
                    links.append((content.strip() or url, url))

            # 高亮背景色 / 文字颜色 —— 无法用 Markdown 还原，用注释备注
            bg_color = style.get("background_color")
            font_color = style.get("font_color")
            annotations = []
            if bg_color and bg_color in _BG_COLOR_NAMES:
                annotations.append(_BG_COLOR_NAMES[bg_color])
            if font_color and font_color in _FONT_COLOR_NAMES:
                annotations.append(f"{_FONT_COLOR_NAMES[font_color]}色字")
            if annotations:
                segment = f"<!-- [不可还原: {'+'.join(annotations)}] -->{segment}"

            parts.append(segment)
            continue
        # 内联文档/文件提及：保留为 Markdown 链接，带出引用文档地址
        md = el.get("mention_doc")
        if md:
            title = md.get("title") or "（引用文档）"
            url = _decode_link_url(md.get("url"))
            parts.append(f"[{title}]({url})" if url else title)
            if url and links is not None:
                links.append((title.strip() or url, url))
            continue
        # @用户：用 open_id 反查到的姓名渲染成 @姓名，查不到则用占位
        mu = el.get("mention_user")
        if mu:
            open_id = mu.get("user_id") or ""
            name = ""
            if user_names and open_id in user_names:
                name = user_names[open_id]
            if not name:
                name = mu.get("content") or mu.get("title") or "某用户"
            parts.append(f"@{name}")
            continue
        # 公式
        eq = el.get("equation")
        if eq:
            text = eq.get("content") or ""
            if text:
                # 用 $...$ 包裹（LaTeX 行内公式通用写法）
                parts.append(f"$`{text}`$")
            continue
        # 未知元素类型 —— 注释标注
        parts.append("<!-- [不可还原: 未知行内元素] -->")
    return "".join(parts)


def collect_mention_user_ids(blocks: list) -> list:
    """扫描所有块，收集 @用户(mention_user) 的 open_id 列表（去重）。"""
    ids = []
    seen = set()
    for b in blocks:
        for field in TEXT_FIELD_BY_TYPE.values():
            node = b.get(field)
            if not isinstance(node, dict):
                continue
            for el in node.get("elements", []) or []:
                mu = el.get("mention_user")
                if mu:
                    oid = mu.get("user_id")
                    if oid and oid not in seen:
                        seen.add(oid)
                        ids.append(oid)
    return ids


def blocks_to_markdown(blocks: list, doc_title: str = "", links_out: list = None,
                       user_names: dict = None, image_saver=None) -> str:
    """把扁平 block 列表按文档顺序转成 Markdown，并在文末附"引用链接"清单。
    若传入 links_out（list），会把收集到的 (标题, URL) 按出现顺序追加进去，供调用方跟随抓取。
    若传入 user_names（dict: open_id->姓名），@用户会渲染成 @姓名。
    若传入 image_saver（callable: token -> 本地相对路径 或 None），图片会被下载并渲染为
    Markdown 图片语法 ![](路径)；返回 None 时退化为注释标注。"""
    by_id = {b.get("block_id"): b for b in blocks}

    # 找根块(page)：block_type==1，children 即文档顶层顺序
    root = next((b for b in blocks if b.get("block_type") == 1), None)

    lines = []
    links = links_out if links_out is not None else []  # 全文收集到的 (标题, URL)
    if doc_title:
        lines.append(f"# {doc_title}")
        lines.append("")

    ordered_counter = {}  # parent_id -> 当前有序列表序号

    def render_block(block_id, depth=0):
        block = by_id.get(block_id)
        if not block:
            return
        btype = block.get("block_type")

        # ---- 表格：单独处理 ----
        if btype == 31:
            _render_table(block, by_id, lines, links, user_names)
            lines.append("")
            return

        # ---- 图片(27)：下载并保存，渲染为 Markdown 图片 ----
        if btype == 27:
            token = (block.get("image") or {}).get("token") or ""
            saved_path = None
            if image_saver and token:
                try:
                    saved_path = image_saver(token)
                except Exception:
                    saved_path = None
            if saved_path:
                lines.append(f"![图片]({saved_path})")
            else:
                lines.append(f"<!-- [图片未下载, token={token}] -->")
            lines.append("")
            return

        # ---- 文件(28)：不可还原，注释标注 ----
        if btype == 28:
            file_info = block.get("file") or {}
            name = file_info.get("name") or file_info.get("token") or "未知文件"
            lines.append(f"<!-- [不可还原: 文件附件, name={name}] -->")
            lines.append("")
            return

        # ---- iframe/嵌入(30)：不可还原，注释标注 ----
        if btype == 30:
            iframe_info = block.get("iframe") or {}
            url = iframe_info.get("url") or ""
            if url:
                lines.append(f"<!-- [不可还原: 嵌入内容] -->[嵌入链接]({url})")
            else:
                lines.append("<!-- [不可还原: 嵌入内容(iframe)] -->")
            lines.append("")
            return

        # ---- 多维表/画板等(33 bitable, 34 board, etc.)：不可还原 ----
        if btype in (33, 34, 35, 36, 37, 38, 39, 40):
            type_names = {33: "多维表格", 34: "画板", 35: "流程图",
                          36: "UML图", 37: "思维导图", 38: "问卷",
                          39: "OKR", 40: "Jira"}
            name = type_names.get(btype, f"未知块类型({btype})")
            lines.append(f"<!-- [不可还原: {name}] -->")
            lines.append("")
            return

        # ---- 文本类 ----
        field = TEXT_FIELD_BY_TYPE.get(btype)
        if field:
            text = _extract_text(block, field, links, user_names)
            indent = "  " * depth

            if btype in HEADING_TYPES:
                lines.append(f"{'#' * HEADING_TYPES[btype]} {text}")
                lines.append("")
            elif btype == 12:  # 无序列表
                lines.append(f"{indent}- {text}")
            elif btype == 13:  # 有序列表
                pid = block.get("parent_id")
                ordered_counter[pid] = ordered_counter.get(pid, 0) + 1
                lines.append(f"{indent}{ordered_counter[pid]}. {text}")
            elif btype == 14:  # 代码块
                lang_raw = (block.get("code") or {}).get("style", {}).get("language", "")
                # 语言可能是数字枚举或字符串
                if isinstance(lang_raw, int):
                    lang = _CODE_LANG_REVERSE.get(lang_raw, "")
                else:
                    lang = lang_raw if isinstance(lang_raw, str) else ""
                lines.append(f"```{lang}")
                lines.append(text)
                lines.append("```")
                lines.append("")
            elif btype == 15:  # 引用
                lines.append(f"> {text}")
                lines.append("")
            elif btype == 17:  # 待办
                checked = (block.get("todo") or {}).get("style", {}).get("done", False)
                lines.append(f"{indent}- [{'x' if checked else ' '}] {text}")
            elif btype == 19:  # callout 高亮块
                # 获取 callout 属性并还原为 :::callout 语法
                callout_info = block.get("callout") or {}
                bg_color = callout_info.get("background_color", 3)
                emoji_id = callout_info.get("emoji_id", "star")
                color_name = _CALLOUT_COLOR_NAMES.get(bg_color, "yellow")
                # 开始标记
                lines.append(f":::callout{{color={color_name} emoji={emoji_id}}}")
                # callout 内文本（自身的 text 内容）
                if text:
                    lines.append(text)
                # callout 的子块递归渲染
                for child_id in block.get("children", []) or []:
                    render_block(child_id, 0)
                lines.append(":::")
                lines.append("")
                return  # 子块已处理，不再走底部的递归
            else:  # 普通正文
                if text:
                    lines.append(text)
                    lines.append("")
        elif btype == 22:  # 分割线
            lines.append("---")
            lines.append("")
        elif btype not in (1, 31, 32):
            # 未知块类型，注释标注
            lines.append(f"<!-- [不可还原: 未知块类型({btype})] -->")
            lines.append("")

        # ---- 递归子块（表格和 callout 已自行处理子块，这里跳过）----
        if btype not in (31, 19):
            for child_id in block.get("children", []) or []:
                render_block(child_id, depth + 1 if btype in (12, 13, 17) else depth)

    if root:
        for child_id in root.get("children", []) or []:
            render_block(child_id, 0)
    else:
        # 没有 page 根块时兜底：按出现顺序平铺
        for b in blocks:
            render_block(b.get("block_id"), 0)

    # 文末附"引用链接"清单（去重，保持首次出现顺序）
    if links:
        seen = set()
        lines.append("")
        lines.append("---")
        lines.append("")
        lines.append("## 引用链接")
        lines.append("")
        idx = 1
        for title, url in links:
            if url in seen:
                continue
            seen.add(url)
            lines.append(f"{idx}. [{title}]({url})")
            idx += 1

    # 收尾：压掉多余空行
    md = "\n".join(lines)
    while "\n\n\n" in md:
        md = md.replace("\n\n\n", "\n\n")
    return md.strip() + "\n"


def _render_table(table_block: dict, by_id: dict, lines: list, links: list = None, user_names: dict = None):
    """把 table 块渲染成 Markdown 表格。
    table.property 里有 row_size/column_size；children 是按行优先排列的 table_cell。"""
    prop = (table_block.get("table") or {}).get("property", {}) or {}
    rows = prop.get("row_size", 0)
    cols = prop.get("column_size", 0)
    cell_ids = table_block.get("children", []) or []
    if not rows or not cols or not cell_ids:
        return

    # 取每个单元格的纯文本（单元格里可能挂多个文本子块，拼接）
    def cell_text(cell_id):
        cell = by_id.get(cell_id)
        if not cell:
            return ""
        texts = []
        for tid in cell.get("children", []) or []:
            tb = by_id.get(tid)
            if not tb:
                continue
            field = TEXT_FIELD_BY_TYPE.get(tb.get("block_type"))
            if field:
                t = _extract_text(tb, field, links, user_names)
                if t:
                    texts.append(t)
        # 单元格内换行用 <br> 保表格不破行
        return "<br>".join(texts)

    # 按 rows x cols 切分
    grid = []
    for r in range(rows):
        row_cells = cell_ids[r * cols:(r + 1) * cols]
        grid.append([cell_text(cid).replace("|", "\\|") for cid in row_cells])

    if not grid:
        return

    # 第一行当表头
    header = grid[0]
    lines.append("| " + " | ".join(header) + " |")
    lines.append("| " + " | ".join("---" for _ in header) + " |")
    for row in grid[1:]:
        # 补齐列数
        while len(row) < cols:
            row.append("")
        lines.append("| " + " | ".join(row) + " |")
