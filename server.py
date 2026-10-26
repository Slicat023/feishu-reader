"""
feishu-mcp MCP Server
=====================
把飞书文档只读能力暴露为 MCP 工具，供 MCP 客户端原生调用。

启动方式（由 MCP 客户端配置自动拉起，用户无需手动运行）：
    python server.py

协议：MCP stdio 模式（stdin/stdout JSON-RPC）。
"""
import os
import sys
import time

# 确保 scripts/ 目录在 import 路径中
PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(PROJECT_ROOT, "scripts"))

from mcp.server.fastmcp import FastMCP

from feishu_client import FeishuClient, FeishuError, load_env
from blocks_to_markdown import blocks_to_markdown
from feishu_sheet import fetch_spreadsheet_as_markdown
from markdown_to_blocks import convert_markdown, batch_units, collect_table_widths
from markdown_to_html import markdown_to_html
from fetch_doc import (
    resolve_wiki,
    fetch_one,
    is_feishu_doc_link,
    sanitize_filename,
    _follow_referenced_docs,
)

# --------------- MCP Server 实例 ---------------
mcp = FastMCP(
    "feishu-mcp",
    instructions="飞书云文档只读工具：读取飞书 docx/wiki/sheet，返回 Markdown 正文",
)

DEFAULT_OUTPUT_DIR = os.path.join(PROJECT_ROOT, "output")


def _get_client() -> FeishuClient:
    """每次工具调用时构造客户端（复用 token 缓存文件）。"""
    return FeishuClient()


# --------------- 工具定义 ---------------


@mcp.tool()
def fetch_document(
    url: str,
    follow: bool = False,
    depth: int = 1,
    save: bool = False,
    images: bool = False,
    save_format: str = "md",
) -> dict:
    """读取飞书文档（docx/wiki），返回 Markdown 正文。

    Args:
        url: 飞书文档链接或 document_id。支持 /docx/、/wiki/ 链接。
        follow: 是否跟随抓取文中引用的飞书文档（作为附录拼接）。
        depth: 跟随抓取的层数，默认 1。仅 follow=True 时生效。
        save: 是否同时存盘到 output/ 目录。
        images: 是否下载文中图片到 output/images/<标题>/ 并在正文以 Markdown 图片引用。
                需 drive:file:download 权限；开启时会自动 save。
        save_format: 存盘格式，"md"(默认) 或 "html"。仅在 save/images 触发存盘时生效；
                     返回给调用方的 content 始终是 Markdown。

    Returns:
        dict: {title, content, links, saved_path?}
    """
    client = _get_client()
    fmt = save_format.lower() if save_format else "md"
    if fmt not in ("md", "html"):
        fmt = "md"
    try:
        obj_type, obj_token = resolve_wiki(client, url)

        if obj_type == "sheet":
            title, md = fetch_spreadsheet_as_markdown(client, obj_token)
            result = {"title": title, "content": md, "type": "sheet", "links": []}
            if save:
                path = _save_content(title or obj_token, md, fmt=fmt)
                result["saved_path"] = path
            return result

        if obj_type != "docx":
            return {"error": f"不支持的文档类型：{obj_type}（当前只支持 docx 和 sheet）"}

        # docx。开启 images 时需要 output_dir 供图片落盘（并强制存盘正文）
        need_output = save or images
        root_links = []
        _, title, md = fetch_one(
            client, url, collect_links=root_links,
            output_dir=DEFAULT_OUTPUT_DIR if need_output else None,
            save_images=images,
        )

        if follow:
            visited = {obj_token}
            appendix, followed_count = _follow_referenced_docs(
                client, root_links, max(1, depth), visited=visited
            )
            if appendix:
                md = md.rstrip() + "\n\n" + appendix

        links = [{"title": t, "url": u} for t, u in root_links]
        result = {"title": title, "content": md, "type": "docx", "links": links}
        if need_output:
            path = _save_content(title or obj_token, md, fmt=fmt)
            result["saved_path"] = path
        return result

    except FeishuError as e:
        return {"error": f"飞书错误: {e}"}
    except Exception as e:
        return {"error": f"{type(e).__name__}: {e}"}


@mcp.tool()
def fetch_sheet(
    url: str,
    max_rows: int = 5000,
    max_display_rows: int = 200,
) -> dict:
    """读取飞书电子表格，返回 Markdown 格式的表格内容。

    Args:
        url: 飞书电子表格链接（/sheets/ 或 /wiki/ 底层为 sheet 的链接）。
        max_rows: 每个工作表最多读取行数，默认 5000。
        max_display_rows: Markdown 表格最多显示行数，默认 200。

    Returns:
        dict: {title, content, type: "sheet"}
    """
    client = _get_client()
    try:
        obj_type, obj_token = resolve_wiki(client, url)
        if obj_type != "sheet":
            return {"error": f"该链接不是电子表格类型（实际为 {obj_type}）。请使用 fetch_document 读取文档。"}
        title, md = fetch_spreadsheet_as_markdown(
            client, obj_token, max_rows_per_sheet=max_rows, max_display_rows=max_display_rows
        )
        return {"title": title, "content": md, "type": "sheet"}
    except FeishuError as e:
        return {"error": f"飞书错误: {e}"}
    except Exception as e:
        return {"error": f"{type(e).__name__}: {e}"}


@mcp.tool()
def create_document(
    title: str,
    markdown: str,
    parent_node_token: str = "",
) -> dict:
    """【输出到飞书文档】在飞书 wiki 个人目录下创建新文档并写入 Markdown 内容。

    重要：仅当用户给出明确指令「输出到飞书文档」（或等价的"写到飞书/生成飞书文档"）时
    才可调用本工具。不要在用户只是让你写/总结内容时主动调用。

    会把 Markdown 转成飞书原生块（标题/正文/列表/代码/引用/分割线/表格，含
    粗体/斜体/删除线/行内代码/链接等行内样式），新文档作为目标 wiki 节点的子节点创建。

    Args:
        title: 新文档标题。
        markdown: 文档正文（Markdown 格式）。
        parent_node_token: 父 wiki 节点 token；留空则用 .env 中 FEISHU_WIKI_PARENT_NODE。

    Returns:
        dict: {title, url, document_id, node_token, blocks_written} 或 {error}
    """
    space_id = os.environ.get("FEISHU_WIKI_SPACE_ID", "").strip()
    parent = (parent_node_token or os.environ.get("FEISHU_WIKI_PARENT_NODE", "")).strip()
    if not space_id or not parent:
        load_env()  # 确保 .env 已加载后再读
        space_id = space_id or os.environ.get("FEISHU_WIKI_SPACE_ID", "").strip()
        parent = parent or os.environ.get("FEISHU_WIKI_PARENT_NODE", "").strip()
    if not space_id or not parent:
        return {"error": "未配置写入目标。请在 feishu-mcp/.env 中填写 "
                         "FEISHU_WIKI_SPACE_ID 与 FEISHU_WIKI_PARENT_NODE。"}
    if not (markdown or "").strip():
        return {"error": "markdown 内容为空，无需创建文档。"}

    client = _get_client()
    try:
        node = client.create_wiki_node(space_id, parent, title or "未命名文档")
        doc_id = node.get("obj_token")
        node_token = node.get("node_token")
        if not doc_id:
            return {"error": f"创建节点未返回 document_id：{node}"}

        units = convert_markdown(markdown)
        batches = batch_units(units, limit=50)
        written = 0
        for children_ids, descendants in batches:
            client.create_block_descendant(doc_id, doc_id, children_ids, descendants, index=-1)
            written += len(descendants)

        # 设置表格列宽：飞书创建表格时忽略 column_width，须写入后用 PATCH 逐列设置。
        # 注意：刚写入的表格在服务端可能尚未就绪，立即 PATCH 会失败，需重试。
        table_widths = collect_table_widths(units)
        width_set = 0
        width_failed = 0
        if table_widths:
            # 等待服务端块就绪后，重新拉取真实 table block_id
            table_ids = []
            for attempt in range(5):
                time.sleep(1.0)
                table_ids = client.list_table_block_ids(doc_id)
                if len(table_ids) >= len(table_widths):
                    break
            for tbl_id, widths in zip(table_ids, table_widths):
                for col_idx, w in enumerate(widths):
                    ok = False
                    for retry in range(3):
                        try:
                            client.set_table_column_width(doc_id, tbl_id, col_idx, int(w))
                            ok = True
                            break
                        except Exception:
                            time.sleep(0.6)  # 失败退避后重试（限频/未就绪）
                    if ok:
                        width_set += 1
                    else:
                        width_failed += 1
                    time.sleep(0.35)  # 飞书编辑限频约每秒 3 次，节流避免 429

        # wiki 链接域名是租户专属的，从 .env 的 FEISHU_WIKI_DOMAIN 读取，例如 your-tenant.feishu.cn
        wiki_domain = os.environ.get("FEISHU_WIKI_DOMAIN", "feishu.cn").strip()
        url = f"https://{wiki_domain}/wiki/{node_token}"
        return {
            "title": title,
            "url": url,
            "document_id": doc_id,
            "node_token": node_token,
            "blocks_written": written,
            "table_columns_sized": width_set,
            "table_columns_failed": width_failed,
        }
    except FeishuError as e:
        return {"error": f"飞书错误: {e}"}
    except Exception as e:
        return {"error": f"{type(e).__name__}: {e}"}


@mcp.tool()
def list_saved_docs() -> dict:
    """列出 output/ 目录中已缓存的飞书文档/表格文件。

    Returns:
        dict: {docs: [{filename, size_kb, path}]}
    """
    docs = []
    if os.path.isdir(DEFAULT_OUTPUT_DIR):
        for fname in sorted(os.listdir(DEFAULT_OUTPUT_DIR)):
            fpath = os.path.join(DEFAULT_OUTPUT_DIR, fname)
            if os.path.isfile(fpath):
                size_kb = round(os.path.getsize(fpath) / 1024, 1)
                docs.append({"filename": fname, "size_kb": size_kb, "path": fpath})
    return {"docs": docs}


@mcp.tool()
def read_saved_doc(filename: str) -> dict:
    """读取 output/ 目录中已缓存的文档内容。

    Args:
        filename: 文件名（如 "某文档标题.md"），来自 list_saved_docs 的结果。

    Returns:
        dict: {filename, content} 或 {error}
    """
    fpath = os.path.join(DEFAULT_OUTPUT_DIR, filename)
    if not os.path.isfile(fpath):
        return {"error": f"文件不存在: {filename}。请先用 list_saved_docs 查看可用文件。"}
    try:
        with open(fpath, "r", encoding="utf-8") as f:
            content = f.read()
        return {"filename": filename, "content": content}
    except Exception as e:
        return {"error": f"读取失败: {e}"}


# --------------- 辅助函数 ---------------


def _save_content(title: str, content: str, fmt: str = "md") -> str:
    """存盘到 output/ 目录，返回文件路径。
    fmt='md' 存 Markdown；fmt='html' 用 markdown_to_html 渲染后存 .html。"""
    os.makedirs(DEFAULT_OUTPUT_DIR, exist_ok=True)
    if fmt == "html":
        fname = sanitize_filename(title) + ".html"
        out = markdown_to_html(content, title=title)
    else:
        fname = sanitize_filename(title) + ".md"
        out = content
    fpath = os.path.join(DEFAULT_OUTPUT_DIR, fname)
    with open(fpath, "w", encoding="utf-8") as f:
        f.write(out)
    return fpath


# --------------- 入口 ---------------

if __name__ == "__main__":
    mcp.run()
