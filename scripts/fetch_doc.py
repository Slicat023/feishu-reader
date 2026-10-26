"""
fetch_doc.py —— 输入飞书文档链接，输出 Markdown 正文（供 AI 工具阅读）
=================================================================
用法（cwd = feishu-doc/）：
  python scripts/fetch_doc.py <文档链接或document_id>            # 仅打印到终端
  python scripts/fetch_doc.py <链接> --save                     # 存到 output/<文档标题>.md
  python scripts/fetch_doc.py <链接> --out output/x.md          # 存到指定路径
  python scripts/fetch_doc.py <链接> --out output/              # 存到目录，文件名用文档标题
  python scripts/fetch_doc.py <链接> --follow                   # 顺带读取文中引用的飞书文档内容
  python scripts/fetch_doc.py <链接> --follow --depth 2         # 跟随抓取的层数(默认1)
  python scripts/fetch_doc.py <链接> --save --images            # 下载文中图片到 output/images/<标题>/ 并在正文引用
  python scripts/fetch_doc.py <链接> --save --html              # 保存为 HTML 格式（默认 md）
  python scripts/fetch_doc.py <链接> --save --format html       # 同上，等价写法
  python scripts/fetch_doc.py <链接> --raw                      # 仅打印纯文本(raw_content,不含表格结构)

保存格式：默认 Markdown(.md)；加 --html 或 --format html 保存为 HTML。未要求保存(不带 --save/--out)时仅返回 Markdown。

引用超链接处理：
  - 文中所有超链接/内联文档引用都会保留为 Markdown 链接，并在文末汇总「引用链接」清单。
  - 加 --follow 时，对其中指向**飞书文档(docx/wiki)**的引用，会自动读取其正文，
    作为「附录」拼接在主文档之后（去重、防循环、按 --depth 控制层数）。
  - 非飞书的外部链接（其他系统、后台地址等）只列出，不抓取。

输出文件名规范：未显式指定文件名时，统一用「飞书文档标题」命名（清洗掉非法字符），落在 output/ 下。

支持的链接形态：
  https://xxx.feishu.cn/docx/<document_id>      新版文档
  https://xxx.feishu.cn/wiki/<node_token>       知识库(自动解析为底层文档)
  直接传 document_id 也可

只读：不修改文档；本期只识别文本与表格，图片不处理。
"""
import os
import re
import sys

# 让 "python scripts/fetch_doc.py" 能 import 同目录模块
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from feishu_client import FeishuClient, FeishuError
from blocks_to_markdown import blocks_to_markdown, collect_mention_user_ids
from feishu_sheet import fetch_spreadsheet_as_markdown, fetch_spreadsheet_as_csv
from markdown_to_html import markdown_to_html

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUTPUT_DIR = os.path.join(PROJECT_ROOT, "output")


def sanitize_filename(name: str, fallback: str = "untitled") -> str:
    """把文档标题清洗成合法文件名：去掉 Windows/类Unix 非法字符与首尾空白点，限长。"""
    name = (name or "").strip()
    # Windows 非法字符: \ / : * ? " < > |  以及控制字符
    name = re.sub(r'[\\/:*?"<>|\r\n\t]', "_", name)
    # 折叠多余空格，去掉首尾的点和空格（Windows 不允许结尾为点/空格）
    name = re.sub(r"\s+", " ", name).strip(" .")
    if not name:
        name = fallback
    # 限制长度，避免超过文件系统上限
    return name[:120]


def parse_link(link: str):
    """从链接里解析出 (类型, token)。
    返回 ('docx', document_id) / ('wiki', node_token) / ('id', 原值)。"""
    link = link.strip()
    m = re.search(r"/docx/([A-Za-z0-9]+)", link)
    if m:
        return "docx", m.group(1)
    m = re.search(r"/wiki/([A-Za-z0-9]+)", link)
    if m:
        return "wiki", m.group(1)
    m = re.search(r"/docs/([A-Za-z0-9]+)", link)
    if m:
        # 老版文档(docs)，docx blocks 接口不一定适用，提示用户
        return "docs", m.group(1)
    # 没有协议/路径，当作直接传入的 document_id
    if re.fullmatch(r"[A-Za-z0-9]+", link):
        return "id", link
    return "unknown", link


def resolve_wiki(client: FeishuClient, link: str):
    """把链接解析为 (obj_type, obj_token)。
    obj_type: 'docx' | 'sheet' | 'bitable' | 'docs' 等。"""
    kind, token = parse_link(link)
    if kind in ("docx", "id"):
        return "docx", token
    if kind == "wiki":
        node = client.resolve_wiki_node(token)
        obj_type = node.get("obj_type", "docx")
        obj_token = node.get("obj_token", token)
        return obj_type, obj_token
    if kind == "docs":
        raise FeishuError(
            "检测到老版文档(/docs/)链接。本工具针对新版文档(/docx/)。"
            "请用新版文档链接，或后续扩展老版接口。"
        )
    raise FeishuError(f"无法识别的链接格式：{link}")


def resolve_document_id(client: FeishuClient, link: str) -> str:
    """把各种链接统一解析为 docx 的 document_id（仅 docx 类型可调用此函数）。"""
    obj_type, obj_token = resolve_wiki(client, link)
    if obj_type == "docx":
        return obj_token
    if obj_type == "sheet":
        raise FeishuError(
            f"该链接底层类型为 sheet（电子表格），请使用 fetch_one 或 main 自动路由。"
        )
    raise FeishuError(
        f"该 wiki 节点底层类型为 {obj_type}，本工具当前只支持 docx 和 sheet。"
    )


def is_feishu_doc_link(url: str) -> bool:
    """判断 URL 是否指向一篇飞书文档/表格（可被本工具读取）。"""
    if not url:
        return False
    kind, _ = parse_link(url)
    return kind in ("docx", "wiki")


def _build_image_saver(client: FeishuClient, image_dir: str, rel_prefix: str):
    """构造图片下载器：给 token 就下载并保存到 image_dir，返回相对 rel_prefix 的路径。
    下载失败返回 None（调用方退化为注释标注）。带 token 去重缓存。"""
    import hashlib
    cache = {}

    def saver(token: str):
        if not token:
            return None
        if token in cache:
            return cache[token]
        content = client.download_media(token)  # 失败会抛异常，由上层 try 兜底
        # 用文件头猜扩展名
        ext = _guess_image_ext(content)
        fname = f"{hashlib.md5(token.encode()).hexdigest()[:12]}{ext}"
        os.makedirs(image_dir, exist_ok=True)
        fpath = os.path.join(image_dir, fname)
        with open(fpath, "wb") as f:
            f.write(content)
        rel = f"{rel_prefix}/{fname}" if rel_prefix else fname
        cache[token] = rel
        return rel

    return saver


def _guess_image_ext(content: bytes) -> str:
    """根据文件头判断图片扩展名。"""
    if content[:8].startswith(b"\x89PNG"):
        return ".png"
    if content[:3] == b"\xff\xd8\xff":
        return ".jpg"
    if content[:6] in (b"GIF87a", b"GIF89a"):
        return ".gif"
    if content[:4] == b"RIFF" and content[8:12] == b"WEBP":
        return ".webp"
    if content[:2] == b"BM":
        return ".bmp"
    return ".png"  # 兜底


def fetch_one(client: FeishuClient, link: str, collect_links: list = None, output_dir: str = None,
              save_images: bool = False, image_title: str = None):
    """读取单篇文档或电子表格。
    - docx: 返回 ('docx', title, markdown正文)
    - sheet: 如果 output_dir 提供则存 CSV 返回 ('sheet', title, [csv_paths])，
             否则返回 ('sheet', title, markdown摘要)
    collect_links 若提供，会被填入该文档内引用的 (标题, URL)（仅 docx 类型会收集）。
    save_images=True 且 output_dir 提供时，下载文中图片到 output_dir/images/<标题>/ 并在正文引用。"""
    obj_type, obj_token = resolve_wiki(client, link)

    if obj_type == "sheet":
        if output_dir:
            title, xlsx_path = fetch_spreadsheet_as_csv(client, obj_token, output_dir)
            return "sheet", title, xlsx_path
        else:
            title, md = fetch_spreadsheet_as_markdown(client, obj_token)
            return "sheet", title, md

    if obj_type == "docx":
        title = ""
        try:
            basic = client.get_document_basic(obj_token)
            title = basic.get("title", "")
        except FeishuError:
            pass
        blocks = client.list_all_blocks(obj_token)
        # 收集 @用户 的 open_id，批量换姓名（无权限/查不到则自动兜底为占位）
        user_names = {}
        try:
            open_ids = collect_mention_user_ids(blocks)
            if open_ids:
                user_names = client.get_user_names(open_ids)
        except Exception:
            user_names = {}
        # 图片下载器：仅当要求存图且有输出目录时启用
        image_saver = None
        if save_images and output_dir:
            safe_title = sanitize_filename(image_title or title or obj_token)
            image_dir = os.path.join(output_dir, "images", safe_title)
            image_saver = _build_image_saver(client, image_dir, rel_prefix=f"images/{safe_title}")
        md = blocks_to_markdown(blocks, doc_title=title, links_out=collect_links,
                                user_names=user_names, image_saver=image_saver)
        return "docx", title, md

    raise FeishuError(f"不支持的文档类型：{obj_type}（当前只支持 docx 和 sheet）")


def main():
    args = sys.argv[1:]
    if not args:
        print(__doc__)
        sys.exit(1)

    link = args[0]
    use_raw = "--raw" in args
    do_save = "--save" in args
    do_follow = "--follow" in args
    save_images = "--images" in args
    # 保存格式：默认 md；--html 或 --format html 切换为 html
    fmt = "md"
    if "--html" in args:
        fmt = "html"
    if "--format" in args:
        i = args.index("--format")
        if len(args) > i + 1 and args[i + 1].lower() in ("md", "html"):
            fmt = args[i + 1].lower()
    out_arg = None
    if "--out" in args:
        i = args.index("--out")
        if len(args) > i + 1:
            out_arg = args[i + 1]
    depth = 1
    if "--depth" in args:
        i = args.index("--depth")
        if len(args) > i + 1:
            try:
                depth = max(1, int(args[i + 1]))
            except ValueError:
                pass

    client = FeishuClient()

    try:
        obj_type, obj_token = resolve_wiki(client, link)

        if use_raw:
            # 纯文本兜底模式：仅适用于 docx
            if obj_type != "docx":
                raise FeishuError(f"--raw 模式仅支持 docx 文档，当前类型为 {obj_type}")
            import requests
            title = ""
            try:
                title = client.get_document_basic(obj_token).get("title", "")
            except FeishuError:
                pass
            resp = requests.get(
                f"{client.host}/open-apis/docx/v1/documents/{obj_token}/raw_content",
                headers=client.headers,
                timeout=60,
            )
            resp.raise_for_status()
            data = resp.json()
            if data.get("code") != 0:
                raise FeishuError(f"raw_content 失败: code={data.get('code')} msg={data.get('msg')}")
            md = data.get("data", {}).get("content", "")
        else:
            # 主文档（收集其引用链接）
            root_links = []
            result = fetch_one(
                client, link, collect_links=root_links,
                output_dir=DEFAULT_OUTPUT_DIR if (do_save or out_arg) else None,
                save_images=save_images,
            )
            doc_type = result[0]
            title = result[1]

            if doc_type == "sheet":
                # 电子表格：已输出为 xlsx
                if do_save or out_arg:
                    xlsx_path = result[2]
                    print(f"[OK] 电子表格「{title}」已导出到 {xlsx_path}")
                    return
                else:
                    # 无 --save/--out 时打印 markdown 摘要到终端
                    md = result[2]
            else:
                # docx
                md = result[2]
                # --follow：对引用到的「飞书文档」逐层读取，作为附录拼接
                if do_follow:
                    visited = {obj_token}
                    appendix, followed = _follow_referenced_docs(
                        client, root_links, depth, visited=visited
                    )
                    if appendix:
                        md = md.rstrip() + "\n\n" + appendix
                        print(f"[follow] 成功附加 {followed} 篇被引用飞书文档（其余未读取的已在附录中标注原因）")
                    else:
                        print("[follow] 未发现可读取的飞书文档引用")

    except FeishuError as e:
        print(f"[飞书错误] {e}")
        sys.exit(1)
    except Exception as e:
        print(f"[执行失败] {type(e).__name__}: {e}")
        sys.exit(1)

    # 计算输出路径：
    # - 未给 --out 且未给 --save：仅打印 Markdown（不保存）
    # - --save 或 --out 指向目录：output/<文档标题>.<fmt>
    # - --out 指向具体文件：用该路径（后缀 .html 自动视为 html 格式）
    ext = ".html" if fmt == "html" else ".md"
    out_path = None
    title_name = sanitize_filename(title, fallback=obj_token)
    if out_arg:
        low = out_arg.lower()
        if out_arg.endswith((os.sep, "/")) or os.path.isdir(out_arg) or not low.endswith((".md", ".html", ".htm")):
            out_path = os.path.join(out_arg, f"{title_name}{ext}")
        else:
            out_path = out_arg
            # 若显式给了 .html 文件名，则以文件后缀为准
            if low.endswith((".html", ".htm")):
                fmt = "html"
    elif do_save:
        out_path = os.path.join(DEFAULT_OUTPUT_DIR, f"{title_name}{ext}")

    if out_path:
        # 按格式生成内容
        if fmt == "html":
            content = markdown_to_html(md, title=title)
        else:
            content = md
        os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(content)
        print(f"[OK] 已保存到 {out_path}（标题：{title or '(未取到)'}，格式：{fmt}，{len(content)} 字）")
    else:
        # 未要求保存：默认返回/打印 Markdown
        print(md)


def _follow_referenced_docs(client: FeishuClient, root_links: list, depth: int, visited: set):
    """对引用到的飞书文档逐层(BFS)读取，返回 (附录Markdown, 已抓取篇数)。
    - 只跟随飞书 docx/wiki 链接；外部链接忽略。
    - visited 用 document_id 去重，防止循环引用。
    - 每层新发现的引用进入下一层，直到 depth 用尽。"""
    appendix_parts = []
    followed = 0
    # 当前层待处理的 (标题, URL)
    frontier = [(t, u) for (t, u) in root_links if is_feishu_doc_link(u)]

    for level in range(depth):
        next_frontier = []
        for ref_title, url in frontier:
            try:
                doc_id = resolve_document_id(client, url)
            except FeishuError as e:
                # wiki 底层是 sheet/bitable 等非 docx，或老版文档：附录里留说明，不静默丢弃
                appendix_parts.append(
                    f"\n---\n\n## [被引用文档未读取] {ref_title}\n\n"
                    f"> 来源：{url}\n>\n> 原因：{e}\n"
                )
                continue
            if doc_id in visited:
                continue
            visited.add(doc_id)
            try:
                child_links = []
                title, md = fetch_one(client, url, collect_links=child_links)
            except Exception as e:
                # 包括无权限(403)、网络错误等：记入附录，不影响主文档与其它引用
                reason = str(e)
                if "403" in reason or "Forbidden" in reason:
                    reason = "无权限（该文档未把本应用加为协作者，或权限未发布生效）"
                appendix_parts.append(
                    f"\n---\n\n## [被引用文档读取失败] {ref_title}\n\n"
                    f"> 来源：{url}\n>\n> 原因：{reason}\n"
                )
                continue
            followed += 1
            appendix_parts.append(
                f"\n---\n\n# 📎 附录：被引用文档「{title or ref_title}」\n\n"
                f"> 来源：{url}\n\n{md.strip()}\n"
            )
            # 收集下一层引用
            next_frontier.extend([(t, u) for (t, u) in child_links if is_feishu_doc_link(u)])
        frontier = next_frontier
        if not frontier:
            break

    return "".join(appendix_parts), followed


if __name__ == "__main__":
    main()
