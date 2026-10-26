# -*- coding: utf-8 -*-
"""
Markdown -> HTML
================
把读取到的 Markdown 正文渲染成一份自包含的 HTML 文档（含基础样式）。
用于用户选择以 html 格式保存本地时。

依赖 markdown 库（pip install markdown），启用 tables/fenced_code 等扩展。
图片引用 ![](images/...) 保持相对路径，与 .md 同目录时可正常显示。
"""
import html as _html


_HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title}</title>
<style>
  body {{ font-family: -apple-system, "Segoe UI", "Microsoft YaHei", sans-serif;
         line-height: 1.7; max-width: 900px; margin: 0 auto; padding: 24px; color: #1f2329; }}
  h1, h2, h3, h4 {{ font-weight: 600; line-height: 1.3; margin: 1.2em 0 0.6em; }}
  h1 {{ font-size: 1.8em; border-bottom: 2px solid #e5e6eb; padding-bottom: 0.3em; }}
  h2 {{ font-size: 1.4em; border-bottom: 1px solid #e5e6eb; padding-bottom: 0.2em; }}
  code {{ background: #f2f3f5; padding: 2px 5px; border-radius: 4px;
          font-family: Consolas, Monaco, monospace; font-size: 0.9em; }}
  pre {{ background: #f7f8fa; padding: 14px; border-radius: 6px; overflow-x: auto; }}
  pre code {{ background: none; padding: 0; }}
  table {{ border-collapse: collapse; width: 100%; margin: 1em 0; }}
  th, td {{ border: 1px solid #dcdfe6; padding: 8px 12px; text-align: left; vertical-align: top; }}
  th {{ background: #f2f3f5; font-weight: 600; }}
  blockquote {{ border-left: 4px solid #c9cdd4; margin: 1em 0; padding: 0.3em 1em; color: #4e5969; background: #fafafa; }}
  img {{ max-width: 100%; height: auto; border-radius: 4px; }}
  a {{ color: #165dff; text-decoration: none; }}
  a:hover {{ text-decoration: underline; }}
  hr {{ border: none; border-top: 1px solid #e5e6eb; margin: 1.5em 0; }}
</style>
</head>
<body>
{body}
</body>
</html>
"""


def markdown_to_html(md_text: str, title: str = "") -> str:
    """把 Markdown 文本转成完整 HTML 文档字符串。"""
    try:
        import markdown as _md
        body = _md.markdown(
            md_text,
            extensions=["tables", "fenced_code", "nl2br", "sane_lists", "toc"],
            output_format="html5",
        )
    except ImportError:
        # 兜底：markdown 库不可用时，做最小转义包裹，至少不丢内容
        body = "<pre>" + _html.escape(md_text) + "</pre>"

    safe_title = _html.escape(title or "飞书文档")
    return _HTML_TEMPLATE.format(title=safe_title, body=body)
