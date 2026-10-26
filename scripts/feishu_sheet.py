"""
飞书电子表格（Sheets）只读模块
==============================
职责：给定 spreadsheet_token，读取所有工作表(sheet)的单元格数据，
      格式化成 Markdown 表格返回。

依赖 feishu_client.py 的 FeishuClient（复用 token 管理）。
"""
import os
import sys
import requests
from typing import Optional

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from feishu_client import FeishuClient, FeishuError


def list_sheets(client: FeishuClient, spreadsheet_token: str) -> tuple:
    """通过 metainfo 接口获取电子表格标题及所有工作表信息。
    返回 (spreadsheet_title, [sheet_info...])。
    sheet_info: {sheet_id, title, row_count, column_count}"""
    resp = requests.get(
        f"{client.host}/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/metainfo",
        headers=client.headers,
        timeout=30,
    )
    resp.raise_for_status()
    data = resp.json()
    if data.get("code") != 0:
        raise FeishuError(
            f"获取电子表格元信息失败: code={data.get('code')} msg={data.get('msg')}。"
            "常见原因：未开通 sheets:spreadsheet:readonly 权限，或电子表格未把应用加为协作者。"
        )
    payload = data.get("data", {})
    sp_title = payload.get("properties", {}).get("title", "")
    sheets = payload.get("sheets", [])
    return sp_title, [
        {
            "sheet_id": s.get("sheetId"),
            "title": s.get("title", ""),
            "row_count": s.get("rowCount", 0),
            "column_count": s.get("columnCount", 0),
        }
        for s in sheets
    ]


def read_sheet_values(
    client: FeishuClient,
    spreadsheet_token: str,
    sheet_id: str,
    row_count: int = 0,
    column_count: int = 0,
    max_rows: int = 5000,
) -> list:
    """读取某个工作表的全部有效数据（二维数组）。
    返回 [[cell, cell, ...], ...]，首行通常是表头。"""
    # 构造范围：如果知道行列数就精确取，否则用大范围兜底
    if row_count and column_count:
        # 限制最大行数防止超大表卡住
        actual_rows = min(row_count, max_rows)
        end_col = _col_letter(column_count)
        range_str = f"{sheet_id}!A1:{end_col}{actual_rows}"
    else:
        range_str = f"{sheet_id}!A1:ZZ{max_rows}"

    resp = requests.get(
        f"{client.host}/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/values/{range_str}",
        headers=client.headers,
        params={"valueRenderOption": "ToString"},
        timeout=120,
    )
    resp.raise_for_status()
    data = resp.json()
    if data.get("code") != 0:
        raise FeishuError(
            f"读取工作表数据失败: code={data.get('code')} msg={data.get('msg')}"
        )
    values = data.get("data", {}).get("valueRange", {}).get("values", [])
    return values


def sheet_to_markdown(
    title: str,
    values: list,
    max_display_rows: int = 200,
) -> str:
    """把二维数组转成 Markdown 表格。首行当表头。"""
    if not values:
        return f"## {title}\n\n（空表）\n"

    lines = [f"## {title}", ""]

    # 首行作为表头
    header = values[0]
    col_count = len(header)
    header_cells = [_cell_str(c) for c in header]
    lines.append("| " + " | ".join(header_cells) + " |")
    lines.append("| " + " | ".join("---" for _ in header_cells) + " |")

    # 数据行
    display = values[1 : max_display_rows + 1]
    for row in display:
        # 补齐/截断到表头列数
        cells = [_cell_str(row[i]) if i < len(row) else "" for i in range(col_count)]
        lines.append("| " + " | ".join(cells) + " |")

    if len(values) - 1 > max_display_rows:
        lines.append(f"\n> ... 共 {len(values) - 1} 行数据，仅显示前 {max_display_rows} 行")

    lines.append("")
    return "\n".join(lines)


def fetch_spreadsheet_as_markdown(
    client: FeishuClient,
    spreadsheet_token: str,
    max_rows_per_sheet: int = 5000,
    max_display_rows: int = 200,
) -> tuple:
    """读取整个电子表格，返回 (spreadsheet_title, summary_markdown)。
    实际数据以 CSV 形式存盘（调用方需传 output_dir），这里返回摘要供终端打印。"""
    sp_title, sheets = list_sheets(client, spreadsheet_token)
    if not sheets:
        return sp_title or "", "（电子表格无工作表）\n"

    parts = [f"# {sp_title}\n"] if sp_title else []
    for s in sheets:
        values = read_sheet_values(
            client,
            spreadsheet_token,
            s["sheet_id"],
            row_count=s.get("row_count", 0),
            column_count=s.get("column_count", 0),
            max_rows=max_rows_per_sheet,
        )
        md = sheet_to_markdown(s["title"], values, max_display_rows=max_display_rows)
        parts.append(md)

    return sp_title, "\n".join(parts)


def fetch_spreadsheet_as_csv(
    client: FeishuClient,
    spreadsheet_token: str,
    output_dir: str,
    max_rows_per_sheet: int = 5000,
) -> tuple:
    """读取整个电子表格，输出为单个 .xlsx 文件，多 sheet 页保持原结构。
    返回 (spreadsheet_title, xlsx_filepath)。"""
    from openpyxl import Workbook

    sp_title, sheets = list_sheets(client, spreadsheet_token)
    if not sheets:
        return sp_title or "", ""

    os.makedirs(output_dir, exist_ok=True)
    wb = Workbook()
    # 删除默认创建的空 sheet
    wb.remove(wb.active)

    for s in sheets:
        values = read_sheet_values(
            client,
            spreadsheet_token,
            s["sheet_id"],
            row_count=s.get("row_count", 0),
            column_count=s.get("column_count", 0),
            max_rows=max_rows_per_sheet,
        )
        # sheet 名最长 31 字符（Excel 限制）
        sheet_title = (s["title"] or s["sheet_id"])[:31]
        ws = wb.create_sheet(title=sheet_title)
        for row in (values or []):
            ws.append([c if c is not None else "" for c in row])

    filename = f"{_sanitize(sp_title) or spreadsheet_token}.xlsx"
    filepath = os.path.join(output_dir, filename)
    wb.save(filepath)

    return sp_title, filepath


def _sanitize(name: str) -> str:
    """简易文件名清洗。"""
    import re as re_mod
    name = (name or "").strip()
    name = re_mod.sub(r'[\\/:*?"<>|\r\n\t]', "_", name)
    name = re_mod.sub(r"\s+", " ", name).strip(" .")
    return name[:120]


def _col_letter(n: int) -> str:
    """列号(1-based) -> Excel 列字母（如 1->A, 27->AA）。"""
    result = ""
    while n > 0:
        n -= 1
        result = chr(n % 26 + ord("A")) + result
        n //= 26
    return result or "A"


def _cell_str(val) -> str:
    """把单元格值转成字符串，处理 None、换行、管道符。"""
    if val is None:
        return ""
    s = str(val).replace("|", "\\|")
    # 换行在 Markdown 表格里用 <br> 保持单行
    s = s.replace("\r\n", "<br>").replace("\n", "<br>").replace("\r", "<br>")
    return s
