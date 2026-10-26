# feishu-mcp

飞书（Feishu/Lark）云文档工具（MCP Server + CLI）。以**读**为主（链接 → Markdown，供 AI 工具总结/分析），同时支持把 Markdown **写**回飞书 wiki 文档。

> 读取数据流是单向的：AI 工具主动「拉」文档内容，**不需要 webhook、不需要事件订阅、不需要在飞书里和机器人对话**。
> 这里的「飞书机器人」本质上是一个飞书自建应用 + 本地脚本。
> 读取只识别**文本与表格**，图片不处理；写入支持文本/表格等常见块，图片同样不支持（无 Drive 权限）。

## 目录结构

```
feishu-mcp/
├── server.py                   # MCP Server 入口：暴露 5 个工具给 MCP 客户端原生调用
├── scripts/
│   ├── feishu_client.py        # 底层客户端：换 token + 读写 docx 块 + wiki 节点操作
│   ├── blocks_to_markdown.py   # 【读】docx 块 → Markdown（文本/标题/列表/代码/引用/表格）
│   ├── feishu_sheet.py         # 【读】电子表格(Sheets) → Markdown 表格
│   ├── markdown_to_blocks.py   # 【写】Markdown → docx 块结构（标题/正文/列表/代码/引用/分割线/表格）
│   ├── markdown_to_html.py     # 【读】Markdown → 自包含 HTML（保存 html 格式时用）
│   └── fetch_doc.py            # CLI 入口：链接 → Markdown（含跟随引用、存盘、格式选择等）
├── output/                     # 存盘的 .md / 缓存文件，文件名=飞书文档标题（已 gitignore，首次运行自动创建）
├── .env.example                # 凭证 + 写入目标模板（复制为 .env 填写）
├── .env                        # 真实凭证（已 gitignore，需自行创建）
├── .feishu_token               # token 缓存（自动生成，已 gitignore）
├── requirements.txt            # 依赖：requests / openpyxl / mcp / markdown
├── .gitignore
└── README.md
```

### 模块职责一览

| 文件 | 方向 | 职责 |
|------|------|------|
| `server.py` | — | MCP Server，注册 `fetch_document` / `fetch_sheet` / `create_document` / `list_saved_docs` / `read_saved_doc` |
| `feishu_client.py` | 读+写 | 封装 `tenant_access_token` 缓存、docx 块读写、wiki 节点创建、表格列宽 PATCH |
| `blocks_to_markdown.py` | 读 | docx 块结构 → Markdown |
| `feishu_sheet.py` | 读 | 电子表格元信息 + 单元格数据 → Markdown 表格 |
| `markdown_to_blocks.py` | 写 | Markdown → 飞书 docx 块（descendant 接口所需的嵌套结构） |
| `fetch_doc.py` | 读 | CLI 入口，封装链接解析、跟随引用、存盘 |

## 一次性配置（飞书后台）

1. **创建自建应用**：[open.feishu.cn](https://open.feishu.cn) → 开发者后台 → 创建企业自建应用 → 拿到 `App ID` / `App Secret`。
2. **开通权限(scope)**：
   - `docx:document:readonly`（读新版文档）— 仅读可用 readonly，写需 `docx:document`
   - `wiki:wiki:readonly`（走 wiki 链接需要）— 写文档需 `wiki:wiki`
   - `sheets:spreadsheet:readonly`（读电子表格 `fetch_sheet` 需要）
3. **发布应用**：版本管理与发布 → 创建版本 → 提交，管理员审核通过后权限才生效。
4. **把目标文档授权给应用**（= 指定权限）：打开文档 → 右上「···/分享」→ 添加协作者 → 搜应用名 → 给「可阅读」。
   - 没加进来的文档读不到，天然权限隔离。

## 本地配置

把 `.env.example` 复制为 `.env`，填入凭证（当前已建好 `.env`，App ID/Secret 留空，待你填）：

```
FEISHU_APP_ID=cli_xxxxxxxx
FEISHU_APP_SECRET=xxxxxxxx
```

安装依赖（只需 requests）：

```
pip install -r requirements.txt
```

## 用法（cwd = feishu-mcp/）

```bash
# 输入链接，打印 Markdown 正文（不存盘）
python scripts/fetch_doc.py https://xxx.feishu.cn/docx/abcd1234

# 存盘：文件名自动用「飞书文档标题」，落在 output/<标题>.md
python scripts/fetch_doc.py <链接> --save

# 也支持 wiki 链接 / 直接传 document_id
python scripts/fetch_doc.py https://xxx.feishu.cn/wiki/xxxxxx --save
python scripts/fetch_doc.py abcd1234 --save

# 跟随抓取：把文中引用的「飞书文档」内容也一并读取，作为附录拼在主文档后
python scripts/fetch_doc.py <链接> --follow --save
python scripts/fetch_doc.py <链接> --follow --depth 2     # 跟随层数(默认1)

# 指定具体文件名 / 指定目录（目录则仍用文档标题命名）
python scripts/fetch_doc.py <链接> --out output/my.md
python scripts/fetch_doc.py <链接> --out output/

# 下载文中图片（存到 output/images/<标题>/，正文用 ![](相对路径) 引用）
python scripts/fetch_doc.py <链接> --save --images

# 保存为 HTML 格式（默认 md）
python scripts/fetch_doc.py <链接> --save --html
python scripts/fetch_doc.py <链接> --save --format html   # 等价写法
python scripts/fetch_doc.py <链接> --out output/x.html    # 指定 .html 文件名也会按 html 存

# 纯文本兜底（raw_content，丢表格结构，排查用）
python scripts/fetch_doc.py <链接> --raw
```

### 保存格式（md / html）

- **默认 Markdown**：不带 `--html`/`--format html` 时，存盘为 `.md`。
- **HTML**：加 `--html` 或 `--format html`，存盘为自包含 `.html`（带基础样式，表格/图片/链接均正常渲染）。`--out` 指定 `.html` 文件名时也自动按 html 存。
- **不保存时**（不带 `--save`/`--out`）：始终返回 Markdown，不生成文件。
- HTML 里图片引用沿用相对路径 `images/<标题>/xxx`，与 `.md` 同目录时可直接显示；建议存 html 时配合 `--images` 一起用。
- MCP 工具 `fetch_document` 对应参数为 `save_format="html"`（返回给调用方的 `content` 始终是 Markdown，仅存盘格式受影响）。

> 输出文件名规范：未显式给 `.md` 文件名时，统一用**飞书文档标题**命名（自动清洗 `\ / : * ? " < > |` 等非法字符，限长120），取不到标题时回退用 document_id。

### 图片读取（`--images`）

加 `--images`（需先开通 `drive:file:download` 权限）会把文中图片下载到本地并在正文引用：

- 图片存到 `output/images/<文档标题>/<hash>.<ext>`，正文中渲染为 `![图片](images/<标题>/xxx.png)`（相对 `.md` 的路径）。
- 扩展名按文件头自动识别（png/jpg/gif/webp/bmp）；同一图片 token 去重，只下载一次。
- 未加 `--images` 或下载失败时，图片位置退化为 `<!-- [图片未下载, token=xxx] -->` 注释，不影响正文。
- MCP 工具 `fetch_document` 对应参数为 `images=true`（开启时自动存盘）。

### 引用超链接处理

文中所有超链接都会被识别，并分三档处理：

1. **原位保留**为 Markdown 链接 `[文字](URL)`，不破坏阅读顺序。
2. **文末汇总**「引用链接」清单（去重、按出现顺序）。
3. **加 `--follow` 时跟随读取**：对其中指向**飞书 docx/wiki 文档**的引用，自动读取正文，作为「📎 附录」拼在主文档之后。
   - 去重 + 防循环引用，按 `--depth` 控制层数（默认1层）。
   - 非飞书外部链接（其他系统、后台地址等）只列出、不抓取。
   - 被引用文档若无权限、或底层是表格/多维表(sheet/bitable)等非 docx 类型，会在附录中**标注原因**而非静默跳过，且不影响主文档输出。

## 与 AI 工具配合的工作流

```
你：(贴链接) https://xxx.feishu.cn/docx/abcd1234  帮我总结这篇文档
AI 工具：调用 MCP 工具 fetch_document → 拿到 Markdown 正文 → 按你的要求总结/分析
```

## 设计要点

- **token 缓存**：`tenant_access_token` 约 2 小时有效，缓存在 `.feishu_token`，到期前自动复用，过期自动重取。
- **表格识别**：docx 表格是「table 块 + 若干 table_cell 子块」结构。读取时按行列还原成 Markdown 表格，单元格内多行序列化为 `<br>`；写入时反向把 `<br>` 还原为单元格内多行（见上文「表格单元格内换行」）。
- **链接识别**：两类链接都会被保留为 Markdown 链接 `[文字](URL)`，并在文末汇总成「引用链接」清单（去重、按出现顺序）：
  - **内联文档引用**（mention_doc，飞书里 @ 出来的子文档/云文档）
  - **选中文字超链接**（text_run 上的 link，如指向其他系统、后台地址等）
  - URL 会做反转义（飞书把 `&`、`/` 等编码过），还原成可点击地址。
- **错误提示**：取块失败时会提示常见原因（文档未授权给应用 / 权限未发布生效）。
- **安全**：`App Secret` 只放 `.env`，不写进代码、不入库。

## 待扩展（本期未做）

- 图片写入（需 Drive 上传权限 `drive:drive`，当前环境无法开通）。
- 图片读取 → 文字（OCR / 视觉大模型），按原位置插回。
- 老版文档(`/docs/`)接口支持。
- 嵌套列表层级、todo、callout 等块类型的写入。

## MCP Server 模式（MCP 客户端原生工具调用）

除了 CLI 模式外，本项目还提供 MCP Server，让支持 MCP 的 AI 工具（Kiro、Claude Code、Cursor 等）直接以原生工具形式调用飞书读取能力。

### 暴露的工具

| 工具名 | 功能 | 核心参数 |
|--------|------|----------|
| `fetch_document` | 读取飞书文档(docx/wiki)，返回 Markdown | `url`, `follow`, `depth`, `save` |
| `fetch_sheet` | 读取飞书电子表格，返回 Markdown 表格 | `url`, `max_rows`, `max_display_rows` |
| `create_document` | **【输出到飞书文档】** 在 wiki 目录下创建新文档并写入 Markdown | `title`, `markdown`, `parent_node_token` |
| `list_saved_docs` | 列出 output/ 中已缓存的文件 | 无 |
| `read_saved_doc` | 读取已缓存的本地文档内容 | `filename` |

> **`create_document` 的触发约束**：仅当用户给出明确指令「输出到飞书文档」（或等价表述）时才会被调用。它**不在 autoApprove 列表**中，每次写入都需人工确认，作为写操作的额外保险。

### 写入能力说明（绕开 drive 权限）

本环境无法开通 `drive:drive`，因此**不走** import_task 导入路径，改用纯 wiki + docx 接口：

1. `POST /wiki/v2/spaces/{space_id}/nodes` —— 在 wiki 节点下建一篇空 docx（需 `wiki:wiki`）
2. `POST /docx/v1/documents/{doc}/blocks/{doc}/descendant` —— 批量写入内容块（需 `docx:document`）

Markdown → 飞书原生块由 `markdown_to_blocks.py` 转换，支持：标题(1-9级)、正文（粗体/斜体/删除线/行内代码/链接）、有序/无序列表、代码块、引用、分割线、表格。块按 50 个/批自动分批提交。**图片暂不支持**（需 Drive 上传权限）。

#### 表格单元格内换行（`<br>`）

表格单元格内用 `<br>`（或 `<br/>` / `<br />`，大小写均可）实现换行：转换时按 `<br>` 切分，每段生成一个独立 text 子块挂到同一 `table_cell` 下，飞书原生渲染为单元格内多行，**无残留字符、行内样式（粗体等）保留**。

- 实现：`markdown_to_blocks.py` 的 `_split_cell_lines()` + `_table_unit()`（一个 cell 挂多个 text 子块）。
- 注意：Markdown 表格的一行必须在同一物理行，**不要在表格行内敲真实回车**，否则会冲断表格结构。
- 列宽：飞书创建表格时忽略 `column_width`，须写入后用 PATCH 接口逐列设置；`_cell_display_width()` 按内容（含 `<br>` 分行）估算列宽。

> ⚠️ **改完 `scripts/*.py` 必须重启 MCP Server 才生效**：MCP Server 是常驻进程，不会热加载磁盘改动。验证代码改动时若不便重启，可写独立脚本直接 import 磁盘模块 + 调 `FeishuClient` 直连验证，避免误判改动未生效。

#### 写入前置条件

- 目标 wiki 节点须分享为「**组织内可编辑**」，应用身份才有写权限（企业知识库无法把应用加为成员时的可行解）。
- 在 `.env` 配置写入目标：

```
FEISHU_WIKI_SPACE_ID=<知识库 space_id>
FEISHU_WIKI_PARENT_NODE=<wiki 链接 /wiki/ 后的 token>
FEISHU_WIKI_DOMAIN=your-tenant.feishu.cn
```

新文档会作为该父节点的子节点创建，归属随节点所在目录。

### 配置方式

在你所用 MCP 客户端的配置文件里添加（以 Kiro 的 `.kiro/settings/mcp.json` 为例）：

```json
{
  "mcpServers": {
    "feishu-mcp": {
      "command": "python",
      "args": ["/absolute/path/to/feishu-mcp/server.py"],
      "disabled": false,
      "autoApprove": ["fetch_document", "fetch_sheet", "list_saved_docs", "read_saved_doc"]
    }
  }
}
```

### 依赖安装

```bash
pip install -r requirements.txt   # 包含 mcp>=1.0.0
```

### 手动测试

```bash
# 以 stdio 模式启动（正常由 MCP 客户端自动拉起，无需手动）
python server.py
```
