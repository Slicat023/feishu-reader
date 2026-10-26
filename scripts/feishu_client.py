"""
飞书（Feishu/Lark）云文档只读客户端
==================================
职责：用自建应用身份（App ID / App Secret）拿 tenant_access_token，
      调云文档 OpenAPI 读取「新版文档(docx)」的全部块(blocks)。

只做「读」，不写、不收消息、不需要 webhook/事件订阅。
图片暂不处理，只识别文本与表格。

token 会缓存到项目根的 .feishu_token（含过期时间），到期前自动复用。
"""
import os
import json
import time
import requests
from typing import Optional

# 飞书开放平台域名（国内版）；如用海外 Lark 改为 https://open.larksuite.com
FEISHU_HOST = "https://open.feishu.cn"

# 项目根 = scripts/ 的上一级，token 缓存与 .env 都落在这里，与脚本位置解耦
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TOKEN_CACHE_FILE = os.path.join(PROJECT_ROOT, ".feishu_token")
ENV_FILE = os.path.join(PROJECT_ROOT, ".env")


def load_env():
    """极简 .env 加载器：把 KEY=VALUE 读进 os.environ（不覆盖已存在的环境变量）。
    避免引入 python-dotenv 依赖。"""
    if not os.path.exists(ENV_FILE):
        return
    with open(ENV_FILE, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            if key and key not in os.environ:
                os.environ[key] = value


class FeishuError(Exception):
    """飞书 API 返回的业务错误（code != 0）"""
    pass


class FeishuClient:
    def __init__(self):
        load_env()
        self.app_id = os.environ.get("FEISHU_APP_ID", "").strip()
        self.app_secret = os.environ.get("FEISHU_APP_SECRET", "").strip()
        self.host = os.environ.get("FEISHU_HOST", FEISHU_HOST).rstrip("/")
        self._token = None
        self._token_expire_at = 0
        self._user_name_cache = {}  # open_id -> name，进程内缓存，避免重复查

    # ---------- token 管理 ----------
    def _load_cached_token(self):
        if os.path.exists(TOKEN_CACHE_FILE):
            try:
                with open(TOKEN_CACHE_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                # 留 60 秒缓冲，避免临界过期
                if data.get("expire_at", 0) - 60 > time.time():
                    return data.get("token")
            except (json.JSONDecodeError, OSError):
                pass
        return None

    def _save_cached_token(self, token: str, expire_in: int):
        data = {"token": token, "expire_at": time.time() + expire_in}
        with open(TOKEN_CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f)

    def get_token(self) -> str:
        """获取 tenant_access_token（应用身份），优先用本地缓存。"""
        if self._token and self._token_expire_at - 60 > time.time():
            return self._token

        cached = self._load_cached_token()
        if cached:
            self._token = cached
            return cached

        if not self.app_id or not self.app_secret:
            raise FeishuError(
                "缺少 FEISHU_APP_ID / FEISHU_APP_SECRET。"
                "请在 feishu-doc/.env 中填写（参考 .env.example）。"
            )

        resp = requests.post(
            f"{self.host}/open-apis/auth/v3/tenant_access_token/internal",
            json={"app_id": self.app_id, "app_secret": self.app_secret},
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        if data.get("code") != 0:
            raise FeishuError(f"获取 token 失败: code={data.get('code')} msg={data.get('msg')}")
        self._token = data["tenant_access_token"]
        expire_in = data.get("expire", 7200)
        self._token_expire_at = time.time() + expire_in
        self._save_cached_token(self._token, expire_in)
        return self._token

    @property
    def headers(self):
        return {"Authorization": f"Bearer {self.get_token()}"}

    # ---------- 文档读取 ----------
    def get_document_basic(self, document_id: str) -> dict:
        """取文档基础信息（主要为了拿标题 title）。"""
        resp = requests.get(
            f"{self.host}/open-apis/docx/v1/documents/{document_id}",
            headers=self.headers,
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        if data.get("code") != 0:
            raise FeishuError(f"获取文档信息失败: code={data.get('code')} msg={data.get('msg')}")
        return data.get("data", {}).get("document", {})

    def list_all_blocks(self, document_id: str) -> list:
        """拉取文档全部块（自动翻页）。返回扁平的 block 列表。"""
        blocks = []
        page_token = None
        while True:
            params = {"page_size": 500}
            if page_token:
                params["page_token"] = page_token
            resp = requests.get(
                f"{self.host}/open-apis/docx/v1/documents/{document_id}/blocks",
                headers=self.headers,
                params=params,
                timeout=60,
            )
            resp.raise_for_status()
            data = resp.json()
            if data.get("code") != 0:
                raise FeishuError(
                    f"获取文档块失败: code={data.get('code')} msg={data.get('msg')}。"
                    "常见原因：文档未把本应用加为协作者(可阅读)，或权限未发布生效。"
                )
            payload = data.get("data", {})
            blocks.extend(payload.get("items", []))
            if payload.get("has_more"):
                page_token = payload.get("page_token")
            else:
                break
        return blocks

    def resolve_wiki_node(self, node_token: str) -> dict:
        """知识库(wiki)链接解析：node_token -> 真实 obj_token / obj_type。
        wiki 链接形如 .../wiki/<node_token>，其 node_token 不是 document_id。"""
        resp = requests.get(
            f"{self.host}/open-apis/wiki/v2/spaces/get_node",
            headers=self.headers,
            params={"token": node_token},
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        if data.get("code") != 0:
            raise FeishuError(f"解析 wiki 节点失败: code={data.get('code')} msg={data.get('msg')}")
        return data.get("data", {}).get("node", {})

    # ---------- 文档写入（wiki 建节点 + docx 写块） ----------
    def create_wiki_node(self, space_id: str, parent_node_token: str, title: str,
                         obj_type: str = "docx") -> dict:
        """在知识库(wiki)指定父节点下创建一个新节点（默认 docx 文档）。
        返回 node 字典，含 node_token（wiki 链接用）与 obj_token（即 document_id）。
        需权限：wiki:wiki，且应用对该节点有编辑权（节点设为"组织内可编辑"即可）。"""
        body = {
            "obj_type": obj_type,
            "node_type": "origin",
            "title": title,
            "parent_node_token": parent_node_token,
        }
        resp = requests.post(
            f"{self.host}/open-apis/wiki/v2/spaces/{space_id}/nodes",
            headers={**self.headers, "Content-Type": "application/json"},
            json=body,
            timeout=60,
        )
        resp.raise_for_status()
        data = resp.json()
        if data.get("code") != 0:
            raise FeishuError(
                f"创建 wiki 节点失败: code={data.get('code')} msg={data.get('msg')}。"
                "常见原因：应用对该节点无编辑权（请把目标节点分享设为『组织内可编辑』），"
                "或缺少 wiki:wiki 权限。"
            )
        return data.get("data", {}).get("node", {})

    def create_block_descendant(self, document_id: str, parent_block_id: str,
                                children_ids: list, descendants: list,
                                index: int = -1) -> dict:
        """用 descendant 接口往文档插入带层级的块（支持表格等嵌套结构）。
        children_ids: 顶层块 id 列表（按顺序）。
        descendants: 全部块定义（顶层 + 后代）。
        index=-1 表示追加到末尾。
        需权限：docx:document。"""
        body = {"children_id": children_ids, "descendants": descendants}
        if index >= 0:
            body["index"] = index
        resp = requests.post(
            f"{self.host}/open-apis/docx/v1/documents/{document_id}"
            f"/blocks/{parent_block_id}/descendant",
            headers={**self.headers, "Content-Type": "application/json"},
            json=body,
            timeout=120,
        )
        resp.raise_for_status()
        data = resp.json()
        if data.get("code") != 0:
            raise FeishuError(
                f"写入文档块失败: code={data.get('code')} msg={data.get('msg')}"
            )
        return data.get("data", {})

    def set_table_column_width(self, document_id: str, table_block_id: str,
                               column_index: int, width: int) -> dict:
        """设置表格指定列的宽度(px)。飞书创建表格时忽略列宽，必须事后用本接口逐列设置。
        需权限：docx:document。"""
        body = {"update_table_property": {"column_index": column_index, "column_width": width}}
        resp = requests.patch(
            f"{self.host}/open-apis/docx/v1/documents/{document_id}/blocks/{table_block_id}",
            headers={**self.headers, "Content-Type": "application/json"},
            json=body,
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        if data.get("code") != 0:
            raise FeishuError(
                f"设置列宽失败: code={data.get('code')} msg={data.get('msg')}"
            )
        return data.get("data", {})

    def list_table_block_ids(self, document_id: str) -> list:
        """按文档顺序返回所有 table 块的 block_id 列表。
        用于写入后定位表格以设置列宽。"""
        blocks = self.list_all_blocks(document_id)
        return [b["block_id"] for b in blocks if b.get("block_type") == 31]

    # ---------- 通讯录：open_id -> 姓名 ----------
    def get_user_names(self, open_ids: list) -> dict:
        """批量把 open_id 换成姓名，返回 {open_id: name}。
        需权限：contact:user.base:readonly（或同等用户读权限），
        且目标用户需在应用通讯录可见范围内。
        查不到的 id 不会进结果（调用方自行兜底）。
        依赖 batch 接口：GET /open-apis/contact/v3/users/batch?user_ids=...&user_id_type=open_id
        每次最多 50 个。带进程内缓存。"""
        result = {}
        to_query = []
        for oid in open_ids:
            if not oid:
                continue
            if oid in self._user_name_cache:
                result[oid] = self._user_name_cache[oid]
            elif oid not in to_query:
                to_query.append(oid)

        for i in range(0, len(to_query), 50):
            batch = to_query[i:i + 50]
            try:
                resp = requests.get(
                    f"{self.host}/open-apis/contact/v3/users/batch",
                    headers=self.headers,
                    params=[("user_ids", b) for b in batch] + [("user_id_type", "open_id")],
                    timeout=30,
                )
                data = resp.json()
            except Exception:
                continue
            if data.get("code") != 0:
                # 权限不足/不可见等，整批跳过；不抛异常以免影响正文
                continue
            for u in data.get("data", {}).get("items", []) or []:
                oid = u.get("open_id")
                name = u.get("name") or u.get("nickname") or ""
                if oid and name:
                    self._user_name_cache[oid] = name
                    result[oid] = name
        return result

    # ---------- 云空间素材：下载图片 ----------
    def download_media(self, file_token: str) -> bytes:
        """下载云空间素材（docx 图片的 image.token 即素材 file_token）。
        需权限：drive:file:download（或 drive:drive:readonly）。
        返回原始二进制内容。失败抛 FeishuError。"""
        resp = requests.get(
            f"{self.host}/open-apis/drive/v1/medias/{file_token}/download",
            headers=self.headers,
            timeout=120,
        )
        # 成功时返回二进制；失败时返回 JSON 错误
        ctype = resp.headers.get("Content-Type", "")
        if resp.status_code == 200 and "application/json" not in ctype:
            return resp.content
        # 尝试解析错误信息
        try:
            data = resp.json()
            raise FeishuError(
                f"下载图片失败: code={data.get('code')} msg={data.get('msg')}。"
                "常见原因：未开通 drive:file:download 权限，或文档未授权给应用。"
            )
        except FeishuError:
            raise
        except Exception:
            raise FeishuError(f"下载图片失败: HTTP {resp.status_code}")
