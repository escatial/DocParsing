"""
MinerU DocParser 后端
- 作为 MinerU 云 API 的代理层，避免前端泄露 Token
- 暴露文件上传 -> 任务提交 -> 进度轮询 -> 结果下载
- 对 docx 文件做脚注后处理（[] -> Word 真实脚注）
"""
import asyncio
import base64
import io
import json
import os
import re
import time
import uuid
import urllib.parse
import zipfile

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn, nsmap
from docx.shared import Pt, RGBColor
from typing import Optional

import httpx
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse, JSONResponse
from pydantic import BaseModel

# ============ 配置 ============
MINERU_API_BASE = "https://mineru.net/api/v4"
# 从环境变量读取 Token，避免硬编码到仓库；未设置时使用默认值（MVP 阶段）
MINERU_TOKEN = os.getenv(
    "MINERU_TOKEN",
    "eyJ0eXBlIjoiSldUIiwiYWxnIjoiSFM1MTIifQ.eyJqdGkiOiI1NzQwMDU0MSIsInJvbCI6IlJPTEVfUkVHSVNURVIiLCJpc3MiOiJPcGVuWExhYiIsImlhdCI6MTc4MzY5NTc0NiwiY2xpZW50SWQiOiJsa3pkeDU3bnZ5MjJqa3BxOXgydyIsInBob25lIjoiMTg0MTkxNjQxMTgiLCJvcGVuSWQiOm51bGwsInV1aWQiOiJiMTdmYzBiZS02ZTk1LTQ1NzUtODdmMy1mNTk3ZGFmYjM1NzAiLCJlbWFpbCI6IiIsImV4cCI6MTc5MTQ3MTc0Nn0.knj2Tgcmn_DWgZr9PcX5OZhHeNl8AcE4yfr-BQxLpiDipVtAEmGfLawxtfpRl4GG1Y5HIA-0qx5EeOUQ8uSlRA",
)

# 内存任务存储（MVP 阶段，重启会丢失）
# 单文件任务：TASKS[internal_id] = {task_id, state, filename, ...}
TASKS: dict[str, dict] = {}
# 批量任务：BATCHES[internal_id] = {batch_id, files: [{internal_id, filename, size, state, full_zip_url, sub_stage}], ...}
BATCHES: dict[str, dict] = {}

# 批量接口上限
MAX_BATCH_FILES = 10

app = FastAPI(title="MinerU MVP", version="0.1.0")

# 允许前端跨域调用
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://127.0.0.1:5173"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ============ 数据模型 ============
# 支持的额外输出格式（markdown / json 为默认导出，无需指定）
SUPPORTED_EXTRA_FORMATS = {"docx", "html", "latex"}


class TaskInfo(BaseModel):
    # 关闭 Pydantic v2 的 model_ 保护命名空间，避免与字段名冲突
    model_config = {"protected_namespaces": ()}
    task_id: str
    state: str  # pending | running | done | failed | converting | cancelled
    extracted_pages: int = 0
    total_pages: int = 0
    full_zip_url: Optional[str] = None
    err_msg: str = ""
    model_version: str = "vlm"
    extra_formats: list[str] = []
    cancelled: bool = False  # 前端是否已取消
    sub_stage: str = ""  # 前端展示的子阶段（后端基于 state 推断）


class ParseResult(BaseModel):
    """解析结果：包含 Markdown、ZIP 中可用的额外格式、大纲、参考文献"""
    filename: str
    title: str = ""              # 文章标题（用于导出文件名）
    markdown: str
    files: list[dict] = []  # 兼容性字段（保留为空）
    available_formats: list[str] = []  # ZIP 内可用的格式（docx）
    # 从 content_list.json 提取的结构化信息
    toc: list[dict] = []         # 大纲 [{level, text, page?}]
    references: list[str] = []   # 参考文献列表（每条是独立字符串）


# ============ 批量解析数据模型 ============
class BatchFileInfo(BaseModel):
    """批量中单个文件的解析状态"""
    model_config = {"protected_namespaces": ()}
    internal_id: str          # 单文件内部 ID（用于单独访问）
    filename: str
    size: int
    state: str = "pending"    # pending | running | done | failed | cancelled
    sub_stage: str = "queued"
    extracted_pages: int = 0
    total_pages: int = 0
    full_zip_url: Optional[str] = None
    err_msg: str = ""
    cancelled: bool = False


class BatchTaskInfo(BaseModel):
    """批量任务整体状态"""
    internal_id: str          # 批次的内部 ID
    state: str = "pending"    # 整体状态：pending | running | done | failed | cancelled | partial
    file_count: int
    done_count: int = 0
    failed_count: int = 0
    cancelled_count: int = 0
    cancelled_all: bool = False  # 用户主动取消整批
    files: list[BatchFileInfo] = []
    created_at: float = 0.0


class BatchParseResponse(BaseModel):
    internal_id: str
    batch_id: str
    files: list[dict]  # [{ internal_id, filename, size }]


# ============ 统一异常处理 ============
@app.exception_handler(HTTPException)
async def http_handler(_, exc: HTTPException):
    return JSONResponse(
        status_code=exc.status_code,
        content={"code": exc.status_code, "msg": exc.detail},
    )


@app.exception_handler(Exception)
async def error_handler(_, exc: Exception):
    return JSONResponse(
        status_code=500,
        content={"code": 500, "msg": f"server error: {exc}"},
    )


# ============ 1. 申请上传链接 + 2. 上传文件 + 3. 提交解析任务 ============
async def upload_and_submit(
    file_bytes: bytes,
    filename: str,
    model_version: str,
    is_ocr: bool,
    enable_formula: bool,
    enable_table: bool,
    extra_formats: list[str],
) -> str:
    """申请上传链接 -> 上传文件 -> 返回 batch_id"""
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {MINERU_TOKEN}",
    }

    payload = {
        "enable_formula": enable_formula,
        "enable_table": enable_table,
        "files": [{"name": filename, "is_ocr": is_ocr}],
        "model_version": model_version,
    }
    # extra_formats 仅支持 docx/html/latex，且对 html 源文件无效
    if extra_formats:
        payload["extra_formats"] = extra_formats

    async with httpx.AsyncClient(timeout=60.0) as client:
        # 第一步：申请上传链接
        apply_resp = await client.post(
            f"{MINERU_API_BASE}/file-urls/batch",
            headers=headers,
            json=payload,
        )
        apply_resp.raise_for_status()
        apply_data = apply_resp.json()

        if apply_data.get("code") != 0:
            raise HTTPException(400, f"申请上传链接失败: {apply_data.get('msg')}")

        batch_id = apply_data["data"]["batch_id"]
        file_url = apply_data["data"]["file_urls"][0]

        # 第二步：上传文件二进制（不要设置 Content-Type）
        upload_resp = await client.put(file_url, content=file_bytes)
        if upload_resp.status_code != 200:
            raise HTTPException(500, f"文件上传失败: {upload_resp.status_code}")

        # 上传完成后 MinerU 会自动提交解析任务，batch_id 即任务标识
        return batch_id


async def submit_single_in_background(
    internal_id: str,
    content: bytes,
    filename: str,
    model_version: str,
    is_ocr: bool,
    enable_formula: bool,
    enable_table: bool,
):
    """后台提交单文件到 MinerU，避免前端卡在 /api/parse 请求上"""
    info = TASKS.get(internal_id)
    if not info or info.get("cancelled"):
        return
    try:
        batch_id = await upload_and_submit(
            content,
            filename or "document.pdf",
            model_version,
            is_ocr,
            enable_formula,
            enable_table,
            ["docx"],
        )
        info["task_id"] = batch_id
        info["state"] = "pending"
        info["err_msg"] = ""
        print(f"[submit_single_in_background] MinerU batch_id: {batch_id}")
    except Exception as e:
        print(f"[submit_single_in_background] MinerU 调用失败: {type(e).__name__}: {e}")
        info["state"] = "failed"
        info["err_msg"] = f"MinerU 调用失败: {e}"


@app.post("/api/parse")
async def parse_document(
    background: BackgroundTasks,
    file: UploadFile = File(...),
    model_version: str = Form("vlm", alias="model_version"),
    is_ocr: bool = Form(False),
    enable_formula: bool = Form(True),
    enable_table: bool = Form(True),
):
    """
    接收前端确认后提交的文件，转交给 MinerU 云端解析。
    返回一个内部 task_id，前端用它轮询进度。
    """
    print(f"[parse_document] 接收到文件: {file.filename}, model={model_version}")
    # 读取文件
    content = await file.read()
    print(f"[parse_document] 文件大小: {len(content)} bytes")
    if not content:
        raise HTTPException(400, "空文件")

    # 先记录到本地任务表并立即返回，MinerU 上传放到后台执行，避免前端卡在“上传文件”
    internal_id = str(uuid.uuid4())
    TASKS[internal_id] = {
        "task_id": "",
        "state": "pending",
        "extracted_pages": 0,
        "total_pages": 0,
        "full_zip_url": None,
        "err_msg": "",
        "model_version": model_version,
        "filename": file.filename,
        "created_at": time.time(),
        "extra_formats": ["docx"],
    }
    background.add_task(
        submit_single_in_background,
        internal_id,
        content,
        file.filename or "document.pdf",
        model_version,
        is_ocr,
        enable_formula,
        enable_table,
    )
    return {"internal_id": internal_id, "batch_id": ""}


# ============ 1.5 批量上传并解析（1-10 个文件） ============
@app.post("/api/parse/batch", response_model=BatchParseResponse)
async def parse_batch(
    files: list[UploadFile] = File(..., description="1-10 个文件"),
    model_version: str = Form("vlm", alias="model_version"),
    is_ocr: bool = Form(False),
    enable_formula: bool = Form(True),
    enable_table: bool = Form(True),
):
    """
    批量解析：一次提交多个文件，MinerU 共享一个 batch_id 统一调度。
    前端拿到每个文件的 internal_id 后可独立查询/取消/下载。
    """
    if not files:
        raise HTTPException(400, "未提供文件")
    if len(files) > MAX_BATCH_FILES:
        raise HTTPException(400, f"批量最多 {MAX_BATCH_FILES} 个文件，当前 {len(files)} 个")
    if len(files) == 1:
        raise HTTPException(400, "单文件请使用 /api/parse 接口")

    # 客户端校验：大小/类型
    allowed_ext = [
        "pdf", "png", "jpg", "jpeg", "jp2", "webp", "gif", "bmp",
        "doc", "docx", "ppt", "pptx", "xls", "xlsx", "html",
    ]
    for f in files:
        if f.size and f.size > 200 * 1024 * 1024:
            raise HTTPException(400, f"文件 {f.filename} 超过 200MB")
        ext = (f.filename or "").split(".")[-1].lower()
        if ext not in allowed_ext:
            raise HTTPException(400, f"文件 {f.filename} 类型不支持")

    print(f"[parse_batch] 收到批量: {len(files)} 个文件, model={model_version}")

    # 读取所有文件
    file_blobs: list[tuple[bytes, str]] = []
    for f in files:
        content = await f.read()
        if not content:
            raise HTTPException(400, f"文件 {f.filename} 为空")
        file_blobs.append((content, f.filename or "document.pdf"))

    # 调用 MinerU 批量接口
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {MINERU_TOKEN}",
    }
    payload = {
        "enable_formula": enable_formula,
        "enable_table": enable_table,
        "files": [
            {"name": fname, "is_ocr": is_ocr} for _, fname in file_blobs
        ],
        "model_version": model_version,
        # 默认开启 docx 输出，让 ZIP 包含 docx 文件
        "extra_formats": ["docx"],
    }

    async with httpx.AsyncClient(timeout=120.0) as client:
        apply_resp = await client.post(
            f"{MINERU_API_BASE}/file-urls/batch", headers=headers, json=payload,
        )
        apply_resp.raise_for_status()
        apply_data = apply_resp.json()

        if apply_data.get("code") != 0:
            raise HTTPException(400, f"申请上传链接失败: {apply_data.get('msg')}")

        batch_id = apply_data["data"]["batch_id"]
        upload_urls = apply_data["data"]["file_urls"]

        if len(upload_urls) != len(file_blobs):
            raise HTTPException(500, "MinerU 返回的上传链接数量与文件数不匹配")

        # 并行上传所有文件
        async def upload_one(idx: int, content: bytes, url: str):
            r = await client.put(url, content=content)
            return idx, r.status_code

        results = await asyncio.gather(*[
            upload_one(i, content, upload_urls[i])
            for i, (content, _) in enumerate(file_blobs)
        ])
        for idx, status in results:
            if status != 200:
                fname = file_blobs[idx][1]
                raise HTTPException(500, f"文件 {fname} 上传失败: HTTP {status}")

    # 构建本地批量任务
    batch_internal_id = str(uuid.uuid4())
    file_records = []
    for i, (content, fname) in enumerate(file_blobs):
        file_internal_id = str(uuid.uuid4())
        file_records.append({
            "internal_id": file_internal_id,
            "filename": fname,
            "size": len(content),
            "state": "pending",
            "sub_stage": "queued",
            "extracted_pages": 0,
            "total_pages": 0,
            "full_zip_url": None,
            "err_msg": "",
            "cancelled": False,
        })

    BATCHES[batch_internal_id] = {
        "batch_id": batch_id,
        "model_version": model_version,
        "files": file_records,
        "cancelled_all": False,
        "created_at": time.time(),
    }

    return BatchParseResponse(
        internal_id=batch_internal_id,
        batch_id=batch_id,
        files=[
            {"internal_id": r["internal_id"], "filename": r["filename"], "size": r["size"]}
            for r in file_records
        ],
    )


# ============ 4. 查询任务进度 ============
def infer_sub_stage(state: str, extracted: int, total: int) -> str:
    """根据 MinerU 状态推断当前子阶段（前端展示用）"""
    if state == "pending":
        return "queued"        # 排队等待
    if state == "converting":
        return "assembling"    # 结果组装
    if state == "running":
        # 用进度比例粗略划分：前 80% 是页面解析，后面是模型推理
        if total > 0 and extracted >= total * 0.8:
            return "inferencing"   # 模型推理
        return "parsing_page"      # 页面解析（OCR/版面/切分）
    if state == "done":
        return "done"
    if state == "failed":
        return "failed"
    if state == "cancelled":
        return "cancelled"
    return "unknown"


@app.get("/api/task/{internal_id}", response_model=TaskInfo)
async def get_task(internal_id: str):
    """前端轮询该接口获取最新状态"""
    info = TASKS.get(internal_id)
    if not info:
        # 任务未注册，可能是后端重启过；返回 pending 而不是 404
        return TaskInfo(
            task_id="",
            state="pending",
            extracted_pages=0,
            total_pages=0,
            full_zip_url=None,
            err_msg="任务尚未注册或已被清理",
            model_version="vlm",
            extra_formats=[],
            sub_stage="queued",
        )

    # 后台尚未拿到 MinerU batch_id：保持 pending，不阻塞前端
    if not info.get("task_id"):
        return TaskInfo(
            task_id="",
            state=info.get("state", "pending"),
            extracted_pages=info.get("extracted_pages", 0),
            total_pages=info.get("total_pages", 0),
            full_zip_url=info.get("full_zip_url"),
            err_msg=info.get("err_msg") or "正在上传到 MinerU，请稍候",
            model_version=info.get("model_version", "vlm"),
            extra_formats=info.get("extra_formats", ["docx"]),
            sub_stage=infer_sub_stage(info.get("state", "pending"), 0, 0),
        )

    # 已取消的任务：直接返回 cancelled，不再轮询 MinerU
    if info.get("cancelled"):
        return TaskInfo(
            task_id=info["task_id"],
            state="cancelled",
            extracted_pages=info.get("extracted_pages", 0),
            total_pages=info.get("total_pages", 0),
            full_zip_url=info.get("full_zip_url"),
            err_msg="任务已被前端取消",
            model_version=info["model_version"],
            extra_formats=info.get("extra_formats", []),
            cancelled=True,
            sub_stage="cancelled",
        )

    headers = {"Authorization": f"Bearer {MINERU_TOKEN}"}

    try:
        async with httpx.AsyncClient(timeout=30.0) as client:
            # 批量上传 API 必须用 /extract-results/batch/{batch_id} 查询
            resp = await client.get(
                f"{MINERU_API_BASE}/extract-results/batch/{info['task_id']}",
                headers=headers,
            )
            resp.raise_for_status()
            data = resp.json()

        if data.get("code") != 0:
            print(f"[get_task] MinerU 业务错误: {data.get('msg')}")
            return TaskInfo(
                task_id=info["task_id"],
                state=info["state"],
                extracted_pages=info["extracted_pages"],
                total_pages=info["total_pages"],
                full_zip_url=info["full_zip_url"],
                err_msg=data.get("msg", "查询失败"),
                model_version=info["model_version"],
                extra_formats=info.get("extra_formats", []),
                sub_stage=infer_sub_stage(
                    info["state"], info["extracted_pages"], info["total_pages"]
                ),
            )

        # 批量接口返回 extract_result 列表，取第一个文件的结果
        d = data["data"]
        results = d.get("extract_result") or []
        if not results:
            return TaskInfo(
                task_id=info["task_id"],
                state="pending",
                extracted_pages=0,
                total_pages=0,
                full_zip_url=None,
                err_msg="任务尚未生成结果",
                model_version=info["model_version"],
                extra_formats=info.get("extra_formats", []),
                sub_stage="queued",
            )

        item = results[0]
        info["state"] = item.get("state", info["state"])
        info["full_zip_url"] = item.get("full_zip_url") or info["full_zip_url"]
        info["err_msg"] = item.get("err_msg", "")
        progress = item.get("extract_progress") or {}
        info["extracted_pages"] = progress.get("extracted_pages", 0)
        info["total_pages"] = progress.get("total_pages", 0)
    except httpx.HTTPError as e:
        print(f"[get_task] HTTP 错误: {type(e).__name__}: {e}")
        info["err_msg"] = f"查询 MinerU 失败: {e}"

    return TaskInfo(
        task_id=info["task_id"],
        state=info["state"],
        extracted_pages=info["extracted_pages"],
        total_pages=info["total_pages"],
        full_zip_url=info["full_zip_url"],
        err_msg=info["err_msg"],
        model_version=info["model_version"],
        extra_formats=info.get("extra_formats", []),
        sub_stage=infer_sub_stage(
            info["state"], info["extracted_pages"], info["total_pages"]
        ),
    )


# ============ 4.5 取消任务 ============
@app.post("/api/task/{internal_id}/cancel")
async def cancel_task(internal_id: str):
    """
    标记任务为已取消。前端停止轮询，MinerU 任务仍会在云端跑完（API 不支持真取消）。
    支持单文件（TASKS）和单文件（批量中的某个，BATCHES[*].files）。
    """
    # 先查 TASKS
    info = TASKS.get(internal_id)
    if info:
        if info["state"] in ("done", "failed"):
            return {"ok": True, "msg": f"任务已处于终态: {info['state']}"}
        info["cancelled"] = True
        info["state"] = "cancelled"
        return {"ok": True, "msg": "已取消"}

    # 再查 BATCHES 中的单个文件
    for batch in BATCHES.values():
        for f in batch["files"]:
            if f["internal_id"] == internal_id:
                if f["state"] in ("done", "failed"):
                    return {"ok": True, "msg": f"任务已处于终态: {f['state']}"}
                f["cancelled"] = True
                f["state"] = "cancelled"
                return {"ok": True, "msg": "已取消该文件"}

    raise HTTPException(404, "任务不存在")


# ============ 4.6 批量任务：查询整体进度 ============
@app.get("/api/batch/{batch_internal_id}", response_model=BatchTaskInfo)
async def get_batch(batch_internal_id: str):
    """轮询批量任务：聚合每个文件状态 + 调用 MinerU 查询最新进度"""
    batch = BATCHES.get(batch_internal_id)
    if not batch:
        return BatchTaskInfo(
            internal_id=batch_internal_id,
            state="pending",
            file_count=0,
        )

    # 整批已取消：不再调用 MinerU
    if batch.get("cancelled_all"):
        return BatchTaskInfo(
            internal_id=batch_internal_id,
            state="cancelled",
            file_count=len(batch["files"]),
            done_count=sum(1 for f in batch["files"] if f["state"] == "done"),
            failed_count=sum(1 for f in batch["files"] if f["state"] == "failed"),
            cancelled_count=sum(1 for f in batch["files"] if f["state"] == "cancelled"),
            cancelled_all=True,
            files=[BatchFileInfo(**f) for f in batch["files"]],
            created_at=batch.get("created_at", 0.0),
        )

    # 调用 MinerU 批量查询接口
    headers = {"Authorization": f"Bearer {MINERU_TOKEN}"}
    try:
        async with httpx.AsyncClient(timeout=30.0) as client:
            resp = await client.get(
                f"{MINERU_API_BASE}/extract-results/batch/{batch['batch_id']}",
                headers=headers,
            )
            resp.raise_for_status()
            data = resp.json()

        if data.get("code") == 0:
            results = data["data"].get("extract_result") or []
            # 按文件名匹配本地记录（顺序与提交顺序一致）
            for i, item in enumerate(results):
                if i >= len(batch["files"]):
                    break
                f = batch["files"][i]
                if f["cancelled"]:
                    continue
                f["state"] = item.get("state", f["state"])
                f["full_zip_url"] = item.get("full_zip_url") or f["full_zip_url"]
                f["err_msg"] = item.get("err_msg", "")
                progress = item.get("extract_progress") or {}
                f["extracted_pages"] = progress.get("extracted_pages", 0)
                f["total_pages"] = progress.get("total_pages", 0)
                f["sub_stage"] = infer_sub_stage(
                    f["state"], f["extracted_pages"], f["total_pages"]
                )
    except httpx.HTTPError as e:
        print(f"[get_batch] HTTP 错误: {e}")

    # 统计整体状态
    done_n = sum(1 for f in batch["files"] if f["state"] == "done")
    failed_n = sum(1 for f in batch["files"] if f["state"] == "failed")
    cancelled_n = sum(1 for f in batch["files"] if f["state"] == "cancelled")
    total = len(batch["files"])

    if done_n == total:
        overall = "done"
    elif failed_n == total:
        overall = "failed"
    elif done_n + failed_n + cancelled_n == total:
        overall = "partial"  # 有成功有失败
    elif cancelled_n > 0 and done_n + failed_n + cancelled_n == total:
        overall = "partial"
    elif cancelled_n == total:
        overall = "cancelled"
    else:
        overall = "running" if any(f["state"] in ("running", "converting", "pending") for f in batch["files"]) else "pending"

    return BatchTaskInfo(
        internal_id=batch_internal_id,
        state=overall,
        file_count=total,
        done_count=done_n,
        failed_count=failed_n,
        cancelled_count=cancelled_n,
        cancelled_all=batch.get("cancelled_all", False),
        files=[BatchFileInfo(**f) for f in batch["files"]],
        created_at=batch.get("created_at", 0.0),
    )


# ============ 4.7 取消整批任务 ============
@app.post("/api/batch/{batch_internal_id}/cancel")
async def cancel_batch(batch_internal_id: str):
    """取消整批任务，标记所有未完成文件为 cancelled"""
    batch = BATCHES.get(batch_internal_id)
    if not batch:
        raise HTTPException(404, "批量任务不存在")
    batch["cancelled_all"] = True
    for f in batch["files"]:
        if f["state"] not in ("done", "failed"):
            f["cancelled"] = True
            f["state"] = "cancelled"
    return {"ok": True, "msg": f"已取消整批 {len(batch['files'])} 个文件"}


# ============ 5. 下载解析结果（提取 Markdown） ============
def _find_file_info(internal_id: str) -> Optional[dict]:
    """统一查找：单文件 TASKS 或批量 BATCHES 中的单个文件"""
    info = TASKS.get(internal_id)
    if info:
        return info
    for batch in BATCHES.values():
        for f in batch["files"]:
            if f["internal_id"] == internal_id:
                # 构造一个伪 info 字典供下游接口使用
                return {
                    "task_id": batch["batch_id"],
                    "state": f["state"],
                    "filename": f["filename"],
                    "full_zip_url": f["full_zip_url"],
                    "model_version": batch.get("model_version", "vlm"),
                    "internal_id": f["internal_id"],
                    "_is_batch_file": True,
                    "_batch_file": f,
                }
    return None


# 重新写 download_result，让它也支持批量文件
@app.get("/api/task/{internal_id}/download")
async def download_result(internal_id: str):
    """支持单文件 TASKS 和批量 BATCHES 中的单个文件"""
    info = _find_file_info(internal_id)
    if not info:
        raise HTTPException(404, "任务不存在")
    if info["state"] != "done" or not info["full_zip_url"]:
        raise HTTPException(400, f"任务未完成，当前状态: {info['state']}")

    async with httpx.AsyncClient(timeout=120.0) as client:
        resp = await client.get(info["full_zip_url"])
        resp.raise_for_status()
        zip_bytes = resp.content

    md_content = ""
    available_formats: list[str] = []
    toc: list[dict] = []
    references: list[str] = []
    try:
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            content_list_data = None
            for name in zf.namelist():
                if name.endswith("full.md"):
                    md_content = zf.read(name).decode("utf-8", errors="replace")
                for fmt in SUPPORTED_EXTRA_FORMATS:
                    if name.endswith(f".{fmt}") and fmt not in available_formats:
                        available_formats.append(fmt)
                # content_list.json：MinerU 的结构化输出（仅用于大纲提取，参考文献用 markdown 解析更可靠）
                if name.endswith("_content_list.json") and content_list_data is None:
                    try:
                        raw = zf.read(name).decode("utf-8", errors="replace")
                        content_list_data = json.loads(raw)
                    except Exception:
                        content_list_data = None

            # 优先从 markdown 提取大纲和参考文献（更可靠，能精确定位章节边界）
            toc = _extract_toc_from_markdown(md_content) if md_content else []
            references = _extract_refs_from_markdown(md_content) if md_content else []

            # 如果 markdown 解析失败，再尝试 content_list 兜底（仅作为补充）
            if not toc and content_list_data is not None:
                items = content_list_data if isinstance(content_list_data, list) else (
                    content_list_data.get("items") if isinstance(content_list_data, dict) else []
                )
                toc, _ = _extract_toc_and_refs(items)
    except Exception as e:
        md_content = md_content or f"[ZIP 解析失败: {e}]"

    # 提取文章标题
    title = _extract_title_from_markdown(md_content) if md_content else ""

    return ParseResult(
        filename=info.get("filename", "result"),
        title=title,
        markdown=md_content,
        files=[],
        available_formats=available_formats,
        toc=toc,
        references=references,
    )


def _extract_toc_and_refs(items: list) -> tuple[list[dict], list[str]]:
    """
    从 MinerU content_list.json 中提取大纲和参考文献。
    参考文献提取规则：
      - type == reference / references / bibliography
      - 或者 text 以 [N] / N. 开头（heuristic fallback）
    """
    toc: list[dict] = []
    refs: list[str] = []

    # 第一轮：标准 type 匹配
    for it in items:
        if not isinstance(it, dict):
            continue
        text = (it.get("text") or "").strip()
        item_type = (it.get("type") or "").lower()
        level = it.get("level") or it.get("heading_level") or 1

        if item_type in ("toc", "outline", "heading"):
            toc.append({"level": int(level) if str(level).isdigit() else 1, "text": text})
            continue

        if item_type in ("reference", "references", "bibliography"):
            if text:
                refs.append(text)

    # 第二轮：如果 refs 为空但有 [N] 开头的项，按编号拼接
    if not refs:
        for it in items:
            if not isinstance(it, dict):
                continue
            text = (it.get("text") or "").strip()
            if re.match(r"^\s*\[?\d+\]?[\.\s]\S", text):
                refs.append(text)

    return toc, refs


def _extract_title_from_markdown(md: str) -> str:
    """
    从 Markdown 中提取文章标题（多策略，按优先级尝试）:
    1. YAML frontmatter 中的 title 字段
    2. 第一个 `# 一级标题`
    3. 第一个 `## 二级标题`
    4. "Title: xxx" / "题目：" / "标题：" 等显式标记
    5. 文档首段（首段连续非空行）
    6. Markdown 第一行非空文本（兜底）

    长度限制 4-150 字符，排除表格行、列表项、URL、邮箱等。
    """
    if not md:
        return ""

    lines = md.split("\n")

    # 策略 1：YAML frontmatter
    if lines and lines[0].strip() == "---":
        for i in range(1, min(20, len(lines))):
            if lines[i].strip() == "---":
                break
            m = re.match(r"^\s*title:\s*(.+)$", lines[i])
            if m:
                t = m.group(1).strip().strip("'\"")
                if _is_valid_title(t):
                    return t

    # 策略 2：第一个 # 一级标题
    for line in lines:
        line = line.strip()
        if not line or line.startswith("```"):
            continue
        if line.startswith("# "):
            title = line[2:].strip()
            title = re.sub(r"\s*#+\s*$", "", title)
            title = re.sub(r"[*_`]+", "", title)
            if _is_valid_title(title):
                return title
            continue

    # 策略 3：第一个 ## 二级标题（仅当文档不含一级标题时）
    for line in lines:
        line = line.strip()
        if not line or line.startswith("```"):
            continue
        if line.startswith("## "):
            title = line[3:].strip()
            title = re.sub(r"\s*#+\s*$", "", title)
            title = re.sub(r"[*_`]+", "", title)
            if _is_valid_title(title):
                return title
            continue

    # 策略 4：显式标题标记
    title_marker = re.compile(
        r"^(?:Title|题目|标题|题\s*目|Title\s*[:：])\s*[:：]?\s*(.+)$",
        re.I,
    )
    for line in lines[:30]:
        m = title_marker.match(line.strip())
        if m:
            t = m.group(1).strip().strip("'\"").strip()
            if _is_valid_title(t):
                return t

    # 策略 5：文档首段（连续非空行合并）
    # 这处理学术论文常把标题放在前面几行不间断段落中的情况
    first_block_lines: list[str] = []
    for line in lines[:20]:
        stripped = line.strip()
        if not stripped:
            break  # 空行 = 段落结束
        first_block_lines.append(stripped)

    # 优先：只取首段中**第一个**像标题的行（避免取到摘要或作者）
    for line in first_block_lines[:3]:
        if _looks_like_academic_title(line):
            return _clean_title(line)

    # 兜底：首段合并
    if first_block_lines:
        joined = " ".join(first_block_lines).strip()
        # 截取到第一个标点（标题通常到第一个句号或冒号前）
        # 但中文标题没有句号，所以只截取到合理长度
        if 4 <= len(joined) <= 150:
            # 排除纯装饰字符（纯 #、*、= 等）
            if re.search(r"[一-龥A-Za-z0-9]", joined):
                return joined
        # 太长：截取前 60 字符
        candidate = joined[:60].strip()
        if 4 <= len(candidate) <= 150 and re.search(r"[一-龥A-Za-z0-9]", candidate):
            return candidate

    return ""


def _looks_like_academic_title(line: str) -> bool:
    """
    判断一行是否像学术论文标题。
    特征：
    - 中等长度（10-100 字符）
    - 不包含明显的元数据标记（@, 网址, DOI, 作者后跟逗号+年份）
    - 通常是中文/英文的短语（不含段末句号后的多余内容）
    """
    line = line.strip()
    if not (8 <= len(line) <= 100):
        return False
    # 排除明显不是标题的内容
    exclude_patterns = [
        r"^摘\s*要",            # 摘要
        r"^Abstract",            # Abstract
        r"^关键词",               # 关键词
        r"^Keywords",             # Keywords
        r"^作者",                  # 作者
        r"@\w",                    # 邮箱
        r"https?://",              # URL
        r"\d{4}\s*年",             # 2020 年
        r"DOI[:：]",
        r"^第\s*\d+\s*[卷期章]",
        r"^#+$",                   # 纯 # 装饰
        r"^[*_~=]+$",              # 纯 * _ ~ = 装饰
    ]
    for p in exclude_patterns:
        if re.search(p, line, re.I):
            return False
    # 必须包含字母或中文（不然全是标点）
    if not re.search(r"[一-龥A-Za-z]", line):
        return False
    return True


def _clean_title(t: str) -> str:
    """清理标题：去掉装饰字符"""
    t = re.sub(r"[*_`]+", "", t)
    # 中文括号
    t = re.sub(r"^[【《「]\s*", "", t)
    t = re.sub(r"\s*[】》」]$", "", t)
    # ASCII 方括号（只在标题首尾存在时去掉）
    t = re.sub(r"^\s*\[\s*", "", t)
    t = re.sub(r"\s*\]\s*$", "", t)
    return t.strip()


def _is_valid_title(t: str) -> bool:
    """判断字符串是否像标题"""
    if not t:
        return False
    t = t.strip()
    # 太长或太短都不是
    if len(t) < 4 or len(t) > 150:
        return False
    # 排除明显不是标题的内容
    if t.startswith("|") or t.endswith("|"):  # 表格
        return False
    if t.startswith("- ") or t.startswith("* ") or t.startswith("> "):
        return False
    # 全是标点或装饰字符（#、*、~、=、-、_）
    if re.fullmatch(r"[\s#*_~=~\-]+", t):
        return False
    return True


def _extract_toc_from_markdown(md: str) -> list[dict]:
    """兜底：从 Markdown 中按 ATX 标题提取大纲"""
    toc: list[dict] = []
    for line in md.split("\n"):
        m = re.match(r"^(#{1,6})\s+(.+?)\s*$", line)
        if m:
            level = len(m.group(1))
            text = m.group(2).strip()
            # 过滤代码块里的 # 标题
            if text and not text.startswith("```"):
                toc.append({"level": level, "text": text})
    return toc


def _extract_refs_from_markdown(md: str) -> list[str]:
    """
    从 Markdown 中提取参考文献段落。
    支持两种情况：
    1. 标准结构：Markdown 中存在 `## 参考文献` 等章节标题
    2. 无章节标题：扫描全文独立的编号行条目（不与正文 [N] 引用混在一起）

    编号格式支持：
    - `[1] xxx` / `[1]. xxx`
    - `1. xxx` / `1) xxx`
    - 跨多行条目（自动合并相邻非编号行）
    """
    refs: list[str] = []

    # 更宽松的章节标题识别
    ref_header_pattern = re.compile(
        r"^#{1,3}\s*(参考文献|References?|Bibliography|引用文献|REFERENCES|参考文\s*献|R\s*E\s*F\s*E\s*R\s*E\s*N\s*C\s*E\s*S)",
        re.I,
    )
    # 编号行模式：以 [N] 或 N. 开头的独立行
    num_pattern = re.compile(r"^\s*\[(\d{1,3})\][\.\s]*(.+)$")
    # 普通编号模式：1. xxx 或 1) xxx
    plain_num_pattern = re.compile(r"^\s*(\d{1,3})[\.\)]\s+(.+)$")
    # 下一个标题模式
    next_header_pattern = re.compile(r"^#{1,3}\s+\S")

    # ============ 阶段 1：尝试定位「参考文献」章节 ============
    in_ref_section = False
    current_ref_lines: list[str] = []
    current_num: int | None = None

    def flush_current(into_list: list[str]):
        """把累积的多行合并成一条 ref"""
        nonlocal current_ref_lines, current_num
        if current_num is not None and current_ref_lines:
            full = " ".join(s for s in current_ref_lines if s).strip()
            full = re.sub(r"\s+", " ", full)
            if full:
                into_list.append(full)
        current_ref_lines = []
        current_num = None

    lines = md.split("\n")
    for line in lines:
        stripped = line.strip()

        if ref_header_pattern.match(stripped):
            in_ref_section = True
            flush_current(refs)
            continue

        if in_ref_section and next_header_pattern.match(stripped):
            flush_current(refs)
            in_ref_section = False
            continue

        if not in_ref_section:
            continue

        # 章节内：跳过空行（不切断）
        if not stripped:
            if current_num is not None:
                flush_current(refs)
            continue

        # 匹配 [1] xxx 格式
        m = num_pattern.match(stripped)
        if not m:
            m = plain_num_pattern.match(stripped)
        if m:
            flush_current(refs)
            current_num = int(m.group(1))
            body = m.group(2).strip()
            if body:
                current_ref_lines.append(body)
        else:
            # 非编号行视为上一条目延续
            if current_num is not None:
                current_ref_lines.append(stripped)

    flush_current(refs)

    # ============ 阶段 2：如果章节法没找到，扫描全文独立编号条目 ============
    if not refs:
        # 收集文档中所有 [N] 引用（正文中）和独立的 [N] xxx 编号行
        # 启发式：匹配以 [N] 开头的非引用行（长度 > 30 字符或包含期刊标记）
        candidate_pattern = re.compile(r"^\s*\[(\d{1,3})\][\s\.]*(.{20,})$")
        seen_nums: set[int] = set()
        for line in lines:
            stripped = line.strip()
            # 跳过标题章节
            if stripped.startswith("#"):
                continue
            m = candidate_pattern.match(stripped)
            if m:
                num = int(m.group(1))
                content = m.group(2).strip()
                # 排除明显的正文引用（短文本 + 通常以"，"、"。" 结尾的引文）
                # 启发规则：参考文献通常比较长，包含期刊标记 [J] / [M] / [C] 或出版社信息
                is_likely_ref = (
                    len(content) > 20
                    or bool(re.search(r"\[\w\]|\(\d{4}\)|\d{4}\s*年|\d+\(\d+\)|et al\.|出版社|Univ", content))
                )
                if is_likely_ref and num not in seen_nums:
                    refs.append(content)
                    seen_nums.add(num)

    return refs[:50]  # 最多 50 条防溢出


# ============ Docx 脚注后处理 ============
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _parse_inline_references(md: str) -> list[dict]:
    """从 Markdown 中解析行内引用 [1] [2] 及其位置"""
    refs = []
    pattern = re.compile(r"\[(\d+)\]")
    for m in pattern.finditer(md):
        refs.append({"num": int(m.group(1)), "raw": m.group(0)})
    # 去重保留首次出现顺序
    seen = set()
    unique = []
    for r in refs:
        if r["num"] not in seen:
            seen.add(r["num"])
            unique.append(r)
    return unique


def _extract_refs_dict_from_md(md: str) -> dict[int, str]:
    """从 Markdown 末尾的参考文献列表中提取 {num: text}"""
    refs_dict: dict[int, str] = {}
    in_refs = False
    pattern = re.compile(r"^\s*\[(\d+)\][\.\s]*(.+)$")
    for line in md.split("\n"):
        stripped = line.strip()
        if re.match(r"^#{1,3}\s*(参考文献|References?|Bibliography|引用文献|REFERENCES)", stripped, re.I):
            in_refs = True
            continue
        if in_refs and re.match(r"^#{1,3}\s+", stripped):
            in_refs = False
        if in_refs:
            m = pattern.match(stripped)
            if m:
                refs_dict[int(m.group(1))] = m.group(2).strip()
    return refs_dict


def _docx_bracket_refs(text: str) -> list[dict]:
    """
    在 docx 段落文本里扫描所有 [...] 引用，返回 [{start, end, nums}] 列表。
    支持：[N] / [N-M] / [N,M,K] / [N－M] / [N—M] / [N~M]
    返回的 nums 列表每个数字都是 references 字典里要查找的编号。
    """
    out: list[dict] = []
    bracket_re = re.compile(
        r"\[\s*"
        r"(\d+)"
        r"("
        r"(?:\s*[\-－\u2013\u2014~]\s*\d+)"
        r"|"
        r"(?:\s*[,，\uff0c]\s*\d+)+"
        r")*"
        r"\s*\]"
    )
    for m in bracket_re.finditer(text):
        body = m.group(0)
        nums = [int(x) for x in re.findall(r"\d+", body)]
        if not nums:
            continue
        if re.search(r"[\-－\u2013\u2014~]", body):
            start, end = nums[0], nums[-1]
            if end >= start:
                nums = list(range(start, end + 1))
            else:
                nums = [start]
        else:
            seen = set()
            dedup = []
            for n in nums:
                if n not in seen:
                    dedup.append(n)
                    seen.add(n)
            nums = dedup
        out.append({"start": m.start(), "end": m.end(), "nums": nums})
    return out


def _process_docx_with_footnotes(docx_bytes: bytes, references: list[str]) -> bytes:
    """
    将 MinerU 生成的 docx 中的 [N]/[N-M]/[N,M,K] 引用转换为 Word 真实脚注。

    参数:
        docx_bytes: 原始 docx 二进制（来自 MinerU）
        references: 参考文献列表

    策略：
    1. 把 references 解析成 {num: text} 字典
    2. 在 docx 末尾插入 Footnote XML 部件
    3. 把正文中每个 [N] / [N-M] / [N,M,K] 文本替换为 footnoteReference
       - [5-7] 展开为 3 个 footnoteReference（5, 6, 7）
       - [1,3,5] 展开为 3 个 footnoteReference
    4. 在 footnotes.xml 中添加对应的 footnote 节点
    """
    try:
        doc = Document(io.BytesIO(docx_bytes))
    except Exception as e:
        print(f"[process_docx_footnotes] 打开 docx 失败: {e}")
        return docx_bytes

    # 收集 references 字典：编号 -> 文本
    refs_dict: dict[int, str] = {}
    for idx, ref in enumerate(references):
        ref = ref.strip()
        if not ref:
            continue
        m = re.match(r"^\s*\[?(\d+)\]?[\.\s]*(.+)$", ref)
        if m:
            num = int(m.group(1))
            text = m.group(2).strip()
        else:
            num = idx + 1
            text = ref
        refs_dict[num] = text

    if not refs_dict:
        print("[process_docx_footnotes] 没有可用的参考文献，不处理")
        return docx_bytes

    print(f"[process_docx_footnotes] 解析到 {len(refs_dict)} 条参考文献，编号: {sorted(refs_dict.keys())[:10]}...")

    body = doc.element.body
    fn_id = 100
    inserted: list[int] = []

    for p in body.iter(qn("w:p")):
        runs_with_text = []
        for r in p.findall(qn("w:r")):
            for t in r.findall(qn("w:t")):
                runs_with_text.append((r, t, t.text or ""))
        if not runs_with_text:
            continue

        # 对每个 w:t 节点单独处理：[N-M] 这种范围在一个 w:t 里要完整替换
        for r, t, txt in runs_with_text:
            new_text = txt
            bracket_matches = _docx_bracket_refs(new_text)
            if not bracket_matches:
                continue

            parent_run = t.getparent()
            parent_para = parent_run.getparent()
            run_index = list(parent_para).index(parent_run)

            new_runs_to_insert: list = []
            cursor_text = 0
            for bm in bracket_matches:
                pre = new_text[cursor_text:bm["start"]]
                if pre:
                    pre_run = OxmlElement("w:r")
                    pre_t = OxmlElement("w:t")
                    pre_t.text = pre
                    pre_t.set(qn("xml:space"), "preserve")
                    pre_run.append(pre_t)
                    new_runs_to_insert.append(pre_run)
                # 对每个有效 num 插入一个 footnoteReference
                valid_nums = [n for n in bm["nums"] if n in refs_dict]
                for num in valid_nums:
                    if num in inserted:
                        continue
                    fn_run = OxmlElement("w:r")
                    rpr = OxmlElement("w:rPr")
                    rstyle = OxmlElement("w:rStyle")
                    rstyle.set(qn("w:val"), "FootnoteReference")
                    rpr.append(rstyle)
                    # 部分 Word/WPS 不会仅凭 FootnoteReference 样式自动显示上标
                    vert_align = OxmlElement("w:vertAlign")
                    vert_align.set(qn("w:val"), "superscript")
                    rpr.append(vert_align)
                    fn_run.append(rpr)
                    fn_ref = OxmlElement("w:footnoteReference")
                    fn_ref.set(qn("w:id"), str(fn_id))
                    fn_run.append(fn_ref)
                    new_runs_to_insert.append(fn_run)
                    inserted.append(num)
                    fn_id += 1
                cursor_text = bm["end"]

            rest = new_text[cursor_text:]
            if rest:
                rest_run = OxmlElement("w:r")
                rest_t = OxmlElement("w:t")
                rest_t.text = rest
                rest_t.set(qn("xml:space"), "preserve")
                rest_run.append(rest_t)
                new_runs_to_insert.append(rest_run)
            t.text = ""
            for offset, nr in enumerate(new_runs_to_insert):
                parent_para.insert(run_index + 1 + offset, nr)
            if not any(rt.text for rt in parent_run.findall(qn("w:t"))):
                parent_para.remove(parent_run)

    if not inserted:
        print("[process_docx_footnotes] 文档中未找到任何 [N]/[N-M]/[N,M,K] 引用，未处理")
        return docx_bytes

    print(f"[process_docx_footnotes] 已插入 {len(inserted)} 个脚注引用，编号: {inserted[:10]}...")

    # 注册 footnotes part
    _ensure_footnotes_part(doc, refs_dict, inserted)

    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()


def _ensure_footnotes_part(doc, refs_dict: dict[int, str], inserted: list[int]):
    """在 docx 中注册 footnotes.xml 部件，添加脚注内容"""
    from docx.opc.constants import RELATIONSHIP_TYPE as RT, CONTENT_TYPE as CT
    from docx.opc.part import Part
    from docx.opc.packuri import PackURI

    package = doc.part.package
    # 检查是否已有 footnotes part
    footnotes_part = None
    for rel in doc.part.rels.values():
        if rel.reltype == RT.FOOTNOTES:
            footnotes_part = rel.target_part
            break

    # 构造 footnotes.xml 内容
    fn_xml_parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                    f'<w:footnotes xmlns:w="{W_NS}">',
                    '  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>',
                    '  <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>']

    # ID 从 100 开始（与 process_docx 中 fn_id 一致）
    for offset, num in enumerate(inserted):
        fn_id = 100 + offset
        ref_text = refs_dict.get(num, "")
        # 转义 XML 特殊字符
        safe_text = ref_text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        # Word 脚注会自动渲染编号（上标小数字），所以正文部分只写参考文献文本内容，不需要再加 [N] 前缀
        fn_xml_parts.append(
            f'  <w:footnote w:id="{fn_id}">'
            f'<w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>'
            f'<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/>'
            f'<w:vertAlign w:val="superscript"/></w:rPr>'
            f'<w:footnoteRef/></w:r>'
            f'<w:r><w:t xml:space="preserve"> {safe_text}</w:t></w:r>'
            f'</w:p></w:footnote>'
        )
    fn_xml_parts.append('</w:footnotes>')
    fn_xml = "\n".join(fn_xml_parts).encode("utf-8")

    if footnotes_part is None:
        # 创建新的 footnotes part
        partname = PackURI("/word/footnotes.xml")
        content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
        footnotes_part = Part(partname, content_type, fn_xml, package)
        # 添加关系
        doc.part.relate_to(footnotes_part, RT.FOOTNOTES)
    else:
        # 更新已有 part
        footnotes_part._blob = fn_xml


# ============ 作者-年份引用模式所需的辅助函数 ============

def _extract_year_from_ref(ref: str) -> str | None:
    """从一条参考文献文本中提取年份（4 位数字 + 可选 abcd 后缀）。"""
    m = re.search(r"\d{4}[a-z]?", ref)
    return m.group(0) if m else None


def _extract_first_author_from_ref(ref: str) -> str | None:
    """
    从一条参考文献文本中提取首位作者「姓氏」：
    - 去掉前导 [N] / N. / N） / 编号点
    - 中文：取开头连续 2-5 个汉字（去掉末尾的「等」）
    - 英文：从最近年份左侧的段里取最后一个完整 token 作为姓
    返回规范化字符串（小写，去前后空格）。
    """
    if not ref:
        return None
    # 步骤 1: 去掉前导编号
    text = re.sub(r"^\s*\[?\s*\d+\s*\]?\s*[.,、)\s]*", "", ref).strip()
    if not text:
        return None
    # 步骤 2: 找到首个 4 位年份 出现位置，取年份前为作者段
    year_match = re.search(r"\d{4}[a-z]?", text)
    before_year = text[: year_match.start()] if year_match else text
    author_part = re.split(
        r"[,;；]|et\s+al\.?|and\s+",
        before_year,
        maxsplit=1,
        flags=re.I,
    )[0].strip().rstrip(".,;:").strip()
    if not author_part:
        return None
    # 步骤 3: 中文：取开头连续 2-5 个汉字（去掉末尾「等」字）
    cn = re.match(r"[\u4e00-\u9fff]{2,5}(等)?", author_part)
    if cn:
        return cn.group(0).rstrip("等").lower()
    # 步骤 4: 英文：取最后一个合适长度 token 作为姓
    parts = author_part.split()
    if not parts:
        return None
    last = parts[-1].rstrip(".,;:")
    if len(last) <= 1 and len(parts) >= 2:
        last = parts[-2].rstrip(".,;:")
    return last.lower() if last else None


def _build_author_year_index(refs: list[str]) -> dict[tuple[str, str], str]:
    """
    由参考文献列表构建索引：{(author, year): ref_text}
    - 同一作者同年多篇用 2024a/2024b 区分
    - 同 key 仅保留首次出现的条目
    """
    index: dict[tuple[str, str], str] = {}
    for ref in refs:
        author = _extract_first_author_from_ref(ref)
        year = _extract_year_from_ref(ref)
        if not author or not year:
            continue
        key = (author, year)
        if key not in index:
            index[key] = ref.strip()
    return index


def _scan_author_year_in_md(
    md: str, index: dict[tuple[str, str], str]
) -> list[dict]:
    """
    在 md 中查找所有作者-年份引用，按 (author, year) 反向搜索，
    同时把多文献并列按 年份/分号 拆分为多段。

    返回: [{"start": int, "end": int, "ref_text": str}, ...]
    按 start 升序排列。
    """
    if not md or not index:
        return []
    out: list[dict] = []
    # 索引条目按年份长度降序（"2024b" 优先于 "2024"）以避免误覆盖
    sorted_index = sorted(index.items(), key=lambda kv: len(kv[0][1]), reverse=True)
    for (author, year), ref_text in sorted_index:
        # 转义作者名，author 可能是 unicode 字符
        author_pat = re.escape(author)
        # 年份允许 abcd 后缀
        year_pat = re.escape(year)
        # 分两段匹配以避免 re.IGNORECASE + 中文字符 兼容性 bug：
        # 1) 带「等」或「et al.」
        pat_with_etal = re.compile(
            rf"({author_pat}\s*(等|et\s*al\.?)\s*[,\uff0c]?\s*{year_pat})",
            flags=re.IGNORECASE,
        )
        # 2) 不带 「等」/「et al.」（仅作者 + 可选分隔符 + 年份）
        pat_simple = re.compile(
            rf"({author_pat}\s*[,\uff0c]?\s*{year_pat})",
            flags=re.IGNORECASE,
        )
        for m in pat_with_etal.finditer(md):
            out.append({
                "start": m.start(),
                "end": m.end(),
                "author": author,
                "year": year,
                "ref_text": ref_text,
            })
        for m in pat_simple.finditer(md):
            # 跳过已被 pat_with_etal 覆盖的位置
            if any(o["start"] == m.start() and o["end"] == m.end() for o in out):
                continue
            out.append({
                "start": m.start(),
                "end": m.end(),
                "author": author,
                "year": year,
                "ref_text": ref_text,
            })
    # 按 start 升序，end 降序排序，剔除重叠
    out.sort(key=lambda x: (x["start"], -x["end"]))
    pruned: list[dict] = []
    last_end = -1
    for item in out:
        if item["start"] >= last_end:
            pruned.append(item)
            last_end = item["end"]
    return pruned


def _detect_citation_style(md: str) -> str:
    """
    探测引用风格：
    - 'numeric'    正文中有 [N] / [N-M] / [N,M,K] 等数字引用形式
    - 'author_year' 正文中有形如 (作者, 年份) 且文献列表含作者-年份
    - 'none'       都不匹配
    """
    if not md:
        return "none"
    # 数字风格：单 [N]、范围 [N-M]、连续 [N,M,K]、[N－M]（全角短横）
    numeric_pattern = re.compile(
        r"\[\s*"
        r"\d+"                                # 第一个数字
        r"(?:\s*[\-－\u2013\u2014]\s*\d+|\s*[,\uff0c]\s*\d+)*"   # 后续 -M 或 ,M
        r"\s*\]"
    )
    if numeric_pattern.search(md):
        return "numeric"
    if re.search(
        r"[\uff08\(]?\s*[\u4e00-\u9fff\w][\u4e00-\u9fff\w\.,\s]{0,30}?"
        r"(?:等|et\s*al\.?)?\s*[,\uff0c]\s*\d{4}[a-z]?",
        md,
        flags=re.I,
    ):
        return "author_year"
    return "none"


def _parse_inline_bracket_refs(text: str) -> list[tuple[int, int]]:
    """
    从一段文本里解析所有方括号引用，返回 [(start, end)] 区间列表。
    支持：[N] / [N-M] / [N, M, K] / [N－M] / [N—M] / [N~M]
    """
    refs: list[tuple[int, int]] = []
    # 用 lookahead / step 扫描整个匹配范围
    bracket_re = re.compile(
        r"\[\s*"
        r"(\d+)"                              # 起始编号
        r"((?:\s*[\-－\u2013\u2014~]\s*\d+)|(?:\s*[,，\uff0c]\s*\d+))*"  # 可选序列
        r"\s*\]"
    )
    for m in bracket_re.finditer(text):
        content = m.group(1) + (m.group(2) or "")
        # 解析所有数字
        num_strs = re.findall(r"\d+", content)
        if not num_strs:
            continue
        try:
            nums = [int(n) for n in num_strs]
        except ValueError:
            continue
        # 区间 [5-7] → [5, 6, 7]；列表 [5,6,7] → [5, 6, 7]
        # 检查是否含 '-' 或 '－' 等区间分隔符
        if re.search(r"[\-－\u2013\u2014~]", content):
            start = nums[0]
            end = nums[-1]
            if end >= start:
                nums = list(range(start, end + 1))
            else:
                nums = [start]
        # 简化列表：保留所有数字，去重
        seen = set()
        for n in nums:
            if n not in seen:
                refs.append((m.start(), m.end()))
                seen.add(n)
                break  # 每个 [N-M] 我们只生成一个区间（在 docx 里后续会一次性替换）
        # 但为了完整性，单独数字也要生成 [start, end]
        if not refs or refs[-1][0] != m.start():
            refs.append((m.start(), m.end()))
    return refs


def _split_multi_refs(bracket_content: str) -> list[str]:
    """
    把多文献并列拆分为多个段：
    - 优先按 `;` / `；` 分
    - 然后按 `, author, year` 模式二次拆分（按年份作为分隔点）
    """
    if not bracket_content:
        return []
    # 按 ; 或 ； 先拆
    segs = re.split(r"[;；]", bracket_content)
    out: list[str] = []
    for seg in segs:
        seg = seg.strip()
        if not seg:
            continue
        # 按年份作为分隔点再次拆：多个年份 = 多段
        year_positions = [m.start() for m in re.finditer(r"\d{4}[a-z]?", seg)]
        if len(year_positions) >= 2:
            # 从年份左侧扫描，每年一段
            for i, pos in enumerate(year_positions):
                # 这段从上一段结束或段首到当前年份末尾
                if i == 0:
                    seg_part = seg[: pos + 5]  # 年份 4 位 + 1 后缀
                else:
                    seg_part = seg[year_positions[i - 1] + 5 : pos + 5]
                out.append(seg_part.strip())
        else:
            out.append(seg)
    return out


def _process_docx_with_author_year_footnotes(
    docx_bytes: bytes,
    md_content: str,
    references: list[str],
) -> bytes:
    """
    将 docx 中 (作者等, 年份) 形式的引用替换为 Word 真实脚注。
    复用 footnotes.xml 生成机制（上标、垂直对齐）。
    多文献并列按 ; 或 年份位置分割。
    对未被引用的文献条目，在脚注末尾自动生成「提示型」脚注。
    """
    try:
        doc = Document(io.BytesIO(docx_bytes))
    except Exception as e:
        print(f"[process_docx_author_year_footnotes] 打开 docx 失败: {e}")
        return docx_bytes

    index = _build_author_year_index(references)
    if not index:
        print("[process_docx_author_year_footnotes] 无法从参考文献构建索引，跳过")
        return docx_bytes

    # 扫描 markdown 中所有作者-年份引用，按 (author, year) 索引匹配
    occurrences = _scan_author_year_in_md(md_content, index)
    if not occurrences:
        print("[process_docx_author_year_footnotes] markdown 中未匹配到作者-年份引用")
        return docx_bytes

    # 实际生效的 ref 文本集合（用于判断孤儿文献）
    used_refs: set[str] = set()
    for item in occurrences:
        used_refs.add(item["ref_text"])

    # 实际处理的序号（用于按出现顺序编号）：1, 2, 3 ...
    # 同一个 ref_text 出现多次，每次都插一个独立脚注（用户要求重复插入跳转）
    footnote_to_text: list[str] = []
    for item in occurrences:
        footnote_to_text.append(item["ref_text"])

    # 加孤儿文献（参考文献里没在正文出现的）作为「提示型脚注」附录
    orphan_refs: list[str] = []
    for ref in references:
        ref_norm = ref.strip()
        if ref_norm and ref_norm not in used_refs:
            orphan_refs.append(f"[未在正文中引用] {ref_norm}")

    print(
        f"[process_docx_author_year_footnotes] 正文中匹配到 {len(footnote_to_text)} 处脚注"
        f"（含 {len(set(footnote_to_text))} 条独立文献），孤儿文献 {len(orphan_refs)} 条"
    )

    # 1. 用 markdown 中的位置信息构建需要在 docx 中查找的模式
    #    docx 是从 MinerU 解析生成的正文字段未必和 markdown 完全一致
    #    简化策略：在每个段落中分别用每个 (author, year) 模式查找
    fn_id = 100
    inserted_idx: list[int] = []  # 已插入的 footnote_to_text 索引

    body = doc.element.body
    for p in body.iter(qn("w:p")):
        # 收集段落所有 w:r/w:t 节点
        runs_with_text: list[tuple] = []
        for r in p.findall(qn("w:r")):
            for t in r.findall(qn("w:t")):
                runs_with_text.append((r, t, t.text or ""))
        if not runs_with_text:
            continue

        for r, t, txt in runs_with_text:
            new_text = txt
            # 收集该 w:t 中所有 (author, year) 模式的位置（一次扫描所有 key）
            insertions: list[tuple[int, int, str]] = []
            for key, ref_text in index.items():
                author, year = key
                author_pat = re.escape(author)
                year_pat = re.escape(year)
                # 分两段以避免 re.I + (?:...) 组合在某些中文场景的 bug
                pat_with_etal = re.compile(
                    rf"({author_pat}\s*(等|et\s*al\.?)\s*[,\uff0c]?\s*{year_pat})",
                    flags=re.IGNORECASE,
                )
                pat_simple = re.compile(
                    rf"({author_pat}\s*[,\uff0c]?\s*{year_pat})",
                    flags=re.IGNORECASE,
                )
                for sm in pat_with_etal.finditer(new_text):
                    insertions.append((sm.start(), sm.end(), ref_text))
                for sm in pat_simple.finditer(new_text):
                    if not any(s[0] == sm.start() and s[1] == sm.end() for s in insertions):
                        insertions.append((sm.start(), sm.end(), ref_text))
            if not insertions:
                continue

            # 按位置排序，剔除重叠（保留更长的匹配）
            insertions.sort(key=lambda x: (x[0], -x[1]))
            pruned: list[tuple[int, int, str]] = []
            last_end = -1
            for ins in insertions:
                if ins[0] >= last_end:
                    pruned.append(ins)
                    last_end = ins[1]
            if not pruned:
                continue

            # 重建段落
            parent_run = t.getparent()
            parent_para = parent_run.getparent()
            run_index = list(parent_para).index(parent_run)

            rebuilt_runs: list = []
            cursor = 0
            for ins_start, ins_end, ins_ref_text in pruned:
                if cursor < ins_start:
                    seg_run = OxmlElement("w:r")
                    seg_t = OxmlElement("w:t")
                    seg_t.text = new_text[cursor:ins_start]
                    seg_t.set(qn("xml:space"), "preserve")
                    seg_run.append(seg_t)
                    rebuilt_runs.append(seg_run)
                # 脚注引用 run
                fn_run = OxmlElement("w:r")
                rpr = OxmlElement("w:rPr")
                rstyle = OxmlElement("w:rStyle")
                rstyle.set(qn("w:val"), "FootnoteReference")
                rpr.append(rstyle)
                vert_align = OxmlElement("w:vertAlign")
                vert_align.set(qn("w:val"), "superscript")
                rpr.append(vert_align)
                fn_run.append(rpr)
                fn_ref = OxmlElement("w:footnoteReference")
                fn_ref.set(qn("w:id"), str(fn_id))
                fn_run.append(fn_ref)
                rebuilt_runs.append(fn_run)
                fn_id += 1
                footnote_to_text.append(ins_ref_text)
                inserted_idx.append(len(footnote_to_text) - 1)
                cursor = ins_end
            # 剩余尾部文本
            if cursor < len(new_text):
                tail = OxmlElement("w:r")
                tail_t = OxmlElement("w:t")
                tail_t.text = new_text[cursor:]
                tail_t.set(qn("xml:space"), "preserve")
                tail.append(tail_t)
                rebuilt_runs.append(tail)

            # 清空原 w:t 内容
            t.text = ""
            insertion_offset = 0
            for nr in rebuilt_runs:
                parent_para.insert(run_index + 1 + insertion_offset, nr)
                insertion_offset += 1

    # 2. 构造 footnotes.xml（含孤儿文献提示）
    # footnote 顺序按 footnote_to_text 列表
    all_footnote_texts = footnote_to_text[:]
    for orphan in orphan_refs:
        all_footnote_texts.append(orphan)

    if not all_footnote_texts:
        return docx_bytes

    # 3. 注册 footnotes part
    _ensure_footnotes_part_with_texts(doc, all_footnote_texts, start_id=100)

    # 4. 保存
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()


def _ensure_footnotes_part_with_texts(doc, footnote_texts: list[str], start_id: int = 100):
    """构造 footnotes.xml，每条 footnote 仅包含文本 + FootnoteReference 上标样式。

    与 _ensure_footnotes_part 区别：直接接收文本列表，不再依赖 (num, ref_text) 字典。
    """
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    from docx.opc.part import Part
    from docx.opc.packuri import PackURI

    package = doc.part.package
    footnotes_part = None
    for rel in doc.part.rels.values():
        if rel.reltype == RT.FOOTNOTES:
            footnotes_part = rel.target_part
            break

    fn_xml_parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<w:footnotes xmlns:w="{W_NS}">',
        '  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>',
        '  <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>',
    ]
    for offset, text in enumerate(footnote_texts):
        fn_id = start_id + offset
        safe_text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        fn_xml_parts.append(
            f'  <w:footnote w:id="{fn_id}">'
            f'<w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>'
            f'<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/>'
            f'<w:vertAlign w:val="superscript"/></w:rPr>'
            f'<w:footnoteRef/></w:r>'
            f'<w:r><w:t xml:space="preserve"> {safe_text}</w:t></w:r>'
            f'</w:p></w:footnote>'
        )
    fn_xml_parts.append('</w:footnotes>')
    fn_xml = "\n".join(fn_xml_parts).encode("utf-8")

    if footnotes_part is None:
        partname = PackURI("/word/footnotes.xml")
        content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
        footnotes_part = Part(partname, content_type, fn_xml, package)
        doc.part.relate_to(footnotes_part, RT.FOOTNOTES)
    else:
        footnotes_part._blob = fn_xml


# ============ 6. 按需下载某格式文件 ============
@app.get("/api/task/{internal_id}/format/{fmt}")
async def download_format(internal_id: str, fmt: str):
    """从已完成的解析结果中提取特定格式的文件"""
    if fmt not in SUPPORTED_EXTRA_FORMATS:
        raise HTTPException(400, f"不支持的格式: {fmt}")

    info = _find_file_info(internal_id)
    if not info:
        raise HTTPException(404, "任务不存在")
    if info["state"] != "done" or not info["full_zip_url"]:
        raise HTTPException(400, f"任务未完成，当前状态: {info['state']}")

    async with httpx.AsyncClient(timeout=120.0) as client:
        resp = await client.get(info["full_zip_url"])
        resp.raise_for_status()
        zip_bytes = resp.content

    # 在 ZIP 中查找匹配文件 + 提取 markdown（用于脚注）
    target_name = None
    target_content = None
    md_content = ""
    md_name = ""
    try:
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            for name in zf.namelist():
                if name.endswith(f".{fmt}"):
                    target_name = name.split("/")[-1]
                    target_content = zf.read(name)
                elif fmt == "docx" and not md_content:
                    # MinerU 输出的 markdown 文件名不一定是 full.md；取第一个 .md
                    if name.endswith("full.md") or (name.endswith(".md") and "/auto/" not in name and "/layout/" not in name):
                        md_name = name
                        md_content = zf.read(name).decode("utf-8", errors="replace")
        if fmt == "docx" and not md_content:
            # 兜底：找任意 .md
            with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
                for name in zf.namelist():
                    if name.endswith(".md"):
                        print(f"[download_format] 警告: 未找到 full.md，使用 {name} 作为 markdown 来源")
                        md_content = zf.read(name).decode("utf-8", errors="replace")
                        md_name = name
                        break
    except Exception as e:
        raise HTTPException(500, f"ZIP 解析失败: {e}")
    print(f"[download_format] ZIP 解析完成, target_name={target_name}, md_name={md_name or '(none)'}")

    if target_content is None:
        raise HTTPException(404, f"该任务结果中未包含 .{fmt} 格式（解析时未启用 extra_formats 或 MinerU 未生成）")

    # docx 特殊处理：把 [N] 引用替换为真实 Word 脚注 + 使用文章标题作为下载文件名
    # 公式按用户要求**保持原样**，不转 Word 自带 OMML；
    # MinerU 输出 LaTeX 字符串（通常已用 $$...$$ 或 \[...\] 包裹），
    # 用户手动复制到 MathType。
    download_name = target_name  # 默认 = full.docx
    if fmt == "docx":
        # 策略 1：提取文章标题
        title = _extract_title_from_markdown(md_content)
        if title:
            safe_title = re.sub(r'[\\/:*?"<>|\r\n\t]', "_", title).strip()
            if safe_title:
                download_name = f"{safe_title}.docx"
                print(f"[download_format] 使用文章标题作为文件名: {download_name}")

        # 策略 2：标题提取失败 → 用原 PDF 文件名（去扩展名 + 清理）
        if download_name == target_name and info.get("filename"):
            original_name = re.sub(r'\.[^.]+$', '', info["filename"])  # 去扩展名
            # 清理：把分隔符替换成空格（如 "_夏后学" → " 夏后学"）
            original_name = re.sub(r'[_\-]+', ' ', original_name).strip()
            # 去掉首尾的英文括号内容（如 "(2024)"）
            original_name = re.sub(r'\(\d{4}\)', '', original_name).strip()
            if original_name:
                download_name = f"{original_name}.docx"
                print(f"[download_format] 标题提取失败，用原文件名: {download_name}")

        # 策略 3：兜底 → 用生成的 internal_id 前 8 位
        if download_name == target_name:
            import uuid
            short_id = str(uuid.uuid4())[:8]
            download_name = f"document_{short_id}.docx"
            print(f"[download_format] 完全兜底: {download_name}")

        references = _extract_refs_from_markdown(md_content)
        print(f"[download_format] md_content 长度: {len(md_content)} | 解析到参考文献: {len(references)} 条")
        if md_content:
            # 取 md_content 前 300 字符以便诊断
            md_preview = md_content[:300].replace("\n", " ")
            print(f"[download_format] md_preview: {md_preview!r}")
        if references:
            print(f"[download_format] 找到 {len(references)} 条参考文献，开始处理 docx 脚注")
            print(f"[download_format] 前 3 条参考文献: {references[:3]}")
            try:
                style = _detect_citation_style(md_content)
                print(f"[download_format] 引用风格: {style}")
                if style == "numeric":
                    target_content = _process_docx_with_footnotes(target_content, references)
                elif style == "author_year":
                    target_content = _process_docx_with_author_year_footnotes(
                        target_content, md_content, references
                    )
                else:
                    # 兜底：如果有 [N] 形式的引用但 md 中看不到 references 头部，仍尝试数字脚注
                    print("[download_format] 风格未识别，尝试以数字脚注兜底")
                    target_content = _process_docx_with_footnotes(target_content, references)
            except Exception as e:
                print(f"[download_format] docx 脚注处理失败，返回原始 docx: {e}")
        elif md_content:
            # 没找到 references 但 md_content 存在 → 看看是否有 [N] 脚注
            if re.search(r"\[\d+\]", md_content):
                print("[download_format] md 有 [N] 但未提取到 references，尝试空脚注处理")
                target_content = _process_docx_with_footnotes(target_content, [])
            else:
                print("[download_format] 未找到参考文献，跳过脚注处理")
        else:
            print("[download_format] 未找到 md_content 也未找到参考文献，跳过脚注处理")

    media_types = {
        "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    }
    ascii_fallback = re.sub(r"[^A-Za-z0-9_.-]+", "_", download_name).strip("_")
    if not ascii_fallback or ascii_fallback in (".docx", "docx"):
        ascii_fallback = "document.docx"
    return StreamingResponse(
        io.BytesIO(target_content),
        media_type=media_types[fmt],
        headers={
            "Content-Disposition": (
                f'attachment; filename="{ascii_fallback}"; '
                f"filename*=UTF-8''{urllib.parse.quote(download_name)}"
            )
        },
    )


@app.get("/")
async def root():
    return {"service": "MinerU MVP", "status": "running"}


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)