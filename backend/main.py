"""
MinerU MVP 后端
- 作为 MinerU 云 API 的代理层，避免前端泄露 Token
- 暴露文件上传 -> 任务提交 -> 进度轮询 -> 结果下载四个端点
"""
import asyncio
import base64
import json
import os
import re
import time
import uuid
import urllib.parse
import zipfile
import io
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
    markdown: str
    files: list[dict] = []  # 兼容性字段（保留为空）
    available_formats: list[str] = []  # ZIP 内可用的格式（docx）
    # 从 content_list.json 提取的结构化信息
    toc: list[dict] = []         # 大纲 [{level, text, page?}]
    references: list[str] = []   # 参考文献列表


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

    # 提交到 MinerU，默认请求 docx/html/latex 格式
    try:
        batch_id = await upload_and_submit(
            content, file.filename or "document.pdf", model_version,
            is_ocr, enable_formula, enable_table,
            ["docx"],
        )
        print(f"[parse_document] MinerU batch_id: {batch_id}")
    except HTTPException:
        raise
    except Exception as e:
        print(f"[parse_document] MinerU 调用失败: {type(e).__name__}: {e}")
        raise HTTPException(500, f"MinerU 调用失败: {e}")

    # 记录到本地任务表
    internal_id = str(uuid.uuid4())
    TASKS[internal_id] = {
        "task_id": batch_id,
        "state": "pending",
        "extracted_pages": 0,
        "total_pages": 0,
        "full_zip_url": None,
        "err_msg": "",
        "model_version": model_version,
        "filename": file.filename,
        "created_at": time.time(),
    }
    return {"internal_id": internal_id, "batch_id": batch_id}


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
                # content_list.json：MinerU 的结构化输出
                if name.endswith("_content_list.json") and content_list_data is None:
                    try:
                        raw = zf.read(name).decode("utf-8", errors="replace")
                        content_list_data = json.loads(raw)
                    except Exception:
                        content_list_data = None

            # 从 content_list 中提取大纲和参考文献
            if isinstance(content_list_data, list):
                toc, references = _extract_toc_and_refs(content_list_data)
            # 部分 ZIP 用 dict 形式（含 items 字段）
            elif isinstance(content_list_data, dict):
                items = content_list_data.get("items") or []
                toc, references = _extract_toc_and_refs(items)

            # 兜底：如果 content_list 没提取到，从 markdown 解析
            if not toc and md_content:
                toc = _extract_toc_from_markdown(md_content)
            if not references and md_content:
                references = _extract_refs_from_markdown(md_content)
    except Exception as e:
        md_content = md_content or f"[ZIP 解析失败: {e}]"

    return ParseResult(
        filename=info.get("filename", "result"),
        markdown=md_content,
        files=[],
        available_formats=available_formats,
        toc=toc,
        references=references,
    )


def _extract_toc_and_refs(items: list) -> tuple[list[dict], list[str]]:
    """从 MinerU content_list.json 中提取大纲和参考文献（结构化优先）"""
    toc: list[dict] = []
    refs: list[str] = []

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
            refs.append(text)

    return toc, refs


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
    """兜底：从 Markdown 中提取参考文献段落"""
    refs: list[str] = []
    in_ref_section = False
    for line in md.split("\n"):
        stripped = line.strip()
        # 识别「参考文献」「References」「Bibliography」等章节
        if re.match(r"^#{1,3}\s*(参考文献|References?|Bibliography|引用文献|REFERENCES)", stripped, re.I):
            in_ref_section = True
            continue
        # 进入参考文献段后，遇到下一个标题则退出
        if in_ref_section and re.match(r"^#{1,3}\s+", stripped):
            in_ref_section = False
        if in_ref_section and stripped:
            refs.append(stripped)
    return refs[:50]  # 最多 50 条防溢出


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

    # 在 ZIP 中查找匹配文件
    target_name = None
    target_content = None
    try:
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            for name in zf.namelist():
                if name.endswith(f".{fmt}"):
                    target_name = name.split("/")[-1]
                    target_content = zf.read(name)
                    break
    except Exception as e:
        raise HTTPException(500, f"ZIP 解析失败: {e}")

    if target_content is None:
        raise HTTPException(404, f"该任务结果中未包含 .{fmt} 格式（解析时未启用 extra_formats 或 MinerU 未生成）")

    media_types = {
        "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "html": "text/html",
        "latex": "application/x-latex",
    }
    return StreamingResponse(
        io.BytesIO(target_content),
        media_type=media_types[fmt],
        headers={
            "Content-Disposition": (
                f"attachment; filename={target_name}; "
                f"filename*=UTF-8''{urllib.parse.quote(target_name)}"
            )
        },
    )


@app.get("/")
async def root():
    return {"service": "MinerU MVP", "status": "running"}


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)