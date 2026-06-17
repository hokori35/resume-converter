import json
import os
import tempfile
import shutil
from pathlib import Path
from typing import Tuple, Union
from urllib.parse import quote

import httpx
from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse

from batch_extract_json import extract_one
from fill_template import default_output_docx_filename, fill_template

app = FastAPI(title="简历格式转换服务")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

TEMPLATE_PATH = Path(__file__).parent / "调整空白版.docx"
RESULT_DIR = Path(__file__).parent / "results"


def _ensure_docx_filename(filename: str) -> str:
    name = Path(str(filename or "").strip()).name
    if not name.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="只支持 .docx 格式的简历文件")
    return name


async def _download_file_from_url(file_url: str, dest: Path) -> None:
    headers = {
        "User-Agent": "ResumeConvertService/1.0",
        "Accept": "*/*",
    }
    try:
        async with httpx.AsyncClient(follow_redirects=True, timeout=120.0) as client:
            resp = await client.get(file_url, headers=headers)
    except httpx.RequestError as exc:
        raise HTTPException(
            status_code=502,
            detail=f"无法访问文件链接，请重新上传后重试（{exc}）",
        ) from exc

    if resp.status_code == 404:
        raise HTTPException(
            status_code=404,
            detail="文件链接不存在或已过期，请在对话中重新上传 .docx 文件后再试",
        )
    if resp.status_code >= 400:
        raise HTTPException(
            status_code=502,
            detail=f"下载文件失败（HTTP {resp.status_code}），请重新上传",
        )
    if not resp.content:
        raise HTTPException(
            status_code=400,
            detail="检测到简历文件格式异常或内容不完整，请检查文件是否为可编辑的 Word 文档（.docx）并确保内容完整后重新上传。",
        )
    dest.write_bytes(resp.content)


def _convert_docx_file(input_path: Path, workdir: Path) -> Tuple[dict, Path, str]:
    if not TEMPLATE_PATH.exists():
        raise HTTPException(status_code=500, detail="服务器缺少 Word 模板文件，请联系管理员")

    try:
        resume_dict = extract_one(input_path)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"简历解析失败：{exc}") from exc

    json_path = workdir / "resume.json"
    json_path.write_text(
        json.dumps(resume_dict, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    out_name = default_output_docx_filename(resume_dict.get("basic_info", {}))
    output_path = workdir / out_name
    try:
        fill_template(
            template_path=str(TEMPLATE_PATH),
            data_json=str(json_path),
            output_path=str(output_path),
            compact=False,
        )
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"简历生成失败：{exc}") from exc

    RESULT_DIR.mkdir(exist_ok=True)
    result_path = RESULT_DIR / out_name
    shutil.copy(output_path, result_path)
    return resume_dict, result_path, out_name


def _file_response(result_path: Path, out_name: str) -> FileResponse:
    return FileResponse(
        path=str(result_path),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=out_name,
    )


def _public_base_url(request: Request) -> str:
    """公网访问根 URL；部署时建议设置环境变量 PUBLIC_BASE_URL。"""
    env = (os.environ.get("PUBLIC_BASE_URL") or "").strip().rstrip("/")
    if env:
        return env
    return str(request.base_url).rstrip("/")


def _safe_result_path(name: str) -> Path:
    safe = Path(str(name or "")).name
    if not safe.lower().endswith(".docx"):
        raise HTTPException(status_code=404, detail="文件不存在或已过期")
    path = (RESULT_DIR / safe).resolve()
    root = RESULT_DIR.resolve()
    if root not in path.parents and path != root:
        raise HTTPException(status_code=404, detail="文件不存在或已过期")
    if not path.is_file():
        raise HTTPException(status_code=404, detail="文件不存在或已过期")
    return path


def _build_download_url(request: Request, out_name: str) -> str:
    encoded = quote(out_name, safe="")
    return f"{_public_base_url(request)}/download/{encoded}"


def _coze_json_response(request: Request, result_path: Path, out_name: str) -> JSONResponse:
    """与 Coze 插件输出参数对齐：status、filename、download_url。"""
    return JSONResponse(
        {
            "status": "success",
            "filename": out_name,
            "download_url": _build_download_url(request, out_name),
        }
    )


async def _handle_convert_resume(
    *,
    request: Request,
    input_path: Path,
    workdir: Path,
    return_json: bool,
) -> Union[FileResponse, JSONResponse]:
    _, result_path, out_name = _convert_docx_file(input_path, workdir)
    if return_json:
        return _coze_json_response(request, result_path, out_name)
    return _file_response(result_path, out_name)


@app.get("/")
def root():
    return {"message": "简历转换服务运行中"}


@app.get("/health")
def health():
    return {
        "status": "ok",
        "template_exists": TEMPLATE_PATH.is_file(),
        "endpoints": ["/convert-resume", "/convert_resume", "/download/{filename}"],
    }


@app.get("/download/{file_name:path}")
def download_result(file_name: str):
    """供 Coze 插件 download_url 指向的下载接口。"""
    path = _safe_result_path(file_name)
    return FileResponse(
        path=str(path),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=path.name,
    )


@app.post("/convert-resume")
@app.post("/convert_resume")
async def convert_resume(
    request: Request,
    file: UploadFile = File(None),
):
    """
    支持两种调用方式（Coze 智能体 / 本地调试）：
    1. JSON：{"file_url": "...", "filename": "xxx.docx"}
    2. multipart：表单字段 file=上传的 .docx
    JSON 请求默认返回 JSON（status、filename、download_url）；加 ?response=file 可改回文件流。
    """
    content_type = (request.headers.get("content-type") or "").lower()
    return_json = "application/json" in content_type and request.query_params.get(
        "response"
    ) != "file"

    with tempfile.TemporaryDirectory() as tmpdir:
        workdir = Path(tmpdir)

        if "application/json" in content_type:
            try:
                body = await request.json()
            except json.JSONDecodeError as exc:
                raise HTTPException(status_code=400, detail="请求体必须是 JSON") from exc

            file_url = (body.get("file_url") or "").strip()
            filename = _ensure_docx_filename(body.get("filename") or "")
            if not file_url:
                raise HTTPException(status_code=400, detail="file_url 为必填参数")

            input_path = workdir / filename
            await _download_file_from_url(file_url, input_path)
            return await _handle_convert_resume(
                request=request,
                input_path=input_path,
                workdir=workdir,
                return_json=return_json,
            )

        if file is not None and file.filename:
            filename = _ensure_docx_filename(file.filename)
            input_path = workdir / filename
            input_path.write_bytes(await file.read())
            return await _handle_convert_resume(
                request=request,
                input_path=input_path,
                workdir=workdir,
                return_json=False,
            )

        if "multipart/form-data" in content_type or "application/x-www-form-urlencoded" in content_type:
            form = await request.form()
            file_url = (form.get("file_url") or "").strip()
            if file_url:
                filename = _ensure_docx_filename(form.get("filename") or "")
                input_path = workdir / filename
                await _download_file_from_url(file_url, input_path)
                return await _handle_convert_resume(
                    request=request,
                    input_path=input_path,
                    workdir=workdir,
                    return_json=False,
                )
            upload = form.get("file")
            if upload is not None and getattr(upload, "filename", None):
                filename = _ensure_docx_filename(upload.filename)
                input_path = workdir / filename
                input_path.write_bytes(await upload.read())
                return await _handle_convert_resume(
                    request=request,
                    input_path=input_path,
                    workdir=workdir,
                    return_json=False,
                )

    raise HTTPException(
        status_code=400,
        detail="请提供 file（multipart 上传）或 JSON 字段 file_url + filename",
    )


@app.post("/extract-json")
async def extract_json(file: UploadFile = File(...)):
    """只做第一步：简历 .docx → JSON，方便调试。"""
    filename = _ensure_docx_filename(file.filename or "")

    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = Path(tmpdir) / filename
        input_path.write_bytes(await file.read())

        try:
            resume_dict = extract_one(input_path)
        except Exception as exc:
            raise HTTPException(status_code=500, detail=f"解析失败：{exc}") from exc

    return resume_dict
