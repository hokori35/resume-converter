import os
import json
import tempfile
import shutil
from pathlib import Path

import requests
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel

# 导入你原有的两个核心函数
from batch_extract_json import extract_one
from fill_template import fill_template

app = FastAPI(title="简历格式转换服务")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# 模板文件路径（和 main.py 放在同一目录）
TEMPLATE_PATH = Path(__file__).parent / "调整空白版.docx"


# ===== 新增：接收 HiAgent 文件 URL =====
class FileInfo(BaseModel):
    name: str
    url: str


class ConvertRequest(BaseModel):
    file: FileInfo


@app.get("/")
def root():
    return {"message": "简历转换服务运行中"}


@app.post("/convert-resume")
async def convert_resume(req: ConvertRequest):
    """
    接收 HiAgent 传来的文件 URL，
    下载 .docx 简历并生成标准格式简历
    """

    # 检查文件格式
    if not req.file.name.endswith(".docx"):
        raise HTTPException(
            status_code=400,
            detail="只支持 .docx 格式的简历文件"
        )

    # 检查模板是否存在
    if not TEMPLATE_PATH.exists():
        raise HTTPException(
            status_code=500,
            detail="服务器缺少 Word 模板文件，请联系管理员"
        )

    # 用临时目录处理所有中间文件
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)

        # ===== 下载 HiAgent 文件 =====
        try:
            response = requests.get(req.file.url)

            if response.status_code != 200:
                raise HTTPException(
                    status_code=400,
                    detail="文件下载失败"
                )

        except Exception as e:
            raise HTTPException(
                status_code=500,
                detail=f"下载文件失败：{str(e)}"
            )

        # 保存下载文件
        input_path = tmpdir / req.file.name

        with open(input_path, "wb") as f:
            f.write(response.content)

        # ===== 提取简历 JSON =====
        try:
            resume_dict = extract_one(input_path)

        except Exception as e:
            raise HTTPException(
                status_code=500,
                detail=f"简历解析失败：{str(e)}"
            )

        # ===== 保存 JSON =====
        json_path = tmpdir / "resume.json"

        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(
                resume_dict,
                f,
                ensure_ascii=False,
                indent=2
            )

        # ===== 生成标准简历 =====
        output_path = tmpdir / "output.docx"

        try:
            fill_template(
                template_path=str(TEMPLATE_PATH),
                data_json=str(json_path),
                output_path=str(output_path),
                compact=False,
            )

        except Exception as e:
            raise HTTPException(
                status_code=500,
                detail=f"简历生成失败：{str(e)}"
            )

        # ===== 保存最终结果 =====
        result_dir = Path(__file__).parent / "results"
        result_dir.mkdir(exist_ok=True)

        name = resume_dict.get(
            "basic_info",
            {}
        ).get(
            "name",
            "简历"
        )

        result_path = result_dir / f"{name}_学术业绩简表.docx"

        shutil.copy(output_path, result_path)

    # ===== 返回下载链接 =====
    download_url = (
        "https://resume-converter-production-983c.up.railway.app"
        f"/download/{result_path.name}"
    )

    return {
        "status": "success",
        "filename": result_path.name,
        "download_url": download_url
    }


# ===== 文件下载接口 =====
from fastapi.responses import FileResponse as FR


@app.get("/download/{filename}")
def download_file(filename: str):

    result_dir = Path(__file__).parent / "results"

    file_path = result_dir / filename

    if not file_path.exists():
        raise HTTPException(
            status_code=404,
            detail="文件不存在"
        )

    return FR(
        path=str(file_path),
        media_type=(
            "application/"
            "vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        ),
        filename=filename,
    )


# ===== 调试接口：提取 JSON =====
@app.post("/extract-json")
async def extract_json(req: ConvertRequest):

    if not req.file.name.endswith(".docx"):
        raise HTTPException(
            status_code=400,
            detail="只支持 .docx 格式"
        )

    with tempfile.TemporaryDirectory() as tmpdir:

        tmpdir = Path(tmpdir)

        response = requests.get(req.file.url)

        if response.status_code != 200:
            raise HTTPException(
                status_code=400,
                detail="文件下载失败"
            )

        input_path = tmpdir / req.file.name

        with open(input_path, "wb") as f:
            f.write(response.content)

        try:
            resume_dict = extract_one(input_path)

        except Exception as e:
            raise HTTPException(
                status_code=500,
                detail=f"解析失败：{str(e)}"
            )

    return resume_dict
