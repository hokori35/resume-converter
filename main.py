import os
import json
import tempfile
import shutil
from pathlib import Path

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

# 导入你原有的两个核心函数
from batch_extract_json import extract_one
from fill_template import fill_template, default_output_docx_filename

app = FastAPI(title="简历格式转换服务")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# 模板文件路径（和 main.py 放在同一目录）
TEMPLATE_PATH = Path(__file__).parent / "调整空白版.docx"


@app.get("/")
def root():
    return {"message": "简历转换服务运行中"}


@app.post("/convert-resume")
async def convert_resume(file: UploadFile = File(...)):
    """
    接收 .docx 简历文件，返回转换后的标准格式简历（.docx）
    """
    # 检查文件格式
    if not file.filename.endswith(".docx"):
        raise HTTPException(status_code=400, detail="只支持 .docx 格式的简历文件")

    # 检查模板是否存在
    if not TEMPLATE_PATH.exists():
        raise HTTPException(status_code=500, detail="服务器缺少 Word 模板文件，请联系管理员")

    # 用临时目录处理所有中间文件，结束后自动清理
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)

        # 1. 把上传的文件保存到临时目录
        input_path = tmpdir / file.filename
        with open(input_path, "wb") as f:
            f.write(await file.read())

        # 2. 调用 extract_one：简历 .docx → dict
        try:
            resume_dict = extract_one(input_path)
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"简历解析失败：{str(e)}")

        # 3. 把 dict 写成临时 JSON 文件（fill_template 需要文件路径）
        json_path = tmpdir / "resume.json"
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(resume_dict, f, ensure_ascii=False, indent=2)

        # 4. 调用 fill_template：JSON → 新简历 .docx
        out_name = default_output_docx_filename(resume_dict.get("basic_info", {}))
        output_path = tmpdir / out_name
        try:
            fill_template(
                template_path=str(TEMPLATE_PATH),
                data_json=str(json_path),
                output_path=str(output_path),
                compact=False,
            )
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"简历生成失败：{str(e)}")

        # 5. 把生成的文件复制到持久目录再返回（临时目录马上会被删除）
        result_dir = Path(__file__).parent / "results"
        result_dir.mkdir(exist_ok=True)

        result_path = result_dir / out_name
        shutil.copy(output_path, result_path)

    return FileResponse(
        path=str(result_path),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=result_path.name,
    )


@app.post("/extract-json")
async def extract_json(file: UploadFile = File(...)):
    """
    只做第一步：简历 .docx → JSON，方便调试
    """
    if not file.filename.endswith(".docx"):
        raise HTTPException(status_code=400, detail="只支持 .docx 格式")

    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = Path(tmpdir) / file.filename
        with open(input_path, "wb") as f:
            f.write(await file.read())

        try:
            resume_dict = extract_one(input_path)
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"解析失败：{str(e)}")

    return resume_dict