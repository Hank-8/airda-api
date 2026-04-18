"""
AIRDA 教案產生器 API
FastAPI 微服務 — 接收主題 → 套用範本 → 產出 PPTX → 回傳檔案下載
"""

import os
import tempfile
import traceback
from urllib.parse import quote

from dotenv import load_dotenv
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel
from starlette.background import BackgroundTask

from lesson_content import generate_content
from pptx_generator import generate_pptx

load_dotenv()

app = FastAPI(title="AIRDA 教案產生器 API")

ALLOWED_ORIGINS = os.getenv("ALLOWED_ORIGINS", "*").split(",")

app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_methods=["GET", "POST"],
    allow_headers=["*"],
)


class LessonRequest(BaseModel):
    topic: str


@app.post("/api/generate-lesson")
async def generate_lesson(req: LessonRequest):
    topic = req.topic.strip()
    if not topic:
        return JSONResponse(status_code=400, content={"error": "主題不能為空"})

    try:
        content = generate_content(topic)

        tmp = tempfile.NamedTemporaryFile(suffix='.pptx', delete=False)
        tmp.close()
        generate_pptx(content, tmp.name)

        filename = f"{topic}_教案.pptx"
        encoded = quote(filename)

        return FileResponse(
            path=tmp.name,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={
                "Content-Disposition": f"attachment; filename*=UTF-8''{encoded}",
            },
            background=BackgroundTask(os.unlink, tmp.name),
        )
    except Exception as e:
        traceback.print_exc()
        return JSONResponse(status_code=500, content={"error": str(e)})


@app.get("/health")
async def health():
    return {"status": "ok"}
