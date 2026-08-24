from pathlib import Path
from tempfile import NamedTemporaryFile
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from .generator import generate_proработка

BASE = Path(__file__).resolve().parents[2]
TEMPLATE = BASE / "templates" / "proработка" / "шаблон.xlsx"
FRONTEND = BASE / "frontend"

app = FastAPI(title="Помощник Опера", version="0.1.0")
app.mount("/static", StaticFiles(directory=FRONTEND), name="static")

@app.get("/", response_class=HTMLResponse)
def index():
    return (FRONTEND / "index.html").read_text(encoding="utf-8")

@app.get("/api/documents")
def documents():
    return {"documents": [{"id": "proработка", "name": "Проработка", "formats": ["xlsx"]}]}

@app.post("/api/generate/proработка")
async def generate(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(400, "Пока поддерживаются только .xlsx")
    source_tmp = NamedTemporaryFile(suffix=".xlsx", delete=False)
    output_tmp = NamedTemporaryFile(suffix=".xlsx", delete=False)
    source_path, output_path = Path(source_tmp.name), Path(output_tmp.name)
    source_tmp.close(); output_tmp.close()
    try:
        source_path.write_bytes(await file.read())
        generate_proработка(str(source_path), str(TEMPLATE), str(output_path))
        return FileResponse(output_path, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename="Проработка_готово.xlsx")
    except ValueError as e:
        raise HTTPException(400, str(e)) from e
    except Exception as e:
        raise HTTPException(500, f"Ошибка генерации: {e}") from e
