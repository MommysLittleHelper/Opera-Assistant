from pathlib import Path
from tempfile import TemporaryDirectory
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles

from .generator import generate_proработка

BASE_DIR = Path(__file__).resolve().parents[2]
TEMPLATE_DIR = BASE_DIR / "templates"
FRONTEND_DIR = BASE_DIR / "frontend"

app = FastAPI(title="Помощник Опера", version="0.1.0")
app.mount("/static", StaticFiles(directory=FRONTEND_DIR), name="static")


@app.get("/", response_class=HTMLResponse)
def index():
    return (FRONTEND_DIR / "index.html").read_text(encoding="utf-8")


@app.get("/api/documents")
def documents():
    return {
        "documents": [
            {
                "id": "proработка",
                "name": "Проработка",
                "formats": ["xlsx"],
            }
        ]
    }


@app.post("/api/generate/proработка")
async def generate_proработка_endpoint(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Пока поддерживаются только .xlsx")

    with TemporaryDirectory() as tmp:
        source = Path(tmp) / "source.xlsx"
        output = Path(tmp) / "Проработка_готово.xlsx"
        source.write_bytes(await file.read())

        try:
            generate_proработка(
                str(source),
                str(TEMPLATE_DIR / "проработка_шаблон.xlsx"),
                str(output),
            )
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        except Exception as exc:
            raise HTTPException(
                status_code=500,
                detail=f"Не удалось сформировать документ: {exc}",
            ) from exc

        # FileResponse читает файл до завершения ответа, поэтому копируем его
        # в постоянное временное место для текущего процесса.
        import tempfile
        persistent_dir = Path(tempfile.mkdtemp(prefix="opera_"))
        persistent_output = persistent_dir / output.name
        persistent_output.write_bytes(output.read_bytes())

    return FileResponse(
        persistent_output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=output.name,
    )
