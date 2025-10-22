# main.py
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
from starlette.background import BackgroundTask
from fastapi.middleware.cors import CORSMiddleware
from pathlib import Path
import tempfile, shutil

from build_report import build_report_template_api  # tu función parametrizada

app = FastAPI(title="Report Builder API")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], allow_methods=["*"], allow_headers=["*"], allow_credentials=True,
)

BASE_DIR = Path("/home/erick/projectos/stata").resolve()
EP_REGION_XLSX = BASE_DIR / "ep_region_macrorregion.xlsx"
PLANTILLA_DOCX = BASE_DIR / "plantilla_word_informe.docx"


@app.get("/")
def read_root():
    return FileResponse("index.html")

@app.post("/generate-report")
async def generate_report(
    eventos_xlsx: UploadFile = File(...),
    vehiculos_xlsx: UploadFile = File(...),
    direcciones_xlsx: UploadFile = File(...),
    mes_nombre: str = "Agosto",
    mes_num: int = 8,
    anio: int = 2025,
):
    for uf in (eventos_xlsx, vehiculos_xlsx, direcciones_xlsx):
        if not uf.filename.lower().endswith(".xlsx"):
            raise HTTPException(status_code=400, detail=f"Archivo inválido: {uf.filename}")

    if not EP_REGION_XLSX.exists():
        raise HTTPException(status_code=500, detail="ep_region_macrorregion.xlsx no encontrado.")
    if not PLANTILLA_DOCX.exists():
        raise HTTPException(status_code=500, detail="plantilla_word_informe.docx no encontrada.")

    # Creamos carpeta temporal manual (no usar 'with TemporaryDirectory()' aquí)
    tmpdir = Path(tempfile.mkdtemp(prefix="reportapi_"))

    try:
        ev_path = tmpdir / f"eventos_{eventos_xlsx.filename}"
        ve_path = tmpdir / f"vehiculos_{vehiculos_xlsx.filename}"
        di_path = tmpdir / f"direcciones_{direcciones_xlsx.filename}"

        with ev_path.open("wb") as f:
            shutil.copyfileobj(eventos_xlsx.file, f)
        with ve_path.open("wb") as f:
            shutil.copyfileobj(vehiculos_xlsx.file, f)
        with di_path.open("wb") as f:
            shutil.copyfileobj(direcciones_xlsx.file, f)

        out_path = tmpdir / f"Informe_Final_{mes_nombre}_{anio}.docx"

        # Genera el informe (tu función parametrizada)
        result_path = build_report_template_api(
            eventos_xlsx=ev_path,
            vehiculos_xlsx=ve_path,
            direcciones_xlsx=di_path,
            ep_region_xlsx=EP_REGION_XLSX,
            plantilla_docx=PLANTILLA_DOCX,
            salida_docx=out_path,
            mes_nombre=mes_nombre,
            mes_num=mes_num,
            anio=anio,
        )

        # Limpieza del directorio temporal DESPUÉS de enviar la respuesta
        cleanup = BackgroundTask(shutil.rmtree, tmpdir, ignore_errors=True)

        return FileResponse(
            path=str(result_path),
            filename=result_path.name,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            background=cleanup,
        )

    except Exception as e:
        shutil.rmtree(tmpdir, ignore_errors=True)
        raise HTTPException(status_code=500, detail=f"Error generando el informe: {e}")

