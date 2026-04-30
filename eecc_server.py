#!/usr/bin/env python3
"""
eecc_server.py — Servidor para generación de EECC
n8n llama a POST /generar con las URLs de los archivos y parámetros del cliente.
"""
import os, subprocess, tempfile, shutil, zipfile, urllib.request
from pathlib import Path
from fastapi import FastAPI, Form, HTTPException
from fastapi.responses import FileResponse
import uvicorn

app = FastAPI(title="EECC Generator")

GEN_SCRIPT = Path(__file__).parent / "gen_eecc_v7.py"


@app.get("/health")
def health():
    return {"status": "ok", "script": str(GEN_SCRIPT), "exists": GEN_SCRIPT.exists()}


@app.post("/generar")
async def generar(
    ss_url:        str = Form(...),
    eecc_url:      str = Form(default=""),
    empresa:       str = Form(...),
    cuit:          str = Form(...),
    nro_ejercicio: int = Form(default=1),
    fecha_cierre:  str = Form(...),   # YYYY-MM-DD
    cof:           float = Form(...),
    cap_nominal:   float = Form(...),
):
    tmp = tempfile.mkdtemp(prefix="eecc_")
    try:
        ss_act_path = os.path.join(tmp, "ss_actual.xlsx")
        out_path    = os.path.join(tmp, "output.xlsx")

        # Descargar SS desde la URL firmada de Notion/S3
        urllib.request.urlretrieve(ss_url, ss_act_path)

        cmd = [
            "python3", str(GEN_SCRIPT),
            "--empresa",       empresa,
            "--cuit",          cuit,
            "--nro-ejercicio", str(nro_ejercicio),
            "--fecha-cierre",  fecha_cierre.strip(),
            "--cof",           str(cof),
            "--cap-nominal",   str(cap_nominal),
            "--ss-actual",     ss_act_path,
            "--output",        out_path,
        ]

        # Descargar EECC anterior PDF si hay URL
        if eecc_url and eecc_url.strip():
            prev_path = os.path.join(tmp, "eecc_anterior.pdf")
            urllib.request.urlretrieve(eecc_url.strip(), prev_path)
            if os.path.getsize(prev_path) > 100:
                cmd += ["--eecc-anterior", prev_path]

        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)

        if result.returncode != 0:
            raise HTTPException(
                status_code=500,
                detail=f"Error en gen_eecc: {result.stderr}\n{result.stdout}"
            )

        if not os.path.exists(out_path):
            raise HTTPException(status_code=500, detail="El script no generó el archivo")

        empresa_slug = empresa.replace(" ", "_").replace(".", "")[:30]
        xlsx_name = f"EECC_{empresa_slug}_{fecha_cierre.strip()[:4]}.xlsx"
        pdf_name  = f"EECC_{empresa_slug}_{fecha_cierre.strip()[:4]}.pdf"
        pdf_path  = os.path.join(tmp, pdf_name)

        _export_pdf(out_path, pdf_path, tmp)

        zip_path = os.path.join(tmp, "eecc.zip")
        with zipfile.ZipFile(zip_path, "w") as zf:
            zf.write(out_path, xlsx_name)
            if os.path.exists(pdf_path):
                zf.write(pdf_path, pdf_name)

        final_zip = Path(tempfile.gettempdir()) / f"eecc_{os.path.basename(tmp)}.zip"
        shutil.copy(zip_path, final_zip)

        return FileResponse(
            path=str(final_zip),
            media_type="application/zip",
            filename="eecc.zip",
            background=_cleanup(tmp, final_zip),
        )

    except HTTPException:
        shutil.rmtree(tmp, ignore_errors=True)
        raise
    except Exception as e:
        shutil.rmtree(tmp, ignore_errors=True)
        raise HTTPException(status_code=500, detail=str(e))


def _export_pdf(xlsx_path: str, pdf_path: str, out_dir: str):
    """Exporta xlsx → pdf usando LibreOffice headless."""
    try:
        result = subprocess.run([
            "libreoffice", "--headless", "--convert-to", "pdf",
            "--outdir", out_dir, xlsx_path
        ], capture_output=True, text=True, timeout=60)

        # LibreOffice nombra el PDF igual que el xlsx
        base = os.path.splitext(os.path.basename(xlsx_path))[0]
        generated = os.path.join(out_dir, base + ".pdf")
        if os.path.exists(generated) and generated != pdf_path:
            os.rename(generated, pdf_path)
    except Exception as e:
        print(f"[PDF] No se pudo exportar: {e}")


class _cleanup:
    def __init__(self, tmp_dir, final_zip):
        self._tmp = tmp_dir
        self._zip = final_zip
    def __call__(self, *_):
        shutil.rmtree(self._tmp, ignore_errors=True)
        try: os.unlink(self._zip)
        except: pass


if __name__ == "__main__":
    print("Iniciando servidor EECC en http://localhost:8000")
    uvicorn.run(app, host="0.0.0.0", port=8000)
