#!/usr/bin/env python3
"""
eecc_server.py — Servidor para generación de EECC
n8n llama a POST /generar con URLs de archivos y parámetros del cliente.
"""
import os, subprocess, tempfile, shutil, zipfile, urllib.request
from pathlib import Path
from datetime import date, datetime
from fastapi import FastAPI, Form, HTTPException
from fastapi.responses import FileResponse
import uvicorn

app = FastAPI(title="EECC Generator")

GEN_SCRIPT       = Path(__file__).parent / "gen_eecc_v7.py"
INFORME_TEMPLATE = Path(__file__).parent / "informe_template.docx"

MONTHS_ES = {
    1:'enero', 2:'febrero', 3:'marzo', 4:'abril', 5:'mayo', 6:'junio',
    7:'julio', 8:'agosto', 9:'septiembre', 10:'octubre', 11:'noviembre', 12:'diciembre'
}


@app.get("/health")
def health():
    return {
        "status": "ok",
        "script": str(GEN_SCRIPT),
        "exists": GEN_SCRIPT.exists(),
        "template": INFORME_TEMPLATE.exists()
    }


@app.post("/generar")
async def generar(
    ss_url:        str   = Form(...),
    eecc_url:      str   = Form(default=""),
    empresa:       str   = Form(...),
    cuit:          str   = Form(...),
    domicilio:     str   = Form(default=""),
    matricula_igj: str   = Form(default=""),
    nro_ejercicio: int   = Form(default=1),
    fecha_cierre:  str   = Form(...),   # YYYY-MM-DD
    cof:           float = Form(...),
    cap_nominal:   float = Form(...),
    sipa_monto:    str   = Form(default=""),
):
    tmp = tempfile.mkdtemp(prefix="eecc_")
    try:
        ss_act_path = os.path.join(tmp, "ss_actual.xlsx")
        out_path    = os.path.join(tmp, "output.xlsx")

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

        if eecc_url and eecc_url.strip():
            prev_path = os.path.join(tmp, "eecc_anterior.pdf")
            urllib.request.urlretrieve(eecc_url.strip(), prev_path)
            if os.path.getsize(prev_path) > 100:
                cmd += ["--eecc-anterior", prev_path]

        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        if result.returncode != 0:
            raise HTTPException(status_code=500,
                detail=f"Error en gen_eecc: {result.stderr}\n{result.stdout}")

        if not os.path.exists(out_path):
            raise HTTPException(status_code=500, detail="El script no generó el archivo")

        empresa_slug = empresa.replace(" ", "_").replace(".", "")[:30]
        year = fecha_cierre.strip()[:4]
        xlsx_name    = f"EECC_{empresa_slug}_{year}.xlsx"
        pdf_name     = f"EECC_{empresa_slug}_{year}.pdf"
        excel_pdf    = os.path.join(tmp, "excel.pdf")
        informe_pdf  = os.path.join(tmp, "informe.pdf")
        merged_pdf   = os.path.join(tmp, pdf_name)

        # 1. Excel → PDF (todas las solapas incluida Notas)
        _libreoffice_pdf(out_path, tmp, excel_pdf)

        # 2. Informe Word → rellenar → PDF
        if INFORME_TEMPLATE.exists():
            informe_filled = os.path.join(tmp, "informe_filled.docx")
            _fill_informe(str(INFORME_TEMPLATE), informe_filled,
                          empresa, cuit, domicilio, matricula_igj,
                          fecha_cierre.strip(), sipa_monto)
            _libreoffice_pdf(informe_filled, tmp, informe_pdf)

        # 3. Mergear PDFs
        _merge_pdfs([excel_pdf, informe_pdf], merged_pdf)

        # 4. ZIP con xlsx + pdf
        zip_path  = os.path.join(tmp, "eecc.zip")
        with zipfile.ZipFile(zip_path, "w") as zf:
            zf.write(out_path, xlsx_name)
            if os.path.exists(merged_pdf):
                zf.write(merged_pdf, pdf_name)

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


def _libreoffice_pdf(input_path: str, out_dir: str, desired_path: str):
    """Convierte cualquier archivo a PDF con LibreOffice headless."""
    subprocess.run([
        "libreoffice", "--headless", "--convert-to", "pdf",
        "--outdir", out_dir, input_path
    ], capture_output=True, timeout=90)
    base = os.path.splitext(os.path.basename(input_path))[0]
    generated = os.path.join(out_dir, base + ".pdf")
    if os.path.exists(generated) and generated != desired_path:
        os.rename(generated, desired_path)


def _fill_informe(template_path: str, out_path: str,
                  empresa: str, cuit: str, domicilio: str, matricula_igj: str,
                  fecha_cierre: str, sipa_monto: str):
    """Rellena el template del Informe de Auditoría con los datos del cliente."""
    from docx import Document

    fecha = datetime.strptime(fecha_cierre, "%Y-%m-%d")
    fecha_larga  = f"{fecha.day} de {MONTHS_ES[fecha.month]} de {fecha.year}"
    mes_anio     = f"{MONTHS_ES[fecha.month]} de {fecha.year}"
    today        = date.today()
    fecha_inf    = f"{today.day} de {MONTHS_ES[today.month]} de {today.year}"
    sipa_fmt     = sipa_monto.strip() if sipa_monto.strip() else "[COMPLETAR SIPA]"

    replacements = {
        "{{EMPRESA}}.": f"{empresa}.",
        "{{EMPRESA}}":  empresa,
        "{{CUIT}}":     cuit,
        "{{DOMICILIO}}":     domicilio or "[DOMICILIO]",
        "{{MATRICULA_IGJ}}": matricula_igj or "[MATRÍCULA]",
        "{{FECHA_CIERRE_LARGA}}": fecha_larga,
        "{{MES_ANIO_CIERRE}}":    mes_anio,
        "{{SIPA_MONTO}}":         sipa_fmt,
        "{{FECHA_INFORME}}":      fecha_inf,
    }

    doc = Document(template_path)
    for para in doc.paragraphs:
        _replace_para(para, replacements)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _replace_para(para, replacements)
    doc.save(out_path)


def _replace_para(para, replacements):
    for key, val in replacements.items():
        if key in para.text:
            for run in para.runs:
                if key in run.text:
                    run.text = run.text.replace(key, val)


def _merge_pdfs(pdf_paths: list, output_path: str):
    """Mergea una lista de PDFs en uno solo."""
    try:
        from pypdf import PdfWriter, PdfReader
        writer = PdfWriter()
        for path in pdf_paths:
            if os.path.exists(path):
                reader = PdfReader(path)
                for page in reader.pages:
                    writer.add_page(page)
        with open(output_path, "wb") as f:
            writer.write(f)
    except Exception as e:
        print(f"[MERGE] Error: {e}")


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
