#!/usr/bin/env python3
"""
eecc_server.py — Servidor para generación de EECC
n8n llama a POST /generar con URLs de archivos y parámetros del cliente.
"""
import os, re, io, subprocess, tempfile, shutil, zipfile, urllib.request
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
    import shutil, subprocess as sp
    lo = shutil.which("libreoffice") or shutil.which("soffice") or "NOT FOUND"
    try:
        ver = sp.run([lo, "--version"], capture_output=True, text=True, timeout=10).stdout.strip()
    except Exception as e:
        ver = str(e)
    return {
        "status": "ok",
        "script": str(GEN_SCRIPT),
        "script_exists": GEN_SCRIPT.exists(),
        "template_exists": INFORME_TEMPLATE.exists(),
        "libreoffice_path": lo,
        "libreoffice_version": ver,
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
        year         = fecha_cierre.strip()[:4]
        xlsx_name    = f"EECC_{empresa_slug}_{year}.xlsx"
        pdf_name     = f"EECC_{empresa_slug}_{year}.pdf"
        excel_pdf    = os.path.join(tmp, "excel.pdf")
        informe_pdf  = os.path.join(tmp, "informe.pdf")
        merged_pdf   = os.path.join(tmp, pdf_name)

        # 1. Excel → PDF
        cc_key = os.environ.get("CLOUDCONVERT_API_KEY", "")
        lo = _find_libreoffice()
        if cc_key:
            _cloudconvert_pdf(cc_key, out_path, excel_pdf)
        elif lo:
            _libreoffice_convert(lo, out_path, tmp, excel_pdf)
        else:
            _xlsx_to_pdf(out_path, excel_pdf)

        # 2. Informe Word → rellenar → PDF
        if INFORME_TEMPLATE.exists():
            informe_filled = os.path.join(tmp, "informe_filled.docx")
            _fill_informe(str(INFORME_TEMPLATE), informe_filled,
                          empresa, cuit, domicilio, matricula_igj,
                          fecha_cierre.strip(), sipa_monto)
            if lo:
                _libreoffice_convert(lo, informe_filled, tmp, informe_pdf)
            else:
                _docx_to_pdf(informe_filled, informe_pdf)

        # 3. Mergear PDFs
        _merge_pdfs([excel_pdf, informe_pdf], merged_pdf)

        # 4. ZIP con xlsx + pdf
        zip_path = os.path.join(tmp, "eecc.zip")
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


def _cloudconvert_pdf(api_key: str, input_path: str, output_path: str):
    """Convierte a PDF via CloudConvert API (usa LibreOffice internamente)."""
    import cloudconvert
    cloudconvert.configure(api_key=api_key, sandbox=False)

    job = cloudconvert.Job.create(payload={
        "tasks": {
            "upload":  {"operation": "import/upload"},
            "convert": {"operation": "convert", "input": "upload",
                        "output_format": "pdf", "engine": "libreoffice"},
            "export":  {"operation": "export/url", "input": "convert"}
        }
    })

    upload_task = next(t for t in job["tasks"] if t["name"] == "upload")
    cloudconvert.Task.upload(file_name=input_path, task=upload_task)

    job = cloudconvert.Job.wait(id=job["id"])
    export_task = next(t for t in job["tasks"] if t["name"] == "export")
    url = export_task["result"]["files"][0]["url"]
    urllib.request.urlretrieve(url, output_path)


def _find_libreoffice():
    import shutil
    return shutil.which("libreoffice") or shutil.which("soffice")


def _libreoffice_convert(lo_bin: str, input_path: str, out_dir: str, desired_path: str):
    """Convierte un archivo a PDF con LibreOffice headless."""
    env = os.environ.copy()
    env["HOME"] = out_dir  # evita conflictos de perfil de usuario
    subprocess.run(
        [lo_bin, "--headless", "--norestore", "--convert-to", "pdf",
         "--outdir", out_dir, input_path],
        capture_output=True, timeout=120, env=env
    )
    base = os.path.splitext(os.path.basename(input_path))[0]
    generated = os.path.join(out_dir, base + ".pdf")
    if os.path.exists(generated) and generated != desired_path:
        os.rename(generated, desired_path)


def _xlsx_to_pdf(xlsx_path: str, pdf_path: str):
    """Convierte Excel a PDF: una página por solapa preservando estilos de xlsx2html."""
    from xlsx2html import xlsx2html as x2h
    from weasyprint import HTML
    from openpyxl import load_workbook
    from pypdf import PdfWriter, PdfReader
    import tempfile

    LANDSCAPE_SHEETS = {'EEPN', 'Anexo I', 'Anexo III'}

    wb = load_workbook(xlsx_path)
    sheet_pdfs = []

    for sheet_name in wb.sheetnames:
        buf = io.StringIO()
        try:
            x2h(xlsx_path, buf, sheet=sheet_name)
            full_html = buf.getvalue()

            landscape = sheet_name in LANDSCAPE_SHEETS
            pw = '297mm' if landscape else '210mm'
            ph = '210mm' if landscape else '297mm'

            # Inyectar @page y reducir font sin pisar estilos de xlsx2html
            inject = (
                f'<style>'
                f'@page {{ size: {pw} {ph}; margin: 0.7cm; }}'
                f'body {{ font-size: 6pt !important; }}'
                f'table {{ width: 100% !important; }}'
                f'td, th {{ overflow: hidden !important; white-space: normal !important; word-break: break-word !important; }}'
                f'</style>'
            )
            if '</head>' in full_html:
                full_html = full_html.replace('</head>', inject + '</head>')
            else:
                full_html = inject + full_html

            tmp = tempfile.mktemp(suffix='.pdf')
            HTML(string=full_html).write_pdf(tmp)
            sheet_pdfs.append(tmp)
        except Exception as e:
            print(f"[XLSX2PDF] {sheet_name}: {e}")

    writer = PdfWriter()
    for p in sheet_pdfs:
        if os.path.exists(p):
            for page in PdfReader(p).pages:
                writer.add_page(page)
            os.unlink(p)

    with open(pdf_path, 'wb') as f:
        writer.write(f)


def _docx_to_pdf(docx_path: str, pdf_path: str):
    """Convierte DOCX a PDF preservando alineación e imágenes."""
    import base64
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from weasyprint import HTML

    doc = Document(docx_path)
    NS_DRAW = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
    NS_A    = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
    NS_R    = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

    def _para_html(para):
        align_map = {
            WD_ALIGN_PARAGRAPH.RIGHT:   'right',
            WD_ALIGN_PARAGRAPH.CENTER:  'center',
            WD_ALIGN_PARAGRAPH.JUSTIFY: 'justify',
            WD_ALIGN_PARAGRAPH.LEFT:    'left',
        }
        align = align_map.get(para.alignment, 'justify')
        parts = []
        for run in para.runs:
            inline = run._element.find(f'.//{NS_DRAW}inline')
            if inline is not None:
                blip = run._element.find(f'.//{NS_A}blip')
                if blip is not None:
                    rId = blip.get(f'{{{NS_R}}}embed')
                    if rId:
                        img_part = doc.part.related_parts[rId]
                        b64 = base64.b64encode(img_part.blob).decode()
                        mime = img_part.content_type
                        parts.append(f'<img src="data:{mime};base64,{b64}" style="width:7.1cm;" />')
            else:
                text = run.text.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
                if not text:
                    continue
                s = ''
                if run.bold:    s += 'font-weight:bold;'
                if run.italic:  s += 'font-style:italic;'
                if run.underline: s += 'text-decoration:underline;'
                parts.append(f'<span style="{s}">{text}</span>' if s else text)
        content = ''.join(parts) if parts else '&nbsp;'
        return f'<p style="text-align:{align};margin:0.25em 0">{content}</p>'

    html_parts = [_para_html(p) for p in doc.paragraphs]

    css = '''
        @page { size: A4; margin: 2.5cm 2cm; }
        body { font-family: Arial, sans-serif; font-size: 11pt; line-height: 1.5; }
    '''
    full_html = f'<html><head><style>{css}</style></head><body>{"".join(html_parts)}</body></html>'
    HTML(string=full_html).write_pdf(pdf_path)


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
        "{{DOMICILIO}}":          domicilio or "[DOMICILIO]",
        "{{MATRICULA_IGJ}}":      matricula_igj or "[MATRÍCULA]",
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
