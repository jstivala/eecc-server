#!/usr/bin/env python3
"""
gen_eecc_v9.py — Generador genérico de Estados Contables (RT 54 / RP CPCECABA 460/2024)

Lee:
  * el Sumas y Saldos del ejercicio (Excel exportado del sistema: Cuenta | Saldo,
    sin columna de rubro, antes del cierre y en valores históricos),
  * el EECC del ejercicio anterior (PDF) para comparativos, PN de inicio y Anexo I.

Calcula (simplificaciones para entidades pequeñas, RT 54):
  * PN de inicio reexpresado con el coeficiente,
  * amortizaciones del ejercicio (Anexo I anterior + altas, año de alta completo),
  * RECPAM / resultados financieros y por tenencia por diferencia.

Genera un .xlsx (con hoja "Control" de clasificación) y un .pdf con:
  Carátula, ESP, ER, EEPN, EFE (método indirecto), Anexos y Notas.

Si algo no se puede resolver (cuenta desconocida, SS que no cuadra) termina con
código 2 y el motivo en stderr, para que el servidor marque "Error" en Notion.
"""
import argparse
import json
import re
import sys
import unicodedata
from datetime import date, datetime, timedelta

# ════════════════════════════════════════════════════════════════════════════
# Utilidades
# ════════════════════════════════════════════════════════════════════════════


class EECCError(Exception):
    pass


def _strip_char(c):
    d = unicodedata.normalize("NFKD", c)
    base = "".join(ch for ch in d if not unicodedata.combining(ch))
    if len(base) != 1:
        base = c
    return base.lower()


def norm(s):
    """minúsculas, sin acentos; conserva la longitud del texto."""
    return "".join(_strip_char(c) for c in str(s))


def skey(s):
    """clave tolerante a plurales: cada palabra sin 's' final."""
    return " ".join(w[:-1] if len(w) > 3 and w.endswith("s") else w for w in nkey(s).split())


def nkey(s):
    """clave de comparación: norm + espacios colapsados, sin $ ni puntuación final."""
    s = norm(s).replace("$", " ").replace("º", "o").replace("°", "o")
    s = re.sub(r"[\s.:]+$", "", s)
    return re.sub(r"\s+", " ", s).strip()


NUM_RE = re.compile(
    r'\(\s*-?\d{1,3}(?:[.,]\d{3})*[.,]\d{2}\s*\)'
    r'|-?\d{1,3}(?:[.,]\d{3})+[.,]\d{2}'
    r'|-?\d+[.,]\d{2}(?!\d)'
    r'|(?<![\w.,])"?[-–]"?(?![\w\d])'
)


def parse_num(tok):
    t = tok.strip().strip('"')
    if t in ("-", "–", ""):
        return 0.0
    neg = t.startswith("(") and t.endswith(")")
    t = t.strip("()").strip()
    if t.startswith("-"):
        neg, t = True, t[1:]
    m = re.match(r"^([\d.,]*?)([.,])(\d{2})$", t)
    if not m:
        return None
    ent = re.sub(r"[.,]", "", m.group(1))
    v = float(f"{ent or 0}.{m.group(3)}")
    return -v if neg else v


def tokens_after(line, pos):
    out = []
    for m in NUM_RE.finditer(line, pos):
        v = parse_num(m.group())
        if v is not None:
            out.append(v)
    return out


def fmt_ar(v, dash=True):
    if v is None:
        return ""
    if abs(v) < 0.005:
        return "-" if dash else "0,00"
    s = f"{abs(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"({s})" if v < 0 else s


def fmt_cuit(c):
    d = re.sub(r"\D", "", str(c))
    return f"{d[:2]}-{d[2:10]}-{d[10:]}" if len(d) == 11 else str(c)


MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
         "septiembre", "octubre", "noviembre", "diciembre"]


def fecha_larga(d):
    return f"{d.day} de {MESES[d.month - 1]} de {d.year}"


# ════════════════════════════════════════════════════════════════════════════
# Clasificación de cuentas
# ════════════════════════════════════════════════════════════════════════════
RUBROS = {
    "caja": ("A", "Caja y bancos"),
    "inv": ("A", "Inversiones"),
    "cv": ("A", "Créditos por ventas"),
    "oc": ("A", "Otros créditos"),
    "bc": ("A", "Bienes de cambio"),
    "bu": ("A", "Bienes de uso"),
    "dc": ("P", "Comerciales"),
    "prest": ("P", "Préstamos"),
    "rem": ("P", "Remuneraciones y cargas sociales"),
    "df": ("P", "Cargas fiscales"),
    "ds": ("P", "Deudas con socios"),
    "od": ("P", "Otras deudas"),
    "cap": ("PN", "Capital social"),
    "aj": ("PN", "Ajuste de capital"),
    "apirr": ("PN", "Aportes irrevocables"),
    "resleg": ("PN", "Reservas"),
    "rna": ("PN", "Resultados no asignados"),
    "ventas": ("R", "Ventas"),
    "costo": ("R", "Costo"),
    "gcom": ("R", "Gastos de comercialización"),
    "gadm": ("R", "Gastos de administración"),
    "amort": ("R", "Amortizaciones"),
    "rfyt": ("R", "Resultados financieros y por tenencia"),
    "oi": ("R", "Otros ingresos y egresos"),
    "ig": ("R", "Impuesto a las ganancias"),
}

RULES = [
    # Patrimonio neto
    ("aj", r"ajuste (de|del) capital"),
    ("apirr", r"aportes? irrevocables?"),
    ("resleg", r"^reserva"),
    ("cap", r"^capital( social| suscripto| integrado)?\b"),
    ("rna", r"resultados? no asignados?|^rna$|resultados? del ejercicio|resultados? acumulados?|"
            r"resultados? (de )?ejercicios? anteriores?"),
    # Resultados financieros (se recalculan por diferencia)
    ("rfyt", r"recpam|diferencias? de cambio|redondeo|resultados? por tenencia|resultados? financieros|"
             r"^intereses|intereses (ganados|perdidos|cobrados|pagados|bancarios)"),
    # Bienes de uso y amortizaciones
    ("bu", r"amortizacion(es)? acumulada|amort\.? acum|depreciacion acumulada"),
    ("amort", r"amortizacion|depreciacion"),
    ("bu", r"muebles|\butiles\b|rodados?$|^rodado|instalaciones|equipos?\b|maquinari|herramientas|"
           r"inmueble|edificio|terreno|computacion|computadoras|^mejoras|hardware"),
    # Pasivo (antes que resultados: "a pagar")
    ("rem", r"(sueldos?|jornales|cargas sociales|sindicato|seguro de vida|sac\b|aguinaldo|vacaciones|"
            r"\bart\b|obra social|f\.? ?931|aportes y contribuciones).*(a pagar|devengad)|"
            r"remuneraciones a pagar"),
    ("df", r"(iva|iibb|ingresos brutos|ganancias|impuestos?|tasas?|abl|monotributo|bienes personales|"
           r"sicore|sircreb|percepcion|retencion).*a pagar|debito fiscal|"
           r"(retencion|retenciones|percepcion|percepciones).*efectuad|"
           r"provision (de |del )?(ig|impuesto|ganancias)|plan(es)? de pagos?|moratoria"),
    ("ds", r"deuda.*socio|socio.*a pagar|cuenta particular|cuentas? (corriente )?socios?|"
           r"dividendos a pagar|honorarios (de )?directores a pagar|aportes? a reintegrar"),
    ("prest", r"prestamo|descubierto|adelanto en cuenta"),
    ("dc", r"proveedor|acreedores comerciales|documentos a pagar|cheques diferidos|tarjetas? de credito"),
    ("od", r"acreedores varios|otras deudas|anticipos? de clientes|deudas varias"),
    # Activo
    ("caja", r"^caja|banco|mercado ?pago|fondo fijo|valores a depositar|moneda extranjera|^efectivo"),
    ("inv", r"plazo fijo|fondos? comun|fondos? de inversion|^fci|^inversion|cauciones?|^acciones|^bonos"),
    ("cv", r"deudores por ventas?|^clientes|documentos a cobrar|cheques en cartera|tarjetas? a cobrar|"
           r"creditos por ventas?"),
    ("bc", r"mercaderias?$|^mercaderia|bienes de cambio|materias primas|productos (en proceso|terminados)|"
           r"^stock|existencias?"),
    ("oc", r"iva (saldo|credito|a favor)|saldo (tecnico|libre)|libre disponibilidad|^saf\b|saldo a favor|"
           r"(retencion|retenciones|percepcion|percepciones).*(sufrid|iibb|iva|ganancias)|"
           r"anticipos?|pagos? a cuenta|25\.?413|\b(db|debitos?) y (cr|creditos?)|credito fiscal|"
           r"^socio|gastos pagados por adelantado|creditos? (varios|diversos)|deudores varios|"
           r"prestamos? a socios|a devengar|depositos? en garantia|reintegros? a cobrar"),
    # Resultados
    ("costo", r"^compras?\b|compra de|^costo|\bcmv\b|mercaderias vendidas"),
    ("ventas", r"^ventas?\b|venta de|ingresos? por|servicios prestados|honorarios (ganados|cobrados|percibidos)|"
               r"^facturacion|ingresos? (operativos|de explotacion)"),
    ("ig", r"impuesto a las ganancias"),
    ("oi", r"otros ingresos|otros egresos|resultado (por )?venta (de )?bienes de uso|recupero"),
    ("gcom", r"iibb|ingresos brutos|publicidad|propaganda|comercializacion|fletes|comisiones|marketing|"
             r"convenio multilateral"),
    ("gadm", r"gasto|sueldo|jornal|cargas sociales|honorario|servicios publicos|alquiler|mantenimiento|"
             r"libreria|papeleria|seguro|impuesto|tasa|movilidad|viatico|combustible|telefon|internet|"
             r"\bluz\b|\bagua\b|\bgas\b|abl|monotributo|sistema|software|donacion|multa|refrigerio|"
             r"limpieza|correo|vigilancia|capacitacion|insumos|reparacion|expensas|bancari|comision|"
             r"cuota|suscripcion|indemnizacion|obra social|contribucion|aportes|remuneracion|premio"),
]
RULES_C = [(k, re.compile(p)) for k, p in RULES]


def clasificar(nombre):
    n = nkey(nombre)
    for k, rx in RULES_C:
        if rx.search(n):
            return k
    return None


MONEDA_EXT = re.compile(r"\busd\b|u\$s|dolar|euro|moneda extranjera|\bme\b")

BU_TIPOS = [
    ("Muebles y útiles", r"mueble|utiles", 10.0),
    ("Rodados", r"rodado", 20.0),
    ("Instalaciones", r"instalacion", 10.0),
    ("Equipos de computación", r"computacion|computadora|hardware", 33.33),
    ("Equipos y herramientas", r"equipo|herramienta|maquinari", 10.0),
    ("Inmueble terreno", r"terreno", 0.0),
    ("Inmueble edificio", r"inmueble|edificio|mejoras", 2.0),
]


def tipo_bu(nombre):
    n = nkey(nombre)
    for t, p, _ in BU_TIPOS:
        if re.search(p, n):
            return t
    return "Otros bienes de uso"


def tasa_default(tipo):
    for t, _, r in BU_TIPOS:
        if t == tipo:
            return r
    return 10.0


# ════════════════════════════════════════════════════════════════════════════
# Lectura del Sumas y Saldos
# ════════════════════════════════════════════════════════════════════════════
def leer_ss(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows = [list(r) for r in ws.iter_rows(values_only=True)]

    hdr_i, c_cta, c_sal, c_debe, c_haber = None, None, None, None, None
    for i, r in enumerate(rows[:15]):
        labels = [nkey(c) if isinstance(c, str) else "" for c in r]
        if any(l == "cuenta" or l.startswith("cuenta") or l == "descripcion" for l in labels):
            hdr_i = i
            for j, l in enumerate(labels):
                if c_cta is None and (l.startswith("cuenta") or l == "descripcion"):
                    c_cta = j
                if "saldo" in l:
                    c_sal = j
                if l == "debe" or l.startswith("debe"):
                    c_debe = j
                if l == "haber" or l.startswith("haber"):
                    c_haber = j
            break
    if hdr_i is None:
        # sin encabezado: primera columna de texto + primera numérica
        hdr_i = -1
        for r in rows:
            txt = [j for j, c in enumerate(r) if isinstance(c, str) and c.strip()]
            num = [j for j, c in enumerate(r) if isinstance(c, (int, float))]
            if txt and num:
                c_cta, c_sal = txt[0], num[-1]
                break
    if c_cta is None or (c_sal is None and (c_debe is None or c_haber is None)):
        raise EECCError("No pude identificar las columnas Cuenta / Saldo en el sumas y saldos.")

    cuentas = []
    for r in rows[hdr_i + 1:]:
        if c_cta >= len(r):
            continue
        nom = r[c_cta]
        if not isinstance(nom, str) or not nom.strip():
            continue
        if re.match(r"^(total|totales|sumas)\b", nkey(nom)):
            continue
        if c_sal is not None:
            v = r[c_sal] if c_sal < len(r) else None
        else:
            d = r[c_debe] or 0
            h = r[c_haber] or 0
            v = (d if isinstance(d, (int, float)) else 0) - (h if isinstance(h, (int, float)) else 0)
        if isinstance(v, str):
            v = parse_num(v)
        if v is None:
            v = 0.0
        cuentas.append((nom.strip(), float(v)))

    total = sum(v for _, v in cuentas)
    if abs(total) > 1.0:
        raise EECCError(
            f"El sumas y saldos no cuadra: saldos deudores menos acreedores = {fmt_ar(total)}. "
            "Revisar la exportación del sistema.")
    return cuentas


# ════════════════════════════════════════════════════════════════════════════
# Lectura del EECC anterior (PDF)
# ════════════════════════════════════════════════════════════════════════════
def texto_pdf(path):
    """pdftotext -layout o pypdf (modo layout): respetan columnas sin partir los números."""
    import shutil
    import subprocess
    if shutil.which("pdftotext"):
        out = subprocess.run(["pdftotext", "-layout", path, "-"], capture_output=True, text=True).stdout
        if out.strip():
            return out
    try:
        import pypdf
        out = "\n".join(pg.extract_text(extraction_mode="layout") or "" for pg in pypdf.PdfReader(path).pages)
        if out.strip():
            return out
    except Exception:
        pass
    import pdfplumber
    with pdfplumber.open(path) as pdf:
        return "\n".join((p.extract_text() or "") for p in pdf.pages)


class Prior:
    """Valores del EECC anterior (columna 'Actual' de ese balance, moneda de su cierre)."""

    LABELS = {
        "caja": r"caja y bancos",
        "inv": r"^\s*inversiones|\binversiones\s*(\(|$|-|\")",
        "cv": r"creditos por ventas",
        "oc": r"otros creditos",
        "bc": r"bienes de cambio",
        "bu": r"bienes de uso",
        "ta": r"total del activo(?! (no )?corriente)",
        "dc": r"comerciales",
        "prest": r"prestamos",
        "rem": r"remuneraciones y cargas sociales",
        "df": r"cargas fiscales",
        "ds": r"deudas sociales|sociales\s*\(",
        "od": r"otras deudas",
        "tp": r"total del pasivo(?! (no )?corriente)(?! y)",
        "pn": r"patrimonio neto",
        "ventas": r"ventas netas",
        "costo": r"costo de (los )?(bienes|mercaderias|servicios)",
        "gcom": r"gastos de comercializacion",
        "gadm": r"gastos de administracion",
        "og": r"^\s*otros gastos",
        "rfyt": r"resultados financieros y por tenencia",
        "oi": r"otros ingresos y egresos",
        "ig": r"^\s*impuesto a las ganancias",
        "res": r"(ganancia|resultado).{0,12}del ejercicio",
        "efe_ini": r"efectivo (y sus equivalentes )?al inicio",
        "efe_cie": r"efectivo (y sus equivalentes )?al cierre",
        "efe_op": r"flujo neto.*actividades operativas",
        "efe_inv": r"flujo neto.*actividades de inversion",
        "efe_fin": r"flujo neto.*actividades de financiacion",
        "efe_dep": r"depreciacion|amortizacion de bienes",
        "efe_cv": r"disminucion\)? en creditos por ventas",
        "efe_oc": r"disminucion\)? en otros creditos",
        "efe_bc": r"disminucion\)? en bienes de cambio",
        "efe_deu": r"disminucion\)? en deudas",
    }

    def __init__(self, path):
        self.text = texto_pdf(path) if path else ""
        self.lines = self.text.splitlines()
        self.nlines = [norm(l) for l in self.lines]
        self.v = {}
        self.warn = []
        if not self.text.strip():
            return
        for k, pat in self.LABELS.items():
            self.v[k] = self._find(pat)
        # notas: cuentas
        self.notas_ini = 0
        for i, nl in enumerate(self.nlines):
            if "notas a los estados contables" in nl and i > 20:
                self.notas_ini = i
                break
        self.anexo_bu = self._leer_anexo_bu()
        self.normas_rt54 = bool(re.search(r"tecnica\s*(n.{0,3})?\s*54|rt\s*54", norm(self.text)))
        m = re.search(r"actividad principal\s*:?\s*(.+)", self.text, re.I)
        self.actividad = ""
        if m:
            act = m.group(1).strip()
            idx = self.text[:m.end()].count("\n")
            if idx + 1 < len(self.lines):
                nxt = self.lines[idx + 1].strip()
                if nxt and nxt.isupper() and ":" not in nxt and len(nxt) < 60:
                    act += " " + nxt
            self.actividad = re.sub(r"\s+", " ", act)

    def _find(self, pat, start=0):
        rx = re.compile(pat)
        for i in range(start, len(self.lines)):
            m = rx.search(self.nlines[i])
            if not m:
                continue
            toks = tokens_after(self.lines[i], m.end())
            if toks:
                return toks[0]
        return None

    def g(self, k):
        return self.v.get(k) or 0.0

    def cuenta(self, nombre):
        """Saldo de una cuenta en las notas del EECC anterior (match exacto del nombre)."""
        key = skey(nombre)
        for i in range(self.notas_ini, len(self.lines)):
            line = self.lines[i]
            m = NUM_RE.search(line)
            while m and parse_num(m.group()) is None:
                m = NUM_RE.search(line, m.end())
            if not m:
                continue
            label = skey(line[:m.start()])
            if label == key:
                return parse_num(m.group())
        return None

    def _leer_anexo_bu(self):
        ini = None
        for i, nl in enumerate(self.nlines):
            if re.search(r"valores\s+de\s+(incorporacion|origen)", nl):
                ini = i
                break
        out = {}
        if ini is None:
            return out
        for i in range(ini, min(ini + 40, len(self.lines))):
            nl = self.nlines[i]
            if re.search(r"anexo ii|costo de (las )?mercaderias|informacion requerida|^\s*totales", nl):
                break
            label_m = re.match(r"\s*([a-z][a-z .y]+?)\s{2,}", nl)
            if not label_m:
                continue
            toks = tokens_after(self.lines[i], label_m.end())
            if len(toks) < 9:
                continue
            tipo = tipo_bu(label_m.group(1))
            origen = toks[3]
            acum = toks[-2]
            neto = toks[-1]
            tasa = toks[6] if len(toks) >= 10 else 0.0
            if 0 < tasa <= 1:
                tasa *= 100
            if origen or neto:
                out[tipo] = dict(origen=origen, acum=acum, neto=neto, tasa=tasa)
        return out


# ════════════════════════════════════════════════════════════════════════════
# Cálculo
# ════════════════════════════════════════════════════════════════════════════
def calcular(a, cuentas, prior):
    COF = a.cof
    rx = lambda v: round((v or 0.0) * COF, 2)
    warn = []
    cierre = datetime.strptime(a.fecha_cierre, "%Y-%m-%d").date()
    try:
        inicio = cierre.replace(year=cierre.year - 1) + timedelta(days=1)
    except ValueError:
        inicio = cierre - timedelta(days=364)
    if a.fecha_inicio:
        inicio = datetime.strptime(a.fecha_inicio, "%Y-%m-%d").date()
    cierre_ant = inicio - timedelta(days=1)
    tiene_prior = bool(prior.text.strip()) and a.nro_ejercicio > 1

    # ── clasificación ──
    clas, desconocidas = [], []
    for nom, sal in cuentas:
        k = clasificar(nom)
        if k == "caja" and sal < -0.5:
            k = "prest"
        if k is None:
            if abs(sal) > 0.005:
                desconocidas.append((nom, sal))
            continue
        clas.append((nom, sal, k))
    if desconocidas:
        lista = "; ".join(f"{n} ({fmt_ar(s)})" for n, s in desconocidas)
        raise EECCError("Cuentas del sumas y saldos que no pude clasificar en un rubro: " + lista)

    def suma(k, signo=1):
        return signo * sum(s for _, s, kk in clas if kk == k)

    moneda_ext = any(MONEDA_EXT.search(nkey(n)) and abs(s) > 0.005 for n, s, _ in clas)

    # ── prior (homogéneo al cierre actual) ──
    P = {k: rx(prior.g(k)) if tiene_prior else 0.0 for k in Prior.LABELS}
    for k in ("costo", "gcom", "gadm", "og", "ig"):
        P[k] = abs(P[k])

    # ── Bienes de uso ──
    bu_ss = {}
    for n, s, k in clas:
        if k == "bu":
            bu_ss.setdefault(tipo_bu(n), 0.0)
            bu_ss[tipo_bu(n)] += s
    anexo_p = prior.anexo_bu if tiene_prior else {}
    if tiene_prior and P["bu"] and not anexo_p:
        warn.append("No pude leer el Anexo I del EECC anterior: se tomó el neto anterior como valor de origen "
                    "y tasas estándar. Revisar el Anexo de bienes de uso.")
        tot_ss = sum(v for v in bu_ss.values() if v > 0) or 1.0
        tipos = [t for t, v in bu_ss.items() if v > 0] or ["Otros bienes de uso"]
        anexo_p = {t: dict(origen=prior.g("bu") * (bu_ss.get(t, tot_ss) / tot_ss), acum=0.0,
                           neto=prior.g("bu") * (bu_ss.get(t, tot_ss) / tot_ss), tasa=0.0) for t in tipos}
    amort_booked = suma("amort")
    bu_rows = []
    for tipo in sorted(set(bu_ss) | set(anexo_p), key=lambda t: [x[0] for x in BU_TIPOS].index(t)
                       if t in [x[0] for x in BU_TIPOS] else 99):
        p = anexo_p.get(tipo, dict(origen=0.0, acum=0.0, neto=0.0, tasa=0.0))
        tasa = p["tasa"] or tasa_default(tipo)
        net_ss = bu_ss.get(tipo, 0.0)
        altas_nom = net_ss - p["neto"]
        if abs(altas_nom) < 1:
            altas_nom = 0.0
        o_ini, a_ini = rx(p["origen"]), rx(p["acum"])
        altas_h = bajas_h = 0.0
        if altas_nom > 0:
            altas_h = round(altas_nom * COF ** 0.5, 2)
        elif altas_nom < 0:
            bajas_h = round(-altas_nom * COF, 2)
            warn.append(f"Bienes de uso '{tipo}': el saldo bajó respecto del ejercicio anterior; "
                        "se expuso como baja. Verificar.")
        o_cie = o_ini + altas_h - bajas_h
        amort = round(min(o_cie * tasa / 100, max(o_cie - a_ini, 0)), 2) if not amort_booked else 0.0
        bu_rows.append(dict(tipo=tipo, o_ini=o_ini, altas=altas_h, bajas=bajas_h, o_cie=o_cie,
                            a_ini=a_ini, tasa=tasa, amort=amort, a_cie=a_ini + amort,
                            neto=o_cie - a_ini - amort, neto_ant=rx(p["neto"])))
    if amort_booked:
        tot_o = sum(r["o_cie"] for r in bu_rows) or 1
        for r in bu_rows:
            r["amort"] = round(amort_booked * r["o_cie"] / tot_o, 2)
            r["a_cie"] = r["a_ini"] + r["amort"]
            r["neto"] = r["o_cie"] - r["a_cie"]
    amort_ej = sum(r["amort"] for r in bu_rows)
    altas_tot = sum(r["altas"] for r in bu_rows)
    bajas_tot = sum(r["bajas"] for r in bu_rows)

    # ── Activo / Pasivo al cierre ──
    C = dict(caja=suma("caja"), inv=suma("inv"), cv=suma("cv"), oc=suma("oc"), bc=suma("bc"),
             bu=sum(r["neto"] for r in bu_rows),
             dc=suma("dc", -1), prest=suma("prest", -1), rem=suma("rem", -1), df=suma("df", -1),
             ds=suma("ds", -1), od=suma("od", -1))
    for k in ("cv", "oc", "bc", "inv"):
        if C[k] < -0.5:
            warn.append(f"El rubro {RUBROS[k][1]} quedó con saldo negativo ({fmt_ar(C[k])}).")
    C["ac"] = C["caja"] + C["inv"] + C["cv"] + C["oc"] + C["bc"]
    C["anc"] = C["bu"]
    C["ta"] = C["ac"] + C["anc"]
    C["tp"] = C["dc"] + C["prest"] + C["rem"] + C["df"] + C["ds"] + C["od"]
    C["pn"] = C["ta"] - C["tp"]

    # ── Patrimonio neto de inicio ──
    cap_ss = suma("cap", -1)
    cap = a.cap_nominal if a.cap_nominal else cap_ss
    if a.cap_nominal and abs(cap_ss - a.cap_nominal) > 1:
        warn.append(f"El capital del sumas y saldos ({fmt_ar(cap_ss)}) no coincide con el cargado en Notion "
                    f"({fmt_ar(a.cap_nominal)}). Se usó el de Notion.")
    aj_p, apirr_p, res_p, rna_p = suma("aj", -1), suma("apirr", -1), suma("resleg", -1), suma("rna", -1)
    pn_ss = cap_ss + aj_p + apirr_p + res_p + rna_p
    aportes_nom = 0.0
    if tiene_prior:
        pn_prior = prior.g("pn")
        dif = pn_ss - pn_prior
        if abs(dif) > max(1.0, abs(pn_prior) * 0.001):
            if dif > 0 and abs(cap_ss - cap) < 1:
                aportes_nom = dif
                warn.append(f"El PN del sumas y saldos supera al del EECC anterior en {fmt_ar(dif)}; "
                            "se expuso como aporte del ejercicio. Verificar.")
            else:
                warn.append(f"El PN del sumas y saldos ({fmt_ar(pn_ss)}) no coincide con el PN del EECC "
                            f"anterior ({fmt_ar(pn_prior)}). Diferencia {fmt_ar(dif)}.")
        # inicio = PN anterior (sin los aportes nuevos)
        base_cap, base_aj, base_apirr, base_res = cap, aj_p, apirr_p, res_p
        base_rna = pn_prior - (cap + aj_p + apirr_p + res_p) if aportes_nom == 0 else rna_p
        I = dict(cap=cap, aj=round((base_cap + base_aj) * COF - base_cap, 2), apirr=rx(base_apirr),
                 resleg=rx(base_res), rna=rx(base_rna))
    else:
        aportes_nom = pn_ss
        I = dict(cap=0.0, aj=0.0, apirr=0.0, resleg=0.0, rna=0.0)
    I["aportes"] = I["cap"] + I["aj"] + I["apirr"]
    I["pn"] = I["aportes"] + I["resleg"] + I["rna"]
    aportes_h = round(aportes_nom, 2)

    # ── Resultados ──
    ventas = suma("ventas", -1)
    compras = suma("costo")
    bc_ini, bc_cie = P["bc"], C["bc"]
    usa_inventario = bool(bc_ini or bc_cie)
    costo = (bc_ini + compras - bc_cie) if usa_inventario else compras
    gcom = suma("gcom")
    gadm_ss = suma("gadm")
    gadm = gadm_ss + amort_ej + (amort_booked if amort_booked else 0)
    oi = suma("oi", -1)
    ig = suma("ig")
    res = C["pn"] - I["pn"] - aportes_h
    rfyt = round(res - (ventas - costo - gcom - gadm + oi - ig), 2)
    R = dict(ventas=ventas, costo=costo, bruta=ventas - costo, gcom=gcom, gadm=gadm, og=0.0,
             rfyt=rfyt, oi=oi, antes_ig=res + ig, ig=ig, res=res, compras=compras,
             bc_ini=bc_ini, bc_cie=bc_cie, usa_inv=usa_inventario)

    # ── Comparativos (ejercicio anterior reexpresado) ──
    A = dict(P)
    if tiene_prior:
        A["ac"] = A["caja"] + A["inv"] + A["cv"] + A["oc"] + A["bc"]
        A["anc"] = A["bu"]
        A["ta"] = A["ta"] or (A["ac"] + A["anc"])
        A["tp"] = A["dc"] + A["prest"] + A["rem"] + A["df"] + A["ds"] + A["od"]
        A["pn"] = A["pn"] or (A["ta"] - A["tp"])
        A["bruta"] = A["ventas"] - A["costo"]
        A["antes_ig"] = A["res"] + A["ig"]
        if abs((A["ta"] - A["tp"]) - A["pn"]) > max(2.0, abs(A["pn"]) * 0.001):
            warn.append("En el EECC anterior no pude reconstruir Activo − Pasivo = PN; revisar comparativos.")

    # ── EFE comparativo ──
    AE = {}
    if tiene_prior:
        if a.nro_ejercicio - 1 == 1:
            # el ejercicio anterior fue el primero: saldos de inicio en cero, se reconstruye completo
            dep = sum(rx(p["acum"]) for p in anexo_p.values())
            AE = dict(ini=0.0, cie=A["caja"], res=A["res"], amort=dep, cv=-A["cv"], oc=-A["oc"], bc=-A["bc"],
                      deu=A["dc"] + A["rem"] + A["df"] + A["od"], bu=-(A["bu"] + dep), invs=-A["inv"],
                      aportes=A["pn"] - A["res"], prest=A["prest"], ds=A["ds"])
            AE["op"] = AE["res"] + AE["amort"] + AE["cv"] + AE["oc"] + AE["bc"] + AE["deu"]
            AE["inv"] = AE["bu"] + AE["invs"]
            AE["fin"] = AE["aportes"] + AE["prest"] + AE["ds"]
        else:
            m = dict(ini="efe_ini", cie="efe_cie", res="res", amort="efe_dep", cv="efe_cv", oc="efe_oc",
                     bc="efe_bc", deu="efe_deu", op="efe_op", inv="efe_inv", fin="efe_fin")
            AE = {k: (A.get(v) if prior.v.get(v) is not None else None) for k, v in m.items()}
            AE["res"] = A["res"]
        AE["var"] = (AE.get("cie") or 0) - (AE.get("ini") or 0)
        if AE.get("op") is not None and AE.get("inv") is not None and AE.get("fin") is not None:
            if abs(AE["op"] + AE["inv"] + AE["fin"] - AE["var"]) > max(2.0, abs(AE["var"]) * 0.01):
                warn.append("El flujo de efectivo del EECC anterior no cierra; revisar la columna comparativa.")

    # ── EFE (método indirecto, moneda homogénea) ──
    d = lambda k: C[k] - A.get(k, 0.0)
    E = dict(ini=A.get("caja", 0.0), cie=C["caja"])
    E["var"] = E["cie"] - E["ini"]
    E["res"] = res
    E["amort"] = amort_ej
    E["cv"] = -d("cv")
    E["oc"] = -d("oc")
    E["bc"] = -d("bc")
    E["deu"] = d("dc") + d("rem") + d("df") + d("od")
    E["op"] = E["res"] + E["amort"] + E["cv"] + E["oc"] + E["bc"] + E["deu"]
    E["bu"] = -(C["bu"] + amort_ej - A.get("bu", 0.0))
    E["invs"] = -d("inv")
    E["inv"] = E["bu"] + E["invs"]
    E["aportes"] = aportes_h
    E["prest"] = d("prest")
    E["ds"] = d("ds")
    E["fin"] = E["aportes"] + E["prest"] + E["ds"]
    E["chk"] = E["op"] + E["inv"] + E["fin"] - E["var"]
    if abs(E["chk"]) > 1:
        raise EECCError(f"El flujo de efectivo no cierra por {fmt_ar(E['chk'])} (error interno).")

    # ── Controles ──
    if abs(C["ta"] - C["tp"] - C["pn"]) > 0.01:
        raise EECCError("Activo − Pasivo ≠ PN (error interno).")
    if abs(I["pn"] + aportes_h + res - C["pn"]) > 0.01:
        raise EECCError("La evolución del PN no cierra (error interno).")
    if C["pn"] < 0:
        warn.append(f"El patrimonio neto al cierre es negativo ({fmt_ar(C['pn'])}).")

    return dict(cuentas=clas, C=C, A=A, I=I, R=R, E=E, AE=AE, bu=bu_rows, amort_ej=amort_ej,
                aportes=aportes_h, inicio=inicio, cierre=cierre, cierre_ant=cierre_ant,
                tiene_prior=tiene_prior, moneda_ext=moneda_ext, warn=warn, cof=COF,
                primer_rt54=not (tiene_prior and prior.normas_rt54),
                actividad=a.actividad or prior.__dict__.get("actividad", ""),
                prior=prior)


# ════════════════════════════════════════════════════════════════════════════
# Armado de los estados (modelo de filas común a XLSX y PDF)
# ════════════════════════════════════════════════════════════════════════════
def armar(a, X):
    C, A, I, R, E = X["C"], X["A"], X["I"], X["R"], X["E"]
    comp = X["tiene_prior"]
    hdr2 = ["Actual", "Anterior"] if comp else ["Actual"]
    v = lambda c, k: [C.get(k, 0.0), A.get(k, 0.0)] if comp else [C.get(k, 0.0)]

    # numeración de anexos
    anexos = []
    if X["bu"]:
        anexos.append("bu")
    if R["usa_inv"]:
        anexos.append("costo")
    anexos.append("gastos")
    ROM = ["I", "II", "III", "IV"]
    anx = {k: ROM[i] for i, k in enumerate(anexos)}

    # numeración de notas de composición (3.x)
    comp_rubros = []
    for k in ["caja", "inv", "cv", "oc", "bc", "dc", "prest", "rem", "df", "ds", "od"]:
        if abs(C.get(k, 0)) > 0.005 or abs(A.get(k, 0)) > 0.005:
            comp_rubros.append(k)
    comp_rubros.append("rfyt")
    nota_n = {k: f"3.{i + 1}" for i, k in enumerate(comp_rubros)}
    nref = lambda k: f"(Nota {nota_n[k]})" if k in nota_n else ""

    S = []

    # ── ESP ──
    rows = []
    rows.append(("sec", "ACTIVO", None))
    rows.append(("sub", "Activo corriente", None))
    for k, lab in [("caja", "Caja y bancos"), ("inv", "Inversiones"), ("cv", "Créditos por ventas"),
                   ("oc", "Otros créditos"), ("bc", "Bienes de cambio")]:
        rows.append(("det", f"{lab} {nref(k)}".strip(), v(C, k)))
    rows.append(("tot", "Total del activo corriente", v(C, "ac")))
    rows.append(("sub", "Activo no corriente", None))
    rows.append(("det", f"Bienes de uso (Anexo {anx['bu']})" if "bu" in anx else "Bienes de uso", v(C, "bu")))
    rows.append(("tot", "Total del activo no corriente", v(C, "anc")))
    rows.append(("grand", "TOTAL DEL ACTIVO", v(C, "ta")))
    rows.append(("blank", "", None))
    rows.append(("sec", "PASIVO", None))
    rows.append(("sub", "Pasivo corriente", None))
    for k, lab in [("dc", "Deudas comerciales"), ("prest", "Préstamos"),
                   ("rem", "Remuneraciones y cargas sociales"), ("df", "Cargas fiscales"),
                   ("ds", "Deudas con socios"), ("od", "Otras deudas")]:
        rows.append(("det", f"{lab} {nref(k)}".strip(), v(C, k)))
    rows.append(("tot", "Total del pasivo corriente", v(C, "tp")))
    rows.append(("grand", "TOTAL DEL PASIVO", v(C, "tp")))
    rows.append(("blank", "", None))
    rows.append(("grand", "PATRIMONIO NETO (según el EEPN)", v(C, "pn")))
    rows.append(("grand", "TOTAL DEL PASIVO Y PATRIMONIO NETO", [x + y for x, y in zip(v(C, "tp"), v(C, "pn"))]))
    S.append(dict(key="ESP", title="ESTADO DE SITUACIÓN PATRIMONIAL", cols=hdr2, rows=rows, land=False))

    # ── ER ──
    rr = lambda k: [R[k], A.get(k, 0.0)] if comp else [R[k]]
    neg = lambda l: [-x for x in l]
    rows = [
        ("det", "Ventas netas de bienes (o servicios)", rr("ventas")),
        ("det", "Costo de los bienes vendidos (o servicios prestados)" +
         (f" (Anexo {anx['costo']})" if "costo" in anx else ""), neg(rr("costo"))),
        ("tot", "Ganancia (Pérdida) bruta", rr("bruta")),
        ("det", f"Gastos de comercialización (Anexo {anx['gastos']})", neg(rr("gcom"))),
        ("det", f"Gastos de administración (Anexo {anx['gastos']})", neg(rr("gadm"))),
        ("det", "Otros ingresos y egresos", rr("oi")),
        ("det", f"Resultados financieros y por tenencia (incluido el RECPAM) {nref('rfyt')}", rr("rfyt")),
        ("tot", "Ganancia (Pérdida) antes del impuesto a las ganancias", rr("antes_ig")),
        ("det", "Impuesto a las ganancias", neg(rr("ig"))),
        ("grand", "GANANCIA (PÉRDIDA) DEL EJERCICIO", rr("res")),
    ]
    S.append(dict(key="ER", title="ESTADO DE RESULTADOS", cols=hdr2, rows=rows, land=False))

    # ── EEPN ──
    cols = ["Capital social", "Ajuste de capital", "Aportes irrevocables", "Total aportes",
            "Reservas", "Resultados no asignados", "Total actual"] + (["Total anterior"] if comp else [])
    # comparativo (ejercicio anterior)
    if comp:
        prior_first = a.nro_ejercicio - 1 == 1
        a_res = A.get("res", 0.0)
        a_pn = A.get("pn", 0.0)
        a_ini = 0.0 if prior_first else a_pn - a_res
        a_apo = (a_pn - a_res) if prior_first else 0.0

    def eep(det, cap, aj, apirr, resl, rna, ant=None):
        tot_apo = cap + aj + apirr
        vals = [cap, aj, apirr, tot_apo, resl, rna, tot_apo + resl + rna]
        if comp:
            vals.append(ant)
        return vals

    rows = []
    rows.append(("det", "Saldos al inicio del ejercicio",
                 eep("", I["cap"], I["aj"], I["apirr"], I["resleg"], I["rna"], a_ini if comp else None)))
    rows.append(("det", "Modificación de saldos de inicio", eep("", 0, 0, 0, 0, 0, 0.0 if comp else None)))
    rows.append(("tot", "Saldos al inicio modificados",
                 eep("", I["cap"], I["aj"], I["apirr"], I["resleg"], I["rna"], a_ini if comp else None)))
    if X["aportes"] or (comp and a_apo):
        rows.append(("det", "Suscripción de capital / aportes",
                     eep("", X["aportes"], 0, 0, 0, 0, a_apo if comp else None)))
    rows.append(("det", "Ganancia (Pérdida) del ejercicio",
                 eep("", 0, 0, 0, 0, R["res"], A.get("res", 0.0) if comp else None)))
    fin = eep("", I["cap"] + X["aportes"], I["aj"], I["apirr"], I["resleg"], I["rna"] + R["res"],
              A.get("pn", 0.0) if comp else None)
    rows.append(("grand", "Saldos al cierre del ejercicio", fin))
    S.append(dict(key="EEPN", title="ESTADO DE EVOLUCIÓN DEL PATRIMONIO NETO", cols=cols, rows=rows,
                  land=True))

    # ── EFE ──
    AE = X["AE"]
    ee = lambda k, pk=None: [E[k], AE.get(k)] if comp else [E[k]]
    rows = [
        ("sec", "VARIACIÓN NETA DEL EFECTIVO Y SUS EQUIVALENTES", None),
        ("det", "Efectivo y sus equivalentes al inicio del ejercicio", ee("ini", "efe_ini")),
        ("det", "Efectivo y sus equivalentes al cierre del ejercicio", ee("cie", "efe_cie")),
        ("tot", "Aumento (Disminución) neto del efectivo y sus equivalentes",
         ee("var")),
        ("blank", "", None),
        ("sec", "CAUSAS DE LAS VARIACIONES DEL EFECTIVO Y SUS EQUIVALENTES", None),
        ("sub", "Actividades operativas", None),
        ("det", "Ganancia (Pérdida) del ejercicio", ee("res", "res")),
        ("det", "Ajustes para arribar al flujo neto de efectivo de las actividades operativas:", None),
        ("det", "   Depreciación de bienes de uso", ee("amort", "efe_dep")),
        ("det", "Cambios en activos y pasivos operativos:", None),
        ("det", "   (Aumento) Disminución en créditos por ventas", ee("cv", "efe_cv")),
        ("det", "   (Aumento) Disminución en otros créditos", ee("oc", "efe_oc")),
        ("det", "   (Aumento) Disminución en bienes de cambio", ee("bc", "efe_bc")),
        ("det", "   Aumento (Disminución) en deudas operativas", ee("deu", "efe_deu")),
        ("tot", "Flujo neto de efectivo generado por (utilizado en) las actividades operativas", ee("op", "efe_op")),
        ("sub", "Actividades de inversión", None),
        ("det", "Pagos por compras de bienes de uso", ee("bu")),
        ("det", "(Aumento) Disminución de inversiones no equivalentes de efectivo", ee("invs")),
        ("tot", "Flujo neto de efectivo generado por (utilizado en) las actividades de inversión", ee("inv", "efe_inv")),
        ("sub", "Actividades de financiación", None),
        ("det", "Aportes de capital", ee("aportes")),
        ("det", "Aumento (Disminución) de préstamos", ee("prest")),
        ("det", "Aumento (Disminución) de deudas con socios", ee("ds")),
        ("tot", "Flujo neto de efectivo generado por (utilizado en) las actividades de financiación", ee("fin", "efe_fin")),
        ("grand", "Aumento (Disminución) neto del efectivo y sus equivalentes", ee("var")),
    ]
    S.append(dict(key="EFE", title="ESTADO DE FLUJO DE EFECTIVO (método indirecto)", cols=hdr2, rows=rows,
                  land=False))

    # ── Anexo Bienes de uso ──
    if "bu" in anx:
        cols = ["Al inicio", "Altas", "Bajas", "Al cierre", "Acum. al inicio", "Tasa %",
                "Del ejercicio", "Acum. al cierre", "Neto actual"] + (["Neto anterior"] if comp else [])
        rows = []
        tot = [0.0] * len(cols)
        for r in X["bu"]:
            vals = [r["o_ini"], r["altas"], r["bajas"], r["o_cie"], r["a_ini"], None, r["amort"], r["a_cie"],
                    r["neto"]] + ([r["neto_ant"]] if comp else [])
            rows.append(("det", r["tipo"], vals, {5: f"{r['tasa']:g}".replace(".", ",")}))
            tot = [t + (x or 0) for t, x in zip(tot, vals)]
        tot[5] = None
        rows.append(("grand", "TOTALES", tot))
        S.append(dict(key=f"Anexo {anx['bu']}", title=f"ANEXO {anx['bu']} — BIENES DE USO", cols=cols,
                      rows=rows, land=True))

    # ── Anexo Costo ──
    if "costo" in anx:
        rows = [("det", "Existencia al inicio del ejercicio", [R["bc_ini"]]),
                ("det", "Compras del ejercicio", [R["compras"]]),
                ("det", "Existencia al cierre del ejercicio", [-R["bc_cie"]]),
                ("grand", "Costo de los bienes vendidos", [R["costo"]])]
        S.append(dict(key=f"Anexo {anx['costo']}",
                      title=f"ANEXO {anx['costo']} — COSTO DE LOS BIENES VENDIDOS", cols=["Actual"],
                      rows=rows, land=False))

    # ── Anexo Gastos (art. 64 inc. b) LGS) ──
    rows = []
    gastos = {}
    for n, s, k in X["cuentas"]:
        if k in ("gcom", "gadm") and abs(s) > 0.005:
            gastos.setdefault(n, [0.0, 0.0])
            gastos[n][0 if k == "gadm" else 1] += s
    if X["amort_ej"]:
        gastos["Amortización bienes de uso"] = [X["amort_ej"], 0.0]
    if not R["usa_inv"] and R["costo"]:
        costo_rows = [(n, s) for n, s, k in X["cuentas"] if k == "costo" and abs(s) > 0.005]
    else:
        costo_rows = []
    ttot = [0.0, 0.0, 0.0, 0.0]
    for n, s in costo_rows:
        rows.append(("det", n, [s, 0.0, 0.0, s]))
        ttot = [t + x for t, x in zip(ttot, [s, 0.0, 0.0, s])]
    for n, (ga, gc) in gastos.items():
        vals = [0.0, ga, gc, ga + gc]
        rows.append(("det", n, vals))
        ttot = [t + x for t, x in zip(ttot, vals)]
    rows.append(("grand", "TOTALES", ttot))
    S.append(dict(key=f"Anexo {anx['gastos']}",
                  title=f"ANEXO {anx['gastos']} — INFORMACIÓN REQUERIDA POR EL ART. 64 INC. b) LEY 19.550",
                  cols=["Costo de servicios", "Gastos de administración", "Gastos de comercialización",
                        "Total"], rows=rows, land=True))

    notas = armar_notas(a, X, comp_rubros, nota_n)
    return S, notas, anx, nota_n


def armar_notas(a, X, comp_rubros, nota_n):
    C, A, R = X["C"], X["A"], X["R"]
    comp = X["tiene_prior"]
    emp = a.empresa
    N = []  # bloques: ("h1"|"h2"|"p"|"table", contenido)
    N.append(("h1", "1. Notas generales"))
    N.append(("h2", "1.1. Bases de preparación de los estados contables"))
    N.append(("p", "Los presentes estados contables están expresados en pesos en moneda homogénea de cierre "
                   "(Nota 1.3) y han sido preparados de conformidad con la Resolución Técnica N° 54 - T.O. RT 59 "
                   "y modificatorias, emitida por la Federación Argentina de Consejos Profesionales de Ciencias "
                   "Económicas (FACPCE), adoptada por la Resolución P. N° 460/2024 del Consejo Profesional de "
                   "Ciencias Económicas de la Ciudad Autónoma de Buenos Aires (CPCECABA)."))
    ini = X["inicio"].strftime("%d/%m/%Y")
    if a.nro_ejercicio <= 1:
        N.append(("p", f"Por tratarse del primer ejercicio económico de la entidad, iniciado el {ini}, "
                       "no se presenta información comparativa."))
    elif X["primer_rt54"]:
        N.append(("p", f"Esta resolución fue aplicada por la entidad a partir del {ini}, utilizando el enfoque "
                       "retroactivo simplificado, no surgiendo impactos patrimoniales ni en resultados como "
                       "consecuencia de su aplicación. Los estados contables comparativos se exponen reexpresados "
                       "en moneda homogénea al cierre del ejercicio actual conforme a lo establecido en la RT 54."))
    else:
        N.append(("p", "Los estados contables comparativos se exponen reexpresados en moneda homogénea al "
                       "cierre del ejercicio actual conforme a lo establecido en la RT 54."))
    N.append(("h2", "1.2. Clasificación de la entidad"))
    N.append(("p", "De acuerdo con lo establecido por la referida norma, la entidad reviste el carácter de entidad "
                   "pequeña (EP) y no presenta temas de alta complejidad contable en función de la información "
                   "suministrada por la administración."))
    N.append(("h2", "1.3. Unidad de medida"))
    N.append(("p", "Los presentes estados contables han sido preparados en moneda homogénea de cierre, "
                   "reconociendo los efectos de la pérdida del poder adquisitivo de la moneda de acuerdo con las "
                   "normas contables profesionales vigentes. En la preparación se utilizaron las simplificaciones "
                   "admitidas para entidades pequeñas, exponiendo los resultados financieros y por tenencia, "
                   "incluyendo el RECPAM, en una sola línea del estado de resultados."))
    N.append(("h2", "1.4. Uso de estimaciones en la preparación de los presentes estados contables"))
    N.append(("p", f"La preparación de estados contables requiere que el órgano de administración de {emp} "
                   "realice estimaciones y evaluaciones que afectan el monto de los activos y pasivos registrados "
                   "y los activos y pasivos contingentes revelados a la fecha de cierre, como así también los "
                   "ingresos y egresos registrados en el ejercicio. Los resultados reales futuros pueden diferir "
                   "de las estimaciones y evaluaciones realizadas a la fecha de preparación de los presentes "
                   "estados contables."))

    # 2. Políticas
    N.append(("h1", "2. Políticas contables sobre criterios de medición contable"))
    pol = []
    tiene = lambda k: abs(C.get(k, 0)) > 0.005 or abs(A.get(k, 0)) > 0.005
    if X["moneda_ext"]:
        pol.append(("Activos y pasivos en moneda extranjera",
                    "Los activos y pasivos en moneda extranjera han sido medidos al tipo de cambio "
                    "correspondiente a la fecha de cierre de los estados contables o, en su defecto, al tipo de "
                    "cambio implícito utilizado por la entidad financiera para su liquidación o cancelación."))
    pol.append(("Caja y bancos", "El efectivo disponible y los saldos en bancos han sido medidos por su importe "
                                 "nominal."))
    if tiene("inv"):
        pol.append(("Inversiones", "Las inversiones temporarias han sido medidas a su valor neto de realización "
                                   "a la fecha de cierre."))
    pol.append(("Cuentas por cobrar a clientes en moneda, Créditos impositivos, Créditos en moneda con partes "
                "relacionadas, Otras cuentas por cobrar en moneda",
                "Las cuentas por cobrar a clientes en moneda, créditos impositivos y otras cuentas por cobrar "
                "se miden a su valor nominal, netos de previsiones por desvalorización, de corresponder."))
    if tiene("bc"):
        pol.append(("Bienes de cambio",
                    "Los bienes de cambio adquiridos se miden a su costo de la última compra, excepto por los "
                    "productos en proceso o terminados que se miden al costo de producción a la fecha de cierre "
                    "de cada ejercicio.\nEl valor registrado de los bienes de cambio no supera su valor recuperable "
                    "estimado en las condiciones actuales. Todos los bienes se encuentran reexpresados a la fecha "
                    "de cierre de acuerdo a lo indicado en la Nota 1.3."))
    if X["bu"]:
        pol.append(("Bienes de uso",
                    "Los bienes de uso se miden a su costo de adquisición reexpresado de acuerdo con lo indicado "
                    "en la Nota 1.3, neto de las depreciaciones acumuladas. Las depreciaciones se calculan por el "
                    "método de la línea recta, aplicando tasas anuales suficientes para extinguir los valores al "
                    "final de la vida útil estimada de los bienes, computándose el año de alta completo. El valor "
                    "registrado de los bienes de uso no supera su valor recuperable."))
    pol.append(("Proveedores de bienes y servicios, Deudas fiscales, Deudas laborales y previsionales, y Deudas "
                "en moneda con partes relacionadas",
                "Las deudas con proveedores y otros pasivos financieros se miden por su valor nominal" +
                (", excepto por los pasivos en moneda extranjera, que se miden de acuerdo con el criterio "
                 "indicado en la Nota 2.1." if X["moneda_ext"] else ".")))
    if tiene("bc"):
        pol.append(("Costo de los bienes vendidos",
                    "El costo de los bienes vendidos se calcula aplicando la técnica de diferencia de inventario "
                    "considerando la reexpresión de la Nota 1.3."))
    pol.append(("Impuesto a las ganancias",
                "El cargo contable por impuesto a las ganancias se reconoce en función del impuesto determinado, "
                "considerando la normativa fiscal vigente."))
    pol.append(("Cuentas de Patrimonio neto",
                "Se encuentran reexpresadas de acuerdo con lo mencionado en la Nota 1.3. Las cuentas “Capital” y "
                "“Aportes Irrevocables” se mantienen a sus valores nominales y el ajuste derivado de dicha "
                "reexpresión monetaria realizada de acuerdo a la Nota 1.3, se expone en las cuentas de “Ajuste de "
                "capital” y “Ajuste de Aportes Irrevocables”, respectivamente."))
    for i, (t, p) in enumerate(pol):
        N.append(("h2", f"2.{i + 1}. {t}"))
        for par in p.split("\n"):
            N.append(("p", par))

    # 3. Composición
    N.append(("h1", "3. Composición de los principales rubros"))
    titulos = dict(caja="Caja y bancos", inv="Inversiones", cv="Créditos por ventas", oc="Otros créditos",
                   bc="Bienes de cambio", dc="Deudas comerciales", prest="Préstamos",
                   rem="Remuneraciones y cargas sociales", df="Cargas fiscales", ds="Deudas con socios",
                   od="Otras deudas", rfyt="Resultados financieros y por tenencia")
    prior = X["prior"]
    signo = lambda k: -1 if RUBROS.get(k, ("A",))[0] == "P" else 1
    for k in comp_rubros:
        filas = []
        if k == "rfyt":
            filas.append(("RECPAM y otros resultados financieros y por tenencia",
                          R["rfyt"], A.get("rfyt", 0.0) if comp else None))
            tot = (R["rfyt"], A.get("rfyt", 0.0) if comp else None)
        else:
            sum_prev = 0.0
            for n, s, kk in X["cuentas"]:
                if kk != k:
                    continue
                act = s * signo(k)
                ant = None
                if comp:
                    pv = prior.cuenta(n)
                    ant = round(pv * X["cof"], 2) if pv is not None else 0.0
                    sum_prev += ant
                if abs(act) > 0.005 or (ant and abs(ant) > 0.005):
                    filas.append((n, act, ant))
            if comp:
                dif = A.get(k, 0.0) - sum_prev
                if abs(dif) > 1:
                    filas.append(("Saldos del ejercicio anterior en otras cuentas", 0.0, dif))
            tot = (C.get(k, 0.0), A.get(k, 0.0) if comp else None)
        N.append(("h2", f"{nota_n[k]}. {titulos[k]}"))
        N.append(("table", dict(cols=["Actual", "Anterior"] if comp else ["Actual"],
                                rows=[(f, [x, y] if comp else [x]) for f, x, y in filas],
                                total=[tot[0], tot[1]] if comp else [tot[0]])))

    if C["pn"] < 0:
        N.append(("h1", "4. Patrimonio neto negativo"))
        N.append(("p", f"Al {X['cierre'].strftime('%d/%m/%Y')} la sociedad presenta un patrimonio neto negativo de "
                       f"$ {fmt_ar(C['pn'])}. Los socios han manifestado su intención de continuar con la "
                       "actividad y de adoptar las medidas necesarias para recomponer el patrimonio, por lo que los "
                       "presentes estados contables han sido preparados bajo el principio de empresa en "
                       "funcionamiento."))
    return N


# ════════════════════════════════════════════════════════════════════════════
# Salida XLSX
# ════════════════════════════════════════════════════════════════════════════
def escribir_xlsx(path, a, X, S, notas):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    NF = '#,##0.00;(#,##0.00);"-"'
    BOLD = Font(bold=True)
    HDR = PatternFill("solid", fgColor="1F3864")
    TOT = PatternFill("solid", fgColor="D9E1F2")
    GR = PatternFill("solid", fgColor="B4C6E7")
    thin = Side(style="thin", color="999999")

    wb = Workbook()
    ws = wb.active
    ws.title = "Carátula"
    car = caratula_lineas(a, X)
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 80
    for i, (k, val) in enumerate(car, 1):
        ws.cell(i, 1, k).font = BOLD
        ws.cell(i, 2, val)

    for st in S:
        ws = wb.create_sheet(st["key"][:31])
        ncol = len(st["cols"])
        ws.column_dimensions["A"].width = 62
        for j in range(ncol):
            ws.column_dimensions[get_column_letter(j + 2)].width = 17
        ws.cell(1, 1, a.empresa).font = Font(bold=True, size=12)
        ws.cell(2, 1, st["title"]).font = Font(bold=True, size=11)
        ws.cell(3, 1, subtitulo(a, X)).font = Font(italic=True, size=9)
        r = 5
        for j, c in enumerate(st["cols"]):
            cell = ws.cell(r, j + 2, c)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = HDR
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
        ws.cell(r, 1, "").fill = HDR
        ws.row_dimensions[r].height = 30
        for row in st["rows"]:
            r += 1
            kind, lab, vals = row[0], row[1], row[2]
            txt = row[3] if len(row) > 3 else {}
            c = ws.cell(r, 1, lab)
            if kind in ("sec", "sub", "tot", "grand"):
                c.font = BOLD
            if kind == "det":
                c.alignment = Alignment(indent=1)
            for j, x in enumerate(vals or []):
                cell = ws.cell(r, j + 2)
                if j in txt:
                    cell.value = txt[j]
                    cell.alignment = Alignment(horizontal="center")
                elif x is not None:
                    cell.value = round(x, 2)
                    cell.number_format = NF
                if kind in ("tot", "grand"):
                    cell.font = BOLD
                    cell.border = Border(top=thin)
            if kind == "tot":
                for j in range(ncol + 1):
                    ws.cell(r, j + 1).fill = TOT
            if kind == "grand":
                for j in range(ncol + 1):
                    ws.cell(r, j + 1).fill = GR
        r += 2
        ws.cell(r, 1, "Las notas y anexos que se acompañan forman parte integrante de los estados contables.") \
            .font = Font(italic=True, size=8)
        ws.page_setup.orientation = "landscape" if st["land"] else "portrait"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True

    ws = wb.create_sheet("Notas")
    ws.column_dimensions["A"].width = 70
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18
    r = 1
    ws.cell(r, 1, f"{a.empresa} — NOTAS A LOS ESTADOS CONTABLES").font = Font(bold=True, size=12)
    r += 1
    ws.cell(r, 1, subtitulo(a, X)).font = Font(italic=True, size=9)
    for kind, cont in notas:
        r += 1
        if kind == "h1":
            r += 1
            ws.cell(r, 1, cont).font = Font(bold=True, size=11)
        elif kind == "h2":
            ws.cell(r, 1, cont).font = BOLD
        elif kind == "p":
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
            c = ws.cell(r, 1, cont)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = max(15, 15 * (len(cont) // 110 + 1))
        elif kind == "table":
            for j, col in enumerate(cont["cols"]):
                ws.cell(r, j + 2, col).font = BOLD
            for lab, vals in cont["rows"]:
                r += 1
                ws.cell(r, 1, lab).alignment = Alignment(indent=1)
                for j, x in enumerate(vals):
                    if x is not None:
                        ws.cell(r, j + 2, round(x, 2)).number_format = NF
            r += 1
            ws.cell(r, 1, "Total").font = BOLD
            for j, x in enumerate(cont["total"]):
                if x is not None:
                    c = ws.cell(r, j + 2, round(x, 2))
                    c.number_format = NF
                    c.font = BOLD
                    c.border = Border(top=thin)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # Control (no va al PDF)
    ws = wb.create_sheet("Control")
    ws.column_dimensions["A"].width = 50
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 34
    ws.append(["Cuenta del sumas y saldos", "Saldo", "Rubro asignado"])
    for c in ws[1]:
        c.font = BOLD
    for n, s, k in X["cuentas"]:
        ws.append([n, round(s, 2), RUBROS[k][1]])
        ws.cell(ws.max_row, 2).number_format = NF
    ws.append([])
    ws.append(["Coeficiente de reexpresión", X["cof"]])
    ws.append(["Amortizaciones del ejercicio calculadas", round(X["amort_ej"], 2)])
    ws.append(["RFyT incl. RECPAM (por diferencia)", round(X["R"]["rfyt"], 2)])
    ws.append([])
    ws.append(["Advertencias"])
    ws.cell(ws.max_row, 1).font = BOLD
    for w in X["warn"] or ["(ninguna)"]:
        ws.append([w])
    wb.save(path)


def subtitulo(a, X):
    return (f"Ejercicio N° {a.nro_ejercicio} iniciado el {X['inicio'].strftime('%d/%m/%Y')} y "
            f"finalizado el {X['cierre'].strftime('%d/%m/%Y')}" +
            (" — comparativo con el ejercicio anterior" if X["tiene_prior"] else "") +
            " — en moneda homogénea")


def caratula_lineas(a, X):
    return [
        ("Denominación", a.empresa),
        ("CUIT", fmt_cuit(a.cuit)),
        ("Domicilio legal", a.domicilio or ""),
        ("Inscripción IGJ", a.matricula_igj or ""),
        ("Actividad principal", X["actividad"] or ""),
        ("Ejercicio económico N°", str(a.nro_ejercicio)),
        ("Iniciado el", fecha_larga(X["inicio"])),
        ("Finalizado el", fecha_larga(X["cierre"])),
        ("Capital suscripto e integrado", f"$ {fmt_ar(a.cap_nominal or X['I']['cap'] + X['aportes'])}"),
        ("Normas contables", "RT 54 (T.O. RT 59) — FACPCE / RP N° 460/2024 CPCECABA"),
        ("Unidad de medida", f"Moneda homogénea al {X['cierre'].strftime('%d/%m/%Y')}"),
    ]


# ════════════════════════════════════════════════════════════════════════════
# Salida PDF (ReportLab)
# ════════════════════════════════════════════════════════════════════════════
def escribir_pdf(path, a, X, S, notas):
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER
    from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer, Table,
                                    TableStyle, NextPageTemplate, PageBreak, KeepTogether)

    AZUL = colors.HexColor("#1F3864")
    TOTF = colors.HexColor("#E7ECF6")
    GRF = colors.HexColor("#C9D5EC")
    st_t = ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=12, leading=15, alignment=TA_CENTER)
    st_s = ParagraphStyle("s", fontName="Helvetica-Oblique", fontSize=8, leading=10, alignment=TA_CENTER)
    st_h1 = ParagraphStyle("h1", fontName="Helvetica-Bold", fontSize=10.5, leading=14, spaceBefore=8,
                           spaceAfter=3)
    st_h2 = ParagraphStyle("h2", fontName="Helvetica-Bold", fontSize=9, leading=12, spaceBefore=5, spaceAfter=2)
    st_p = ParagraphStyle("p", fontName="Helvetica", fontSize=9, leading=12, alignment=TA_JUSTIFY, spaceAfter=3)
    st_c = ParagraphStyle("c", fontName="Helvetica", fontSize=7.5, leading=9)
    st_cb = ParagraphStyle("cb", parent=st_c, fontName="Helvetica-Bold")
    st_hd = ParagraphStyle("hd", fontName="Helvetica-Bold", fontSize=7, leading=8.5, alignment=TA_CENTER,
                           textColor=colors.white)

    def deco(canv, doc):
        canv.saveState()
        w, h = canv._pagesize
        canv.setFont("Helvetica-Bold", 8)
        canv.drawString(1.5 * cm, h - 1.0 * cm, a.empresa)
        canv.setFont("Helvetica", 7.5)
        canv.drawRightString(w - 1.5 * cm, h - 1.0 * cm, f"CUIT {fmt_cuit(a.cuit)}"
                             + (f" — Inscripción IGJ N° {a.matricula_igj}" if a.matricula_igj else ""))
        canv.setStrokeColor(AZUL)
        canv.line(1.5 * cm, h - 1.2 * cm, w - 1.5 * cm, h - 1.2 * cm)
        canv.setFont("Helvetica", 7)
        canv.drawString(1.5 * cm, 0.9 * cm, "Las notas y anexos que se acompañan forman parte integrante "
                                            "de los estados contables.")
        canv.drawRightString(w - 1.5 * cm, 0.9 * cm, f"Página {doc.page}")
        canv.restoreState()

    pw, ph = A4
    lw, lh = landscape(A4)
    doc = BaseDocTemplate(path, pagesize=A4, title=f"EECC {a.empresa} {X['cierre'].year}",
                          author="Juan Pablo Stivala")
    fr_p = Frame(1.5 * cm, 1.4 * cm, pw - 3 * cm, ph - 3.0 * cm, id="p")
    fr_l = Frame(1.5 * cm, 1.4 * cm, lw - 3 * cm, lh - 3.0 * cm, id="l")
    doc.addPageTemplates([PageTemplate(id="P", frames=[fr_p], pagesize=A4, onPage=deco),
                          PageTemplate(id="L", frames=[fr_l], pagesize=landscape(A4), onPage=deco)])
    E = []

    # Carátula
    E.append(Spacer(1, 3 * cm))
    E.append(Paragraph("DENOMINACIÓN DE LA ENTIDAD", st_s))
    E.append(Paragraph(a.empresa, ParagraphStyle("big", parent=st_t, fontSize=18, leading=24)))
    E.append(Spacer(1, 0.4 * cm))
    E.append(Paragraph(f"ESTADOS CONTABLES CORRESPONDIENTES AL EJERCICIO ECONÓMICO N° {a.nro_ejercicio}", st_t))
    E.append(Paragraph(f"iniciado el {fecha_larga(X['inicio'])} y finalizado el {fecha_larga(X['cierre'])}",
                       ParagraphStyle("s2", parent=st_s, fontSize=10, leading=13)))
    E.append(Paragraph("(Expresados en moneda homogénea — pesos — a la fecha de cierre)", st_s))
    E.append(Spacer(1, 1.2 * cm))
    data = [[Paragraph(k, st_cb), Paragraph(v or "", st_c)] for k, v in caratula_lineas(a, X)]
    t = Table(data, colWidths=[5.5 * cm, 11 * cm])
    t.setStyle(TableStyle([("LINEBELOW", (0, 0), (-1, -1), 0.3, colors.HexColor("#BBBBBB")),
                           ("VALIGN", (0, 0), (-1, -1), "TOP"), ("TOPPADDING", (0, 0), (-1, -1), 4),
                           ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]))
    E.append(t)
    cur = "P"

    for stt in S:
        tpl = "L" if stt["land"] else "P"
        if tpl != cur:
            E.append(NextPageTemplate(tpl))
            cur = tpl
        E.append(PageBreak())
        E.append(Paragraph(stt["title"], st_t))
        E.append(Paragraph(subtitulo(a, X), st_s))
        E.append(Spacer(1, 0.35 * cm))
        width = (lw if stt["land"] else pw) - 3 * cm
        n = len(stt["cols"])
        if stt["land"]:
            numw = min(3.4 * cm, (width - 6.5 * cm) / n)
        else:
            numw = 2.6 * cm if n <= 2 else min(2.6 * cm, (width - 6.0 * cm) / n)
        labw = width - numw * n
        data = [[Paragraph("", st_hd)] + [Paragraph(c, st_hd) for c in stt["cols"]]]
        styles = [("BACKGROUND", (0, 0), (-1, 0), AZUL), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                  ("ALIGN", (1, 1), (-1, -1), "RIGHT"), ("FONTSIZE", (0, 1), (-1, -1), 7.5),
                  ("TOPPADDING", (0, 0), (-1, -1), 1.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5)]
        for i, row in enumerate(stt["rows"], 1):
            kind, lab, vals = row[0], row[1], row[2]
            txt = row[3] if len(row) > 3 else {}
            bold = kind in ("sec", "sub", "tot", "grand")
            ind = "&nbsp;&nbsp;&nbsp;" if kind == "det" else ""
            cells = [Paragraph(ind + lab, st_cb if bold else st_c)]
            for j in range(n):
                x = (vals or [None] * n)[j] if vals and j < len(vals) else None
                s = txt[j] if j in txt else (fmt_ar(x) if x is not None else "")
                cells.append(s)
            data.append(cells)
            if bold:
                styles.append(("FONTNAME", (1, i), (-1, i), "Helvetica-Bold"))
            if kind == "tot":
                styles += [("BACKGROUND", (0, i), (-1, i), TOTF), ("LINEABOVE", (1, i), (-1, i), 0.5, colors.grey)]
            if kind == "grand":
                styles += [("BACKGROUND", (0, i), (-1, i), GRF), ("LINEABOVE", (1, i), (-1, i), 0.8, AZUL)]
            for j in txt:
                styles.append(("ALIGN", (j + 1, i), (j + 1, i), "CENTER"))
        t = Table(data, colWidths=[labw] + [numw] * n, repeatRows=1)
        t.setStyle(TableStyle(styles))
        E.append(t)

    # Notas
    if cur != "P":
        E.append(NextPageTemplate("P"))
    E.append(PageBreak())
    E.append(Paragraph("NOTAS A LOS ESTADOS CONTABLES", st_t))
    E.append(Paragraph(subtitulo(a, X), st_s))
    E.append(Spacer(1, 0.3 * cm))
    for kind, cont in notas:
        if kind == "h1":
            E.append(Paragraph(cont, st_h1))
        elif kind == "h2":
            E.append(Paragraph(cont, st_h2))
        elif kind == "p":
            E.append(Paragraph(cont, st_p))
        elif kind == "table":
            n = len(cont["cols"])
            data = [[""] + [Paragraph(c, ParagraphStyle("nh", parent=st_cb, alignment=2)) for c in cont["cols"]]]
            for lab, vals in cont["rows"]:
                data.append([Paragraph(lab, st_c)] + [fmt_ar(x) if x is not None else "" for x in vals])
            data.append([Paragraph("Total", st_cb)] + [fmt_ar(x) if x is not None else "" for x in cont["total"]])
            t = Table(data, colWidths=[11.5 * cm - (n - 1) * 2.8 * cm + 2.8 * cm] + [2.8 * cm] * n)
            t.setStyle(TableStyle([("ALIGN", (1, 0), (-1, -1), "RIGHT"), ("FONTSIZE", (0, 0), (-1, -1), 7.5),
                                   ("FONTNAME", (1, -1), (-1, -1), "Helvetica-Bold"),
                                   ("LINEABOVE", (1, -1), (-1, -1), 0.5, colors.grey),
                                   ("TOPPADDING", (0, 0), (-1, -1), 1), ("BOTTOMPADDING", (0, 0), (-1, -1), 1)]))
            E.append(KeepTogether([t, Spacer(1, 0.2 * cm)]))
    doc.build(E)


# ════════════════════════════════════════════════════════════════════════════
# Main
# ════════════════════════════════════════════════════════════════════════════
def main(argv=None):
    ap = argparse.ArgumentParser(description="Genera EECC (RT 54) desde sumas y saldos + EECC anterior")
    ap.add_argument("--empresa", required=True)
    ap.add_argument("--cuit", required=True)
    ap.add_argument("--nro-ejercicio", type=int, default=1)
    ap.add_argument("--fecha-cierre", required=True, help="YYYY-MM-DD")
    ap.add_argument("--fecha-inicio", default="", help="YYYY-MM-DD (opcional, primer ejercicio)")
    ap.add_argument("--cof", type=float, required=True)
    ap.add_argument("--cap-nominal", type=float, default=0.0)
    ap.add_argument("--ss-actual", required=True)
    ap.add_argument("--eecc-anterior", default=None)
    ap.add_argument("--domicilio", default="")
    ap.add_argument("--matricula-igj", default="")
    ap.add_argument("--actividad", default="")
    ap.add_argument("--output", required=True, help="ruta del .xlsx")
    ap.add_argument("--pdf", default=None, help="ruta del .pdf (por defecto junto al xlsx)")
    a = ap.parse_args(argv)
    a.empresa = a.empresa.strip()
    try:
        if a.cof <= 0:
            raise EECCError("El coeficiente de reexpresión tiene que ser mayor a cero.")
        if a.nro_ejercicio > 1 and not a.eecc_anterior:
            raise EECCError("Falta el EECC del ejercicio anterior (PDF) para los comparativos.")
        cuentas = leer_ss(a.ss_actual)
        prior = Prior(a.eecc_anterior if a.nro_ejercicio > 1 else None)
        if a.nro_ejercicio > 1 and not prior.text.strip():
            raise EECCError("No pude leer el texto del EECC anterior (¿PDF escaneado?).")
        if a.nro_ejercicio > 1 and prior.v.get("pn") is None:
            raise EECCError("No encontré el Patrimonio Neto en el EECC anterior: revisar que el PDF sea el balance "
                            "del ejercicio anterior de esta sociedad.")
        cuit_pdf = re.sub(r"\D", "", prior.text[:4000])
        X = calcular(a, cuentas, prior)
        S, notas, anx, nota_n = armar(a, X)
        escribir_xlsx(a.output, a, X, S, notas)
        pdf_path = a.pdf or re.sub(r"\.xlsx$", "", a.output) + ".pdf"
        escribir_pdf(pdf_path, a, X, S, notas)
        meta = dict(notas_rango=f"1 a {4 if X['C']['pn'] < 0 else 3}",
                    anexos_rango=("I" if len(anx) == 1 else f"I a {list(anx.values())[-1]}"),
                    advertencias=X["warn"], pdf=pdf_path)
        with open(re.sub(r"\.xlsx$", "", a.output) + ".meta.json", "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
        for w in X["warn"]:
            print("ADVERTENCIA:", w)
        print(f"OK {a.output} {pdf_path}")
        return 0
    except EECCError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
