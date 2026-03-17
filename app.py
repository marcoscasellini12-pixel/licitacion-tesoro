import streamlit as st
import io, re, copy
from datetime import datetime
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Licitación del Tesoro", page_icon="🏦", layout="centered")
st.markdown("""
<style>
.stButton>button{background:#1a3a5c;color:white;font-weight:bold;width:100%;padding:.6rem;font-size:1.05rem;border-radius:6px;border:none;}
.stButton>button:hover{background:#e26b0a;}
</style>""", unsafe_allow_html=True)

# ── Constantes ────────────────────────────────────────────────────────────────
NARANJA   = "FFE26B0A"
BLANCO    = "FFFFFFFF"
GRIS_FILAS = "FFD9D9D9"   # fondo gris claro filas de label
AZUL_OSCURO = "FF1F3864"

MESES = {"ENERO":1,"FEBRERO":2,"MARZO":3,"ABRIL":4,"MAYO":5,"JUNIO":6,
         "JULIO":7,"AGOSTO":8,"SEPTIEMBRE":9,"OCTUBRE":10,"NOVIEMBRE":11,"DICIEMBRE":12}
MESES_ES = {1:"enero",2:"febrero",3:"marzo",4:"abril",5:"mayo",6:"junio",
            7:"julio",8:"agosto",9:"septiembre",10:"octubre",11:"noviembre",12:"diciembre"}

# ── Helpers fecha ─────────────────────────────────────────────────────────────
def pf(texto):
    if not texto: return None
    m = re.search(r"(\d{1,2})\s+DE\s+(\w+)\s+(?:DE\s+)?(\d{4})", str(texto).upper())
    if not m: return None
    mes = MESES.get(m.group(2))
    if not mes: return None
    try: return datetime(int(m.group(3)), mes, int(m.group(1)))
    except: return None

def pddmm(texto):
    m = re.search(r"(\d{2})/(\d{2})/(\d{4})", str(texto))
    if m:
        try: return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except: pass
    return None

def fstr(dt):
    if not dt: return ""
    return f"{dt.day} de {MESES_ES[dt.month]} de {dt.year}"

def dias_entre(venc, emision):
    if not venc or not emision: return ""
    d = (venc - emision).days
    if d > 365:
        a = d // 365; m = round((d % 365) / 30)
        return f"Aprox. {a} año{'s' if a>1 else ''} y {m} meses" if m else f"Aprox. {a} año{'s' if a>1 else ''}"
    return f"{d} días"

def ticker_fecha(T, ticker):
    idx = T.find(ticker)
    if idx < 0: return None
    frag = T[max(0,idx-500):idx+500]
    m = re.search(r"(\d{2}/\d{2}/\d{4})", frag)
    if m: return pddmm(m.group(1))
    m2 = re.search(r"VENCIMIENTO\s+([\d\w\s]+?)\s*\("+ticker, T)
    if m2: return pf(m2.group(1))
    return None

# ── Extracción PDF ────────────────────────────────────────────────────────────
def extraer_pdf(pdf_bytes):
    txt = ""
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for p in pdf.pages: txt += (p.extract_text() or "") + "\n"
    T = re.sub(r"\s+", " ", txt.upper()).strip()

    # Fecha emisión / liquidación
    m_liq = re.search(r"LIQUIDACI[ÓO]N.*?(\d{1,2}\s+DE\s+\w+\s+DE\s+\d{4})", T)
    fecha_liq = pf(m_liq.group(1)) if m_liq else None
    if not fecha_liq:
        m2 = re.search(r"LIQUIDACI[ÓO]N.*?(\d{2}/\d{2}/\d{4})", T)
        if m2: fecha_liq = pddmm(m2.group(1))
    # Fecha emisión = fecha liquidación (T+2 → emisión = liquidación en este contexto)
    fecha_emision = fecha_liq  # usada para calcular plazo

    # Horario y fecha licitación (para header)
    mh = re.search(r"(\d{1,2}:\d{2})\s+HORAS.*?(\d{1,2}:\d{2})\s+HORAS", T)
    h_ini = mh.group(1) if mh else "10:00"; h_fin = mh.group(2) if mh else "15:00"
    mf = re.search(r"(?:día|dia)\s+\w+\s+(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})", txt, re.IGNORECASE)
    fecha_lic_str = mf.group(1).strip() if mf else ""

    periodo = f"Período de Licitación Pública: desde las {h_ini} hs hasta las {h_fin} hs del  {fecha_lic_str}"

    sv = ""
    msv = re.search(r"segunda vuelta[^\n]*?(\d{1,2}:\d{2}).*?(\d{1,2}:\d{2}).*?(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})", txt, re.IGNORECASE|re.DOTALL)
    if msv: sv = f"Segunda Vuelta BONAR 2027: desde las {msv.group(1)} hs hasta las {msv.group(2)} hs del  {msv.group(3).strip()}"

    sus_lelink = 'Pesos al tipo de cambio de referencia publicado por el BCRA  (Comunicación "A" 3500 ) correspondiente al día hábil previo a la fecha de licitación (T-1)'

    pesos_fija, pesos_var, cer_list, usd_list = [], [], [], []

    seen = set()
    def add(lista, key, d):
        if key not in seen:
            seen.add(key); lista.append(d)

    # ── LECAP nuevas (tasa fija) ──────────────────────────────────────────────
    for m in re.finditer(r"LETRA DEL TESORO NACIONAL CAPITALIZABLE EN PESOS CON VENCIMIENTO ([\d\w\s]+?)\(NUEVA\)", T):
        f = pf(m.group(1))
        if f:
            add(pesos_fija, f"LECAP-nueva-{f}", {
                "label": "LECAP (nueva)", "vencimiento": f,
                "tasa": "A licitar",
                "precio": "$ 1.000,00 por cada VNO $ 1.000",
                "ajuste": "N/A", "parametro": "TEM",
                "moneda": "Pesos", "amort": "Íntegra al vencimiento",
                "monto": "Hasta el monto máximo autorizado por la normativa vigente",
            })

    # ── LECAP reaperturas (tasa fija) ─────────────────────────────────────────
    for m in re.finditer(r"LETRA DEL TESORO NACIONAL CAPITALIZABLE EN PESOS CON VENCIMIENTO [\d\w\s]+?\((\w+)\s*-\s*REAPERTURA\)", T):
        tk = m.group(1); f = ticker_fecha(T, tk)
        add(pesos_fija, f"LECAP-{tk}", {
            "label": f"LECAP ({tk} - reapertura)", "vencimiento": f,
            "tasa": "A licitar",
            "precio": "$ 1.000,00 por cada VNO $ 1.000",
            "ajuste": "N/A", "parametro": "TEM",
            "moneda": "Pesos", "amort": "Íntegra al vencimiento",
            "monto": "Hasta el monto máximo autorizado por la normativa vigente",
        })

    # ── LETAMAR (tasa variable) ───────────────────────────────────────────────
    for m in re.finditer(r"LETRA DEL TESORO NACIONAL EN PESOS A TASA TAMAR CON VENCIMIENTO ([\d\w\s]+?)\((\w+)\s*-\s*REAPERTURA\)", T):
        f = pf(m.group(1)); tk = m.group(2)
        add(pesos_var, f"LETAM-{tk}", {
            "label": f"LETAM ({tk} - reapertura)", "vencimiento": f,
            "tasa": "",   # no está en PDF, se deja en blanco
            "precio": "A licitar", "ajuste": "N/A", "parametro": "Precio",
            "moneda": "Pesos", "amort": "Íntegra al vencimiento",
            "monto": "Hasta el monto máximo autorizado por la normativa vigente",
        })

    # ── BOTAMAR (tasa variable) ───────────────────────────────────────────────
    for m in re.finditer(r"BONO DEL TESORO NACIONAL EN PESOS A TASA TAMAR CON VENCIMIENTO ([\d\w\s]+?)\((\w+)\s*[-–]\s*REAPERTURA\)", T):
        f = pf(m.group(1)); tk = m.group(2)
        add(pesos_var, f"BOTAM-{tk}", {
            "label": f"BOTAM ({tk} - reapertura)", "vencimiento": f,
            "tasa": "",   # no está en PDF
            "precio": "A licitar", "ajuste": "N/A", "parametro": "Precio",
            "moneda": "Pesos", "amort": "Íntegra al vencimiento",
            "monto": "Hasta el monto máximo autorizado por la normativa vigente",
        })

    # ── LECER reaperturas ─────────────────────────────────────────────────────
    for m in re.finditer(r"LETRA DEL TESORO NACIONAL EN PESOS AJUSTADA? POR CER A DESCUENTO VENCIMIENTO [\d\w\s]+?\((\w+)\s*-\s*REAPERTURA\)", T):
        tk = m.group(1); f = ticker_fecha(T, tk)
        add(cer_list, f"LECER-{tk}", {
            "label": f"LECER ({tk} - reapertura)", "vencimiento": f,
            "tasa": "Cero Cupón", "precio": "A licitar",
            "ajuste": "CER", "parametro": "Precio",
            "moneda": "Pesos", "amort": "Íntegra al vencimiento",
            "monto": "Hasta el monto máximo autorizado por la normativa vigente",
        })

    # ── LECER nuevas ──────────────────────────────────────────────────────────
    for m in re.finditer(r"LETRA DEL TESORO NACIONAL EN PESOS AJUSTADA? POR CER A DESCUENTO VENCIMIENTO ([\d\w\s]+?)\(NUEVA\)", T):
        f = pf(m.group(1))
        if f:
            add(cer_list, f"LECER-nueva-{f}", {
                "label": "LECER (Nueva)", "vencimiento": f,
                "tasa": "Cero Cupón", "precio": "A licitar",
                "ajuste": "CER", "parametro": "Precio",
                "moneda": "Pesos", "amort": "Íntegra al vencimiento",
                "monto": "Hasta el monto máximo autorizado por la normativa vigente",
            })

    # ── BONCER reaperturas ────────────────────────────────────────────────────
    for m in re.finditer(r"BONO DEL TESORO NACIONAL EN PESOS CERO CUP[ÓO]N CON AJUSTE POR CER VENCIMIENTO [\d\w\s]+?\((\w+)\s*[-–]\s*REAPERTURA\)", T):
        tk = m.group(1); f = ticker_fecha(T, tk)
        add(cer_list, f"BONCER-{tk}", {
            "label": f"BONCER ({tk} – reapertura)", "vencimiento": f,
            "tasa": "Cero Cupón", "precio": "A licitar",
            "ajuste": "CER", "parametro": "Precio",
            "moneda": "Pesos", "amort": "Íntegra al vencimiento",
            "monto": "Hasta el monto máximo autorizado por la normativa vigente",
        })

    # ── BONCER nuevos ─────────────────────────────────────────────────────────
    for m in re.finditer(r"BONO DEL TESORO NACIONAL EN PESOS CERO CUP[ÓO]N CON AJUSTE POR CER VENCIMIENTO ([\d\w\s]+?)\(NUEVA\)", T):
        f = pf(m.group(1))
        if f:
            add(cer_list, f"BONCER-nueva-{f}", {
                "label": "BONCER (nuevo)", "vencimiento": f,
                "tasa": "Cero Cupón", "precio": "A licitar",
                "ajuste": "CER", "parametro": "Precio",
                "moneda": "Pesos", "amort": "Íntegra al vencimiento",
                "monto": "Hasta el monto máximo autorizado por la normativa vigente",
            })

    # ── LELINK reaperturas ────────────────────────────────────────────────────
    for m in re.finditer(r"LETRA DEL TESORO NACIONAL VINCULADA AL D[ÓO]LAR ESTADOUNIDENSE CERO CUP[ÓO]N CON VENCIMIENTO [\d\w\s]+?\((\w+)\s*-\s*REAPERTURA\)", T):
        tk = m.group(1); f = ticker_fecha(T, tk)
        add(usd_list, f"LELINK-{tk}", {
            "tipo": "LELINK", "label": f"LELINK ({tk} - reapertura)", "vencimiento": f,
            "tasa": "Cero Cupón", "precio": "A licitar", "parametro": "Precio",
            "mon_em": "Dólares Estadounidenses",
            "mon_sus": sus_lelink,
            "mon_pago": "Pesos al tipo de cambio aplicable",
            "amort": "Íntegra al vencimiento",
            "monto": "Hasta el monto máximo  autorizado por la normativa vigente",
        })

    # ── LELINK nuevas ─────────────────────────────────────────────────────────
    for m in re.finditer(r"LETRA DEL TESORO NACIONAL VINCULADA AL D[ÓO]LAR ESTADOUNIDENSE CERO CUP[ÓO]N CON VENCIMIENTO ([\d\w\s]+?)\(NUEVA\)", T):
        f = pf(m.group(1))
        if f:
            add(usd_list, f"LELINK-nueva-{f}", {
                "tipo": "LELINK", "label": "LELINK (Nuevo)", "vencimiento": f,
                "tasa": "Cero Cupón", "precio": "A licitar", "parametro": "Precio",
                "mon_em": "Dólares Estadounidenses",
                "mon_sus": sus_lelink,
                "mon_pago": "Pesos al tipo de cambio aplicable",
                "amort": "Íntegra al vencimiento",
                "monto": "Hasta el monto máximo  autorizado por la normativa vigente",
            })

    # ── BONAR ─────────────────────────────────────────────────────────────────
    for m in re.finditer(r"BONO DEL TESORO NACIONAL EN DOLARES ESTADOUNIDENSES.*?VENCIMIENTO [\d\w\s]+?\((\w+)\s*-\s*REAPERTURA\)", T):
        tk = m.group(1); f = ticker_fecha(T, tk)
        mmx = re.search(r"USD\s*(\d+)\s*MILLONES.*?(?:EN\s+)?(?:LA\s+)?PRIMERA VUELTA", T)
        monto_str = f"USD {mmx.group(1)} Millones en primera vuelta. (*)" if mmx else ""
        m_tna = re.search(r"ESTADOUNIDENSES\s+(\d+%)", T)
        tna = m_tna.group(1) if m_tna else "6%"
        mr = re.search(r"RESCATE ANTICIPADO.*?(\d{1,2}\s+DE\s+\w+\s+DE\s+\d{4})", T)
        fr_str = fstr(pf(mr.group(1))) if mr else ""
        add(usd_list, f"BONAR-{tk}", {
            "tipo": "BONAR", "label": f"BONAR 2027 ({tk} - reapertura)", "vencimiento": f,
            "tasa": f"TNA {tna} a pagar mensualmente", "precio": "A licitar", "parametro": "Precio",
            "mon_em": "Dólares Estadounidenses",
            "mon_sus": "En Dólares Estadounidenses",
            "mon_pago": "En Dólares Estadounidenses",
            "amort": "Íntegra al vencimiento",
            "monto": monto_str,
            "opcion": "Los tenedores de los Bonos podrán ejercer, por única vez, una opción de rescate anticipado, total o parcial, del capital del Bono.",
            "fecha_opcion": fr_str,
            "pago_int": "Los intereses serán pagaderos en Pesos por semestre vencido los días 30 de mayo y 30 de noviembre de cada año hasta la fecha de vencimiento",
        })

    return {
        "pesos_fija": pesos_fija,
        "pesos_var": pesos_var,
        "cer": cer_list,
        "usd": usd_list,
        "header": {
            "periodo": periodo, "segunda_vuelta": sv,
            "fecha_liq": fecha_liq, "fecha_liq_str": fstr(fecha_liq),
            "fecha_emision": fecha_emision,
        }
    }

# ── Estilos Excel ─────────────────────────────────────────────────────────────
def _fnt(bold=False, size=16, color="FF000000", italic=False):
    return Font(bold=bold, size=size, color=color, italic=italic,
                name="Calibri")

def _fill(rgb):
    return PatternFill("solid", start_color=rgb)

def _aln(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def _border_thin():
    s = Side(style="thin", color="FF000000")
    return Border(left=s, right=s, top=s, bottom=s)

def estilo_hdr_inst(cell):
    cell.fill = _fill(NARANJA)
    cell.font = _fnt(color=BLANCO, size=16)
    cell.alignment = _aln("center", "center", wrap=True)
    cell.border = _border_thin()

def estilo_label(cell):
    cell.font = _fnt(size=16)
    cell.alignment = _aln("left", "center", wrap=True)

def estilo_valor(cell, wrap=True):
    cell.font = _fnt(size=16)
    cell.alignment = _aln("left", "center", wrap=wrap)

def estilo_seccion(cell):
    cell.font = _fnt(bold=True, size=16)
    cell.alignment = _aln("left", "center")

# ── Construcción del Excel desde cero ─────────────────────────────────────────
def generar_excel(datos):
    wb = Workbook()
    ws = wb.active
    header = datos["header"]
    fecha_liq = header["fecha_liq"]
    fecha_emision = header["fecha_emision"]

    if fecha_liq:
        ws.title = fecha_liq.strftime("%d.%m.%Y")

    # Anchos de columna
    ws.column_dimensions["A"].width = 17.7
    ws.column_dimensions["B"].width = 5.6
    ws.column_dimensions["C"].width = 31.0
    for col_idx in range(4, 15):
        ws.column_dimensions[get_column_letter(col_idx)].width = 38.5

    # ── Header del cuadro ────────────────────────────────────────────────────
    # R7: título
    c = ws.cell(row=7, column=6, value="LICITACIÓN DEL TESORO")
    c.font = _fnt(bold=True, size=16); c.alignment = _aln("center")

    # R9: subtítulo
    c = ws.cell(row=9, column=3, value="LICITACION POR EFECTIVO")
    c.font = _fnt(bold=True, size=16); c.alignment = _aln("left")

    # R11: período
    c = ws.cell(row=11, column=3, value=header["periodo"])
    c.font = _fnt(bold=True, size=16); c.alignment = _aln("left", wrap=True)

    # R12: segunda vuelta (si existe)
    if header["segunda_vuelta"]:
        c = ws.cell(row=12, column=3, value=header["segunda_vuelta"])
        c.font = _fnt(bold=True, size=16); c.alignment = _aln("left", wrap=True)

    # R13: fecha liquidación
    c = ws.cell(row=13, column=3, value=f"Fecha de Liquidación: {header['fecha_liq_str']} (T+2)")
    c.font = _fnt(size=16); c.alignment = _aln("left")

    fila_actual = [15]  # lista mutable para poder modificar en subfunción

    def next_row(n=1):
        r = fila_actual[0]; fila_actual[0] += n; return r

    def escribir_bloque(titulo, instrumentos, es_usd=False):
        if not instrumentos: return

        n = len(instrumentos)
        col_fin = 3 + n  # col C = labels, D en adelante = instrumentos

        # Fila de título de sección
        r = next_row()
        ws.row_dimensions[r].height = 24.75
        c = ws.cell(row=r, column=3, value=titulo)
        estilo_seccion(c)

        # Fila vacía
        next_row()

        # Fila de headers de instrumentos
        r_hdr = next_row()
        ws.row_dimensions[r_hdr].height = 42.0
        # Celda vacía col C con fondo naranja
        c0 = ws.cell(row=r_hdr, column=3)
        c0.fill = _fill(NARANJA); c0.border = _border_thin()
        for i, inst in enumerate(instrumentos):
            c = ws.cell(row=r_hdr, column=4+i, value=inst["label"])
            estilo_hdr_inst(c)

        # Filas de datos
        FILAS_PESOS = [
            ("Vencimiento",       lambda inst: inst["vencimiento"], "fecha"),
            ("Plazo",             lambda inst: dias_entre(inst["vencimiento"], fecha_emision), "texto"),
            ("Moneda de emision", lambda inst: inst.get("moneda","Pesos") if inst == instrumentos[0] else None, "merged"),
            ("Moneda de Suscripcion", lambda inst: inst.get("moneda","Pesos") if inst == instrumentos[0] else None, "merged"),
            ("Moneda de Pago",    lambda inst: inst.get("moneda","Pesos") if inst == instrumentos[0] else None, "merged"),
            ("Tasa de interés ",  lambda inst: inst.get("tasa",""), "texto"),
            ("Precio",            lambda inst: inst.get("precio",""), "texto"),
            ("Ajuste de capital", lambda inst: inst.get("ajuste",""), "texto"),
            ("Párametro a licitar", lambda inst: inst.get("parametro",""), "texto"),
            ("Amortización",      lambda inst: inst.get("amort","") if inst == instrumentos[0] else None, "merged"),
            ("Monto Máximo a Licitar", lambda inst: inst.get("monto","") if inst == instrumentos[0] else None, "merged"),
            ("Ley aplicable",     lambda inst: "Ley de la REPÚBLICA ARGENTINA" if inst == instrumentos[0] else None, "merged"),
        ]

        FILAS_USD = [
            ("Vencimiento",       lambda inst: inst["vencimiento"], "fecha"),
            ("Plazo",             lambda inst: dias_entre(inst["vencimiento"], fecha_emision), "texto"),
            ("Moneda de emision", lambda inst: inst.get("mon_em",""), "texto"),
            ("Moneda de Suscripcion", lambda inst: inst.get("mon_sus",""), "texto"),
            ("Moneda de Pago",    lambda inst: inst.get("mon_pago",""), "texto"),
            ("Tasa de interés",   lambda inst: inst.get("tasa",""), "texto"),
            ("Precio",            lambda inst: inst.get("precio",""), "texto"),
            ("Párametro a licitar", lambda inst: inst.get("parametro",""), "texto"),
            ("Amortización",      lambda inst: inst.get("amort",""), "texto"),
            ("Monto Máximo a Licitar", lambda inst: inst.get("monto",""), "texto"),
            ("Ley aplicable",     lambda inst: "Ley de la REPÚBLICA ARGENTINA", "texto"),
        ]

        filas_def = FILAS_USD if es_usd else FILAS_PESOS

        for label, fn, tipo in filas_def:
            r = next_row()
            ws.row_dimensions[r].height = 24.75 if label not in ("Moneda de Suscripcion","Amortización","Monto Máximo a Licitar") else 42.0

            # Label col C
            cl = ws.cell(row=r, column=3, value=label)
            estilo_label(cl)
            cl.border = _border_thin()

            for i, inst in enumerate(instrumentos):
                cv = ws.cell(row=r, column=4+i)
                cv.border = _border_thin()
                val = fn(inst)
                if tipo == "fecha" and val:
                    cv.value = val
                    cv.number_format = "DD/MM/YYYY"
                    cv.alignment = _aln("center", "center")
                    cv.font = _fnt(size=16)
                elif tipo == "merged":
                    if i == 0 and val is not None:
                        cv.value = val
                        estilo_valor(cv)
                        # Merge si hay más de 1 instrumento
                        if n > 1:
                            ws.merge_cells(start_row=r, start_column=4,
                                           end_row=r, end_column=3+n)
                    # si i>0, la celda queda vacía (es parte del merge)
                    break  # solo una iteración para merged
                else:
                    if val is not None:
                        cv.value = val
                    estilo_valor(cv)

        # Fila vacía entre secciones
        next_row()

    # ── Sección BONAR especial (filas previas a la tabla) ────────────────────
    def escribir_bonar_especial(usd_list):
        bonar = [x for x in usd_list if x.get("tipo") == "BONAR"]
        if not bonar: return
        b = bonar[0]

        r = next_row()
        ws.row_dimensions[r].height = 21.75
        c = ws.cell(row=r, column=3, value="Instrumentros a licitar en dólares")
        estilo_seccion(c)

        for label, val, h in [
            ("Opción de rescate anticipado", b.get("opcion",""), 62.1),
            ("Fecha de ejercicio de la Opción", b.get("fecha_opcion",""), 21.75),
            ("Forma de pago de los servicios\n de interés", b.get("pago_int",""), 94.5),
        ]:
            r = next_row()
            ws.row_dimensions[r].height = h
            cl = ws.cell(row=r, column=3, value=label)
            estilo_label(cl); cl.border = _border_thin()
            if val:
                cv = ws.cell(row=r, column=4, value=val)
                estilo_valor(cv); cv.border = _border_thin()
                if len(usd_list) > 1:
                    ws.merge_cells(start_row=r, start_column=4,
                                   end_row=r, end_column=3+len(usd_list))

        next_row()  # fila vacía antes de la tabla

    # ── Escribir secciones ────────────────────────────────────────────────────
    pesos_fija = datos["pesos_fija"]
    pesos_var  = datos["pesos_var"]
    pesos_todos = pesos_fija + pesos_var   # juntos en una sección

    if pesos_todos:
        escribir_bloque("Instrumentos a licitar en pesos a tasa fija y tasa variable:", pesos_todos)

    if datos["cer"]:
        escribir_bloque("Instrumentos a licitar en pesos ajustados por CER:", datos["cer"])

    if datos["usd"]:
        escribir_bonar_especial(datos["usd"])
        escribir_bloque("Instrumentros a licitar en dólares", datos["usd"], es_usd=True)

        # Nota (*) si hay BONAR
        if any(x.get("tipo") == "BONAR" for x in datos["usd"]):
            r = next_row(2)
            ws.cell(row=r, column=3,
                    value="(*) En segunda vuelta se emitirá un monto tal que en conjunto con la primera vuelta no supere el VNO USD de 250 Millones"
                    ).font = _fnt(size=16)

    out = io.BytesIO()
    wb.save(out); out.seek(0)
    return out

# ── UI ────────────────────────────────────────────────────────────────────────
st.markdown("## 🏦 Licitación del Tesoro")
st.markdown("**Banco Hipotecario · Mercado de Capitales**")
st.markdown("---")
st.markdown("### Subí los archivos")

col1, col2 = st.columns(2)
with col1:
    pdf_file = st.file_uploader("📄 PDF del nuevo llamado", type=["pdf"])
with col2:
    st.info("ℹ️ Solo necesitás el PDF. El Excel se genera desde cero.")

st.markdown("---")

if pdf_file:
    if st.button("⚙️ Generar Excel"):
        with st.spinner("Procesando PDF..."):
            try:
                datos = extraer_pdf(pdf_file.read())
                excel_out = generar_excel(datos)

                header = datos["header"]
                total = sum(len(datos[k]) for k in ["pesos_fija","pesos_var","cer","usd"])
                st.success(f"✅ {total} instrumentos detectados")

                for bloque, nombre in [("pesos_fija","💵 Tasa fija"),("pesos_var","📊 Tasa variable"),
                                        ("cer","📈 CER"),("usd","💲 Dólares")]:
                    if datos[bloque]:
                        st.markdown(f"**{nombre}**")
                        for inst in datos[bloque]:
                            v = inst["vencimiento"].strftime("%d/%m/%Y") if inst["vencimiento"] else "N/A"
                            st.markdown(f"&nbsp;&nbsp;&nbsp;• `{inst['label']}` — vto. {v}")

                if header["fecha_liq_str"]:
                    st.markdown(f"📅 **Liquidación:** {header['fecha_liq_str']}")

                fecha_liq = header["fecha_liq"]
                nombre_archivo = f"Licitacion_Tesoro_{fecha_liq.strftime('%d_%m_%Y')}.xlsx" if fecha_liq else "Licitacion_nueva.xlsx"

                st.markdown("---")
                st.download_button("⬇️ Descargar Excel", data=excel_out,
                                   file_name=nombre_archivo,
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
                st.exception(e)
else:
    st.info("⬆️ Subí el PDF para habilitar la generación.")

st.markdown("---")
st.caption("Banco Hipotecario · Mercado de Capitales · Emisiones Primarias")
