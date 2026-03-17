import streamlit as st
import io
import re
import copy
from datetime import datetime
import pdfplumber
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.cell.cell import MergedCell

# ─── Config página ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Licitación del Tesoro",
    page_icon="🏦",
    layout="centered"
)

# ─── Estilos ──────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { max-width: 750px; margin: 0 auto; }
    .stButton>button {
        background-color: #1a3a5c;
        color: white;
        font-weight: bold;
        width: 100%;
        padding: 0.6rem;
        font-size: 1.05rem;
        border-radius: 6px;
        border: none;
    }
    .stButton>button:hover { background-color: #e26b0a; }
    .resultado-box {
        background-color: #f0f7f0;
        border-left: 4px solid #28a745;
        padding: 1rem 1.2rem;
        border-radius: 4px;
        margin-top: 1rem;
    }
    .instrumento-item { margin: 0.15rem 0; font-size: 0.9rem; }
    .bloque-header { font-weight: bold; color: #1a3a5c; margin-top: 0.5rem; }
</style>
""", unsafe_allow_html=True)

# ─── Constantes ───────────────────────────────────────────────────────────────
NARANJA_HDR = "FFE26B0A"
BLANCO_TXT  = "FFFFFFFF"
MESES = {
    "ENERO":1,"FEBRERO":2,"MARZO":3,"ABRIL":4,"MAYO":5,"JUNIO":6,
    "JULIO":7,"AGOSTO":8,"SEPTIEMBRE":9,"OCTUBRE":10,"NOVIEMBRE":11,"DICIEMBRE":12
}
MESES_ES = {
    1:"enero",2:"febrero",3:"marzo",4:"abril",5:"mayo",6:"junio",
    7:"julio",8:"agosto",9:"septiembre",10:"octubre",11:"noviembre",12:"diciembre"
}

# ─── Helpers ──────────────────────────────────────────────────────────────────
def parsear_fecha(texto):
    if not texto: return None
    m = re.search(r"(\d{1,2})\s+DE\s+(\w+)\s+(?:DE\s+)?(\d{4})", texto.upper())
    if not m: return None
    mes = MESES.get(m.group(2))
    if not mes: return None
    try: return datetime(int(m.group(3)), mes, int(m.group(1)))
    except: return None

def parsear_ddmmyyyy(texto):
    m = re.search(r"(\d{2})/(\d{2})/(\d{4})", texto)
    if m:
        try: return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except: pass
    return None

def fecha_str(dt):
    if not dt: return ""
    return f"{dt.day} de {MESES_ES[dt.month]} de {dt.year}"

def plazo_str(venc, liq):
    if not venc or not liq: return ""
    dias = (venc - liq).days
    if dias > 365:
        a = dias // 365; m = round((dias % 365) / 30)
        return f"Aprox. {a} año{'s' if a>1 else ''} y {m} meses" if m else f"Aprox. {a} año{'s' if a>1 else ''}"
    return f"{dias} días"

def hdr_style(cell):
    cell.fill = PatternFill("solid", start_color=NARANJA_HDR)
    cell.font = Font(color=BLANCO_TXT, size=10)
    cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")

def safe_set(ws, row, col, value):
    cell = ws.cell(row=row, column=col)
    if not isinstance(cell, MergedCell): cell.value = value

def fecha_ticker(T, ticker):
    idx = T.find(ticker)
    if idx == -1: return None
    frag = T[max(0,idx-400):idx+400]
    m = re.search(r"(\d{2}/\d{2}/\d{4})", frag)
    if m: return parsear_ddmmyyyy(m.group(1))
    m2 = re.search(r"VENCIMIENTO\s+([\d\w\s]+?)\s*\("+ticker, T)
    if m2: return parsear_fecha(m2.group(1))
    return None

def leer_tasas_template(wb):
    tasas = {}
    try:
        ws = wb.active
        for row_hdr in range(1, ws.max_row):
            for col in range(4, 10):
                cell_hdr = ws.cell(row=row_hdr, column=col)
                if cell_hdr.value and isinstance(cell_hdr.value, str):
                    m = re.search(r"\(([A-Z0-9]+)\s*[-–]\s*reapertura\)", cell_hdr.value, re.IGNORECASE)
                    if m:
                        ticker = m.group(1).upper()
                        for row_offset in range(1, 15):
                            label_cell = ws.cell(row=row_hdr+row_offset, column=3)
                            if label_cell.value and "tasa" in str(label_cell.value).lower():
                                tasa_val = ws.cell(row=row_hdr+row_offset, column=col).value
                                if tasa_val and str(tasa_val).strip():
                                    tasas[ticker] = str(tasa_val).strip()
                                break
    except: pass
    return tasas

# ─── Extracción PDF ───────────────────────────────────────────────────────────
def extraer_pdf(pdf_bytes, tasas_template=None):
    texto = ""
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages: texto += (page.extract_text() or "") + "\n"
    T = re.sub(r"\s+", " ", texto.upper()).strip()

    m = re.search(r"(\d{1,2}:\d{2})\s+HORAS.*?(\d{1,2}:\d{2})\s+HORAS", T)
    h_ini = m.group(1) if m else "10:00"; h_fin = m.group(2) if m else "15:00"
    m2 = re.search(r"(?:día|dia)\s+\w+\s+(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})", texto, re.IGNORECASE)
    fecha_lic_orig = m2.group(1).strip() if m2 else ""
    periodo_lic = f"Período de Licitación Pública: desde las {h_ini} hs hasta las {h_fin} hs del  {fecha_lic_orig}"

    m_liq = re.search(r"LIQUIDACI[ÓO]N.*?(\d{1,2}\s+DE\s+\w+\s+DE\s+\d{4})", T)
    fecha_liq = parsear_fecha(m_liq.group(1)) if m_liq else None
    if not fecha_liq:
        m_liq2 = re.search(r"LIQUIDACI[ÓO]N.*?(\d{2}/\d{2}/\d{4})", T)
        if m_liq2: fecha_liq = parsear_ddmmyyyy(m_liq2.group(1))

    segunda_vuelta = ""
    m_sv = re.search(r"segunda vuelta[^\n]*?(\d{1,2}:\d{2}).*?(\d{1,2}:\d{2}).*?(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})", texto, re.IGNORECASE|re.DOTALL)
    if m_sv: segunda_vuelta = f"Segunda Vuelta BONAR 2027: desde las {m_sv.group(1)} hs hasta las {m_sv.group(2)} hs del  {m_sv.group(3).strip()}"

    sus_lelink = '\tPesos al tipo de cambio de referencia publicado por el BCRA  (Comunicación "A" 3500 ) correspondiente al día hábil previo a la fecha de licitación (T-1)'
    tamar_spreads = re.findall(r"TAMAR\s*\+\s*([\d,\.]+%)", T)
    tamar_idx = [0]
    if tasas_template is None: tasas_template = {}

    pesos, cer, usd = [], [], []

    def ya_existe(lista, ticker, tipo, venc):
        for x in lista:
            if ticker and x.get("ticker") == ticker: return True
            if not ticker and x.get("tipo") == tipo and x.get("vencimiento") == venc: return True
        return False

    def inst_pesos_base(tipo, ticker, label, f, tasa, precio, ajuste, parametro):
        return {"tipo":tipo,"ticker":ticker,"label":label,"vencimiento":f,
                "tasa":tasa,"precio":precio,"ajuste":ajuste,"parametro":parametro,
                "moneda_emision":"Pesos","moneda_suscripcion":"Pesos","moneda_pago":"Pesos",
                "amortizacion":"Íntegra al vencimiento",
                "monto_max":"Hasta el monto máximo autorizado por la normativa vigente",
                "ley":"Ley de la REPÚBLICA ARGENTINA"}

    # LECAP nuevas
    for m in re.finditer(r"LETRA DEL TESORO NACIONAL CAPITALIZABLE EN PESOS CON VENCIMIENTO ([\d\w\s]+?)\(NUEVA\)", T):
        f = parsear_fecha(m.group(1))
        if f and not ya_existe(pesos, None, "LECAP", f):
            pesos.append(inst_pesos_base("LECAP", None, "LECAP (nueva)", f,
                "A licitar", "$ 1.000,00 por cada VNO \n$ 1.000", "N/A", "TEM"))

    # LECAP reaperturas
    for m in re.finditer(r"LETRA DEL TESORO NACIONAL CAPITALIZABLE EN PESOS CON VENCIMIENTO [\d\w\s]+?\((\w+)\s*-\s*REAPERTURA\)", T):
        tk = m.group(1); f = fecha_ticker(T, tk)
        if not ya_existe(pesos, tk, "LECAP", f):
            pesos.append(inst_pesos_base("LECAP", tk, f"LECAP ({tk} - reapertura)", f,
                "A licitar", "$ 1.000,00 por cada VNO \n$ 1.000", "N/A", "TEM"))

    # LETAMAR
    for m in re.finditer(r"LETRA DEL TESORO NACIONAL EN PESOS A TASA TAMAR CON VENCIMIENTO ([\d\w\s]+?)\((\w+)\s*-\s*REAPERTURA\)", T):
        f = parsear_fecha(m.group(1)); tk = m.group(2)
        sp = tamar_spreads[tamar_idx[0]] if tamar_idx[0] < len(tamar_spreads) else ""; tamar_idx[0] += 1
        tasa = tasas_template.get(tk) or (f"Tamar + {sp}" if sp else "A licitar")
        if not ya_existe(pesos, tk, "LETAM", f):
            pesos.append(inst_pesos_base("LETAM", tk, f"LETAM ({tk} - reapertura)", f,
                tasa, "A licitar", "N/A", "Precio"))

    # BOTAMAR
    for m in re.finditer(r"BONO DEL TESORO NACIONAL EN PESOS A TASA TAMAR CON VENCIMIENTO ([\d\w\s]+?)\((\w+)\s*[-–]\s*REAPERTURA\)", T):
        f = parsear_fecha(m.group(1)); tk = m.group(2)
        sp = tamar_spreads[tamar_idx[0]] if tamar_idx[0] < len(tamar_spreads) else ""; tamar_idx[0] += 1
        tasa = tasas_template.get(tk) or (f"Tamar + {sp}" if sp else "A licitar")
        if not ya_existe(pesos, tk, "BOTAM", f):
            pesos.append(inst_pesos_base("BOTAM", tk, f"BOTAM ({tk} - reapertura)", f,
                tasa, "A licitar", "N/A", "Precio"))

    # LECER reaperturas
    for m in re.finditer(r"LETRA DEL TESORO NACIONAL EN PESOS AJUSTADA? POR CER A DESCUENTO VENCIMIENTO [\d\w\s]+?\((\w+)\s*-\s*REAPERTURA\)", T):
        tk = m.group(1); f = fecha_ticker(T, tk)
        if not ya_existe(cer, tk, "LECER", f):
            cer.append(inst_pesos_base("LECER", tk, f"LECER ({tk} - reapertura)", f,
                "Cero Cupón", "A licitar", "CER", "Precio"))

    # LECER nuevas
    for m in re.finditer(r"LETRA DEL TESORO NACIONAL EN PESOS AJUSTADA? POR CER A DESCUENTO VENCIMIENTO ([\d\w\s]+?)\(NUEVA\)", T):
        f = parsear_fecha(m.group(1))
        if f and not ya_existe(cer, None, "LECER", f):
            cer.append(inst_pesos_base("LECER", None, "LECER (Nueva)", f,
                "Cero Cupón", "A licitar", "CER", "Precio"))

    # BONCER reaperturas
    for m in re.finditer(r"BONO DEL TESORO NACIONAL EN PESOS CERO CUP[ÓO]N CON AJUSTE POR CER VENCIMIENTO [\d\w\s]+?\((\w+)\s*[-–]\s*REAPERTURA\)", T):
        tk = m.group(1); f = fecha_ticker(T, tk)
        if not ya_existe(cer, tk, "BONCER", f):
            cer.append(inst_pesos_base("BONCER", tk, f"BONCER ({tk} – reapertura)", f,
                "Cero Cupón", "A licitar", "CER", "Precio"))

    # BONCER nuevos
    for m in re.finditer(r"BONO DEL TESORO NACIONAL EN PESOS CERO CUP[ÓO]N CON AJUSTE POR CER VENCIMIENTO ([\d\w\s]+?)\(NUEVA\)", T):
        f = parsear_fecha(m.group(1))
        if f and not ya_existe(cer, None, "BONCER", f):
            cer.append(inst_pesos_base("BONCER", None, "BONCER (nuevo)", f,
                "Cero Cupón", "A licitar", "CER", "Precio"))

    # LELINK reaperturas
    for m in re.finditer(r"LETRA DEL TESORO NACIONAL VINCULADA AL D[ÓO]LAR ESTADOUNIDENSE CERO CUP[ÓO]N CON VENCIMIENTO [\d\w\s]+?\((\w+)\s*-\s*REAPERTURA\)", T):
        tk = m.group(1); f = fecha_ticker(T, tk)
        if not ya_existe(usd, tk, "LELINK", f):
            usd.append({"tipo":"LELINK","ticker":tk,"label":f"LELINK ({tk} - reapertura)","vencimiento":f,
                "tasa":"Cero Cupón","precio":"A licitar","ajuste":None,"parametro":"Precio",
                "moneda_emision":"Dólares Estadounidenses","moneda_suscripcion":sus_lelink,
                "moneda_pago":"Pesos al tipo de cambio aplicable","amortizacion":"Íntegra al vencimiento",
                "monto_max":"Hasta el monto máximo  autorizado por la normativa vigente","ley":"Ley de la REPÚBLICA ARGENTINA"})

    # LELINK nuevas
    for m in re.finditer(r"LETRA DEL TESORO NACIONAL VINCULADA AL D[ÓO]LAR ESTADOUNIDENSE CERO CUP[ÓO]N CON VENCIMIENTO ([\d\w\s]+?)\(NUEVA\)", T):
        f = parsear_fecha(m.group(1))
        if f and not ya_existe(usd, None, "LELINK", f):
            usd.append({"tipo":"LELINK","ticker":None,"label":"LELINK (Nuevo)","vencimiento":f,
                "tasa":"Cero Cupón","precio":"A licitar","ajuste":None,"parametro":"Precio",
                "moneda_emision":"Dólares Estadounidenses","moneda_suscripcion":sus_lelink,
                "moneda_pago":"Pesos al tipo de cambio aplicable","amortizacion":"Íntegra al vencimiento",
                "monto_max":"Hasta el monto máximo  autorizado por la normativa vigente","ley":"Ley de la REPÚBLICA ARGENTINA"})

    # BONAR
    for m in re.finditer(r"BONO DEL TESORO NACIONAL EN DOLARES ESTADOUNIDENSES.*?VENCIMIENTO [\d\w\s]+?\((\w+)\s*-\s*REAPERTURA\)", T):
        tk = m.group(1); f = fecha_ticker(T, tk)
        m_mm = re.search(r"USD\s*(\d+)\s*MILLONES.*?(?:EN\s+)?(?:LA\s+)?PRIMERA VUELTA", T)
        monto = f"USD {m_mm.group(1)} Millones en primera vuelta. (*)" if m_mm else "A determinar"
        m_tna = re.search(r"(\d+%)\s+(?:CON VENCIMIENTO|A PAGAR|TNA)", T)
        tna = m_tna.group(1) if m_tna else "6%"
        m_r = re.search(r"RESCATE ANTICIPADO.*?(\d{1,2}\s+DE\s+\w+\s+DE\s+\d{4})", T)
        fr = fecha_str(parsear_fecha(m_r.group(1))) if m_r else ""
        if not ya_existe(usd, tk, "BONAR", f):
            usd.append({"tipo":"BONAR","ticker":tk,"label":f"BONAR 2027 ({tk} - reapertura)","vencimiento":f,
                "tasa":f"TNA {tna} a pagar mensualmente","precio":"A licitar","ajuste":None,"parametro":"Precio",
                "moneda_emision":"Dólares Estadounidenses","moneda_suscripcion":"En Dólares Estadounidenses",
                "moneda_pago":"En Dólares Estadounidenses","amortizacion":"Íntegra al vencimiento",
                "monto_max":monto,"ley":"Ley de la REPÚBLICA ARGENTINA",
                "opcion_rescate":"Los tenedores de los Bonos podrán ejercer, por única vez, una opción de rescate anticipado, total o parcial, del capital del Bono.",
                "fecha_rescate":fr,
                "pago_intereses":"Los intereses serán pagaderos en Pesos por semestre vencido los días 30 de mayo y 30 de noviembre de cada año hasta la fecha de vencimiento"})

    header = {"periodo_licitacion":periodo_lic,"segunda_vuelta":segunda_vuelta,
              "fecha_liquidacion":fecha_liq,"fecha_liquidacion_str":fecha_str(fecha_liq)}
    return {"pesos_fija_variable":pesos,"cer":cer,"dolares":usd}, header


# ─── Generación Excel ─────────────────────────────────────────────────────────
def generar_excel(instrumentos, header, wb):
    ws = wb.active
    fecha_liq = header["fecha_liquidacion"]
    if fecha_liq: ws.title = fecha_liq.strftime("%d.%m.%Y")

    for mc in list(ws.merged_cells.ranges): ws.unmerge_cells(str(mc))

    for row in ws.iter_rows():
        for cell in row:
            if not isinstance(cell.value, str): continue
            if "Período de Licitación Pública" in cell.value:
                cell.value = header["periodo_licitacion"]
            elif "Segunda Vuelta BONAR" in cell.value:
                cell.value = header["segunda_vuelta"] or ""
            elif "Fecha de Liquidación" in cell.value and fecha_liq:
                cell.value = f"Fecha de Liquidación: {header['fecha_liquidacion_str']} (T+2)"

    _bloque(ws, instrumentos["pesos_fija_variable"], fecha_liq, 20,21,22,23,24,25,26,27,28,29,30,31,32)
    _bloque(ws, instrumentos["cer"], fecha_liq, 36,37,38,39,40,41,42,43,44,45,46,47,48)
    _bloque_usd(ws, instrumentos["dolares"], fecha_liq)

    for bloque, r0, r1 in [(instrumentos["pesos_fija_variable"],20,32),
                            (instrumentos["cer"],36,48),(instrumentos["dolares"],55,66)]:
        for col in range(4+len(bloque), 4+6):
            for row in range(r0, r1+1):
                safe_set(ws, row, col, None)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def _bloque(ws, insts, fecha_liq, R0,R1,R2,R3,R4,R5,R6,R7,R8,R9,R10,R11,R12):
    if not insts: return
    for i, inst in enumerate(insts):
        col = 4 + i
        cell_h = ws.cell(row=R0, column=col)
        cell_h.value = inst["label"]; hdr_style(cell_h)
        cell_v = ws.cell(row=R1, column=col)
        cell_v.value = inst["vencimiento"]
        if inst["vencimiento"]:
            cell_v.number_format = "DD/MM/YYYY"
            cell_v.alignment = Alignment(horizontal="center")
        safe_set(ws, R2, col, plazo_str(inst["vencimiento"], fecha_liq))
        if i == 0: safe_set(ws, R3, col, inst["moneda_emision"])
        safe_set(ws, R6, col, inst["tasa"])
        safe_set(ws, R7, col, inst["precio"])
        safe_set(ws, R8, col, inst.get("ajuste") or "N/A")
        safe_set(ws, R9, col, inst["parametro"])
        if i == 0:
            safe_set(ws, R10, col, inst["amortizacion"])
            safe_set(ws, R11, col, inst["monto_max"])
            safe_set(ws, R12, col, inst["ley"])

def _bloque_usd(ws, insts, fecha_liq):
    if not insts: return
    for inst in insts:
        if inst["tipo"] == "BONAR":
            col = 4 + insts.index(inst)
            safe_set(ws, 51, col, inst.get("opcion_rescate",""))
            safe_set(ws, 52, col, inst.get("fecha_rescate",""))
            safe_set(ws, 53, col, inst.get("pago_intereses",""))
    for i, inst in enumerate(insts):
        col = 4 + i
        cell_h = ws.cell(row=55, column=col)
        cell_h.value = inst["label"]; hdr_style(cell_h)
        cell_v = ws.cell(row=56, column=col)
        cell_v.value = inst["vencimiento"]
        if inst["vencimiento"]:
            cell_v.number_format = "DD/MM/YYYY"
            cell_v.alignment = Alignment(horizontal="center")
        safe_set(ws, 57, col, plazo_str(inst["vencimiento"], fecha_liq))
        safe_set(ws, 58, col, inst["moneda_emision"])
        safe_set(ws, 59, col, inst["moneda_suscripcion"])
        safe_set(ws, 60, col, inst["moneda_pago"])
        safe_set(ws, 61, col, inst["tasa"])
        safe_set(ws, 62, col, inst["precio"])
        safe_set(ws, 63, col, inst["parametro"])
        safe_set(ws, 64, col, inst["amortizacion"])
        safe_set(ws, 65, col, inst["monto_max"])
        safe_set(ws, 66, col, inst["ley"])


# ─── UI ───────────────────────────────────────────────────────────────────────
st.markdown("## 🏦 Licitación del Tesoro")
st.markdown("**Banco Hipotecario · Mercado de Capitales**")
st.markdown("---")

st.markdown("### 1. Subí los archivos")

col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("📊 Excel de la licitación anterior", type=["xlsx"], key="excel")
with col2:
    pdf_file = st.file_uploader("📄 PDF del nuevo llamado", type=["pdf"], key="pdf")

st.markdown("---")
st.markdown("### 2. Generá el Excel")

if excel_file and pdf_file:
    if st.button("⚙️ Generar Excel de la nueva licitación"):
        with st.spinner("Procesando..."):
            try:
                # Leer archivos
                excel_bytes = excel_file.read()
                pdf_bytes = pdf_file.read()

                # Cargar wb template para leer tasas
                wb_template = load_workbook(io.BytesIO(excel_bytes))
                tasas_template = leer_tasas_template(wb_template)

                # Extraer datos del PDF
                instrumentos, header = extraer_pdf(pdf_bytes, tasas_template)

                # Cargar template fresco para generar output
                wb_output = load_workbook(io.BytesIO(excel_bytes))
                excel_output = generar_excel(instrumentos, header, wb_output)

                # Nombre del archivo output
                fecha_liq = header["fecha_liquidacion"]
                nombre_output = f"Licitacion_Tesoro_{fecha_liq.strftime('%d_%m_%Y')}.xlsx" if fecha_liq else "Licitacion_nueva.xlsx"

                # Mostrar resumen
                total = sum(len(v) for v in instrumentos.values())
                st.success(f"✅ Excel generado con {total} instrumentos detectados")

                nombres_bloques = {
                    "pesos_fija_variable": "💵 Pesos tasa fija y variable",
                    "cer": "📈 Ajustados por CER",
                    "dolares": "💲 Dólares"
                }
                for bloque, insts in instrumentos.items():
                    if insts:
                        st.markdown(f"**{nombres_bloques[bloque]}**")
                        for inst in insts:
                            venc = inst["vencimiento"].strftime("%d/%m/%Y") if inst["vencimiento"] else "N/A"
                            st.markdown(f"&nbsp;&nbsp;&nbsp;• `{inst['label']}` — vto. {venc}")

                if header["fecha_liquidacion_str"]:
                    st.markdown(f"📅 **Fecha de liquidación:** {header['fecha_liquidacion_str']}")

                st.markdown("---")
                st.download_button(
                    label="⬇️ Descargar Excel",
                    data=excel_output,
                    file_name=nombre_output,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"❌ Error al procesar los archivos: {str(e)}")
                st.exception(e)
else:
    st.info("⬆️ Subí los dos archivos para habilitar la generación.")

st.markdown("---")
st.caption("Banco Hipotecario · Mercado de Capitales · Emisiones Primarias")
