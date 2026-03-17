import streamlit as st
import io, re
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

NARANJA="FFE26B0A"; BLANCO="FFFFFFFF"
MESES={"ENERO":1,"FEBRERO":2,"MARZO":3,"ABRIL":4,"MAYO":5,"JUNIO":6,"JULIO":7,"AGOSTO":8,"SEPTIEMBRE":9,"OCTUBRE":10,"NOVIEMBRE":11,"DICIEMBRE":12}
MESES_ES={1:"enero",2:"febrero",3:"marzo",4:"abril",5:"mayo",6:"junio",7:"julio",8:"agosto",9:"septiembre",10:"octubre",11:"noviembre",12:"diciembre"}

# ── Fechas ────────────────────────────────────────────────────────────────────
def pf(t):
    if not t: return None
    m=re.search(r"(\d{1,2})\s+DE\s+(\w+)\s+(?:DE\s+)?(\d{4})",str(t).upper())
    if not m: return None
    mes=MESES.get(m.group(2))
    if not mes: return None
    try: return datetime(int(m.group(3)),mes,int(m.group(1)))
    except: return None

def pddmm(t):
    m=re.search(r"(\d{2})/(\d{2})/(\d{4})",str(t))
    if m:
        try: return datetime(int(m.group(3)),int(m.group(2)),int(m.group(1)))
        except: pass
    return None

def fstr(dt):
    if not dt: return ""
    return f"{dt.day} de {MESES_ES[dt.month]} de {dt.year}"

def dias_entre(venc, emision):
    if not venc or not emision: return ""
    d=(venc-emision).days
    if d>365:
        a=d//365; m=round((d%365)/30)
        return f"Aprox. {a} año{'s' if a>1 else ''} y {m} meses" if m else f"Aprox. {a} año{'s' if a>1 else ''}"
    return f"{d} días"

def ticker_fecha(T, ticker):
    idx=T.find(ticker)
    if idx<0: return None
    frag=T[max(0,idx-500):idx+500]
    m=re.search(r"(\d{2}/\d{2}/\d{4})",frag)
    if m: return pddmm(m.group(1))
    m2=re.search(r"VENCIMIENTO\s+([\d\w\s]+?)\s*\("+ticker,T)
    if m2: return pf(m2.group(1))
    return None

# ── Estilos ───────────────────────────────────────────────────────────────────
SZ = 11  # font size uniforme

def thin():
    s=Side(style="thin",color="FF000000")
    return Border(left=s,right=s,top=s,bottom=s)

def set_border_range(ws, r, c1, c2):
    """Pone borde en todas las celdas de un rango (necesario para merged)."""
    for c in range(c1, c2+1):
        ws.cell(row=r, column=c).border = thin()

def fnt(bold=False, color="FF000000"):
    return Font(bold=bold, size=SZ, color=color, name="Calibri")

def aln(h="left", v="center", wrap=True):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def estilo_hdr(cell):
    cell.fill=PatternFill("solid",start_color=NARANJA)
    cell.font=fnt(color=BLANCO)
    cell.alignment=aln("center")
    cell.border=thin()

def estilo_label(cell):
    cell.font=fnt()
    cell.alignment=aln("left")
    cell.border=thin()

def estilo_val(cell, centrar=False):
    cell.font=fnt()
    cell.alignment=aln("center" if centrar else "left")
    cell.border=thin()

def merge_y_escribir(ws, r, c1, c2, valor, centrar=False):
    """Escribe valor, mergea c1:c2 y pone bordes en todo el rango."""
    cell = ws.cell(row=r, column=c1, value=valor)
    cell.font = fnt()
    cell.alignment = aln("center" if centrar else "left")
    if c2 > c1:
        ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    set_border_range(ws, r, c1, c2)

# ── Extracción PDF ────────────────────────────────────────────────────────────
def extraer_pdf(pdf_bytes):
    txt=""
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for p in pdf.pages: txt+=(p.extract_text() or "")+"\n"
    T=re.sub(r"\s+"," ",txt.upper()).strip()

    m_liq=re.search(r"LIQUIDACI[ÓO]N.*?(\d{1,2}\s+DE\s+\w+\s+DE\s+\d{4})",T)
    fecha_liq=pf(m_liq.group(1)) if m_liq else None
    if not fecha_liq:
        m2=re.search(r"LIQUIDACI[ÓO]N.*?(\d{2}/\d{2}/\d{4})",T)
        if m2: fecha_liq=pddmm(m2.group(1))
    fecha_emision=fecha_liq

    mh=re.search(r"(\d{1,2}:\d{2})\s+HORAS.*?(\d{1,2}:\d{2})\s+HORAS",T)
    h_ini=mh.group(1) if mh else "10:00"; h_fin=mh.group(2) if mh else "15:00"
    mf=re.search(r"(?:día|dia)\s+\w+\s+(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})",txt,re.IGNORECASE)
    fecha_lic_str=mf.group(1).strip() if mf else ""
    periodo=f"Período de Licitación Pública: desde las {h_ini} hs hasta las {h_fin} hs del  {fecha_lic_str}"

    sv=""
    msv=re.search(r"segunda vuelta[^\n]*?(\d{1,2}:\d{2}).*?(\d{1,2}:\d{2}).*?(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})",txt,re.IGNORECASE|re.DOTALL)
    if msv: sv=f"Segunda Vuelta BONAR 2027: desde las {msv.group(1)} hs hasta las {msv.group(2)} hs del  {msv.group(3).strip()}"

    sus_lelink='Pesos al tipo de cambio de referencia publicado por el BCRA (Comunicación "A" 3500) correspondiente al día hábil previo a la fecha de licitación (T-1)'

    pesos_fija,pesos_var,cer_list,usd_list=[],[],[],[]
    seen=set()
    def add(lista,key,d):
        if key not in seen: seen.add(key); lista.append(d)

    # LECAP nuevas
    for m in re.finditer(r"LETRA DEL TESORO NACIONAL CAPITALIZABLE EN PESOS CON VENCIMIENTO ([\d\w\s]+?)\(NUEVA\)",T):
        f=pf(m.group(1))
        if f: add(pesos_fija,f"LECAP-nueva-{f}",{"label":"LECAP (nueva)","vencimiento":f,"tasa":"A licitar","precio":"$ 1.000,00 por cada VNO $ 1.000","ajuste":"N/A","parametro":"TEM","moneda":"Pesos","amort":"Íntegra al vencimiento","monto":"Hasta el monto máximo autorizado por la normativa vigente"})

    # LECAP reaperturas
    for m in re.finditer(r"LETRA DEL TESORO NACIONAL CAPITALIZABLE EN PESOS CON VENCIMIENTO [\d\w\s]+?\((\w+)\s*-\s*REAPERTURA\)",T):
        tk=m.group(1); f=ticker_fecha(T,tk)
        add(pesos_fija,f"LECAP-{tk}",{"label":f"LECAP ({tk} - reapertura)","vencimiento":f,"tasa":"A licitar","precio":"$ 1.000,00 por cada VNO $ 1.000","ajuste":"N/A","parametro":"TEM","moneda":"Pesos","amort":"Íntegra al vencimiento","monto":"Hasta el monto máximo autorizado por la normativa vigente"})

    # LETAMAR
    for m in re.finditer(r"LETRA DEL TESORO NACIONAL EN PESOS A TASA TAMAR CON VENCIMIENTO ([\d\w\s]+?)\((\w+)\s*-\s*REAPERTURA\)",T):
        f=pf(m.group(1)); tk=m.group(2)
        add(pesos_var,f"LETAM-{tk}",{"label":f"LETAM ({tk} - reapertura)","vencimiento":f,"tasa":"","precio":"A licitar","ajuste":"N/A","parametro":"Precio","moneda":"Pesos","amort":"Íntegra al vencimiento","monto":"Hasta el monto máximo autorizado por la normativa vigente"})

    # BOTAMAR
    for m in re.finditer(r"BONO DEL TESORO NACIONAL EN PESOS A TASA TAMAR CON VENCIMIENTO ([\d\w\s]+?)\((\w+)\s*[-–]\s*REAPERTURA\)",T):
        f=pf(m.group(1)); tk=m.group(2)
        add(pesos_var,f"BOTAM-{tk}",{"label":f"BOTAM ({tk} - reapertura)","vencimiento":f,"tasa":"","precio":"A licitar","ajuste":"N/A","parametro":"Precio","moneda":"Pesos","amort":"Íntegra al vencimiento","monto":"Hasta el monto máximo autorizado por la normativa vigente"})

    # LECER reaperturas
    for m in re.finditer(r"LETRA DEL TESORO NACIONAL EN PESOS AJUSTADA? POR CER A DESCUENTO VENCIMIENTO [\d\w\s]+?\((\w+)\s*-\s*REAPERTURA\)",T):
        tk=m.group(1); f=ticker_fecha(T,tk)
        add(cer_list,f"LECER-{tk}",{"label":f"LECER ({tk} - reapertura)","vencimiento":f,"tasa":"Cero Cupón","precio":"A licitar","ajuste":"CER","parametro":"Precio","moneda":"Pesos","amort":"Íntegra al vencimiento","monto":"Hasta el monto máximo autorizado por la normativa vigente"})

    # LECER nuevas
    for m in re.finditer(r"LETRA DEL TESORO NACIONAL EN PESOS AJUSTADA? POR CER A DESCUENTO VENCIMIENTO ([\d\w\s]+?)\(NUEVA\)",T):
        f=pf(m.group(1))
        if f: add(cer_list,f"LECER-nueva-{f}",{"label":"LECER (Nueva)","vencimiento":f,"tasa":"Cero Cupón","precio":"A licitar","ajuste":"CER","parametro":"Precio","moneda":"Pesos","amort":"Íntegra al vencimiento","monto":"Hasta el monto máximo autorizado por la normativa vigente"})

    # BONCER reaperturas
    for m in re.finditer(r"BONO DEL TESORO NACIONAL EN PESOS CERO CUP[ÓO]N CON AJUSTE POR CER VENCIMIENTO [\d\w\s]+?\((\w+)\s*[-–]\s*REAPERTURA\)",T):
        tk=m.group(1); f=ticker_fecha(T,tk)
        add(cer_list,f"BONCER-{tk}",{"label":f"BONCER ({tk} – reapertura)","vencimiento":f,"tasa":"Cero Cupón","precio":"A licitar","ajuste":"CER","parametro":"Precio","moneda":"Pesos","amort":"Íntegra al vencimiento","monto":"Hasta el monto máximo autorizado por la normativa vigente"})

    # LELINK reaperturas
    for m in re.finditer(r"LETRA DEL TESORO NACIONAL VINCULADA AL D[ÓO]LAR ESTADOUNIDENSE CERO CUP[ÓO]N CON VENCIMIENTO [\d\w\s]+?\((\w+)\s*-\s*REAPERTURA\)",T):
        tk=m.group(1); f=ticker_fecha(T,tk)
        add(usd_list,f"LELINK-{tk}",{"tipo":"LELINK","label":f"LELINK ({tk} - reapertura)","vencimiento":f,"tasa":"Cero Cupón","precio":"A licitar","parametro":"Precio","mon_em":"Dólares Estadounidenses","mon_sus":sus_lelink,"mon_pago":"Pesos al tipo de cambio aplicable","amort":"Íntegra al vencimiento","monto":"Hasta el monto máximo autorizado por la normativa vigente"})

    # LELINK nuevas
    for m in re.finditer(r"LETRA DEL TESORO NACIONAL VINCULADA AL D[ÓO]LAR ESTADOUNIDENSE CERO CUP[ÓO]N CON VENCIMIENTO ([\d\w\s]+?)\(NUEVA\)",T):
        f=pf(m.group(1))
        if f: add(usd_list,f"LELINK-nueva-{f}",{"tipo":"LELINK","label":"LELINK (Nuevo)","vencimiento":f,"tasa":"Cero Cupón","precio":"A licitar","parametro":"Precio","mon_em":"Dólares Estadounidenses","mon_sus":sus_lelink,"mon_pago":"Pesos al tipo de cambio aplicable","amort":"Íntegra al vencimiento","monto":"Hasta el monto máximo autorizado por la normativa vigente"})

    # BONAR
    for m in re.finditer(r"BONO DEL TESORO NACIONAL EN DOLARES ESTADOUNIDENSES.*?VENCIMIENTO [\d\w\s]+?\((\w+)\s*-\s*REAPERTURA\)",T):
        tk=m.group(1); f=ticker_fecha(T,tk)
        mmx=re.search(r"USD\s*(\d+)\s*MILLONES.*?(?:EN\s+)?(?:LA\s+)?PRIMERA VUELTA",T)
        monto_str=f"USD {mmx.group(1)} Millones en primera vuelta. (*)" if mmx else ""
        m_tna=re.search(r"ESTADOUNIDENSES\s+(\d+%)",T); tna=m_tna.group(1) if m_tna else "6%"
        mr=re.search(r"RESCATE ANTICIPADO.*?(\d{1,2}\s+DE\s+\w+\s+DE\s+\d{4})",T)
        fr_str=fstr(pf(mr.group(1))) if mr else ""
        add(usd_list,f"BONAR-{tk}",{"tipo":"BONAR","label":f"BONAR 2027 ({tk} - reapertura)","vencimiento":f,"tasa":f"TNA {tna} a pagar mensualmente","precio":"A licitar","parametro":"Precio","mon_em":"Dólares Estadounidenses","mon_sus":"En Dólares Estadounidenses","mon_pago":"En Dólares Estadounidenses","amort":"Íntegra al vencimiento","monto":monto_str,"opcion":"Los tenedores de los Bonos podrán ejercer, por única vez, una opción de rescate anticipado, total o parcial, del capital del Bono.","fecha_opcion":fr_str,"pago_int":"Los intereses serán pagaderos en Pesos por semestre vencido los días 30 de mayo y 30 de noviembre de cada año hasta la fecha de vencimiento"})

    return {"pesos_fija":pesos_fija,"pesos_var":pesos_var,"cer":cer_list,"usd":usd_list,
            "header":{"periodo":periodo,"segunda_vuelta":sv,"fecha_liq":fecha_liq,"fecha_liq_str":fstr(fecha_liq),"fecha_emision":fecha_emision}}

# ── Generación Excel ──────────────────────────────────────────────────────────
def generar_excel(datos):
    wb=Workbook(); ws=wb.active
    header=datos["header"]
    fecha_liq=header["fecha_liq"]
    fecha_em=header["fecha_emision"]
    if fecha_liq: ws.title=fecha_liq.strftime("%d.%m.%Y")

    # Anchos
    ws.column_dimensions["A"].width=5
    ws.column_dimensions["B"].width=5
    ws.column_dimensions["C"].width=32
    for i in range(4,16):
        ws.column_dimensions[get_column_letter(i)].width=28

    COL_LABEL=3   # col C = labels
    COL_D=4       # primera col de datos

    # ── Header ────────────────────────────────────────────────────────────────
    ws.row_dimensions[7].height=21
    c=ws.cell(row=7,column=6,value="LICITACIÓN DEL TESORO")
    c.font=Font(bold=True,size=SZ,name="Calibri"); c.alignment=aln("center",wrap=False)

    ws.row_dimensions[9].height=21
    c=ws.cell(row=9,column=COL_LABEL,value="LICITACION POR EFECTIVO")
    c.font=Font(bold=True,size=SZ,name="Calibri"); c.alignment=aln("left",wrap=False)

    ws.row_dimensions[11].height=42
    c=ws.cell(row=11,column=COL_LABEL,value=header["periodo"])
    c.font=Font(bold=True,size=SZ,name="Calibri"); c.alignment=aln("left")

    if header["segunda_vuelta"]:
        ws.row_dimensions[12].height=21
        c=ws.cell(row=12,column=COL_LABEL,value=header["segunda_vuelta"])
        c.font=Font(bold=True,size=SZ,name="Calibri"); c.alignment=aln("left")

    ws.row_dimensions[13].height=21
    c=ws.cell(row=13,column=COL_LABEL,value=f"Fecha de Liquidación: {header['fecha_liq_str']} (T+2)")
    c.font=Font(size=SZ,name="Calibri"); c.alignment=aln("left",wrap=False)

    fila=[15]
    def R(inc=1):
        r=fila[0]; fila[0]+=inc; return r

    # ── Función principal para escribir un bloque ─────────────────────────────
    def bloque(titulo, insts, es_usd=False):
        if not insts: return
        n=len(insts)
        c1=COL_D          # primera col datos
        c2=COL_D+n-1      # última col datos

        # Título sección
        r=R()
        ws.row_dimensions[r].height=21
        cell=ws.cell(row=r,column=COL_LABEL,value=titulo)
        cell.font=Font(bold=True,size=SZ,name="Calibri"); cell.alignment=aln("left",wrap=False)

        R()  # fila vacía

        # Header instrumentos
        r=R()
        ws.row_dimensions[r].height=42
        # celda vacía col C con naranja
        cc=ws.cell(row=r,column=COL_LABEL); cc.fill=PatternFill("solid",start_color=NARANJA); cc.border=thin()
        for i,inst in enumerate(insts):
            c=ws.cell(row=r,column=c1+i,value=inst["label"])
            estilo_hdr(c)

        # ── Filas según tipo ──────────────────────────────────────────────────
        if not es_usd:
            # VENCIMIENTO - valor por columna, centrado
            r=R(); ws.row_dimensions[r].height=21
            estilo_label(ws.cell(row=r,column=COL_LABEL,value="Vencimiento"))
            for i,inst in enumerate(insts):
                cv=ws.cell(row=r,column=c1+i,value=inst["vencimiento"])
                cv.number_format="DD/MM/YYYY"; estilo_val(cv,centrar=True)

            # PLAZO - valor por columna
            r=R(); ws.row_dimensions[r].height=21
            estilo_label(ws.cell(row=r,column=COL_LABEL,value="Plazo"))
            for i,inst in enumerate(insts):
                estilo_val(ws.cell(row=r,column=c1+i,value=dias_entre(inst["vencimiento"],fecha_em)))

            # MONEDA EMISION - combinada
            r=R(); ws.row_dimensions[r].height=21
            estilo_label(ws.cell(row=r,column=COL_LABEL,value="Moneda de emision"))
            merge_y_escribir(ws,r,c1,c2,"Pesos")

            # MONEDA SUSCRIPCION - combinada
            r=R(); ws.row_dimensions[r].height=21
            estilo_label(ws.cell(row=r,column=COL_LABEL,value="Moneda de Suscripcion"))
            merge_y_escribir(ws,r,c1,c2,"Pesos")

            # MONEDA PAGO - combinada
            r=R(); ws.row_dimensions[r].height=21
            estilo_label(ws.cell(row=r,column=COL_LABEL,value="Moneda de Pago"))
            merge_y_escribir(ws,r,c1,c2,"Pesos")

            # TASA - valor por columna
            r=R(); ws.row_dimensions[r].height=21
            estilo_label(ws.cell(row=r,column=COL_LABEL,value="Tasa de interés "))
            for i,inst in enumerate(insts):
                estilo_val(ws.cell(row=r,column=c1+i,value=inst.get("tasa","")))

            # PRECIO - valor por columna
            r=R(); ws.row_dimensions[r].height=42
            estilo_label(ws.cell(row=r,column=COL_LABEL,value="Precio"))
            for i,inst in enumerate(insts):
                estilo_val(ws.cell(row=r,column=c1+i,value=inst.get("precio","")))

            # AJUSTE - valor por columna
            r=R(); ws.row_dimensions[r].height=21
            estilo_label(ws.cell(row=r,column=COL_LABEL,value="Ajuste de capital"))
            for i,inst in enumerate(insts):
                estilo_val(ws.cell(row=r,column=c1+i,value=inst.get("ajuste","")))

            # PARAMETRO - valor por columna
            r=R(); ws.row_dimensions[r].height=21
            estilo_label(ws.cell(row=r,column=COL_LABEL,value="Párametro a licitar"))
            for i,inst in enumerate(insts):
                estilo_val(ws.cell(row=r,column=c1+i,value=inst.get("parametro","")))

            # AMORTIZACION - combinada
            r=R(); ws.row_dimensions[r].height=21
            estilo_label(ws.cell(row=r,column=COL_LABEL,value="Amortización"))
            merge_y_escribir(ws,r,c1,c2,"Íntegra al vencimiento")

            # MONTO - combinado
            r=R(); ws.row_dimensions[r].height=21
            estilo_label(ws.cell(row=r,column=COL_LABEL,value="Monto Máximo a Licitar"))
            merge_y_escribir(ws,r,c1,c2,"Hasta el monto máximo autorizado por la normativa vigente")

            # LEY - combinada
            r=R(); ws.row_dimensions[r].height=21
            estilo_label(ws.cell(row=r,column=COL_LABEL,value="Ley aplicable"))
            merge_y_escribir(ws,r,c1,c2,"Ley de la REPÚBLICA ARGENTINA")

        else:  # USD - cada celda tiene su propio valor
            # VENCIMIENTO
            r=R(); ws.row_dimensions[r].height=21
            estilo_label(ws.cell(row=r,column=COL_LABEL,value="Vencimiento"))
            for i,inst in enumerate(insts):
                cv=ws.cell(row=r,column=c1+i,value=inst["vencimiento"])
                cv.number_format="DD/MM/YYYY"; estilo_val(cv,centrar=True)

            # PLAZO
            r=R(); ws.row_dimensions[r].height=21
            estilo_label(ws.cell(row=r,column=COL_LABEL,value="Plazo"))
            for i,inst in enumerate(insts):
                estilo_val(ws.cell(row=r,column=c1+i,value=dias_entre(inst["vencimiento"],fecha_em)))

            # MONEDA EMISION - cada col su valor
            r=R(); ws.row_dimensions[r].height=21
            estilo_label(ws.cell(row=r,column=COL_LABEL,value="Moneda de emision"))
            for i,inst in enumerate(insts):
                estilo_val(ws.cell(row=r,column=c1+i,value=inst.get("mon_em","")))

            # MONEDA SUSCRIPCION - cada col su valor (texto largo)
            r=R(); ws.row_dimensions[r].height=84
            estilo_label(ws.cell(row=r,column=COL_LABEL,value="Moneda de Suscripcion"))
            for i,inst in enumerate(insts):
                estilo_val(ws.cell(row=r,column=c1+i,value=inst.get("mon_sus","")))

            # MONEDA PAGO
            r=R(); ws.row_dimensions[r].height=42
            estilo_label(ws.cell(row=r,column=COL_LABEL,value="Moneda de Pago"))
            for i,inst in enumerate(insts):
                estilo_val(ws.cell(row=r,column=c1+i,value=inst.get("mon_pago","")))

            # TASA
            r=R(); ws.row_dimensions[r].height=42
            estilo_label(ws.cell(row=r,column=COL_LABEL,value="Tasa de interés"))
            for i,inst in enumerate(insts):
                estilo_val(ws.cell(row=r,column=c1+i,value=inst.get("tasa","")))

            # PRECIO
            r=R(); ws.row_dimensions[r].height=21
            estilo_label(ws.cell(row=r,column=COL_LABEL,value="Precio"))
            for i,inst in enumerate(insts):
                estilo_val(ws.cell(row=r,column=c1+i,value=inst.get("precio","")))

            # PARAMETRO
            r=R(); ws.row_dimensions[r].height=21
            estilo_label(ws.cell(row=r,column=COL_LABEL,value="Párametro a licitar"))
            for i,inst in enumerate(insts):
                estilo_val(ws.cell(row=r,column=c1+i,value=inst.get("parametro","")))

            # AMORTIZACION
            r=R(); ws.row_dimensions[r].height=21
            estilo_label(ws.cell(row=r,column=COL_LABEL,value="Amortización"))
            for i,inst in enumerate(insts):
                estilo_val(ws.cell(row=r,column=c1+i,value=inst.get("amort","")))

            # MONTO
            r=R(); ws.row_dimensions[r].height=42
            estilo_label(ws.cell(row=r,column=COL_LABEL,value="Monto Máximo a Licitar"))
            for i,inst in enumerate(insts):
                estilo_val(ws.cell(row=r,column=c1+i,value=inst.get("monto","")))

            # LEY - combinada (siempre igual para todos)
            r=R(); ws.row_dimensions[r].height=21
            estilo_label(ws.cell(row=r,column=COL_LABEL,value="Ley aplicable"))
            merge_y_escribir(ws,r,c1,c2,"Ley de la REPÚBLICA ARGENTINA")

        R()  # fila vacía al final

    # ── Sección especial BONAR (antes de tabla USD) ───────────────────────────
    def bloque_bonar_especial(usd_list):
        bonar=[x for x in usd_list if x.get("tipo")=="BONAR"]
        if not bonar: return
        b=bonar[0]; n=len(usd_list); c1=COL_D; c2=COL_D+n-1

        r=R(); ws.row_dimensions[r].height=21
        cell=ws.cell(row=r,column=COL_LABEL,value="Instrumentros a licitar en dólares")
        cell.font=Font(bold=True,size=SZ,name="Calibri"); cell.alignment=aln("left",wrap=False)

        for label,val,h in [
            ("Opción de rescate anticipado", b.get("opcion",""), 42),
            ("Fecha de ejercicio de la Opción", b.get("fecha_opcion",""), 21),
            ("Forma de pago de los servicios\n de interés", b.get("pago_int",""), 63),
        ]:
            r=R(); ws.row_dimensions[r].height=h
            estilo_label(ws.cell(row=r,column=COL_LABEL,value=label))
            merge_y_escribir(ws,r,c1,c2,val)

        R()  # fila vacía

    # ── Escribir secciones ────────────────────────────────────────────────────
    pesos=datos["pesos_fija"]+datos["pesos_var"]
    if pesos:
        bloque("Instrumentos a licitar en pesos a tasa fija y tasa variable:", pesos)
    if datos["cer"]:
        bloque("Instrumentos a licitar en pesos ajustados por CER:", datos["cer"])
    if datos["usd"]:
        bloque_bonar_especial(datos["usd"])
        bloque("Instrumentros a licitar en dólares", datos["usd"], es_usd=True)
        if any(x.get("tipo")=="BONAR" for x in datos["usd"]):
            r=R(2); ws.cell(row=r,column=COL_LABEL,
                value="(*) En segunda vuelta se emitirá un monto tal que en conjunto con la primera vuelta no supere el VNO USD de 250 Millones"
                ).font=Font(size=SZ,name="Calibri")

    out=io.BytesIO(); wb.save(out); out.seek(0)
    return out

# ── UI ────────────────────────────────────────────────────────────────────────
st.markdown("## 🏦 Licitación del Tesoro")
st.markdown("**Banco Hipotecario · Mercado de Capitales**")
st.markdown("---")
st.markdown("### Subí el PDF del nuevo llamado")

pdf_file=st.file_uploader("📄 PDF del llamado", type=["pdf"])
st.markdown("---")

if pdf_file:
    if st.button("⚙️ Generar Excel"):
        with st.spinner("Procesando..."):
            try:
                datos=extraer_pdf(pdf_file.read())
                excel_out=generar_excel(datos)
                total=sum(len(datos[k]) for k in ["pesos_fija","pesos_var","cer","usd"])
                st.success(f"✅ {total} instrumentos detectados")
                for bloque,nombre in [("pesos_fija","💵 Tasa fija"),("pesos_var","📊 Tasa variable"),("cer","📈 CER"),("usd","💲 Dólares")]:
                    if datos[bloque]:
                        st.markdown(f"**{nombre}**")
                        for inst in datos[bloque]:
                            v=inst["vencimiento"].strftime("%d/%m/%Y") if inst["vencimiento"] else "N/A"
                            st.markdown(f"&nbsp;&nbsp;&nbsp;• `{inst['label']}` — vto. {v}")
                if datos["header"]["fecha_liq_str"]:
                    st.markdown(f"📅 **Liquidación:** {datos['header']['fecha_liq_str']}")
                fecha_liq=datos["header"]["fecha_liq"]
                nombre_archivo=f"Licitacion_Tesoro_{fecha_liq.strftime('%d_%m_%Y')}.xlsx" if fecha_liq else "Licitacion_nueva.xlsx"
                st.markdown("---")
                st.download_button("⬇️ Descargar Excel",data=excel_out,file_name=nombre_archivo,mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"❌ Error: {str(e)}"); st.exception(e)
else:
    st.info("⬆️ Subí el PDF para habilitar la generación.")

st.markdown("---")
st.caption("Banco Hipotecario · Mercado de Capitales · Emisiones Primarias")
