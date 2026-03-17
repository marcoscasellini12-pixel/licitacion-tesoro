import streamlit as st
import io, re
from datetime import datetime
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Licitación del Tesoro", page_icon="🏦", layout="centered")
st.markdown("""<style>
.stButton>button{background:#1a3a5c;color:white;font-weight:bold;width:100%;padding:.6rem;font-size:1.05rem;border-radius:6px;border:none;}
.stButton>button:hover{background:#e26b0a;}
</style>""", unsafe_allow_html=True)

# ── Colores exactos del template ──────────────────────────────────────────────
NARANJA  = "FFE26B0A"   # header instrumentos
BLANCO   = "FFFFFFFF"   # texto blanco / borde interno
GRIS_OSC = "FFBFBFBF"   # gris oscuro  (tint -0.25)
GRIS_CLA = "FFD8D8D8"   # gris claro   (tint -0.15)
NEGRO    = "FF000000"

# Patrón de colores por fila (mismo orden que template):
# Vencimiento=GRIS_CLA, Plazo=GRIS_OSC, Mon.Em=GRIS_CLA, Mon.Sus=GRIS_CLA, Mon.Pago=GRIS_CLA,
# Tasa=GRIS_OSC, Precio=GRIS_CLA, Ajuste=GRIS_OSC, Param=GRIS_CLA,
# Amort=GRIS_OSC, Monto=GRIS_CLA, Ley=GRIS_OSC
SZ = 11

MESES={"ENERO":1,"FEBRERO":2,"MARZO":3,"ABRIL":4,"MAYO":5,"JUNIO":6,"JULIO":7,"AGOSTO":8,"SEPTIEMBRE":9,"OCTUBRE":10,"NOVIEMBRE":11,"DICIEMBRE":12}
MESES_ES={1:"enero",2:"febrero",3:"marzo",4:"abril",5:"mayo",6:"junio",7:"julio",8:"agosto",9:"septiembre",10:"octubre",11:"noviembre",12:"diciembre"}

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
def fill(rgb): return PatternFill("solid", start_color=rgb)
def fnt(bold=False, color=NEGRO, size=SZ): return Font(bold=bold,size=size,color=color,name="Calibri")
def aln(h="left",v="center",wrap=True): return Alignment(horizontal=h,vertical=v,wrap_text=wrap)

def borde_blanco():
    s=Side(style="thin",color=BLANCO)
    return Border(left=s,right=s,top=s,bottom=s)

def borde_externo():
    """Borde externo negro, interno blanco."""
    bl=Side(style="thin",color=NEGRO)
    bw=Side(style="thin",color=BLANCO)
    return Border(left=bl,right=bl,top=bw,bottom=bw)

def set_fila(ws, r, color_fill, c_inicio, c_fin):
    """Aplica fondo y bordes a toda una fila del rango de datos."""
    f=fill(color_fill)
    for c in range(c_inicio, c_fin+1):
        cell=ws.cell(row=r,column=c)
        cell.fill=f
        # borde blanco entre celdas internas, pero dejamos los bordes extremos
        s_int=Side(style="thin",color=BLANCO)
        s_ext=Side(style="thin",color=NEGRO)
        cell.border=Border(
            left=s_ext if c==c_inicio else s_int,
            right=s_ext if c==c_fin else s_int,
            top=s_int,
            bottom=s_int
        )

def escribir_label(ws, r, color_fill, texto):
    c=ws.cell(row=r,column=3,value=texto)
    c.fill=fill(color_fill); c.font=fnt(size=SZ)
    c.alignment=aln("left","center",wrap=True)
    s_int=Side(style="thin",color=BLANCO)
    s_ext=Side(style="thin",color=NEGRO)
    c.border=Border(left=s_ext,right=s_int,top=s_int,bottom=s_int)

def escribir_valor(ws, r, col, valor, color_fill, centrar=False):
    c=ws.cell(row=r,column=col,value=valor)
    c.fill=fill(color_fill); c.font=fnt(size=SZ)
    c.alignment=aln("center" if centrar else "left","center",wrap=True)

def merge_fila(ws, r, c1, c2, valor, color_fill, centrar=False):
    """Escribe, mergea y aplica estilo a toda la fila c1:c2."""
    set_fila(ws,r,color_fill,c1,c2)
    cell=ws.cell(row=r,column=c1,value=valor)
    cell.font=fnt(size=SZ)
    cell.alignment=aln("center" if centrar else "left","center",wrap=True)
    if c2>c1:
        ws.merge_cells(start_row=r,start_column=c1,end_row=r,end_column=c2)

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

    sv=""
    msv=re.search(r"segunda vuelta[^\n]*?(\d{1,2}:\d{2}).*?(\d{1,2}:\d{2}).*?(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})",txt,re.IGNORECASE|re.DOTALL)
    if msv: sv=f"Segunda Vuelta BONAR 2027: desde las {msv.group(1)} hs hasta las {msv.group(2)} hs del  {msv.group(3).strip()}"

    sus_lelink='Pesos al tipo de cambio de referencia publicado por el BCRA (Comunicación "A" 3500) correspondiente al día hábil previo a la fecha de licitación (T-1)'
    pesos_fija,pesos_var,cer_list,usd_list=[],[],[],[]
    seen=set()
    def add(lista,key,d):
        if key not in seen: seen.add(key); lista.append(d)

    for m in re.finditer(r"LETRA DEL TESORO NACIONAL CAPITALIZABLE EN PESOS CON VENCIMIENTO ([\d\w\s]+?)\(NUEVA\)",T):
        f=pf(m.group(1))
        if f: add(pesos_fija,f"LECAP-nueva-{f}",{"label":"LECAP (nueva)","vencimiento":f,"tasa":"A licitar","precio":"$ 1.000,00 por cada VNO $ 1.000","ajuste":"N/A","parametro":"TEM","moneda":"Pesos","amort":"Íntegra al vencimiento","monto":"Hasta el monto máximo autorizado por la normativa vigente"})
    for m in re.finditer(r"LETRA DEL TESORO NACIONAL CAPITALIZABLE EN PESOS CON VENCIMIENTO [\d\w\s]+?\((\w+)\s*-\s*REAPERTURA\)",T):
        tk=m.group(1); f=ticker_fecha(T,tk)
        add(pesos_fija,f"LECAP-{tk}",{"label":f"LECAP ({tk} - reapertura)","vencimiento":f,"tasa":"A licitar","precio":"$ 1.000,00 por cada VNO $ 1.000","ajuste":"N/A","parametro":"TEM","moneda":"Pesos","amort":"Íntegra al vencimiento","monto":"Hasta el monto máximo autorizado por la normativa vigente"})
    for m in re.finditer(r"LETRA DEL TESORO NACIONAL EN PESOS A TASA TAMAR CON VENCIMIENTO ([\d\w\s]+?)\((\w+)\s*-\s*REAPERTURA\)",T):
        f=pf(m.group(1)); tk=m.group(2)
        add(pesos_var,f"LETAM-{tk}",{"label":f"LETAM ({tk} - reapertura)","vencimiento":f,"tasa":"","precio":"A licitar","ajuste":"N/A","parametro":"Precio","moneda":"Pesos","amort":"Íntegra al vencimiento","monto":"Hasta el monto máximo autorizado por la normativa vigente"})
    for m in re.finditer(r"BONO DEL TESORO NACIONAL EN PESOS A TASA TAMAR CON VENCIMIENTO ([\d\w\s]+?)\((\w+)\s*[-–]\s*REAPERTURA\)",T):
        f=pf(m.group(1)); tk=m.group(2)
        add(pesos_var,f"BOTAM-{tk}",{"label":f"BOTAM ({tk} - reapertura)","vencimiento":f,"tasa":"","precio":"A licitar","ajuste":"N/A","parametro":"Precio","moneda":"Pesos","amort":"Íntegra al vencimiento","monto":"Hasta el monto máximo autorizado por la normativa vigente"})
    for m in re.finditer(r"LETRA DEL TESORO NACIONAL EN PESOS AJUSTADA? POR CER A DESCUENTO VENCIMIENTO [\d\w\s]+?\((\w+)\s*-\s*REAPERTURA\)",T):
        tk=m.group(1); f=ticker_fecha(T,tk)
        add(cer_list,f"LECER-{tk}",{"label":f"LECER ({tk} - reapertura)","vencimiento":f,"tasa":"Cero Cupón","precio":"A licitar","ajuste":"CER","parametro":"Precio","moneda":"Pesos","amort":"Íntegra al vencimiento","monto":"Hasta el monto máximo autorizado por la normativa vigente"})
    for m in re.finditer(r"LETRA DEL TESORO NACIONAL EN PESOS AJUSTADA? POR CER A DESCUENTO VENCIMIENTO ([\d\w\s]+?)\(NUEVA\)",T):
        f=pf(m.group(1))
        if f: add(cer_list,f"LECER-nueva-{f}",{"label":"LECER (Nueva)","vencimiento":f,"tasa":"Cero Cupón","precio":"A licitar","ajuste":"CER","parametro":"Precio","moneda":"Pesos","amort":"Íntegra al vencimiento","monto":"Hasta el monto máximo autorizado por la normativa vigente"})
    for m in re.finditer(r"BONO DEL TESORO NACIONAL EN PESOS CERO CUP[ÓO]N CON AJUSTE POR CER VENCIMIENTO [\d\w\s]+?\((\w+)\s*[-–]\s*REAPERTURA\)",T):
        tk=m.group(1); f=ticker_fecha(T,tk)
        add(cer_list,f"BONCER-{tk}",{"label":f"BONCER ({tk} – reapertura)","vencimiento":f,"tasa":"Cero Cupón","precio":"A licitar","ajuste":"CER","parametro":"Precio","moneda":"Pesos","amort":"Íntegra al vencimiento","monto":"Hasta el monto máximo autorizado por la normativa vigente"})
    for m in re.finditer(r"LETRA DEL TESORO NACIONAL VINCULADA AL D[ÓO]LAR ESTADOUNIDENSE CERO CUP[ÓO]N CON VENCIMIENTO [\d\w\s]+?\((\w+)\s*-\s*REAPERTURA\)",T):
        tk=m.group(1); f=ticker_fecha(T,tk)
        add(usd_list,f"LELINK-{tk}",{"tipo":"LELINK","label":f"LELINK ({tk} - reapertura)","vencimiento":f,"tasa":"Cero Cupón","precio":"A licitar","parametro":"Precio","mon_em":"Dólares Estadounidenses","mon_sus":sus_lelink,"mon_pago":"Pesos al tipo de cambio aplicable","amort":"Íntegra al vencimiento","monto":"Hasta el monto máximo autorizado por la normativa vigente"})
    for m in re.finditer(r"LETRA DEL TESORO NACIONAL VINCULADA AL D[ÓO]LAR ESTADOUNIDENSE CERO CUP[ÓO]N CON VENCIMIENTO ([\d\w\s]+?)\(NUEVA\)",T):
        f=pf(m.group(1))
        if f: add(usd_list,f"LELINK-nueva-{f}",{"tipo":"LELINK","label":"LELINK (Nuevo)","vencimiento":f,"tasa":"Cero Cupón","precio":"A licitar","parametro":"Precio","mon_em":"Dólares Estadounidenses","mon_sus":sus_lelink,"mon_pago":"Pesos al tipo de cambio aplicable","amort":"Íntegra al vencimiento","monto":"Hasta el monto máximo autorizado por la normativa vigente"})
    for m in re.finditer(r"BONO DEL TESORO NACIONAL EN DOLARES ESTADOUNIDENSES.*?VENCIMIENTO [\d\w\s]+?\((\w+)\s*-\s*REAPERTURA\)",T):
        tk=m.group(1); f=ticker_fecha(T,tk)
        mmx=re.search(r"USD\s*(\d+)\s*MILLONES.*?(?:EN\s+)?(?:LA\s+)?PRIMERA VUELTA",T)
        monto_str=f"USD {mmx.group(1)} Millones en primera vuelta. (*)" if mmx else ""
        m_tna=re.search(r"ESTADOUNIDENSES\s+(\d+%)",T); tna=m_tna.group(1) if m_tna else "6%"
        mr=re.search(r"RESCATE ANTICIPADO.*?(\d{1,2}\s+DE\s+\w+\s+DE\s+\d{4})",T)
        fr_str=fstr(pf(mr.group(1))) if mr else ""
        add(usd_list,f"BONAR-{tk}",{"tipo":"BONAR","label":f"BONAR 2027 ({tk} - reapertura)","vencimiento":f,"tasa":f"TNA {tna} a pagar mensualmente","precio":"A licitar","parametro":"Precio","mon_em":"Dólares Estadounidenses","mon_sus":"En Dólares Estadounidenses","mon_pago":"En Dólares Estadounidenses","amort":"Íntegra al vencimiento","monto":monto_str,"opcion":"Los tenedores de los Bonos podrán ejercer, por única vez, una opción de rescate anticipado, total o parcial, del capital del Bono.","fecha_opcion":fr_str,"pago_int":"Los intereses serán pagaderos en Pesos por semestre vencido los días 30 de mayo y 30 de noviembre de cada año hasta la fecha de vencimiento"})

    return {"pesos_fija":pesos_fija,"pesos_var":pesos_var,"cer":cer_list,"usd":usd_list,
            "header":{"h_ini":h_ini,"h_fin":h_fin,"fecha_lic_str":fecha_lic_str,
                      "segunda_vuelta":sv,"fecha_liq":fecha_liq,"fecha_liq_str":fstr(fecha_liq),"fecha_emision":fecha_emision}}

# ── Generación Excel ──────────────────────────────────────────────────────────
def generar_excel(datos):
    wb=Workbook(); ws=wb.active
    hdr=datos["header"]
    fecha_liq=hdr["fecha_liq"]; fecha_em=hdr["fecha_emision"]
    if fecha_liq: ws.title=fecha_liq.strftime("%d.%m.%Y")

    ws.sheet_view.showGridLines=False  # sin grilla

    ws.column_dimensions["A"].width=5
    ws.column_dimensions["B"].width=5
    ws.column_dimensions["C"].width=32
    for i in range(4,16): ws.column_dimensions[get_column_letter(i)].width=27

    COL_L=3; COL_D=4  # col labels=C, datos desde D

    # ── Header con negrita hasta los dos puntos ───────────────────────────────
    def header_bicolor(ws, r, parte_bold, parte_normal, col_ini=3, col_fin=9):
        """Escribe texto en dos partes: bold + normal, usando rich text simulado con dos celdas."""
        # openpyxl no soporta rich text nativo, usamos una celda con el texto completo
        # y ponemos negrita en toda la celda para la parte bold, luego celda contigua para el resto
        ws.row_dimensions[r].height=18
        c1=ws.cell(row=r,column=col_ini,value=parte_bold)
        c1.font=Font(bold=True,size=SZ,name="Calibri")
        c1.alignment=aln("left","center",wrap=False)
        if col_fin>col_ini:
            ws.merge_cells(start_row=r,start_column=col_ini,end_row=r,end_column=col_ini)
        # Celda contigua con texto normal
        c2=ws.cell(row=r,column=col_ini+1,value=parte_normal)
        c2.font=Font(bold=False,size=SZ,name="Calibri")
        c2.alignment=aln("left","center",wrap=False)
        if col_fin>col_ini+1:
            ws.merge_cells(start_row=r,start_column=col_ini+1,end_row=r,end_column=col_fin)

    # Título
    ws.row_dimensions[7].height=18
    c=ws.cell(row=7,column=3,value="LICITACIÓN DEL TESORO")
    c.font=Font(bold=True,size=SZ,name="Calibri"); c.alignment=aln("left",wrap=False)

    ws.row_dimensions[9].height=18
    c=ws.cell(row=9,column=3,value="LICITACION POR EFECTIVO")
    c.font=Font(bold=True,size=SZ,name="Calibri"); c.alignment=aln("left",wrap=False)

    # Período - bold hasta ":" luego normal
    ws.row_dimensions[11].height=18
    header_bicolor(ws,11,"Período de Licitación Pública:",
                   f" desde las {hdr['h_ini']} hs hasta las {hdr['h_fin']} hs del  {hdr['fecha_lic_str']}")

    if hdr["segunda_vuelta"]:
        ws.row_dimensions[12].height=18
        header_bicolor(ws,12,"Segunda Vuelta BONAR 2027:",
                       f" {hdr['segunda_vuelta'].split(':',1)[1] if ':' in hdr['segunda_vuelta'] else hdr['segunda_vuelta']}")

    ws.row_dimensions[13].height=18
    header_bicolor(ws,13,"Fecha de Liquidación:", f" {hdr['fecha_liq_str']} (T+2)")

    fila=[15]
    def R(inc=1): r=fila[0]; fila[0]+=inc; return r

    # ── Bloque de instrumentos ────────────────────────────────────────────────
    def bloque(titulo, insts, es_usd=False):
        if not insts: return
        n=len(insts); c1=COL_D; c2=COL_D+n-1

        # Título sección en itálica
        r=R(); ws.row_dimensions[r].height=21
        ct=ws.cell(row=r,column=COL_L,value=titulo)
        ct.font=Font(italic=True,bold=True,size=SZ,name="Calibri")
        ct.alignment=aln("left",wrap=False)
        R()  # vacía

        # Header naranja
        r=R(); ws.row_dimensions[r].height=42
        # col C vacía con naranja
        cc=ws.cell(row=r,column=COL_L)
        cc.fill=fill(NARANJA)
        s=Side(style="thin",color=BLANCO)
        cc.border=Border(left=Side(style="thin",color=NEGRO),right=s,top=s,bottom=s)
        for i,inst in enumerate(insts):
            c=ws.cell(row=r,column=c1+i,value=inst["label"])
            c.fill=fill(NARANJA); c.font=fnt(color=BLANCO,size=SZ)
            c.alignment=aln("center","center")
            sl=Side(style="thin",color=NEGRO if i==0 else BLANCO)
            sr=Side(style="thin",color=NEGRO if i==n-1 else BLANCO)
            sw=Side(style="thin",color=BLANCO)
            c.border=Border(left=sl,right=sr,top=sw,bottom=sw)

        def fila_individual(label, getter, gris, centrar=False, fmt_fecha=False, altura=21):
            r=R(); ws.row_dimensions[r].height=altura
            escribir_label(ws,r,gris,label)
            set_fila(ws,r,gris,c1,c2)
            for i,inst in enumerate(insts):
                cv=ws.cell(row=r,column=c1+i)
                val=getter(inst)
                if fmt_fecha and val:
                    cv.value=val; cv.number_format="D/M/YYYY"
                    cv.alignment=aln("center","center",wrap=False)
                else:
                    cv.value=val if val is not None else ""
                    cv.alignment=aln("center" if centrar else "left","center")
                cv.font=fnt(size=SZ); cv.fill=fill(gris)

        def fila_merged(label, valor, gris, centrar=False, altura=21):
            r=R(); ws.row_dimensions[r].height=altura
            escribir_label(ws,r,gris,label)
            merge_fila(ws,r,c1,c2,valor,gris,centrar)

        if not es_usd:
            fila_individual("Vencimiento",    lambda x: x["vencimiento"], GRIS_CLA, centrar=True, fmt_fecha=True)
            fila_individual("Plazo",          lambda x: dias_entre(x["vencimiento"],fecha_em), GRIS_OSC)
            fila_merged("Moneda de emision",      "Pesos", GRIS_CLA)
            fila_merged("Moneda de Suscripcion",  "Pesos", GRIS_CLA)
            fila_merged("Moneda de Pago",          "Pesos", GRIS_CLA)
            fila_individual("Tasa de interés ",   lambda x: x.get("tasa",""), GRIS_OSC, centrar=True)
            fila_individual("Precio",             lambda x: x.get("precio",""), GRIS_CLA, altura=42)
            fila_individual("Ajuste de capital",  lambda x: x.get("ajuste",""), GRIS_OSC, centrar=True)
            fila_individual("Párametro a licitar",lambda x: x.get("parametro",""), GRIS_CLA, centrar=True)
            fila_merged("Amortización",           "Íntegra al vencimiento", GRIS_OSC, centrar=True)
            fila_merged("Monto Máximo a Licitar", "Hasta el monto máximo autorizado por la normativa vigente", GRIS_CLA)
            fila_merged("Ley aplicable",          "Ley de la REPÚBLICA ARGENTINA", GRIS_OSC, centrar=True)
        else:
            fila_individual("Vencimiento",        lambda x: x["vencimiento"], GRIS_CLA, centrar=True, fmt_fecha=True)
            fila_individual("Moneda de emision",  lambda x: x.get("mon_em",""), GRIS_OSC)
            fila_individual("Moneda de Suscripcion", lambda x: x.get("mon_sus",""), GRIS_CLA, altura=84)
            fila_individual("Moneda de Pago",     lambda x: x.get("mon_pago",""), GRIS_OSC, altura=42)
            fila_individual("Tasa de interés",    lambda x: x.get("tasa",""), GRIS_CLA, centrar=True, altura=42)
            fila_individual("Precio",             lambda x: x.get("precio",""), GRIS_OSC, centrar=True)
            fila_individual("Párametro a licitar",lambda x: x.get("parametro",""), GRIS_CLA, centrar=True)
            fila_individual("Amortización",       lambda x: x.get("amort",""), GRIS_OSC, centrar=True)
            fila_individual("Monto Máximo a Licitar", lambda x: x.get("monto",""), GRIS_CLA, altura=42)
            fila_merged("Ley aplicable",          "Ley de la REPÚBLICA ARGENTINA", GRIS_OSC, centrar=True)

        R()  # fila vacía

    def bloque_bonar_especial(usd_list):
        bonar=[x for x in usd_list if x.get("tipo")=="BONAR"]
        if not bonar: return
        b=bonar[0]; n=len(usd_list); c1=COL_D; c2=COL_D+n-1

        r=R(); ws.row_dimensions[r].height=21
        ct=ws.cell(row=r,column=COL_L,value="Instrumentros a licitar en dólares")
        ct.font=Font(italic=True,bold=True,size=SZ,name="Calibri")
        ct.alignment=aln("left",wrap=False)

        for label,val,h,gris in [
            ("Opción de rescate anticipado",     b.get("opcion",""),    42, GRIS_CLA),
            ("Fecha de ejercicio de la Opción",  b.get("fecha_opcion",""), 21, GRIS_OSC),
            ("Forma de pago de los servicios\n de interés", b.get("pago_int",""), 63, GRIS_CLA),
        ]:
            r=R(); ws.row_dimensions[r].height=h
            escribir_label(ws,r,gris,label)
            merge_fila(ws,r,c1,c2,val,gris)

        R()

    # ── Escribir todo ─────────────────────────────────────────────────────────
    pesos=datos["pesos_fija"]+datos["pesos_var"]
    if pesos:   bloque("Instrumentos a licitar en pesos a tasa fija y tasa variable:", pesos)
    if datos["cer"]: bloque("Instrumentos a licitar en pesos ajustados por CER:", datos["cer"])
    if datos["usd"]:
        bloque_bonar_especial(datos["usd"])
        bloque("Instrumentros a licitar en dólares", datos["usd"], es_usd=True)
        if any(x.get("tipo")=="BONAR" for x in datos["usd"]):
            r=R(2)
            ws.cell(row=r,column=COL_L,
                value="(*) En segunda vuelta se emitirá un monto tal que en conjunto con la primera vuelta no supere el VNO USD de 250 Millones"
            ).font=Font(size=SZ,name="Calibri")

    out=io.BytesIO(); wb.save(out); out.seek(0)
    return out

# ── UI ────────────────────────────────────────────────────────────────────────
st.markdown("## 🏦 Licitación del Tesoro")
st.markdown("**Banco Hipotecario · Mercado de Capitales**")
st.markdown("---")
st.markdown("### Subí el PDF del nuevo llamado")
pdf_file=st.file_uploader("📄 PDF del llamado",type=["pdf"])
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
                st.download_button("⬇️ Descargar Excel",data=excel_out,file_name=nombre_archivo,
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"❌ Error: {str(e)}"); st.exception(e)
else:
    st.info("⬆️ Subí el PDF para habilitar la generación.")

st.markdown("---")
st.caption("Banco Hipotecario · Mercado de Capitales · Emisiones Primarias")
