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

NARANJA="FFE26B0A"; BLANCO="FFFFFFFF"; NEGRO="FF000000"
GRIS_OSC="FFBFBFBF"; GRIS_CLA="FFD8D8D8"
SZ=11

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

def dias_entre(venc,emision):
    if not venc or not emision: return ""
    d=(venc-emision).days
    if d>365:
        a=d//365; m=round((d%365)/30)
        return f"Aprox. {a} año{'s' if a>1 else ''} y {m} meses" if m else f"Aprox. {a} año{'s' if a>1 else ''}"
    return f"{d} días"

def ticker_fecha(T,ticker):
    idx=T.find(ticker)
    if idx<0: return None
    frag=T[max(0,idx-500):idx+500]
    m=re.search(r"(\d{2}/\d{2}/\d{4})",frag)
    if m: return pddmm(m.group(1))
    m2=re.search(r"VENCIMIENTO\s+([\d\w\s]+?)\s*\("+ticker,T)
    if m2: return pf(m2.group(1))
    return None

# ── Estilos ───────────────────────────────────────────────────────────────────
def F(rgb): return PatternFill("solid",start_color=rgb)
def fn(bold=False,color=NEGRO): return Font(bold=bold,size=SZ,color=color,name="Calibri")
def al(h="center",v="center",wrap=True): return Alignment(horizontal=h,vertical=v,wrap_text=wrap)
def B(): # todos los bordes blancos medium
    s=Side(style="medium",color=BLANCO)
    return Border(left=s,right=s,top=s,bottom=s)

def aplicar(cell,rgb,valor=None,bold=False,color=NEGRO,h="center",wrap=True,fmt=None):
    if valor is not None: cell.value=valor
    cell.fill=F(rgb); cell.font=fn(bold=bold,color=color)
    cell.alignment=al(h=h,wrap=wrap); cell.border=B()
    if fmt: cell.number_format=fmt

def aplicar_rango(ws,r,c1,c2,rgb):
    for c in range(c1,c2+1): aplicar(ws.cell(row=r,column=c),rgb)

def merge_escribir(ws,r,c1,c2,rgb,valor,h="center"):
    aplicar_rango(ws,r,c1,c2,rgb)
    cell=ws.cell(row=r,column=c1,value=valor)
    cell.font=fn(); cell.alignment=al(h=h); cell.border=B()
    if c2>c1: ws.merge_cells(start_row=r,start_column=c1,end_row=r,end_column=c2)

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
    fecha_em=fecha_liq

    mh=re.search(r"(\d{1,2}:\d{2})\s+HORAS.*?(\d{1,2}:\d{2})\s+HORAS",T)
    h_ini=mh.group(1) if mh else "10:00"; h_fin=mh.group(2) if mh else "15:00"
    mf=re.search(r"(?:día|dia)\s+\w+\s+(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})",txt,re.IGNORECASE)
    fecha_lic_str=mf.group(1).strip() if mf else ""
    sv=""
    msv=re.search(r"segunda vuelta[^\n]*?(\d{1,2}:\d{2}).*?(\d{1,2}:\d{2}).*?(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})",txt,re.IGNORECASE|re.DOTALL)
    if msv: sv=f"Segunda Vuelta BONAR 2027: desde las {msv.group(1)} hs hasta las {msv.group(2)} hs del  {msv.group(3).strip()}"

    sus_ll='Pesos al tipo de cambio de referencia publicado por el BCRA (Comunicación "A" 3500) correspondiente al día hábil previo a la fecha de licitación (T-1)'
    pf_,pv_,cer_,usd_=[],[],[],[]
    seen=set()
    def add(l,k,d):
        if k not in seen: seen.add(k); l.append(d)

    def inst_pesos(label,f,tasa,precio,ajuste,parametro):
        return {"label":label,"vencimiento":f,"tasa":tasa,"precio":precio,"ajuste":ajuste,
                "parametro":parametro,"amort":"Íntegra al vencimiento",
                "monto":"Hasta el monto máximo autorizado por la normativa vigente"}

    # LECAP nuevas
    for m in re.finditer(r"LETRA DEL TESORO NACIONAL CAPITALIZABLE EN PESOS CON VENCIMIENTO ([\d\w\s]+?)\(NUEVA\)",T):
        f=pf(m.group(1))
        if f: add(pf_,f"LN{f}",inst_pesos("LECAP (nueva)",f,"A licitar","$ 1.000,00 por cada VNO $ 1.000","N/A","TEM"))
    # LECAP reaperturas
    for m in re.finditer(r"LETRA DEL TESORO NACIONAL CAPITALIZABLE EN PESOS CON VENCIMIENTO [\d\w\s]+?\((\w+)\s*-\s*REAPERTURA\)",T):
        tk=m.group(1); f=ticker_fecha(T,tk)
        add(pf_,f"LR{tk}",inst_pesos(f"LECAP ({tk} - reapertura)",f,"A licitar","A licitar","N/A","Precio"))
    # BONCAP reaperturas
    for m in re.finditer(r"BONO DEL TESORO NACIONAL CAPITALIZABLE EN PESOS CON VENCIMIENTO [\d\w\s]+?\((\w+)\s*-\s*REAPERTURA\)",T):
        tk=m.group(1); f=ticker_fecha(T,tk)
        add(pf_,f"BCAP{tk}",inst_pesos(f"BONCAP ({tk} - reapertura)",f,"A licitar","A licitar","N/A","Precio"))
    # LETAMAR reaperturas
    for m in re.finditer(r"LETRA DEL TESORO NACIONAL EN PESOS A TASA TAMAR CON VENCIMIENTO ([\d\w\s]+?)\((\w+)\s*-\s*REAPERTURA\)",T):
        f=pf(m.group(1)); tk=m.group(2)
        add(pv_,f"LT{tk}",inst_pesos(f"LETAM ({tk} - reapertura)",f,"","A licitar","N/A","Precio"))
    # BOTAMAR reaperturas
    for m in re.finditer(r"BONO DEL TESORO NACIONAL EN PESOS A TASA TAMAR CON VENCIMIENTO ([\d\w\s]+?)\((\w+)\s*[-–]\s*REAPERTURA\)",T):
        f=pf(m.group(1)); tk=m.group(2)
        add(pv_,f"BT{tk}",inst_pesos(f"BOTAM ({tk} - reapertura)",f,"","A licitar","N/A","Precio"))
    # BOTAMAR nuevos
    for m in re.finditer(r"BONO DEL TESORO NACIONAL EN PESOS A TASA TAMAR CON VENCIMIENTO ([\d\w\s]+?)\(NUEVO\)",T):
        f=pf(m.group(1))
        if f: add(pv_,f"BTN{f}",inst_pesos("BOTAM (nuevo)",f,"A licitar","$ 1.000,00 por cada VNO $ 1.000","N/A","Tasa"))
    # LECER reaperturas
    for m in re.finditer(r"LETRA DEL TESORO NACIONAL EN PESOS AJUSTADA? POR CER A DESCUENTO VENCIMIENTO [\d\w\s]+?\((\w+)\s*-\s*REAPERTURA\)",T):
        tk=m.group(1); f=ticker_fecha(T,tk)
        add(cer_,f"CR{tk}",inst_pesos(f"LECER ({tk} - reapertura)",f,"Cero Cupón","A licitar","CER","Precio"))
    # LECER nuevas
    for m in re.finditer(r"LETRA DEL TESORO NACIONAL EN PESOS AJUSTADA? POR CER A DESCUENTO VENCIMIENTO ([\d\w\s]+?)\(NUEVA\)",T):
        f=pf(m.group(1))
        if f: add(cer_,f"CN{f}",inst_pesos("LECER (Nueva)",f,"Cero Cupón","A licitar","CER","Precio"))
    # BONCER reaperturas
    for m in re.finditer(r"BONO DEL TESORO NACIONAL EN PESOS CERO CUP[ÓO]N CON AJUSTE POR CER VENCIMIENTO [\d\w\s]+?\((\w+)\s*[-–]\s*REAPERTURA\)",T):
        tk=m.group(1); f=ticker_fecha(T,tk)
        add(cer_,f"BR{tk}",inst_pesos(f"BONCER ({tk} – reapertura)",f,"Cero Cupón","A licitar","CER","Precio"))
    # BONCER nuevos
    for m in re.finditer(r"BONO DEL TESORO NACIONAL EN PESOS CERO CUP[ÓO]N CON AJUSTE POR CER VENCIMIENTO ([\d\w\s]+?)\(NUEVO\)",T):
        f=pf(m.group(1))
        if f: add(cer_,f"BRN{f}",inst_pesos("BONCER (nuevo)",f,"Cero Cupón","A licitar","CER","Precio"))
    # LELINK reaperturas
    for m in re.finditer(r"LETRA DEL TESORO NACIONAL VINCULADA AL D[ÓO]LAR ESTADOUNIDENSE CERO CUP[ÓO]N CON VENCIMIENTO [\d\w\s]+?\((\w+)\s*-\s*REAPERTURA\)",T):
        tk=m.group(1); f=ticker_fecha(T,tk)
        add(usd_,f"LL{tk}",{"tipo":"LELINK","label":f"LELINK ({tk} - reapertura)","vencimiento":f,"tasa":"Cero Cupón","precio":"A licitar","parametro":"Precio","mon_em":"Dólares Estadounidenses","mon_sus":sus_ll,"mon_pago":"Pesos al tipo de cambio aplicable","amort":"Íntegra al vencimiento","monto":"Hasta el monto máximo autorizado por la normativa vigente"})
    # LELINK nuevas
    for m in re.finditer(r"LETRA DEL TESORO NACIONAL VINCULADA AL D[ÓO]LAR ESTADOUNIDENSE CERO CUP[ÓO]N CON VENCIMIENTO ([\d\w\s]+?)\(NUEVA\)",T):
        f=pf(m.group(1))
        if f: add(usd_,f"LLN{f}",{"tipo":"LELINK","label":"LELINK (Nuevo)","vencimiento":f,"tasa":"Cero Cupón","precio":"A licitar","parametro":"Precio","mon_em":"Dólares Estadounidenses","mon_sus":sus_ll,"mon_pago":"Pesos al tipo de cambio aplicable","amort":"Íntegra al vencimiento","monto":"Hasta el monto máximo autorizado por la normativa vigente"})
    # BONAR
    for m in re.finditer(r"BONO DEL TESORO NACIONAL EN DOLARES ESTADOUNIDENSES.*?VENCIMIENTO [\d\w\s]+?\((\w+)\s*-\s*REAPERTURA\)",T):
        tk=m.group(1); f=ticker_fecha(T,tk)
        mmx=re.search(r"USD\s*(\d+)\s*MILLONES.*?(?:EN\s+)?(?:LA\s+)?PRIMERA VUELTA",T)
        monto_str=f"USD {mmx.group(1)} Millones en primera vuelta. (*)" if mmx else ""
        m_tna=re.search(r"ESTADOUNIDENSES\s+(\d+%)",T); tna=m_tna.group(1) if m_tna else "6%"
        mr=re.search(r"RESCATE ANTICIPADO.*?(\d{1,2}\s+DE\s+\w+\s+DE\s+\d{4})",T)
        fr_str=fstr(pf(mr.group(1))) if mr else ""
        add(usd_,f"BN{tk}",{"tipo":"BONAR","label":f"BONAR 2027 ({tk} - reapertura)","vencimiento":f,"tasa":f"TNA {tna} a pagar mensualmente","precio":"A licitar","parametro":"Precio","mon_em":"Dólares Estadounidenses","mon_sus":"En Dólares Estadounidenses","mon_pago":"En Dólares Estadounidenses","amort":"Íntegra al vencimiento","monto":monto_str,"opcion":"Los tenedores de los Bonos podrán ejercer, por única vez, una opción de rescate anticipado, total o parcial, del capital del Bono.","fecha_opcion":fr_str,"pago_int":"Los intereses serán pagaderos en Pesos por semestre vencido los días 30 de mayo y 30 de noviembre de cada año hasta la fecha de vencimiento"})

    return {"pf":pf_,"pv":pv_,"cer":cer_,"usd":usd_,
            "h":{"h_ini":h_ini,"h_fin":h_fin,"fecha_lic":fecha_lic_str,"sv":sv,
                 "liq":fecha_liq,"liq_str":fstr(fecha_liq),"em":fecha_em}}

# ── Excel ─────────────────────────────────────────────────────────────────────
def generar_excel(datos):
    wb=Workbook(); ws=wb.active
    h=datos["h"]; fecha_em=h["em"]
    if h["liq"]: ws.title=h["liq"].strftime("%d.%m.%Y")
    ws.sheet_view.showGridLines=False

    ws.column_dimensions["A"].width=4
    ws.column_dimensions["B"].width=4
    ws.column_dimensions["C"].width=30
    for i in range(4,18): ws.column_dimensions[get_column_letter(i)].width=24

    CL=3; CD=4

    def hdr_row(r,negrita,normal,h_row=18):
        ws.row_dimensions[r].height=h_row
        c1=ws.cell(row=r,column=CL,value=negrita)
        c1.font=Font(bold=True,size=SZ,name="Calibri"); c1.alignment=al(h="left",wrap=False)
        c2=ws.cell(row=r,column=CL+1,value=normal)
        c2.font=Font(bold=False,size=SZ,name="Calibri"); c2.alignment=al(h="left",wrap=False)
        ws.merge_cells(start_row=r,start_column=CL+1,end_row=r,end_column=12)

    ws.row_dimensions[7].height=18
    c=ws.cell(row=7,column=CL,value="LICITACIÓN DEL TESORO")
    c.font=Font(bold=True,size=SZ,name="Calibri"); c.alignment=al(h="left",wrap=False)
    ws.row_dimensions[9].height=18
    c=ws.cell(row=9,column=CL,value="LICITACION POR EFECTIVO")
    c.font=Font(bold=True,size=SZ,name="Calibri"); c.alignment=al(h="left",wrap=False)
    hdr_row(11,"Período de Licitación Pública:",f" desde las {h['h_ini']} hs hasta las {h['h_fin']} hs del  {h['fecha_lic']}")
    if h["sv"]:
        partes=h["sv"].split(":",1)
        hdr_row(12,partes[0]+":",partes[1] if len(partes)>1 else "")
    hdr_row(13,"Fecha de Liquidación:",f" {h['liq_str']} (T+2)")

    fila=[15]
    def R(n=1): r=fila[0]; fila[0]+=n; return r

    def bloque(titulo,insts,es_usd=False):
        if not insts: return
        n=len(insts); c1=CD; c2=CD+n-1

        r=R(); ws.row_dimensions[r].height=21
        ct=ws.cell(row=r,column=CL,value=titulo)
        ct.font=Font(italic=True,bold=True,size=SZ,name="Calibri")
        ct.alignment=al(h="left",wrap=False)
        R()  # fila vacía

        # Header naranja
        r=R(); ws.row_dimensions[r].height=42
        cc=ws.cell(row=r,column=CL); cc.fill=F(NARANJA); cc.border=B()
        for i,inst in enumerate(insts):
            c=ws.cell(row=r,column=c1+i,value=inst["label"])
            c.fill=F(NARANJA); c.font=fn(color=BLANCO)
            c.alignment=al(); c.border=B()

        def fila_ind(label,getter,gris,h_row=21,fmt=None):
            r=R(); ws.row_dimensions[r].height=h_row
            aplicar(ws.cell(row=r,column=CL),gris,valor=label,h="left")
            for i,inst in enumerate(insts):
                val=getter(inst)
                c=ws.cell(row=r,column=c1+i)
                aplicar(c,gris,valor=val if val is not None else "")
                if fmt: c.number_format=fmt

        def fila_merge(label,valor,gris,h_row=21):
            r=R(); ws.row_dimensions[r].height=h_row
            aplicar(ws.cell(row=r,column=CL),gris,valor=label,h="left")
            merge_escribir(ws,r,c1,c2,gris,valor)

        if not es_usd:
            fila_ind("Vencimiento",    lambda x:x["vencimiento"], GRIS_CLA, fmt="D/M/YYYY")
            fila_ind("Plazo",          lambda x:dias_entre(x["vencimiento"],fecha_em), GRIS_OSC)
            fila_merge("Moneda de emision",     "Pesos", GRIS_CLA)
            fila_merge("Moneda de Suscripcion", "Pesos", GRIS_CLA)
            fila_merge("Moneda de Pago",        "Pesos", GRIS_CLA)
            fila_ind("Tasa de interés ",  lambda x:x.get("tasa",""),     GRIS_OSC)
            fila_ind("Precio",            lambda x:x.get("precio",""),   GRIS_CLA, h_row=42)
            fila_ind("Ajuste de capital", lambda x:x.get("ajuste",""),   GRIS_OSC)
            fila_ind("Párametro a licitar",lambda x:x.get("parametro",""),GRIS_CLA)
            fila_merge("Amortización",           "Íntegra al vencimiento", GRIS_OSC)
            fila_merge("Monto Máximo a Licitar", "Hasta el monto máximo autorizado por la normativa vigente", GRIS_CLA)
            fila_merge("Ley aplicable",          "Ley de la REPÚBLICA ARGENTINA", GRIS_OSC)
        else:
            fila_ind("Vencimiento",       lambda x:x["vencimiento"],      GRIS_CLA, fmt="D/M/YYYY")
            fila_ind("Moneda de emision", lambda x:x.get("mon_em",""),    GRIS_OSC)
            fila_ind("Moneda de Suscripcion",lambda x:x.get("mon_sus",""),GRIS_CLA, h_row=84)
            fila_ind("Moneda de Pago",    lambda x:x.get("mon_pago",""),  GRIS_OSC, h_row=42)
            fila_ind("Tasa de interés",   lambda x:x.get("tasa",""),      GRIS_CLA, h_row=42)
            fila_ind("Precio",            lambda x:x.get("precio",""),    GRIS_OSC)
            fila_ind("Párametro a licitar",lambda x:x.get("parametro",""),GRIS_CLA)
            fila_ind("Amortización",      lambda x:x.get("amort",""),     GRIS_OSC)
            fila_ind("Monto Máximo a Licitar",lambda x:x.get("monto",""), GRIS_CLA, h_row=42)
            fila_merge("Ley aplicable",   "Ley de la REPÚBLICA ARGENTINA", GRIS_OSC)

        R()  # fila vacía al final

    def bonar_especial(usd_list):
        bonar=[x for x in usd_list if x.get("tipo")=="BONAR"]
        if not bonar: return
        b=bonar[0]; n=len(usd_list); c1=CD; c2=CD+n-1
        r=R(); ws.row_dimensions[r].height=21
        ct=ws.cell(row=r,column=CL,value="Instrumentros a licitar en dólares")
        ct.font=Font(italic=True,bold=True,size=SZ,name="Calibri"); ct.alignment=al(h="left",wrap=False)
        for label,val,h_row,gris in [
            ("Opción de rescate anticipado",    b.get("opcion",""),    42,GRIS_CLA),
            ("Fecha de ejercicio de la Opción", b.get("fecha_opcion",""),21,GRIS_OSC),
            ("Forma de pago de los servicios\n de interés",b.get("pago_int",""),63,GRIS_CLA),
        ]:
            r=R(); ws.row_dimensions[r].height=h_row
            aplicar(ws.cell(row=r,column=CL),gris,valor=label,h="left")
            merge_escribir(ws,r,c1,c2,gris,val,h="left")
        R()

    pesos=datos["pf"]+datos["pv"]
    if pesos:        bloque("Instrumentos a licitar en pesos a tasa fija y tasa variable:",pesos)
    if datos["cer"]: bloque("Instrumentos a licitar en pesos ajustados por CER:",datos["cer"])
    if datos["usd"]:
        bonar_especial(datos["usd"])
        bloque("Instrumentros a licitar en dólares",datos["usd"],es_usd=True)
        if any(x.get("tipo")=="BONAR" for x in datos["usd"]):
            r=R(2)
            ws.cell(row=r,column=CL,
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
                total=sum(len(datos[k]) for k in ["pf","pv","cer","usd"])
                st.success(f"✅ {total} instrumentos detectados")
                for bq,nm in [("pf","💵 Tasa fija"),("pv","📊 Tasa variable"),("cer","📈 CER"),("usd","💲 Dólares")]:
                    if datos[bq]:
                        st.markdown(f"**{nm}**")
                        for inst in datos[bq]:
                            v=inst["vencimiento"].strftime("%d/%m/%Y") if inst["vencimiento"] else "N/A"
                            st.markdown(f"&nbsp;&nbsp;&nbsp;• `{inst['label']}` — vto. {v}")
                if datos["h"]["liq_str"]:
                    st.markdown(f"📅 **Liquidación:** {datos['h']['liq_str']}")
                fl=datos["h"]["liq"]
                nombre=f"Licitacion_Tesoro_{fl.strftime('%d_%m_%Y')}.xlsx" if fl else "Licitacion_nueva.xlsx"
                st.markdown("---")
                st.download_button("⬇️ Descargar Excel",data=excel_out,file_name=nombre,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"❌ Error: {str(e)}"); st.exception(e)
else:
    st.info("⬆️ Subí el PDF para habilitar la generación.")

st.markdown("---")
st.caption("Banco Hipotecario · Mercado de Capitales · Emisiones Primarias")
