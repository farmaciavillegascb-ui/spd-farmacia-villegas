import os
import io
import re  # Librería para procesar el texto del DataMatrix
# ----------------------------------------------------
# FORZAR TEMA CLARO (LIGHT MODE) AUTOMÁTICAMENTE
# ----------------------------------------------------
if not os.path.exists(".streamlit"):
    os.makedirs(".streamlit")
config_path = ".streamlit/config.toml"
if not os.path.exists(config_path):
    with open(config_path, "w") as f:
        f.write("[theme]\nbase='light'\nprimaryColor='#0ea5e9'\n")

import streamlit as st
import pandas as pd
import time
import uuid
import unicodedata
from fpdf import FPDF
from datetime import datetime

# Configuración de la página optimizada
st.set_page_config(
    page_title="SPD FARMACIA VILLEGAS",
    page_icon="💊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Estilos CSS avanzados (AUMENTADO TAMAÑO DE BOTONES E ICONOS)
st.markdown("""
<style>
    .block-container { padding-top: 0.5rem !important; padding-bottom: 2rem !important; padding-left: 1rem !important; padding-right: 1rem !important; }
    
    /* Forzar fondo de la aplicación claro */
    .stApp, .main { background-color: #f7f9fc !important; }
    
    #MainMenu {visibility: hidden;} footer {visibility: hidden;} header {visibility: hidden;}
    [data-testid="collapsedControl"] { display: none; }
    
    .dashboard-header { background: #ffffff !important; padding: 12px 16px; border-radius: 12px; box-shadow: 0 4px 20px rgba(0, 0, 0, 0.05); margin-bottom: 12px; border: 1px solid #e2e8f0; }
    .logo-container { display: flex; align-items: center; justify-content: center; gap: 8px; margin-bottom: 8px; }
    .logo-title { font-size: 20px; font-weight: 800; color: #1e293b; letter-spacing: 0.5px; }
    .status-bar { display: flex; justify-content: space-between; align-items: center; background: #f8fafc !important; padding: 6px 12px; border-radius: 8px; font-size: 12px; color: #475569; font-weight: 600; margin-bottom: 10px; border: 1px solid #e2e8f0; }
    
    /* FORZAR INPUTS (USUARIO/CLAVE/BUSCADOR) EN BLANCO */
    div[data-baseweb="input"] > div { background-color: #ffffff !important; border: 1px solid #cbd5e1 !important; }
    input { background-color: #ffffff !important; color: #1e293b !important; -webkit-text-fill-color: #1e293b !important; font-weight: 500 !important; }
    
    /* FORZAR SELECTORES EN BLANCO */
    div[data-baseweb="select"] > div { background-color: #ffffff !important; border: 1px solid #cbd5e1 !important; }
    div[data-baseweb="select"] * { color: #1e293b !important; }
    
    /* FORZAR BOTONES EN BLANCO/CLARO Y AUMENTAR SU TAMAÑO/ICONOS */
    div.stButton > button { 
        width: 100% !important; 
        height: 52px !important; 
        border-radius: 10px !important; 
        font-weight: 800 !important; 
        font-size: 13px !important; 
        background-color: #ffffff !important; 
        color: #334155 !important; 
        border: 2px solid #cbd5e1 !important; 
        box-shadow: 0 2px 6px rgba(0,0,0,0.03) !important; 
        transition: all 0.2s ease-in-out !important; 
    }
    div.stButton > button:hover { background-color: #f1f5f9 !important; border-color: #0ea5e9 !important; color: #0284c7 !important; transform: translateY(-1px); }
    
    /* FORZAR FORMULARIOS EN BLANCO */
    div[data-testid="stForm"] { background-color: #ffffff !important; border-color: #e2e8f0 !important; }
    
    [data-testid="column"] { display: flex !important; flex-direction: column !important; align-items: stretch !important; }
    [data-testid="column"] > div { display: flex !important; flex-direction: column !important; flex-grow: 1 !important; }
    
    @keyframes pulse-subtle { 
        0% { transform: scale(1); box-shadow: 0 0 0 0 rgba(239, 68, 68, 0.6); } 
        50% { transform: scale(1.03); box-shadow: 0 0 0 12px rgba(239, 68, 68, 0); } 
        100% { transform: scale(1); box-shadow: 0 0 0 0 rgba(239, 68, 68, 0); } 
    }

    @media (max-width: 768px) {
        .block-container { padding-top: 0.3rem !important; padding-left: 0.5rem !important; padding-right: 0.5rem !important; }
        .logo-title { font-size: 16px !important; }
        .status-bar { font-size: 11px !important; padding: 4px 8px !important; }
        /* TAMAÑO DE BOTONES MÁS GRANDE PARA MÓVIL */
        div.stButton > button { height: 48px !important; font-size: 12px !important; padding: 2px !important; }
        h2 { font-size: 1.25rem !important; text-align: center; }
        h3 { font-size: 1.1rem !important; }
    }
</style>
""", unsafe_allow_html=True)

# ----------------------------------------------------
# CLASES Y FUNCIONES PARA PDF
# ----------------------------------------------------
class PDFAlbaran(FPDF):
    def footer(self):
        self.set_y(-20)
        y_line = self.get_y()
        self.line(40, y_line, 100, y_line)
        self.line(197, y_line, 257, y_line)
        self.set_y(y_line + 2)
        self.set_font('Arial', 'I', 10)
        self.set_x(40)
        self.cell(60, 5, "Firma Farmaceutico", align='C')
        self.set_x(197)
        self.cell(60, 5, "Firma Enfermera", align='C')

class PDFReporteIncidencias(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 15)
        self.cell(0, 10, limpiar_texto_pdf('FARMACIA VILLEGAS C.B. - REPORTE DE INCIDENCIAS'), 0, 1, 'C')
        self.set_font('Arial', 'I', 10)
        self.cell(0, 6, limpiar_texto_pdf(f'Fecha de emision: {datetime.now().strftime("%d/%m/%Y %H:%M")}'), 0, 1, 'C')
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Pagina {self.page_no()}', 0, 0, 'C')

def limpiar_texto_pdf(texto):
    if not texto: return ""
    return unicodedata.normalize('NFKD', str(texto).replace('ñ','n').replace('Ñ','N').replace('º','.').replace('ª','.')).encode('ascii', 'ignore').decode('ascii')

def generar_pdf_incidencias(incidencias):
    pdf = PDFReporteIncidencias(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    
    agrupadas = {}
    for inc in incidencias:
        pac = f"{inc.get('paciente', 'Desconocido')} (Ref: {inc.get('ref_paciente', '')})"
        if pac not in agrupadas:
            agrupadas[pac] = []
        agrupadas[pac].append(inc)
        
    for pac, lista in agrupadas.items():
        pdf.set_font("Arial", 'B', 12)
        pdf.set_fill_color(230, 240, 255)
        pdf.cell(0, 8, limpiar_texto_pdf(f"  Paciente: {pac}"), 0, 1, 'L', fill=True)
        pdf.ln(3)
        
        for inc in lista:
            med = str(inc.get('medicamento', ''))
            cn = str(inc.get('cn', ''))
            mot = str(inc.get('motivo', ''))
            obs = str(inc.get('observaciones', ''))
            fecha = str(inc.get('fecha', ''))
            resuelta = inc.get('resuelta_por_enfermera', False)
            
            estado = "[RESUELTA]" if resuelta else "[PENDIENTE]"
            
            pdf.set_font("Arial", 'B', 10)
            pdf.cell(5, 5, "", 0, 0)
            pdf.cell(0, 5, limpiar_texto_pdf(f"> {med} (CN: {cn}) {estado}"), 0, 1)
            
            pdf.set_font("Arial", '', 9)
            pdf.cell(10, 5, "", 0, 0)
            pdf.multi_cell(0, 5, limpiar_texto_pdf(f"Motivo: {mot} | Registrada: {fecha}"))
            
            if obs:
                pdf.cell(10, 5, "", 0, 0)
                pdf.multi_cell(0, 5, limpiar_texto_pdf(f"Observaciones: {obs}"))
            
            pdf.ln(3)
        pdf.ln(2)
        
    return pdf.output(dest='S').encode('latin1')

def dibujar_tabla_pdf(pdf, headers, rows_data, col_widths, align_list=None):
    header_height = 8
    pdf.set_font("Arial", 'B', 8)
    x_start, y_start = pdf.get_x(), pdf.get_y()
    
    for i, h_text in enumerate(headers):
        w = col_widths[i]
        cx, cy = pdf.get_x(), pdf.get_y()
        pdf.rect(cx, cy, w, header_height)
        pdf.set_xy(cx, cy + 1.5)
        pdf.cell(w, 5, limpiar_texto_pdf(h_text), align='C', ln=0)
        pdf.set_xy(cx + w, y_start)
    pdf.set_xy(x_start, y_start + header_height)
    
    pdf.set_font("Arial", '', 7.5)
    for row in rows_data:
        cell_lines = []
        max_num_lines = 1
        for i, val in enumerate(row):
            w = col_widths[i]
            lines = pdf.multi_cell(w - 2, 3.5, str(val if val is not None else ""), split_only=True)
            cell_lines.append(lines)
            if len(lines) > max_num_lines: max_num_lines = len(lines)
        
        row_height = max(6, max_num_lines * 3.5 + 2)
        xr_start, yr_start = pdf.get_x(), pdf.get_y()
        
        if yr_start + row_height > 180:
            pdf.add_page(); pdf.set_font("Arial", 'B', 8)
            hx, hy = pdf.get_x(), pdf.get_y()
            for i, h_text in enumerate(headers):
                w = col_widths[i]
                cx, cy = pdf.get_x(), pdf.get_y()
                pdf.rect(cx, cy, w, header_height)
                pdf.set_xy(cx, cy + 1.5)
                pdf.cell(w, 5, limpiar_texto_pdf(h_text), align='C', ln=0)
                pdf.set_xy(cx + w, hy)
            pdf.set_xy(hx, hy + header_height)
            pdf.set_font("Arial", '', 7.5)
            yr_start, xr_start = pdf.get_y(), pdf.get_x()

        for i, lines in enumerate(cell_lines):
            w = col_widths[i]
            align = align_list[i] if align_list and i < len(align_list) else 'L'
            cx, cy = pdf.get_x(), pdf.get_y()
            pdf.rect(cx, cy, w, row_height)
            current_y = cy + 1
            for line in lines:
                pdf.set_xy(cx + 1, current_y)
                pdf.cell(w - 2, 3.5, limpiar_texto_pdf(line), align=align, ln=0)
                current_y += 3.5
            pdf.set_xy(cx + w, yr_start)
        pdf.set_xy(xr_start, yr_start + row_height)

def generar_albaran_devolucion_pdf(nombre_paciente, ref_paciente, lista_devolucion):
    pdf = PDFAlbaran(orientation='L', unit='mm', format='A4') 
    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, limpiar_texto_pdf("FARMACIA VILLEGAS C.B. - ALBARAN DE DEVOLUCION"), ln=True, align='C')
    pdf.set_font("Arial", '', 11)
    pdf.cell(0, 6, limpiar_texto_pdf(f"Paciente: {nombre_paciente} (Ref: {ref_paciente})"), ln=True, align='L')
    pdf.ln(5)
    headers = ["Medicamento", "Descripcion", "CN", "Lote", "Caducidad", "Serie", "Restantes"]
    col_widths = [62, 78, 22, 38, 28, 25, 22]
    align_list = ['L', 'L', 'C', 'C', 'C', 'C', 'C']
    rows_data = [[str(r.get('Medicamento','')), str(r.get('Descripción','')), str(r.get('CN','')), str(r.get('Lote','')), str(r.get('Caducidad','')), str(r.get('Serie','')), str(r.get('Pastillas restantes','0'))] for r in lista_devolucion]
    dibujar_tabla_pdf(pdf, headers, rows_data, col_widths, align_list)
    return pdf.output(dest='S').encode('latin1')

# ----------------------------------------------------
# CARGA DE BASE DE DATOS DE MEDICAMENTOS (EN CACHÉ)
# ----------------------------------------------------
@st.cache_data
def cargar_base_medicamentos():
    ruta_bd = "listado_de_medicamentos.xlsx"
    if not os.path.exists(ruta_bd): return {}
    try:
        df_bd = pd.read_excel(ruta_bd, sheet_name=0, usecols=['Cod. Nacional', 'Laboratorio', 'Presentación'])
        df_bd = df_bd.dropna(subset=['Cod. Nacional'])
        df_bd['CN'] = df_bd['Cod. Nacional'].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(6)
        partes = df_bd['Presentación'].astype(str).str.split(',', n=1, expand=True)
        df_bd['farmaco'] = partes[0].fillna('').astype(str).str.strip().str.slice(0, 45)
        df_bd['tamano'] = partes[1].fillna('').astype(str).str.strip().str.slice(0, 25) if partes.shape[1] > 1 else ""
        df_bd['marca'] = df_bd['Laboratorio'].fillna('').astype(str).str.slice(0, 25)
        df_bd = df_bd.set_index('CN')
        return df_bd[['marca', 'farmaco', 'tamano']].to_dict(orient='index')
    except Exception: return {}

BD_MEDICAMENTOS = cargar_base_medicamentos()
EXCEL_PATH = "Tratamientos_Por_Paciente.xlsx"

def cargar_datos_excel():
    if not os.path.exists(EXCEL_PATH): return {}
    xls = pd.ExcelFile(EXCEL_PATH)
    pacientes_dict = {}
    hojas_validas = [sh for sh in xls.sheet_names if sh not in ['Hoja2', 'Hoja3']]
    
    for hoja in hojas_validas:
        df_hoja = pd.read_excel(EXCEL_PATH, sheet_name=hoja)
        if not df_hoja.empty and 'Nombre' in df_hoja.columns:
            nombre_paciente = str(df_hoja['Nombre'].iloc[0]).strip()
            ref_paciente = str(df_hoja['Ref. paciente'].iloc[0]).strip() if 'Ref. paciente' in df_hoja.columns else hoja
            cip_paciente = str(df_hoja['CIP'].iloc[0]).strip() if 'CIP' in df_hoja.columns else ""
            etiqueta = f"{ref_paciente} — {nombre_paciente}"
            
            if 'Fecha inicio' not in df_hoja.columns: df_hoja['Fecha inicio'] = ""
            if 'Ultima Entrega' not in df_hoja.columns: df_hoja['Ultima Entrega'] = ""
            if 'Pedido' not in df_hoja.columns: df_hoja['Pedido'] = False
            else: df_hoja['Pedido'] = df_hoja['Pedido'].astype(bool)
            if 'Incidencia' not in df_hoja.columns: df_hoja['Incidencia'] = False
            else: df_hoja['Incidencia'] = df_hoja['Incidencia'].astype(bool)
            
            pacientes_dict[etiqueta] = {
                "ref": ref_paciente, "nombre": nombre_paciente, "cip": cip_paciente, "hoja": hoja, "datos": df_hoja
            }
    return pacientes_dict

# ----------------------------------------------------
# MOTOR DE LECTURA DATAMATRIX (ACTUALIZADO: LOTE = AI 10)
# ----------------------------------------------------
def traducir_datamatrix(raw_code, bd_medicamentos):
    res = {'marca': '', 'farmaco': '', 'tamano': '', 'cn': '', 'lote': '', 'caducidad': '', 'serie': ''}
    if not raw_code: return res
    
    # 1. Normalizar el carácter de control GS1
    texto = str(raw_code).replace('\x1D', '<GS>').replace(chr(29), '<GS>')
    
    try:
        # 2. Extraer el Código Nacional (CN)
        cn_712 = re.search(r'712(\d{6})', texto)
        if cn_712:
            res['cn'] = cn_712.group(1)
        else:
            gtin_match = re.search(r'01(\d{14})', texto)
            if gtin_match:
                gtin = gtin_match.group(1)
                res['cn'] = gtin[7:13]

        # 3. Extraer Caducidad (AI 17): Siempre son 6 dígitos AAMMDD
        cad_match = re.search(r'17(\d{6})', texto)
        if cad_match:
            aammdd = cad_match.group(1)
            yy, mm, dd = aammdd[0:2], aammdd[2:4], aammdd[4:6]
            if dd == '00': dd = '01'
            res['caducidad'] = f"{dd}/{mm}/20{yy}"
            
        # 4. Extraer el Número de Lote de Fabricación (AI 10)
        # Asignado directamente a la variable 'lote' para que aparezca en la columna Lote
        lote_match = re.search(r'10(.*?)(?:<GS>|$)', texto)
        if lote_match:
            lote_capturado = lote_match.group(1)
            if '<GS>' not in texto and '21' in lote_capturado and len(lote_capturado) > 6:
                lote_capturado = lote_capturado.split('21')[0]
            res['lote'] = lote_capturado.strip()[:20]

        # 5. Extraer Número de Serie por si acaso (AI 21)
        serie_match = re.search(r'21(.*?)(?:<GS>|$)', texto)
        if serie_match:
            res['serie'] = serie_match.group(1).strip()

    except Exception:
        pass

    # --- FALLBACKS ---
    if not res['cn']:
        nums = re.sub(r'\D', '', texto)
        res['cn'] = nums[-6:] if len(nums) >= 6 else "000000"
        
    if not res['lote']: res['lote'] = "LOTE_NO_LEIDO"
    if not res['caducidad']: res['caducidad'] = "31/12/2028"

    # --- BÚSQUEDA EN BASE DE DATOS DE MEDICAMENTOS ---
    cn_busqueda = res['cn'].zfill(6)
    if cn_busqueda in bd_medicamentos:
        datos = bd_medicamentos[cn_busqueda]
        res['marca'] = str(datos.get('marca', ''))
        res['farmaco'] = str(datos.get('farmaco', ''))
        res['tamano'] = str(datos.get('tamano', ''))
    else: 
        res['farmaco'] = f"Medicamento (CN: {res['cn']})"
        
    return res

@st.cache_resource
def get_shared_data():
    return {
        "lista_pacientes": cargar_datos_excel(),
        "solicitud_pedido": [], "pedidos_definitivos": [], "incidencias_activas": [], 
        "solicitudes_alta": [], "solicitudes_baja": [],
        "roles_sistema": {
            "admin": ["pacientes", "altas", "bajas", "propuesta", "pedidos_definitivos", "incidencias", "validar_incidencias", "usuarios"],
            "enfermera": ["pacientes", "altas", "bajas", "propuesta", "pedidos_definitivos", "incidencias"]
        },
        "usuarios_sistema": {
            "farmaciaB": {"clave": "farmaciaB2026", "rol": "admin"}, "farmaciaR": {"clave": "farmaciaR2026", "rol": "admin"},
            "FarmaciaC": {"clave": "FarmaciaC2026", "rol": "admin"}, "FarmaciaA": {"clave": "FarmaciaA2026", "rol": "admin"},
            "FarmaciasCH": {"clave": "FarmaciasCH2026", "rol": "admin"},
            "Enfermera1": {"clave": "enfermera12026", "rol": "enfermera"}, "Enfermera2": {"clave": "enfermera22026", "rol": "enfermera"}
        }
    }

shared_data = get_shared_data()
lista_pacientes = shared_data["lista_pacientes"]
USUARIOS_VALIDOS = shared_data["usuarios_sistema"]

if "usuario_autenticado" not in st.session_state: st.session_state["usuario_autenticado"] = None
if "rol_usuario" not in st.session_state: st.session_state["rol_usuario"] = None
if "pagina" not in st.session_state: st.session_state["pagina"] = "inicio"
if "paciente_seleccionado_key" not in st.session_state: st.session_state["paciente_seleccionado_key"] = None
if "modo_incidencia" not in st.session_state: st.session_state["modo_incidencia"] = False
if "borrador_incidencias" not in st.session_state: st.session_state["borrador_incidencias"] = pd.DataFrame()
if "baja_proceso_devolucion" not in st.session_state: st.session_state["baja_proceso_devolucion"] = None

def obtener_parametro_url(nombre):
    try: return st.query_params.get(nombre)
    except AttributeError:
        try: return st.experimental_get_query_params().get(nombre, [None])[0]
        except Exception: return None

def establecer_parametro_url(nombre, valor):
    try: st.query_params[nombre] = valor
    except AttributeError:
        try: st.experimental_set_query_params(**{nombre: valor})
        except Exception: pass

def limpiar_parametros_url():
    try: st.query_params.clear()
    except AttributeError:
        try: st.experimental_set_query_params()
        except Exception: pass

if "sesiones_activas" not in shared_data: shared_data["sesiones_activas"] = {}
token_url = obtener_parametro_url("session_token")

if token_url and token_url in shared_data["sesiones_activas"]:
    datos_sesion = shared_data["sesiones_activas"][token_url]
    if time.time() - datos_sesion["ultimo_acceso"] < 3600:
        shared_data["sesiones_activas"][token_url]["ultimo_acceso"] = time.time()
        st.session_state["usuario_autenticado"] = datos_sesion["usuario"]
        st.session_state["rol_usuario"] = datos_sesion["rol"]
    else:
        del shared_data["sesiones_activas"][token_url]; limpiar_parametros_url()
        st.session_state["usuario_autenticado"] = None; st.session_state["rol_usuario"] = None

# LOGIN
if st.session_state["usuario_autenticado"] is None:
    col1, col2, col3 = st.columns([0.5, 2, 0.5])
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("<h2 style='text-align: center; color: #1e293b;'>💊 SPD FARMACIA VILLEGAS</h2>", unsafe_allow_html=True)
        with st.form("form_login"):
            usuario_input = st.text_input("Usuario")
            clave_input = st.text_input("Clave de acceso", type="password")
            if st.form_submit_button("Iniciar Sesión", use_container_width=True):
                if usuario_input in USUARIOS_VALIDOS and USUARIOS_VALIDOS[usuario_input]["clave"] == clave_input:
                    nuevo_token = str(uuid.uuid4())
                    shared_data["sesiones_activas"][nuevo_token] = {"usuario": usuario_input, "rol": USUARIOS_VALIDOS[usuario_input]["rol"], "ultimo_acceso": time.time()}
                    establecer_parametro_url("session_token", nuevo_token)
                    st.session_state["usuario_autenticado"] = usuario_input
                    st.session_state["rol_usuario"] = USUARIOS_VALIDOS[usuario_input]["rol"]
                    st.session_state["pagina"] = "inicio"
                    st.rerun()
                else: st.error("❌ Usuario o clave incorrectos.")
    st.stop()

rol_actual = st.session_state["rol_usuario"]
permisos_usuario = shared_data["roles_sistema"].get(rol_actual, [])

# CABECERA RESPONSIVE
st.markdown('<div class="dashboard-header">', unsafe_allow_html=True)
st.markdown('<div class="logo-container"><span style="font-size: 22px;">💊</span><span class="logo-title">SPD FARMACIA VILLEGAS</span></div>', unsafe_allow_html=True)
st.markdown(f'<div class="status-bar"><span>Sistema activo <span style="color: #22c55e;">●</span></span><span>Usuario: <b>{st.session_state["usuario_autenticado"]}</b> ({rol_actual.upper()})</span></div>', unsafe_allow_html=True)

num_ped = len(shared_data["pedidos_definitivos"])
num_prop = len(shared_data["solicitud_pedido"])
num_inc = len(shared_data["incidencias_activas"])
num_altas = len([a for a in shared_data["solicitudes_alta"] if a["estado"] == "Pendiente"])
num_bajas = len([b for b in shared_data["solicitudes_baja"] if b["estado"] == "Pendiente"])

alert_ped = (num_ped > 0) or (num_prop > 0) or (num_altas > 0 and rol_actual == "admin") or (num_bajas > 0 and rol_actual == "admin")
alert_inc = any(inc.get("resuelta_por_enfermera", False) for inc in shared_data["incidencias_activas"]) or (num_inc > 0 and rol_actual == "admin")

if alert_ped: st.markdown("""<style>[data-testid="column"]:nth-child(5) div.stButton > button { background: linear-gradient(135deg, #ef4444, #dc2626) !important; color: white !important; }</style>""", unsafe_allow_html=True)
if alert_inc: st.markdown("""<style>[data-testid="column"]:nth-child(6) div.stButton > button { background: linear-gradient(135deg, #ef4444, #dc2626) !important; color: white !important; }</style>""", unsafe_allow_html=True)

col_inicio, col_sync, col_alta, col_baja, col_ped, col_inc, col_user, col_logout = st.columns([1.1, 1.1, 1.3, 1.3, 1.5, 1.5, 1.2, 1.1], gap="small")

with col_inicio:
    if st.button("🏠 INICIO", key="btn_hdr_inicio", use_container_width=True): 
        st.session_state["pagina"] = "inicio"; st.rerun()
with col_sync:
    if st.button("🔄 SYNC", key="btn_hdr_sync", use_container_width=True): 
        st.rerun()  
with col_alta:
    if "altas" in permisos_usuario:
        lbl = f"➕ ALTAS ({num_altas})" if num_altas>0 and rol_actual=="admin" else "➕ ALTAS"
        if st.button(lbl, key="btn_hdr_alta", use_container_width=True): 
            st.session_state["pagina"] = "alta_paciente"; st.rerun()
with col_baja:
    if "bajas" in permisos_usuario:
        lbl = f"➖ BAJAS ({num_bajas})" if num_bajas>0 and rol_actual=="admin" else "➖ BAJAS"
        if st.button(lbl, key="btn_hdr_bajas", use_container_width=True): 
            st.session_state["pagina"] = "baja_paciente"; st.rerun()
with col_ped:
    if "propuesta" in permisos_usuario or "pedidos_definitivos" in permisos_usuario:
        if rol_actual == "admin":
            lbl = f"🛒 PEDIDOS ({num_ped})" if num_ped>0 else "🛒 PEDIDOS"
            destino = "pedidos_definitivos_admin"
        else:
            lbl = f"📦 PROP. ({num_prop})" if num_prop>0 else "📦 PROPUESTA"
            destino = "seleccion_productos_enfermera"
            
        if st.button(lbl, key="btn_hdr_ped", use_container_width=True):
            st.session_state["pagina"] = destino
            st.rerun()
with col_inc:
    if "incidencias" in permisos_usuario:
        hay_aviso_enf = any(inc.get("resuelta_por_enfermera", False) for inc in shared_data["incidencias_activas"])
        lbl = "⚠️ ¡AVISO INCID!" if hay_aviso_enf else (f"⚠️ INCID. ({num_inc})" if num_inc>0 else "⚠️ INCIDENCIAS")
        if st.button(lbl, key="btn_hdr_inc", use_container_width=True): 
            st.session_state["pagina"] = "incidencias"; st.rerun()
with col_user:
    if "usuarios" in permisos_usuario:
        if st.button("👥 USUS.", key="btn_hdr_usu", use_container_width=True): 
            st.session_state["pagina"] = "gestion_usuarios"; st.rerun()
with col_logout:
    if st.button("🚪 SALIR", key="btn_hdr_out", use_container_width=True):
        if token_url in shared_data["sesiones_activas"]: del shared_data["sesiones_activas"][token_url]
        limpiar_parametros_url(); st.session_state["usuario_autenticado"] = None; st.session_state["pagina"] = "inicio"; st.rerun()

st.markdown('</div>', unsafe_allow_html=True)

# VISTAS PRINCIPALES
if st.session_state["pagina"] == "inicio":
    c1, c2 = st.columns(2, gap="medium")
    with c1:
        if "pacientes" in permisos_usuario:
            st.markdown(f'<div style="background: #eff6ff; padding: 18px; border-radius: 12px; border: 1px solid #bfdbfe; margin-bottom: 10px;"><h4 style="color: #1e3a8a; margin-top: 0; font-size: 16px;">👤 PACIENTES ({len(shared_data["lista_pacientes"])})</h4><p style="color: #334155; font-size: 13px; margin-bottom: 0;">Listado completo de pacientes y tratamientos.</p></div>', unsafe_allow_html=True)
            if st.button("🧓 **VER PACIENTES**", use_container_width=True): st.session_state["pagina"] = "lista_pacientes"; st.rerun()
    with c2:
        if "incidencias" in permisos_usuario:
            st.markdown(f'<div style="background: #fef2f2; padding: 18px; border-radius: 12px; border: 1px solid #fecaca; margin-bottom: 10px;"><h4 style="color: #7f1d1d; margin-top: 0; font-size: 16px;">⚠️ INCIDENCIAS ({len(shared_data["incidencias_activas"])})</h4><p style="color: #334155; font-size: 13px; margin-bottom: 0;">Panel de incidencias y recetas.</p></div>', unsafe_allow_html=True)
            if st.button("📋 **VER INCIDENCIAS**", use_container_width=True): st.session_state["pagina"] = "incidencias"; st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)
    c3, c4 = st.columns(2, gap="medium")
    with c3:
        if "propuesta" in permisos_usuario:
            st.markdown(f'<div style="background: #fefce8; padding: 18px; border-radius: 12px; border: 1px solid #fef08a; margin-bottom: 10px;"><h4 style="color: #713f12; margin-top: 0; font-size: 16px;">🚚 PROPUESTA ({len(shared_data["solicitud_pedido"])})</h4><p style="color: #334155; font-size: 13px; margin-bottom: 0;">Revisión y selección de pedidos.</p></div>', unsafe_allow_html=True)
            if st.button("📦 **VER PROPUESTA**", use_container_width=True): 
                st.session_state["pagina"] = "seleccion_productos_enfermera" if rol_actual != "admin" else "solicitud_pedido_admin"
                st.rerun()
    with c4:
        if "pedidos_definitivos" in permisos_usuario:
            st.markdown(f'<div style="background: #f0fdf4; padding: 18px; border-radius: 12px; border: 1px solid #bbf7d0; margin-bottom: 10px;"><h4 style="color: #14532d; margin-top: 0; font-size: 16px;">📄 PEDIDOS ({len(shared_data["pedidos_definitivos"])})</h4><p style="color: #334155; font-size: 13px; margin-bottom: 0;">Validación DataMatrix y albaranes.</p></div>', unsafe_allow_html=True)
            if st.button("🛒 **VER PEDIDOS**", use_container_width=True): st.session_state["pagina"] = "pedidos_definitivos_admin"; st.rerun()

# ----------------------------------------------------
# GESTIÓN DE BAJAS
# ----------------------------------------------------
elif st.session_state["pagina"] == "baja_paciente":
    if "bajas" not in permisos_usuario: st.stop()
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 800;'>👴 GESTIÓN DE BAJAS</h2>", unsafe_allow_html=True)

    if rol_actual != "admin":
        paciente_seleccionado = st.selectbox("Seleccione paciente activo:", [""] + list(shared_data["lista_pacientes"].keys()))
        if st.button("📤 Enviar Propuesta de Baja", use_container_width=True) and paciente_seleccionado:
            info_p = shared_data["lista_pacientes"][paciente_seleccionado]
            shared_data["solicitudes_baja"].append({"id": str(uuid.uuid4())[:8], "etiqueta": paciente_seleccionado, "nombre": info_p["nombre"], "ref": info_p["ref"], "estado": "Pendiente", "fecha": datetime.now().strftime("%d/%m/%Y %H:%M")})
            st.success("¡Enviada correctamente!"); time.sleep(1.5); st.rerun()

        st.markdown("---")
        
        c_tit, c_btn = st.columns([0.6, 0.4], gap="large")
        with c_tit:
            st.markdown("##### 📋 Mis Propuestas de Baja:")
        with c_btn:
            if any(b.get("estado", "").startswith("Validada") for b in shared_data["solicitudes_baja"]):
                if st.button("🧹 Quitar Validadas", use_container_width=True):
                    shared_data["solicitudes_baja"] = [b for b in shared_data["solicitudes_baja"] if not b.get("estado", "").startswith("Validada")]
                    st.rerun()

        if shared_data["solicitudes_baja"]: 
            st.dataframe(pd.DataFrame(shared_data["solicitudes_baja"])[['fecha', 'nombre', 'ref', 'estado']], use_container_width=True, hide_index=True)
            
    else:
        if "baja_a_confirmar" not in st.session_state: st.session_state["baja_a_confirmar"] = None
        
        if st.session_state["baja_a_confirmar"]:
            info_conf = st.session_state["baja_a_confirmar"]
            st.error(f"### ⚠️ ATENCIÓN: Va a proceder con la BAJA DEFINITIVA")
            st.markdown(f"**Paciente:** {info_conf['nombre']}")
            st.markdown("¿Desea continuar con la baja?")
            
            c_yes, c_no = st.columns(2)
            with c_yes:
                if st.button("✔️ CONFIRMAR BAJA", use_container_width=True):
                    if info_conf["tipo"] == "directa_sin_dev":
                        if info_conf["etiqueta"] in shared_data["lista_pacientes"]: del shared_data["lista_pacientes"][info_conf["etiqueta"]]
                        st.success("¡Paciente dado de baja correctamente!")
                    elif info_conf["tipo"] == "propuesta_sin_dev":
                        if info_conf["etiqueta"] in shared_data["lista_pacientes"]: del shared_data["lista_pacientes"][info_conf["etiqueta"]]
                        if "baja_ref" in info_conf and info_conf["baja_ref"] in shared_data["solicitudes_baja"]:
                            info_conf["baja_ref"]["estado"] = "Validada sin devolución"
                        st.success("¡Baja procesada!")
                    elif info_conf["tipo"] == "finalizar_dev":
                        if info_conf["etiqueta"] in shared_data["lista_pacientes"]: del shared_data["lista_pacientes"][info_conf["etiqueta"]]
                        if "baja_ref" in info_conf and info_conf["baja_ref"] in shared_data["solicitudes_baja"]:
                            info_conf["baja_ref"]["estado"] = "Validada con devolución"
                        st.session_state["baja_proceso_devolucion"] = None
                        st.session_state["df_devolucion_admin"] = pd.DataFrame(columns=['Medicamento', 'Descripción', 'CN', 'Lote', 'Caducidad', 'Serie', 'Pastillas restantes'])
                        st.success("¡Baja completada con devolución!")
                    
                    st.session_state["baja_a_confirmar"] = None
                    time.sleep(1.5); st.rerun()
            with c_no:
                if st.button("❌ ANULAR BAJA", use_container_width=True):
                    st.session_state["baja_a_confirmar"] = None
                    st.warning("Baja anulada.")
                    time.sleep(1.0); st.rerun()
                    
            st.stop()

        st.markdown("##### ⚡ Dar de Baja Directamente")
        paciente_directo = st.selectbox("Seleccione paciente para dar de baja:", [""] + list(shared_data["lista_pacientes"].keys()), key="select_baja_directa")
        col_dir1, col_dir2 = st.columns(2)
        with col_dir1:
            if st.button("✅ Dar de Baja Directa (SIN Devolución)", use_container_width=True) and paciente_directo:
                st.session_state["baja_a_confirmar"] = {"tipo": "directa_sin_dev", "nombre": shared_data["lista_pacientes"][paciente_directo]["nombre"], "etiqueta": paciente_directo}
                st.rerun()
        with col_dir2:
            if st.button("📦 Dar de Baja Directa (CON Devolución)", use_container_width=True) and paciente_directo:
                info_p = shared_data["lista_pacientes"][paciente_directo]
                baja_rapida = {"id": str(uuid.uuid4())[:8], "etiqueta": paciente_directo, "nombre": info_p["nombre"], "ref": info_p["ref"], "estado": "Pendiente"}
                st.session_state["baja_proceso_devolucion"] = baja_rapida
                st.rerun()

        st.markdown("---"); st.markdown("##### 📋 Propuestas Pendientes de Enfermería / Devoluciones:")
        pendientes_baja = [b for b in shared_data["solicitudes_baja"] if b["estado"] == "Pendiente"]
        
        if not pendientes_baja and not st.session_state.get("baja_proceso_devolucion"): 
            st.info("No hay propuestas de baja pendientes.")
        else:
            if "baja_proceso_devolucion" not in st.session_state: st.session_state["baja_proceso_devolucion"] = None
            
            if not st.session_state["baja_proceso_devolucion"]:
                for baja in pendientes_baja:
                    with st.container(border=True):
                        st.markdown(f"**Paciente:** {baja['nombre']} (Ref: {baja['ref']})")
                        col_b1, col_b2 = st.columns(2)
                        with col_b1:
                            if st.button(f"✅ Validar SIN Devolución", key=f"sin_dev_{baja['id']}", use_container_width=True):
                                st.session_state["baja_a_confirmar"] = {"tipo": "propuesta_sin_dev", "nombre": baja["nombre"], "etiqueta": baja["etiqueta"], "baja_ref": baja}
                                st.rerun()
                        with col_b2:
                            if st.button(f"📦 Validar CON Devolución", key=f"con_dev_{baja['id']}", use_container_width=True):
                                st.session_state["baja_proceso_devolucion"] = baja; st.rerun()

            if st.session_state["baja_proceso_devolucion"]:
                baja_activa = st.session_state["baja_proceso_devolucion"]
                st.markdown("---"); st.markdown(f"##### 📦 Proceso de Devolución para: **{baja_activa['nombre']}**")
                
                c_back1, c_back2 = st.columns(2)
                with c_back1:
                    if st.button("⬅️ Cancelar (Echar para atrás)", use_container_width=True):
                        st.session_state["baja_proceso_devolucion"] = None
                        st.rerun()
                with c_back2:
                    if st.button("🔄 Cambiar a SIN Devolución", use_container_width=True):
                        st.session_state["baja_a_confirmar"] = {"tipo": "propuesta_sin_dev", "nombre": baja_activa["nombre"], "etiqueta": baja_activa["etiqueta"], "baja_ref": baja_activa}
                        st.session_state["baja_proceso_devolucion"] = None
                        st.rerun()
                        
                if "df_devolucion_admin" not in st.session_state: st.session_state["df_devolucion_admin"] = pd.DataFrame(columns=['Medicamento', 'Descripción', 'CN', 'Lote', 'Caducidad', 'Serie', 'Pastillas restantes'])
                with st.form("form_dm_baja_admin", clear_on_submit=True):
                    cad_dm = st.text_input("Escanee DataMatrix devuelto:")
                    if st.form_submit_button("Añadir"):
                        if cad_dm:
                            parsed = traducir_datamatrix(cad_dm, BD_MEDICAMENTOS)
                            nuevo_reg = {'Medicamento': parsed['farmaco'], 'Descripción': f"{parsed['marca']} - {parsed['tamano']}".strip(" -"), 'CN': parsed['cn'], 'Lote': parsed['lote'], 'Caducidad': parsed['caducidad'], 'Serie': parsed['serie'], 'Pastillas restantes': 0}
                            st.session_state["df_devolucion_admin"] = pd.concat([st.session_state["df_devolucion_admin"], pd.DataFrame([nuevo_reg])], ignore_index=True); st.success("¡Añadido!")

                if not st.session_state["df_devolucion_admin"].empty:
                    st.session_state["df_devolucion_admin"] = st.data_editor(st.session_state["df_devolucion_admin"], use_container_width=True, hide_index=True)
                    c_pdf, c_fin = st.columns(2)
                    with c_pdf:
                        st.download_button("📄 Imprimir PDF", data=generar_albaran_devolucion_pdf(baja_activa['nombre'], baja_activa['ref'], st.session_state["df_devolucion_admin"].to_dict(orient="records")), file_name=f"Devolucion_{baja_activa['ref']}.pdf", mime="application/pdf", use_container_width=True)
                    with c_fin:
                        if st.button("💾 Finalizar Baja", use_container_width=True):
                            st.session_state["baja_a_confirmar"] = {"tipo": "finalizar_dev", "nombre": baja_activa["nombre"], "etiqueta": baja_activa["etiqueta"], "baja_ref": baja_activa}
                            st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("⬅ Volver al Menú", use_container_width=True): st.session_state["pagina"] = "inicio"; st.rerun()

# ----------------------------------------------------
# GESTIÓN DE ALTAS
# ----------------------------------------------------
elif st.session_state["pagina"] == "alta_paciente":
    if "altas" not in permisos_usuario: st.stop()
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 800;'>👴 GESTIÓN DE ALTAS</h2>", unsafe_allow_html=True)

    if rol_actual != "admin":
        if "df_alta_cargado" not in st.session_state: st.session_state["df_alta_cargado"] = pd.DataFrame(columns=['Medicamento', 'CN', 'Posologia', 'Ultima Entrega'])
        if "reset_alta_enf" not in st.session_state: st.session_state["reset_alta_enf"] = 0
        
        rk = st.session_state["reset_alta_enf"]

        st.markdown("##### 📝 Nueva Propuesta de Alta")
        with st.container(border=True):
            ref_input = st.text_input("Código del Paciente (Ref) - [Lo asignará Farmacia]:", value=st.session_state.get("alta_input_ref", ""), key=f"ref_enf_{rk}")
            nom_input = st.text_input("Nombre Completo (Obligatorio):", value=st.session_state.get("alta_input_nombre", ""), key=f"nom_enf_{rk}")
            cip_input = st.text_input("Código CIP (Obligatorio):", value=st.session_state.get("alta_input_cip", ""), key=f"cip_enf_{rk}")
            
            meds_editadas = st.data_editor(st.session_state["df_alta_cargado"], num_rows="dynamic", key=f"editor_alta_paciente_{rk}", use_container_width=True)
            
            if st.button("📤 Enviar Propuesta de Alta", use_container_width=True):
                if nom_input.strip() and cip_input.strip():
                    df_final = meds_editadas.copy()
                    df_final['Ultima Entrega'] = ""
                    shared_data["solicitudes_alta"].append({
                        "id": str(uuid.uuid4())[:8], 
                        "nombre": nom_input.strip(), 
                        "cip": cip_input.strip(), 
                        "ref": ref_input.strip(), 
                        "datos": df_final, 
                        "estado": "Pendiente", 
                        "observacion": "", 
                        "fecha": datetime.now().strftime("%d/%m/%Y %H:%M")
                    })
                    st.success("¡Propuesta enviada correctamente!")
                    st.session_state["alta_input_ref"] = ""
                    st.session_state["alta_input_nombre"] = ""
                    st.session_state["alta_input_cip"] = ""
                    st.session_state["df_alta_cargado"] = pd.DataFrame(columns=['Medicamento', 'CN', 'Posologia', 'Ultima Entrega'])
                    st.session_state["reset_alta_enf"] += 1 
                    time.sleep(1.0)
                    st.rerun()
                else: 
                    st.warning("⚠️ El Nombre Completo y el Código CIP son campos obligatorios.")

        st.markdown("---")
        c_tit, c_btn = st.columns([0.6, 0.4], gap="large")
        with c_tit:
            st.markdown("##### 📋 Mis Propuestas de Alta:")
        with c_btn:
            if any(a["estado"] == "Validada" for a in shared_data["solicitudes_alta"]):
                if st.button("🧹 Quitar Validadas", use_container_width=True):
                    shared_data["solicitudes_alta"] = [a for a in shared_data["solicitudes_alta"] if a["estado"] != "Validada"]
                    st.rerun()

        for alta in shared_data["solicitudes_alta"]:
            color_estado = "#fef08a" if alta["estado"] == "Pendiente" else ("#bbf7d0" if alta["estado"] == "Validada" else "#fecaca")
            with st.container(border=True):
                st.markdown(f"**{alta['nombre']}** (Cód: {alta['ref'] if alta['ref'] else 'Pendiente'}) — <span style='background-color: {color_estado}; padding: 2px 6px; border-radius: 4px;'><b>{alta['estado']}</b></span>", unsafe_allow_html=True)
                if alta["estado"] == "Rechazada":
                    st.error(f"❌ Motivo de rechazo: {alta['observacion']}")
                    if st.button(f"🔄 Corregir (Volcar a la ficha)", key=f"re_enviar_{alta['id']}", use_container_width=True): 
                        st.session_state["alta_input_nombre"] = alta["nombre"]
                        st.session_state["alta_input_cip"] = alta["cip"]
                        st.session_state["alta_input_ref"] = alta["ref"]
                        st.session_state["df_alta_cargado"] = alta["datos"]
                        st.session_state["reset_alta_enf"] += 1 
                        shared_data["solicitudes_alta"].remove(alta)
                        st.rerun()
    else:
        tabs_admin = st.tabs(["⚡ Alta Directa", "📥 Carga Masiva (Excel)", "📋 Propuestas Pendientes"])
        
        with tabs_admin[0]:
            if "df_alta_admin_directo" not in st.session_state: st.session_state["df_alta_admin_directo"] = pd.DataFrame(columns=['Medicamento', 'CN', 'Posologia', 'Ultima Entrega'])
            if "reset_alta_admin" not in st.session_state: st.session_state["reset_alta_admin"] = 0
            
            rk_admin = st.session_state["reset_alta_admin"]

            st.markdown("##### ⚡ Alta Directa Individual")
            with st.container(border=True):
                ref_dir = st.text_input("Código del Paciente (Ref):", value=st.session_state.get("admin_alta_ref", ""), key=f"ref_adm_{rk_admin}")
                nom_dir = st.text_input("Nombre Completo:", value=st.session_state.get("admin_alta_nombre", ""), key=f"nom_adm_{rk_admin}")
                cip_dir = st.text_input("Código CIP:", value=st.session_state.get("admin_alta_cip", ""), key=f"cip_adm_{rk_admin}")
                meds_dir_edit = st.data_editor(st.session_state["df_alta_admin_directo"], num_rows="dynamic", key=f"editor_alta_admin_{rk_admin}", use_container_width=True)
                
                if st.button("✅ Dar de Alta Directamente", use_container_width=True):
                    r_str = ref_dir.strip()
                    n_str = nom_dir.strip()
                    c_str = cip_dir.strip()

                    if n_str and r_str and c_str:
                        df_final_dir = meds_dir_edit.copy()
                        df_final_dir['Ultima Entrega'] = ""
                        if 'Pedido' not in df_final_dir.columns: df_final_dir['Pedido'] = False
                        if 'Incidencia' not in df_final_dir.columns: df_final_dir['Incidencia'] = False
                        
                        shared_data["lista_pacientes"][f"{r_str} — {n_str}"] = {
                            "ref": r_str, "nombre": n_str, "cip": c_str, "hoja": r_str, "datos": df_final_dir
                        }
                        st.success("¡Paciente dado de alta correctamente!")
                        st.session_state["admin_alta_ref"] = ""
                        st.session_state["admin_alta_nombre"] = ""
                        st.session_state["admin_alta_cip"] = ""
                        st.session_state["df_alta_admin_directo"] = pd.DataFrame(columns=['Medicamento', 'CN', 'Posologia', 'Ultima Entrega'])
                        st.session_state["reset_alta_admin"] += 1
                        time.sleep(1.0)
                        st.rerun()
                    else: 
                        st.warning("⚠️ Rellene los campos obligatorios (Ref, Nombre y CIP).")

        with tabs_admin[1]:
            st.markdown("##### 1️⃣ Descargar Plantilla")
            st.info("Descarga este archivo Excel, rellénalo con todos los pacientes y sus medicamentos (una fila por medicamento), y súbelo en el paso 2.")
            
            df_template = pd.DataFrame(columns=["Ref_Paciente", "Nombre", "CIP", "Medicamento", "CN", "Posologia"])
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_template.to_excel(writer, index=False, sheet_name='Carga_Masiva')
            
            st.download_button(
                label="📥 Descargar Plantilla Excel", 
                data=buffer.getvalue(), 
                file_name="Plantilla_Carga_Masiva.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.markdown("---")
            st.markdown("##### 2️⃣ Subir Archivo Completado")
            archivo_subido = st.file_uploader("Selecciona el archivo Excel con los datos rellenos:", type=["xlsx"])
            
            if archivo_subido is not None:
                if st.button("🚀 Procesar Carga Masiva", use_container_width=True):
                    try:
                        df_carga = pd.read_excel(archivo_subido)
                        req_cols = ["Ref_Paciente", "Nombre", "CIP", "Medicamento", "CN", "Posologia"]
                        
                        cols_excel = [str(c).strip().lower() for c in df_carga.columns]
                        req_cols_lower = [c.lower() for c in req_cols]
                        
                        if all(rc in cols_excel for rc in req_cols_lower):
                            map_cols = {c: req_cols[req_cols_lower.index(str(c).strip().lower())] for c in df_carga.columns if str(c).strip().lower() in req_cols_lower}
                            df_carga = df_carga.rename(columns=map_cols)
                            
                            df_carga = df_carga.dropna(subset=['Ref_Paciente', 'Nombre'])
                            
                            grupos = df_carga.groupby('Ref_Paciente')
                            count_nuevos = 0
                            
                            for ref_val, grupo in grupos:
                                ref_str = str(ref_val).strip()
                                nombre_str = str(grupo['Nombre'].iloc[0]).strip()
                                cip_str = str(grupo['CIP'].iloc[0]).strip() if pd.notna(grupo['CIP'].iloc[0]) else ""
                                
                                df_meds = grupo[['Medicamento', 'CN', 'Posologia']].copy()
                                df_meds['Ultima Entrega'] = ""
                                df_meds['Pedido'] = False
                                df_meds['Incidencia'] = False
                                
                                df_meds['CN'] = df_meds['CN'].astype(str).str.replace(r'\.0$', '', regex=True)
                                
                                etiqueta = f"{ref_str} — {nombre_str}"
                                shared_data["lista_pacientes"][etiqueta] = {
                                    "ref": ref_str,
                                    "nombre": nombre_str,
                                    "cip": cip_str,
                                    "hoja": ref_str,
                                    "datos": df_meds.reset_index(drop=True)
                                }
                                count_nuevos += 1
                                
                            st.success(f"✅ ¡Carga masiva completada! Se han procesado {count_nuevos} pacientes exitosamente.")
                            time.sleep(2.5)
                            st.rerun()
                        else:
                            st.error(f"❌ El archivo no tiene las columnas correctas. Se requieren: {', '.join(req_cols)}")
                    except Exception as e:
                        st.error(f"❌ Error al procesar el archivo: {str(e)}")

        with tabs_admin[2]:
            st.markdown("##### 📋 Propuestas Pendientes de Enfermería:")
            pendientes_alta = [a for a in shared_data["solicitudes_alta"] if a["estado"] == "Pendiente"]
            if not pendientes_alta: st.info("No hay propuestas de alta pendientes.")
            else:
                for alta in pendientes_alta:
                    with st.container(border=True):
                        st.markdown(f"**Paciente:** {alta['nombre']} | **CIP:** {alta['cip']}")
                        
                        ref_asignado = st.text_input("Asignar/Modificar Código (Ref):", value=alta.get('ref', ''), key=f"ref_val_{alta['id']}")
                        
                        st.dataframe(alta["datos"], use_container_width=True, hide_index=True)
                        observacion_input = st.text_input("Observación si rechaza:", key=f"obs_{alta['id']}")
                        c_val, c_rec = st.columns(2)
                        with c_val:
                            if st.button(f"✅ Validar Propuesta", key=f"val_{alta['id']}", use_container_width=True):
                                if not ref_asignado.strip():
                                    st.error("⚠️ Debe asignar obligatoriamente un Código (Ref).")
                                else:
                                    alta["ref"] = ref_asignado.strip()
                                    shared_data["lista_pacientes"][f"{alta['ref']} — {alta['nombre']}"] = {"ref": alta["ref"], "nombre": alta["nombre"], "cip": alta["cip"], "hoja": alta["ref"], "datos": alta["datos"]}
                                    alta["estado"] = "Validada"
                                    st.success("¡Alta realizada!")
                                    time.sleep(1.0)
                                    st.rerun()
                        with c_rec:
                            if st.button(f"❌ Rechazar", key=f"rec_{alta['id']}", use_container_width=True):
                                if observacion_input.strip(): 
                                    alta["estado"] = "Rechazada"
                                    alta["observacion"] = observacion_input.strip()
                                    st.warning("Propuesta Rechazada.")
                                    time.sleep(1.0)
                                    st.rerun()
                                else: st.error("Escriba un motivo para el rechazo.")

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("⬅ Volver", use_container_width=True): st.session_state["pagina"] = "inicio"; st.rerun()

# ----------------------------------------------------
# LISTA DE PACIENTES Y BUSCADOR
# ----------------------------------------------------
elif st.session_state["pagina"] == "lista_pacientes":
    if "pacientes" not in permisos_usuario: st.stop()
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 800;'>PACIENTES</h2>", unsafe_allow_html=True)
    busqueda_paciente = st.text_input("🔍 Buscar paciente:", value="", placeholder="Escribe nombre o código...", key="input_busq_paciente")
    pacientes_filtrados = {pk: info for pk, info in lista_pacientes.items() if busqueda_paciente.lower() in pk.lower() or busqueda_paciente.lower() in info['nombre'].lower() or busqueda_paciente.lower() in info['ref'].lower() or busqueda_paciente.lower() in str(info.get('cip', '')).lower()}
    
    st.markdown("<br>", unsafe_allow_html=True)
    if not pacientes_filtrados: st.info("❌ No hay coincidencias.")
    else:
        for pk in list(pacientes_filtrados.keys()):
            if st.button(pk, key=f"p_{pk}", use_container_width=True): st.session_state["paciente_seleccionado_key"] = pk; st.session_state["pagina"] = "detalle_paciente"; st.rerun()
                
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("⬅ Volver", use_container_width=True): st.session_state["pagina"] = "inicio"; st.rerun()

# ----------------------------------------------------
# FICHA DE PACIENTE
# ----------------------------------------------------
elif st.session_state["pagina"] == "detalle_paciente":
    if "pacientes" not in permisos_usuario: st.stop()
    pk = st.session_state["paciente_seleccionado_key"]
    info = lista_pacientes.get(pk)
    
    if info:
        st.markdown(f"<h3 style='text-align: center;'>{info['nombre']} <br><span style='font-size: 13px; color: #64748b;'>Cód: {info['ref']} | CIP: {info['cip']}</span></h3>", unsafe_allow_html=True)
        
        if st.session_state["modo_incidencia"]:
            st.markdown("### 📝 Detalles de Incidencia")
            df_edit = st.data_editor(
                st.session_state["borrador_incidencias"],
                key="editor_borrador_incidencias",
                column_config={"Código Paciente": st.column_config.TextColumn("Código Paciente", disabled=True), "Nombre Paciente": st.column_config.TextColumn("Nombre Paciente", disabled=True), "Medicamento": st.column_config.TextColumn("Medicamento", disabled=True), "C.N.": st.column_config.TextColumn("C.N.", disabled=True), "Motivo": st.column_config.SelectboxColumn("Motivo", options=['Falta de receta electrónica', 'Modificar posología', 'Medicación adelantada', 'Lo consume?', 'Falta de abastecimiento', 'Otros'], required=True)},
                use_container_width=True, hide_index=True
            )
            st.session_state["borrador_incidencias"] = df_edit

            col_conf, col_canc = st.columns(2)
            with col_conf:
                if st.button("✅ Confirmar", use_container_width=True):
                    for _, row in st.session_state["borrador_incidencias"].iterrows():
                        shared_data["incidencias_activas"].append({"fecha": datetime.now().strftime("%d/%m/%Y %H:%M"), "ref_paciente": row["Código Paciente"], "paciente": row["Nombre Paciente"], "medicamento": row["Medicamento"], "cn": row["C.N."], "motivo": row.get("Motivo", ""), "observaciones": row.get("Observaciones", ""), "resuelta_por_enfermera": False})
                    info["datos"]["Incidencia"] = False
                    shared_data["lista_pacientes"][pk]["datos"] = info["datos"]
                    st.session_state["modo_incidencia"] = False; st.success("¡Incidencia enviada!"); time.sleep(1.5); st.rerun()
            with col_canc:
                if st.button("❌ Cancelar", use_container_width=True): st.session_state["modo_incidencia"] = False; st.rerun()
        else:
            df_pac = info["datos"].copy()
            cols = df_pac.columns.tolist()
            
            if rol_actual == "admin":
                for col_name in ['Incidencia', 'Pedido']:
                    if col_name in cols: cols.remove(col_name)
                new_cols = ['Pedido', 'Incidencia'] + cols
            else:
                for col_name in ['Incidencia', 'Pedido']:
                    if col_name in cols: cols.remove(col_name)
                new_cols = ['Pedido'] + cols
                
            new_cols = [c for c in new_cols if c in df_pac.columns]
            df_mostrar = df_pac[new_cols].copy()
            
            if 'Pedido' in df_mostrar.columns: df_mostrar['Pedido'] = df_mostrar['Pedido'].astype(bool)
            if 'Incidencia' in df_mostrar.columns: df_mostrar['Incidencia'] = df_mostrar['Incidencia'].astype(bool)

            def color_filas_paciente(row):
                if row.get('Incidencia', False) == True:
                    return ['background-color: #fecaca; color: #7f1d1d; font-weight: bold;'] * len(row) 
                elif row.get('Pedido', False) == True:
                    return ['background-color: #bbf7d0; color: #14532d; font-weight: bold;'] * len(row) 
                return [''] * len(row)

            styled_df = df_mostrar.style.apply(color_filas_paciente, axis=1)
            
            col_config_dict = {}
            for col in df_mostrar.columns:
                if col not in ['Pedido', 'Incidencia']: col_config_dict[col] = st.column_config.TextColumn(disabled=True)
                elif col == 'Pedido': col_config_dict[col] = st.column_config.CheckboxColumn("📦 Pedido", default=False)
                elif col == 'Incidencia': col_config_dict[col] = st.column_config.CheckboxColumn("⚠️ Incidencia", default=False)

            df_edited_result = st.data_editor(
                styled_df, 
                use_container_width=True, 
                hide_index=True, 
                key=f"editor_paciente_{pk}",
                column_config=col_config_dict
            )
            
            if not df_edited_result.equals(df_mostrar):
                for idx in range(len(df_edited_result)):
                    p_val = df_edited_result.loc[idx, 'Pedido'] if 'Pedido' in df_edited_result.columns else False
                    i_val = df_edited_result.loc[idx, 'Incidencia'] if 'Incidencia' in df_edited_result.columns else False
                    
                    if p_val and i_val:
                        old_p = df_mostrar.loc[idx, 'Pedido'] if 'Pedido' in df_mostrar.columns else False
                        old_i = df_mostrar.loc[idx, 'Incidencia'] if 'Incidencia' in df_mostrar.columns else False
                        if p_val and not old_p: df_edited_result.loc[idx, 'Incidencia'] = False
                        elif i_val and not old_i: df_edited_result.loc[idx, 'Pedido'] = False
                        else: df_edited_result.loc[idx, 'Incidencia'] = False

                for idx in range(len(df_edited_result)):
                    if 'Pedido' in df_edited_result.columns and 'Pedido' in info["datos"].columns: info["datos"].loc[idx, 'Pedido'] = df_edited_result.loc[idx, 'Pedido']
                    if rol_actual == "admin" and 'Incidencia' in df_edited_result.columns and 'Incidencia' in info["datos"].columns: info["datos"].loc[idx, 'Incidencia'] = df_edited_result.loc[idx, 'Incidencia']
                        
                shared_data["lista_pacientes"][pk]["datos"] = info["datos"]
                st.rerun()

            st.markdown("<br>", unsafe_allow_html=True)
            col_btn_ped, col_btn_inc = st.columns(2) if rol_actual == "admin" else (st.container(), None)

            with col_btn_ped:
                if rol_actual == "admin":
                    if st.button("📦 Enviar a Propuesta de Pedido", use_container_width=True):
                        df_pedidos = info["datos"][info["datos"]["Pedido"] == True]
                        if not df_pedidos.empty:
                            for _, row in df_pedidos.iterrows():
                                shared_data["solicitud_pedido"].append({"ref": info["ref"], "paciente": info["nombre"], "medicamento": row.get("Medicamento", ""), "cn": row.get("CN", ""), "posologia": row.get("Posologia", ""), "datamatrix": "", "lote": "", "caducidad": ""})
                            info["datos"]["Pedido"] = False
                            shared_data["lista_pacientes"][pk]["datos"] = info["datos"]
                            st.success("¡Enviado a Propuesta!"); time.sleep(1.5); st.rerun()
                        else: st.warning("Seleccione algún medicamento.")
                else: st.info("ℹ️ Solo el farmacéutico envía a propuesta desde aquí.")

            if rol_actual == "admin" and col_btn_inc:
                with col_btn_inc:
                    if st.button("⚠️ Enviar a Incidencias", use_container_width=True):
                        df_incidencias = info["datos"][info["datos"]["Incidencia"] == True]
                        if not df_incidencias.empty:
                            filas_borrador = []
                            for _, row in df_incidencias.iterrows():
                                filas_borrador.append({"Código Paciente": str(info.get("ref", "")), "Nombre Paciente": str(info.get("nombre", "")), "Medicamento": str(row.get("Medicamento", "")), "C.N.": str(row.get("CN", "")), "Motivo": "Falta de receta electrónica", "Observaciones": ""})
                            st.session_state["borrador_incidencias"] = pd.DataFrame(filas_borrador)
                            st.session_state["modo_incidencia"] = True; st.rerun()
                        else: st.warning("Seleccione algún medicamento.")

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("⬅ Volver a Lista"): st.session_state["modo_incidencia"] = False; st.session_state["pagina"] = "lista_pacientes"; st.rerun()

# ----------------------------------------------------
# PROPUESTA DE PEDIDO (ENFERMERÍA)
# ----------------------------------------------------
elif st.session_state["pagina"] == "seleccion_productos_enfermera":
    if "propuesta" not in permisos_usuario: st.stop()
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 800;'>📦 PROPUESTA (ENFERMERÍA)</h2>", unsafe_allow_html=True)
    
    if not shared_data["solicitud_pedido"]: 
        st.info("No hay propuestas pendientes.")
        if "propuesta_seleccion" in st.session_state: 
            del st.session_state["propuesta_seleccion"]
    else:
        st.markdown("<p style='text-align: center; color: #64748b; font-size: 14px;'>Toca el botón debajo de cada medicamento para seleccionarlo.</p>", unsafe_allow_html=True)
        
        if "propuesta_seleccion" not in st.session_state or len(st.session_state["propuesta_seleccion"]) != len(shared_data["solicitud_pedido"]):
            st.session_state["propuesta_seleccion"] = {i: False for i in range(len(shared_data["solicitud_pedido"]))}
        
        for i, item in enumerate(shared_data["solicitud_pedido"]):
            is_selected = st.session_state["propuesta_seleccion"].get(i, False)
            
            bg_color = "#bbf7d0" if is_selected else "#ffffff"
            border_color = "#22c55e" if is_selected else "#cbd5e1"
            text_color = "#14532d" if is_selected else "#1e293b"
            icon = "✅" if is_selected else "📦"
            
            st.markdown(f'''
            <div style="background-color: {bg_color}; border: 2px solid {border_color}; border-radius: 10px; padding: 12px; margin-bottom: 5px; color: {text_color}; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
                <div style="font-weight: 800; font-size: 15px;">{icon} {item.get('paciente', '')} <span style="font-size:12px; font-weight:normal; opacity:0.8;">(Ref: {item.get('ref', '')})</span></div>
                <div style="font-size: 14px; margin-top: 4px;">💊 <b>{item.get('medicamento', '')}</b> <span style="font-size:12px; opacity:0.8;">(CN: {item.get('cn', '')})</span></div>
                <div style="font-size: 13px; margin-top: 4px;">📝 Posología: {item.get('posologia', '')}</div>
            </div>
            ''', unsafe_allow_html=True)
            
            btn_label = "✅ SELECCIONADO (Tocar para desmarcar)" if is_selected else "👆 TOCAR PARA SELECCIONAR"
            
            if st.button(btn_label, key=f"btn_prop_{i}", use_container_width=True):
                st.session_state["propuesta_seleccion"][i] = not is_selected
                st.rerun()
                
            st.markdown("<div style='height: 15px;'></div>", unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🚀 Solicitar Pedido Definitivo", use_container_width=True):
            seleccionados = []
            restantes = []
            for i, item in enumerate(shared_data["solicitud_pedido"]):
                if st.session_state["propuesta_seleccion"].get(i, False):
                    seleccionados.append(item)
                else:
                    restantes.append(item)
            
            if seleccionados:
                for s in seleccionados:
                    shared_data["pedidos_definitivos"].append({
                        "ref": s.get("ref", ""), "paciente": s.get("paciente", ""), "medicamento": s.get("medicamento", ""),
                        "cn": s.get("cn", ""), "posologia": s.get("posologia", ""), "datamatrix": "", "lote": "", "caducidad": ""
                    })
                shared_data["solicitud_pedido"] = restantes
                del st.session_state["propuesta_seleccion"]
                st.success(f"¡{len(seleccionados)} medicamentos solicitados!")
                time.sleep(1.5); st.rerun()
            else: 
                st.warning("⚠️ Selecciona al menos un medicamento.")
            
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("⬅ Volver"): st.session_state["pagina"] = "inicio"; st.rerun()

elif st.session_state["pagina"] == "solicitud_pedido_admin":
    if rol_actual != "admin": st.stop()
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 800;'>📦 PROPUESTA (FARMACÉUTICO)</h2>", unsafe_allow_html=True)
    if shared_data["solicitud_pedido"]: 
        st.dataframe(pd.DataFrame(shared_data["solicitud_pedido"]), use_container_width=True, hide_index=True)
        st.info("ℹ️ Propuestas enviadas. Enfermería las revisará para generar el pedido definitivo.")
    else: st.info("No hay propuestas pendientes.")
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("⬅ Volver"): st.session_state["pagina"] = "inicio"; st.rerun()

# ----------------------------------------------------
# PEDIDOS DEFINITIVOS
# ----------------------------------------------------
elif st.session_state["pagina"] == "pedidos_definitivos_admin":
    if "pedidos_definitivos" not in permisos_usuario: st.stop()
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 800;'>🛒 PEDIDOS DEFINITIVOS</h2>", unsafe_allow_html=True)
    
    if not shared_data["pedidos_definitivos"]: 
        st.info("No hay pedidos definitivos pendientes.")
    else:
        st.markdown("##### 📋 Listado de Pedidos (Clic en la columna 'DataMatrix' y escanee):")
        
        df_defs = pd.DataFrame(shared_data["pedidos_definitivos"])
        
        columnas_requeridas = ['ref', 'paciente', 'medicamento', 'cn', 'posologia', 'datamatrix', 'lote', 'caducidad']
        for col in columnas_requeridas:
            if col not in df_defs.columns: 
                df_defs[col] = ""
                
        df_defs = df_defs.astype(str)
        df_defs = df_defs.replace(["nan", "None", "<NA>"], "")
            
        df_defs_edited = st.data_editor(
            df_defs, 
            use_container_width=True, 
            hide_index=True, 
            num_rows="dynamic", 
            key="editor_pedidos_definitivos", 
            column_config={
                "datamatrix": st.column_config.TextColumn("📷 Clic y Escanear (DataMatrix)"), 
                "lote": st.column_config.TextColumn("Lote de Fabricación"), 
                "caducidad": st.column_config.TextColumn("Caducidad"),
                "ref": st.column_config.TextColumn("Ref.", disabled=True),
                "paciente": st.column_config.TextColumn("Paciente", disabled=True),
                "medicamento": st.column_config.TextColumn("Medicamento", disabled=True),
                "cn": st.column_config.TextColumn("C.N.", disabled=True),
                "posologia": st.column_config.TextColumn("Posología", disabled=True)
            }
        )
        
        cambios = False
        for i in range(len(df_defs_edited)):
            nuevo_dm = str(df_defs_edited.iloc[i].get("datamatrix", "")).strip()
            viejo_dm = str(shared_data["pedidos_definitivos"][i].get("datamatrix", "")).strip() if i < len(shared_data["pedidos_definitivos"]) else ""
            
            if nuevo_dm and nuevo_dm != viejo_dm:
                parsed = traducir_datamatrix(nuevo_dm, BD_MEDICAMENTOS)
                df_defs_edited.at[i, "lote"] = parsed["lote"]
                df_defs_edited.at[i, "caducidad"] = parsed["caducidad"]
                cambios = True

        shared_data["pedidos_definitivos"] = df_defs_edited.to_dict(orient="records")
        
        if cambios:
            st.rerun()
        
        st.markdown("<br>", unsafe_allow_html=True)
        col_pdf, col_act = st.columns(2)
        with col_pdf:
            pdf = PDFAlbaran(orientation='L', unit='mm', format='A4')
            pdf.add_page()
            pdf.set_font("Arial", 'B', 14)
            pdf.cell(0, 10, limpiar_texto_pdf("FARMACIA VILLEGAS C.B."), ln=True, align='C')
            pdf.set_font("Arial", '', 11)
            pdf.cell(0, 6, limpiar_texto_pdf("C/ INDEPENDENCIA, 5 - TOMELLOSO"), ln=True, align='C')
            pdf.ln(10)
            pdf.set_font("Arial", 'B', 12)
            pdf.cell(0, 10, limpiar_texto_pdf(f"ALBARAN DE ENTREGA - Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M')}"), ln=True, align='L')
            pdf.ln(5)
            
            headers = ["Ref.", "Paciente", "Medicamento", "C.N.", "Posologia", "Lote", "Caducidad"]
            col_widths = [20, 60, 90, 25, 35, 30, 22] 
            align_list = ['C', 'L', 'L', 'C', 'C', 'C', 'C']
            
            rows_data = [[str(r.get('ref', '')), str(r.get('paciente', '')), str(r.get('medicamento', '')), str(r.get('cn', '')), str(r.get('posologia', '')), str(r.get('lote', '')), str(r.get('caducidad', ''))] for r in shared_data["pedidos_definitivos"]]
            
            dibujar_tabla_pdf(pdf, headers, rows_data, col_widths, align_list)
            st.download_button("📄 Imprimir Albarán (PDF)", data=pdf.output(dest='S').encode('latin1'), file_name="Albaran_Entrega.pdf", mime="application/pdf", use_container_width=True)
        
        with col_act:
            if st.button("📌 Actualizar y Limpiar", use_container_width=True):
                fecha_hoy = datetime.now().strftime("%d/%m/%Y")
                for item in shared_data["pedidos_definitivos"]:
                    for pk, p_info in shared_data["lista_pacientes"].items():
                        if str(p_info.get("ref", "")) == str(item.get("ref", "")):
                            df_p = p_info["datos"]
                            if 'Ultima Entrega' not in df_p.columns: df_p['Ultima Entrega'] = ""
                            mask = (df_p['CN'].astype(str).str.zfill(6) == str(item.get("cn", "")).zfill(6)) | (df_p['Medicamento'].astype(str) == str(item.get("medicamento", "")))
                            if mask.any(): df_p.loc[mask, 'Ultima Entrega'] = fecha_hoy
                            shared_data["lista_pacientes"][pk]["datos"] = df_p
                shared_data["pedidos_definitivos"] = []
                st.success("¡Actualizado y limpio!")
                time.sleep(1.5)
                st.rerun()
                
    if st.button("⬅ Volver"): st.session_state["pagina"] = "inicio"; st.rerun()

# ----------------------------------------------------
# INCIDENCIAS
# ----------------------------------------------------
elif st.session_state["pagina"] == "incidencias":
    if "incidencias" not in permisos_usuario: st.stop()
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 800;'>⚠️ PANEL DE INCIDENCIAS</h2>", unsafe_allow_html=True)
    
    if not shared_data["incidencias_activas"]: 
        st.info("No hay incidencias activas.")
    else:
        col_inst, col_print = st.columns([0.6, 0.4])
        with col_inst:
            st.markdown("<p style='text-align: left; color: #64748b; font-size: 14px;'>Toca el botón debajo de cada incidencia para marcarla como solucionada.</p>", unsafe_allow_html=True)
        with col_print:
            pdf_bytes = generar_pdf_incidencias(shared_data["incidencias_activas"])
            st.download_button("🖨️ Imprimir PDF (Agrupado por Paciente)", data=pdf_bytes, file_name="Reporte_Incidencias.pdf", mime="application/pdf", use_container_width=True)

        st.markdown("<hr style='margin-top: 0px; margin-bottom: 20px;'>", unsafe_allow_html=True)

        for i, item in enumerate(shared_data["incidencias_activas"]):
            is_resuelta = item.get("resuelta_por_enfermera", False)
            
            bg_color = "#ffd1d1" if is_resuelta else "#ffffff"
            border_color = "#ff7b7b" if is_resuelta else "#cbd5e1"
            text_color = "#5c1d1d" if is_resuelta else "#1e293b"
            icon = "🛑" if is_resuelta else "⚠️"
            text_decor = "line-through" if is_resuelta else "none"
            
            obs_html = f"<div style='font-size: 13px; margin-top: 4px;'>📝 Obs: {item.get('observaciones', '')}</div>" if item.get('observaciones', '') else ""
            aviso_avanzado = f"<div style='background-color: #ffebee; color: #b71c1c; padding: 6px; border-radius: 6px; font-weight: bold; font-size: 13px; margin-top: 6px;'>🔔 AVISO: La enfermera indica que esta incidencia está RESUELTA. Comprobar y validar.</div>" if is_resuelta else ""
            
            st.markdown(f'''
            <div style="background-color: {bg_color}; border: 2px solid {border_color}; border-radius: 10px; padding: 12px; margin-bottom: 5px; color: {text_color}; box-shadow: 0 2px 4px rgba(0,0,0,0.05); transition: all 0.3s ease;">
                <div style="font-weight: 800; font-size: 15px; text-decoration: {text_decor};">{icon} {item.get('paciente', '')} <span style="font-size:12px; font-weight:normal; opacity:0.8;">(Ref: {item.get('ref_paciente', '')})</span></div>
                <div style="font-size: 14px; margin-top: 4px; text-decoration: {text_decor};">💊 <b>{item.get('medicamento', '')}</b> <span style="font-size:12px; opacity:0.8;">(CN: {item.get('cn', '')})</span></div>
                <div style="font-size: 13px; margin-top: 4px; font-weight: bold; text-decoration: {text_decor};">🚨 Motivo: {item.get('motivo', '')}</div>
                {obs_html}
                {aviso_avanzado}
            </div>
            ''', unsafe_allow_html=True)
            
            st.caption(f"📅 {item.get('fecha', '')}")
            
            btn_label = "🛑 MARCADA COMO RESUELTA (Tocar para deshacer)" if is_resuelta else "👆 MARCAR COMO RESUELTA"
            
            if st.button(btn_label, key=f"btn_inc_{i}", use_container_width=True):
                item["resuelta_por_enfermera"] = not is_resuelta
                st.rerun()
                
            st.markdown("<div style='height: 15px;'></div>", unsafe_allow_html=True)
        
        if "validar_incidencias" in permisos_usuario:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("🗑️ Validar y Eliminar Marcadas", use_container_width=True):
                restantes = [item for item in shared_data["incidencias_activas"] if not item.get("resuelta_por_enfermera", False)]
                shared_data["incidencias_activas"] = restantes
                st.success("¡Incidencias validadas y eliminadas correctamente!")
                time.sleep(1.5)
                st.rerun()
        
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("⬅ Volver"): st.session_state["pagina"] = "inicio"; st.rerun()
    
elif st.session_state["pagina"] == "gestion_usuarios":
    if "usuarios" not in permisos_usuario:
        st.error("⛔ ACCESO RESTRINGIDO.")
        if st.button("⬅ Volver"): st.session_state["pagina"] = "inicio"; st.rerun()
        st.stop()
        
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 800;'>🔐 GESTIÓN DE USUARIOS</h2>", unsafe_allow_html=True)
    tabs = st.tabs(["👥 Cuentas", "🛡️ Roles"])
    with tabs[0]:
        df_u = pd.DataFrame([{"Usuario": k, "Clave": v["clave"], "Rol": v["rol"]} for k, v in shared_data["usuarios_sistema"].items()])
        df_u_edit = st.data_editor(df_u, use_container_width=True, num_rows="dynamic", key="editor_usuarios_sistema", column_config={"Rol": st.column_config.SelectboxColumn("Rol", options=list(shared_data["roles_sistema"].keys()), required=True), "Usuario": st.column_config.TextColumn("Usuario", required=True), "Clave": st.column_config.TextColumn("Clave", required=True)})
        if st.button("💾 Guardar Usuarios", use_container_width=True):
            nuevo_dict = {str(row["Usuario"]).strip(): {"clave": str(row["Clave"]), "rol": str(row["Rol"])} for _, row in df_u_edit.iterrows() if pd.notna(row["Usuario"]) and str(row["Usuario"]).strip()}
            shared_data["usuarios_sistema"] = nuevo_dict; st.success("Guardado."); time.sleep(1.5); st.rerun()

    with tabs[1]:
        modulos = ["pacientes", "altas", "bajas", "propuesta", "pedidos_definitivos", "incidencias", "validar_incidencias", "usuarios"]
        df_r = pd.DataFrame([{"Nombre del Rol": rol, **{m: m in perms for m in modulos}} for rol, perms in shared_data["roles_sistema"].items()])
        cols_config = {"Nombre del Rol": st.column_config.TextColumn("Rol", required=True)}
        for m in modulos: cols_config[m] = st.column_config.CheckboxColumn(m.replace("_", " ").title())
        df_r_edit = st.data_editor(df_r, use_container_width=True, num_rows="dynamic", key="editor_roles_sistema", column_config=cols_config)
        if st.button("💾 Guardar Permisos", use_container_width=True):
            nuevo_roles = {str(row["Nombre del Rol"]).strip(): [m for m in modulos if row.get(m, False) == True] for _, row in df_r_edit.iterrows() if pd.notna(row["Nombre del Rol"]) and str(row["Nombre del Rol"]).strip()}
            shared_data["roles_sistema"] = nuevo_roles; st.success("Guardado."); time.sleep(1.5); st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("⬅ Volver"): st.session_state["pagina"] = "inicio"; st.rerun()
