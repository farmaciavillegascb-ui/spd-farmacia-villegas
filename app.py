import streamlit as st
import pandas as pd
import os
import time
import uuid
import unicodedata
from fpdf import FPDF
from datetime import datetime

# Configuración de la página
st.set_page_config(
    page_title="SPD FARMACIA VILLEGAS",
    page_icon="💊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Estilos CSS
st.markdown("""
<style>
    .block-container { padding-top: 0.8rem !important; padding-bottom: 2rem !important; }
    .stApp { background-color: #f7f9fc; }
    #MainMenu {visibility: hidden;} footer {visibility: hidden;} header {visibility: hidden;}
    [data-testid="collapsedControl"] { display: none; }
    .dashboard-header { background: #ffffff; padding: 16px 24px; border-radius: 16px; box-shadow: 0 4px 20px rgba(0, 0, 0, 0.05); margin-bottom: 18px; border: 1px solid #e2e8f0; }
    .logo-container { display: flex; align-items: center; justify-content: center; gap: 10px; margin-bottom: 10px; }
    .logo-title { font-size: 22px; font-weight: 800; color: #1e293b; letter-spacing: 0.5px; }
    .status-bar { display: flex; justify-content: space-between; align-items: center; background: #f8fafc; padding: 8px 16px; border-radius: 10px; font-size: 13px; color: #475569; font-weight: 600; margin-bottom: 15px; border: 1px solid #e2e8f0; }
    [data-testid="column"] { display: flex !important; flex-direction: column !important; align-items: stretch !important; }
    [data-testid="column"] > div { display: flex !important; flex-direction: column !important; flex-grow: 1 !important; }
    div.stButton > button { width: 100% !important; height: 48px !important; border-radius: 10px !important; font-weight: 700 !important; font-size: 12px !important; background-color: #ffffff !important; color: #334155 !important; border: 2px solid #cbd5e1 !important; box-shadow: 0 2px 6px rgba(0,0,0,0.03) !important; transition: all 0.2s ease-in-out !important; flex-grow: 1 !important; }
    div.stButton > button:hover { background-color: #f1f5f9 !important; border-color: #0ea5e9 !important; color: #0284c7 !important; transform: translateY(-1px); }
    
    @keyframes pulse-subtle { 
        0% { transform: scale(1); box-shadow: 0 0 0 0 rgba(239, 68, 68, 0.6); } 
        50% { transform: scale(1.03); box-shadow: 0 0 0 12px rgba(239, 68, 68, 0); } 
        100% { transform: scale(1); box-shadow: 0 0 0 0 rgba(239, 68, 68, 0); } 
    }
    
    [data-testid="column"]:nth-child(1) div.stButton > button { border-color: #0ea5e9 !important; color: #0284c7 !important; background-color: #f0f9ff !important; }
    [data-testid="column"]:nth-child(2) div.stButton > button { border-color: #10b981 !important; color: #047857 !important; background-color: #ecfdf5 !important; }
    [data-testid="column"]:nth-child(4) div.stButton > button { border-color: #f59e0b !important; color: #d97706 !important; background-color: #fffbeb !important; }
</style>
""", unsafe_allow_html=True)

# ----------------------------------------------------
# CARGA DE BASE DE DATOS DE MEDICAMENTOS (EN CACHÉ)
# ----------------------------------------------------
@st.cache_data
def cargar_base_medicamentos():
    ruta_bd = "listado_de_medicamentos.xlsx"
    if not os.path.exists(ruta_bd):
        return {}
    try:
        df_bd = pd.read_excel(ruta_bd, sheet_name=0, usecols=['Cod. Nacional', 'Laboratorio', 'Presentación'])
        df_bd = df_bd.dropna(subset=['Cod. Nacional'])
        df_bd['CN'] = df_bd['Cod. Nacional'].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(6)
        
        partes = df_bd['Presentación'].astype(str).str.split(',', n=1, expand=True)
        df_bd['farmaco'] = partes[0].fillna('').astype(str).str.strip().str.slice(0, 45)
        if partes.shape[1] > 1:
            df_bd['tamano'] = partes[1].fillna('').astype(str).str.strip().str.slice(0, 25)
        else:
            df_bd['tamano'] = ""
            
        df_bd['marca'] = df_bd['Laboratorio'].fillna('').astype(str).str.slice(0, 25)
        
        df_bd = df_bd.set_index('CN')
        return df_bd[['marca', 'farmaco', 'tamano']].to_dict(orient='index')
    except Exception:
        return {}

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
            if 'Incidencia' not in df_hoja.columns: df_hoja['Incidencia'] = False
            
            pacientes_dict[etiqueta] = {
                "ref": ref_paciente, "nombre": nombre_paciente, "cip": cip_paciente, "hoja": hoja, "datos": df_hoja
            }
    return pacientes_dict

# ----------------------------------------------------
# MOTOR DE LECTURA DATAMATRIX
# ----------------------------------------------------
def traducir_datamatrix(raw_code, bd_medicamentos):
    res = {'marca': '', 'farmaco': '', 'tamano': '', 'cn': '', 'lote': '', 'caducidad': '', 'serie': ''}
    if not raw_code: return res
    
    clean = str(raw_code).replace('\x1D', '<GS>')
    
    try:
        if '712' in clean:
            idx = clean.find('712')
            if len(clean) >= idx + 9:
                res['cn'] = clean[idx+3 : idx+9].strip()
                
        if not res['cn'] and '01' in clean:
            idx = clean.find('01')
            if len(clean) >= idx + 16:
                gtin = clean[idx+2 : idx+16]
                if len(gtin) == 14:
                    res['cn'] = gtin[7:13]
                    
        if not res['cn'] and len(clean) >= 6:
            res['cn'] = clean[-6:].strip()

        if '17' in clean:
            idx = clean.find('17')
            if len(clean) >= idx + 8:
                cad_raw = clean[idx+2 : idx+8]
                if len(cad_raw) == 6 and cad_raw.isdigit():
                    yy, mm, dd = cad_raw[0:2], cad_raw[2:4], cad_raw[4:6]
                    if dd == '00': dd = '01'
                    res['caducidad'] = f"{dd}/{mm}/20{yy}"
        if not res['caducidad']: res['caducidad'] = "31/12/2028"

        idx_10 = clean.find('10')
        if idx_10 != -1:
            sub = clean[idx_10 + 2:]
            if '<GS>' in sub:
                res['lote'] = sub.split('<GS>')[0]
            else:
                for ai in ['21', '17', '712', '01']:
                    if ai in sub: sub = sub.split(ai)[0]
                res['lote'] = sub[:20].strip()
        if not res['lote']: res['lote'] = "LOTE01"

        idx_21 = clean.find('21')
        if idx_21 != -1:
            sub = clean[idx_21 + 2:]
            if '<GS>' in sub:
                res['serie'] = sub.split('<GS>')[0]
            else:
                for ai in ['17', '10', '712']:
                    if ai in sub: sub = sub.split(ai)[0]
                res['serie'] = sub[:20].strip()

    except Exception:
        res['cn'] = clean[-6:] if len(clean)>=6 else "000000"
        res['lote'] = "LOTE01"
        res['caducidad'] = "31/12/2028"
        
    cn_busqueda = res['cn'].zfill(6)
    if cn_busqueda in bd_medicamentos:
        datos = bd_medicamentos[cn_busqueda]
        res['marca'] = str(datos.get('marca', ''))
        res['farmaco'] = str(datos.get('farmaco', ''))
        res['tamano'] = str(datos.get('tamano', ''))
    else:
        res['farmaco'] = f"Medicamento (CN: {res['cn']})"
        
    return res

def limpiar_texto_pdf(texto):
    if not texto: return ""
    texto_str = str(texto)
    texto_str = texto_str.replace('ñ', 'n').replace('Ñ', 'N').replace('º', '.').replace('ª', '.').replace('—', '-')
    return unicodedata.normalize('NFKD', texto_str).encode('ascii', 'ignore').decode('ascii')

def dibujar_tabla_pdf(pdf, headers, rows_data, col_widths, align_list=None):
    header_height = 8
    pdf.set_font("Arial", 'B', 8)
    x_start = pdf.get_x()
    y_start = pdf.get_y()
    
    for i, h_text in enumerate(headers):
        w = col_widths[i]
        cx = pdf.get_x()
        cy = pdf.get_y()
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
            txt = str(val if val is not None else "")
            lines = pdf.multi_cell(w - 2, 3.5, txt, split_only=True)
            cell_lines.append(lines)
            if len(lines) > max_num_lines:
                max_num_lines = len(lines)
        
        row_height = max(6, max_num_lines * 3.5 + 2)
        
        xr_start = pdf.get_x()
        yr_start = pdf.get_y()
        
        if yr_start + row_height > 195:
            pdf.add_page()
            pdf.set_font("Arial", 'B', 8)
            hx = pdf.get_x()
            hy = pdf.get_y()
            for i, h_text in enumerate(headers):
                w = col_widths[i]
                cx = pdf.get_x()
                cy = pdf.get_y()
                pdf.rect(cx, cy, w, header_height)
                pdf.set_xy(cx, cy + 1.5)
                pdf.cell(w, 5, limpiar_texto_pdf(h_text), align='C', ln=0)
                pdf.set_xy(cx + w, hy)
            pdf.set_xy(hx, hy + header_height)
            pdf.set_font("Arial", '', 7.5)
            yr_start = pdf.get_y()
            xr_start = pdf.get_x()

        for i, lines in enumerate(cell_lines):
            w = col_widths[i]
            align = align_list[i] if align_list and i < len(align_list) else 'L'
            cx = pdf.get_x()
            cy = pdf.get_y()
            
            pdf.rect(cx, cy, w, row_height)
            current_y = cy + 1
            for line in lines:
                pdf.set_xy(cx + 1, current_y)
                pdf.cell(w - 2, 3.5, limpiar_texto_pdf(line), align=align, ln=0)
                current_y += 3.5
            
            pdf.set_xy(cx + w, yr_start)
            
        pdf.set_xy(xr_start, yr_start + row_height)

def generar_albaran_devolucion_pdf(nombre_paciente, ref_paciente, lista_devolucion):
    pdf = FPDF(orientation='L', unit='mm', format='A4') 
    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, limpiar_texto_pdf("FARMACIA VILLEGAS C.B. - ALBARAN DE DEVOLUCION"), ln=True, align='C')
    pdf.set_font("Arial", '', 11)
    pdf.cell(0, 6, limpiar_texto_pdf(f"Paciente: {nombre_paciente} (Ref: {ref_paciente})"), ln=True, align='L')
    pdf.ln(5)
    
    headers = ["Medicamento", "Descripcion", "CN", "Lote", "Caducidad", "Serie", "Restantes"]
    col_widths = [62, 78, 22, 38, 28, 25, 22]
    align_list = ['L', 'L', 'C', 'C', 'C', 'C', 'C']
    
    rows_data = []
    for row in lista_devolucion:
        rows_data.append([
            str(row.get('Medicamento', '')),
            str(row.get('Descripción', '')),
            str(row.get('CN', '')),
            str(row.get('Lote', '')),
            str(row.get('Caducidad', '')),
            str(row.get('Serie', '')),
            str(row.get('Pastillas restantes', '0'))
        ])
        
    dibujar_tabla_pdf(pdf, headers, rows_data, col_widths, align_list)
    return pdf.output(dest='S').encode('latin1')

@st.cache_resource
def get_shared_data():
    return {
        "lista_pacientes": cargar_datos_excel(),
        "solicitud_pedido": [], "pedidos_definitivos": [], "incidencias_activas": [], "solicitudes_alta": [], 
        "roles_sistema": {
            "admin": ["pacientes", "altas", "propuesta", "pedidos_def", "incidencias", "usuarios", "roles"],
            "enfermera": ["pacientes", "altas", "propuesta", "incidencias"]
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

TIEMPO_EXPIRACION = 3600  

def obtener_parametro_url(nombre):
    try: return st.query_params.get(nombre)
    except AttributeError:
        try:
            params = st.experimental_get_query_params()
            val = params.get(nombre)
            return val[0] if val else None
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
    if time.time() - datos_sesion["ultimo_acceso"] < TIEMPO_EXPIRACION:
        shared_data["sesiones_activas"][token_url]["ultimo_acceso"] = time.time()
        st.session_state["usuario_autenticado"] = datos_sesion["usuario"]
        st.session_state["rol_usuario"] = datos_sesion["rol"]
    else:
        del shared_data["sesiones_activas"][token_url]
        limpiar_parametros_url()
        st.session_state["usuario_autenticado"] = None
        st.session_state["rol_usuario"] = None

# LOGIN
if st.session_state["usuario_autenticado"] is None:
    col1, col2, col3 = st.columns([1, 1.5, 1])
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

# CABECERA: 8 COLUMNAS (INCLUYE SYNC)
st.markdown('<div class="dashboard-header">', unsafe_allow_html=True)
st.markdown('<div class="logo-container"><span style="font-size: 24px;">💊</span><span class="logo-title">SPD FARMACIA VILLEGAS</span></div>', unsafe_allow_html=True)

rol_actual = st.session_state["rol_usuario"]
st.markdown(f'<div class="status-bar"><span>Sistema activo <span style="color: #22c55e; font-size: 16px;">●</span></span><span>Usuario: <b>{st.session_state["usuario_autenticado"]}</b> ({rol_actual.upper()})</span></div>', unsafe_allow_html=True)

num_ped = len(shared_data["pedidos_definitivos"])
num_prop = len(shared_data["solicitud_pedido"])
num_inc = len(shared_data["incidencias_activas"])

if rol_actual == "admin":
    txt_ped = f"PEDIDOS ({num_ped})" if num_ped > 0 else "PEDIDOS"
    alert_ped = (num_ped > 0)
else:
    txt_ped = f"PROPUESTA ({num_prop})" if num_prop > 0 else "PROPUESTA"
    alert_ped = (num_prop > 0)

txt_inc = f"INCIDENCIAS ({num_inc})" if num_inc > 0 else "INCIDENCIAS"
alert_inc = (num_inc > 0)

if alert_ped:
    st.markdown("""<style>[data-testid="column"]:nth-child(5) div.stButton > button { background: linear-gradient(135deg, #ef4444, #dc2626) !important; color: white !important; border: 2px solid #fca5a5 !important; animation: pulse-subtle 1.8s infinite; }</style>""", unsafe_allow_html=True)
if alert_inc:
    st.markdown("""<style>[data-testid="column"]:nth-child(6) div.stButton > button { background: linear-gradient(135deg, #ef4444, #dc2626) !important; color: white !important; border: 2px solid #fca5a5 !important; animation: pulse-subtle 1.8s infinite; }</style>""", unsafe_allow_html=True)

col_inicio, col_sync, col_alta, col_baja, col_ped, col_inc, col_user, col_logout = st.columns([1,1,1,1,1.2,1.2,1,1], gap="small")

with col_inicio:
    if st.button("🏠 INICIO", key="btn_hdr_inicio", use_container_width=True):
        st.session_state["pagina"] = "inicio"; st.rerun()

with col_sync:
    if st.button("🔄 SYNC", key="btn_hdr_sync", use_container_width=True):
        st.rerun()  

with col_alta:
    if st.button("ALTA", key="btn_hdr_alta", use_container_width=True):
        st.session_state["pagina"] = "alta_paciente"; st.rerun()

with col_baja:
    if st.button("🚨 BAJAS", key="btn_hdr_bajas", use_container_width=True):
        st.session_state["pagina"] = "baja_paciente"; st.rerun()

with col_ped:
    if st.button(txt_ped, key="btn_hdr_ped", use_container_width=True):
        st.session_state["pagina"] = "pedidos_definitivos_admin" if rol_actual == "admin" else "seleccion_productos_enfermera"
        st.rerun()

with col_inc:
    if st.button(txt_inc, key="btn_hdr_inc", use_container_width=True):
        st.session_state["pagina"] = "incidencias"; st.rerun()

with col_user:
    if st.button("USUARIOS", key="btn_hdr_usu", use_container_width=True):
        st.session_state["pagina"] = "gestion_usuarios"; st.rerun()

with col_logout:
    if st.button("SALIR", key="btn_hdr_out", use_container_width=True):
        if token_url in shared_data["sesiones_activas"]: del shared_data["sesiones_activas"][token_url]
        limpiar_parametros_url(); st.session_state["usuario_autenticado"] = None; st.session_state["pagina"] = "inicio"; st.rerun()

st.markdown('</div>', unsafe_allow_html=True)

# VISTAS PRINCIPALES
if st.session_state["pagina"] == "inicio":
    c1, c2 = st.columns(2, gap="large")
    with c1:
        st.markdown(f'<div style="background: #eff6ff; padding: 22px; border-radius: 16px; border: 1px solid #bfdbfe; margin-bottom: 10px;"><h4 style="color: #1e3a8a; margin-top: 0;">👤 PACIENTE ({len(shared_data["lista_pacientes"])} activos)</h4><p style="color: #334155; font-size: 14px; margin-bottom: 0;">Listado completo de pacientes y tratamientos.</p></div>', unsafe_allow_html=True)
        if st.button("🧓 **ENTRAR A PACIENTES**", use_container_width=True): st.session_state["pagina"] = "lista_pacientes"; st.rerun()
    with c2:
        st.markdown(f'<div style="background: #fef2f2; padding: 22px; border-radius: 16px; border: 1px solid #fecaca; margin-bottom: 10px;"><h4 style="color: #7f1d1d; margin-top: 0;">⚠️ PANEL INCIDENCIAS ({len(shared_data["incidencias_activas"])} abiertas)</h4><p style="color: #334155; font-size: 14px; margin-bottom: 0;">Gestión de incidencias y recetas.</p></div>', unsafe_allow_html=True)
        if st.button("📋 **VER PANEL DE INCIDENCIAS**", use_container_width=True): st.session_state["pagina"] = "incidencias"; st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)
    c3, c4 = st.columns(2, gap="large")
    with c3:
        st.markdown(f'<div style="background: #fefce8; padding: 22px; border-radius: 16px; border: 1px solid #fef08a; margin-bottom: 10px;"><h4 style="color: #713f12; margin-top: 0;">🚚 PROPUESTA DE PEDIDO ({len(shared_data["solicitud_pedido"])})</h4><p style="color: #334155; font-size: 14px; margin-bottom: 0;">Seguimiento de propuestas con enfermería.</p></div>', unsafe_allow_html=True)
        if st.button("📦 **VER PROPUESTA DE PEDIDO**", use_container_width=True): st.session_state["pagina"] = "solicitud_pedido_admin" if rol_actual == "admin" else "seleccion_productos_enfermera"; st.rerun()
    with c4:
        st.markdown(f'<div style="background: #f0fdf4; padding: 22px; border-radius: 16px; border: 1px solid #bbf7d0; margin-bottom: 10px;"><h4 style="color: #14532d; margin-top: 0;">📄 PEDIDO DEFINITIVO ({len(shared_data["pedidos_definitivos"])})</h4><p style="color: #334155; font-size: 14px; margin-bottom: 0;">Validación con DataMatrix y albaranes PDF.</p></div>', unsafe_allow_html=True)
        if st.button("🛒 **VER PEDIDOS DEFINITIVOS**", use_container_width=True): st.session_state["pagina"] = "pedidos_definitivos_admin" if rol_actual == "admin" else "seleccion_productos_enfermera"; st.rerun()

# ----------------------------------------------------
# MÓDULO: BAJAS DE PACIENTES CON/SIN DEVOLUCIÓN
# ----------------------------------------------------
elif st.session_state["pagina"] == "baja_paciente":
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 800;'>👴 GESTIÓN DE BAJAS DE PACIENTES</h2>", unsafe_allow_html=True)
    
    if "baja_paso" not in st.session_state: st.session_state["baja_paso"] = "seleccion_paciente"
    if "paciente_baja_obj" not in st.session_state: st.session_state["paciente_baja_obj"] = None
    if "df_devolucion" not in st.session_state: 
        st.session_state["df_devolucion"] = pd.DataFrame(columns=['Medicamento', 'Descripción', 'CN', 'Lote', 'Caducidad', 'Serie', 'Pastillas restantes'])

    if st.session_state["baja_paso"] == "seleccion_paciente":
        st.markdown("##### 🔍 Seleccione el paciente que causa baja:")
        paciente_seleccionado = st.selectbox("Paciente activo:", [""] + list(shared_data["lista_pacientes"].keys()), key="select_baja_paciente")
        
        c_btn1, c_btn2 = st.columns(2)
        with c_btn1:
            if st.button("Continuar", use_container_width=True) and paciente_seleccionado:
                info_p = shared_data["lista_pacientes"][paciente_seleccionado]
                st.session_state["paciente_baja_obj"] = {"etiqueta": paciente_seleccionado, "nombre": info_p["nombre"], "ref": info_p["ref"]}
                st.session_state["baja_paso"] = "elegir_modalidad"
                st.rerun()
        with c_btn2:
            if st.button("⬅ Volver al Menú Principal", use_container_width=True):
                st.session_state["pagina"] = "inicio"; st.rerun()

    elif st.session_state["baja_paso"] == "elegir_modalidad":
        pac = st.session_state["paciente_baja_obj"]
        st.info(f"👤 Paciente seleccionado: **{pac['nombre']}** (Ref: {pac['ref']})")
        st.markdown("##### Seleccione cómo desea procesar esta baja:")
        
        col_m1, col_m2, col_m3 = st.columns(3, gap="medium")
        with col_m1:
            if st.button("❌ Baja SIN Devolución", use_container_width=True):
                if pac["etiqueta"] in shared_data["lista_pacientes"]:
                    del shared_data["lista_pacientes"][pac["etiqueta"]]
                st.success("¡Baja procesada correctamente sin devolución!")
                st.session_state["baja_paso"] = "seleccion_paciente"
                time.sleep(1.5); st.rerun()
        with col_m2:
            if st.button("📦 Baja CON Devolución", use_container_width=True):
                st.session_state["baja_paso"] = "pantalla_devolucion"
                st.rerun()
        with col_m3:
            if st.button("↩ Cancelar / Cambiar Paciente", use_container_width=True):
                st.session_state["baja_paso"] = "seleccion_paciente"
                st.rerun()

    elif st.session_state["baja_paso"] == "pantalla_devolucion":
        pac = st.session_state["paciente_baja_obj"]
        st.info(f"📦 Registrando Devolución de Medicación para: **{pac['nombre']}** (Ref: {pac['ref']})")
        
        st.markdown("##### 📷 Escaneo de Código DataMatrix")
        with st.form("form_escanear_dm_baja", clear_on_submit=True):
            cadena_dm = st.text_input("Escanee o introduzca la cadena continua del DataMatrix:")
            submit_scan = st.form_submit_button("Añadir Medicamento a la Devolución", use_container_width=True)
            
            if submit_scan and cadena_dm:
                parsed = traducir_datamatrix(cadena_dm, BD_MEDICAMENTOS)
                nuevo_reg = {
                    'Medicamento': parsed['farmaco'],
                    'Descripción': f"{parsed['marca']} - {parsed['tamano']}".strip(" -"),
                    'CN': parsed['cn'],
                    'Lote': parsed['lote'],
                    'Caducidad': parsed['caducidad'],
                    'Serie': parsed['serie'],
                    'Pastillas restantes': 0
                }
                st.session_state["df_devolucion"] = pd.concat([st.session_state["df_devolucion"], pd.DataFrame([nuevo_reg])], ignore_index=True)
                st.success(f"✅ ¡{parsed['farmaco']} añadido correctamente!")

        if not st.session_state["df_devolucion"].empty:
            if st.button("🗑️ Eliminar último escaneo", use_container_width=False):
                st.session_state["df_devolucion"] = st.session_state["df_devolucion"].iloc[:-1]
                st.rerun()
            
            st.markdown("##### 📋 Listado de Devolución (Edite la columna 'Pastillas restantes')")
            st.session_state["df_devolucion"] = st.data_editor(
                st.session_state["df_devolucion"], 
                use_container_width=True, 
                hide_index=True,
                num_rows="dynamic",
                key="editor_devoluciones_baja",
                disabled=['Medicamento', 'Descripción', 'CN', 'Lote', 'Caducidad', 'Serie'] 
            )
            st.markdown("---")
            col_pdf, col_fin = st.columns(2)
            with col_pdf:
                pdf_bytes = generar_albaran_devolucion_pdf(pac['nombre'], pac['ref'], st.session_state["df_devolucion"].to_dict(orient="records"))
                st.download_button("📄 Imprimir PDF", data=pdf_bytes, file_name=f"Devolucion_{pac['ref']}.pdf", mime="application/pdf", use_container_width=True)
            with col_fin:
                if st.button("💾 Finalizar Baja y Guardar", use_container_width=True):
                    if pac["etiqueta"] in shared_data["lista_pacientes"]: 
                        del shared_data["lista_pacientes"][pac["etiqueta"]]
                    st.session_state["baja_paso"] = "seleccion_paciente"
                    st.session_state["df_devolucion"] = pd.DataFrame(columns=['Medicamento', 'Descripción', 'CN', 'Lote', 'Caducidad', 'Serie', 'Pastillas restantes'])
                    st.success("¡Baja registrada con éxito!"); time.sleep(1.5); st.rerun()

        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("⬅ Volver al Menú de Modalidades"):
            st.session_state["baja_paso"] = "elegir_modalidad"; st.rerun()

# ----------------------------------------------------
# RESTO DE MÓDULOS
# ----------------------------------------------------
elif st.session_state["pagina"] == "alta_paciente":
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 800;'>👴 ALTA DE PACIENTE</h2>", unsafe_allow_html=True)
    if "df_alta_cargado" not in st.session_state:
        st.session_state["df_alta_cargado"] = pd.DataFrame(columns=['Medicamento', 'CN', 'Posologia', 'Ultima Entrega'])
    with st.form("form_alta"):
        nuevo_nombre = st.text_input("Nombre completo:")
        nuevo_cip = st.text_input("CIP:")
        nueva_ref = st.text_input("Referencia:", value="NUEVO")
        meds_editadas = st.data_editor(st.session_state["df_alta_cargado"], num_rows="dynamic", key="editor_alta_paciente", use_container_width=True)
        if st.form_submit_button("Guardar y Dar de Alta"):
            if nuevo_nombre.strip():
                df_final = meds_editadas.copy()
                df_final['Ultima Entrega'] = ""
                etiqueta = f"{nueva_ref} — {nuevo_nombre.strip()}"
                shared_data["lista_pacientes"][etiqueta] = {
                    "ref": nueva_ref, "nombre": nuevo_nombre.strip(), "cip": nuevo_cip, "hoja": nueva_ref, "datos": df_final
                }
                st.success("¡Paciente dado de alta!")
                st.session_state["df_alta_cargado"] = pd.DataFrame(columns=['Medicamento', 'CN', 'Posologia', 'Ultima Entrega'])
                time.sleep(1.5); st.rerun()
    if st.button("⬅ Volver"): st.session_state["pagina"] = "inicio"; st.rerun()

elif st.session_state["pagina"] == "lista_pacientes":
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 800;'>PACIENTES</h2>", unsafe_allow_html=True)
    for pk in list(lista_pacientes.keys()):
        if st.button(pk, key=f"p_{pk}", use_container_width=True):
            st.session_state["paciente_seleccionado_key"] = pk
            st.session_state["pagina"] = "detalle_paciente"; st.rerun()
    if st.button("⬅ Volver"): st.session_state["pagina"] = "inicio"; st.rerun()

elif st.session_state["pagina"] == "detalle_paciente":
    pk = st.session_state["paciente_seleccionado_key"]
    info = lista_pacientes.get(pk)
    
    if info:
        st.markdown(f"<h3 style='text-align: center;'>Paciente: {info['nombre']}</h3>", unsafe_allow_html=True)
        
        if st.session_state["modo_incidencia"]:
            st.markdown("### 📝 Completar Detalles de Incidencia")
            st.info("Rellene el motivo y las observaciones para cada incidencia y confirme el envío.")
            
            df_edit = st.data_editor(
                st.session_state["borrador_incidencias"],
                key="editor_borrador_incidencias",
                column_config={
                    "Código Paciente": st.column_config.TextColumn("Código Paciente", disabled=True),
                    "Nombre Paciente": st.column_config.TextColumn("Nombre Paciente", disabled=True),
                    "Medicamento": st.column_config.TextColumn("Medicamento", disabled=True),
                    "C.N.": st.column_config.TextColumn("C.N.", disabled=True),
                    "Motivo": st.column_config.SelectboxColumn(
                        "Motivo",
                        options=[
                            'Falta de receta electrónica', 
                            'Modificar posología', 
                            'Medicación adelantada', 
                            'Lo consume?', 
                            'Falta de abastecimiento', 
                            'Otros'
                        ],
                        required=True
                    ),
                    "Observaciones": st.column_config.TextColumn("Observaciones")
                },
                use_container_width=True,
                hide_index=True
            )
            st.session_state["borrador_incidencias"] = df_edit

            col_conf, col_canc = st.columns(2)
            with col_conf:
                if st.button("✅ Confirmar y Enviar a Enfermería", use_container_width=True):
                    for _, row in st.session_state["borrador_incidencias"].iterrows():
                        incidencia = {
                            "fecha": datetime.now().strftime("%d/%m/%Y %H:%M"),
                            "ref_paciente": row["Código Paciente"],
                            "paciente": row["Nombre Paciente"],
                            "medicamento": row["Medicamento"],
                            "cn": row["C.N."],
                            "motivo": row.get("Motivo", ""),
                            "observaciones": row.get("Observaciones", ""),
                            "estado": "Pendiente"
                        }
                        shared_data["incidencias_activas"].append(incidencia)
                    
                    info["datos"]["Incidencia"] = False
                    shared_data["lista_pacientes"][pk]["datos"] = info["datos"]
                    
                    st.session_state["modo_incidencia"] = False
                    st.success("¡Incidencias enviadas correctamente a la bandeja!")
                    time.sleep(1.5); st.rerun()
            with col_canc:
                if st.button("❌ Cancelar", use_container_width=True):
                    st.session_state["modo_incidencia"] = False
                    st.rerun()
        else:
            df_pac = info["datos"].copy()
            cols = df_pac.columns.tolist()
            for col_name in ['Incidencia', 'Pedido']:
                if col_name in cols: cols.remove(col_name)
            new_cols = ['Pedido', 'Incidencia'] + cols
            new_cols = [c for c in new_cols if c in df_pac.columns]
            
            df_mostrar = df_pac[new_cols].copy()

            def color_filas_paciente(row):
                if row.get('Incidencia', False) == True:
                    return ['background-color: #fecaca; color: #7f1d1d;'] * len(row) 
                elif row.get('Pedido', False) == True:
                    return ['background-color: #bbf7d0; color: #14532d;'] * len(row) 
                return [''] * len(row)

            styled_df = df_mostrar.style.apply(color_filas_paciente, axis=1)

            df_edited_result = st.data_editor(styled_df, use_container_width=True, hide_index=True, key=f"editor_paciente_{pk}")
            
            cambio_realizado = False
            for idx in range(len(df_edited_result)):
                p_val = df_edited_result.loc[idx, 'Pedido'] if 'Pedido' in df_edited_result.columns else False
                i_val = df_edited_result.loc[idx, 'Incidencia'] if 'Incidencia' in df_edited_result.columns else False
                
                if p_val and i_val:
                    old_p = df_mostrar.loc[idx, 'Pedido'] if 'Pedido' in df_mostrar.columns else False
                    old_i = df_mostrar.loc[idx, 'Incidencia'] if 'Incidencia' in df_mostrar.columns else False
                    
                    if p_val and not old_p:
                        df_edited_result.loc[idx, 'Incidencia'] = False
                        cambio_realizado = True
                    elif i_val and not old_i:
                        df_edited_result.loc[idx, 'Pedido'] = False
                        cambio_realizado = True
                    else:
                        df_edited_result.loc[idx, 'Incidencia'] = False
                        cambio_realizado = True

            if cambio_realizado:
                st.rerun()

            info["datos"] = df_edited_result
            shared_data["lista_pacientes"][pk]["datos"] = info["datos"]

            st.markdown("<br>", unsafe_allow_html=True)
            col_btn_ped, col_btn_inc = st.columns(2)
            
            with col_btn_ped:
                if st.button("📦 Enviar a Propuesta de Pedido", use_container_width=True):
                    df_pedidos = info["datos"][info["datos"]["Pedido"] == True]
                    if not df_pedidos.empty:
                        for _, row in df_pedidos.iterrows():
                            item = {
                                "seleccion_enfermera": False,
                                "ref": info["ref"],
                                "paciente": info["nombre"],
                                "medicamento": row.get("Medicamento", ""),
                                "cn": row.get("CN", ""),
                                "posologia": row.get("Posologia", ""),
                                "datamatrix": "",
                                "lote": "",
                                "caducidad": ""
                            }
                            shared_data["solicitud_pedido"].append(item)
                        info["datos"]["Pedido"] = False
                        shared_data["lista_pacientes"][pk]["datos"] = info["datos"]
                        st.success("¡Medicamentos enviados a la bandeja de Pedidos!")
                        time.sleep(1.5); st.rerun()
                    else:
                        st.warning("Marca la casilla 'Pedido' en algún medicamento primero.")

            with col_btn_inc:
                if st.button("⚠️ Enviar a Incidencias", use_container_width=True):
                    df_incidencias = info["datos"][info["datos"]["Incidencia"] == True]
                    if not df_incidencias.empty:
                        filas_borrador = []
                        for _, row in df_incidencias.iterrows():
                            filas_borrador.append({
                                "Código Paciente": str(info.get("ref", "")),
                                "Nombre Paciente": str(info.get("nombre", "")),
                                "Medicamento": str(row.get("Medicamento", "")),
                                "C.N.": str(row.get("CN", "")),
                                "Motivo": "Falta de receta electrónica",
                                "Observaciones": ""
                            })
                        st.session_state["borrador_incidencias"] = pd.DataFrame(filas_borrador)
                        st.session_state["modo_incidencia"] = True
                        st.rerun()
                    else:
                        st.warning("Marca la casilla 'Incidencia' en algún medicamento primero.")

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("⬅ Volver a Lista"): 
        st.session_state["modo_incidencia"] = False
        st.session_state["pagina"] = "lista_pacientes"
        st.rerun()

elif st.session_state["pagina"] == "seleccion_productos_enfermera":
    st.markdown("<h2 style='text-align: center;'>📦 PROPUESTA DE PEDIDO</h2>", unsafe_allow_html=True)
    if not shared_data["solicitud_pedido"]: st.info("No hay propuestas.")
    else:
        df_sol = pd.DataFrame(shared_data["solicitud_pedido"])
        if 'seleccion_enfermera' in df_sol.columns:
            cols = df_sol.columns.tolist()
            cols.remove('seleccion_enfermera')
            cols.insert(0, 'seleccion_enfermera')
            df_sol = df_sol[cols]
            
        df_edited = st.data_editor(
            df_sol, 
            use_container_width=True, 
            hide_index=True, 
            num_rows="dynamic", 
            key="editor_enfermera_propuesta",
            column_config={
                "seleccion_enfermera": st.column_config.CheckboxColumn("Seleccionar", default=False)
            }
        )
        shared_data["solicitud_pedido"] = df_edited.to_dict(orient="records")
        
        if st.button("🚀 Solicitar Pedido Definitivo"):
            sel = [i for i in shared_data["solicitud_pedido"] if i.get("seleccion_enfermera")]
            for it in sel: shared_data["pedidos_definitivos"].append(it)
            shared_data["solicitud_pedido"] = [i for i in shared_data["solicitud_pedido"] if not i.get("seleccion_enfermera")]
            st.success("Enviado al farmacéutico."); time.sleep(1.5); st.rerun()
            
    if st.button("⬅ Volver"): st.session_state["pagina"] = "inicio"; st.rerun()

elif st.session_state["pagina"] == "solicitud_pedido_admin":
    st.markdown("<h2 style='text-align: center;'>📦 PROPUESTA (ENVIADO A ENFERMERÍA)</h2>", unsafe_allow_html=True)
    if shared_data["solicitud_pedido"]: 
        df_sol = pd.DataFrame(shared_data["solicitud_pedido"])
        if 'seleccion_enfermera' in df_sol.columns:
            cols = df_sol.columns.tolist()
            cols.remove('seleccion_enfermera')
            cols.insert(0, 'seleccion_enfermera')
            df_sol = df_sol[cols]
        st.dataframe(df_sol, use_container_width=True, hide_index=True)
    else: st.info("Vacío.")
    if st.button("⬅ Volver"): st.session_state["pagina"] = "inicio"; st.rerun()

elif st.session_state["pagina"] == "pedidos_definitivos_admin":
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 800;'>🛒 PEDIDOS DEFINITIVOS Y ESCÁNER DATAMATRIX</h2>", unsafe_allow_html=True)
    if not shared_data["pedidos_definitivos"]: 
        st.info("No hay pedidos definitivos pendientes en este momento.")
    else:
        st.markdown("##### 📥 Escanee el código DataMatrix de cada medicamento para rellenar Lote y Caducidad automáticamente:")
        with st.form("form_pedidos_dm", clear_on_submit=True):
            cadena_dm_pedido = st.text_input("Cadena DataMatrix escaneada:")
            btn_escaneo = st.form_submit_button("🔍 Procesar y Asignar al Siguiente Medicamento", use_container_width=True)
            
            if btn_escaneo and cadena_dm_pedido:
                parsed_ped = traducir_datamatrix(cadena_dm_pedido, BD_MEDICAMENTOS)
                asignado = False
                for item in shared_data["pedidos_definitivos"]:
                    if not item.get("datamatrix") or item.get("datamatrix") == "":
                        item["datamatrix"] = cadena_dm_pedido
                        item["lote"] = parsed_ped['lote']
                        item["caducidad"] = parsed_ped['caducidad']
                        asignado = True
                        break
                if asignado:
                    st.success(f"✅ Medicamento asignado correctamente (Lote: {parsed_ped['lote']}, Cad: {parsed_ped['caducidad']})")
                else:
                    st.warning("⚠️ Todos los medicamentos de la lista ya tienen un DataMatrix asignado.")

        st.markdown("---")
        st.markdown("##### 📋 Listado de Pedidos Definitivos:")
        df_defs = pd.DataFrame(shared_data["pedidos_definitivos"])
        
        for col in ['datamatrix', 'lote', 'caducidad']:
            if col not in df_defs.columns: df_defs[col] = ""

        df_defs_edited = st.data_editor(
            df_defs, 
            use_container_width=True, 
            hide_index=True, 
            num_rows="dynamic",
            key="editor_pedidos_definitivos",
            column_config={
                "datamatrix": st.column_config.TextColumn("DataMatrix", help="Cadena escaneada"),
                "lote": st.column_config.TextColumn("Lote"),
                "caducidad": st.column_config.TextColumn("Caducidad")
            }
        )
        shared_data["pedidos_definitivos"] = df_defs_edited.to_dict(orient="records")
        
        st.markdown("<br>", unsafe_allow_html=True)
        col_pdf, col_act = st.columns(2)
        with col_pdf:
            pdf = FPDF(orientation='L', unit='mm', format='A4') 
            pdf.add_page()
            pdf.set_font("Arial", 'B', 14)
            pdf.cell(0, 10, limpiar_texto_pdf("FARMACIA VILLEGAS C.B."), ln=True, align='C')
            pdf.set_font("Arial", '', 11)
            pdf.cell(0, 6, limpiar_texto_pdf("C/ INDEPENDENCIA, 5 - TOMELLOSO"), ln=True, align='C')
            pdf.ln(10)
            pdf.set_font("Arial", 'B', 12)
            fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M")
            pdf.cell(0, 10, limpiar_texto_pdf(f"ALBARAN DE ENTREGA - Fecha: {fecha_actual}"), ln=True, align='L')
            pdf.ln(5)
            
            headers = ["Ref.", "Paciente", "Medicamento", "C.N.", "Posologia", "DataMatrix", "Lote", "Caducidad"]
            col_widths = [20, 55, 75, 22, 25, 45, 20, 20] 
            align_list = ['C', 'L', 'L', 'C', 'C', 'L', 'C', 'C']
            
            rows_data = []
            for row in shared_data["pedidos_definitivos"]:
                rows_data.append([
                    str(row.get('ref', '')),
                    str(row.get('paciente', '')),
                    str(row.get('medicamento', '')),
                    str(row.get('cn', '')),
                    str(row.get('posologia', '')),
                    str(row.get('datamatrix', '')),
                    str(row.get('lote', '')),
                    str(row.get('caducidad', ''))
                ])
                
            dibujar_tabla_pdf(pdf, headers, rows_data, col_widths, align_list)
            pdf_bytes = pdf.output(dest='S').encode('latin1')
            st.download_button("📄 Imprimir Albarán de Entrega (PDF)", data=pdf_bytes, file_name="Albaran_Entrega.pdf", mime="application/pdf", use_container_width=True)
        with col_act:
            if st.button("📌 Actualizar Última Entrega y Limpiar", use_container_width=True):
                fecha_hoy = datetime.now().strftime("%d/%m/%Y")
                # Actualizar la columna 'Ultima Entrega' en cada paciente correspondiente
                for item in shared_data["pedidos_definitivos"]:
                    ref_item = str(item.get("ref", ""))
                    med_item = str(item.get("medicamento", ""))
                    cn_item = str(item.get("cn", ""))
                    
                    for pk, p_info in shared_data["lista_pacientes"].items():
                        if str(p_info.get("ref", "")) == ref_item:
                            df_p = p_info["datos"]
                            if 'Ultima Entrega' not in df_p.columns:
                                df_p['Ultima Entrega'] = ""
                            
                            mask = (df_p['CN'].astype(str).str.zfill(6) == str(cn_item).zfill(6)) | (df_p['Medicamento'].astype(str) == med_item)
                            if mask.any():
                                df_p.loc[mask, 'Ultima Entrega'] = fecha_hoy
                            shared_data["lista_pacientes"][pk]["datos"] = df_p

                shared_data["pedidos_definitivos"] = []
                st.success("¡Fechas de última entrega actualizadas en las fichas y lista limpiada!")
                time.sleep(1.5); st.rerun()
    if st.button("⬅ Volver al Menú"): st.session_state["pagina"] = "inicio"; st.rerun()

elif st.session_state["pagina"] == "incidencias":
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 800;'>⚠️ PANEL DE INCIDENCIAS</h2>", unsafe_allow_html=True)
    if shared_data["incidencias_activas"]:
        df_inc = pd.DataFrame(shared_data["incidencias_activas"])
        
        if 'Solucionada' not in df_inc.columns:
            df_inc.insert(0, 'Solucionada', False)
            
        df_edit_inc = st.data_editor(df_inc, use_container_width=True, hide_index=True, key="editor_panel_incidencias")
        
        if rol_actual == "admin":
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("🗑️ Validar y Eliminar Incidencias Marcadas", use_container_width=True):
                restantes = df_edit_inc[df_edit_inc['Solucionada'] != True]
                if 'Solucionada' in restantes.columns:
                    restantes = restantes.drop(columns=['Solucionada'])
                shared_data["incidencias_activas"] = restantes.to_dict(orient="records")
                st.success("¡Incidencias validadas y eliminadas correctamente!")
                time.sleep(1.5); st.rerun()
    else: 
        st.info("No hay incidencias activas en este momento.")
        
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("⬅ Volver"): st.session_state["pagina"] = "inicio"; st.rerun()
    
elif st.session_state["pagina"] == "gestion_usuarios":
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 800;'>🔐 GESTIÓN DE USUARIOS Y ROLES</h2>", unsafe_allow_html=True)
    if st.button("⬅ Volver"): st.session_state["pagina"] = "inicio"; st.rerun()
