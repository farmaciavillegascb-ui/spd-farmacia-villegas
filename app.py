import streamlit as st
import pandas as pd
import os
import time
import uuid
import gdown
from fpdf import FPDF
from datetime import datetime

# Configuración de la página
st.set_page_config(
    page_title="SPD FARMACIA VILLEGAS",
    page_icon="💊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Estilos CSS definitivos
st.markdown("""
<style>
    .block-container {
        padding-top: 0.8rem !important;
        padding-bottom: 2rem !important;
    }
    .stApp {
        background-color: #f7f9fc;
    }
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    [data-testid="collapsedControl"] { display: none; }
    
    .dashboard-header {
        background: #ffffff;
        padding: 16px 24px;
        border-radius: 16px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.05);
        margin-bottom: 18px;
        border: 1px solid #e2e8f0;
    }
    .logo-container {
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 10px;
        margin-bottom: 10px;
    }
    .logo-title {
        font-size: 22px;
        font-weight: 800;
        color: #1e293b;
        letter-spacing: 0.5px;
    }
    .status-bar {
        display: flex;
        justify-content: space-between;
        align-items: center;
        background: #f8fafc;
        padding: 8px 16px;
        border-radius: 10px;
        font-size: 13px;
        color: #475569;
        font-weight: 600;
        margin-bottom: 15px;
        border: 1px solid #e2e8f0;
    }
    [data-testid="column"] {
        display: flex !important;
        flex-direction: column !important;
        align-items: stretch !important;
    }
    [data-testid="column"] > div {
        display: flex !important;
        flex-direction: column !important;
        flex-grow: 1 !important;
    }
    .alerta-wrapper {
        display: flex !important;
        flex-direction: column !important;
        flex-grow: 1 !important;
        width: 100% !important;
    }
    .alerta-wrapper > div {
        display: flex !important;
        flex-direction: column !important;
        flex-grow: 1 !important;
    }
    div.stButton > button {
        width: 100% !important;
        height: 48px !important;
        border-radius: 10px !important;
        font-weight: 700 !important;
        font-size: 12px !important;
        background-color: #ffffff !important;
        color: #334155 !important;
        border: 2px solid #cbd5e1 !important;
        box-shadow: 0 2px 6px rgba(0,0,0,0.03) !important;
        transition: all 0.2s ease-in-out !important;
        flex-grow: 1 !important;
    }
    div.stButton > button:hover {
        background-color: #f1f5f9 !important;
        border-color: #0ea5e9 !important;
        color: #0284c7 !important;
        transform: translateY(-1px);
    }
    @keyframes pulse-subtle {
        0% { transform: scale(1); box-shadow: 0 0 0 0 rgba(239, 68, 68, 0.4); }
        70% { transform: scale(1.02); box-shadow: 0 0 0 10px rgba(239, 68, 68, 0); }
        100% { transform: scale(1); box-shadow: 0 0 0 0 rgba(239, 68, 68, 0); }
    }
    .alerta-activa button {
        background: linear-gradient(135deg, #ef4444, #dc2626) !important;
        color: white !important;
        border: 2px solid #fca5a5 !important;
        font-weight: bold !important;
        animation: pulse-subtle 1.8s infinite;
    }
</style>
""", unsafe_allow_html=True)

EXCEL_PATH = "Tratamientos_Por_Paciente.xlsx"

def cargar_datos_excel():
    if not os.path.exists(EXCEL_PATH):
        return {}
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
                "ref": ref_paciente,
                "nombre": nombre_paciente,
                "cip": cip_paciente,
                "hoja": hoja,
                "datos": df_hoja
            }
    return pacientes_dict

def traducir_datamatrix(raw_code):
    """Decodificador inteligente GS1 DataMatrix farmacéutico."""
    res = {
        'marca': 'Genérico Farmacia',
        'farmaco': 'Medicamento Genérico',
        'dosificacion': '100mg',
        'tamano': '28 comp',
        'cn': '',
        'lote': '',
        'caducidad': ''
    }
    if not raw_code:
        return res
    
    clean = str(raw_code).replace('\x1D', '').replace('(', '').replace(')', '').replace(']', '').strip()
    
    try:
        if '01' in clean:
            idx = clean.find('01')
            if len(clean) >= idx + 16:
                gtin = clean[idx+2 : idx+16]
                if len(gtin) == 14:
                    res['cn'] = gtin[7:13]
                    res['farmaco'] = f"Fármaco CN {res['cn']}"
        elif len(clean) >= 6 and clean[:6].isdigit():
            res['cn'] = clean[:6]
            res['farmaco'] = f"Fármaco CN {res['cn']}"
        else:
            res['cn'] = clean[:6] if len(clean) >= 6 else "123456"
            res['farmaco'] = clean[:25] if len(clean) > 0 else "Medicamento"

        if '17' in clean:
            idx = clean.find('17')
            if len(clean) >= idx + 8:
                cad_raw = clean[idx+2 : idx+8]
                if len(cad_raw) == 6:
                    yy, mm, dd = cad_raw[0:2], cad_raw[2:4], cad_raw[4:6]
                    if dd == '00': dd = '01'
                    res['caducidad'] = f"{dd}/{mm}/20{yy}"
        if not res['caducidad']:
            res['caducidad'] = "31/12/2028"

        if '10' in clean:
            idx = clean.find('10')
            lote_val = clean[idx+2:]
            for ai in ['21', '17', '30', '11']:
                if ai in lote_val:
                    lote_val = lote_val.split(ai)[0]
            res['lote'] = lote_val[:20].strip()
        if not res['lote']:
            res['lote'] = "LOTE01"

    except Exception:
        res['cn'] = clean[:6] if len(clean)>=6 else "123456"
        res['lote'] = "LOTE01"
        res['caducidad'] = "31/12/2028"
        
    return res

def generar_albaran_pdf(lista_pedidos):
    pdf = FPDF(orientation='L') 
    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, "FARMACIA VILLEGAS C.B.", ln=True, align='C')
    pdf.set_font("Arial", '', 11)
    pdf.cell(0, 6, "C/ INDEPENDENCIA, 5 — TOMELLOSO", ln=True, align='C')
    pdf.ln(10)
    pdf.set_font("Arial", 'B', 12)
    fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M")
    pdf.cell(0, 10, f"ALBARÁN DE ENTREGA - Fecha: {fecha_actual}", ln=True, align='L')
    pdf.ln(5)
    
    pdf.set_font("Arial", 'B', 9)
    col_widths = [15, 55, 65, 20, 20, 45, 25, 25] 
    headers = ["Ref.", "Paciente", "Medicamento", "C.N.", "Posología", "DataMatrix", "Lote", "Caducidad"]
    for i in range(len(headers)):
        pdf.cell(col_widths[i], 8, headers[i], border=1, align='C')
    pdf.ln()
    
    pdf.set_font("Arial", '', 8)
    for row in lista_pedidos:
        pdf.cell(col_widths[0], 8, str(row.get('ref', ''))[:8], border=1)
        pdf.cell(col_widths[1], 8, str(row.get('paciente', ''))[:35], border=1)
        pdf.cell(col_widths[2], 8, str(row.get('medicamento', ''))[:45], border=1)
        pdf.cell(col_widths[3], 8, str(row.get('cn', ''))[:10], border=1, align='C')
        pdf.cell(col_widths[4], 8, str(row.get('posologia', ''))[:12], border=1, align='C')
        pdf.cell(col_widths[5], 8, str(row.get('datamatrix', ''))[:35], border=1)
        pdf.cell(col_widths[6], 8, str(row.get('lote', ''))[:15], border=1, align='C')
        pdf.cell(col_widths[7], 8, str(row.get('caducidad', ''))[:12], border=1, align='C')
        pdf.ln()
    return pdf.output(dest='S').encode('latin1')

def generar_albaran_devolucion_pdf(nombre_paciente, ref_paciente, lista_devolucion):
    pdf = FPDF(orientation='L') 
    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, "FARMACIA VILLEGAS C.B. - ALBARÁN DE DEVOLUCIÓN", ln=True, align='C')
    pdf.set_font("Arial", '', 11)
    pdf.cell(0, 6, f"Paciente: {nombre_paciente} (Ref: {ref_paciente})", ln=True, align='L')
    pdf.ln(5)
    
    pdf.set_font("Arial", 'B', 9)
    col_widths = [35, 55, 25, 25, 20, 25, 20, 25] 
    headers = ["Marca", "Fármaco", "Dosificación", "Tamaño", "CN", "Lote", "Caducidad", "Restantes"]
    for i in range(len(headers)):
        pdf.cell(col_widths[i], 8, headers[i], border=1, align='C')
    pdf.ln()
    
    pdf.set_font("Arial", '', 8)
    for row in lista_devolucion:
        pdf.cell(col_widths[0], 8, str(row.get('Marca', ''))[:20], border=1)
        pdf.cell(col_widths[1], 8, str(row.get('Fármaco', ''))[:30], border=1)
        pdf.cell(col_widths[2], 8, str(row.get('Dosificación', ''))[:15], border=1, align='C')
        pdf.cell(col_widths[3], 8, str(row.get('Tamaño envase', ''))[:15], border=1, align='C')
        pdf.cell(col_widths[4], 8, str(row.get('CN', ''))[:10], border=1, align='C')
        pdf.cell(col_widths[5], 8, str(row.get('Lote', ''))[:12], border=1, align='C')
        pdf.cell(col_widths[6], 8, str(row.get('Caducidad', ''))[:10], border=1, align='C')
        pdf.cell(col_widths[7], 8, str(row.get('Pastillas restantes', '0'))[:8], border=1, align='C')
        pdf.ln()
    return pdf.output(dest='S').encode('latin1')

def obtener_parametro_url(nombre):
    try:
        return st.query_params.get(nombre)
    except AttributeError:
        try:
            params = st.experimental_get_query_params()
            val = params.get(nombre)
            return val[0] if val else None
        except Exception:
            return None

def establecer_parametro_url(nombre, valor):
    try:
        st.query_params[nombre] = valor
    except AttributeError:
        try:
            st.experimental_set_query_params(**{nombre: valor})
        except Exception:
            pass

def limpiar_parametros_url():
    try:
        st.query_params.clear()
    except AttributeError:
        try:
            st.experimental_set_query_params()
        except Exception:
            pass

@st.cache_resource
def get_shared_data():
    return {
        "lista_pacientes": cargar_datos_excel(),
        "solicitud_pedido": [],       
        "pedidos_definitivos": [],
        "incidencias_activas": [],
        "solicitudes_alta": [], 
        "solicitudes_baja": [], 
        "roles_sistema": {
            "admin": ["pacientes", "altas", "bajas", "propuesta", "pedidos_def", "incidencias", "usuarios", "roles"],
            "enfermera": ["pacientes", "altas", "bajas", "propuesta", "incidencias"]
        },
        "usuarios_sistema": {
            "farmaciaB": {"clave": "farmaciaB2026", "rol": "admin"},
            "farmaciaR": {"clave": "farmaciaR2026", "rol": "admin"},
            "FarmaciaC": {"clave": "FarmaciaC2026", "rol": "admin"},
            "FarmaciaA": {"clave": "FarmaciaA2026", "rol": "admin"},
            "FarmaciasCH": {"clave": "FarmaciasCH2026", "rol": "admin"},
            "Enfermera1": {"clave": "enfermera12026", "rol": "enfermera"},
            "Enfermera2": {"clave": "enfermera22026", "rol": "enfermera"}
        },
        "sesiones_activas": {}        
    }

shared_data = get_shared_data()
lista_pacientes = shared_data["lista_pacientes"]
USUARIOS_VALIDOS = shared_data["usuarios_sistema"]
ROLES_VALIDOS = shared_data["roles_sistema"]

if "usuario_autenticado" not in st.session_state:
    st.session_state["usuario_autenticado"] = None
if "rol_usuario" not in st.session_state:
    st.session_state["rol_usuario"] = None
if "pagina" not in st.session_state:
    st.session_state["pagina"] = "inicio"
if "paciente_seleccionado_key" not in st.session_state:
    st.session_state["paciente_seleccionado_key"] = None

TIEMPO_EXPIRACION = 600  
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

def tiene_permiso(rol, modulo):
    if rol not in ROLES_VALIDOS:
        return False
    return modulo in ROLES_VALIDOS[rol]

# LOGIN
if st.session_state["usuario_autenticado"] is None:
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("<h2 style='text-align: center; color: #1e293b;'>💊 SPD FARMACIA VILLEGAS</h2>", unsafe_allow_html=True)
        with st.form("form_login"):
            st.markdown("#### Portal de Acceso Profesional")
            usuario_input = st.text_input("Usuario")
            clave_input = st.text_input("Clave de acceso", type="password")
            submit_login = st.form_submit_button("Iniciar Sesión", use_container_width=True)
            if submit_login:
                if usuario_input in USUARIOS_VALIDOS and USUARIOS_VALIDOS[usuario_input]["clave"] == clave_input:
                    nuevo_token = str(uuid.uuid4())
                    shared_data["sesiones_activas"][nuevo_token] = {
                        "usuario": usuario_input,
                        "rol": USUARIOS_VALIDOS[usuario_input]["rol"],
                        "ultimo_acceso": time.time()
                    }
                    establecer_parametro_url("session_token", nuevo_token)
                    st.session_state["usuario_autenticado"] = usuario_input
                    st.session_state["rol_usuario"] = USUARIOS_VALIDOS[usuario_input]["rol"]
                    st.session_state["pagina"] = "inicio"
                    st.rerun()
                else:
                    st.error("❌ Usuario o clave incorrectos.")
    st.stop()

# CABECERA SUPERIOR
st.markdown('<div class="dashboard-header">', unsafe_allow_html=True)
st.markdown("""
    <div class="logo-container">
        <span style="font-size: 24px;">💊</span>
        <span class="logo-title">SPD FARMACIA VILLEGAS</span>
    </div>
""", unsafe_allow_html=True)

rol_actual = st.session_state["rol_usuario"]
st.markdown(f"""
    <div class="status-bar">
        <span>Sistema activo <span style="color: #22c55e; font-size: 16px;">●</span></span>
        <span>Usuario: <b>{st.session_state['usuario_autenticado']}</b> ({rol_actual.upper()})</span>
    </div>
""", unsafe_allow_html=True)

col_alta, col_baja, col_ped, col_inc, col_user, col_logout = st.columns(6, gap="small")

with col_alta:
    if tiene_permiso(rol_actual, "altas"):
        num_altas = len(shared_data["solicitudes_alta"]) if rol_actual == "admin" else 0
        lbl_alta = f"ALTA ({num_altas})" if num_altas > 0 else "ALTA"
        if num_altas > 0: st.markdown('<div class="alerta-wrapper alerta-activa">', unsafe_allow_html=True)
        if st.button(lbl_alta, key="btn_alta", use_container_width=True): 
            st.session_state["pagina"] = "alta_paciente"; st.rerun()
        if num_altas > 0: st.markdown('</div>', unsafe_allow_html=True)
    else: st.markdown('<div style="height: 48px;"></div>', unsafe_allow_html=True)
        
with col_baja:
    if tiene_permiso(rol_actual, "bajas"):
        num_bajas = len(shared_data["solicitudes_baja"]) if rol_actual == "admin" else 0
        lbl_baja = f"BAJA ({num_bajas})" if num_bajas > 0 else "BAJA"
        if num_bajas > 0: st.markdown('<div class="alerta-wrapper alerta-activa">', unsafe_allow_html=True)
        if st.button(lbl_baja, key="btn_baja", use_container_width=True): 
            st.session_state["pagina"] = "baja_paciente"; st.rerun()
        if num_bajas > 0: st.markdown('</div>', unsafe_allow_html=True)
    else: st.markdown('<div style="height: 48px;"></div>', unsafe_allow_html=True)

with col_ped:
    if rol_actual == "admin" and tiene_permiso(rol_actual, "pedidos_def"):
        num_defs = len(shared_data["pedidos_definitivos"])
        lbl_ped = "PEDIDOS"
        if num_defs > 0: st.markdown('<div class="alerta-wrapper alerta-activa">', unsafe_allow_html=True)
        if st.button(lbl_ped, key="btn_pedidos_admin", use_container_width=True):
            st.session_state["pagina"] = "pedidos_definitivos_admin"; st.rerun()
        if num_defs > 0: st.markdown('</div>', unsafe_allow_html=True)
    elif tiene_permiso(rol_actual, "propuesta"):
        num_sol = len(shared_data["solicitud_pedido"])
        lbl_sol = "PROPUESTA"
        if num_sol > 0: st.markdown('<div class="alerta-wrapper alerta-activa">', unsafe_allow_html=True)
        if st.button(lbl_sol, key="btn_sol_enf_top", use_container_width=True):
            st.session_state["pagina"] = "seleccion_productos_enfermera"; st.rerun()
        if num_sol > 0: st.markdown('</div>', unsafe_allow_html=True)
    else: st.markdown('<div style="height: 48px;"></div>', unsafe_allow_html=True)
            
with col_inc:
    if tiene_permiso(rol_actual, "incidencias"):
        num_inc = len(shared_data["incidencias_activas"])
        lbl_inc = f"INCIDENCIAS ({num_inc})" if num_inc > 0 else "INCIDENCIAS"
        if num_inc > 0: st.markdown('<div class="alerta-wrapper alerta-activa">', unsafe_allow_html=True)
        if st.button(lbl_inc, key="btn_incidencias", use_container_width=True):
            st.session_state["pagina"] = "incidencias"; st.rerun()
        if num_inc > 0: st.markdown('</div>', unsafe_allow_html=True)
    else: st.markdown('<div style="height: 48px;"></div>', unsafe_allow_html=True)

with col_user:
    if tiene_permiso(rol_actual, "usuarios") or tiene_permiso(rol_actual, "roles"):
        if st.button("USUARIOS", key="btn_usuarios_top", use_container_width=True):
            st.session_state["pagina"] = "gestion_usuarios"; st.rerun()
    else: st.markdown('<div style="height: 48px;"></div>', unsafe_allow_html=True)

with col_logout:
    if st.button("SALIR", key="btn_logout_top", use_container_width=True):
        t_url = obtener_parametro_url("session_token")
        if t_url and t_url in shared_data["sesiones_activas"]: del shared_data["sesiones_activas"][t_url]
        limpiar_parametros_url()
        st.session_state["usuario_autenticado"] = None
        st.session_state["rol_usuario"] = None
        st.session_state["pagina"] = "inicio"
        st.rerun()

st.markdown('</div>', unsafe_allow_html=True)

# VISTAS PRINCIPALES
if st.session_state["pagina"] == "inicio":
    total_pacientes = len(shared_data["lista_pacientes"])
    total_incidencias = len(shared_data["incidencias_activas"])
    total_pedidos_def = len(shared_data["pedidos_definitivos"])
    total_solicitudes = len(shared_data["solicitud_pedido"])

    c_card1, c_card2 = st.columns(2, gap="large")
    with c_card1:
        if tiene_permiso(rol_actual, "pacientes"):
            st.markdown(f"""
                <div style="background: #eff6ff; padding: 22px; border-radius: 16px; border: 1px solid #bfdbfe; margin-bottom: 10px;">
                    <h4 style="color: #1e3a8a; margin-top: 0;">👤 PACIENTE ({total_pacientes} activos)</h4>
                    <p style="color: #334155; font-size: 14px; margin-bottom: 0;">Listado completo de pacientes y tratamientos.</p>
                </div>
            """, unsafe_allow_html=True)
            if st.button("🧓 **ENTRAR A PACIENTES**", key="button_pacientes", use_container_width=True):
                st.session_state["pagina"] = "lista_pacientes"; st.rerun()
            
    with c_card2:
        if tiene_permiso(rol_actual, "incidencias"):
            st.markdown(f"""
                <div style="background: #fef2f2; padding: 22px; border-radius: 16px; border: 1px solid #fecaca; margin-bottom: 10px;">
                    <h4 style="color: #7f1d1d; margin-top: 0;">⚠️ PANEL INCIDENCIAS ({total_incidencias} abiertas)</h4>
                    <p style="color: #334155; font-size: 14px; margin-bottom: 0;">Gestión de incidencias y recetas.</p>
                </div>
            """, unsafe_allow_html=True)
            if st.button("📋 **VER PANEL DE INCIDENCIAS**", key="button_incidencias", use_container_width=True):
                st.session_state["pagina"] = "incidencias"; st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)
    c_card3, c_card4 = st.columns(2, gap="large")
    with c_card3:
        if tiene_permiso(rol_actual, "propuesta") or tiene_permiso(rol_actual, "pedidos_def"):
            st.markdown(f"""
                <div style="background: #fefce8; padding: 22px; border-radius: 16px; border: 1px solid #fef08a; margin-bottom: 10px;">
                    <h4 style="color: #713f12; margin-top: 0;">🚚 PROPUESTA DE PEDIDO ({total_solicitudes})</h4>
                    <p style="color: #334155; font-size: 14px; margin-bottom: 0;">Seguimiento de propuestas con enfermería.</p>
                </div>
            """, unsafe_allow_html=True)
            if rol_actual == "admin":
                if st.button("📦 **VER PROPUESTA DE PEDIDO**", key="button_solicitudes", use_container_width=True):
                    st.session_state["pagina"] = "solicitud_pedido_admin"; st.rerun()
            else:
                if st.button("📦 **BANDEJA DE PROPUESTAS DE PEDIDO**", key="button_bandeja", use_container_width=True):
                    st.session_state["pagina"] = "seleccion_productos_enfermera"; st.rerun()
                
    with c_card4:
        if tiene_permiso(rol_actual, "pedidos_def") or tiene_permiso(rol_actual, "propuesta"):
            st.markdown(f"""
                <div style="background: #f0fdf4; padding: 22px; border-radius: 16px; border: 1px solid #bbf7d0; margin-bottom: 10px;">
                    <h4 style="color: #14532d; margin-top: 0;">📄 PEDIDO DEFINITIVO ({total_pedidos_def})</h4>
                    <p style="color: #334155; font-size: 14px; margin-bottom: 0;">Validación con DataMatrix y albaranes PDF.</p>
                </div>
            """, unsafe_allow_html=True)
            if rol_actual == "admin":
                if st.button("🛒 **VER PEDIDOS DEFINITIVOS**", key="button_definitivos", use_container_width=True):
                    st.session_state["pagina"] = "pedidos_definitivos_admin"; st.rerun()
            else:
                if st.button("📦 **SELECCIÓN DE ENFERMERÍA**", key="button_sel_enf", use_container_width=True):
                    st.session_state["pagina"] = "seleccion_productos_enfermera"; st.rerun()

elif st.session_state["pagina"] == "alta_paciente":
    if not tiene_permiso(rol_actual, "altas"): st.error("No tienes permiso."); st.stop()
        
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 800;'>👴 ALTA DE PACIENTE</h2>", unsafe_allow_html=True)
    es_admin_rol = (rol_actual == "admin")
    
    if es_admin_rol and shared_data["solicitudes_alta"]:
        st.warning("⚠️ Solicitudes de alta pendientes de Enfermería:")
        for idx, sol in enumerate(shared_data["solicitudes_alta"]):
            c1, c2, c3 = st.columns([3, 1, 1])
            c1.write(f"**Nombre:** {sol['nombre']} | **CIP:** {sol.get('cip', 'N/A')}")
            if c2.button("✅ Aprobar Alta", key=f"apr_alta_{idx}"):
                etiqueta = f"{sol['ref']} — {sol['nombre']}"
                shared_data["lista_pacientes"][etiqueta] = {
                    "ref": sol['ref'], "nombre": sol['nombre'], "cip": sol['cip'], "hoja": sol['ref'], "datos": sol['datos']
                }
                shared_data["solicitudes_alta"].pop(idx)
                st.success(f"¡Paciente {sol['nombre']} dado de alta!")
                time.sleep(1); st.rerun()
            if c3.button("❌ Rechazar", key=f"rec_alta_{idx}"):
                shared_data["solicitudes_alta"].pop(idx); st.rerun()
        st.markdown("---")

    st.subheader("Carga de Medicación")
    if "df_alta_cargado" not in st.session_state:
        st.session_state["df_alta_cargado"] = pd.DataFrame(columns=['Medicamento', 'CN', 'Posologia', 'Ultima Entrega'])

    tab_local, tab_drive = st.tabs(["📁 Archivo local", "☁️ Google Drive"])
    with tab_local:
        archivo_subido = st.file_uploader("Sube un Excel (.xlsx)", type=["xlsx", "xls"], key="up_excel")
        if archivo_subido is not None:
            try:
                df_leido = pd.read_excel(archivo_subido)
                df_leido.columns = [str(c).strip().capitalize() for c in df_leido.columns]
                st.session_state["df_alta_cargado"] = pd.DataFrame({
                    'Medicamento': df_leido.iloc[:, 0],
                    'CN': df_leido.iloc[:, 1],
                    'Posologia': df_leido.iloc[:, 2],
                    'Ultima Entrega': ""
                })
                st.success("¡Archivo leído con éxito!")
            except Exception as e:
                st.error(f"Error: {e}")

    with tab_drive:
        gdrive_url = st.text_input("Enlace público de Google Drive:")
        if st.button("📥 Descargar de Drive"):
            if gdrive_url:
                try:
                    output_path = "temp_gdrive.xlsx"
                    gdown.download(gdrive_url, output_path, quiet=False)
                    if os.path.exists(output_path):
                        df_leido = pd.read_excel(output_path)
                        st.session_state["df_alta_cargado"] = pd.DataFrame({
                            'Medicamento': df_leido.iloc.values[:, 0],
                            'CN': df_leido.iloc.values[:, 1],
                            'Posologia': df_leido.iloc.values[:, 2],
                            'Ultima Entrega': ""
                        })
                        st.success("¡Descargado de Google Drive!")
                except Exception as e:
                    st.error(f"Error: {e}")

    with st.form("form_alta"):
        c1, c2, c3 = st.columns(3)
        with c1: nuevo_nombre = st.text_input("Nombre completo:")
        with c2: nuevo_cip = st.text_input("CIP:")
        with c3: nueva_ref = st.text_input("Referencia:", value="NUEVO")
            
        meds_editadas = st.data_editor(st.session_state["df_alta_cargado"], num_rows="dynamic", use_container_width=True)
        if st.form_submit_button("Guardar y Dar de Alta" if es_admin_rol else "Enviar Solicitud"):
            if nuevo_nombre.strip():
                df_final = meds_editadas.copy()
                df_final['Ultima Entrega'] = ""
                df_final['Pedido'] = False
                df_final['Incidencia'] = False
                if es_admin_rol:
                    etiqueta = f"{nueva_ref} — {nuevo_nombre.strip()}"
                    shared_data["lista_pacientes"][etiqueta] = {
                        "ref": nueva_ref, "nombre": nuevo_nombre.strip(), "cip": nuevo_cip, "hoja": nueva_ref, "datos": df_final
                    }
                    st.success("¡Paciente dado de alta!")
                    st.session_state["df_alta_cargado"] = pd.DataFrame(columns=['Medicamento', 'CN', 'Posologia', 'Ultima Entrega'])
                    time.sleep(1); st.rerun()
                else:
                    shared_data["solicitudes_alta"].append({"nombre": nuevo_nombre.strip(), "ref": nueva_ref, "cip": nuevo_cip, "datos": df_final})
                    st.success("Solicitud enviada al farmacéutico.")
            else:
                st.error("Introduce el nombre del paciente.")
    if st.button("⬅ Volver"): st.session_state["pagina"] = "inicio"; st.rerun()

# BAJAS DE PACIENTE Y DEVOLUCIÓN CON FICHA VISUAL DE MEDICAMENTO
elif st.session_state["pagina"] == "baja_paciente":
    if not tiene_permiso(rol_actual, "bajas"): st.error("Sin permiso."); st.stop()
        
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 800;'>👴 BAJA DE PACIENTE Y DEVOLUCIÓN</h2>", unsafe_allow_html=True)
    es_admin_rol = (rol_actual == "admin")
    
    if "devolucion_activa" not in st.session_state: st.session_state["devolucion_activa"] = False
    if "paciente_a_baja_obj" not in st.session_state: st.session_state["paciente_a_baja_obj"] = None
    if "df_devolucion" not in st.session_state:
        st.session_state["df_devolucion"] = pd.DataFrame(columns=['Marca', 'Fármaco', 'Dosificación', 'Tamaño envase', 'CN', 'Lote', 'Caducidad', 'Pastillas restantes'])

    if es_admin_rol and shared_data["solicitudes_baja"]:
        st.warning("⚠️ Solicitudes de baja pendientes:")
        for idx, sol in enumerate(shared_data["solicitudes_baja"]):
            c1, c2, c3 = st.columns([3, 1, 1])
            c1.write(f"**Nombre:** {sol['nombre']} | **Ref:** {sol['ref']}")
            if c2.button("✅ Aprobar", key=f"apr_b_{idx}"):
                st.session_state["paciente_a_baja_obj"] = {"etiqueta": f"{sol['ref']} — {sol['nombre']}", "nombre": sol['nombre'], "ref": sol['ref']}
                st.session_state["devolucion_activa"] = True
                shared_data["solicitudes_baja"].pop(idx); st.rerun()
            if c3.button("❌ Rechazar", key=f"rec_b_{idx}"):
                shared_data["solicitudes_baja"].pop(idx); st.rerun()
        st.markdown("---")

    if not st.session_state["devolucion_activa"]:
        with st.form("form_baja"):
            paciente_a_baja = st.selectbox("Seleccione el paciente:", [""] + list(shared_data["lista_pacientes"].keys()))
            hacer_devolucion = st.checkbox("🔄 ¿Registrar devolución de medicación restante?", value=True)
            if st.form_submit_button("Continuar") and paciente_a_baja:
                info_pac = shared_data["lista_pacientes"][paciente_a_baja]
                if hacer_devolucion:
                    st.session_state["paciente_a_baja_obj"] = {"etiqueta": paciente_a_baja, "nombre": info_pac["nombre"], "ref": info_pac["ref"]}
                    st.session_state["devolucion_activa"] = True
                    st.rerun()
                else:
                    if es_admin_rol:
                        del shared_data["lista_pacientes"][paciente_a_baja]
                        st.success("Paciente eliminado."); time.sleep(1); st.rerun()
    else:
        pac_obj = st.session_state["paciente_a_baja_obj"]
        st.info(f"📦 Paciente en baja: **{pac_obj['nombre']}** (Ref: {pac_obj['ref']})")
        
        st.markdown("##### 📷 Escáner de Código DataMatrix")
        cadena_dm = st.text_input("Escanee o pegue aquí el código DataMatrix del medicamento:", key="input_dm_baja")
        
        # TRADUCCIÓN AUTOMÁTICA DEL CÓDIGO
        parsed = traducir_datamatrix(cadena_dm)

        # ----------------------------------------------------
        # FICHA VISUAL DE MEDICAMENTO DECODIFICADA
        # ----------------------------------------------------
        if cadena_dm:
            st.markdown(f"""
            <div style="background: #ffffff; padding: 20px; border-radius: 12px; border: 2px solid #0ea5e9; box-shadow: 0 4px 12px rgba(14,165,233,0.1); margin-bottom: 20px;">
                <h4 style="color: #0369a1; margin-top: 0; border-bottom: 2px solid #bae6fd; padding-bottom: 8px;">💊 Ficha de Medicamento Decodificada</h4>
                <div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 12px; font-size: 14px; color: #334155;">
                    <div><b>Marca:</b> {parsed['marca']}</div>
                    <div><b>Fármaco:</b> {parsed['farmaco']}</div>
                    <div><b>Dosificación:</b> {parsed['dosificacion']}</div>
                    <div><b>Tamaño envase:</b> {parsed['tamano']}</div>
                    <div><b>Código Nacional (CN):</b> {parsed['cn']}</div>
                    <div><b>Lote:</b> {parsed['lote']}</div>
                    <div><b>Caducidad:</b> {parsed['caducidad']}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

        with st.form("form_escanear_datamatrix_baja"):
            st.markdown("Verifique o ajuste los datos extraídos e indique las pastillas restantes:")
            c1, c2, c3 = st.columns(3)
            with c1: m_val = st.text_input("Marca:", value=parsed['marca'])
            with c2: f_val = st.text_input("Fármaco:", value=parsed['farmaco'])
            with c3: d_val = st.text_input("Dosificación:", value=parsed['dosificacion'])
                
            c4, c5, c6, c7 = st.columns(4)
            with c4: t_val = st.text_input("Tamaño:", value=parsed['tamano'])
            with c5: cn_val = st.text_input("CN:", value=parsed['cn'])
            with c6: l_val = st.text_input("Lote:", value=parsed['lote'])
            with c7: cad_val = st.text_input("Caducidad:", value=parsed['caducidad'])
                
            pastillas_input = st.number_input("Pastillas restantes en el envase:", min_value=0, value=0, step=1)
            
            if st.form_submit_button("➕ Añadir a la Lista de Devolución"):
                nuevo_reg = {
                    'Marca': m_val.strip(), 'Fármaco': f_val.strip(), 'Dosificación': d_val.strip(),
                    'Tamaño envase': t_val.strip(), 'CN': cn_val.strip(), 'Lote': l_val.strip(),
                    'Caducidad': cad_val.strip(), 'Pastillas restantes': int(pastillas_input)
                }
                st.session_state["df_devolucion"] = pd.concat([st.session_state["df_devolucion"], pd.DataFrame([nuevo_reg])], ignore_index=True)
                st.success("¡Medicamento añadido a la devolución correctamente!")

        if not st.session_state["df_devolucion"].empty:
            st.markdown("##### 📋 Listado de Devolución Actual")
            st.session_state["df_devolucion"] = st.data_editor(st.session_state["df_devolucion"], use_container_width=True, hide_index=True)
            
            col_pdf, col_fin = st.columns(2)
            with col_pdf:
                pdf_bytes = generar_albaran_devolucion_pdf(pac_obj['nombre'], pac_obj['ref'], st.session_state["df_devolucion"].to_dict(orient="records"))
                st.download_button("📄 Imprimir Albarán de Devolución (PDF)", data=pdf_bytes, file_name=f"Devolucion_{pac_obj['nombre']}.pdf", mime="application/pdf", use_container_width=True)
            with col_fin:
                if st.button("💾 Finalizar y Dar de Baja Definitiva", use_container_width=True):
                    if pac_obj["etiqueta"] in shared_data["lista_pacientes"]:
                        del shared_data["lista_pacientes"][pac_obj["etiqueta"]]
                    st.success("¡Paciente dado de baja con éxito!")
                    st.session_state["devolucion_activa"] = False
                    st.session_state["df_devolucion"] = pd.DataFrame(columns=['Marca', 'Fármaco', 'Dosificación', 'Tamaño envase', 'CN', 'Lote', 'Caducidad', 'Pastillas restantes'])
                    time.sleep(1); st.rerun()
                    
        if st.button("❌ Cancelar Devolución"):
            st.session_state["devolucion_activa"] = False
            st.rerun()

    if not st.session_state["devolucion_activa"] and st.button("⬅ Volver"):
        st.session_state["pagina"] = "inicio"; st.rerun()

# GESTIÓN DE USUARIOS
elif st.session_state["pagina"] == "gestion_usuarios":
    if not (tiene_permiso(rol_actual, "usuarios") or tiene_permiso(rol_actual, "roles")): st.error("Sin permiso."); st.stop()
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 800;'>🔐 GESTIÓN DE USUARIOS Y ROLES</h2>", unsafe_allow_html=True)
    if st.button("⬅ Volver"): st.session_state["pagina"] = "inicio"; st.rerun()

# LISTA PACIENTES Y DETALLE
elif st.session_state["pagina"] == "lista_pacientes":
    if not tiene_permiso(rol_actual, "pacientes"): st.error("Sin permiso."); st.stop()
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 800;'>PACIENTES</h2>", unsafe_allow_html=True)
    for pk in list(lista_pacientes.keys()):
        if st.button(pk, key=f"p_{pk}", use_container_width=True):
            st.session_state["paciente_seleccionado_key"] = pk
            st.session_state["pagina"] = "detalle_paciente"; st.rerun()
    if st.button("⬅ Volver"): st.session_state["pagina"] = "inicio"; st.rerun()

elif st.session_state["pagina"] == "detalle_paciente":
    if not tiene_permiso(rol_actual, "pacientes"): st.error("Sin permiso."); st.stop()
    pk = st.session_state["paciente_seleccionado_key"]
    info = lista_pacientes.get(pk)
    if info:
        st.markdown(f"<h3 style='text-align: center;'>Paciente: {info['nombre']}</h3>", unsafe_allow_html=True)
        info["datos"] = st.data_editor(info["datos"], use_container_width=True, hide_index=True)
    if st.button("⬅ Volver"): st.session_state["pagina"] = "lista_pacientes"; st.rerun()

elif st.session_state["pagina"] == "seleccion_productos_enfermera":
    if not tiene_permiso(rol_actual, "propuesta"): st.error("Sin permiso."); st.stop()
    st.markdown("<h2 style='text-align: center;'>📦 PROPUESTA DE PEDIDO</h2>", unsafe_allow_html=True)
    if not shared_data["solicitud_pedido"]: st.info("No hay propuestas.")
    else:
        df_sol = pd.DataFrame(shared_data["solicitud_pedido"])
        shared_data["solicitud_pedido"] = st.data_editor(df_sol, use_container_width=True, hide_index=True).to_dict(orient="records")
        if st.button("🚀 Solicitar Pedido Definitivo"):
            sel = [i for i in shared_data["solicitud_pedido"] if i.get("seleccion_enfermera")]
            for it in sel: shared_data["pedidos_definitivos"].append(it)
            shared_data["solicitud_pedido"] = [i for i in shared_data["solicitud_pedido"] if not i.get("seleccion_enfermera")]
            st.success("Enviado al farmacéutico."); time.sleep(1); st.rerun()
    if st.button("⬅ Volver"): st.session_state["pagina"] = "inicio"; st.rerun()

elif st.session_state["pagina"] == "solicitud_pedido_admin":
    if not tiene_permiso(rol_actual, "propuesta"): st.error("Sin permiso."); st.stop()
    st.markdown("<h2 style='text-align: center;'>📦 PROPUESTA (ENVIADO A ENFERMERÍA)</h2>", unsafe_allow_html=True)
    if shared_data["solicitud_pedido"]: st.dataframe(pd.DataFrame(shared_data["solicitud_pedido"]), use_container_width=True, hide_index=True)
    else: st.info("Vacío.")
    if st.button("⬅ Volver"): st.session_state["pagina"] = "inicio"; st.rerun()

# PEDIDOS DEFINITIVOS CON TRADUCCIÓN DATAMATRIX
elif st.session_state["pagina"] == "pedidos_definitivos_admin":
    if not tiene_permiso(rol_actual, "pedidos_def"): st.error("Sin permiso."); st.stop()
        
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 800;'>🛒 PEDIDOS DEFINITIVOS Y ESCÁNER</h2>", unsafe_allow_html=True)
    if not shared_data["pedidos_definitivos"]:
        st.info("No hay pedidos definitivos pendientes.")
    else:
        st.markdown("##### ⚡ Escáner Rápido DataMatrix")
        cadena_dm_pedido = st.text_input("Escanee el DataMatrix del medicamento:", key="input_dm_pedido")
        
        parsed_ped = traducir_datamatrix(cadena_dm_pedido)
        if cadena_dm_pedido:
            for item in shared_data["pedidos_definitivos"]:
                if not item.get("datamatrix"):
                    item["datamatrix"] = cadena_dm_pedido
                    item["lote"] = parsed_ped['lote']
                    item["caducidad"] = parsed_ped['caducidad']
                    break

        df_defs = pd.DataFrame(shared_data["pedidos_definitivos"])
        edited_defs = st.data_editor(
            df_defs, use_container_width=True, hide_index=True,
            column_config={
                "paciente": st.column_config.TextColumn("Paciente", disabled=True),
                "ref": st.column_config.TextColumn("Ref.", disabled=True),
                "medicamento": st.column_config.TextColumn("Medicamento", disabled=True),
                "cn": st.column_config.TextColumn("C.N.", disabled=True),
                "posologia": st.column_config.TextColumn("Posología", disabled=True),
                "datamatrix": st.column_config.TextColumn("DataMatrix"),
                "lote": st.column_config.TextColumn("Lote"),
                "caducidad": st.column_config.TextColumn("Caducidad")
            }
        )
        shared_data["pedidos_definitivos"] = edited_defs.to_dict(orient="records")
        
        st.markdown("---")
        col_pdf, col_act = st.columns(2)
        with col_pdf:
            pdf_bytes = generar_albaran_pdf(shared_data["pedidos_definitivos"])
            st.download_button("📄 Imprimir Albarán de Entrega (PDF)", data=pdf_bytes, file_name="Albaran_Entrega.pdf", mime="application/pdf", use_container_width=True)
        with col_act:
            if st.button("📌 Actualizar Última Entrega y Limpiar", use_container_width=True):
                fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M")
                for item_ped in shared_data["pedidos_definitivos"]:
                    nom_p = item_ped.get("paciente")
                    cn_p = str(item_ped.get("cn"))
                    for pk, pval in shared_data["lista_pacientes"].items():
                        if pval["nombre"] == nom_p:
                            dfm = pval["datos"]
                            mask = dfm['CN'].astype(str).str.strip() == cn_p.strip()
                            dfm.loc[mask, 'Ultima Entrega'] = fecha_actual
                shared_data["pedidos_definitivos"] = []
                st.success("¡Fechas actualizadas en las fichas de los pacientes!")
                time.sleep(1); st.rerun()

    if st.button("⬅ Volver al Menú"): st.session_state["pagina"] = "inicio"; st.rerun()

elif st.session_state["pagina"] == "incidencias":
    if not tiene_permiso(rol_actual, "incidencias"): st.error("Sin permiso."); st.stop()
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 800;'>⚠️ PANEL DE INCIDENCIAS</h2>", unsafe_allow_html=True)
    if shared_data["incidencias_activas"]:
        df_inc = pd.DataFrame(shared_data["incidencias_activas"])
        shared_data["incidencias_activas"] = st.data_editor(df_inc, use_container_width=True, hide_index=True).to_dict(orient="records")
    else: st.info("No hay incidencias.")
    if st.button("⬅ Volver"): st.session_state["pagina"] = "inicio"; st.rerun()
