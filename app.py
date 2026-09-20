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

# Estilos CSS definitivos para alineación perfecta y optimización del espacio superior
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
    [data-testid="collapsedControl"] {
        display: none;
    }
    
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

    /* FORZAR ALINEACIÓN ABSOLUTA EN LAS 6 COLUMNAS SUPERIORES */
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

    div.stButton {
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
        box-shadow: 0 4px 12px rgba(14, 165, 233, 0.15) !important;
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

def parsear_datamatrix(dm):
    """Analiza de forma segura un código DataMatrix GS1 para extraer CN, Lote y Caducidad."""
    resultado = {
        'marca': 'Marca Genérica',
        'farmaco': 'Medicamento',
        'dosificacion': '100mg',
        'tamano': '28 comp',
        'cn': '',
        'lote': '',
        'caducidad': ''
    }
    if not dm:
        return resultado
    
    limpio = str(dm).replace('\x1D', '').replace('(', '').replace(')', '').strip()
    
    try:
        if '01' in limpio:
            idx = limpio.find('01')
            if len(limpio) >= idx + 16:
                gtin = limpio[idx+2 : idx+16]
                if len(gtin) == 14:
                    resultado['cn'] = gtin[7:13]
                    resultado['farmaco'] = f"Fármaco CN {resultado['cn']}"
        
        if '17' in limpio:
            idx = limpio.find('17')
            if len(limpio) >= idx + 8:
                cad = limpio[idx+2 : idx+8]
                if len(cad) == 6:
                    yy, mm, dd = cad[0:2], cad[2:4], cad[4:6]
                    if dd == '00': dd = '01'
                    resultado['caducidad'] = f"{dd}/{mm}/20{yy}"
        
        if '10' in limpio:
            idx = limpio.find('10')
            lote_val = limpio[idx+2:]
            for ai in ['21', '17', '30', '11']:
                if ai in lote_val:
                    lote_val = lote_val.split(ai)[0]
            resultado['lote'] = lote_val[:20].strip()
            
    except Exception:
        pass
        
    return resultado

def generar_albaran_pdf(lista_pedidos):
    pdf = FPDF(orientation='L') 
    pdf.add_page()
    
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, "FARMACIA VILLEGAS C.B.", ln=True, align='C')
    pdf.set_font("Arial", '', 11)
    pdf.cell(0, 6, "C/ INDEPENDENCIA, 5", ln=True, align='C')
    pdf.cell(0, 6, "13700 TOMELLOSO, CIUDAD REAL", ln=True, align='C')
    
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
        ref = str(row.get('ref', ''))[:8]
        pac = str(row.get('paciente', ''))[:35]
        med = str(row.get('medicamento', ''))[:45]
        cn = str(row.get('cn', ''))[:10]
        pos = str(row.get('posologia', ''))[:12]
        dm = str(row.get('datamatrix', ''))[:35]
        lote = str(row.get('lote', ''))[:15]
        cad = str(row.get('caducidad', ''))[:12]
        
        pdf.cell(col_widths[0], 8, ref, border=1)
        pdf.cell(col_widths[1], 8, pac, border=1)
        pdf.cell(col_widths[2], 8, med, border=1)
        pdf.cell(col_widths[3], 8, cn, border=1, align='C')
        pdf.cell(col_widths[4], 8, pos, border=1, align='C')
        pdf.cell(col_widths[5], 8, dm, border=1)
        pdf.cell(col_widths[6], 8, lote, border=1, align='C')
        pdf.cell(col_widths[7], 8, cad, border=1, align='C')
        pdf.ln()
        
    pdf.ln(20)
    pdf.set_font("Arial", 'I', 10)
    pdf.cell(0, 10, "Firma del receptor: ___________________________", ln=True, align='R')
    return pdf.output(dest='S').encode('latin1')

def generar_albaran_devolucion_pdf(nombre_paciente, ref_paciente, lista_devolucion):
    pdf = FPDF(orientation='L') 
    pdf.add_page()
    
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, "FARMACIA VILLEGAS C.B. - ALBARÁN DE DEVOLUCIÓN", ln=True, align='C')
    pdf.set_font("Arial", '', 11)
    pdf.cell(0, 6, "C/ INDEPENDENCIA, 5 — 13700 TOMELLOSO, CIUDAD REAL", ln=True, align='C')
    
    pdf.ln(8)
    pdf.set_font("Arial", 'B', 11)
    fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M")
    pdf.cell(0, 8, f"Paciente: {nombre_paciente} (Ref: {ref_paciente})", ln=True, align='L')
    pdf.cell(0, 8, f"Fecha de Devolución: {fecha_actual}", ln=True, align='L')
    pdf.ln(4)
    
    pdf.set_font("Arial", 'B', 9)
    col_widths = [35, 55, 25, 25, 20, 25, 20, 25] 
    headers = ["Marca", "Fármaco", "Dosificación", "Tamaño", "CN", "Lote", "Caducidad", "Restantes"]
    
    for i in range(len(headers)):
        pdf.cell(col_widths[i], 8, headers[i], border=1, align='C')
    pdf.ln()
    
    pdf.set_font("Arial", '', 8)
    for row in lista_devolucion:
        marca = str(row.get('Marca', ''))[:20]
        farmaco = str(row.get('Fármaco', ''))[:30]
        dosif = str(row.get('Dosificación', ''))[:15]
        tam = str(row.get('Tamaño envase', ''))[:15]
        cn = str(row.get('CN', ''))[:10]
        lote = str(row.get('Lote', ''))[:12]
        cad = str(row.get('Caducidad', ''))[:10]
        restantes = str(row.get('Pastillas restantes', '0'))[:8]
        
        pdf.cell(col_widths[0], 8, marca, border=1)
        pdf.cell(col_widths[1], 8, farmaco, border=1)
        pdf.cell(col_widths[2], 8, dosif, border=1, align='C')
        pdf.cell(col_widths[3], 8, tam, border=1, align='C')
        pdf.cell(col_widths[4], 8, cn, border=1, align='C')
        pdf.cell(col_widths[5], 8, lote, border=1, align='C')
        pdf.cell(col_widths[6], 8, cad, border=1, align='C')
        pdf.cell(col_widths[7], 8, restantes, border=1, align='C')
        pdf.ln()
        
    pdf.ln(15)
    pdf.set_font("Arial", 'I', 10)
    pdf.cell(0, 10, "Firma del farmacéutico: ___________________________       Firma del receptor: ___________________________", ln=True, align='C')
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

# ----------------------------------------------------
# MEMORIA GLOBAL COMPARTIDA
# ----------------------------------------------------
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

# ----------------------------------------------------
# PÁGINA 0: LOGIN
# ----------------------------------------------------
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

# ----------------------------------------------------
# CABECERA SUPERIOR (6 BOTONES)
# ----------------------------------------------------
st.markdown('<div class="dashboard-header">', unsafe_allow_html=True)

st.markdown("""
    <div class="logo-container">
        <span style="font-size: 24px;">💊</span>
        <span class="logo-title">SPD FARMACIA VILLEGAS</span>
    </div>
""", unsafe_allow_html=True)

rol_actual = st.session_state["rol_usuario"]
rol_txt = rol_actual.upper()
st.markdown(f"""
    <div class="status-bar">
        <span>Sistema activo <span style="color: #22c55e; font-size: 16px;">●</span></span>
        <span>Usuario: <b>{st.session_state['usuario_autenticado']}</b> ({rol_txt})</span>
    </div>
""", unsafe_allow_html=True)

col_alta, col_baja, col_ped, col_inc, col_user, col_logout = st.columns(6, gap="small")

with col_alta:
    if tiene_permiso(rol_actual, "altas"):
        num_altas = len(shared_data["solicitudes_alta"]) if rol_actual == "admin" else 0
        lbl_alta = f"ALTA ({num_altas})" if num_altas > 0 else "ALTA"
        if num_altas > 0:
            st.markdown('<div class="alerta-wrapper alerta-activa">', unsafe_allow_html=True)
        if st.button(lbl_alta, key="btn_alta", use_container_width=True): 
            st.session_state["pagina"] = "alta_paciente"
            st.rerun()
        if num_altas > 0:
            st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div style="height: 48px;"></div>', unsafe_allow_html=True)
        
with col_baja:
    if tiene_permiso(rol_actual, "bajas"):
        num_bajas = len(shared_data["solicitudes_baja"]) if rol_actual == "admin" else 0
        lbl_baja = f"BAJA ({num_bajas})" if num_bajas > 0 else "BAJA"
        if num_bajas > 0:
            st.markdown('<div class="alerta-wrapper alerta-activa">', unsafe_allow_html=True)
        if st.button(lbl_baja, key="btn_baja", use_container_width=True): 
            st.session_state["pagina"] = "baja_paciente"
            st.rerun()
        if num_bajas > 0:
            st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div style="height: 48px;"></div>', unsafe_allow_html=True)

with col_ped:
    if rol_actual == "admin" and tiene_permiso(rol_actual, "pedidos_def"):
        num_defs = len(shared_data["pedidos_definitivos"])
        lbl_ped = "PEDIDOS"
        if num_defs > 0:
            st.markdown('<div class="alerta-wrapper alerta-activa">', unsafe_allow_html=True)
        if st.button(lbl_ped, key="btn_pedidos_admin", use_container_width=True):
            st.session_state["pagina"] = "pedidos_definitivos_admin"
            st.rerun()
        if num_defs > 0:
            st.markdown('</div>', unsafe_allow_html=True)
    elif tiene_permiso(rol_actual, "propuesta"):
        num_sol = len(shared_data["solicitud_pedido"])
        lbl_sol = "PROPUESTA"
        if num_sol > 0:
            st.markdown('<div class="alerta-wrapper alerta-activa">', unsafe_allow_html=True)
        if st.button(lbl_sol, key="btn_sol_enf_top", use_container_width=True):
            st.session_state["pagina"] = "seleccion_productos_enfermera"
            st.rerun()
        if num_sol > 0:
            st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div style="height: 48px;"></div>', unsafe_allow_html=True)
            
with col_inc:
    if tiene_permiso(rol_actual, "incidencias"):
        num_inc = len(shared_data["incidencias_activas"])
        lbl_inc = f"INCIDENCIAS ({num_inc})" if num_inc > 0 else "INCIDENCIAS"
        if num_inc > 0:
            st.markdown('<div class="alerta-wrapper alerta-activa">', unsafe_allow_html=True)
        if st.button(lbl_inc, key="btn_incidencias", use_container_width=True):
            st.session_state["pagina"] = "incidencias"
            st.rerun()
        if num_inc > 0:
            st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div style="height: 48px;"></div>', unsafe_allow_html=True)

with col_user:
    if tiene_permiso(rol_actual, "usuarios") or tiene_permiso(rol_actual, "roles"):
        if st.button("USUARIOS", key="btn_usuarios_top", use_container_width=True):
            st.session_state["pagina"] = "gestion_usuarios"
            st.rerun()
    else:
        st.markdown('<div style="height: 48px;"></div>', unsafe_allow_html=True)

with col_logout:
    if st.button("SALIR", key="btn_logout_top", use_container_width=True):
        token_url = obtener_parametro_url("session_token")
        if token_url and token_url in shared_data["sesiones_activas"]:
            del shared_data["sesiones_activas"][token_url]
        limpiar_parametros_url()
        st.session_state["usuario_autenticado"] = None
        st.session_state["rol_usuario"] = None
        st.session_state["pagina"] = "inicio"
        st.rerun()

st.markdown('</div>', unsafe_allow_html=True)

# ----------------------------------------------------
# VISTAS PRINCIPALES (DASHBOARD CON PANELES PASTEL)
# ----------------------------------------------------
if st.session_state["pagina"] == "inicio":
    total_pacientes = len(shared_data["lista_pacientes"])
    total_incidencias = len(shared_data["incidencias_activas"])
    total_pedidos_def = len(shared_data["pedidos_definitivos"])
    total_solicitudes = len(shared_data["solicitud_pedido"])

    c_card1, c_card2 = st.columns(2, gap="large")
    
    with c_card1:
        if tiene_permiso(rol_actual, "pacientes"):
            st.markdown(f"""
                <div style="background: #eff6ff; padding: 22px; border-radius: 16px; border: 1px solid #bfdbfe; box-shadow: 0 4px 12px rgba(0,0,0,0.03); margin-bottom: 10px;">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <h4 style="color: #1e3a8a; margin-top: 0;">👤 PACIENTE</h4>
                        <span style="background: #dbeafe; color: #1e40af; padding: 4px 12px; border-radius: 20px; font-weight: 700; font-size: 13px;">{total_pacientes} activos</span>
                    </div>
                    <p style="color: #334155; font-size: 14px; font-weight: 500; margin-bottom: 0;">Acceso al listado completo de pacientes, revisión de tratamientos y gestión de estados clínicos.</p>
                </div>
            """, unsafe_allow_html=True)
            
            if st.button("🧓 **ENTRAR A PACIENTES**", key="button_pacientes", use_container_width=True):
                st.session_state["pagina"] = "lista_pacientes"
                st.rerun()
            
    with c_card2:
        if tiene_permiso(rol_actual, "incidencias"):
            st.markdown(f"""
                <div style="background: #fef2f2; padding: 22px; border-radius: 16px; border: 1px solid #fecaca; box-shadow: 0 4px 12px rgba(0,0,0,0.03); margin-bottom: 10px;">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <h4 style="color: #7f1d1d; margin-top: 0;">⚠️ PANEL INCIDENCIAS</h4>
                        <span style="background: #fee2e2; color: #991b1b; padding: 4px 12px; border-radius: 20px; font-weight: 700; font-size: 13px;">{total_incidencias} abiertas</span>
                    </div>
                    <p style="color: #334155; font-size: 14px; font-weight: 500; margin-bottom: 0;">Gestión de incidencias reportadas (falta de receta, modificación de posología, abastecimiento).</p>
                </div>
            """, unsafe_allow_html=True)
            
            if st.button("📋 **VER PANEL DE INCIDENCIAS**", key="button_incidencias", use_container_width=True):
                st.session_state["pagina"] = "incidencias"
                st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)
    c_card3, c_card4 = st.columns(2, gap="large")
    
    with c_card3:
        if tiene_permiso(rol_actual, "propuesta") or tiene_permiso(rol_actual, "pedidos_def"):
            st.markdown(f"""
                <div style="background: #fefce8; padding: 22px; border-radius: 16px; border: 1px solid #fef08a; box-shadow: 0 4px 12px rgba(0,0,0,0.03); margin-bottom: 10px;">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <h4 style="color: #713f12; margin-top: 0;">🚚 PROPUESTA DE PEDIDO</h4>
                        <span style="background: #fef08a; color: #854d0e; padding: 4px 12px; border-radius: 20px; font-weight: 700; font-size: 13px;">{total_solicitudes} propuestas</span>
                    </div>
                    <p style="color: #334155; font-size: 14px; font-weight: 500; margin-bottom: 0;">Seguimiento en tiempo real de las propuestas de pedido enviadas entre farmacia y enfermería.</p>
                </div>
            """, unsafe_allow_html=True)
            
            if rol_actual == "admin":
                if st.button("📦 **VER PROPUESTA DE PEDIDO**", key="button_solicitudes", use_container_width=True):
                    st.session_state["pagina"] = "solicitud_pedido_admin"
                    st.rerun()
            else:
                if st.button("📦 **BANDEJA DE PROPUESTAS DE PEDIDO**", key="button_bandeja", use_container_width=True):
                    st.session_state["pagina"] = "seleccion_productos_enfermera"
                    st.rerun()
                
    with c_card4:
        if tiene_permiso(rol_actual, "pedidos_def") or tiene_permiso(rol_actual, "propuesta"):
            st.markdown(f"""
                <div style="background: #f0fdf4; padding: 22px; border-radius: 16px; border: 1px solid #bbf7d0; box-shadow: 0 4px 12px rgba(0,0,0,0.03); margin-bottom: 10px;">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <h4 style="color: #14532d; margin-top: 0;">📄 PEDIDO DEFINITIVO</h4>
                        <span style="background: #dcfce7; color: #166534; padding: 4px 12px; border-radius: 20px; font-weight: 700; font-size: 13px;">{total_pedidos_def} listos</span>
                    </div>
                    <p style="color: #334155; font-size: 14px; font-weight: 500; margin-bottom: 0;">Validación final con códigos DataMatrix, control de lotes y generación de albaranes PDF.</p>
                </div>
            """, unsafe_allow_html=True)
            
            if rol_actual == "admin":
                if st.button("🛒 **VER PEDIDOS DEFINITIVOS**", key="button_definitivos", use_container_width=True):
                    st.session_state["pagina"] = "pedidos_definitivos_admin"
                    st.rerun()
            else:
                if st.button("📦 **SELECCIÓN DE ENFERMERÍA**", key="button_sel_enf", use_container_width=True):
                    st.session_state["pagina"] = "seleccion_productos_enfermera"
                    st.rerun()

# ----------------------------------------------------
# ALTAS DE PACIENTE
# ----------------------------------------------------
elif st.session_state["pagina"] == "alta_paciente":
    if not tiene_permiso(rol_actual, "altas"):
        st.error("No tienes permiso para acceder a este módulo.")
        st.stop()
        
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 800;'>👴 ALTA DE PACIENTE</h2>", unsafe_allow_html=True)
    es_admin_rol = (rol_actual == "admin")
    
    if es_admin_rol and shared_data["solicitudes_alta"]:
        st.warning("⚠️ Tienes solicitudes de alta pendientes de Enfermería:")
        for idx, sol in enumerate(shared_data["solicitudes_alta"]):
            c1, c2, c3 = st.columns([3, 1, 1])
            c1.write(f"**Nombre:** {sol['nombre']} | **CIP:** {sol.get('cip', 'N/A')} | **Ref:** {sol.get('ref', 'N/A')}")
            if c2.button("✅ Aprobar Alta", key=f"apr_alta_{idx}"):
                etiqueta = f"{sol['ref']} — {sol['nombre']}"
                shared_data["lista_pacientes"][etiqueta] = {
                    "ref": sol['ref'], "nombre": sol['nombre'], "cip": sol['cip'], "hoja": sol['ref'], "datos": sol['datos']
                }
                shared_data["solicitudes_alta"].pop(idx)
                st.success(f"¡Paciente {sol['nombre']} dado de alta con éxito!")
                time.sleep(1)
                st.rerun()
            if c3.button("❌ Rechazar", key=f"rec_alta_{idx}"):
                shared_data["solicitudes_alta"].pop(idx)
                st.rerun()
        st.markdown("---")

    st.subheader("Formulario de Alta y Carga de Medicación")
    
    if "df_alta_cargado" not in st.session_state:
        st.session_state["df_alta_cargado"] = pd.DataFrame(columns=['Medicamento', 'CN', 'Posologia', 'Ultima Entrega'])

    tab_local, tab_drive = st.tabs(["📁 Subir archivo local", "☁️ Importar desde Google Drive"])

    with tab_local:
        archivo_subido = st.file_uploader("Sube un fichero Excel (.xlsx)", type=["xlsx", "xls"], key="uploader_excel_alta")
        if archivo_subido is not None:
            try:
                df_excel_leido = pd.read_excel(archivo_subido)
                df_excel_leido.columns = [str(c).strip().capitalize() for c in df_excel_leido.columns]
                cols_disp = df_excel_leido.columns
                col_m = next((c for c in cols_disp if 'medicamento' in c.lower() or 'fármaco' in c.lower()), None)
                col_c = next((c for c in cols_disp if 'cn' in c.lower() or 'codigo' in c.lower()), None)
                col_p = next((c for c in cols_disp if 'posología' in c.lower() or 'posologia' in c.lower()), None)
                
                if col_m and col_c and col_p:
                    st.session_state["df_alta_cargado"] = pd.DataFrame({
                        'Medicamento': df_excel_leido[col_m],
                        'CN': df_excel_leido[col_c],
                        'Posologia': df_excel_leido[col_p],
                        'Ultima Entrega': ""
                    })
                    st.success("¡Excel local leído con éxito!")
                elif len(df_excel_leido.columns) >= 3:
                    st.session_state["df_alta_cargado"] = pd.DataFrame({
                        'Medicamento': df_excel_leido.iloc[:, 0],
                        'CN': df_excel_leido.iloc[:, 1],
                        'Posologia': df_excel_leido.iloc[:, 2],
                        'Ultima Entrega': ""
                    })
                    st.success("¡Datos cargados correctamente!")
            except Exception as e:
                st.error(f"Error al leer el archivo: {e}")

    with tab_drive:
        st.markdown("Pega el enlace de compartir de **Google Drive** del archivo Excel (configurado como *'Cualquier persona con el enlace puede ver'*):")
        gdrive_url = st.text_input("Enlace de Google Drive:")
        if st.button("📥 Descargar y Leer desde Google Drive"):
            if gdrive_url:
                try:
                    with st.spinner("Conectando con Google Drive y descargando archivo..."):
                        output_path = "temp_gdrive.xlsx"
                        import gdown
                        gdown.download(gdrive_url, output_path, quiet=False)
                        if os.path.exists(output_path):
                            df_excel_leido = pd.read_excel(output_path)
                            df_excel_leido.columns = [str(c).strip().capitalize() for c in df_excel_leido.columns]
                            cols_disp = df_excel_leido.columns
                            col_m = next((c for c in cols_disp if 'medicamento' in c.lower() or 'fármaco' in c.lower()), None)
                            col_c = next((c for c in cols_disp if 'cn' in c.lower() or 'codigo' in c.lower()), None)
                            col_p = next((c for c in cols_disp if 'posología' in c.lower() or 'posologia' in c.lower()), None)
                            
                            if col_m and col_c and col_p:
                                st.session_state["df_alta_cargado"] = pd.DataFrame({
                                    'Medicamento': df_excel_leido[col_m],
                                    'CN': df_excel_leido[col_c],
                                    'Posologia': df_excel_leido[col_p],
                                    'Ultima Entrega': ""
                                })
                            elif len(df_excel_leido.columns) >= 3:
                                st.session_state["df_alta_cargado"] = pd.DataFrame({
                                    'Medicamento': df_excel_leido.iloc[:, 0],
                                    'CN': df_excel_leido.iloc[:, 1],
                                    'Posologia': df_excel_leido.iloc[:, 2],
                                    'Ultima Entrega': ""
                                })
                            st.success("¡Archivo descargado y cargado correctamente desde Google Drive!")
                        else:
                            st.error("No se pudo descargar el archivo. Comprueba que el enlace sea público.")
                except Exception as e:
                    st.error(f"Error al procesar el enlace de Google Drive: {e}")
            else:
                st.warning("Por favor, introduce un enlace válido.")

    with st.form("form_alta"):
        col_datos1, col_datos2, col_datos3 = st.columns(3)
        with col_datos1:
            nuevo_nombre = st.text_input("Nombre completo del paciente:")
        with col_datos2:
            nuevo_cip = st.text_input("Código CIP del paciente:")
        with col_datos3:
            nueva_ref = st.text_input("Código o Referencia:", value="NUEVO")
            
        st.write("💊 **Listado de medicación (editable):**")
        meds_editadas = st.data_editor(
            st.session_state["df_alta_cargado"], 
            num_rows="dynamic", 
            use_container_width=True,
            key="meds_alta_editor"
        )
        
        lbl_submit = "Guardar y Dar de Alta" if es_admin_rol else "Enviar Solicitud de Alta"
        submit = st.form_submit_button(lbl_submit)
        
        if submit:
            if nuevo_nombre and nuevo_nombre.strip():
                df_final = meds_editadas.copy()
                if 'Ultima Entrega' not in df_final.columns:
                    df_final['Ultima Entrega'] = ""
                df_final['Pedido'] = False
                df_final['Incidencia'] = False
                df_final['Fecha inicio'] = datetime.now().strftime("%Y-%m-%d")
                
                if es_admin_rol:
                    etiqueta = f"{nueva_ref} — {nuevo_nombre.strip()}"
                    shared_data["lista_pacientes"][etiqueta] = {
                        "ref": nueva_ref, "nombre": nuevo_nombre.strip(), "cip": nuevo_cip, "hoja": nueva_ref, "datos": df_final
                    }
                    st.success(f"¡Paciente '{nuevo_nombre.strip()}' dado de alta con éxito!")
                    st.session_state["df_alta_cargado"] = pd.DataFrame(columns=['Medicamento', 'CN', 'Posologia', 'Ultima Entrega'])
                    st.rerun()
                else:
                    shared_data["solicitudes_alta"].append({
                        "nombre": nuevo_nombre.strip(), "ref": nueva_ref, "cip": nuevo_cip, "datos": df_final
                    })
                    st.success("Su solicitud de alta y medicación ha sido enviada al farmacéutico.")
            else:
                st.error("⚠️ Por favor, introduce el nombre completo del paciente antes de guardar.")
                
    st.write("")
    if st.button("⬅ Volver al Menú Principal"):
        st.session_state["pagina"] = "inicio"
        st.rerun()

# ----------------------------------------------------
# BAJAS DE PACIENTE Y DEVOLUCIÓN AUTOMATIZADA CON DATAMATRIX
# ----------------------------------------------------
elif st.session_state["pagina"] == "baja_paciente":
    if not tiene_permiso(rol_actual, "bajas"):
        st.error("No tienes permiso para acceder a este módulo.")
        st.stop()
        
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 800;'>👴 BAJA DE PACIENTE Y DEVOLUCIÓN</h2>", unsafe_allow_html=True)
    es_admin_rol = (rol_actual == "admin")
    
    if "devolucion_activa" not in st.session_state:
        st.session_state["devolucion_activa"] = False
    if "paciente_a_baja_obj" not in st.session_state:
        st.session_state["paciente_a_baja_obj"] = None
    if "df_devolucion" not in st.session_state:
        st.session_state["df_devolucion"] = pd.DataFrame(columns=[
            'Marca', 'Fármaco', 'Dosificación', 'Tamaño envase', 'CN', 'Lote', 'Caducidad', 'Pastillas restantes'
        ])
    if "dm_input_val" not in st.session_state:
        st.session_state["dm_input_val"] = ""

    if es_admin_rol and shared_data["solicitudes_baja"]:
        st.warning("⚠️ Tienes solicitudes de baja pendientes de Enfermería:")
        for idx, sol in enumerate(shared_data["solicitudes_baja"]):
            c1, c2, c3 = st.columns([3, 1, 1])
            c1.write(f"**Nombre:** {sol['nombre']} | **Ref:** {sol.get('ref', 'N/A')}")
            if c2.button("✅ Aprobar Baja y Devolución", key=f"apr_baja_{idx}"):
                etiqueta = f"{sol['ref']} — {sol['nombre']}"
                st.session_state["paciente_a_baja_obj"] = {
                    "etiqueta": etiqueta,
                    "nombre": sol['nombre'],
                    "ref": sol['ref']
                }
                st.session_state["devolucion_activa"] = True
                shared_data["solicitudes_baja"].pop(idx)
                st.rerun()
            if c3.button("❌ Rechazar", key=f"rec_baja_{idx}"):
                shared_data["solicitudes_baja"].pop(idx)
                st.rerun()
        st.markdown("---")

    if not st.session_state["devolucion_activa"]:
        st.subheader("Seleccionar Paciente para Baja")
        with st.form("form_baja"):
            opciones_pacientes = [""] + list(shared_data["lista_pacientes"].keys())
            paciente_a_baja = st.selectbox("Seleccione el paciente:", opciones_pacientes)
            hacer_devolucion = st.checkbox("🔄 ¿Registrar devolución de medicación restante?", value=True)
            
            lbl_submit = "Continuar con Baja y Devolución" if hacer_devolucion else ("Confirmar Baja Definitiva" if es_admin_rol else "Enviar Solicitud de Baja")
            submit = st.form_submit_button(lbl_submit)
            
            if submit and paciente_a_baja:
                info_pac = shared_data["lista_pacientes"][paciente_a_baja]
                if hacer_devolucion:
                    st.session_state["paciente_a_baja_obj"] = {
                        "etiqueta": paciente_a_baja,
                        "nombre": info_pac["nombre"],
                        "ref": info_pac["ref"]
                    }
                    st.session_state["devolucion_activa"] = True
                    st.rerun()
                else:
                    if es_admin_rol:
                        del shared_data["lista_pacientes"][paciente_a_baja]
                        st.success(f"Paciente '{info_pac['nombre']}' eliminado del sistema.")
                        time.sleep(1)
                        st.rerun()
                    else:
                        shared_data["solicitudes_baja"].append({"nombre": info_pac["nombre"], "ref": info_pac["ref"]})
                        st.success("Su solicitud de baja ha sido enviada al farmacéutico para su validación.")
    else:
        pac_obj = st.session_state["paciente_a_baja_obj"]
        st.info(f"📦 Registrando devolución de medicación para: **{pac_obj['nombre']}** (Ref: {pac_obj['ref']})")
        
        # Callback para autocompletar instantáneamente al teclear/escanear DataMatrix
        def actualizar_campos_dm():
            val = st.session_state.get("widget_dm_baja", "")
            parsed = parsear_datamatrix(val)
            st.session_state["baja_marca"] = parsed['marca']
            st.session_state["baja_farmaco"] = parsed['farmaco']
            st.session_state["baja_dosificacion"] = parsed['dosificacion']
            st.session_state["baja_tamano"] = parsed['tamano']
            st.session_state["baja_cn"] = parsed['cn']
            st.session_state["baja_lote"] = parsed['lote']
            st.session_state["baja_caducidad"] = parsed['caducidad']

        st.markdown("##### 📷 Lector de Código DataMatrix (Autocompletado Automático)")
        st.text_input(
            "Escanee el código DataMatrix completo o introduzca la cadena GS1:",
            key="widget_dm_baja",
            on_change=actualizar_campos_dm
        )
        
        # Inicializar estados si no existen
        for k, v in [("baja_marca",""), ("baja_farmaco",""), ("baja_dosificacion",""), ("baja_tamano",""), ("baja_cn",""), ("baja_lote",""), ("baja_caducidad","")]:
            if k not in st.session_state: st.session_state[k] = v

        with st.form("form_escanear_datamatrix_baja"):
            c_m1, c_m2, c_m3 = st.columns(3)
            with c_m1:
                m_val = st.text_input("Marca:", value=st.session_state["baja_marca"])
            with c_m2:
                f_val = st.text_input("Nombre del fármaco:", value=st.session_state["baja_farmaco"])
            with c_m3:
                d_val = st.text_input("Dosificación:", value=st.session_state["baja_dosificacion"])
                
            c_m4, c_m5, c_m6, c_m7 = st.columns(4)
            with c_m4:
                t_val = st.text_input("Tamaño envase:", value=st.session_state["baja_tamano"])
            with c_m5:
                cn_val = st.text_input("CN (Código Nacional):", value=st.session_state["baja_cn"])
            with c_m6:
                l_val = st.text_input("Número de Lote:", value=st.session_state["baja_lote"])
            with c_m7:
                cad_val = st.text_input("Fecha Caducidad:", value=st.session_state["baja_caducidad"])
                
            pastillas_input = st.number_input("Número de pastillas restantes:", min_value=0, value=0, step=1)
            
            btn_anadir_dev = st.form_submit_button("➕ Añadir a la Lista de Devolución")
            if btn_anadir_dev:
                nuevo_reg = {
                    'Marca': m_val.strip(),
                    'Fármaco': f_val.strip(),
                    'Dosificación': d_val.strip(),
                    'Tamaño envase': t_val.strip(),
                    'CN': cn_val.strip(),
                    'Lote': l_val.strip(),
                    'Caducidad': cad_val.strip(),
                    'Pastillas restantes': int(pastillas_input)
                }
                st.session_state["df_devolucion"] = pd.concat([st.session_state["df_devolucion"], pd.DataFrame([nuevo_reg])], ignore_index=True)
                st.success("¡Medicamento decodificado y añadido a la devolución correctamente!")

        if not st.session_state["df_devolucion"].empty:
            st.markdown("##### 📋 Listado de Medicación a Devolver")
            edited_devolucion = st.data_editor(
                st.session_state["df_devolucion"],
                use_container_width=True,
                hide_index=True,
                key="editor_tabla_devolucion"
            )
            st.session_state["df_devolucion"] = edited_devolucion
            
            st.markdown("---")
            col_pdf_dev, col_fin_dev = st.columns([1, 1])
            
            with col_pdf_dev:
                fecha_str_f = datetime.now().strftime('%Y%m%d_%H%M')
                nombre_archivo_pdf = f"Devolucion_{pac_obj['nombre'].replace(' ', '_')}_{fecha_str_f}.pdf"
                
                pdf_dev_bytes = generar_albaran_devolucion_pdf(
                    pac_obj['nombre'], 
                    pac_obj['ref'], 
                    st.session_state["df_devolucion"].to_dict(orient="records")
                )
                
                st.download_button(
                    label="📄 Imprimir Albarán de Devolución (PDF)",
                    data=pdf_dev_bytes, 
                    file_name=nombre_archivo_pdf,
                    mime="application/pdf", 
                    use_container_width=True
                )
                
            with col_fin_dev:
                if st.button("💾 Finalizar y Dar de Baja Definitiva", use_container_width=True):
                    etiqueta_borrar = pac_obj["etiqueta"]
                    if etiqueta_borrar in shared_data["lista_pacientes"]:
                        del shared_data["lista_pacientes"][etiqueta_borrar]
                    
                    st.success(f"¡Paciente '{pac_obj['nombre']}' dado de baja definitiva del sistema con éxito!")
                    st.session_state["devolucion_activa"] = False
                    st.session_state["paciente_a_baja_obj"] = None
                    st.session_state["df_devolucion"] = pd.DataFrame(columns=[
                        'Marca', 'Fármaco', 'Dosificación', 'Tamaño envase', 'CN', 'Lote', 'Caducidad', 'Pastillas restantes'
                    ])
                    time.sleep(1.5)
                    st.rerun()
        
        if st.button("❌ Cancelar Devolución"):
            st.session_state["devolucion_activa"] = False
            st.session_state["paciente_a_baja_obj"] = None
            st.session_state["df_devolucion"] = pd.DataFrame(columns=[
                'Marca', 'Fármaco', 'Dosificación', 'Tamaño envase', 'CN', 'Lote', 'Caducidad', 'Pastillas restantes'
            ])
            st.rerun()
                
    st.write("")
    if not st.session_state["devolucion_activa"]:
        if st.button("⬅ Volver al Menú Principal"):
            st.session_state["pagina"] = "inicio"
            st.rerun()

# ----------------------------------------------------
# GESTIÓN DE USUARIOS Y ROLES DINÁMICOS
# ----------------------------------------------------
elif st.session_state["pagina"] == "gestion_usuarios":
    if not (tiene_permiso(rol_actual, "usuarios") or tiene_permiso(rol_actual, "roles")):
        st.error("No tienes permisos para acceder a la gestión de usuarios y roles.")
        st.stop()
        
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 800;'>🔐 GESTIÓN DE USUARIOS Y ROLES</h2>", unsafe_allow_html=True)
    st.markdown("---")
    
    tab_usr, tab_nuevo_usr, tab_roles = st.tabs(["👥 Usuarios Activos", "➕ Crear Nuevo Usuario", "⚙️ Crear / Gestionar Roles"])
    
    MODULOS_DISPONIBLES = {
        "pacientes": "Acceso a Lista y Fichas de Pacientes",
        "altas": "Módulo de Altas de Pacientes",
        "bajas": "Módulo de Bajas de Pacientes",
        "propuesta": "Propuesta de Pedido (Bandeja Enfermería)",
        "pedidos_def": "Pedidos Definitivos y Albaranes PDF",
        "incidencias": "Panel de Incidencias",
        "usuarios": "Gestión de Usuarios",
        "roles": "Gestión de Roles y Permisos"
    }

    with tab_usr:
        st.subheader("Listado de Usuarios Registrados")
        for usr, info in list(USUARIOS_VALIDOS.items()):
            with st.expander(f"👤 {usr} — Rol: {info['rol'].upper()}"):
                with st.form(f"form_edit_user_{usr}"):
                    nuevo_pass = st.text_input("Nueva contraseña", value=info['clave'], type="password", key=f"pwd_{usr}")
                    
                    lista_roles_disponibles = list(ROLES_VALIDOS.keys())
                    rol_actual_idx = lista_roles_disponibles.index(info['rol']) if info['rol'] in lista_roles_disponibles else 0
                    nuevo_rol = st.selectbox("Rol asignado", lista_roles_disponibles, index=rol_actual_idx, key=f"rol_{usr}")
                    
                    col_u1, col_u2 = st.columns(2)
                    with col_u1:
                        guardar_cambios = st.form_submit_button("💾 Guardar Cambios")
                    with col_u2:
                        borrar_usr = st.form_submit_button("🗑️ Dar de Baja (Eliminar)")
                        
                    if guardar_cambios:
                        USUARIOS_VALIDOS[usr]["clave"] = nuevo_pass
                        USUARIOS_VALIDOS[usr]["rol"] = nuevo_rol
                        st.success(f"¡Usuario {usr} actualizado correctamente!")
                        time.sleep(0.8)
                        st.rerun()
                        
                    if borrar_usr:
                        if usr == st.session_state["usuario_autenticado"]:
                            st.error("No puedes eliminar tu propia cuenta mientras estás conectado.")
                        else:
                            del USUARIOS_VALIDOS[usr]
                            st.success(f"¡Usuario {usr} dado de baja con éxito!")
                            time.sleep(0.8)
                            st.rerun()

    with tab_nuevo_usr:
        st.subheader("Dar de Alta Nuevo Usuario")
        with st.form("form_nuevo_usuario"):
            nombre_nuevo = st.text_input("Nombre de Usuario (Login):")
            clave_nueva = st.text_input("Contraseña de acceso:", type="password")
            rol_nuevo = st.selectbox("Rol profesional:", list(ROLES_VALIDOS.keys()))
            
            crear_submit = st.form_submit_button("✅ Crear Usuario")
            if crear_submit:
                nom_limpio = nombre_nuevo.strip()
                if not nom_limpio or not clave_nueva.strip():
                    st.error("⚠️ El nombre de usuario y la contraseña no pueden estar vacíos.")
                elif nom_limpio in USUARIOS_VALIDOS:
                    st.error("⚠️ Ya existe un usuario con ese nombre.")
                else:
                    USUARIOS_VALIDOS[nom_limpio] = {
                        "clave": clave_nueva.strip(),
                        "rol": rol_nuevo
                    }
                    st.success(f"¡Usuario '{nom_limpio}' dado de alta con éxito!")
                    time.sleep(1)
                    st.rerun()

    with tab_roles:
        st.subheader("⚙️ Configuración de Roles y Permisos Granulares")
        
        with st.form("form_nuevo_rol"):
            st.markdown("##### ➕ Crear Nuevo Tipo de Rol")
            nombre_nuevo_rol = st.text_input("Nombre del nuevo Rol (ej: auxiliar, supervisor):")
            st.write("Selecciona los permisos que tendrá este rol:")
            
            permisos_seleccionados = {}
            for mod_key, mod_desc in MODULOS_DISPONIBLES.items():
                permisos_seleccionados[mod_key] = st.checkbox(mod_desc, value=False, key=f"chk_new_{mod_key}")
                
            crear_rol_btn = st.form_submit_button("Crear Nuevo Rol")
            if crear_rol_btn:
                r_limpio = nombre_nuevo_rol.strip().lower()
                if not r_limpio:
                    st.error("El nombre del rol no puede estar vacío.")
                elif r_limpio in ROLES_VALIDOS:
                    st.error("Ya existe un rol con ese nombre.")
                else:
                    mods_activos = [k for k, v in permisos_seleccionados.items() if v]
                    ROLES_VALIDOS[r_limpio] = mods_activos
                    st.success(f"¡Rol '{r_limpio}' creado con éxito!")
                    time.sleep(1)
                    st.rerun()

        st.markdown("---")
        st.markdown("##### 📝 Modificar Permisos de Roles Existentes")
        
        rol_a_editar = st.selectbox("Selecciona un rol para editar sus permisos:", list(ROLES_VALIDOS.keys()))
        if rol_a_editar:
            with st.form(f"form_editar_rol_{rol_a_editar}"):
                permisos_actuales = ROLES_VALIDOS[rol_a_editar]
                nuevos_permisos = {}
                for mod_key, mod_desc in MODULOS_DISPONIBLES.items():
                    tiene_actualmente = mod_key in permisos_actuales
                    nuevos_permisos[mod_key] = st.checkbox(mod_desc, value=tiene_actualmente, key=f"edit_chk_{rol_a_editar}_{mod_key}")
                    
                guardar_rol_btn = st.form_submit_button("💾 Guardar Permisos del Rol")
                if guardar_rol_btn:
                    ROLES_VALIDOS[rol_a_editar] = [k for k, v in nuevos_permisos.items() if v]
                    st.success(f"¡Permisos del rol '{rol_a_editar}' actualizados correctamente!")
                    time.sleep(0.8)
                    st.rerun()

    st.write("")
    if st.button("⬅ Volver al Menú Principal", key="btn_volver_usr"):
        st.session_state["pagina"] = "inicio"
        st.rerun()

# ----------------------------------------------------
# LISTA DE PACIENTES Y DETALLE
# ----------------------------------------------------
elif st.session_state["pagina"] == "lista_pacientes":
    if not tiene_permiso(rol_actual, "pacientes"):
        st.error("No tienes permiso para acceder a este módulo.")
        st.stop()
        
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 800;'>PACIENTES</h2>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 3, 1])
    with col2:
        if not lista_pacientes:
            st.error("No hay pacientes registrados actualmente.")
        else:
            busqueda = st.text_input("🔍 Buscar paciente por nombre o código:", "")
            pacientes_filtrados = [p for p in lista_pacientes.keys() if busqueda.lower() in p.lower()]
            for pac_key in pacientes_filtrados[:15]: 
                if st.button(pac_key, key=f"pac_{pac_key}", use_container_width=True):
                    st.session_state["paciente_seleccionado_key"] = pac_key
                    st.session_state["pagina"] = "detalle_paciente"
                    st.rerun()
        st.write("")
        if st.button("⬅ Volver al Menú Principal"):
            st.session_state["pagina"] = "inicio"
            st.rerun()

elif st.session_state["pagina"] == "detalle_paciente":
    if not tiene_permiso(rol_actual, "pacientes"):
        st.error("No tienes permiso para acceder a este módulo.")
        st.stop()
        
    pac_key = st.session_state["paciente_seleccionado_key"]
    info_paciente = lista_pacientes.get(pac_key)
    
    if info_paciente:
        es_admin_rol = (rol_actual == "admin")
        
        st.markdown(f"<h3 style='text-align: center; color: #1e293b; font-weight: 800;'>Ficha de Paciente: {info_paciente['nombre']}</h3>", unsafe_allow_html=True)
        
        if es_admin_rol:
            with st.form(f"form_editar_datos_{pac_key}"):
                st.markdown("##### 📝 Modificar Datos Personales")
                c_ed1, c_ed2, c_ed3 = st.columns(3)
                with c_ed1:
                    nuevo_nom_edit = st.text_input("Nombre completo:", value=info_paciente['nombre'])
                with c_ed2:
                    nueva_ref_edit = st.text_input("Código / Referencia:", value=info_paciente['ref'])
                with c_ed3:
                    nuevo_cip_edit = st.text_input("Código CIP:", value=info_paciente.get('cip', ''))
                
                guardar_datos_pers = st.form_submit_button("💾 Guardar Cambios Personales")
                if guardar_datos_pers:
                    if nuevo_nom_edit.strip():
                        nueva_etiqueta = f"{nueva_ref_edit.strip()} — {nuevo_nom_edit.strip()}"
                        info_paciente['nombre'] = nuevo_nom_edit.strip()
                        info_paciente['ref'] = nueva_ref_edit.strip()
                        info_paciente['cip'] = nuevo_cip_edit.strip()
                        
                        if nueva_etiqueta != pac_key:
                            shared_data["lista_pacientes"][nueva_etiqueta] = info_paciente
                            del shared_data["lista_pacientes"][pac_key]
                            st.session_state["paciente_seleccionado_key"] = nueva_etiqueta
                        
                        st.success("¡Datos personales actualizados correctamente!")
                        time.sleep(0.8)
                        st.rerun()
                    else:
                        st.error("El nombre no puede estar vacío.")
            st.markdown("---")
        else:
            cip_mostrar = info_paciente.get('cip', 'Sin asignar')
            st.markdown(f"<p style='text-align: center; font-size: 16px;'><b>Referencia:</b> {info_paciente['ref']} | <b>CIP:</b> {cip_mostrar}</p>", unsafe_allow_html=True)
        
        df_pac = info_paciente["datos"]
        
        if 'Medicamento' not in df_pac.columns: df_pac['Medicamento'] = ""
        if 'CN' not in df_pac.columns: df_pac['CN'] = ""
        if 'Posologia' not in df_pac.columns: df_pac['Posologia'] = ""
        if 'Ultima Entrega' not in df_pac.columns: df_pac['Ultima Entrega'] = ""
        if 'Fecha inicio' not in df_pac.columns: df_pac['Fecha inicio'] = ""
        if 'Pedido' not in df_pac.columns: df_pac['Pedido'] = False
        if 'Incidencia' not in df_pac.columns: df_pac['Incidencia'] = False
        
        st.markdown("##### 💊 Listado de Medicación y Tratamientos")
        if es_admin_rol:
            edited_df = st.data_editor(
                df_pac[['Medicamento', 'CN', 'Posologia', 'Ultima Entrega', 'Pedido', 'Incidencia']],
                use_container_width=True, hide_index=True, num_rows="dynamic", key=f"editor_{pac_key}",
                column_config={
                    "Ultima Entrega": st.column_config.TextColumn("Última Entrega", disabled=True),
                    "Pedido": st.column_config.CheckboxColumn("Incluir en Propuesta", default=False), 
                    "Incidencia": st.column_config.CheckboxColumn("Incidencia", default=False)
                }
            )
            info_paciente["datos"] = edited_df
        else:
            if not df_pac.empty:
                st.dataframe(df_pac[['Medicamento', 'CN', 'Posologia', 'Ultima Entrega']], use_container_width=True, hide_index=True)
            else:
                st.info("No hay medicación registrada para este paciente.")
        
        col_b1, col_b2, col_b3 = st.columns([1, 1.5, 1.5])
        with col_b1:
            if st.button("⬅ Volver", use_container_width=True):
                st.session_state["pagina"] = "lista_pacientes"
                st.rerun()
                
        if es_admin_rol and not df_pac.empty:
            with col_b2:
                if st.button("💾 Enviar a Propuesta de Pedido", use_container_width=True):
                    for pk, p_info in lista_pacientes.items():
                        df_p = p_info["datos"]
                        if df_p.empty: continue
                        marcados = df_p[df_p['Pedido'] == True]
                        for idx, row in marcados.iterrows():
                            if not any(i['cn'] == str(row['CN']) and i['paciente'] == p_info['nombre'] for i in shared_data["solicitud_pedido"]):
                                shared_data["solicitud_pedido"].append({
                                    "paciente": p_info['nombre'], "ref": p_info['ref'], "medicamento": row['Medicamento'],
                                    "cn": str(row['CN']), "posologia": row['Posologia'], "seleccion_enfermera": False,
                                    "datamatrix": "", "lote": "", "caducidad": ""
                                })
                        df_p['Pedido'] = False
                    st.success("Enviados a Propuesta de Pedido.")
            with col_b3:
                if st.button("⚠️ Enviar a Incidencias", use_container_width=True):
                    count_inc = 0
                    for pk, p_info in lista_pacientes.items():
                        df_p = p_info["datos"]
                        if df_p.empty: continue
                        marcadas = df_p[df_p['Incidencia'] == True]
                        for idx, row in marcadas.iterrows():
                            if not any(i['cn'] == str(row['CN']) and i['paciente'] == p_info['nombre'] for i in shared_data["incidencias_activas"]):
                                shared_data["incidencias_activas"].append({
                                    "paciente": p_info['nombre'], "ref": p_info['ref'], "medicamento": row['Medicamento'],
                                    "cn": str(row['CN']), "posologia": row['Posologia'], 
                                    "tipo_incidencia": None, "observaciones": "", "resuelta": False
                                })
                                count_inc += 1
                        df_p['Incidencia'] = False
                    st.success(f"¡{count_inc} enviadas al panel de Incidencias!")

elif st.session_state["pagina"] == "seleccion_productos_enfermera":
    if not tiene_permiso(rol_actual, "propuesta"):
        st.error("No tienes permiso para acceder a este módulo.")
        st.stop()
        
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 800;'>📦 PROPUESTA DE PEDIDO</h2>", unsafe_allow_html=True)
    if not shared_data["solicitud_pedido"]:
        st.info("No hay propuestas de pedido pendientes.")
    else:
        df_sol = pd.DataFrame(shared_data["solicitud_pedido"])
        edited_sol = st.data_editor(
            df_sol, use_container_width=True, hide_index=True,
            column_config={
                "paciente": st.column_config.TextColumn("Paciente", disabled=True),
                "medicamento": st.column_config.TextColumn("Medicamento", disabled=True),
                "seleccion_enfermera": st.column_config.CheckboxColumn("Seleccionar")
            }
        )
        shared_data["solicitud_pedido"] = edited_sol.to_dict(orient="records")
        
        if st.button("🚀 Solicitar Pedido Definitivo"):
            seleccionados = [i for i in shared_data["solicitud_pedido"] if i["seleccion_enfermera"] == True]
            for item in seleccionados:
                shared_data["pedidos_definitivos"].append(item)
            shared_data["solicitud_pedido"] = [i for i in shared_data["solicitud_pedido"] if i["seleccion_enfermera"] == False]
            st.success("Enviado al farmacéutico.")
            time.sleep(1)
            st.rerun()
    if st.button("⬅ Volver"):
        st.session_state["pagina"] = "inicio"
        st.rerun()

elif st.session_state["pagina"] == "solicitud_pedido_admin":
    if not tiene_permiso(rol_actual, "propuesta"):
        st.error("No tienes permiso para acceder a este módulo.")
        st.stop()
        
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 800;'>📦 PROPUESTA DE PEDIDO (Enviado a Enfermería)</h2>", unsafe_allow_html=True)
    if not shared_data["solicitud_pedido"]:
        st.info("No hay propuestas de pedido enviadas a enfermería actualmente.")
    else:
        df_sol = pd.DataFrame(shared_data["solicitud_pedido"])
        st.dataframe(df_sol[['paciente', 'ref', 'medicamento', 'cn', 'posologia']], use_container_width=True, hide_index=True)
    if st.button("⬅ Volver"):
        st.session_state["pagina"] = "inicio"
        st.rerun()

# ----------------------------------------------------
# PEDIDOS DEFINITIVOS CON ESCÁNER RÁPIDO DATAMATRIX
# ----------------------------------------------------
elif st.session_state["pagina"] == "pedidos_definitivos_admin":
    if not tiene_permiso(rol_actual, "pedidos_def"):
        st.error("No tienes permiso para acceder a este módulo.")
        st.stop()
        
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 800;'>🛒 PEDIDOS DEFINITIVOS Y ESCÁNER</h2>", unsafe_allow_html=True)
    if not shared_data["pedidos_definitivos"]:
        st.info("No hay pedidos definitivos en este momento.")
    else:
        # Escáner Rápido DataMatrix para autocompletar la tabla de pedidos
        def procesar_escaner_pedido():
            val = st.session_state.get("quick_dm_scan", "")
            if val:
                parsed = parsear_datamatrix(val)
                # Buscar el primer pedido definitivo sin DataMatrix registrado o coincidente por CN
                encontrado = False
                for item in shared_data["pedidos_definitivos"]:
                    if not item.get("datamatrix"):
                        # Si coincide el CN o si está libre, lo rellenamos automáticamente
                        if not item.get("cn") or str(item.get("cn")) == str(parsed['cn']):
                            item["datamatrix"] = val
                            item["lote"] = parsed['lote'] if parsed['lote'] else "L001"
                            item["caducidad"] = parsed['caducidad'] if parsed['caducidad'] else "12/2028"
                            encontrado = True
                            break
                if not encontrado and shared_data["pedidos_definitivos"]:
                    # Rellenar al menos la primera línea disponible
                    shared_data["pedidos_definitivos"][0]["datamatrix"] = val
                    if parsed['lote']: shared_data["pedidos_definitivos"][0]["lote"] = parsed['lote']
                    if parsed['caducidad']: shared_data["pedidos_definitivos"][0]["caducidad"] = parsed['caducidad']
                st.session_state["quick_dm_scan"] = ""

        st.markdown("##### ⚡ Escáner Rápido de DataMatrix para Pedidos")
        st.text_input(
            "Escanee aquí el código DataMatrix del medicamento para autocompletar la siguiente línea:",
            key="quick_dm_scan",
            on_change=procesar_escaner_pedido
        )

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
        col_pdf, col_vaciar = st.columns([1, 1])
        with col_pdf:
            pdf_bytes = generar_albaran_pdf(shared_data["pedidos_definitivos"])
            
            if st.download_button(
                label="📄 Imprimir Albarán de Entrega (PDF)",
                data=pdf_bytes, file_name=f"Albaran_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                mime="application/pdf", use_container_width=True
            ):
                pass
            
            if st.button("📌 Actualizar Fechas de Última Entrega y Borrar Pedidos", use_container_width=True):
                fecha_entrega_actual = datetime.now().strftime("%d/%m/%Y %H:%M")
                for item_ped in shared_data["pedidos_definitivos"]:
                    nom_pac_ped = item_ped.get("paciente")
                    cn_ped = str(item_ped.get("cn"))
                    
                    for p_key, p_val in shared_data["lista_pacientes"].items():
                        if p_val["nombre"] == nom_pac_ped:
                            df_pac_meds = p_val["datos"]
                            if not df_pac_meds.empty and 'CN' in df_pac_meds.columns:
                                mask = df_pac_meds['CN'].astype(str).str.strip() == cn_ped.strip()
                                df_pac_meds.loc[mask, 'Ultima Entrega'] = fecha_entrega_actual
                
                shared_data["pedidos_definitivos"] = []
                st.success("¡Fechas de última entrega actualizadas en las fichas de los pacientes correctamente!")
                time.sleep(1)
                st.rerun()

    if st.button("⬅ Volver al Menú Principal"):
        st.session_state["pagina"] = "inicio"
        st.rerun()

elif st.session_state["pagina"] == "incidencias":
    if not tiene_permiso(rol_actual, "incidencias"):
        st.error("No tienes permiso para acceder a este módulo.")
        st.stop()
        
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 800;'>⚠️ PANEL DE INCIDENCIAS GLOBAL</h2>", unsafe_allow_html=True)
    st.markdown("---")
    
    if not shared_data["incidencias_activas"]:
        st.info("No hay incidencias activas reportadas en este momento.")
    else:
        df_inc = pd.DataFrame(shared_data["incidencias_activas"])
        
        if rol_actual == "admin":
            st.markdown("Edite el **Tipo de incidencia** u **Observaciones**. Para eliminarla, marque la casilla **'Solventar'** y confirme.")
            
            opciones_incidencia = [
                "Falta de receta electrónica",
                "Modificar posología",
                "Medicación adelantada",
                "Lo consume?",
                "Falta de abastecimiento",
                "Otra"
            ]
            
            edited_inc = st.data_editor(
                df_inc,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "paciente": st.column_config.TextColumn("Paciente", disabled=True),
                    "ref": st.column_config.TextColumn("Ref.", disabled=True),
                    "medicamento": st.column_config.TextColumn("Medicamento", disabled=True),
                    "cn": st.column_config.TextColumn("C.N.", disabled=True),
                    "posologia": st.column_config.TextColumn("Posología", disabled=True),
                    "tipo_incidencia": st.column_config.SelectboxColumn(
                        "Tipo Incidencia",
                        help="Seleccione el motivo de la incidencia",
                        options=opciones_incidencia,
                        required=False
                    ),
                    "observaciones": st.column_config.TextColumn("Observaciones"),
                    "resuelta": st.column_config.CheckboxColumn("Solventar (Borrar)")
                },
                key="editor_inc_admin"
            )
            shared_data["incidencias_activas"] = edited_inc.to_dict(orient="records")
            
            if st.button("✅ Confirmar Incidencias Resueltas"):
                shared_data["incidencias_activas"] = [i for i in shared_data["incidencias_activas"] if not i.get("resuelta", False)]
                st.success("Listado actualizado. Se han borrado las incidencias resueltas.")
                time.sleep(1)
                st.rerun()
        else:
            st.markdown("Lista de incidencias reportadas por Farmacia para su conocimiento.")
            st.dataframe(
                df_inc[['paciente', 'medicamento', 'tipo_incidencia', 'observaciones']],
                use_container_width=True,
                hide_index=True
            )
            
    st.write("")
    if st.button("⬅ Volver al Menú"):
        st.session_state["pagina"] = "inicio"
        st.rerun()
        st.rerun()

elif st.session_state["pagina"] == "incidencias":
    if not tiene_permiso(rol_actual, "incidencias"):
        st.error("No tienes permiso para acceder a este módulo.")
        st.stop()
        
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 800;'>⚠️ PANEL DE INCIDENCIAS GLOBAL</h2>", unsafe_allow_html=True)
    st.markdown("---")
    
    if not shared_data["incidencias_activas"]:
        st.info("No hay incidencias activas reportadas en este momento.")
    else:
        df_inc = pd.DataFrame(shared_data["incidencias_activas"])
        
        if rol_actual == "admin":
            st.markdown("Edite el **Tipo de incidencia** u **Observaciones**. Para eliminarla, marque la casilla **'Solventar'** y confirme.")
            
            opciones_incidencia = [
                "Falta de receta electrónica",
                "Modificar posología",
                "Medicación adelantada",
                "Lo consume?",
                "Falta de abastecimiento",
                "Otra"
            ]
            
            edited_inc = st.data_editor(
                df_inc,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "paciente": st.column_config.TextColumn("Paciente", disabled=True),
                    "ref": st.column_config.TextColumn("Ref.", disabled=True),
                    "medicamento": st.column_config.TextColumn("Medicamento", disabled=True),
                    "cn": st.column_config.TextColumn("C.N.", disabled=True),
                    "posologia": st.column_config.TextColumn("Posología", disabled=True),
                    "tipo_incidencia": st.column_config.SelectboxColumn(
                        "Tipo Incidencia",
                        help="Seleccione el motivo de la incidencia",
                        options=opciones_incidencia,
                        required=False
                    ),
                    "observaciones": st.column_config.TextColumn("Observaciones"),
                    "resuelta": st.column_config.CheckboxColumn("Solventar (Borrar)")
                },
                key="editor_inc_admin"
            )
            shared_data["incidencias_activas"] = edited_inc.to_dict(orient="records")
            
            if st.button("✅ Confirmar Incidencias Resueltas"):
                shared_data["incidencias_activas"] = [i for i in shared_data["incidencias_activas"] if not i.get("resuelta", False)]
                st.success("Listado actualizado. Se han borrado las incidencias resueltas.")
                time.sleep(1)
                st.rerun()
        else:
            st.markdown("Lista de incidencias reportadas por Farmacia para su conocimiento.")
            st.dataframe(
                df_inc[['paciente', 'medicamento', 'tipo_incidencia', 'observaciones']],
                use_container_width=True,
                hide_index=True
            )
            
    st.write("")
    if st.button("⬅ Volver al Menú"):
        st.session_state["pagina"] = "inicio"
        st.rerun()
