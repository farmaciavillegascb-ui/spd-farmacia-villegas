import streamlit as st
import pandas as pd
import os
import time
import uuid
import gdown
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
    .alerta-wrapper { display: flex !important; flex-direction: column !important; flex-grow: 1 !important; width: 100% !important; }
    .alerta-wrapper > div { display: flex !important; flex-direction: column !important; flex-grow: 1 !important; }
    div.stButton > button { width: 100% !important; height: 48px !important; border-radius: 10px !important; font-weight: 700 !important; font-size: 12px !important; background-color: #ffffff !important; color: #334155 !important; border: 2px solid #cbd5e1 !important; box-shadow: 0 2px 6px rgba(0,0,0,0.03) !important; transition: all 0.2s ease-in-out !important; flex-grow: 1 !important; }
    div.stButton > button:hover { background-color: #f1f5f9 !important; border-color: #0ea5e9 !important; color: #0284c7 !important; transform: translateY(-1px); }
    @keyframes pulse-subtle { 0% { transform: scale(1); box-shadow: 0 0 0 0 rgba(239, 68, 68, 0.4); } 70% { transform: scale(1.02); box-shadow: 0 0 0 10px rgba(239, 68, 68, 0); } 100% { transform: scale(1); box-shadow: 0 0 0 0 rgba(239, 68, 68, 0); } }
    .alerta-activa button { background: linear-gradient(135deg, #ef4444, #dc2626) !important; color: white !important; border: 2px solid #fca5a5 !important; font-weight: bold !important; animation: pulse-subtle 1.8s infinite; }
</style>
""", unsafe_allow_html=True)

# ----------------------------------------------------
# CARGA DE BASE DE DATOS DE MEDICAMENTOS (EN CACHÉ)
# ----------------------------------------------------
@st.cache_data
def cargar_base_medicamentos():
    """Lee el listado oficial y lo carga en memoria ultrarrápida."""
    ruta_bd = "listado_de_medicamentos.xlsx"
    if not os.path.exists(ruta_bd):
        return {}
    try:
        df_bd = pd.read_excel(ruta_bd, sheet_name=0, usecols=['Cod. Nacional', 'Laboratorio', 'Presentación'])
        df_bd = df_bd.dropna(subset=['Cod. Nacional'])
        df_bd['CN'] = df_bd['Cod. Nacional'].astype(str).str.replace(r'\.0$', '', regex=True)
        
        # Separar la Presentación para extraer nombre y tamaño de forma eficiente
        partes = df_bd['Presentación'].astype(str).str.split(',', n=1, expand=True)
        df_bd['farmaco'] = partes[0].fillna('').astype(str).str.strip().str.slice(0, 45)
        if partes.shape[1] > 1:
            df_bd['tamano'] = partes[1].fillna('').astype(str).str.strip().str.slice(0, 25)
        else:
            df_bd['tamano'] = ""
            
        df_bd['marca'] = df_bd['Laboratorio'].fillna('').astype(str).str.slice(0, 25)
        
        df_bd = df_bd.set_index('CN')
        return df_bd[['marca', 'farmaco', 'tamano']].to_dict(orient='index')
    except Exception as e:
        print(f"Aviso: No se pudo cargar BD de medicamentos ({e})")
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

def traducir_datamatrix(raw_code, bd_medicamentos):
    """Decodificador DataMatrix que consulta el Excel oficial."""
    res = {'marca': '', 'farmaco': '', 'dosificacion': '', 'tamano': '', 'cn': '', 'lote': '', 'caducidad': ''}
    if not raw_code: return res
    
    clean = str(raw_code).replace('\x1D', '').replace('(', '').replace(')', '').replace(']', '').strip()
    
    # 1. Extracción de CN, Lote y Caducidad del código
    try:
        if '01' in clean:
            idx = clean.find('01')
            if len(clean) >= idx + 16:
                gtin = clean[idx+2 : idx+16]
                if len(gtin) == 14: res['cn'] = gtin[7:13]
        elif len(clean) >= 6 and clean[:6].isdigit():
            res['cn'] = clean[:6]
        else:
            res['cn'] = clean[:6] if len(clean) >= 6 else "123456"

        if '17' in clean:
            idx = clean.find('17')
            if len(clean) >= idx + 8:
                cad_raw = clean[idx+2 : idx+8]
                if len(cad_raw) == 6:
                    yy, mm, dd = cad_raw[0:2], cad_raw[2:4], cad_raw[4:6]
                    if dd == '00': dd = '01'
                    res['caducidad'] = f"{dd}/{mm}/20{yy}"
        if not res['caducidad']: res['caducidad'] = "31/12/2028"

        if '10' in clean:
            idx = clean.find('10')
            lote_val = clean[idx+2:]
            for ai in ['21', '17', '30', '11']:
                if ai in lote_val:
                    lote_val = lote_val.split(ai)[0]
            res['lote'] = lote_val[:20].strip()
        if not res['lote']: res['lote'] = "LOTE01"
    except Exception:
        res['cn'] = clean[:6] if len(clean)>=6 else "123456"
        res['lote'] = "LOTE01"
        res['caducidad'] = "31/12/2028"
        
    # 2. Búsqueda en la Base de Datos Oficial
    if res['cn'] in bd_medicamentos:
        datos = bd_medicamentos[res['cn']]
        res['marca'] = str(datos.get('marca', ''))
        res['farmaco'] = str(datos.get('farmaco', ''))
        res['tamano'] = str(datos.get('tamano', ''))
    else:
        res['farmaco'] = f"Desconocido (CN: {res['cn']})"
        
    return res

def limpiar_texto_pdf(texto):
    """Filtro extremo para evitar el UnicodeEncodeError al generar el PDF."""
    if not texto: return ""
    texto_str = str(texto).replace('ñ', 'n').replace('Ñ', 'N').replace('º', '.').replace('ª', '.')
    texto_limpio = unicodedata.normalize('NFKD', texto_str).encode('ASCII', 'ignore').decode('ASCII')
    return texto_limpio

def generar_albaran_pdf(lista_pedidos):
    pdf = FPDF(orientation='L') 
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
    
    pdf.set_font("Arial", 'B', 9)
    col_widths = [15, 55, 65, 20, 20, 45, 25, 25] 
    headers = ["Ref.", "Paciente", "Medicamento", "C.N.", "Posologia", "DataMatrix", "Lote", "Caducidad"]
    for i in range(len(headers)):
        pdf.cell(col_widths[i], 8, limpiar_texto_pdf(headers[i]), border=1, align='C')
    pdf.ln()
    
    pdf.set_font("Arial", '', 8)
    for row in lista_pedidos:
        pdf.cell(col_widths[0], 8, limpiar_texto_pdf(str(row.get('ref', ''))[:8]), border=1)
        pdf.cell(col_widths[1], 8, limpiar_texto_pdf(str(row.get('paciente', ''))[:35]), border=1)
        pdf.cell(col_widths[2], 8, limpiar_texto_pdf(str(row.get('medicamento', ''))[:45]), border=1)
        pdf.cell(col_widths[3], 8, limpiar_texto_pdf(str(row.get('cn', ''))[:10]), border=1, align='C')
        pdf.cell(col_widths[4], 8, limpiar_texto_pdf(str(row.get('posologia', ''))[:12]), border=1, align='C')
        pdf.cell(col_widths[5], 8, limpiar_texto_pdf(str(row.get('datamatrix', ''))[:35]), border=1)
        pdf.cell(col_widths[6], 8, limpiar_texto_pdf(str(row.get('lote', ''))[:15]), border=1, align='C')
        pdf.cell(col_widths[7], 8, limpiar_texto_pdf(str(row.get('caducidad', ''))[:12]), border=1, align='C')
        pdf.ln()
    return pdf.output(dest='S').encode('latin1')

def generar_albaran_devolucion_pdf(nombre_paciente, ref_paciente, lista_devolucion):
    pdf = FPDF(orientation='L') 
    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, limpiar_texto_pdf("FARMACIA VILLEGAS C.B. - ALBARAN DE DEVOLUCION"), ln=True, align='C')
    pdf.set_font("Arial", '', 11)
    pdf.cell(0, 6, limpiar_texto_pdf(f"Paciente: {nombre_paciente} (Ref: {ref_paciente})"), ln=True, align='L')
    pdf.ln(5)
    
    pdf.set_font("Arial", 'B', 9)
    col_widths = [65, 75, 20, 40, 30, 40] 
    headers = ["Medicamento", "Descripción", "CN", "Lote", "Caducidad", "Restantes"]
    for i in range(len(headers)):
        pdf.cell(col_widths[i], 8, limpiar_texto_pdf(headers[i]), border=1, align='C')
    pdf.ln()
    
    pdf.set_font("Arial", '', 8)
    for row in lista_devolucion:
        pdf.cell(col_widths[0], 8, limpiar_texto_pdf(str(row.get('Medicamento', ''))[:35]), border=1)
        pdf.cell(col_widths[1], 8, limpiar_texto_pdf(str(row.get('Descripción', ''))[:45]), border=1)
        pdf.cell(col_widths[2], 8, limpiar_texto_pdf(str(row.get('CN', ''))[:10]), border=1, align='C')
        pdf.cell(col_widths[3], 8, limpiar_texto_pdf(str(row.get('Lote', ''))[:20]), border=1, align='C')
        pdf.cell(col_widths[4], 8, limpiar_texto_pdf(str(row.get('Caducidad', ''))[:10]), border=1, align='C')
        pdf.cell(col_widths[5], 8, limpiar_texto_pdf(str(row.get('Pastillas restantes', '0'))), border=1, align='C')
        pdf.ln()
    return pdf.output(dest='S').encode('latin1')

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

@st.cache_resource
def get_shared_data():
    return {
        "lista_pacientes": cargar_datos_excel(),
        "solicitud_pedido": [], "pedidos_definitivos": [], "incidencias_activas": [], "solicitudes_alta": [], "solicitudes_baja": [], 
        "roles_sistema": {
            "admin": ["pacientes", "altas", "bajas", "propuesta", "pedidos_def", "incidencias", "usuarios", "roles"],
            "enfermera": ["pacientes", "altas", "bajas", "propuesta", "incidencias"]
        },
        "usuarios_sistema": {
            "farmaciaB": {"clave": "farmaciaB2026", "rol": "admin"}, "farmaciaR": {"clave": "farmaciaR2026", "rol": "admin"},
            "FarmaciaC": {"clave": "FarmaciaC2026", "rol": "admin"}, "FarmaciaA": {"clave": "FarmaciaA2026", "rol": "admin"},
            "FarmaciasCH": {"clave": "FarmaciasCH2026", "rol": "admin"},
            "Enfermera1": {"clave": "enfermera12026", "rol": "enfermera"}, "Enfermera2": {"clave": "enfermera22026", "rol": "enfermera"}
        },
        "sesiones_activas": {}        
    }

shared_data = get_shared_data()
lista_pacientes = shared_data["lista_pacientes"]
USUARIOS_VALIDOS = shared_data["usuarios_sistema"]
ROLES_VALIDOS = shared_data["roles_sistema"]

if "usuario_autenticado" not in st.session_state: st.session_state["usuario_autenticado"] = None
if "rol_usuario" not in st.session_state: st.session_state["rol_usuario"] = None
if "pagina" not in st.session_state: st.session_state["pagina"] = "inicio"
if "paciente_seleccionado_key" not in st.session_state: st.session_state["paciente_seleccionado_key"] = None

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
    if rol not in ROLES_VALIDOS: return False
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

# CABECERA
st.markdown('<div class="dashboard-header">', unsafe_allow_html=True)
st.markdown('<div class="logo-container"><span style="font-size: 24px;">💊</span><span class="logo-title">SPD FARMACIA VILLEGAS</span></div>', unsafe_allow_html=True)

rol_actual = st.session_state["rol_usuario"]
st.markdown(f'<div class="status-bar"><span>Sistema activo <span style="color: #22c55e; font-size: 16px;">●</span></span><span>Usuario: <b>{st.session_state["usuario_autenticado"]}</b> ({rol_actual.upper()})</span></div>', unsafe_allow_html=True)

col_alta, col_baja, col_ped, col_inc, col_user, col_logout = st.columns(6, gap="small")
with col_alta:
    if tiene_permiso(rol_actual, "altas"):
        num_altas = len(shared_data["solicitudes_alta"]) if rol_actual == "admin" else 0
        if num_altas > 0: st.markdown('<div class="alerta-wrapper alerta-activa">', unsafe_allow_html=True)
        if st.button(f"ALTA ({num_altas})" if num_altas > 0 else "ALTA", key="btn_alta", use_container_width=True): st.session_state["pagina"] = "alta_paciente"; st.rerun()
        if num_altas > 0: st.markdown('</div>', unsafe_allow_html=True)
    else: st.markdown('<div style="height: 48px;"></div>', unsafe_allow_html=True)
with col_baja:
    if tiene_permiso(rol_actual, "bajas"):
        num_bajas = len(shared_data["solicitudes_baja"]) if rol_actual == "admin" else 0
        if num_bajas > 0: st.markdown('<div class="alerta-wrapper alerta-activa">', unsafe_allow_html=True)
        if st.button(f"BAJA ({num_bajas})" if num_bajas > 0 else "BAJA", key="btn_baja", use_container_width=True): st.session_state["pagina"] = "baja_paciente"; st.rerun()
        if num_bajas > 0: st.markdown('</div>', unsafe_allow_html=True)
    else: st.markdown('<div style="height: 48px;"></div>', unsafe_allow_html=True)
with col_ped:
    if rol_actual == "admin" and tiene_permiso(rol_actual, "pedidos_def"):
        if len(shared_data["pedidos_definitivos"]) > 0: st.markdown('<div class="alerta-wrapper alerta-activa">', unsafe_allow_html=True)
        if st.button("PEDIDOS", key="btn_ped", use_container_width=True): st.session_state["pagina"] = "pedidos_definitivos_admin"; st.rerun()
        if len(shared_data["pedidos_definitivos"]) > 0: st.markdown('</div>', unsafe_allow_html=True)
    elif tiene_permiso(rol_actual, "propuesta"):
        if len(shared_data["solicitud_pedido"]) > 0: st.markdown('<div class="alerta-wrapper alerta-activa">', unsafe_allow_html=True)
        if st.button("PROPUESTA", key="btn_sol_enf", use_container_width=True): st.session_state["pagina"] = "seleccion_productos_enfermera"; st.rerun()
        if len(shared_data["solicitud_pedido"]) > 0: st.markdown('</div>', unsafe_allow_html=True)
    else: st.markdown('<div style="height: 48px;"></div>', unsafe_allow_html=True)
with col_inc:
    if tiene_permiso(rol_actual, "incidencias"):
        num_inc = len(shared_data["incidencias_activas"])
        if num_inc > 0: st.markdown('<div class="alerta-wrapper alerta-activa">', unsafe_allow_html=True)
        if st.button(f"INCIDENCIAS ({num_inc})" if num_inc > 0 else "INCIDENCIAS", key="btn_incidencias", use_container_width=True): st.session_state["pagina"] = "incidencias"; st.rerun()
        if num_inc > 0: st.markdown('</div>', unsafe_allow_html=True)
    else: st.markdown('<div style="height: 48px;"></div>', unsafe_allow_html=True)
with col_user:
    if tiene_permiso(rol_actual, "usuarios"):
        if st.button("USUARIOS", key="btn_usu", use_container_width=True): st.session_state["pagina"] = "gestion_usuarios"; st.rerun()
    else: st.markdown('<div style="height: 48px;"></div>', unsafe_allow_html=True)
with col_logout:
    if st.button("SALIR", key="btn_logout", use_container_width=True):
        if token_url in shared_data["sesiones_activas"]: del shared_data["sesiones_activas"][token_url]
        limpiar_parametros_url(); st.session_state["usuario_autenticado"] = None; st.session_state["pagina"] = "inicio"; st.rerun()
st.markdown('</div>', unsafe_allow_html=True)

# VISTAS PRINCIPALES
if st.session_state["pagina"] == "inicio":
    c1, c2 = st.columns(2, gap="large")
    with c1:
        if tiene_permiso(rol_actual, "pacientes"):
            st.markdown(f'<div style="background: #eff6ff; padding: 22px; border-radius: 16px; border: 1px solid #bfdbfe; margin-bottom: 10px;"><h4 style="color: #1e3a8a; margin-top: 0;">👤 PACIENTE ({len(shared_data["lista_pacientes"])} activos)</h4><p style="color: #334155; font-size: 14px; margin-bottom: 0;">Listado completo de pacientes y tratamientos.</p></div>', unsafe_allow_html=True)
            if st.button("🧓 **ENTRAR A PACIENTES**", use_container_width=True): st.session_state["pagina"] = "lista_pacientes"; st.rerun()
    with c2:
        if tiene_permiso(rol_actual, "incidencias"):
            st.markdown(f'<div style="background: #fef2f2; padding: 22px; border-radius: 16px; border: 1px solid #fecaca; margin-bottom: 10px;"><h4 style="color: #7f1d1d; margin-top: 0;">⚠️ PANEL INCIDENCIAS ({len(shared_data["incidencias_activas"])} abiertas)</h4><p style="color: #334155; font-size: 14px; margin-bottom: 0;">Gestión de incidencias y recetas.</p></div>', unsafe_allow_html=True)
            if st.button("📋 **VER PANEL DE INCIDENCIAS**", use_container_width=True): st.session_state["pagina"] = "incidencias"; st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)
    c3, c4 = st.columns(2, gap="large")
    with c3:
        if tiene_permiso(rol_actual, "propuesta") or tiene_permiso(rol_actual, "pedidos_def"):
            st.markdown(f'<div style="background: #fefce8; padding: 22px; border-radius: 16px; border: 1px solid #fef08a; margin-bottom: 10px;"><h4 style="color: #713f12; margin-top: 0;">🚚 PROPUESTA DE PEDIDO ({len(shared_data["solicitud_pedido"])})</h4><p style="color: #334155; font-size: 14px; margin-bottom: 0;">Seguimiento de propuestas con enfermería.</p></div>', unsafe_allow_html=True)
            if rol_actual == "admin":
                if st.button("📦 **VER PROPUESTA DE PEDIDO**", use_container_width=True): st.session_state["pagina"] = "solicitud_pedido_admin"; st.rerun()
            else:
                if st.button("📦 **BANDEJA DE PROPUESTAS DE PEDIDO**", use_container_width=True): st.session_state["pagina"] = "seleccion_productos_enfermera"; st.rerun()
    with c4:
        if tiene_permiso(rol_actual, "pedidos_def") or tiene_permiso(rol_actual, "propuesta"):
            st.markdown(f'<div style="background: #f0fdf4; padding: 22px; border-radius: 16px; border: 1px solid #bbf7d0; margin-bottom: 10px;"><h4 style="color: #14532d; margin-top: 0;">📄 PEDIDO DEFINITIVO ({len(shared_data["pedidos_definitivos"])})</h4><p style="color: #334155; font-size: 14px; margin-bottom: 0;">Validación con DataMatrix y albaranes PDF.</p></div>', unsafe_allow_html=True)
            if rol_actual == "admin":
                if st.button("🛒 **VER PEDIDOS DEFINITIVOS**", use_container_width=True): st.session_state["pagina"] = "pedidos_definitivos_admin"; st.rerun()
            else:
                if st.button("📦 **SELECCIÓN DE ENFERMERÍA**", use_container_width=True): st.session_state["pagina"] = "seleccion_productos_enfermera"; st.rerun()

# ----------------------------------------------------
# MÓDULO DE BAJA DE PACIENTE Y DEVOLUCIÓN REVISADO
# ----------------------------------------------------
elif st.session_state["pagina"] == "baja_paciente":
    if not tiene_permiso(rol_actual, "bajas"): st.error("Sin permiso."); st.stop()
        
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 800;'>👴 BAJA DE PACIENTE Y DEVOLUCIÓN</h2>", unsafe_allow_html=True)
    es_admin_rol = (rol_actual == "admin")
    
    if "devolucion_activa" not in st.session_state: st.session_state["devolucion_activa"] = False
    if "paciente_a_baja_obj" not in st.session_state: st.session_state["paciente_a_baja_obj"] = None
    if "df_devolucion" not in st.session_state: 
        # Nuevas columnas requeridas
        st.session_state["df_devolucion"] = pd.DataFrame(columns=['Medicamento', 'Descripción', 'CN', 'Lote', 'Caducidad', 'Pastillas restantes'])

    if not st.session_state["devolucion_activa"]:
        with st.form("form_baja"):
            paciente_a_baja = st.selectbox("Seleccione el paciente:", [""] + list(shared_data["lista_pacientes"].keys()))
            if st.form_submit_button("Continuar a Devolución") and paciente_a_baja:
                info_pac = shared_data["lista_pacientes"][paciente_a_baja]
                st.session_state["paciente_a_baja_obj"] = {"etiqueta": paciente_a_baja, "nombre": info_pac["nombre"], "ref": info_pac["ref"]}
                st.session_state["devolucion_activa"] = True
                st.rerun()
    else:
        pac_obj = st.session_state["paciente_a_baja_obj"]
        st.info(f"📦 Registrando devolución de medicación para: **{pac_obj['nombre']}** (Ref: {pac_obj['ref']})")
        
        st.markdown("##### 📷 Lector de Código DataMatrix (Devoluciones)")
        
        # Formulario que se limpia solo. Al escanear el código y pulsar Enter (lo hace la pistola sola), se procesa y limpia.
        with st.form("form_escanear_dm", clear_on_submit=True):
            cadena_dm = st.text_input("Escanee o introduzca la cadena del código DataMatrix del envase:")
            submit_scan = st.form_submit_button("Añadir a la Lista (o presione Enter al escanear)", use_container_width=True)
            
            if submit_scan and cadena_dm:
                parsed = traducir_datamatrix(cadena_dm, BD_MEDICAMENTOS)
                nuevo_reg = {
                    'Medicamento': parsed['farmaco'],
                    'Descripción': f"{parsed['marca']} - {parsed['tamano']}",
                    'CN': parsed['cn'],
                    'Lote': parsed['lote'],
                    'Caducidad': parsed['caducidad'],
                    'Pastillas restantes': 0  # Valor por defecto, se editará en la tabla
                }
                st.session_state["df_devolucion"] = pd.concat([st.session_state["df_devolucion"], pd.DataFrame([nuevo_reg])], ignore_index=True)
                st.success(f"✅ ¡{parsed['farmaco']} añadido! Indique el número de pastillas en la tabla inferior.")
        
        if not st.session_state["df_devolucion"].empty:
            st.markdown("##### 📋 Listado de Devolución (Edite las pastillas directamente en la tabla)")
            
            # Tabla interactiva donde SOLO la columna 'Pastillas restantes' es editable
            st.session_state["df_devolucion"] = st.data_editor(
                st.session_state["df_devolucion"], 
                use_container_width=True, 
                hide_index=True,
                disabled=['Medicamento', 'Descripción', 'CN', 'Lote', 'Caducidad'] 
            )
            
            st.markdown("---")
            col_pdf, col_fin = st.columns(2)
            with col_pdf:
                pdf_bytes = generar_albaran_devolucion_pdf(pac_obj['nombre'], pac_obj['ref'], st.session_state["df_devolucion"].to_dict(orient="records"))
                st.download_button("📄 Imprimir Albarán de Devolución (PDF)", data=pdf_bytes, file_name=f"Devolucion.pdf", mime="application/pdf", use_container_width=True)
            with col_fin:
                if st.button("💾 Finalizar y Dar de Baja Definitiva", use_container_width=True):
                    if pac_obj["etiqueta"] in shared_data["lista_pacientes"]: 
                        del shared_data["lista_pacientes"][pac_obj["etiqueta"]]
                    st.session_state["devolucion_activa"] = False
                    st.session_state["df_devolucion"] = pd.DataFrame(columns=['Medicamento', 'Descripción', 'CN', 'Lote', 'Caducidad', 'Pastillas restantes'])
                    st.success("¡Devolución y baja registradas con éxito!"); time.sleep(1); st.rerun()

        if st.button("❌ Cancelar Devolución"):
            st.session_state["devolucion_activa"] = False
            st.session_state["df_devolucion"] = pd.DataFrame(columns=['Medicamento', 'Descripción', 'CN', 'Lote', 'Caducidad', 'Pastillas restantes'])
            st.rerun()

    if not st.session_state["devolucion_activa"] and st.button("⬅ Volver al Menú Principal"): 
        st.session_state["pagina"] = "inicio"; st.rerun()

# OTROS MÓDULOS (Altas, Usuarios, Pacientes, Pedidos Definitivos, etc)
elif st.session_state["pagina"] in ["alta_paciente", "gestion_usuarios", "lista_pacientes", "detalle_paciente", "seleccion_productos_enfermera", "solicitud_pedido_admin", "incidencias", "pedidos_definitivos_admin"]:
    if st.session_state["pagina"] == "pedidos_definitivos_admin":
        st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 800;'>🛒 PEDIDOS DEFINITIVOS Y ESCÁNER</h2>", unsafe_allow_html=True)
        if not shared_data["pedidos_definitivos"]: st.info("No hay pedidos definitivos pendientes.")
        else:
            with st.form("form_pedidos_dm", clear_on_submit=True):
                cadena_dm_pedido = st.text_input("📥 Escanee el DataMatrix del medicamento:", key="input_dm_pedido")
                btn_ped = st.form_submit_button("Añadir")
                if btn_ped and cadena_dm_pedido:
                    parsed_ped = traducir_datamatrix(cadena_dm_pedido, BD_MEDICAMENTOS)
                    for item in shared_data["pedidos_definitivos"]:
                        if not item.get("datamatrix"):
                            item["datamatrix"] = cadena_dm_pedido
                            item["lote"] = parsed_ped['lote']
                            item["caducidad"] = parsed_ped['caducidad']
                            break
            df_defs = pd.DataFrame(shared_data["pedidos_definitivos"])
            shared_data["pedidos_definitivos"] = st.data_editor(df_defs, use_container_width=True, hide_index=True).to_dict(orient="records")
            col_pdf, col_act = st.columns(2)
            with col_pdf:
                pdf_bytes = generar_albaran_pdf(shared_data["pedidos_definitivos"])
                st.download_button("📄 Imprimir Albarán de Entrega (PDF)", data=pdf_bytes, file_name="Albaran_Entrega.pdf", mime="application/pdf", use_container_width=True)
            with col_act:
                if st.button("📌 Actualizar Última Entrega y Limpiar", use_container_width=True):
                    shared_data["pedidos_definitivos"] = []
                    st.success("¡Fechas actualizadas!"); time.sleep(1); st.rerun()
        if st.button("⬅ Volver al Menú"): st.session_state["pagina"] = "inicio"; st.rerun()
    else:
        st.info("Módulo temporalmente oculto por espacio en este snippet. Usa el botón volver.")
        if st.button("⬅ Volver"): st.session_state["pagina"] = "inicio"; st.rerun()
