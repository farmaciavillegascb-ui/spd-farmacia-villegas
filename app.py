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
    
    /* Animación de parpadeo rojo (alerta) */
    @keyframes pulse-subtle { 
        0% { transform: scale(1); box-shadow: 0 0 0 0 rgba(239, 68, 68, 0.6); } 
        50% { transform: scale(1.03); box-shadow: 0 0 0 12px rgba(239, 68, 68, 0); } 
        100% { transform: scale(1); box-shadow: 0 0 0 0 rgba(239, 68, 68, 0); } 
    }
    
    /* Estilos fijos para Columna 1 (INICIO) y Columna 3 (BAJAS) */
    [data-testid="column"]:nth-child(1) div.stButton > button { border-color: #0ea5e9 !important; color: #0284c7 !important; background-color: #f0f9ff !important; }
    [data-testid="column"]:nth-child(3) div.stButton > button { border-color: #f59e0b !important; color: #d97706 !important; background-color: #fffbeb !important; }
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

def generar_albaran_devolucion_pdf(nombre_paciente, ref_paciente, lista_devolucion):
    pdf = FPDF(orientation='L') 
    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, limpiar_texto_pdf("FARMACIA VILLEGAS C.B. - ALBARAN DE DEVOLUCION"), ln=True, align='C')
    pdf.set_font("Arial", '', 11)
    pdf.cell(0, 6, limpiar_texto_pdf(f"Paciente: {nombre_paciente} (Ref: {ref_paciente})"), ln=True, align='L')
    pdf.ln(5)
    
    pdf.set_font("Arial", 'B', 9)
    col_widths = [55, 65, 20, 35, 25, 30, 25] 
    headers = ["Medicamento", "Descripcion", "CN", "Lote", "Caducidad", "Serie", "Restantes"]
    for i in range(len(headers)):
        pdf.cell(col_widths[i], 8, limpiar_texto_pdf(headers[i]), border=1, align='C')
    pdf.ln()
    
    pdf.set_font("Arial", '', 8)
    for row in lista_devolucion:
        pdf.cell(col_widths[0], 8, limpiar_texto_pdf(str(row.get('Medicamento', ''))[:30]), border=1)
        pdf.cell(col_widths[1], 8, limpiar_texto_pdf(str(row.get('Descripción', ''))[:40]), border=1)
        pdf.cell(col_widths[2], 8, limpiar_texto_pdf(str(row.get('CN', ''))[:10]), border=1, align='C')
        pdf.cell(col_widths[3], 8, limpiar_texto_pdf(str(row.get('Lote', ''))[:18]), border=1, align='C')
        pdf.cell(col_widths[4], 8, limpiar_texto_pdf(str(row.get('Caducidad', ''))[:10]), border=1, align='C')
        pdf.cell(col_widths[5], 8, limpiar_texto_pdf(str(row.get('Serie', ''))[:15]), border=1, align='C')
        pdf.cell(col_widths[6], 8, limpiar_texto_pdf(str(row.get('Pastillas restantes', '0'))), border=1, align='C')
        pdf.ln()
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
                    st.session_state["usuario_autenticado"] = usuario_input
                    st.session_state["rol_usuario"] = USUARIOS_VALIDOS[usuario_input]["rol"]
                    st.session_state["pagina"] = "inicio"
                    st.rerun()
                else: st.error("❌ Usuario o clave incorrectos.")
    st.stop()

# CABECERA: 7 COLUMNAS INCLUYENDO "INICIO" AL PRINCIPIO
st.markdown('<div class="dashboard-header">', unsafe_allow_html=True)
st.markdown('<div class="logo-container"><span style="font-size: 24px;">💊</span><span class="logo-title">SPD FARMACIA VILLEGAS</span></div>', unsafe_allow_html=True)

rol_actual = st.session_state["rol_usuario"]
st.markdown(f'<div class="status-bar"><span>Sistema activo <span style="color: #22c55e; font-size: 16px;">●</span></span><span>Usuario: <b>{st.session_state["usuario_autenticado"]}</b> ({rol_actual.upper()})</span></div>', unsafe_allow_html=True)

# Lógica de conteos para alertas (rojo parpadeante)
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

# Inyectar CSS dinámico en la Columna 4 (Pedidos) si hay alertas
if alert_ped:
    st.markdown("""
    <style>
        [data-testid="column"]:nth-child(4) div.stButton > button { background: linear-gradient(135deg, #ef4444, #dc2626) !important; color: white !important; border: 2px solid #fca5a5 !important; animation: pulse-subtle 1.8s infinite; }
    </style>
    """, unsafe_allow_html=True)

# Inyectar CSS dinámico en la Columna 5 (Incidencias) si hay alertas
if alert_inc:
    st.markdown("""
    <style>
        [data-testid="column"]:nth-child(5) div.stButton > button { background: linear-gradient(135deg, #ef4444, #dc2626) !important; color: white !important; border: 2px solid #fca5a5 !important; animation: pulse-subtle 1.8s infinite; }
    </style>
    """, unsafe_allow_html=True)

# Dibujar las 7 columnas
col_inicio, col_alta, col_baja, col_ped, col_inc, col_user, col_logout = st.columns(7, gap="small")

with col_inicio:
    if st.button("🏠 INICIO", key="btn_hdr_inicio", use_container_width=True):
        st.session_state["pagina"] = "inicio"
        st.rerun()

with col_alta:
    if st.button("ALTA", key="btn_hdr_alta", use_container_width=True):
        st.session_state["pagina"] = "alta_paciente"
        st.rerun()

with col_baja:
    if st.button("🚨 BAJAS", key="btn_hdr_bajas", use_container_width=True):
        st.session_state["pagina"] = "baja_paciente"
        st.rerun()

with col_ped:
    if st.button(txt_ped, key="btn_hdr_ped", use_container_width=True):
        st.session_state["pagina"] = "pedidos_definitivos_admin" if rol_actual == "admin" else "seleccion_productos_enfermera"
        st.rerun()

with col_inc:
    if st.button(txt_inc, key="btn_hdr_inc", use_container_width=True):
        st.session_state["pagina"] = "incidencias"
        st.rerun()

with col_user:
    if st.button("USUARIOS", key="btn_hdr_usu", use_container_width=True):
        st.session_state["pagina"] = "gestion_usuarios"
        st.rerun()

with col_logout:
    if st.button("SALIR", key="btn_hdr_out", use_container_width=True):
        st.session_state["usuario_autenticado"] = None
        st.session_state["pagina"] = "inicio"
        st.rerun()

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
            # BOTÓN PARA ELIMINAR EL ÚLTIMO MEDICAMENTO ESCANEADO
            if st.button("🗑️ Eliminar último escaneo", use_container_width=False):
                st.session_state["df_devolucion"] = st.session_state["df_devolucion"].iloc[:-1]
                st.rerun()
            
            st.markdown("##### 📋 Listado de Devolución (Edite la columna 'Pastillas restantes')")
            # num_rows="dynamic" permite borrar filas desde la tabla marcándolas y pulsando suprimir o la papelera
            st.session_state["df_devolucion"] = st.data_editor(
                st.session_state["df_devolucion"], 
                use_container_width=True, 
                hide_index=True,
                num_rows="dynamic",
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
        if st.button("⬅ Volver a Modalidades"):
            st.session_state["baja_paso"] = "elegir_modalidad"
            st.rerun()

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
        meds_editadas = st.data_editor(st.session_state["df_alta_cargado"], num_rows="dynamic", use_container_width=True)
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
                time.sleep(1); st.rerun()
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
        
        # Guardamos la edición del dataframe en session/shared context
        info["datos"] = st.data_editor(info["datos"], use_container_width=True, hide_index=True)
        shared_data["lista_pacientes"][pk]["datos"] = info["datos"]

        st.markdown("<br>", unsafe_allow_html=True)
        # NUEVOS BOTONES DE ACCIÓN PARA LA FICHA
        col_btn_ped, col_btn_inc = st.columns(2)
        
        with col_btn_ped:
            if st.button("📦 Enviar a Propuesta de Pedido", use_container_width=True):
                df_pedidos = info["datos"][info["datos"]["Pedido"] == True]
                if not df_pedidos.empty:
                    for _, row in df_pedidos.iterrows():
                        item = {
                            "ref": info["ref"],
                            "paciente": info["nombre"],
                            "medicamento": row.get("Medicamento", ""),
                            "cn": row.get("CN", ""),
                            "posologia": row.get("Posologia", ""),
                            "seleccion_enfermera": False
                        }
                        shared_data["solicitud_pedido"].append(item)
                    # Desmarcamos las casillas una vez enviado
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
                    for _, row in df_incidencias.iterrows():
                        incidencia = {
                            "fecha": datetime.now().strftime("%d/%m/%Y %H:%M"),
                            "paciente": info["nombre"],
                            "medicamento": row.get("Medicamento", ""),
                            "estado": "Pendiente",
                            "descripcion": "Generada automáticamente desde ficha"
                        }
                        shared_data["incidencias_activas"].append(incidencia)
                    # Desmarcamos las casillas una vez enviado
                    info["datos"]["Incidencia"] = False
                    shared_data["lista_pacientes"][pk]["datos"] = info["datos"]
                    st.success("¡Incidencia registrada correctamente!")
                    time.sleep(1.5); st.rerun()
                else:
                    st.warning("Marca la casilla 'Incidencia' en algún medicamento primero.")

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("⬅ Volver a Lista"): st.session_state["pagina"] = "lista_pacientes"; st.rerun()

elif st.session_state["pagina"] == "seleccion_productos_enfermera":
    st.markdown("<h2 style='text-align: center;'>📦 PROPUESTA DE PEDIDO</h2>", unsafe_allow_html=True)
    if not shared_data["solicitud_pedido"]: st.info("No hay propuestas.")
    else:
        df_sol = pd.DataFrame(shared_data["solicitud_pedido"])
        shared_data["solicitud_pedido"] = st.data_editor(df_sol, use_container_width=True, hide_index=True, num_rows="dynamic").to_dict(orient="records")
        if st.button("🚀 Solicitar Pedido Definitivo"):
            sel = [i for i in shared_data["solicitud_pedido"] if i.get("seleccion_enfermera")]
            for it in sel: shared_data["pedidos_definitivos"].append(it)
            shared_data["solicitud_pedido"] = [i for i in shared_data["solicitud_pedido"] if not i.get("seleccion_enfermera")]
            st.success("Enviado al farmacéutico."); time.sleep(1); st.rerun()
    if st.button("⬅ Volver"): st.session_state["pagina"] = "inicio"; st.rerun()

elif st.session_state["pagina"] == "solicitud_pedido_admin":
    st.markdown("<h2 style='text-align: center;'>📦 PROPUESTA (ENVIADO A ENFERMERÍA)</h2>", unsafe_allow_html=True)
    if shared_data["solicitud_pedido"]: st.dataframe(pd.DataFrame(shared_data["solicitud_pedido"]), use_container_width=True, hide_index=True)
    else: st.info("Vacío.")
    if st.button("⬅ Volver"): st.session_state["pagina"] = "inicio"; st.rerun()

elif st.session_state["pagina"] == "pedidos_definitivos_admin":
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 800;'>🛒 PEDIDOS DEFINITIVOS Y ESCÁNER</h2>", unsafe_allow_html=True)
    if not shared_data["pedidos_definitivos"]: st.info("No hay pedidos definitivos pendientes.")
    else:
        with st.form("form_pedidos_dm", clear_on_submit=True):
            cadena_dm_pedido = st.text_input("📥 Escanee el DataMatrix del medicamento:")
            btn_ped = st.form_submit_button("Añadir Escaneo")
            if btn_ped and cadena_dm_pedido:
                parsed_ped = traducir_datamatrix(cadena_dm_pedido, BD_MEDICAMENTOS)
                for item in shared_data["pedidos_definitivos"]:
                    if not item.get("datamatrix"):
                        item["datamatrix"] = cadena_dm_pedido
                        item["lote"] = parsed_ped['lote']
                        item["caducidad"] = parsed_ped['caducidad']
                        break
        
        st.markdown("*(Puedes borrar escaneos erróneos marcando la casilla de la izquierda en la tabla y pulsando el icono de papelera)*")
        df_defs = pd.DataFrame(shared_data["pedidos_definitivos"])
        shared_data["pedidos_definitivos"] = st.data_editor(df_defs, use_container_width=True, hide_index=True, num_rows="dynamic").to_dict(orient="records")
        
        col_pdf, col_act = st.columns(2)
        with col_pdf:
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
            for row in shared_data["pedidos_definitivos"]:
                pdf.cell(col_widths[0], 8, limpiar_texto_pdf(str(row.get('ref', ''))[:8]), border=1)
                pdf.cell(col_widths[1], 8, limpiar_texto_pdf(str(row.get('paciente', ''))[:35]), border=1)
                pdf.cell(col_widths[2], 8, limpiar_texto_pdf(str(row.get('medicamento', ''))[:45]), border=1)
                pdf.cell(col_widths[3], 8, limpiar_texto_pdf(str(row.get('cn', ''))[:10]), border=1, align='C')
                pdf.cell(col_widths[4], 8, limpiar_texto_pdf(str(row.get('posologia', ''))[:12]), border=1, align='C')
                pdf.cell(col_widths[5], 8, limpiar_texto_pdf(str(row.get('datamatrix', ''))[:35]), border=1)
                pdf.cell(col_widths[6], 8, limpiar_texto_pdf(str(row.get('lote', ''))[:15]), border=1, align='C')
                pdf.cell(col_widths[7], 8, limpiar_texto_pdf(str(row.get('caducidad', ''))[:12]), border=1, align='C')
                pdf.ln()
            pdf_bytes = pdf.output(dest='S').encode('latin1')
            st.download_button("📄 Imprimir Albarán de Entrega (PDF)", data=pdf_bytes, file_name="Albaran_Entrega.pdf", mime="application/pdf", use_container_width=True)
        with col_act:
            if st.button("📌 Actualizar Última Entrega y Limpiar", use_container_width=True):
                shared_data["pedidos_definitivos"] = []
                st.success("¡Fechas actualizadas!"); time.sleep(1); st.rerun()
    if st.button("⬅ Volver al Menú"): st.session_state["pagina"] = "inicio"; st.rerun()

elif st.session_state["pagina"] == "incidencias":
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 800;'>⚠️ PANEL DE INCIDENCIAS</h2>", unsafe_allow_html=True)
    if shared_data["incidencias_activas"]:
        df_inc = pd.DataFrame(shared_data["incidencias_activas"])
        shared_data["incidencias_activas"] = st.data_editor(df_inc, use_container_width=True, hide_index=True, num_rows="dynamic").to_dict(orient="records")
    else: st.info("No hay incidencias.")
    if st.button("⬅ Volver"): st.session_state["pagina"] = "inicio"; st.rerun()
    
elif st.session_state["pagina"] == "gestion_usuarios":
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 800;'>🔐 GESTIÓN DE USUARIOS Y ROLES</h2>", unsafe_allow_html=True)
    if st.button("⬅ Volver"): st.session_state["pagina"] = "inicio"; st.rerun()
