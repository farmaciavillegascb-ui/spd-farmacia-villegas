import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime
from fpdf import FPDF
import requests

# --- PEGA AQUÍ TU URL DE GOOGLE APPS SCRIPT ---
WEB_APP_URL = "TU_URL_DE_GOOGLE_APPS_SCRIPT_AQUÍ"

st.set_page_config(
    page_title="SPD FARMACIA VILLEGAS", 
    page_icon="💊", 
    layout="centered",
    initial_sidebar_state="collapsed"
)

# --- ESTILOS VISUALES ---
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .card-incidencia { background-color: #fff5f5; border-left: 5px solid #dc3545; padding: 15px; border-radius: 8px; margin-bottom: 10px; }
    .avisos-box { background-color: #ffffff; border: 2px solid #e2e8f0; padding: 20px; border-radius: 12px; margin-bottom: 20px; }
    div.stButton > button:first-child {
        background: linear-gradient(135deg, #007bff 0%, #0056b3 100%);
        color: white !important; border-radius: 12px !important; padding: 12px !important; font-weight: 800 !important; width: 100% !important; border: none !important;
    }
    div.stFormSubmitButton > button:first-child {
        background: linear-gradient(135deg, #28a745 0%, #1e7e34 100%);
        color: white !important; border-radius: 12px !important; padding: 12px !important; font-weight: 800 !important; width: 100% !important; border: none !important;
    }
    </style>
""", unsafe_allow_html=True)

ARCHIVO_USUARIOS = 'usuarios.json'
ARCHIVO_INCIDENCIAS = 'incidencias.json'

def cargar_json(archivo, def_val):
    if not os.path.exists(archivo):
        with open(archivo, 'w') as f: json.dump(def_val, f)
        return def_val
    with open(archivo, 'r') as f: return json.load(f)

def guardar_json(archivo, datos):
    with open(archivo, 'w') as f: json.dump(datos, f)

def obtener_solicitudes():
    try:
        res = requests.get(f"{WEB_APP_URL}?action=leer", timeout=10)
        if res.status_code == 200: return res.json()
    except: pass
    return []

def enviar_servidor(payload):
    try:
        requests.post(WEB_APP_URL, json=payload, timeout=10)
    except: pass

usuarios_base = {
    "farmacia": {"password": "admin", "rol": "administrador"},
    "enfermera 1": {"password": "1234", "rol": "usuario"},
    "enfermera 2": {"password": "1234", "rol": "usuario"}
}
usuarios_db = cargar_json(ARCHIVO_USUARIOS, usuarios_base)
incidencias_db = cargar_json(ARCHIVO_INCIDENCIAS, {})

if 'usuario' not in st.session_state: st.session_state['usuario'] = None
if 'rol' not in st.session_state: st.session_state['rol'] = None
if 'pantalla' not in st.session_state: st.session_state['pantalla'] = 'menu'
if 'paciente_actual' not in st.session_state: st.session_state['paciente_actual'] = None

# --- LOGIN ---
if st.session_state['usuario'] is None:
    c1, c2, c3 = st.columns([0.1, 2, 0.1])
    with c2:
        st.markdown("<br><h2 style='text-align: center; color: #0066cc;'>💊 SPD FARMACIA VILLEGAS</h2>", unsafe_allow_html=True)
        with st.form("login"):
            usr = st.text_input("👤 USUARIO")
            pwd = st.text_input("🔑 CONTRASEÑA", type="password")
            if st.form_submit_button("🚀 ENTRAR"):
                if usr in usuarios_db and usuarios_db[usr]["password"] == pwd:
                    st.session_state['usuario'] = usr
                    st.session_state['rol'] = usuarios_db[usr]["rol"]
                    st.session_state['pantalla'] = 'menu'
                    st.rerun()
                else:
                    st.error("❌ Datos incorrectos.")
else:
    es_admin = (st.session_state['rol'] == 'administrador')
    solicitudes_db = obtener_solicitudes()
    
    with st.sidebar:
        st.info(f"👤 **{st.session_state['usuario'].upper()}**")
        if st.button("🏠 Menú Principal", use_container_width=True): st.session_state['pantalla'] = 'menu'; st.rerun()
        if st.button("📦 Pedidos / Albaranes", use_container_width=True): st.session_state['pantalla'] = 'albaranes' if es_admin else 'pedido_enfermeria'; st.rerun()
        if st.button("🚪 Cerrar Sesión", use_container_width=True): st.session_state['usuario'] = None; st.rerun()

    @st.cache_data
    def cargar_excel():
        excel = next((f for f in os.listdir('.') if f.endswith('.xlsx') and not f.startswith('~$')), None)
        if not excel: return None
        xls = pd.ExcelFile(excel)
        df = pd.concat([xls.parse(s) for s in xls.sheet_names], ignore_index=True)
        df.columns = df.columns.str.strip().str.upper()
        return df

    datos = cargar_excel()
    if datos is None:
        st.error("⚠️ Sube el archivo Excel de tratamientos (.xlsx) a GitHub.")
        st.stop()

    col_nombre = next((c for c in datos.columns if 'NOMBRE' in str(c).upper()), None)

    # --- PANTALLA: MENÚ PRINCIPAL ---
    if st.session_state['pantalla'] == 'menu':
        st.markdown("<h1 style='text-align:center;'>🏥 PANEL DE CONTROL</h1>", unsafe_allow_html=True)
        
        st.markdown("<div class='avisos-container'>", unsafe_allow_html=True)
        st.markdown("<h3>🔔 TAREAS Y AVISOS</h3>", unsafe_allow_html=True)
        
        if es_admin:
            pendientes = [s for s in solicitudes_db if str(s.get('estado')).strip().lower() == 'pendiente']
            if not pendientes:
                st.success("✨ Todo al día. No hay solicitudes pendientes.")
            else:
                st.info(f"Tienes **{len(pendientes)}** solicitud(es) pendientes:")
                for sol in pendientes:
                    s_id = sol.get('id')
                    tipo = sol.get('tipo')
                    pac = sol.get('paciente')
                    med = sol.get('medicamento', '')
                    notas = sol.get('notas', '')
                    
                    with st.expander(f"📌 [{tipo}] - {pac} {('(' + med + ')') if med else ''}"):
                        st.write(f"**Detalles:** {notas}")
                        c_a, c_b = st.columns(2)
                        with c_a:
                            if st.button("✅ APROBAR / ENVIAR A ALBARÁN", key=f"apr_{s_id}", use_container_width=True):
                                enviar_servidor({"action": "actualizar", "id": s_id, "estado": "aprobada"})
                                st.success("¡Aprobado!")
                                st.rerun()
                        with c_b:
                            if st.button("❌ RECHAZAR", key=f"rec_{s_id}", use_container_width=True):
                                enviar_servidor({"action": "actualizar", "id": s_id, "estado": "rechazada"})
                                st.error("Rechazado.")
                                st.rerun()
        else:
            mis_sols = [s for s in solicitudes_db if str(s.get('solicitante')).lower() == str(st.session_state['usuario']).lower()]
            if not mis_sols:
                st.info("ℹ️ No has enviado solicitudes recientes.")
            else:
                for s in mis_sols:
                    st.write(f"• **[{s.get('tipo')}]** Paciente: {s.get('paciente')} - Estado: **{s.get('estado').upper()}**")
        st.markdown("</div>", unsafe_allow_html=True)

        c_1, c_2 = st.columns(2)
        with c_1:
            if st.button("👥 Directorio Pacientes", use_container_width=True): st.session_state['pantalla'] = 'pacientes'; st.rerun()
            if st.button("➕ Solicitud Alta", use_container_width=True): st.session_state['pantalla'] = 'alta'; st.rerun()
        with c_2:
            if st.button("📦 Pedido Fuera de Blíster", use_container_width=True): st.session_state['pantalla'] = 'pedido_enfermeria'; st.rerun()
            if st.button("➖ Solicitud Baja", use_container_width=True): st.session_state['pantalla'] = 'baja'; st.rerun()

    # --- PANTALLA: DIRECTORIO PACIENTES ---
    elif st.session_state['pantalla'] == 'pacientes':
        if st.button("⬅️ Volver", use_container_width=True): st.session_state['pantalla'] = 'menu'; st.rerun()
        st.markdown("<h2>👥 Directorio de Pacientes</h2>", unsafe_allow_html=True)
        busq = st.text_input("🔍 Buscar paciente:")
        df_p = datos[[col_nombre]].drop_duplicates().dropna()
        if busq: df_p = df_p[df_p[col_nombre].str.lower().str.contains(busq.lower())]
        
        for idx, row in df_p.iterrows():
            p_nom = str(row[col_nombre])
            if st.button(f"👤 {p_nom}", key=f"pac_{idx}", use_container_width=True):
                st.session_state['paciente_actual'] = p_nom
                st.session_state['pantalla'] = 'ficha'
                st.rerun()

    # --- PANTALLA: FICHA PACIENTE (MARCAR FUERA DE BLÍSTER) ---
    elif st.session_state['pantalla'] == 'ficha':
        pac = st.session_state['paciente_actual']
        if st.button("⬅️ Volver al Directorio", use_container_width=True): st.session_state['pantalla'] = 'pacientes'; st.rerun()
        st.markdown(f"<h2>📋 {pac}</h2>", unsafe_allow_html=True)
        
        meds = datos[datos[col_nombre].astype(str) == pac]
        st.markdown("### Selecciona medicación para **Fuera de Blíster**:")
        
        with st.form("form_ficha_blister"):
            checks = []
            for i, r in meds.iterrows():
                m_nom = ""
                for col in meds.columns:
                    if 'MEDICAMENTO' in col or 'MEDICINA' in col:
                        m_nom = str(r[col])
                        break
                cn_val = ""
                for col in meds.columns:
                    if 'CN' in col or 'CÓDIGO' in col:
                        cn_val = str(r[col])
                        break
                
                chk = st.checkbox(f"💊 {m_nom} (CN: {cn_val})", key=f"chk_med_{i}")
                if chk: checks.append({"med": m_nom, "cn": cn_val})
                st.markdown("---")
            
            if st.form_submit_button("🚀 ENVIAR SELECCIÓN A FARMACIA (ALBARÁN)"):
                if checks:
                    for item in checks:
                        enviar_servidor({
                            "action": "insertar",
                            "tipo": "FUERA DE BLISTER",
                            "paciente": pac,
                            "medicamento": item['med'],
                            "cn": item['cn'],
                            "solicitante": st.session_state['usuario'],
                            "notas": f"Fuera de blíster: {item['med']}",
                            "estado": "pendiente"
                        })
                    st.success("✅ ¡Enviado a la farmacia correctamente para su albarán!")
                else:
                    st.warning("⚠️ Selecciona al menos un medicamento.")

    # --- PANTALLA: GESTIÓN DE ALBARANES Y DATAMATRIX (ADMIN) ---
    elif st.session_state['pantalla'] == 'albaranes':
        if not es_admin: st.session_state['pantalla'] = 'menu'; st.rerun()
        if st.button("⬅️ Volver", use_container_width=True): st.session_state['pantalla'] = 'menu'; st.rerun()
        
        st.markdown("<h2>🖨️ GESTIÓN DE ALBARANES Y DATAMATRIX</h2>", unsafe_allow_html=True)
        pendientes_blister = [s for s in solicitudes_db if 'FUERA DE BLISTER' in str(s.get('tipo')).upper() and str(s.get('estado')).lower() == 'pendiente']
        
        if not pendientes_blister:
            st.info("✨ No hay medicamentos pendientes de fuera de blíster. Selecciónalos desde el Directorio de Pacientes.")
        else:
            with st.form("form_albaran_dm"):
                ingresados = {}
                for s in pendientes_blister:
                    s_id = s.get('id')
                    st.markdown(f"👤 **Paciente:** {s.get('paciente')} | 💊 **Fármaco:** {s.get('medicamento')} | 🏷️ **CN:** {s.get('cn')}")
                    lote = st.text_input("LOTE", key=f"l_{s_id}")
                    cad = st.text_input("CADUCIDAD (MM/AAAA)", key=f"c_{s_id}")
                    ingresados[s_id] = {"paciente": s.get('paciente'), "med": s.get('medicamento'), "cn": s.get('cn'), "lote": lote, "cad": cad}
                    st.markdown("---")
                
                if st.form_submit_button("🖨️ GENERAR Y DESCARGAR PDF DE ALBARÁN"):
                    fecha_gen = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                    pdf = FPDF()
                    pdf.add_page()
                    pdf.set_font("Arial", "B", 16)
                    pdf.cell(200, 10, txt="ALBARÁN DE ENTREGA - FARMACIA VILLEGAS", ln=True, align="C")
                    pdf.set_font("Arial", "", 11)
                    pdf.cell(200, 10, txt=f"Fecha: {fecha_gen}", ln=True, align="C")
                    pdf.ln(10)
                    
                    pdf.set_font("Arial", "B", 10)
                    pdf.cell(45, 10, "PACIENTE", 1)
                    pdf.cell(75, 10, "MEDICAMENTO", 1)
                    pdf.cell(20, 10, "CN", 1)
                    pdf.cell(50, 10, "LOTE / CAD", 1)
                    pdf.ln()
                    
                    pdf.set_font("Arial", "", 9)
                    for s_id, d in ingresados.items():
                        enviar_servidor({"action": "actualizar", "id": s_id, "estado": "aprobada", "fecha_albaran": fecha_gen})
                        pdf.cell(45, 10, str(d['paciente'][:22]), 1)
                        pdf.cell(75, 10, str(d['med'][:40]), 1)
                        pdf.cell(20, 10, str(d['cn'] or '000000'), 1)
                        pdf.cell(50, 10, f"{d['lote'] or 'L-01'} / {d['cad'] or '12/2028'}", 1)
                        pdf.ln()
                        
                    pdf_bytes = pdf.output(dest='S').encode('latin1')
                    st.success("✅ ¡Albarán generado con éxito!")
                    st.download_button("📥 DESCARGAR PDF", data=pdf_bytes, file_name=f"albaran_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf", mime="application/pdf")

    # --- PANTALLA: PEDIDO ENFERMERÍA ---
    elif st.session_state['pantalla'] == 'pedido_enfermeria':
        if st.button("⬅️ Volver", use_container_width=True): st.session_state['pantalla'] = 'menu'; st.rerun()
        st.markdown("<h2>📦 Solicitud de Medicación / Fuera de Blíster</h2>", unsafe_allow_html=True)
        st.info("💡 Para solicitar medicamentos fuera de blíster de forma rápida, entra en el **Directorio de Pacientes**, selecciona al paciente y marca sus medicamentos directamente.")

    # --- PANTALLA: SOLICITUD ALTA ---
    elif st.session_state['pantalla'] == 'alta':
        if st.button("⬅️ Volver", use_container_width=True): st.session_state['pantalla'] = 'menu'; st.rerun()
        st.markdown("<h2>➕ Solicitud de Alta de Paciente</h2>", unsafe_allow_html=True)
        with st.form("f_alta"):
            nom = st.text_input("👤 Nombre del Paciente:")
            obs = st.text_area("📝 Observaciones:")
            if st.form_submit_button("Enviar Alta"):
                if nom:
                    enviar_servidor({"action": "insertar", "tipo": "ALTA PACIENTE", "paciente": nom, "solicitante": st.session_state['usuario'], "notas": obs, "estado": "pendiente"})
                    st.success("✅ Solicitud de alta enviada.")
                else:
                    st.warning("Escribe un nombre.")

    # --- PANTALLA: SOLICITUD BAJA ---
    elif st.session_state['pantalla'] == 'baja':
        if st.button("⬅️ Volver", use_container_width=True): st.session_state['pantalla'] = 'menu'; st.rerun()
        st.markdown("<h2>➖ Solicitud de Baja de Paciente</h2>", unsafe_allow_html=True)
        lista_p = sorted(datos[col_nombre].dropna().astype(str).unique())
        with st.form("f_baja"):
            p_sel = st.selectbox("👤 Selecciona Paciente:", lista_p)
            mot = st.text_area("📝 Motivo de la baja:")
            if st.form_submit_button("Enviar Baja"):
                if mot:
                    enviar_servidor({"action": "insertar", "tipo": "BAJA PACIENTE", "paciente": p_sel, "solicitante": st.session_state['usuario'], "notas": mot, "estado": "pendiente"})
                    st.success("✅ Solicitud de baja enviada.")
                else:
                    st.warning("Indica el motivo.")
