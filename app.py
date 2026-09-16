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
    .avisos-box { background-color: #ffffff; border: 2px solid #e2e8f0; padding: 20px; border-radius: 12px; margin-bottom: 20px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); }
    div.stButton > button:first-child {
        background: linear-gradient(135deg, #007bff 0%, #0056b3 100%);
        color: white !important; border-radius: 14px !important; padding: 15px !important; font-weight: 900 !important; font-size: 18px !important; width: 100% !important; border: none !important; text-transform: uppercase !important;
    }
    div.stFormSubmitButton > button:first-child {
        background: linear-gradient(135deg, #28a745 0%, #1e7e34 100%);
        color: white !important; border-radius: 12px !important; padding: 15px !important; font-weight: 800 !important; font-size: 18px !important; width: 100% !important; border: none !important; text-transform: uppercase !important;
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
        st.markdown("<br><h2 style='text-align: center; color: #0066cc;'>💊 SPD FARMACIA VILLEGAS</h2><p style='text-align: center; color: #666;'>Sincronización Cloud</p>", unsafe_allow_html=True)
        with st.form("login"):
            usr = st.text_input("👤 USUARIO")
            pwd = st.text_input("🔑 CONTRASEÑA", type="password")
            if st.form_submit_button("🚀 INICIAR SESIÓN"):
                if usr in usuarios_db and usuarios_db[usr]["password"] == pwd:
                    st.session_state['usuario'] = usr
                    st.session_state['rol'] = usuarios_db[usr]["rol"]
                    st.session_state['pantalla'] = 'menu'
                    st.rerun()
                else:
                    st.error("❌ Usuario o contraseña incorrectos.")
else:
    es_admin = (st.session_state['rol'] == 'administrador')
    solicitudes_db = obtener_solicitudes()
    
    with st.sidebar:
        st.info(f"👩‍⚕️ **{st.session_state['usuario'].upper()}**")
        if st.button("🏠 MENÚ PRINCIPAL", use_container_width=True): st.session_state['pantalla'] = 'menu'; st.rerun()
        if st.button("🚨 VER INCIDENCIAS", use_container_width=True): st.session_state['pantalla'] = 'panel_incidencias'; st.rerun()
        if st.button("📦 PEDIDO FUERA DE BLÍSTER", use_container_width=True): st.session_state['pantalla'] = 'panel_presolicitudes'; st.rerun()
        if es_admin:
            ped_b = [s for s in solicitudes_db if 'FUERA DE BLISTER' in str(s.get('tipo')).upper() and str(s.get('estado')).lower() == 'pendiente']
            if ped_b:
                if st.button(f"🖨️ ALBARÁN ({len(ped_b)})", use_container_width=True): st.session_state['pantalla'] = 'admin_gestion_datamatrix'; st.rerun()
        if st.button("🚪 CERRAR SESIÓN", use_container_width=True): st.session_state['usuario'] = None; st.rerun()

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
        st.markdown("<h1 style='text-align:center;'>🏥 SPD FARMACIA</h1><br>", unsafe_allow_html=True)
        
        # CENTRO DE AVISOS
        st.markdown("<div class='avisos-box'>", unsafe_allow_html=True)
        st.markdown("<h3 style='margin-top:0; color:#1a334e;'>🔔 CENTRO DE AVISOS Y TAREAS</h3>", unsafe_allow_html=True)
        
        if es_admin:
            sols_pen = [s for s in solicitudes_db if str(s.get('estado')).lower() == 'pendiente' and 'FUERA DE BLISTER' not in str(s.get('tipo')).upper()]
            peds_pen = [s for s in solicitudes_db if 'FUERA DE BLISTER' in str(s.get('tipo')).upper() and str(s.get('estado')).lower() == 'pendiente']
            total_t = len(sols_pen) + len(peds_pen) + len(incidencias_db)
            
            if total_t == 0:
                st.success("✨ ¡Todo al día! No hay tareas pendientes.")
            else:
                st.info(f"Tienes **{total_t}** asunto(s) pendientes de administración:")
                if sols_pen:
                    st.markdown("#### 📋 Altas / Bajas / Nuevos Fármacos:")
                    for sol in sols_pen:
                        s_id = sol.get('id')
                        with st.expander(f"📌 [{sol.get('tipo')}] - {sol.get('paciente')} (Solicitante: {sol.get('solicitante', 'N/A')})"):
                            st.write(f"**Detalles:** {sol.get('notas')}")
                            c1, c2 = st.columns(2)
                            with c1:
                                if st.button("✅ APROBAR", key=f"apr_{s_id}", use_container_width=True):
                                    enviar_servidor({"action": "actualizar", "id": s_id, "estado": "aprobada"})
                                    st.success("¡Aprobado!")
                                    st.rerun()
                            with c2:
                                if st.button("❌ RECHAZAR", key=f"rec_{s_id}", use_container_width=True):
                                    enviar_servidor({"action": "actualizar", "id": s_id, "estado": "rechazada"})
                                    st.error("Rechazado.")
                                    st.rerun()
                if peds_pen:
                    st.markdown("#### 📦 Pedidos Fuera de Blíster:")
                    st.warning(f"Hay **{len(peds_pen)}** medicamento(s) listos para Albarán de Entrega.")
                    if st.button("🖨️ IR A GENERAR ALBARÁN", key="btn_ir_alb", use_container_width=True):
                        st.session_state['pantalla'] = 'admin_gestion_datamatrix'
                        st.rerun()
                if incidencias_db:
                    st.markdown("#### 🚨 Incidencias Activas:")
                    incs_del = []
                    for cl, det in incidencias_db.items():
                        p = cl.split("____")[0] if "____" in cl else "Desc."
                        m = cl.split("____")[1] if "____" in cl else cl
                        st.write(f"• **{p}** - {m}")
                        if st.button(f"Resolver incidencia de {p}", key=f"res_{cl}", use_container_width=True):
                            incs_del.append(cl)
                    if incs_del:
                        for cl in incs_del: del incidencias_db[cl]
                        guardar_json(ARCHIVO_INCIDENCIAS, incidencias_db)
                        st.success("✅ Incidencia resuelta.")
                        st.rerun()
        else:
            mis_s = [s for s in solicitudes_db if str(s.get('solicitante')).lower() == str(st.session_state['usuario']).lower()]
            if not mis_s:
                st.info("ℹ️ No has enviado solicitudes recientes.")
            else:
                for s in mis_s:
                    st.write(f"• **[{s.get('tipo')}]** Paciente: {s.get('paciente')} - Estado: **{s.get('estado').upper()}**")
        st.markdown("</div>", unsafe_allow_html=True)

        # BOTONES DEL MENÚ PRINCIPAL
        g1, g2 = st.columns(2)
        with g1:
            if st.button("👥 DIRECTORIO PACIENTES", use_container_width=True): st.session_state['pantalla'] = 'pacientes'; st.rerun()
            if st.button("➕ SOLICITUD ALTAS", use_container_width=True): st.session_state['pantalla'] = 'alta'; st.rerun()
        with g2:
            if st.button("🚨 PANEL INCIDENCIAS", use_container_width=True): st.session_state['pantalla'] = 'panel_incidencias'; st.rerun()
            if st.button("➖ SOLICITUD BAJAS", use_container_width=True): st.session_state['pantalla'] = 'baja'; st.rerun()

        g3, g4 = st.columns(2)
        with g3:
            if st.button("📦 PEDIDO FUERA DE BLÍSTER", use_container_width=True): st.session_state['pantalla'] = 'panel_presolicitudes'; st.rerun()
        with g4:
            if es_admin:
                if st.button("🖨️ ALBARÁN DE ENTREGA", use_container_width=True): st.session_state['pantalla'] = 'admin_gestion_datamatrix'; st.rerun()

    # --- DIRECTORIO DE PACIENTES ---
    elif st.session_state['pantalla'] == 'pacientes':
        if st.button("⬅️ VOLVER AL MENÚ", use_container_width=True): st.session_state['pantalla'] = 'menu'; st.rerun()
        st.markdown("<h2>👥 DIRECTORIO DE PACIENTES</h2>", unsafe_allow_html=True)
        busq = st.text_input("🔍 Buscar paciente:")
        df_p = datos[[col_nombre]].drop_duplicates().dropna()
        if busq: df_p = df_p[df_p[col_nombre].str.lower().str.contains(busq.lower())]
        
        for idx, row in df_p.iterrows():
            p_nom = str(row[col_nombre])
            if st.button(f"👤 {p_nom}", key=f"pac_{idx}", use_container_width=True):
                st.session_state['paciente_actual'] = p_nom
                st.session_state['pantalla'] = 'ficha_paciente'
                st.rerun()

    # --- FICHA PACIENTE ---
    elif st.session_state['pantalla'] == 'ficha_paciente':
        pac = st.session_state['paciente_actual']
        if st.button("⬅️ VOLVER AL DIRECTORIO", use_container_width=True): st.session_state['pantalla'] = 'pacientes'; st.rerun()
        st.markdown(f"<h2>📋 {pac}</h2>", unsafe_allow_html=True)
        
        t1, t2 = st.tabs(["💊 TRATAMIENTOS", "➕ NUEVO FÁRMACO"])
        with t1:
            meds = datos[datos[col_nombre].astype(str) == pac]
            with st.form("form_t"):
                blisters = []
                for idx, fila in meds.iterrows():
                    med = next((str(fila[c]) for c in meds.columns if 'MEDICAMENTO' in c or 'MEDICINA' in c), "N/A")
                    cn = next((str(fila[c]) for c in meds.columns if 'CN' in c or 'CÓDIGO' in c), "N/A")
                    st.markdown(f"**{med}** (CN: {cn})")
                    if st.checkbox("📦 FUERA DE BLÍSTER", key=f"b_{idx}"):
                        blisters.append({"paciente": pac, "medicamento": med, "cn": cn})
                    st.markdown("---")
                
                if st.form_submit_button("📨 AÑADIR A PEDIDO FUERA DE BLÍSTER"):
                    if blisters:
                        presols = cargar_json('presolicitudes_blister.json', [])
                        for item in blisters:
                            if not any(p['paciente'] == item['paciente'] and p['medicamento'] == item['medicamento'] for p in presols):
                                presols.append(item)
                        guardar_json('presolicitudes_blister.json', presols)
                        st.success("✅ ¡Añadido! Ve a 'Pedido Fuera de Blíster' para enviarlo a farmacia.")
        with t2:
            with st.form("form_n"):
                n_med = st.text_input("💊 MEDICAMENTO:")
                n_pau = st.text_input("🕒 PAUTA:")
                if st.form_submit_button("ENVIAR SOLICITUD"):
                    if n_med:
                        enviar_servidor({
                            "action": "insertar", "tipo": "NUEVO MEDICAMENTO", "paciente": pac,
                            "solicitante": st.session_state['usuario'], "notas": f"{n_med} - {n_pau}", "estado": "pendiente"
                        })
                        st.success("✅ Solicitud enviada.")

    # --- PANEL PRE-SOLICITUDES ---
    elif st.session_state['pantalla'] == 'panel_presolicitudes':
        if st.button("⬅️ VOLVER AL MENÚ", use_container_width=True): st.session_state['pantalla'] = 'menu'; st.rerun()
        st.markdown("<h2>📦 PEDIDO FUERA DE BLÍSTER</h2>", unsafe_allow_html=True)
        presols = cargar_json('presolicitudes_blister.json', [])
        
        if not presols:
            st.info("✨ No hay medicamentos señalados. Entra en el **Directorio de Pacientes** y marca la casilla correspondiente.")
        else:
            with st.form("form_presol"):
                chks = []
                for idx, item in enumerate(presols):
                    m = st.checkbox(f"👤 {item.get('paciente')} - 💊 {item.get('medicamento')} (CN: {item.get('cn')})", key=f"chk_p_{idx}")
                    if m: chks.append(idx)
                
                if st.form_submit_button("🚀 ENVIAR A FARMACIA (APARTADO ALBARÁN)"):
                    if chks:
                        restantes = []
                        for idx, item in enumerate(presols):
                            if idx in chks:
                                enviar_servidor({
                                    "action": "insertar", "tipo": "FUERA DE BLISTER",
                                    "paciente": item.get('paciente'), "medicamento": item.get('medicamento'),
                                    "cn": item.get('cn'), "solicitante": st.session_state['usuario'],
                                    "notas": f"Fuera de blíster: {item.get('medicamento')}", "estado": "pendiente"
                                })
                            else:
                                restantes.append(item)
                        guardar_json('presolicitudes_blister.json', restantes)
                        st.success("✅ ¡Enviado correctamente a farmacia!")
                        st.rerun()

    # --- GESTIÓN DE ALBARANES Y DATAMATRIX (ADMIN) ---
    elif st.session_state['pantalla'] == 'admin_gestion_datamatrix':
        if not es_admin: st.session_state['pantalla'] = 'menu'; st.rerun()
        if st.button("⬅️ VOLVER AL MENÚ", use_container_width=True): st.session_state['pantalla'] = 'menu'; st.rerun()
        
        st.markdown("<h2>🖨️ ALBARÁN DE ENTREGA Y DATAMATRIX</h2>", unsafe_allow_html=True)
        ped_b = [s for s in solicitudes_db if 'FUERA DE BLISTER' in str(s.get('tipo')).upper() and str(s.get('estado')).lower() == 'pendiente']
        
        if not ped_b:
            st.info("✨ No hay pedidos pendientes de fuera de blíster.")
        else:
            with st.form("form_alb"):
                ingresados = {}
                for s in ped_b:
                    s_id = s.get('id')
                    st.markdown(f"👤 **Paciente:** {s.get('paciente')} | 💊 **Fármaco:** {s.get('medicamento')} | 🏷️ **CN:** {s.get('cn')}")
                    lote = st.text_input("LOTE", key=f"l_{s_id}")
                    cad = st.text_input("CADUCIDAD (MM/AAAA)", key=f"c_{s_id}")
                    ingresados[s_id] = {"paciente": s.get('paciente'), "med": s.get('medicamento'), "cn": s.get('cn'), "lote": lote, "cad": cad}
                    st.markdown("---")
                
                if st.form_submit_button("🖨️ GENERAR Y DESCARGAR PDF"):
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

    # --- SOLICITUD ALTA ---
    elif st.session_state['pantalla'] == 'alta':
        if st.button("⬅️ VOLVER AL MENÚ", use_container_width=True): st.session_state['pantalla'] = 'menu'; st.rerun()
        st.markdown("<h2>➕ SOLICITUD DE ALTA DE PACIENTE</h2>", unsafe_allow_html=True)
        with st.form("f_alta"):
            nom = st.text_input("👤 Nombre del Paciente:")
            obs = st.text_area("📝 Observaciones / Tratamiento:")
            if st.form_submit_button("ENVIAR SOLICITUD DE ALTA"):
                if nom:
                    enviar_servidor({"action": "insertar", "tipo": "ALTA PACIENTE", "paciente": nom, "solicitante": st.session_state['usuario'], "notas": obs, "estado": "pendiente"})
                    st.success("✅ Solicitud de alta enviada al administrador.")
                else:
                    st.warning("Escribe un nombre.")

    # --- SOLICITUD BAJA ---
    elif st.session_state['pantalla'] == 'baja':
        if st.button("⬅️ VOLVER AL MENÚ", use_container_width=True): st.session_state['pantalla'] = 'menu'; st.rerun()
        st.markdown("<h2>➖ SOLICITUD DE BAJA DE PACIENTE</h2>", unsafe_allow_html=True)
        lista_p = sorted(datos[col_nombre].dropna().astype(str).unique())
        with st.form("f_baja"):
            p_sel = st.selectbox("👤 Selecciona Paciente:", lista_p)
            mot = st.text_area("📝 Motivo de la baja:")
            if st.form_submit_button("ENVIAR SOLICITUD DE BAJA"):
                if mot:
                    enviar_servidor({"action": "insertar", "tipo": "BAJA PACIENTE", "paciente": p_sel, "solicitante": st.session_state['usuario'], "notas": mot, "estado": "pendiente"})
                    st.success("✅ Solicitud de baja enviada al administrador.")
                else:
                    st.warning("Indica el motivo.")

    # --- PANEL DE INCIDENCIAS ---
    elif st.session_state['pantalla'] == 'panel_incidencias':
        if st.button("⬅️ VOLVER AL MENÚ", use_container_width=True): st.session_state['pantalla'] = 'menu'; st.rerun()
        st.markdown("<h2>🚨 PANEL DE INCIDENCIAS</h2>", unsafe_allow_html=True)
        
        if not incidencias_db:
            st.info("✨ No hay incidencias activas.")
        else:
            del_list = []
            for cl, det in incidencias_db.items():
                p = cl.split("____")[0] if "____" in cl else "Desc."
                m = cl.split("____")[1] if "____" in cl else cl
                st.markdown(f"""
                    <div class='card-incidencia'>
                        <b>👤 Paciente:</b> {p}<br>
                        <b>💊 Fármaco:</b> {m}<br>
                        <b>Observaciones:</b> {det.get('observaciones', 'N/A')}
                    </div>
                """, unsafe_allow_html=True)
                if st.button(f"🗑️ Resolver e eliminar incidencia", key=f"res_inc_{cl}", use_container_width=True):
                    del_list.append(cl)
            if del_list:
                for cl in del_list: del incidencias_db[cl]
                guardar_json(ARCHIVO_INCIDENCIAS, incidencias_db)
                st.success("✅ Incidencia resuelta y eliminada.")
                st.rerun()
