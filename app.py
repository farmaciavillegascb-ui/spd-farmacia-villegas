import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime
from fpdf import FPDF
import requests

# --- URL DE TU WEB APP DE GOOGLE SHEETS ---
WEB_APP_URL = "TU_URL_DE_GOOGLE_APPS_SCRIPT_AQUÍ"

# Configuración inicial de la página optimizada para móviles
st.set_page_config(
    page_title="SPD FARMACIA VILLEGAS", 
    page_icon="💊", 
    layout="centered",
    initial_sidebar_state="collapsed"
)

# --- ESTILOS CSS ---
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .card-medicamento, .card-incidencia, .card-presolicitud {
        background-color: #ffffff; padding: 16px; border-radius: 12px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.06); margin-bottom: 14px; border-left: 5px solid #0066cc;
    }
    .card-incidencia { background-color: #fff5f5; border-left: 5px solid #dc3545; }
    
    .avisos-container {
        background-color: #ffffff;
        border: 2px solid #e2e8f0;
        padding: 20px;
        border-radius: 14px;
        margin-bottom: 25px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.05);
    }
    
    div.stButton > button:first-child, div.stButton > button:first-child p {
        color: white !important; font-weight: 900 !important; font-size: 24px !important; text-transform: uppercase !important;
    }
    div.stButton > button:first-child {
        background: linear-gradient(135deg, #007bff 0%, #0056b3 100%);
        border-radius: 16px !important; padding: 15px !important; border: none !important;
        height: 140px !important; display: flex !important; align-items: center !important; justify-content: center !important; text-align: center !important;
    }
    div.stFormSubmitButton > button:first-child {
        background: linear-gradient(135deg, #28a745 0%, #1e7e34 100%);
        color: white !important; border-radius: 12px !important; padding: 15px 20px !important;
        font-weight: 800 !important; font-size: 18px !important; text-transform: uppercase !important; border: none !important; width: 100% !important;
    }
    </style>
""", unsafe_allow_html=True)

# --- BASES DE DATOS LOCALES Y GOOGLE SHEETS ---
ARCHIVO_USUARIOS = 'usuarios.json'
ARCHIVO_INCIDENCIAS = 'incidencias.json'
ARCHIVO_EXCEL = 'Tratamientos_Por_Paciente.xlsx'

def cargar_json(archivo, valor_por_defecto):
    if not os.path.exists(archivo):
        with open(archivo, 'w') as f:
            json.dump(valor_por_defecto, f)
        return valor_por_defecto
    with open(archivo, 'r') as f:
        return json.load(f)

def guardar_json(archivo, datos):
    with open(archivo, 'w') as f:
        json.dump(datos, f)

# Funciones de sincronización con Google Sheets
def obtener_solicitudes():
    try:
        response = requests.get(f"{WEB_APP_URL}?action=leer")
        if response.status_code == 200:
            return response.json()
    except Exception as e:
        print(f"Error al leer de Google Sheets: {e}")
    return []

def guardar_solicitud(item):
    try:
        payload = {"action": "insertar"}
        payload.update(item)
        requests.post(WEB_APP_URL, json=payload)
    except Exception as e:
        print(f"Error al guardar en Google Sheets: {e}")

def actualizar_estado_solicitud(sol_id, nuevo_estado, fecha_albaran=None):
    try:
        payload = {"action": "actualizar", "id": sol_id, "estado": nuevo_estado}
        if fecha_albaran:
            payload["fecha_albaran"] = fecha_albaran
        requests.post(WEB_APP_URL, json=payload)
    except Exception as e:
        print(f"Error al actualizar en Google Sheets: {e}")

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
    col_l1, col_l2, col_l3 = st.columns([0.1, 2, 0.1])
    with col_l2:
        st.markdown("<br><h2 style='text-align: center; color: #0066cc;'>💊 SPD FARMACIA VILLEGAS</h2><p style='text-align: center; color: #666;'>Sincronización Cloud (Google Sheets)</p>", unsafe_allow_html=True)
        with st.form("login_form"):
            usuario_input = st.text_input("👤 USUARIO")
            password_input = st.text_input("🔑 CONTRASEÑA", type="password")
            st.markdown("<br>", unsafe_allow_html=True)
            if st.form_submit_button("🚀 INICIAR SESIÓN"):
                if usuario_input in usuarios_db and usuarios_db[usuario_input]["password"] == password_input:
                    st.session_state['usuario'] = usuario_input
                    st.session_state['rol'] = usuarios_db[usuario_input]["rol"]
                    st.session_state['pantalla'] = 'menu'
                    st.rerun()
                else:
                    st.error("❌ USUARIO O CONTRASEÑA INCORRECTOS.")

# --- APLICACIÓN PRINCIPAL ---
else:
    es_admin = (st.session_state['rol'] == 'administrador')
    solicitudes_db = obtener_solicitudes()
    
    with st.sidebar:
        st.markdown(f"### 👩‍⚕️ SESIÓN ACTIVA")
        st.info(f"Usuario: **{st.session_state['usuario'].upper()}**")
        
        if st.button("🏠 MENÚ PRINCIPAL", key="sb_home", use_container_width=True):
            st.session_state['pantalla'] = 'menu'; st.rerun()
        if st.button("🚨 VER INCIDENCIAS", key="sb_inc", use_container_width=True):
            st.session_state['pantalla'] = 'panel_incidencias'; st.rerun()
        if st.button("📦 PEDIDO FUERA DE BLISTER", key="sb_pre", use_container_width=True):
            st.session_state['pantalla'] = 'panel_presolicitudes'; st.rerun()
            
        if es_admin:
            pedidos_blister_pendientes = [s for s in solicitudes_db if str(s.get('tipo')) == 'FUERA DE BLISTER (CONFIRMADO)' and str(s.get('estado')) == 'pendiente']
            if pedidos_blister_pendientes:
                if st.button(f"🔔 ALBARÁN DE ENTREGA ({len(pedidos_blister_pendientes)})", key="sb_ped_bli", use_container_width=True):
                    st.session_state['pantalla'] = 'admin_gestion_datamatrix'; st.rerun()

        if st.button("🚪 CERRAR SESIÓN", key="sb_logout", use_container_width=True):
            st.session_state['usuario'] = None; st.session_state['rol'] = None; st.rerun()

    @st.cache_data
    def cargar_datos():
        if not os.path.exists(ARCHIVO_EXCEL): return None, "No se encuentra el Excel de tratamientos."
        try:
            xls = pd.ExcelFile(ARCHIVO_EXCEL)
            df = pd.concat([xls.parse(sheet) for sheet in xls.sheet_names], ignore_index=True)
            df.columns = df.columns.str.strip().str.upper()
            return df, None
        except Exception as e: return None, str(e)

    datos, error_carga = cargar_datos()
    if datos is None:
        st.error(f"⚠️ {error_carga}")
        st.stop()

    def obtener_valor(fila, posibles_nombres, por_defecto="N/A"):
        for col in fila.index:
            if str(col).strip().upper() in [n.upper() for n in posibles_nombres]:
                val = fila[col]
                if pd.isna(val) or str(val).strip() == "": return por_defecto
                return str(val)
        return por_defecto

    col_nombre = next((c for c in datos.columns if 'NOMBRE' in str(c).upper()), None)
    col_codigo = next((c for c in datos.columns if 'CODIGO' in str(c).upper() or 'CÓDIGO' in str(c).upper()), None)

    # --- MENÚ PRINCIPAL ---
    if st.session_state['pantalla'] == 'menu':
        st.markdown("<h1 style='text-align:center;'>🏥 SPD FARMACIA</h1>", unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        
        # ==========================================
        # --- CENTRO DE AVISOS Y GESTIÓN INTELIGENTE ---
        # ==========================================
        st.markdown("<div class='avisos-container'>", unsafe_allow_html=True)
        st.markdown("<h3 style='margin-top: 0; color: #1a334e;'>🔔 CENTRO DE AVISOS Y TAREAS</h3>", unsafe_allow_html=True)
        
        if es_admin:
            solicitudes_pendientes = [s for s in solicitudes_db if str(s.get('estado')).strip().lower() == 'pendiente' and str(s.get('tipo')).strip().upper() != 'FUERA DE BLISTER (CONFIRMADO)']
            pedidos_blister_pendientes = [s for s in solicitudes_db if str(s.get('tipo')).strip().upper() == 'FUERA DE BLISTER (CONFIRMADO)' and str(s.get('estado')).strip().lower() == 'pendiente']
            
            total_tareas_admin = len(solicitudes_pendientes) + len(pedidos_blister_pendientes) + len(incidencias_db)
            
            if total_tareas_admin == 0:
                st.success("✨ ¡Todo al día! No hay tareas pendientes de resolver.")
            else:
                st.info(f"Tienes **{total_tareas_admin}** asunto(s) pendientes de administración:")
                
                if solicitudes_pendientes:
                    st.markdown("#### 📋 Solicitudes de Altas / Bajas / Fármacos:")
                    for sol in solicitudes_pendientes:
                        sol_id = sol.get('id')
                        with st.expander(f"📌 [{sol.get('tipo')}] - {sol.get('paciente')} (Solicitante: {sol.get('solicitante', 'N/A')})"):
                            st.markdown(f"**Detalles / Motivo:** {sol.get('notas')}")
                            st.markdown(f"📅 *Fecha:* {sol.get('fecha_solicitud', 'N/A')}")
                            c1, c2 = st.columns(2)
                            with c1:
                                if st.button("✅ APROBAR", key=f"cent_apr_{sol_id}", use_container_width=True):
                                    actualizar_estado_solicitud(sol_id, 'aprobada')
                                    st.success("¡Solicitud aprobada correctamente!")
                                    st.rerun()
                            with c2:
                                if st.button("❌ RECHAZAR", key=f"cent_rech_{sol_id}", use_container_width=True):
                                    actualizar_estado_solicitud(sol_id, 'rechazada')
                                    st.error("Solicitud rechazada.")
                                    st.rerun()

                if pedidos_blister_pendientes:
                    st.markdown("#### 📦 Pedidos Fuera de Blíster:")
                    st.warning(f"Hay **{len(pedidos_blister_pendientes)}** pedido(s) pendientes de Albarán de Entrega.")
                    if st.button("🖨️ IR A GENERAR ALBARÁN DE ENTREGA", key="btn_ir_albaran", use_container_width=True):
                        st.session_state['pantalla'] = 'admin_gestion_datamatrix'
                        st.rerun()

                if incidencias_db:
                    st.markdown("#### 🚨 Incidencias Activas:")
                    incidencias_a_resolver = []
                    for clave, detalles in incidencias_db.items():
                        pac = clave.split("____")[0] if "____" in clave else "Desconocido"
                        med = clave.split("____")[1] if "____" in clave else clave
                        st.markdown(f"• **{pac}** - Fármaco: *{med}* ({detalles.get('tipo', 'Incidencia')})")
                        if st.button(f"Resolver incidencia de {pac}", key=f"res_aviso_{clave}", use_container_width=True):
                            incidencias_a_resolver.append(clave)
                    
                    if incidencias_a_resolver:
                        for c in incidencias_a_resolver:
                            del incidencias_db[c]
                        guardar_json(ARCHIVO_INCIDENCIAS, incidencias_db)
                        st.success("✅ Incidencia resuelta y retirada de avisos.")
                        st.rerun()

        else:
            mis_solicitudes = [s for s in solicitudes_db if str(s.get('solicitante')).strip().lower() == str(st.session_state['usuario']).strip().lower()]
            pendientes_mias = [s for s in mis_solicitudes if str(s.get('estado')).strip().lower() == 'pendiente']
            resueltas_mias = [s for s in mis_solicitudes if str(s.get('estado')).strip().lower() in ['aprobada', 'rechazada']]
            
            if not mis_solicitudes:
                st.info("ℹ️ No has enviado ninguna solicitud reciente. Utiliza los botones inferiores para tramitar altas, bajas o pedidos.")
            else:
                st.markdown(f"**Estado de tus solicitudes enviadas:** (Pendientes: {len(pendientes_mias)} | Resueltas: {len(resueltas_mias)})")
                for sol in mis_solicitudes:
                    estado_txt = str(sol.get('estado')).upper()
                    color_badge = "#ffc107" if estado_txt == "PENDIENTE" else ("#28a745" if estado_txt == "APROBADA" else "#dc3545")
                    
                    st.markdown(f"""
                        <div style='background-color: #f8f9fa; padding: 10px; border-radius: 8px; margin-bottom: 8px; border-left: 4px solid {color_badge};'>
                            <b>[{sol.get('tipo')}]</b> Paciente: {sol.get('paciente')}<br>
                            <span style='font-size: 0.9em; color: #555;'>Detalles: {sol.get('notas')}</span><br>
                            <b>Estado:</b> <span style='color: {color_badge};'><b>{estado_txt}</b></span>
                        </div>
                    """, unsafe_allow_html=True)
                    
        st.markdown("</div>", unsafe_allow_html=True)
        # ==========================================

        r1_c1, r1_c2 = st.columns(2)
        with r1_c1:
            if st.button("👥 DIRECTORIO PACIENTES", key="grid_pac", use_container_width=True):
                st.session_state['pantalla'] = 'pacientes'; st.rerun()
        with r1_c2:
            if st.button("🚨 PANEL INCIDENCIAS", key="grid_inc", use_container_width=True):
                st.session_state['pantalla'] = 'panel_incidencias'; st.rerun()
                
        r2_c1, r2_c2 = st.columns(2)
        with r2_c1:
            if st.button("➕ SOLICITUD ALTAS", key="grid_alt", use_container_width=True):
                st.session_state['pantalla'] = 'alta'; st.rerun()
        with r2_c2:
            if st.button("➖ SOLICITUD BAJAS", key="grid_baj", use_container_width=True):
                st.session_state['pantalla'] = 'baja'; st.rerun()
                
        r3_c1, r3_c2 = st.columns(2)
        with r3_c1:
            if st.button("📦 PEDIDO FUERA DE BLISTER", key="grid_pre", use_container_width=True):
                st.session_state['pantalla'] = 'panel_presolicitudes'; st.rerun()
        with r3_c2:
            if es_admin:
                if st.button("🖨️ ALBARÁN DE ENTREGA", key="grid_dm", use_container_width=True):
                    st.session_state['pantalla'] = 'admin_gestion_datamatrix'; st.rerun()
            else:
                if st.button("ℹ️ INFO Y AYUDA", key="grid_info", use_container_width=True):
                    st.info("ℹ️ Sistema sincronizado con Google Sheets.")

    elif st.session_state['pantalla'] == 'admin_gestion_datamatrix':
        if not es_admin:
            st.warning("⚠️ Acceso restringido exclusivamente al perfil de administrador.")
            st.session_state['pantalla'] = 'menu'
            st.rerun()
            
        if st.button("⬅️ VOLVER AL MENÚ", key="vol_dm", use_container_width=True):
            st.session_state['pantalla'] = 'menu'; st.rerun()
            
        st.markdown("<h2>🖨️ ALBARÁN DE ENTREGA</h2>", unsafe_allow_html=True)
        pedidos_blister = [s for s in solicitudes_db if str(s.get('tipo')).strip().upper() == 'FUERA DE BLISTER (CONFIRMADO)' and str(s.get('estado')).strip().lower() == 'pendiente']
        
        if not pedidos_blister:
            st.info("✨ No hay solicitudes de blister pendientes. Todo limpio y listo.")
        else:
            with st.form("form_lectura_datamatrix"):
                datos_ingresados = {}
                for idx, sol in enumerate(pedidos_blister):
                    sol_id = sol.get('id')
                    med_clean = sol.get('medicamento', '')
                    fecha_sol = sol.get('fecha_solicitud', 'N/A')
                    st.markdown(f"**Paciente:** 👤 {sol.get('paciente')} | **Fármaco:** 💊 {med_clean}")
                    
                    cn_input = st.text_input("CN", value=sol.get('cn', ''), key=f"cn_{sol_id}")
                    lote_input = st.text_input("LOTE", key=f"lote_{sol_id}")
                    cad_input = st.text_input("CADUCIDAD", key=f"cad_{sol_id}")
                    
                    datos_ingresados[sol_id] = {"nombre": sol.get('paciente'), "med": med_clean, "cn": cn_input, "lote": lote_input, "cad": cad_input}
                    st.markdown("---")
                
                if st.form_submit_button("🖨️ GENERAR PDF DE ALBARÁN"):
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
                    pdf.cell(85, 10, "MEDICAMENTO", 1)
                    pdf.cell(25, 10, "CN", 1)
                    pdf.cell(35, 10, "LOTE / CAD", 1)
                    pdf.ln()
                    
                    pdf.set_font("Arial", "", 9)
                    for sol in pedidos_blister:
                        sol_id = sol.get('id')
                        d = datos_ingresados[sol_id]
                        actualizar_estado_solicitud(sol_id, 'aprobada', fecha_albaran=fecha_gen)
                        pdf.cell(45, 10, str(d['nombre'][:22]), 1)
                        pdf.cell(85, 10, str(d['med'][:45]), 1)
                        pdf.cell(25, 10, str(d['cn'] or '000000'), 1)
                        pdf.cell(35, 10, f"{d['lote'] or 'L-01'} / {d['cad'] or '12/2028'}", 1)
                        pdf.ln()
                        
                    pdf_output = pdf.output(dest='S').encode('latin1')
                    st.success("✅ ¡Albarán generado con éxito!")
                    st.download_button(
                        label="📥 DESCARGAR ALBARÁN PDF EN TU EQUIPO",
                        data=pdf_output,
                        file_name=f"albaran_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                        mime="application/pdf"
                    )

    elif st.session_state['pantalla'] == 'panel_presolicitudes':
        if st.button("⬅️ VOLVER AL MENÚ", key="vol_pre", use_container_width=True):
            st.session_state['pantalla'] = 'menu'; st.rerun()
            
        st.markdown("<h2>📦 PEDIDO FUERA DE BLISTER</h2>", unsafe_allow_html=True)
        presolicitudes_locales = cargar_json('presolicitudes_blister.json', [])
        
        if not presolicitudes_locales:
            st.info("✨ No hay medicamentos señalados para fuera de blíster.")
        else:
            with st.form("form_gestion_presolicitudes"):
                checks_idx = []
                for idx, item in enumerate(presolicitudes_locales):
                    marcado = st.checkbox(f"👤 {item.get('paciente')} - 💊 {item.get('medicamento')} (CN: {item.get('cn', 'N/A')})", key=f"chk_p_{idx}")
                    if marcado: checks_idx.append(idx)
                
                if st.form_submit_button("🚀 ENVIAR A FARMACIA"):
                    if checks_idx:
                        fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                        nuevos_restantes = []
                        for idx, item in enumerate(presolicitudes_locales):
                            if idx in checks_idx:
                                guardar_solicitud({
                                    "tipo": "FUERA DE BLISTER (CONFIRMADO)",
                                    "paciente": item.get('paciente'),
                                    "medicamento": item.get('medicamento'),
                                    "cn": item.get('cn', 'N/A'),
                                    "solicitante": st.session_state['usuario'],
                                    "notas": f"Solicitado fuera de blíster: {item.get('medicamento')}",
                                    "estado": "pendiente",
                                    "fecha_solicitud": fecha_actual
                                })
                            else:
                                nuevos_restantes.append(item)
                        
                        guardar_json('presolicitudes_blister.json', nuevos_restantes)
                        st.success("✅ ¡Enviado a la farmacia y reiniciado!")
                        st.rerun()
                    else:
                        st.warning("⚠️ Selecciona al menos un medicamento.")

    elif st.session_state['pantalla'] == 'pacientes':
        if st.button("⬅️ VOLVER AL MENÚ", key="vol_menu_dir", use_container_width=True):
            st.session_state['pantalla'] = 'menu'; st.rerun()
            
        st.markdown("<h2>👥 DIRECTORIO DE PACIENTES</h2>", unsafe_allow_html=True)
        busq = st.text_input("🔍 Buscar paciente:")
        df_p = datos.copy()
        df_p['N_F'] = df_p[col_nombre].fillna('Desconocido').astype(str)
        df_p = df_p[['N_F']].drop_duplicates()
        if busq: df_p = df_p[df_p['N_F'].str.lower().str.contains(busq.lower())]
        
        for idx, fila in df_p.iterrows():
            if st.button(f"👤 {fila['N_F']}", key=f"p_{idx}", use_container_width=True):
                st.session_state['paciente_actual'] = fila['N_F']
                st.session_state['pantalla'] = 'ficha_paciente'; st.rerun()

    elif st.session_state['pantalla'] == 'ficha_paciente':
        paciente = st.session_state['paciente_actual']
        if st.button("⬅️ VOLVER AL DIRECTORIO", key="vol_dir", use_container_width=True):
            st.session_state['pantalla'] = 'pacientes'; st.rerun()
            
        st.markdown(f"<h2>📋 {paciente}</h2>", unsafe_allow_html=True)
        tab1, tab2 = st.tabs(["💊 TRATAMIENTOS", "➕ NUEVO FÁRMACO"])
        
        with tab1:
            meds = datos[datos[col_nombre].astype(str) == paciente]
            with st.form("form_t"):
                blisters = []
                for idx, fila in meds.iterrows():
                    med = obtener_valor(fila, ['MEDICAMENTO', 'MEDICINA'])
                    cn = obtener_valor(fila, ['CN', 'C.N.', 'CODIGO'])
                    st.markdown(f"**{med}** (CN: {cn})")
                    if es_admin and st.checkbox("📦 FUERA DE BLISTER", key=f"b_{idx}"):
                        blisters.append({"paciente": paciente, "medicamento": med, "cn": cn})
                    st.markdown("---")
                
                if es_admin and st.form_submit_button("📨 AÑADIR A FUERA DE BLISTER"):
                    if blisters:
                        presolicitudes_locales = cargar_json('presolicitudes_blister.json', [])
                        for b_item in blisters:
                            if not any(p['paciente'] == b_item['paciente'] and p['medicamento'] == b_item['medicamento'] for p in presolicitudes_locales):
                                presolicitudes_locales.append(b_item)
                        guardar_json('presolicitudes_blister.json', presolicitudes_locales)
                        st.success("✅ ¡Añadido al pedido fuera de blíster!")

        with tab2:
            with st.form("form_n"):
                n_med = st.text_input("💊 MEDICAMENTO:")
                n_pau = st.text_input("🕒 PAUTA:")
                if st.form_submit_button("ENVIAR SOLICITUD"):
                    if n_med:
                        guardar_solicitud({
                            "tipo": "NUEVO MEDICAMENTO", "paciente": paciente,
                            "solicitante": st.session_state['usuario'],
                            "notas": f"{n_med} - {n_pau}", "estado": "pendiente",
                            "fecha_solicitud": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                        })
                        st.success("✅ Solicitud enviada.")

    elif st.session_state['pantalla'] == 'alta':
        if st.button("⬅️ VOLVER AL MENÚ", key="vol_alta", use_container_width=True):
            st.session_state['pantalla'] = 'menu'; st.rerun()
        st.markdown("<h2>➕ SOLICITUD DE ALTA DE PACIENTE</h2>", unsafe_allow_html=True)
        with st.form("form_alta"):
            nombre = st.text_input("👤 NOMBRE DEL NUEVO PACIENTE:")
            notas = st.text_area("📝 OBSERVACIONES / TRATAMIENTO INICIAL:")
            if st.form_submit_button("ENVIAR SOLICITUD DE ALTA"):
                if nombre:
                    guardar_solicitud({
                        "tipo": "ALTA PACIENTE", "paciente": nombre,
                        "solicitante": st.session_state['usuario'], "notas": notas, "estado": "pendiente",
                        "fecha_solicitud": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                    })
                    st.success("✅ Solicitud de alta enviada al administrador para su revisión.")
                else:
                    st.warning("⚠️ Debes introducir el nombre del paciente.")

    elif st.session_state['pantalla'] == 'baja':
        if st.button("⬅️ VOLVER AL MENÚ", key="vol_baja", use_container_width=True):
            st.session_state['pantalla'] = 'menu'; st.rerun()
        st.markdown("<h2>➖ SOLICITUD DE BAJA DE PACIENTE</h2>", unsafe_allow_html=True)
        df_p = datos.copy()
        lista = sorted(df_p[col_nombre].fillna('Desconocido').astype(str).unique())
        with st.form("form_baja"):
            pac = st.selectbox("👤 PACIENTE:", lista)
            motivo = st.text_area("📝 MOTIVO DE LA BAJA:")
            if st.form_submit_button("ENVIAR SOLICITUD DE BAJA"):
                if motivo:
                    guardar_solicitud({
                        "tipo": "BAJA PACIENTE", "paciente": pac,
                        "solicitante": st.session_state['usuario'], "notas": motivo, "estado": "pendiente",
                        "fecha_solicitud": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                    })
                    st.success("✅ Solicitud de baja enviada al administrador para su revisión.")
                else:
                    st.warning("⚠️ Debes indicar el motivo de la baja.")

    elif st.session_state['pantalla'] == 'panel_incidencias':
        if st.button("⬅️ VOLVER AL MENÚ", key="vol_inc", use_container_width=True):
            st.session_state['pantalla'] = 'menu'; st.rerun()
        st.markdown("<h2>🚨 PANEL DE INCIDENCIAS</h2>", unsafe_allow_html=True)
        
        if not incidencias_db:
            st.info("✨ No hay incidencias activas.")
        else:
            incidencias_a_borrar = []
            for clave, detalles in incidencias_db.items():
                pac = clave.split("____")[0] if "____" in clave else "Desconocido"
                med = clave.split("____")[1] if "____" in clave else clave
                
                st.markdown(f"""
                    <div class='card-incidencia'>
                        <b>👤 Paciente:</b> {pac}<br>
                        <b>💊 Fármaco:</b> {med}<br>
                        <b>Tipo:</b> {detalles.get('tipo', 'N/A')}<br>
                        <b>Observaciones:</b> {detalles.get('observaciones', 'N/A')}
                    </div>
                """, unsafe_allow_html=True)
                
                if st.button(f"🗑️ Resolver e eliminar incidencia", key=f"res_inc_{clave}", use_container_width=True):
                    incidencias_a_borrar.append(clave)
            
            if incidencias_a_borrar:
                for c in incidencias_a_borrar:
                    del incidencias_db[c]
                guardar_json(ARCHIVO_INCIDENCIAS, incidencias_db)
                st.success("✅ Incidencia resuelta y eliminada correctamente.")
                st.rerun()
