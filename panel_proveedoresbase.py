import streamlit as st
import pandas as pd
from datetime import datetime, date
import io
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import gspread
from google.oauth2.service_account import Credentials

# --- 1. CONFIGURACIÓN ---
st.set_page_config(page_title="Panel RMA Proveedores", layout="wide")

# Estilos CSS Originales
st.markdown("""
    <style>
        .block-container { max-width: 100% !important; padding-left: 2rem !important; padding-right: 2rem !important; padding-top: 4rem; }
        div[data-testid="stExpander"] { border: 1px solid #444; margin-bottom: 1rem; }
        [data-testid="stDataEditor"] div, .stDataTable td { border-bottom: 4px solid #000 !important; }
        .stDataTable td, .stDataTable th, [data-testid="stDataEditor"] * { font-family: sans-serif !important; font-size: 14px !important; font-weight: 400 !important; }
        .stDataTable td, .stDataTable th { border-right: 1px solid #444 !important; }
    </style>
    """, unsafe_allow_html=True)

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
    st.session_state.usuario = ""
    st.session_state.rol = ""

# --- LOGIN ---
def login():
    st.markdown("<h2 style='text-align: center;'>Control de Acceso - Panel RMA</h2>", unsafe_allow_html=True)
    with st.form("formulario_login"):
        usuario = st.text_input("Usuario:").strip()
        clave = st.text_input("Contraseña:", type="password").strip()
        bot_login = st.form_submit_button("Iniciar Sesión", use_container_width=True)
        
        if bot_login:
            if "USUARIOS" not in st.secrets:
                st.error("Error: No se encontró la sección [USUARIOS] en los Secrets.")
                return
            usuarios_secretos = st.secrets["USUARIOS"]
            if usuario in usuarios_secretos and str(usuarios_secretos[usuario]) == clave:
                st.session_state.autenticado = True
                st.session_state.usuario = usuario
                st.session_state.rol = "admin" if usuario == "admin" else "user"
                st.success("¡Acceso concedido!")
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos.")

if not st.session_state.autenticado:
    login()
    st.stop()

# --- BARRA LATERAL ---
st.sidebar.write(f"Conectado como: **{st.session_state.usuario}** ({st.session_state.rol.upper()})")
if st.sidebar.button("Cerrar Sesión", type="secondary", use_container_width=True):
    st.session_state.autenticado = False
    st.session_state.usuario = ""
    st.session_state.rol = ""
    st.rerun()

# --- CONEXIÓN DIRECTA A GOOGLE SHEETS (VERSIÓN SEGURA PARA GIT) ---
@st.cache_resource
def conectar_google_sheets():
    try:
        # Definimos los accesos requeridos para Drive y Sheets
        scope = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        
        # Intentamos cargar la cuenta de servicio desde los Secrets de Streamlit
        if "gcp_service_account" in st.secrets:
            # Convertimos explícitamente a diccionario estándar de Python
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
        else:
            # Fallback de respaldo por si sigues usando el archivo local en tu entorno de desarrollo
            creds = Credentials.from_service_account_file("creds_google.json", scopes=scope)
        
        gc = gspread.authorize(creds)
        
        # Abrimos el archivo utilizando el cliente de gspread directamente
        sh = gc.open("Proveedores")
        return sh
    except gspread.exceptions.SpreadsheetNotFound:
        st.error("❌ Error: No se encontró la planilla 'RMA Proveedores Base' en tu cuenta de Google. Revisa que el nombre sea exacto.")
        st.stop()
    except Exception as e:
        st.error(f"❌ Falló la conexión con Google de forma inesperada: {str(e)}")
        st.stop()

# Inicializamos el objeto del libro
sh = conectar_google_sheets()

# Conectamos las pestañas asegurando la lectura correcta de datos
try:
    ws_prov = sh.worksheet("Proveedores")
    ws_rma = sh.worksheet("RMA_Proveedores")
except gspread.exceptions.WorksheetNotFound as e:
    st.error(f"❌ Error de pestañas: No se encontró la hoja dentro del archivo. {str(e)}")
    st.stop()

# --- FUNCIONES DE ASISTENCIA ---
def despachar_correo(config_section, destinatario, asunto, cuerpo_texto):
    try:
        if config_section not in st.secrets: return False
        smtp_user = st.secrets[config_section]["SMTP_USER"]
        smtp_password = st.secrets[config_section]["SMTP_PASSWORD"]
        msg = MIMEMultipart()
        msg['From'] = smtp_user
        msg['To'] = str(destinatario).strip().lower()
        msg['Subject'] = asunto
        msg.attach(MIMEText(cuerpo_texto, 'plain', 'utf-8'))
        server = smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=15)
        server.login(smtp_user, smtp_password)
        server.sendmail(smtp_user, str(destinatario).strip().lower(), msg.as_string())
        server.quit()
        return True
    except Exception:
        return False

def estilo_filas(row):
    estado = str(row.get('Estado del RMA', "")).upper()
    verde, naranja, celeste, rojo, gris = 'background-color: #28a745; color: white;', 'background-color: #fd7e14; color: black;', 'background-color: #17a2b8; color: white;', 'background-color: #dc3545; color: white;', 'background-color: #6c757d; color: white;'
    style = ''
    if estado in ["CAMBIO", "CREDITO"]: style = verde
    elif estado in ["GARANTIA", "GARANTIA OFICIAL"]: style = naranja
    elif estado == "NO FALLO - DEVOLVER A CLIENTE": style = celeste
    elif estado == "FUERA DE GARANTIA": style = rojo
    elif estado == "REPARADO": style = gris
    return [style for _ in row.index]

def formatear_para_leer(fecha_raw):
    if not fecha_raw or str(fecha_raw).strip() in ["None", "none", "nan", "NaN", ""]: return ""
    fecha_str = str(fecha_raw).replace('-', '/').strip()
    for formato in ['%Y/%m/%d', '%Y-%m-%d', '%d/%m/%Y', '%d/%m/%y']:
        try:
            return datetime.strptime(fecha_str, formato).strftime('%d/%m/%Y')
        except ValueError: continue
    return str(fecha_raw)

# --- CARGAR DATOS DESDE GOOGLE SHEETS ---
@st.cache_data(ttl=5)
def cargar_todos_los_datos():
    data = ws_rma.get_all_records()
    if not data: return pd.DataFrame()
    df = pd.DataFrame(data)
    # Creamos un ID basado en el número de fila física de Google Sheets (comienza en fila 2)
    df['id_interno'] = range(2, len(df) + 2)
    return df

df_all = cargar_todos_los_datos()

# --- BOTÓN DE REFRESH ---
if st.button("🔄 Actualizar Datos", use_container_width=True):
    cargar_todos_los_datos.clear()
    st.rerun()

if df_all.empty:
    st.warning("La planilla 'RMA_Proveedores' está vacía o no tiene registros.")
    st.stop()

# --- VALIDACIÓN DE COLUMNAS ---
columnas_requeridas = ['Aceptado', 'Finalizado', 'Ingreso', 'Resolucion', 'diagnostico', 'Estado del RMA', 'Compra', 'Producto', 'comentario', 'Falla', 'Serial', 'autonumero', 'Cliente']
for col in columnas_requeridas:
    if col not in df_all.columns: df_all[col] = False if col in ['Aceptado', 'Finalizado'] else ""

# Normalización de textos
for col in df_all.columns:
    if col not in ['Aceptado', 'Finalizado', 'id_interno']:
        df_all[col] = df_all[col].astype(str).str.strip().replace(["None", "none", "nan", "NaN"], "")

# --- TABLA 1: POR ACEPTAR ---
df1 = df_all[(df_all['Aceptado'].astype(str).str.upper() != "TRUE") & (df_all['Producto'] != "")].copy()

with st.expander("📥 1. TICKETS POR ACEPTAR (Entrada)", expanded=True):
    if not df1.empty:
        df1['Compra'] = df1['Compra'].apply(formatear_para_leer)
        with st.form("f1"):
            c1_cols = ['Cliente', 'Producto', 'Serial', 'Falla', 'Compra', 'Aceptado']
            disabled_cols = ['Serial', 'Falla'] if st.session_state.rol == "admin" else c1_cols
            
            ed1 = st.data_editor(df1[['id_interno'] + c1_cols], column_config={"id_interno": None, "Aceptado": st.column_config.CheckboxColumn("Aceptar")}, disabled=disabled_cols, hide_index=True, use_container_width=True)
            
            if st.form_submit_button("GUARDAR ENTRADAS", disabled=(st.session_state.rol != "admin")):
                for _, r in ed1.iterrows():
                    if r['Aceptado'] == True:
                        idx_col_aceptado = ws_rma.row_values(1).index('Aceptado') + 1
                        ws_rma.update_cell(int(r['id_interno']), idx_col_aceptado, "TRUE")
                cargar_todos_los_datos.clear()
                st.rerun()
    else:
        st.info("No hay pendientes por aceptar.")

# --- TABLA 2: EN PROCESO ---
df2 = df_all[(df_all['Aceptado'].astype(str).str.upper() == "TRUE") & (df_all['Finalizado'].astype(str).str.upper() != "TRUE")].copy()

with st.expander("⚙️ 2. TICKETS EN PROCESO (Aceptados)", expanded=True):
    if not df2.empty:
        for c in ['Compra','Ingreso','Resolucion']: df2[c] = df2[c].apply(formatear_para_leer)
        with st.form("f2"):
            c2_cols = ['autonumero', 'Cliente', 'Producto', 'Serial', 'Falla', 'Ingreso', 'diagnostico', 'Estado del RMA', 'Finalizado']
            disabled_cols = ['autonumero', 'Cliente', 'Producto', 'Serial', 'Falla'] if st.session_state.rol == "admin" else c2_cols
            
            ed2 = st.data_editor(
                df2[['id_interno'] + c2_cols].style.apply(estilo_filas, axis=1),
                column_config={
                    "id_interno": None,
                    "Estado del RMA": st.column_config.SelectboxColumn(options=["CAMBIO", "CREDITO", "GARANTIA OFICIAL", "GARANTIA", "FUERA DE GARANTIA", "NO FALLO - DEVOLVER A CLIENTE", "REPARADO"])
                },
                disabled=disabled_cols, hide_index=True, use_container_width=True
            )
            
            if st.form_submit_button("ACTUALIZAR PROCESOS"):
                headers = ws_rma.row_values(1)
                for _, r in ed2.iterrows():
                    row_num = int(r['id_interno'])
                    
                    if 'diagnostico' in r and st.session_state.rol == "admin":
                        ws_rma.update_cell(row_num, headers.index('diagnostico') + 1, r['diagnostico'])
                    
                    if 'Estado del RMA' in r and st.session_state.rol == "admin":
                        ws_rma.update_cell(row_num, headers.index('Estado del RMA') + 1, r['Estado del RMA'])
                        
                    if r['Finalizado'] == True:
                        ws_rma.update_cell(row_num, headers.index('Finalizado') + 1, "TRUE")
                        ws_rma.update_cell(row_num, headers.index('Resolucion') + 1, date.today().strftime('%Y-%m-%d'))
                        
                cargar_todos_los_datos.clear()
                st.rerun()
    else:
        st.info("No hay tickets en proceso.")

# --- TABLA 3: HISTÓRICO ---
df3 = df_all[(df_all['Aceptado'].astype(str).str.upper() == "TRUE") & (df_all['Finalizado'].astype(str).str.upper() == "TRUE")].copy()
with st.expander("✅ 3. CASOS RESUELTOS (Histórico)"):
    if not df3.empty:
        df3['Resolucion'] = df3['Resolucion'].apply(formatear_para_leer)
        c3_cols = ['autonumero', 'comentario', 'Cliente', 'Producto', 'diagnostico', 'Estado del RMA', 'Resolucion']
        st.dataframe(df3[c3_cols].style.apply(estilo_filas, axis=1), hide_index=True, use_container_width=True)
