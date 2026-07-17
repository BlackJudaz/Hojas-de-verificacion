from pathlib import Path
import streamlit as st

BASE_DIR = Path(__file__).resolve().parent

st.set_page_config(page_title="Hojas de Verificación", layout="wide")

if "inventario_df" not in st.session_state:
    st.session_state.inventario_df = None
if "nombre_ingeniero" not in st.session_state:
    st.session_state.nombre_ingeniero = ""
if "nombre_jefe" not in st.session_state:
    st.session_state.nombre_jefe = ""
if "nombre_hospital" not in st.session_state:
    st.session_state.nombre_hospital = ""
if "google_drive_credentials" not in st.session_state:
    st.session_state.google_drive_credentials = None
if "google_drive_usuario" not in st.session_state:
    st.session_state.google_drive_usuario = {}
if "ultimo_paquete_zip_bytes" not in st.session_state:
    st.session_state.ultimo_paquete_zip_bytes = b""
if "ultimo_paquete_zip_nombre" not in st.session_state:
    st.session_state.ultimo_paquete_zip_nombre = ""
if "ultimo_paquete_drive_folder" not in st.session_state:
    st.session_state.ultimo_paquete_drive_folder = ""
if "ultimo_paquete_periodo" not in st.session_state:
    st.session_state.ultimo_paquete_periodo = ""
if "ultimo_paquete_periodo_mixto" not in st.session_state:
    st.session_state.ultimo_paquete_periodo_mixto = False
if "ultimo_paquete_generado_en" not in st.session_state:
    st.session_state.ultimo_paquete_generado_en = ""

inventario_cargado = st.session_state.inventario_df is not None
informacion_usuario_completa = bool(
    str(st.session_state.nombre_ingeniero).strip() and str(st.session_state.nombre_hospital).strip()
)
paquete_listo_para_drive = bool(st.session_state.ultimo_paquete_zip_bytes)
drive_conectado = bool(st.session_state.google_drive_credentials)

pagina_1 = st.Page(str(BASE_DIR / "pages" / "1_Hojas de Verificacion.py"), title="Hojas de Verificación")
pagina_2 = st.Page(str(BASE_DIR / "pages" / "2_Inventario.py"), title="Cargar Inventario")
pagina_3 = st.Page(str(BASE_DIR / "pages" / "3_Informacion del usuario.py"), title="Información del Usuario")
pagina_4 = st.Page(str(BASE_DIR / "pages" / "4_Google Drive.py"), title="Google Drive")

pg = st.navigation([pagina_1, pagina_2, pagina_3, pagina_4], position="hidden")

with st.sidebar:
    st.image(str(BASE_DIR / "image_5976e1.png"), use_container_width=True)
    st.divider()
    st.markdown("### Menú")
    st.page_link(pagina_1, label="🔧 Hojas de Verificación")
    st.page_link(pagina_2, label="📂 Inventario")
    st.page_link(pagina_3, label="👤 Información del Usuario")
    st.page_link(pagina_4, label="☁️ Google Drive")
    st.divider()
    st.markdown("### Estado rápido")
    st.write(f"{'✅' if inventario_cargado else '⬜'} Inventario cargado")
    st.write(f"{'✅' if informacion_usuario_completa else '⬜'} Información del usuario lista")
    st.write(f"{'✅' if paquete_listo_para_drive else '⬜'} Último ZIP listo para Drive")
    st.write(f"{'✅' if drive_conectado else '⬜'} Google Drive conectado")
    #st.caption("Orden recomendado: 1) Información del usuario, 2) Inventario, 3) Hojas de verificación.")

pg.run()