#app.py

import streamlit as st

st.set_page_config(page_title="Hojas de Verificación", layout="wide")

if "inventario_df" not in st.session_state:
    st.session_state.inventario_df = None

if "nombre_ingeniero" not in st.session_state:
    st.session_state.nombre_ingeniero = ""

if "nombre_jefe" not in st.session_state:
    st.session_state.nombre_jefe = ""

if "nombre_hospital" not in st.session_state:
    st.session_state.nombre_hospital = ""

pagina_1 = st.Page("pages/1_Hojas de Verificacion.py", title="Hojas de Verificación")
pagina_2 = st.Page("pages/2_Inventario.py", title="Cargar Inventario")
pagina_3 = st.Page("pages/3_Informacion del usuario.py", title="Información del Usuario")

pg = st.navigation([pagina_1, pagina_2, pagina_3], position="hidden")

with st.sidebar:
    st.image("image_5976e1.png", use_container_width=True)
    st.divider()
    st.markdown("### Menú")
    st.page_link(pagina_1, label="🔧 Hojas de Verificación")
    st.page_link(pagina_2, label="📂 Inventario")
    st.page_link(pagina_3, label="👤 Información del Usuario")

pg.run()