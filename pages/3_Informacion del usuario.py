# pages/informacion del usuario.py
import random

import streamlit as st


def _guardar_informacion_usuario():
    st.session_state.nombre_ingeniero = st.session_state.input_nombre_ingeniero.strip()
    st.session_state.nombre_jefe = st.session_state.input_nombre_jefe.strip()
    st.session_state.nombre_hospital = st.session_state.input_nombre_hospital.strip()
    st.session_state.info_usuario_autoguardada = True


if "input_nombre_ingeniero" not in st.session_state:
    st.session_state.input_nombre_ingeniero = st.session_state.nombre_ingeniero
if "input_nombre_jefe" not in st.session_state:
    st.session_state.input_nombre_jefe = st.session_state.nombre_jefe
if "input_nombre_hospital" not in st.session_state:
    st.session_state.input_nombre_hospital = st.session_state.nombre_hospital
if "info_usuario_autoguardada" not in st.session_state:
    st.session_state.info_usuario_autoguardada = False

st.title("👤 Información del Usuario")
st.info("Esta información se usará para generar las hojas de verificación y las etiquetas. Completa al menos el nombre del ingeniero y el hospital antes de generar archivos.")
st.divider()

nombre_ejemplo = random.choice([
    "Ej. José Andrés Quijada López",
    "Ej. Alain Jesús Aguilera Noriega",
])

with st.container(border=True):
    col_form_1, col_form_2 = st.columns(2)
    with col_form_1:
        st.text_input(
            label="Nombre del Ingeniero de Servicio",
            key="input_nombre_ingeniero",
            placeholder=nombre_ejemplo,
            on_change=_guardar_informacion_usuario,
        )
        st.text_input(
            label="Nombre del Jefe de Departamento",
            key="input_nombre_jefe",
            placeholder="Ej. Dr. Juan Pérez",
            on_change=_guardar_informacion_usuario,
        )
    with col_form_2:
        st.text_input(
            label="Nombre del Hospital",
            key="input_nombre_hospital",
            placeholder="Ej. Medyarthros",
            on_change=_guardar_informacion_usuario,
        )
        st.caption("Campos obligatorios: nombre del ingeniero y nombre del hospital.")

    if st.session_state.info_usuario_autoguardada:
        st.success("✅ Información guardada automáticamente.")

st.divider()

if st.session_state.nombre_ingeniero or st.session_state.nombre_jefe or st.session_state.nombre_hospital:
    with st.container(border=True):
        st.markdown("### Información guardada")
        st.markdown(f"👤 **Ingeniero:** {st.session_state.nombre_ingeniero}")
        st.markdown(f"👔 **Jefe:** {st.session_state.nombre_jefe or 'No capturado'}")
        st.markdown(f"🏥 **Hospital:** {st.session_state.nombre_hospital}")
        if not st.session_state.nombre_ingeniero or not st.session_state.nombre_hospital:
            st.warning("⚠️ La información se guardó, pero aún faltan el nombre del ingeniero o el hospital para generar hojas o etiquetas.")
else:
    st.warning("⚠️ Aún no has guardado tu información. Completa este formulario antes de generar hojas o etiquetas.")
    ###