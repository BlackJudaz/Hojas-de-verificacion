import random

import streamlit as st
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
        nombre_ingeniero = st.text_input(
            label="Nombre del Ingeniero de Servicio",
            value=st.session_state.nombre_ingeniero,
            placeholder=nombre_ejemplo
        )
        nombre_jefe = st.text_input(
            label="Nombre del Jefe de Departamento",
            value=st.session_state.nombre_jefe,
            placeholder="Ej. Dr. Juan Pérez"
        )
    with col_form_2:
        nombre_hospital = st.text_input(
            label="Nombre del Hospital",
            value=st.session_state.nombre_hospital,
            placeholder="Ej. Medyarthros"
        )
        st.caption("Campos obligatorios: nombre del ingeniero y nombre del hospital.")

    if st.button("💾 Guardar", type="primary", use_container_width=True):
        if nombre_ingeniero.strip() == "" or nombre_hospital.strip() == "":
            st.error("❌ El nombre del ingeniero y el hospital son obligatorios.")
        else:
            st.session_state.nombre_ingeniero = nombre_ingeniero.strip()
            st.session_state.nombre_jefe      = nombre_jefe.strip()
            st.session_state.nombre_hospital  = nombre_hospital.strip()
            st.success("✅ Información guardada correctamente.")

st.divider()

if st.session_state.nombre_ingeniero:
    with st.container(border=True):
        st.markdown("### Información guardada")
        st.markdown(f"👤 **Ingeniero:** {st.session_state.nombre_ingeniero}")
        st.markdown(f"👔 **Jefe:** {st.session_state.nombre_jefe or 'No capturado'}")
        st.markdown(f"🏥 **Hospital:** {st.session_state.nombre_hospital}")
else:
    st.warning("⚠️ Aún no has guardado tu información. Completa este formulario antes de generar hojas o etiquetas.")