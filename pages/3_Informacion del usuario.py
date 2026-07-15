import streamlit as st

st.title("👤 Información del Usuario")
st.info("Esta información se usará para generar las hojas de verificación y etiquetas. Deberás llenarla cada vez que abras la aplicación.")
st.divider()

with st.container(border=True):
    nombre_ingeniero = st.text_input(
        label="Nombre del Ingeniero de Servicio",
        value=st.session_state.nombre_ingeniero,
        placeholder="Ej. Jose Andres Quijada Lopez"
    )
    nombre_jefe = st.text_input(
        label="Nombre del Jefe de Departamento",
        value=st.session_state.nombre_jefe,
        placeholder="Ej. Dr. Juan Pérez"
    )
    nombre_hospital = st.text_input(
        label="Nombre del Hospital",
        value=st.session_state.nombre_hospital,
        placeholder="Ej. Medyarthros"
    )

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
    st.markdown(f"👤 **Ingeniero:** {st.session_state.nombre_ingeniero}")
    st.markdown(f"👔 **Jefe:** {st.session_state.nombre_jefe}")
    st.markdown(f"🏥 **Hospital:** {st.session_state.nombre_hospital}")
else:
    st.warning("⚠️ Aún no has guardado tu información.")