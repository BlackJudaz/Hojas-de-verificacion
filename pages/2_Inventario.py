import streamlit as st
from utils.lector_inventario import cargar_inventario

st.title("Gestión de Inventario")
st.divider()

if st.session_state.inventario_df is None:
    archivo = st.file_uploader(
        label="Sube tu archivo Excel de inventario",
        type=["xlsx"]
    )

    if archivo is not None:
        resultado = cargar_inventario(archivo)

        df_temporal = None
        error = None

        if isinstance(resultado, tuple):
            if len(resultado) == 3:
                df_temporal, error, _ = resultado
            elif len(resultado) == 2:
                df_temporal, error = resultado
        else:
            error = "Error al procesar el inventario: respuesta no válida."

        if df_temporal is not None:
            st.session_state.inventario_df = df_temporal
            st.success(f"Inventario guardado: {len(df_temporal)} equipos encontrados.")
            st.info("Dirígete a 'Hojas de Verificación' en el menú lateral.")
            st.button("Actualizar vista", on_click=st.rerun)
        else:
            st.error(error or "No se pudo leer el archivo.")
else:
    df = st.session_state.inventario_df
    st.success(f"Inventario cargado con {len(df)} equipos.")
    st.dataframe(df.head(10), use_container_width=True)

    if st.button("Eliminar e Importar otro inventario"):
        st.session_state.inventario_df = None
        st.rerun()