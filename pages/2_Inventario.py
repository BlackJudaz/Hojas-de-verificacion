#Inventario

import streamlit as st
from utils.lector_inventario import cargar_inventario

st.title("Gestión de Inventario")
st.divider()

# Si no hay inventario, se muestra el uploader de forma habitual
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
                error = "Error al procesar el resultado de carga del inventario."
        else:
            error = "Error al procesar el inventario: respuesta no válida."

        if df_temporal is not None:
            st.session_state.inventario_df = df_temporal
            st.success(f"Inventario guardado en la sesión: {len(df_temporal)} equipos encontrados.")
            st.info("Dirigete a la pestaña 'Hojas de Verificación' en el menú de la izquierda.")
            st.button("Actualizar vista", on_click=st.rerun)
        else:
            st.error(error or "No se pudo leer el archivo. Verifica que sea el formato correcto.")
    else:
        st.write("")

# Si ya existe un inventario guardado, mostramos un resumen y opción de borrarlo
else:
    df = st.session_state.inventario_df
    st.success(f"Actualmente hay un inventario cargado con {len(df)} equipos.")
    
    # Mostramos una vista previa corta de los datos guardados
    st.dataframe(df.head(10), use_container_width=True)
    
    if st.button("Eliminar e Importar otro inventario"):
        st.session_state.inventario_df = None
        st.rerun()