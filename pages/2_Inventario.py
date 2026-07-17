import streamlit as st
from utils.lector_inventario import (
    cargar_inventario,
    parsear_programacion_tinc,
    aplicar_programacion_tinc,
)

if "programacion_tinc_texto" not in st.session_state:
    st.session_state.programacion_tinc_texto = ""
if "programacion_tinc_df" not in st.session_state:
    st.session_state.programacion_tinc_df = None
if "programacion_tinc_html" not in st.session_state:
    st.session_state.programacion_tinc_html = ""

st.title("Gestión de Inventario")
st.caption("Carga tu inventario, pega la programación mensual y verifica qué equipos quedaron emparejados con su folio TiNC.")
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
            st.rerun()
        else:
            st.error(error or "No se pudo leer el archivo.")
else:
    df = st.session_state.inventario_df
    st.success(f"Inventario cargado con {len(df)} equipos.")

    total_equipos = len(df)
    total_emparejados = 0
    if "FOLIO TINC" in df.columns:
        total_emparejados = int(df["FOLIO TINC"].fillna("").astype(str).str.strip().ne("").sum())
    total_pendientes = max(total_equipos - total_emparejados, 0)

    col_resumen_1, col_resumen_2, col_resumen_3 = st.columns(3)
    col_resumen_1.metric("Equipos en inventario", total_equipos)
    col_resumen_2.metric("Mantenimientos preventivos por programar", total_emparejados)
    #col_resumen_3.metric("Equipos pendientes", total_pendientes)

    st.markdown("### Programación de mantenimiento del mes")
    st.caption("Pega aquí la programación mensual copiada desde TiNC y luego guárdala para emparejar los folios con el inventario.")

    texto_programacion = st.text_area(
        "Texto de programación TINC",
        key="programacion_tinc_texto",
        height=260,
        placeholder="Pega aquí manualmente el contenido de la programación del mes."
    )

    if st.button("Guardar programación TINC", type="primary", use_container_width=True):
        st.session_state.programacion_tinc_html = ""
        df_programacion = parsear_programacion_tinc(texto_programacion)
        st.session_state.programacion_tinc_df = df_programacion
        st.session_state.inventario_df = aplicar_programacion_tinc(df, df_programacion)
        st.success(
            f"Programación de mantenimientos preventivos mensuales cargada.\n"
            f"{len(df_programacion)} mantenimientos preventivos programados detectados."
        )
    inventario_verificacion = st.session_state.inventario_df.copy()
    if "FOLIO TINC" in inventario_verificacion.columns:
        inventario_verificacion = inventario_verificacion[inventario_verificacion["FOLIO TINC"].astype(str).str.strip() != ""]

    columnas_verificacion = [
        columna for columna in [
            "FOLIO TINC",
            "# ACTIVO",
            "CONCEPTO",
           # "URL TINC",
            "MARCA",
            "MODELO",
            "No. DE SERIE",
            "UBICACIÓN",
        ]
        if columna in inventario_verificacion.columns
    ]

    st.markdown("### Tabla de verificación del emparejamiento")
    st.caption("Aquí se muestran únicamente los equipos que ya quedaron emparejados con un folio TiNC.")
    if inventario_verificacion.empty:
        st.info("Aún no hay equipos con folio TiNC emparejado.")
    else:
        st.dataframe(
            inventario_verificacion[columnas_verificacion].head(50),
            use_container_width=True,
            hide_index=True,
            column_config={
                "URL TINC": st.column_config.LinkColumn(
                    "Liga TiNC",
                    display_text="Abrir enlace"
                )
            }
        )

        if len(inventario_verificacion) > 50:
            st.caption(f"Se muestran los primeros 50 registros emparejados de un total de {len(inventario_verificacion)}.")

    if st.button("Eliminar e Importar otro inventario"):
        st.session_state.inventario_df = None
        st.session_state.programacion_tinc_texto = ""
        st.session_state.programacion_tinc_df = None
        st.session_state.programacion_tinc_html = ""
        st.rerun()