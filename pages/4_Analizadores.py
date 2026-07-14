#Analizadores
import streamlit as st
from utils.lector_analizadores import RUTA_ANALIZADORES, cargar_analizadores, obtener_analizadores, filtrar_analizadores

st.title("Analizadores disponibles")
st.divider()

analizadores_df = cargar_analizadores()
if analizadores_df is None:
    st.error(
        f"No se pudo cargar la lista de analizadores. Verifica que exista el archivo '{RUTA_ANALIZADORES}' "
        "o que el proyecto se ejecute desde su carpeta raíz."
    )
    st.stop()

analizador_tipo = st.selectbox(
    label="Selecciona un tipo de analizador",
    options=[""] + obtener_analizadores(analizadores_df),
    format_func=lambda x: "-- Seleccionar --" if x == "" else x,
    key="analizador_tipo"
)

filtros = {}
if analizador_tipo:
    filtros["tipo"] = analizador_tipo

marca_opciones = []
modelo_opciones = []
if analizador_tipo:
    df_tipo = analizadores_df[analizadores_df["SIMULADOR / ANALIZADOR"] == analizador_tipo]
    marca_opciones = sorted(df_tipo["MARCA"].dropna().unique().tolist())
    modelo_opciones = sorted(df_tipo["MODELO"].dropna().unique().tolist())

marca = st.selectbox(
    label="Marca",
    options=[""] + marca_opciones,
    key="analizador_marca"
)
if marca:
    filtros["marca"] = marca

modelo = st.selectbox(
    label="Modelo",
    options=[""] + modelo_opciones,
    key="analizador_modelo"
)
if modelo:
    filtros["modelo"] = modelo

num_serie = st.text_input(
    label="Número de serie",
    value=st.session_state.get("analizador_num_serie", ""),
    placeholder="Buscar por número de serie"
)
if num_serie:
    filtros["serie"] = num_serie.strip()

analizadores_filtrados = filtrar_analizadores(
    analizadores_df,
    tipo=filtros.get("tipo"),
    marca=filtros.get("marca"),
    modelo=filtros.get("modelo"),
    serie=filtros.get("serie")
)

st.markdown("### Resultados")
if analizadores_filtrados is None or analizadores_filtrados.empty:
    st.info("No hay analizadores que coincidan con los filtros seleccionados.")
else:
    st.dataframe(analizadores_filtrados, use_container_width=True)

if st.button("Seleccionar este analizador", type="primary", use_container_width=True):
    if analizador_tipo == "":
        st.error("Selecciona primero un tipo de analizador.")
    else:
        st.session_state.analizador_seleccionado = analizadores_filtrados.iloc[0].to_dict() if not analizadores_filtrados.empty else None
        if st.session_state.analizador_seleccionado:
            st.success("Analizador guardado en la sesión. Ahora puedes generar las listas de verificación con este analizador seleccionado.")
        else:
            st.error("No se encontró un analizador para seleccionar.")

if st.session_state.get("analizador_seleccionado"):
    st.divider()
    st.markdown("### Analizador seleccionado")
    for clave, valor in st.session_state.analizador_seleccionado.items():
        st.markdown(f"- **{clave}**: {valor}")
