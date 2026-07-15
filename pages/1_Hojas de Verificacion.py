import re
import streamlit as st
from datetime import datetime

from utils.gestor_plantillas import crear_paquete_reporte
from utils.lector_analizadores import (
    cargar_analizadores,
    buscar_analizadores_por_concepto,
    obtener_analizadores_display,
    parse_analizador_display
)
from utils.lector_inventario import aplicar_filtros, opciones_disponibles, df_con_filtros

config = {
    "nombre":   st.session_state.get("nombre_ingeniero", ""),
    "jefe":     st.session_state.get("nombre_jefe", ""),
    "hospital": st.session_state.get("nombre_hospital", "")
}


def inicializar_estado():
    if "inventario_df" not in st.session_state:
        st.session_state.inventario_df = None
    if "clic_buscar" not in st.session_state:
        st.session_state.clic_buscar = False
    if "analizadores_seleccionados" not in st.session_state:
        st.session_state.analizadores_seleccionados = []
    if "periodicidad_por_concepto" not in st.session_state:
        st.session_state.periodicidad_por_concepto = {}
    for key in ("filtro_concepto", "filtro_marca", "filtro_activo", "filtro_ubicacion"):
        if key not in st.session_state:
            st.session_state[key] = []


def limpiar_filtros():
    st.session_state.clic_buscar = False
    st.session_state.filtro_concepto = []
    st.session_state.filtro_marca = []
    st.session_state.filtro_activo = []
    st.session_state.filtro_ubicacion = []


# ── Inicio de la página ──────────────────────────────────────────────────────
st.title("Genera Hojas de Verificación")

inicializar_estado()

if st.session_state.inventario_df is None:
    st.warning("No se ha detectado ningún inventario en el sistema.")
    st.info("Por favor, ve a la sección 'Inventario' en el menú lateral antes de continuar.")
    st.stop()

analizadores_df = cargar_analizadores()
if analizadores_df is None:
    st.warning("No se pudo cargar la lista de analizadores.")

df = st.session_state.inventario_df

# ── Filtros ──────────────────────────────────────────────────────────────────
with st.container(border=True):
    col_filtros, col_acciones = st.columns([8, 2])
    with col_filtros:
        estados = {
            "filtro_concepto":  st.session_state.filtro_concepto,
            "filtro_marca":     st.session_state.filtro_marca,
            "filtro_activo":    st.session_state.filtro_activo,
            "filtro_ubicacion": st.session_state.filtro_ubicacion
        }
        opciones = {
            "CONCEPTO":  opciones_disponibles(df_con_filtros(df, estados, excluir="filtro_concepto"), "CONCEPTO"),
            "MARCA":     opciones_disponibles(df_con_filtros(df, estados, excluir="filtro_marca"), "MARCA"),
            "# ACTIVO":  opciones_disponibles(df_con_filtros(df, estados, excluir="filtro_activo"), "# ACTIVO"),
            "UBICACIÓN": opciones_disponibles(df_con_filtros(df, estados, excluir="filtro_ubicacion"), "UBICACIÓN")
        }

        fila1_col1, fila1_col2 = st.columns(2)
        with fila1_col1:
            st.multiselect(label="Concepto", options=opciones["CONCEPTO"],
                           key="filtro_concepto", placeholder="Concepto",
                           label_visibility="collapsed")
        with fila1_col2:
            st.multiselect(label="Marca", options=opciones["MARCA"],
                           key="filtro_marca", placeholder="Marca",
                           label_visibility="collapsed")

        fila2_col1, fila2_col2 = st.columns(2)
        with fila2_col1:
            st.multiselect(label="Activo", options=opciones["# ACTIVO"],
                           key="filtro_activo", placeholder="Activo",
                           label_visibility="collapsed")
        with fila2_col2:
            st.multiselect(label="Ubicación", options=opciones["UBICACIÓN"],
                           key="filtro_ubicacion", placeholder="Ubicacion",
                           label_visibility="collapsed")

    with col_acciones:
        if st.button("Buscar", use_container_width=True, type="primary"):
            st.session_state.clic_buscar = True
        st.button("Limpiar", use_container_width=True, on_click=limpiar_filtros)

# ── Resultados ───────────────────────────────────────────────────────────────
if st.session_state.clic_buscar:
    filtros = {
        "CONCEPTO":  st.session_state.filtro_concepto,
        "MARCA":     st.session_state.filtro_marca,
        "# ACTIVO":  st.session_state.filtro_activo,
        "UBICACIÓN": st.session_state.filtro_ubicacion
    }
    df_final = aplicar_filtros(df, filtros)

    columnas_mostrar = [
        c for c in ["# ACTIVO", "CONCEPTO", "MARCA", "MODELO", "UBICACIÓN", "SUB UBICACIÓN"]
        if c in df.columns
    ]

    with st.container(border=True):
        seleccion_tabla = st.dataframe(
            df_final[columnas_mostrar],
            use_container_width=True,
            hide_index=True,
            on_select="rerun",
            selection_mode="multi-row"
        )

    filas_seleccionadas = seleccion_tabla.get("selection", {}).get("rows", [])

    if filas_seleccionadas:
        equipos_a_mantener = df_final.iloc[filas_seleccionadas].copy()
        st.success(f"Se seleccionaron {len(equipos_a_mantener)} equipos")

        conceptos_seleccionados = equipos_a_mantener["CONCEPTO"].dropna().unique().tolist()

        # ── Periodicidad ─────────────────────────────────────────────────────
        st.markdown("### Periodicidad de mantenimiento por tipo de activo")
        opciones_periodicidad = ["Bimestral", "Cuatrimestral", "Semestral", "Anual"]
        periodicidades_por_concepto = st.session_state.periodicidad_por_concepto.copy()

        for concepto in conceptos_seleccionados:
            clave_limpia = re.sub(r'\W+', '_', concepto.strip().lower()).strip('_')
            clave = f"periodicidad_{clave_limpia}"            
            default = periodicidades_por_concepto.get(concepto, "Anual")
            if default not in opciones_periodicidad:
                default = "Anual"
            seleccion = st.selectbox(
                label=f"Periodicidad para {concepto}",
                options=opciones_periodicidad,
                index=opciones_periodicidad.index(default),
                key=clave
            )
            periodicidades_por_concepto[concepto] = seleccion

        st.session_state.periodicidad_por_concepto = periodicidades_por_concepto
        equipos_a_mantener["PERIODICIDAD"] = equipos_a_mantener["CONCEPTO"].map(
            periodicidades_por_concepto).fillna("Anual")

        # ── Analizadores ──────────────────────────────────────────────────────
        analizador_seleccionado_por_concepto = []
        analizadores_por_concepto = {}

        if analizadores_df is not None:
            st.markdown("### Selección de analizadores por tipo de equipo")
            for concepto in conceptos_seleccionados:
                sugeridos = buscar_analizadores_por_concepto(analizadores_df, [concepto])
                opciones_sugeridas = obtener_analizadores_display(sugeridos)
                clave_concepto = re.sub(r"\W+", "_", concepto.strip().lower()).strip("_")
                key = f"analizadores_{clave_concepto}"
                if key not in st.session_state:
                    st.session_state[key] = []

                if opciones_sugeridas:
                    st.markdown(f"**{concepto}**")
                    seleccion = st.multiselect(
                        label=f"Analizadores para {concepto}",
                        options=opciones_sugeridas,
                        default=st.session_state.get(key, []),
                        key=key
                    )
                    analizador_seleccionado_por_concepto.extend(seleccion)
                    analizadores_por_concepto[concepto] = [
                        parse_analizador_display(item) for item in seleccion
                    ]
                else:
                    st.warning(f"No se encontraron analizadores recomendados para '{concepto}'.")
                    analizadores_por_concepto[concepto] = [
                        parse_analizador_display(item)
                        for item in st.session_state.get(key, [])
                    ]

        st.session_state["analizadores_seleccionados"] = list(
            dict.fromkeys(analizador_seleccionado_por_concepto))
        st.session_state["analizadores_por_concepto"] = analizadores_por_concepto

        # ── Nombre del paquete ────────────────────────────────────────────────
        nombre_carpeta = st.text_input(
            "Nombre del paquete",
            value=f"reporte_{datetime.now():%Y%m%d_%H%M%S}",
            help="Nombre del archivo ZIP que se descargará."
        ).strip()
        if not nombre_carpeta:
            nombre_carpeta = f"reporte_{datetime.now():%Y%m%d_%H%M%S}"
        nombre_carpeta = re.sub(r"\W+", "_", nombre_carpeta).strip("_")

        # ── Generar ───────────────────────────────────────────────────────────
        col_boton, col_hojas, col_etiquetas = st.columns([4, 3, 3])
        with col_boton:
            generar = st.button("Generar archivo", type="primary", use_container_width=True)
        with col_hojas:
            hacer_hojas = st.checkbox("Hojas de verificación", value=True)
        with col_etiquetas:
            hacer_etiquetas = st.checkbox("Etiquetas", value=True)

        if generar:
            if not hacer_hojas and not hacer_etiquetas:
                st.error("❌ Selecciona al menos una opción.")
            else:
                progress_bar = st.progress(0)
                status_text  = st.empty()

                buffer_zip, errores, exitos = crear_paquete_reporte(
                    equipos                    = equipos_a_mantener,
                    nombre_carpeta             = nombre_carpeta,
                    ingeniero                  = config.get("nombre", ""),
                    jefe                       = config.get("jefe", ""),
                    hospital                   = config.get("hospital", ""),
                    progress_bar               = progress_bar,
                    status_text                = status_text,
                    hacer_hojas                = hacer_hojas,
                    hacer_etiquetas            = hacer_etiquetas,
                    analizadores_por_concepto  = st.session_state.get("analizadores_por_concepto", {}),
                    analizadores_seleccionados = st.session_state.get("analizadores_seleccionados", [])
                )

                if exitos > 0 or hacer_etiquetas:
                    st.download_button(
                        label               = "⬇️ Descargar Paquete (.zip)",
                        data                = buffer_zip,
                        file_name           = f"{nombre_carpeta}.zip",
                        mime                = "application/zip",
                        use_container_width = True
                    )

                if errores:
                    with st.expander("Detalles de advertencias u omisiones"):
                        for err in errores:
                            st.warning(err)
    else:
        st.info("Seleccione los equipos que desea trabajar")