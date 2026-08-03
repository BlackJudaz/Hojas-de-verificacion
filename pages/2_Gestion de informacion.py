# pages/2_Gestion de informacion.py
import pandas as pd
import streamlit as st
from utils.lector_inventario import (
    cargar_inventario,
    parsear_programacion_tinc,
    aplicar_programacion_tinc,
    normalizar_texto,
)

st.session_state["_pagina_actual"] = "page_gestion_informacion"


def _preseleccionar_equipos_programados(df_inventario):
    """Aplica y guarda la selección/filtros automáticos derivados de TiNC."""
    if df_inventario is None or "# ACTIVO" not in df_inventario.columns:
        st.session_state.tinc_filtros_automaticos = {}
        return 0

    if "FOLIO TINC" in df_inventario.columns:
        mask_programados = df_inventario["FOLIO TINC"].astype(str).str.strip().ne("")
        df_programados = df_inventario.loc[mask_programados].copy()
    else:
        df_programados = df_inventario.iloc[0:0].copy()

    activos_programados = (
        df_programados["# ACTIVO"]
        .astype(str)
        .str.strip()
        .replace("", pd.NA)
        .dropna()
        .drop_duplicates()
        .tolist()
        if not df_programados.empty else []
    )
    marcas_programadas = (
        sorted(df_programados["MARCA"].dropna().astype(str).unique().tolist())
        if "MARCA" in df_programados.columns else []
    )
    ubicaciones_programadas = (
        sorted(df_programados["UBICACIÓN"].dropna().astype(str).unique().tolist())
        if "UBICACIÓN" in df_programados.columns else []
    )
    conceptos_programados = (
        sorted(df_programados["CONCEPTO"].dropna().astype(str).unique().tolist())
        if "CONCEPTO" in df_programados.columns else []
    )

    conceptos_display = []
    for concepto in conceptos_programados:
        normalizado = normalizar_texto(concepto)
        if normalizado and normalizado not in conceptos_display:
            conceptos_display.append(normalizado)

    filtros_auto = {
        "ids": list(activos_programados),
        "marcas": list(marcas_programadas),
        "ubicaciones": list(ubicaciones_programadas),
        "conceptos": list(conceptos_programados),
        "conceptos_display": list(conceptos_display),
    }
    st.session_state.tinc_filtros_automaticos = filtros_auto

    st.session_state.ids_equipos_seleccionados = list(activos_programados)
    st.session_state.filtro_activo = list(activos_programados)
    st.session_state.filtro_marca = list(marcas_programadas)
    st.session_state.filtro_ubicacion = list(ubicaciones_programadas)
    st.session_state.filtro_concepto = list(conceptos_programados)
    st.session_state.filtro_tipo_activo_display = list(conceptos_display)
    st.session_state.ui_filtro_activo = list(activos_programados)
    st.session_state.ui_filtro_marca = list(marcas_programadas)
    st.session_state.ui_filtro_ubicacion = list(ubicaciones_programadas)
    st.session_state.ui_filtro_tipo_activo_display = list(conceptos_display)
    st.session_state.clic_buscar = bool(activos_programados)
    st.session_state._selector_autoseleccion_pendiente = bool(activos_programados)

    st.session_state.pop("_selector_cache_clave", None)
    st.session_state.pop("_selector_df_base", None)
    st.session_state["_selector_df_editor"] = None
    st.session_state["_selector_grid_key_rendered"] = ""
    st.session_state["_selector_ids_visibles"] = []
    st.session_state["_selector_reset_token"] = st.session_state.get("_selector_reset_token", 0) + 1

    return len(activos_programados)


def _resetear_estado_inventario():
    st.session_state.inventario_df                      = None
    st.session_state.clic_buscar                        = False
    st.session_state.filtro_concepto                    = []
    st.session_state.filtro_marca                       = []
    st.session_state.filtro_activo                      = []
    st.session_state.filtro_ubicacion                   = []
    st.session_state.analizadores_seleccionados         = []
    st.session_state.analizadores_por_concepto          = {}
    st.session_state.analizadores_bel_por_concepto      = {}
    st.session_state.analizadores_propios_por_concepto  = {}
    st.session_state.periodicidad_por_concepto          = {}
    st.session_state.tiempo_mantenimiento_por_concepto  = {}
    st.session_state.fecha_mantenimiento_por_concepto   = {}
    st.session_state.programacion_tinc_texto            = ""
    st.session_state.programacion_tinc_df               = None
    st.session_state.tinc_filtros_automaticos           = {}
    st.session_state._selector_autoseleccion_pendiente  = False
    st.session_state.usar_programacion_tinc             = "No"


def _guardar_informacion_usuario():
    st.session_state.nombre_ingeniero = st.session_state.get("input_nombre_ingeniero", "").strip()
    st.session_state.nombre_jefe      = st.session_state.get("input_nombre_jefe",      "").strip()
    st.session_state.nombre_hospital  = st.session_state.get("input_nombre_hospital",  "").strip()


# ── Inicializar estados ──────────────────────────────────────────────────────
for key, default in {
    "programacion_tinc_texto": "",
    "programacion_tinc_df":    None,
    "tinc_filtros_automaticos": {},
    "_selector_autoseleccion_pendiente": False,
    "usar_programacion_tinc":  "No",
    "input_nombre_ingeniero":  st.session_state.get("nombre_ingeniero", ""),
    "input_nombre_jefe":       st.session_state.get("nombre_jefe",      ""),
    "input_nombre_hospital":   st.session_state.get("nombre_hospital",  ""),
}.items():
    if key not in st.session_state:
        st.session_state[key] = default

# ── Título ───────────────────────────────────────────────────────────────────
st.title("Gestión de Información")
st.divider()

# ════════════════════════════════════════════════════════════════════════════
# PASO 1 — Inventario
# ════════════════════════════════════════════════════════════════════════════
with st.container(border=True):
    st.markdown("### 1. Sube el inventario")
    st.caption("El inventario será utilizado para la generación tanto de las hojas de verificación como de las etiquetas.")

    if st.session_state.inventario_df is None:
        archivo = st.file_uploader(
            label="Sube tu archivo Excel de inventario",
            type=["xlsx"],
            label_visibility="collapsed",
        )
        if archivo is not None:
            resultado = cargar_inventario(archivo)
            df_temporal, error = None, None
            if isinstance(resultado, tuple):
                if len(resultado) == 3:
                    df_temporal, error, _ = resultado
                elif len(resultado) == 2:
                    df_temporal, error = resultado
            else:
                error = "Error al procesar el inventario: respuesta no válida."

            if df_temporal is not None:
                st.session_state.inventario_df = df_temporal
                st.rerun()
            else:
                st.error(error or "No se pudo leer el archivo.")
    else:
        df = st.session_state.inventario_df

        st.dataframe(
            df.head(10)[[c for c in ["# ACTIVO", "CONCEPTO", "MARCA", "MODELO", "UBICACIÓN", "SUB UBICACIÓN"] if c in df.columns]],
            use_container_width=True,
            hide_index=True,
        )

        st.success(f"Se registraron {len(df)} equipos en el inventario")

        if st.button("Eliminar e importar otro inventario", use_container_width=False):
            _resetear_estado_inventario()
            st.rerun()

# ════════════════════════════════════════════════════════════════════════════
# PASO 2 — Información del usuario
# ════════════════════════════════════════════════════════════════════════════
with st.container(border=True):
    st.markdown("### 2. Registra la Información del Usuario")
    st.caption("La información será utilizada para la generación tanto de las hojas de verificación como de las etiquetas.")

    col_izq, col_der = st.columns(2)
    with col_izq:
        st.text_input(
            label="Nombre del Ingeniero de Servicio",
            key="input_nombre_ingeniero",
            placeholder="Jose Andres Quijada",
            on_change=_guardar_informacion_usuario,
        )
        st.text_input(
            label="Nombre del Jefe de Departamento",
            key="input_nombre_jefe",
            placeholder="Dr. Juan",
            on_change=_guardar_informacion_usuario,
        )
    with col_der:
        st.text_input(
            label="Nombre del Hospital",
            key="input_nombre_hospital",
            placeholder="Medyarthros",
            on_change=_guardar_informacion_usuario,
        )
        st.caption("Campos obligatorios: nombre del ingeniero y nombre del hospital.")

    # Resumen de información guardada
    nombre_ing  = st.session_state.get("nombre_ingeniero", "")
    nombre_jefe = st.session_state.get("nombre_jefe",      "")
    nombre_hosp = st.session_state.get("nombre_hospital",  "")

    if nombre_ing or nombre_jefe or nombre_hosp:
        st.markdown("#### Información guardada")
        if nombre_ing:
            st.markdown(f"👤 **Ingeniero:** {nombre_ing}")
        if nombre_jefe:
            st.markdown(f"👔 **Jefe:** {nombre_jefe}")
        if nombre_hosp:
            st.markdown(f"🏥 **Hospital:** {nombre_hosp}")

        if nombre_ing and nombre_hosp:
            st.success("Información Guardada")
        else:
            st.warning("⚠️ Faltan campos obligatorios: nombre del ingeniero y nombre del hospital.")

# ════════════════════════════════════════════════════════════════════════════
# PASO 3 — TiNC
# ════════════════════════════════════════════════════════════════════════════
with st.container(border=True):
    st.markdown("### 3. Agregar Inventario de TiNC")
    st.caption("Con el inventario de TiNC las hojas de Verificación relacionadas al equipo tendrán el Folio de TiNC.")

    st.radio(
        "¿Deseas agregar la programación de TiNC?",
        options=["No", "Sí"],
        key="usar_programacion_tinc",
        horizontal=True,
    )

    if st.session_state.usar_programacion_tinc == "Sí":
        texto_programacion = st.text_area(
            label="Programación TiNC",
            key="programacion_tinc_texto",
            height=200,
            placeholder="Pega aquí manualmente el contenido de la programación del mes.",
            label_visibility="collapsed",
        )

        if st.button("Guardar programación TINC", type="primary", use_container_width=True):
            if st.session_state.inventario_df is not None:
                df_prog = parsear_programacion_tinc(texto_programacion)
                st.session_state.programacion_tinc_df  = df_prog
                inventario_actualizado = aplicar_programacion_tinc(
                    st.session_state.inventario_df, df_prog
                )
                st.session_state.inventario_df = inventario_actualizado
                _preseleccionar_equipos_programados(inventario_actualizado)
                st.rerun()
            else:
                st.warning("Primero sube el inventario en el paso 1.")

    # Tabla de emparejamiento
    if st.session_state.inventario_df is not None:
        df_actual = st.session_state.inventario_df

        if "FOLIO TINC" in df_actual.columns:
            df_emparejados = df_actual[
                df_actual["FOLIO TINC"].astype(str).str.strip().ne("")
            ]

            columnas_tabla = [c for c in [
                "FOLIO TINC", "# ACTIVO", "CONCEPTO", "MARCA",
                "MODELO", "No. DE SERIE", "UBICACIÓN",
            ] if c in df_emparejados.columns]

            if not df_emparejados.empty:
                st.markdown("#### Tabla de verificación del emparejamiento")
                st.caption("Aquí se muestran únicamente los equipos que ya quedaron emparejados con un folio TiNC.")

                st.dataframe(
                    df_emparejados[columnas_tabla].head(50),
                    use_container_width=True,
                    hide_index=True,
                )

                st.success(f"Se emparejaron {len(df_emparejados)} equipos")

                if len(df_emparejados) > 50:
                    st.caption(f"Se muestran los primeros 50 de {len(df_emparejados)} emparejados.")