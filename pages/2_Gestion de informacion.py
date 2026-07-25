# pages/2_Gestion de informacion.py
import streamlit as st
from utils.lector_inventario import (
    cargar_inventario,
    parsear_programacion_tinc,
    aplicar_programacion_tinc,
)


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
    st.session_state.usar_programacion_tinc             = "No"


def _guardar_informacion_usuario():
    st.session_state.nombre_ingeniero = st.session_state.get("input_nombre_ingeniero", "").strip()
    st.session_state.nombre_jefe      = st.session_state.get("input_nombre_jefe",      "").strip()
    st.session_state.nombre_hospital  = st.session_state.get("input_nombre_hospital",  "").strip()


# ── Inicializar estados ──────────────────────────────────────────────────────
for key, default in {
    "programacion_tinc_texto": "",
    "programacion_tinc_df":    None,
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
                st.session_state.inventario_df = aplicar_programacion_tinc(
                    st.session_state.inventario_df, df_prog
                )
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