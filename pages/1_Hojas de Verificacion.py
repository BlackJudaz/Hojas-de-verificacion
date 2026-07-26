# pages/Hojas de Verificacion.py
import re
import streamlit as st
import pandas as pd
from datetime import datetime, date

from utils.gestor_plantillas import crear_paquete_reporte
from utils.google_drive import (
    formatear_nombre_carpeta_documentacion,
    resolver_fecha_referencia_drive,
)
from utils.lector_analizadores import (
    cargar_analizadores,
    buscar_analizadores_por_concepto,
    obtener_analizadores_display,
    parse_analizador_display,
)
from utils.lector_inventario import (
    aplicar_filtros,
    opciones_disponibles,
    df_con_filtros,
    normalizar_texto,
)

config = {
    "nombre":   st.session_state.get("nombre_ingeniero", ""),
    "jefe":     st.session_state.get("nombre_jefe", ""),
    "hospital": st.session_state.get("nombre_hospital", ""),
}

MAX_ANALIZADORES = 3
COLUMNA_ID = "# ACTIVO"


@st.cache_data(show_spinner=False)
def _cargar_analizadores_cached():
    return cargar_analizadores()


def inicializar_estado():
    defaults = {
        "inventario_df":                    None,
        "clic_buscar":                      False,
        "ids_equipos_seleccionados":        [],
        "analizadores_seleccionados":       [],
        "periodicidad_por_concepto":        {},
        "tiempo_mantenimiento_por_concepto":{},
        "analizadores_bel_por_concepto":    {},
        "analizadores_propios_por_concepto":{},
        "fecha_mantenimiento_por_concepto": {},
        "ultimo_paquete_zip_bytes":         b"",
        "ultimo_paquete_zip_nombre":        "",
        "ultimo_paquete_drive_folder":      "",
        "ultimo_paquete_periodo":           "",
        "ultimo_paquete_periodo_mixto":     False,
        "ultimo_paquete_generado_en":       "",
        "filtro_concepto":                  [],
        "filtro_tipo_activo_display":       [],
        "filtro_marca":                     [],
        "filtro_activo":                    [],
        "filtro_ubicacion":                 [],
        "_selector_reset_token":            0,
        "_selector_sincronizar_visual":     False,
        "_selector_limpiar_visual":         False,
        "_selector_ids_visibles":           [],
        "_selector_df_editor":             None,
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val


def limpiar_filtros():
    st.session_state.clic_buscar             = False
    st.session_state.filtro_concepto         = []
    st.session_state.filtro_tipo_activo_display = []
    st.session_state.filtro_marca            = []
    st.session_state.filtro_activo           = []
    st.session_state.filtro_ubicacion        = []
    st.session_state.pop("_selector_cache_clave", None)
    st.session_state.pop("_selector_df_base", None)
    st.session_state["_selector_df_editor"] = None
    st.session_state["_selector_sincronizar_visual"] = False
    st.session_state["_selector_limpiar_visual"] = False
    st.session_state["_selector_ids_visibles"] = []


def _clave(concepto):
    return re.sub(r"\W+", "_", concepto.strip().lower()).strip("_")


def _total_analizadores(concepto):
    bel     = [x for x in st.session_state.analizadores_bel_por_concepto.get(concepto, []) if x]
    propios = st.session_state.analizadores_propios_por_concepto.get(concepto, [])
    return len(bel) + len(propios)


def _toggle_seleccionar_todo_visible(toggle_key):
    df_editor = st.session_state.get("_selector_df_editor")
    if df_editor is None or "SELECCIONAR" not in df_editor.columns:
        return

    marcar_todo = bool(st.session_state.get(toggle_key, False))
    df_editor = df_editor.copy()
    df_editor["SELECCIONAR"] = marcar_todo
    st.session_state["_selector_df_editor"] = df_editor

    ids_en_tabla = set(df_editor[COLUMNA_ID].astype(str))
    ids_previos = set(st.session_state.ids_equipos_seleccionados)
    ids_seleccionados_tabla = set(
        df_editor.loc[df_editor["SELECCIONAR"], COLUMNA_ID].astype(str)
    )
    st.session_state.ids_equipos_seleccionados = sorted(
        (ids_previos - ids_en_tabla) | ids_seleccionados_tabla
    )

    # Fuerza remontaje del editor para que el estado visual de casillas se actualice al instante.
    st.session_state["_selector_reset_token"] = st.session_state.get("_selector_reset_token", 0) + 1
    st.rerun()


# ── Inicio de la página ──────────────────────────────────────────────────────
st.title("Generador de Hojas de Verificación")
st.caption("Filtra los equipos, selecciona los activos a trabajar y genera el paquete de hojas y etiquetas en un solo flujo.")

inicializar_estado()

if st.session_state.inventario_df is None:
    st.warning("No se ha detectado ningún inventario en el sistema.")
    st.info("Por favor, ve a la sección 'Inventario' en el menú lateral antes de continuar.")
    st.stop()

analizadores_df = _cargar_analizadores_cached()
df = st.session_state.inventario_df

# ════════════════════════════════════════════════════════════════════════════
# PASO 1 — Filtrar
# ════════════════════════════════════════════════════════════════════════════
main_card = st.container(border=True)
with main_card:
    st.markdown("### 1. Filtra y selecciona los equipos")
    st.caption("Primero filtra por tipo de activo, marca, número de activo o ubicación; luego selecciona en la misma vista.")

    col_filtros, col_acciones = st.columns([8, 2])
    with col_filtros:
        estados = {
            "filtro_concepto":  st.session_state.filtro_concepto,
            "filtro_marca":     st.session_state.filtro_marca,
            "filtro_activo":    st.session_state.filtro_activo,
            "filtro_ubicacion": st.session_state.filtro_ubicacion,
        }
        opciones_calc = {
            "CONCEPTO":  opciones_disponibles(df_con_filtros(df, estados, excluir="filtro_concepto"), "CONCEPTO"),
            "MARCA":     opciones_disponibles(df_con_filtros(df, estados, excluir="filtro_marca"),    "MARCA"),
            "# ACTIVO":  opciones_disponibles(df_con_filtros(df, estados, excluir="filtro_activo"),   "# ACTIVO"),
            "UBICACIÓN": opciones_disponibles(df_con_filtros(df, estados, excluir="filtro_ubicacion"),"UBICACIÓN"),
        }

        mapa_display = {}
        opciones_display = []
        for c_orig in opciones_calc["CONCEPTO"]:
            clave_d = normalizar_texto(c_orig)
            if not clave_d:
                continue
            if clave_d not in mapa_display:
                mapa_display[clave_d] = []
                opciones_display.append(clave_d)
            mapa_display[clave_d].append(c_orig)

        sel_display = [v for v in st.session_state.filtro_tipo_activo_display if v in mapa_display]
        if sel_display != st.session_state.filtro_tipo_activo_display:
            st.session_state.filtro_tipo_activo_display = sel_display

        f1c1, f1c2 = st.columns(2)
        with f1c1:
            st.multiselect(
                label="Tipo de activo",
                options=opciones_display,
                key="filtro_tipo_activo_display",
                placeholder="Selecciona uno o varios tipos de activo",
                label_visibility="collapsed",
                format_func=lambda x: str(x).title(),
            )
            conceptos_filtrados = []
            for d in st.session_state.filtro_tipo_activo_display:
                conceptos_filtrados.extend(mapa_display.get(d, []))
            st.session_state.filtro_concepto = list(dict.fromkeys(conceptos_filtrados))

        with f1c2:
            st.multiselect(label="Marca", options=opciones_calc["MARCA"],
                           key="filtro_marca", placeholder="Selecciona una o varias marcas",
                           label_visibility="collapsed")

        f2c1, f2c2 = st.columns(2)
        with f2c1:
            st.multiselect(label="Activo", options=opciones_calc["# ACTIVO"],
                           key="filtro_activo", placeholder="Selecciona uno o varios activos",
                           label_visibility="collapsed")
        with f2c2:
            st.multiselect(label="Ubicación", options=opciones_calc["UBICACIÓN"],
                           key="filtro_ubicacion", placeholder="Selecciona una o varias ubicaciones",
                           label_visibility="collapsed")

    with col_acciones:
        if st.button("Buscar", use_container_width=True, type="primary"):
            st.session_state.clic_buscar = True
        st.button("Limpiar", use_container_width=True, on_click=limpiar_filtros)

# ════════════════════════════════════════════════════════════════════════════
# PASO 2 — Seleccionar
# ════════════════════════════════════════════════════════════════════════════
if not st.session_state.clic_buscar:
    st.stop()

filtros = {
    "CONCEPTO":  st.session_state.filtro_concepto,
    "MARCA":     st.session_state.filtro_marca,
    "# ACTIVO":  st.session_state.filtro_activo,
    "UBICACIÓN": st.session_state.filtro_ubicacion,
}
clave_filtros = "|".join([
    ",".join(sorted(st.session_state.filtro_concepto)),
    ",".join(sorted(st.session_state.filtro_marca)),
    ",".join(sorted(st.session_state.filtro_activo)),
    ",".join(sorted(st.session_state.filtro_ubicacion)),
])

columnas_mostrar = [
    c for c in ["# ACTIVO", "CONCEPTO", "MARCA", "MODELO", "UBICACIÓN", "SUB UBICACIÓN"]
    if c in df.columns
]


if st.session_state.get("_selector_cache_clave") != clave_filtros:
    ids_previos_cache = set(st.session_state.ids_equipos_seleccionados)
    ids_visibles_previos = set(st.session_state.get("_selector_ids_visibles", []))
    # Solo arrastrar seleccionados que estaban en la tabla visible previa.
    ids_arrastrables = ids_previos_cache & ids_visibles_previos if ids_visibles_previos else ids_previos_cache
    df_filtrado = aplicar_filtros(df, filtros)

    # Mantiene visibles seleccionados y los coloca al final para no romper el flujo al ampliar filtros.
    mask_sel_filtrado = df_filtrado[COLUMNA_ID].astype(str).isin(ids_arrastrables)
    df_filtrado_no_sel = df_filtrado[~mask_sel_filtrado]
    df_sel = df[df[COLUMNA_ID].astype(str).isin(ids_arrastrables)]
    df_mostrar = pd.concat([df_filtrado_no_sel, df_sel], ignore_index=True)
    df_mostrar = df_mostrar.drop_duplicates(subset=[COLUMNA_ID], keep="first")
    df_editor = df_mostrar[columnas_mostrar].copy()
    df_editor.insert(
        0,
        "SELECCIONAR",
        df_editor[COLUMNA_ID].astype(str).isin(ids_previos_cache),
    )

    st.session_state["_selector_cache_clave"] = clave_filtros
    st.session_state["_selector_df_base"] = df_mostrar[columnas_mostrar].copy()
    st.session_state["_selector_df_editor"] = df_editor


def _render_selector_equipos(columnas):
    with st.container(border=False):
    #    st.markdown("#### Selecciona los equipos")
     #   st.caption(
      #      "Marca los equipos a trabajar usando las casillas de la tabla."
       # )
       # st.caption(
       #     "Si cambias filtros, los equipos ya seleccionados que no coincidan se mostraran al final y conservaran su seleccion."
       # )

        ids_previos = set(st.session_state.ids_equipos_seleccionados)
        df_base_tabla = st.session_state.get("_selector_df_base")
        if df_base_tabla is None:
            st.info("Aplica los filtros para cargar la tabla de equipos.")
            st.stop()

        clave_df = f"selector_df_equipos_{st.session_state.get('_selector_reset_token', 0)}"
        df_tabla = st.session_state.get("_selector_df_editor")
        if df_tabla is None:
            df_tabla = df_base_tabla.copy()
            df_tabla.insert(
                0,
                "SELECCIONAR",
                df_tabla[COLUMNA_ID].astype(str).isin(ids_previos),
            )
            st.session_state["_selector_df_editor"] = df_tabla.copy()

        toggle_all_key = f"selector_toggle_all_{st.session_state.get('_selector_reset_token', 0)}"
        st.session_state[toggle_all_key] = bool(df_tabla["SELECCIONAR"].all())
        st.checkbox(
            "Seleccionar todo lo visible",
            key=toggle_all_key,
            help="Marca todo lo visible; si lo desmarcas, se deselecciona todo lo visible.",
            on_change=_toggle_seleccionar_todo_visible,
            args=(toggle_all_key,),
        )

        with st.form(key=f"selector_form_{st.session_state.get('_selector_reset_token', 0)}"):
            tabla_editada = st.data_editor(
                df_tabla,
                use_container_width=True,
                hide_index=True,
                key=clave_df,
                column_config={
                    "SELECCIONAR": st.column_config.CheckboxColumn(
                        "Seleccionar",
                        help="Marca los equipos a trabajar",
                    )
                },
                disabled=[col for col in df_tabla.columns if col != "SELECCIONAR"],
            )
            aplicar_seleccion = st.form_submit_button(
                "Aplicar selección",
                type="primary",
                use_container_width=True,
            )

        ids_en_tabla = set(df_base_tabla[COLUMNA_ID].astype(str))
        st.session_state["_selector_ids_visibles"] = sorted(ids_en_tabla)
        if aplicar_seleccion:
            st.session_state["_selector_df_editor"] = tabla_editada.copy()
            ids_seleccionados_tabla = set(
                tabla_editada.loc[tabla_editada["SELECCIONAR"], COLUMNA_ID].astype(str)
            )
            ids_finales = (ids_previos - ids_en_tabla) | ids_seleccionados_tabla
            st.session_state.ids_equipos_seleccionados = sorted(ids_finales)
            st.rerun()

        col_info, col_borrar = st.columns([10, 2])
        with col_borrar:
            if st.button("Borrar selección", use_container_width=True):
                st.session_state.ids_equipos_seleccionados = []
                st.session_state["_selector_reset_token"] = st.session_state.get("_selector_reset_token", 0) + 1
                st.session_state["_selector_limpiar_visual"] = False
                st.session_state["_selector_sincronizar_visual"] = False
                # Fuerza reconstruccion de la tabla con los filtros actuales,
                # sin filas arrastradas por seleccion previa.
                st.session_state.pop("_selector_cache_clave", None)
                st.session_state.pop("_selector_df_base", None)
                st.session_state["_selector_df_editor"] = None
                st.rerun()

        with col_info:
            total = len(st.session_state.ids_equipos_seleccionados)
            if total:
                st.success(f"Se seleccionaron {total} equipos")
            else:
                st.info("Selecciona uno o varios equipos de la tabla para continuar.")


with main_card:
    _render_selector_equipos(columnas_mostrar)

ids_finales = set(st.session_state.ids_equipos_seleccionados)
if not ids_finales:
    st.stop()

equipos_a_mantener = df[df[COLUMNA_ID].astype(str).isin(ids_finales)].copy()
conceptos_seleccionados = equipos_a_mantener["CONCEPTO"].dropna().unique().tolist()

# ════════════════════════════════════════════════════════════════════════════
# PASO 3 — Configurar información
# ════════════════════════════════════════════════════════════════════════════
opciones_periodicidad    = ["Anual", "Semestral", "Cuatrimestral", "Bimestral"]
periodicidades_concepto  = st.session_state.periodicidad_por_concepto.copy()
tiempos_concepto         = st.session_state.tiempo_mantenimiento_por_concepto.copy()

with st.container(border=True):
    st.markdown("### 3. Configura la información")
    st.caption("Selecciona la información correspondiente al servicio realizado: Fecha, Periodicidad, Tiempo y Analizadores usados.")

    for concepto in conceptos_seleccionados:
        ck = _clave(concepto)

        # Inicializar listas de analizadores si no existen
        if concepto not in st.session_state.analizadores_bel_por_concepto:
            st.session_state.analizadores_bel_por_concepto[concepto] = []
        if concepto not in st.session_state.analizadores_propios_por_concepto:
            st.session_state.analizadores_propios_por_concepto[concepto] = []

        with st.container(border=True):
            st.markdown(f"#### {concepto}")

            # ── Fila superior: Fecha | Periodicidad | Tiempo ─────────────────
            col_fecha, col_period, col_tiempo = st.columns(3)

            with col_fecha:
                usar_fecha_manual = st.checkbox(
                    "Seleccionar fecha de manera manual",
                    key=f"usar_fecha_manual_{ck}",
                    value=concepto in st.session_state.fecha_mantenimiento_por_concepto,
                )
                if usar_fecha_manual:
                    fecha_guardada = st.session_state.fecha_mantenimiento_por_concepto.get(
                        concepto, date.today()
                    )
                    fecha_elegida = st.date_input(
                        label="Fecha",
                        value=fecha_guardada,
                        key=f"fecha_manual_{ck}",
                        label_visibility="collapsed",
                    )
                    st.session_state.fecha_mantenimiento_por_concepto[concepto] = fecha_elegida
                else:
                    st.session_state.fecha_mantenimiento_por_concepto.pop(concepto, None)
                    st.markdown(
                        f"<div style='padding:8px 0 4px 0; font-size:14px; color:#888;'>"
                        f"{date.today().strftime('%Y/%m/%d')}</div>",
                        unsafe_allow_html=True,
                    )

            with col_period:
                st.caption("Periodicidad entre cada mantenimiento")
                default_p = periodicidades_concepto.get(concepto, "Anual")
                if default_p not in opciones_periodicidad:
                    default_p = "Anual"
                periodicidad_elegida = st.selectbox(
                    label=f"Periodicidad_{ck}",
                    options=opciones_periodicidad,
                    index=opciones_periodicidad.index(default_p),
                    key=f"periodicidad_{ck}",
                    label_visibility="collapsed",
                )
                periodicidades_concepto[concepto] = periodicidad_elegida

            with col_tiempo:
                st.caption("Tiempo en realizar el mantenimiento")
                tiempo_actual = tiempos_concepto.get(concepto, "")
                tiempo_nuevo  = st.text_input(
                    label=f"Tiempo_{ck}",
                    value=tiempo_actual,
                    placeholder="Ej. 1 h, 2 h",
                    key=f"tiempo_{ck}",
                    label_visibility="collapsed",
                )
                tiempos_concepto[concepto] = str(tiempo_nuevo or "").strip()

            # ── Analizadores ─────────────────────────────────────────────────
            st.markdown("**Analizadores utilizados en el mantenimiento**")

            if analizadores_df is not None:
                opciones_bel_concepto = list(dict.fromkeys(
                    obtener_analizadores_display(
                        buscar_analizadores_por_concepto(analizadores_df, [concepto])
                    ) + obtener_analizadores_display(analizadores_df)
                ))
            else:
                opciones_bel_concepto = []

            # Lista actual de analizadores BEL
            lista_bel    = st.session_state.analizadores_bel_por_concepto[concepto]
            lista_propios = st.session_state.analizadores_propios_por_concepto[concepto]

            # Mostrar selectbox por cada analizador BEL ya agregado
            nuevos_bel = []
            for idx, val_actual in enumerate(lista_bel):
                col_sel, col_menos = st.columns([11, 1])
                with col_sel:
                    opciones_slot = [""] + [
                        o for o in opciones_bel_concepto
                        if o not in lista_bel or o == val_actual
                    ]
                    if val_actual not in opciones_slot:
                        val_actual = ""
                    nuevo_val = st.selectbox(
                        label=f"Analizador BEL {idx+1}",
                        options=opciones_slot,
                        index=opciones_slot.index(val_actual) if val_actual in opciones_slot else 0,
                        key=f"bel_{ck}_{idx}",
                        label_visibility="collapsed",
                        placeholder="Escoge un analizador...",
                    )
                    nuevos_bel.append(nuevo_val)
                with col_menos:
                    if st.button("−", key=f"quitar_bel_{ck}_{idx}", use_container_width=True):
                        lista_bel.pop(idx)
                        st.session_state.analizadores_bel_por_concepto[concepto] = lista_bel
                        st.rerun()

            st.session_state.analizadores_bel_por_concepto[concepto] = nuevos_bel

            # Mostrar filas de analizadores propios ya agregados
            nuevos_propios = []
            for idx, propio in enumerate(lista_propios):
                cp1, cp2, cp3, cp4, cp_menos = st.columns([3, 3, 3, 3, 1])
                with cp1:
                    tipo  = st.text_input("Analizador",    value=propio.get("tipo", ""),
                                          key=f"propio_tipo_{ck}_{idx}",  label_visibility="collapsed",
                                          placeholder="Analizador")
                with cp2:
                    marca = st.text_input("Marca",         value=propio.get("marca", ""),
                                          key=f"propio_marca_{ck}_{idx}", label_visibility="collapsed",
                                          placeholder="Marca")
                with cp3:
                    modelo= st.text_input("Modelo",        value=propio.get("modelo", ""),
                                          key=f"propio_modelo_{ck}_{idx}",label_visibility="collapsed",
                                          placeholder="Modelo")
                with cp4:
                    serie = st.text_input("Número de Serie",value=propio.get("serie", ""),
                                          key=f"propio_serie_{ck}_{idx}", label_visibility="collapsed",
                                          placeholder="Número de Serie")
                with cp_menos:
                    if st.button("−", key=f"quitar_propio_{ck}_{idx}", use_container_width=True):
                        lista_propios.pop(idx)
                        st.session_state.analizadores_propios_por_concepto[concepto] = lista_propios
                        st.rerun()

                nuevos_propios.append({
                    "tipo": tipo, "marca": marca, "modelo": modelo, "serie": serie
                })

            st.session_state.analizadores_propios_por_concepto[concepto] = nuevos_propios

            # Botones de agregar — solo si no se llegó al límite
            total_actual = _total_analizadores(concepto)
            if total_actual < MAX_ANALIZADORES:
                col_agregar_bel, col_agregar_propio, _ = st.columns([2, 2, 6])
                with col_agregar_bel:
                    if st.button(f"＋ Agregar analizador", key=f"add_bel_{ck}",
                                 use_container_width=True):
                        st.session_state.analizadores_bel_por_concepto[concepto].append("")
                        st.rerun()
                with col_agregar_propio:
                    if st.button(f"＋ Agregar analizador propio", key=f"add_propio_{ck}",
                                 use_container_width=True):
                        st.session_state.analizadores_propios_por_concepto[concepto].append(
                            {"tipo": "", "marca": "", "modelo": "", "serie": ""}
                        )
                        st.rerun()

# Guardar periodicidad y tiempo en session_state y en el DataFrame
st.session_state.periodicidad_por_concepto           = periodicidades_concepto
st.session_state.tiempo_mantenimiento_por_concepto   = tiempos_concepto
equipos_a_mantener["PERIODICIDAD"]         = equipos_a_mantener["CONCEPTO"].map(periodicidades_concepto).fillna("Anual")
equipos_a_mantener["TIEMPO MANTENIMIENTO"] = equipos_a_mantener["CONCEPTO"].map(tiempos_concepto).fillna("")

# Construir analizadores por concepto para gestor_plantillas
analizadores_por_concepto = {}
analizadores_seleccionados_lista = []
for concepto in conceptos_seleccionados:
    bel     = [parse_analizador_display(v)
               for v in st.session_state.analizadores_bel_por_concepto.get(concepto, []) if v]
    propios = [p for p in st.session_state.analizadores_propios_por_concepto.get(concepto, [])
               if any(p.get(k) for k in ("tipo", "marca", "modelo", "serie"))]
    analizadores_por_concepto[concepto] = bel + propios
    analizadores_seleccionados_lista.extend(
        [v for v in st.session_state.analizadores_bel_por_concepto.get(concepto, []) if v]
    )

st.session_state.analizadores_por_concepto    = analizadores_por_concepto
st.session_state.analizadores_seleccionados   = list(dict.fromkeys(analizadores_seleccionados_lista))

# ════════════════════════════════════════════════════════════════════════════
# PASO 4 — Descargar documentación
# ════════════════════════════════════════════════════════════════════════════
with st.container(border=True):
    st.markdown("### 4. Descargar documentación")

    nombre_carpeta = st.text_input(
        "Nombre del paquete",
        value=f"reporte_{datetime.now():%Y%m%d_%H%M%S}",
        help="Este será el nombre del archivo ZIP que se descargará.",
        key="nombre_paquete_input",
    ).strip()
    if not nombre_carpeta:
        nombre_carpeta = f"reporte_{datetime.now():%Y%m%d_%H%M%S}"
    nombre_carpeta = re.sub(r"\W+", "_", nombre_carpeta).strip("_")

    col_boton, col_hojas, col_etiquetas = st.columns([4, 3, 3])
    with col_boton:
        generar = st.button("Generar paquete", type="primary", use_container_width=True)
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
                analizadores_seleccionados = st.session_state.get("analizadores_seleccionados", []),
                fecha_mantenimiento_base   = st.session_state.get("fecha_mantenimiento_por_concepto", {}),
            )

            if exitos > 0 or hacer_etiquetas:
                fecha_ref, periodo_mixto = resolver_fecha_referencia_drive(
                    st.session_state.get("fecha_mantenimiento_por_concepto", {}),
                    datetime.now().date(),
                )
                contenido_zip = buffer_zip.getvalue()
                st.session_state.ultimo_paquete_zip_bytes    = contenido_zip
                st.session_state.ultimo_paquete_zip_nombre   = f"{nombre_carpeta}.zip"
                st.session_state.ultimo_paquete_drive_folder = formatear_nombre_carpeta_documentacion(fecha_ref)
                st.session_state.ultimo_paquete_periodo      = fecha_ref.isoformat()
                st.session_state.ultimo_paquete_periodo_mixto = periodo_mixto
                st.session_state.ultimo_paquete_generado_en  = datetime.now().isoformat()

                st.download_button(
                    label               = "⬇️ Descargar Paquete (.zip)",
                    data                = contenido_zip,
                    file_name           = f"{nombre_carpeta}.zip",
                    mime                = "application/zip",
                    use_container_width = True,
                )
                st.page_link(
                    "pages/4_Google Drive.py",
                    label="Guardar también en Google Drive",
                    use_container_width=True,
                )

            if errores:
                with st.expander("Detalles de advertencias u omisiones"):
                    for err in errores:
                        st.warning(err)