import re
import unicodedata
import streamlit as st
import pandas as pd
from datetime import datetime

from utils.gestor_plantillas import crear_paquete_reporte
from utils.google_drive import (
    formatear_nombre_carpeta_documentacion,
    resolver_fecha_referencia_drive,
)
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


def _normalizar_texto(valor):
    texto = str(valor or "").strip().lower()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = re.sub(r"\s+", " ", texto)
    return texto.strip()


def _resolver_serie_desde_fila(fila):
    for clave in (
        "serie", "sn", "ns", "n_s", "num_serie", "numero_serie", "numero de serie",
        "SERIE", "SN", "NS", "NUMERO DE SERIE"
    ):
        valor = fila.get(clave, "") if hasattr(fila, "get") else ""
        texto = str(valor or "").strip()
        if texto:
            return texto
    return ""


def inicializar_estado():
    if "inventario_df" not in st.session_state:
        st.session_state.inventario_df = None
    if "clic_buscar" not in st.session_state:
        st.session_state.clic_buscar = False
    if "analizadores_seleccionados" not in st.session_state:
        st.session_state.analizadores_seleccionados = []
    if "periodicidad_por_concepto" not in st.session_state:
        st.session_state.periodicidad_por_concepto = {}
    if "tiempo_mantenimiento_por_concepto" not in st.session_state:
        st.session_state.tiempo_mantenimiento_por_concepto = {}
    if "analizadores_propios_por_concepto" not in st.session_state:
        st.session_state.analizadores_propios_por_concepto = {}
    if "fecha_mantenimiento_por_concepto" not in st.session_state:
        st.session_state.fecha_mantenimiento_por_concepto = {}
    if "ultimo_paquete_zip_bytes" not in st.session_state:
        st.session_state.ultimo_paquete_zip_bytes = b""
    if "ultimo_paquete_zip_nombre" not in st.session_state:
        st.session_state.ultimo_paquete_zip_nombre = ""
    if "ultimo_paquete_drive_folder" not in st.session_state:
        st.session_state.ultimo_paquete_drive_folder = ""
    if "ultimo_paquete_periodo" not in st.session_state:
        st.session_state.ultimo_paquete_periodo = ""
    if "ultimo_paquete_periodo_mixto" not in st.session_state:
        st.session_state.ultimo_paquete_periodo_mixto = False
    if "ultimo_paquete_generado_en" not in st.session_state:
        st.session_state.ultimo_paquete_generado_en = ""
    for key in ("filtro_concepto", "filtro_marca", "filtro_activo", "filtro_ubicacion"):
        if key not in st.session_state:
            st.session_state[key] = []
    if "filtro_tipo_activo_display" not in st.session_state:
        st.session_state.filtro_tipo_activo_display = []


def limpiar_filtros():
    st.session_state.clic_buscar = False
    st.session_state.filtro_concepto = []
    st.session_state.filtro_tipo_activo_display = []
    st.session_state.filtro_marca = []
    st.session_state.filtro_activo = []
    st.session_state.filtro_ubicacion = []


# ── Inicio de la página ──────────────────────────────────────────────────────
st.title("Generador de Hojas de Verificación")
st.caption("Filtra los equipos, selecciona los activos a trabajar y genera el paquete de hojas y etiquetas en un solo flujo.")

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
    st.markdown("### 1. Filtra los equipos")
    st.caption("Puedes filtrar por concepto, marca, número de activo o ubicación antes de seleccionar los equipos.")
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

        # Para que el buscador del multiselect no dependa de acentos,
        # se muestran opciones normalizadas y se traducen de vuelta
        # al concepto real para aplicar filtros sin romper la lógica.
        mapa_tipo_display_a_conceptos = {}
        opciones_tipo_display = []
        for concepto_original in opciones["CONCEPTO"]:
            clave_display = _normalizar_texto(concepto_original)
            if not clave_display:
                continue
            if clave_display not in mapa_tipo_display_a_conceptos:
                mapa_tipo_display_a_conceptos[clave_display] = []
                opciones_tipo_display.append(clave_display)
            mapa_tipo_display_a_conceptos[clave_display].append(concepto_original)

        seleccion_display_actual = [
            valor for valor in st.session_state.get("filtro_tipo_activo_display", [])
            if valor in mapa_tipo_display_a_conceptos
        ]
        if seleccion_display_actual != st.session_state.get("filtro_tipo_activo_display", []):
            st.session_state.filtro_tipo_activo_display = seleccion_display_actual

        fila1_col1, fila1_col2 = st.columns(2)
        with fila1_col1:
            st.multiselect(
                label="Tipo de activo",
                options=opciones_tipo_display,
                key="filtro_tipo_activo_display",
                placeholder="Selecciona uno o varios tipos de activo",
                label_visibility="collapsed",
                format_func=lambda x: str(x).title()
            )

            conceptos_filtrados = []
            for display in st.session_state.get("filtro_tipo_activo_display", []):
                conceptos_filtrados.extend(mapa_tipo_display_a_conceptos.get(display, []))
            st.session_state.filtro_concepto = list(dict.fromkeys(conceptos_filtrados))
        with fila1_col2:
            st.multiselect(label="Marca", options=opciones["MARCA"],
                           key="filtro_marca", placeholder="Selecciona una o varias marcas",
                           label_visibility="collapsed")

        fila2_col1, fila2_col2 = st.columns(2)
        with fila2_col1:
            st.multiselect(label="Activo", options=opciones["# ACTIVO"],
                           key="filtro_activo", placeholder="Selecciona uno o varios activos",
                           label_visibility="collapsed")
        with fila2_col2:
            st.multiselect(label="Ubicación", options=opciones["UBICACIÓN"],
                           key="filtro_ubicacion", placeholder="Selecciona una o varias ubicaciones",
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
        st.markdown("### 2. Selecciona los equipos")
        st.caption("Marca una o varias filas de la tabla para preparar sus hojas de verificación y etiquetas.")
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
        st.caption("Completa la periodicidad y los analizadores de cada tipo de activo antes de generar el paquete.")

        conceptos_seleccionados = equipos_a_mantener["CONCEPTO"].dropna().unique().tolist()

        # ── Analizadores ──────────────────────────────────────────────────────
        analizador_seleccionado_por_concepto = []
        analizadores_por_concepto = {}
        opciones_periodicidad = ["Bimestral", "Cuatrimestral", "Semestral", "Anual"]
        periodicidades_por_concepto = st.session_state.periodicidad_por_concepto.copy()
        tiempos_por_concepto = st.session_state.tiempo_mantenimiento_por_concepto.copy()

        st.markdown("### Selección de analizadores por tipo de equipo")
        if analizadores_df is None:
            st.info("No se pudo cargar la lista BEL. Aun así puedes capturar frecuencia y analizadores propios.")
            opciones_bel = []
        else:
            opciones_bel = list(dict.fromkeys(obtener_analizadores_display(analizadores_df)))

        for concepto in conceptos_seleccionados:
                if analizadores_df is not None:
                    sugeridos = buscar_analizadores_por_concepto(analizadores_df, [concepto])
                    opciones_sugeridas = list(dict.fromkeys(obtener_analizadores_display(sugeridos)))
                else:
                    opciones_sugeridas = []
                opciones_disponibles_bel = list(dict.fromkeys(opciones_sugeridas + opciones_bel))
                clave_concepto = re.sub(r"\W+", "_", concepto.strip().lower()).strip("_")
                key = f"analizadores_{clave_concepto}"
                key_propios_flag = f"usar_analizadores_propios_{clave_concepto}"
                key_propios_editor = f"analizadores_propios_editor_{clave_concepto}"
                if key not in st.session_state:
                    st.session_state[key] = []

                with st.container(border=True):
                    st.markdown(f"#### {concepto}")

                    st.markdown("**Frecuencia y tiempo de mantenimiento**")
                    col_freq, col_tiempo = st.columns([3, 2])
                    with col_freq:
                        default_periodicidad = periodicidades_por_concepto.get(concepto, "Anual")
                        if default_periodicidad not in opciones_periodicidad:
                            default_periodicidad = "Anual"
                        periodicidad_elegida = st.selectbox(
                            label=f"Periodicidad para {concepto}",
                            options=opciones_periodicidad,
                            index=opciones_periodicidad.index(default_periodicidad),
                            key=f"periodicidad_{clave_concepto}"
                        )
                        periodicidades_por_concepto[concepto] = periodicidad_elegida

                    with col_tiempo:
                        tiempo_actual = tiempos_por_concepto.get(concepto, "")
                        tiempo_manual = st.text_input(
                            label=f"Tiempo de mantenimiento para {concepto}",
                            key=f"tiempo_mantenimiento_{clave_concepto}",
                            value=tiempo_actual,
                            placeholder="Ej. 1 h, 2 h",
                            help="Sugerencia de formato: 1 h, 2 h. Puedes capturarlo como prefieras."
                        )
                        tiempos_por_concepto[concepto] = str(tiempo_manual or "").strip()

                    st.caption("El formato de tiempo es sugerido; el sistema permite capturarlo libremente.")

                    analizadores_propios_guardados = st.session_state["analizadores_propios_por_concepto"].get(concepto, [])
                    total_propios_guardados = len(analizadores_propios_guardados)
                    cupo_sugeridos = max(0, 3 - total_propios_guardados)
                    seleccion_actual = st.session_state.get(key, [])
                    if len(seleccion_actual) > cupo_sugeridos:
                        st.session_state[key] = seleccion_actual[:cupo_sugeridos]
                        seleccion_actual = st.session_state[key]
                        st.warning(
                            f"Solo se permiten 3 analizadores por tipo de equipo. Ya tienes {total_propios_guardados} propios guardados, por eso solo puedes elegir {cupo_sugeridos} sugeridos para '{concepto}'."
                        )

                    col_resumen_1, col_resumen_2 = st.columns(2)
                    col_resumen_1.metric("Máximo permitido", "3")
                    col_resumen_2.metric("Propios guardados", str(total_propios_guardados))

                    st.markdown("**Analizadores BEL disponibles**")
                    if opciones_sugeridas and opciones_sugeridas != opciones_bel:
                        st.caption("Las opciones sugeridas para este equipo aparecen primero, pero puedes escoger cualquier analizador BEL. Si borras analizadores propios guardados, aquí volverás a tener cupo disponible automáticamente.")
                    else:
                        st.caption("Puedes escoger cualquier analizador BEL. Si borras analizadores propios guardados, aquí volverás a tener cupo disponible automáticamente.")

                    seleccion = []
                    if cupo_sugeridos > 0 and opciones_disponibles_bel:
                        for indice in range(cupo_sugeridos):
                            clave_slot = f"{key}_slot_{indice}"
                            valor_actual = seleccion_actual[indice] if indice < len(seleccion_actual) else ""
                            valor_widget = st.session_state.get(clave_slot, valor_actual)
                            if valor_widget in opciones_disponibles_bel and valor_widget not in seleccion:
                                seleccion.append(valor_widget)

                            opciones_slot = [""] + [
                                opcion for opcion in opciones_disponibles_bel
                                if opcion not in seleccion or opcion == valor_widget
                            ]
                            if valor_widget not in opciones_slot:
                                valor_widget = ""

                            seleccion_slot = st.selectbox(
                                label=f"Analizador BEL {indice + 1} para {concepto}",
                                options=opciones_slot,
                                index=opciones_slot.index(valor_widget) if valor_widget in opciones_slot else 0,
                                key=clave_slot,
                                help="Selecciona un analizador BEL o deja el campo vacío si no lo necesitas."
                            )
                            if seleccion_slot and seleccion_slot not in seleccion:
                                seleccion.append(seleccion_slot)
                    elif cupo_sugeridos == 0:
                        st.info("Ya alcanzaste el máximo de 3 analizadores con los registros propios guardados.")

                    st.session_state[key] = seleccion
                    analizador_seleccionado_por_concepto.extend(seleccion)
                    analizadores_del_concepto = [
                        parse_analizador_display(item) for item in seleccion
                    ]

                    if not opciones_sugeridas:
                        st.info(f"No se encontraron sugerencias específicas para '{concepto}', pero puedes elegir cualquier analizador BEL o capturar uno propio si lo necesitas.")

                    st.markdown("**Analizadores propios**")
                    usar_propios = st.checkbox(
                        f"Agregar analizadores propios para {concepto}",
                        key=key_propios_flag,
                        value=bool(st.session_state["analizadores_propios_por_concepto"].get(concepto, [])),
                        help="Úsalo cuando el analizador no aparezca en las sugerencias o necesites registrar uno específico."
                    )

                    analizadores_propios = st.session_state["analizadores_propios_por_concepto"].get(concepto, [])
                    if usar_propios:
                        cupo_propios = max(0, 3 - len(seleccion))
                        if cupo_propios == 0:
                            st.warning(
                                f"Ya seleccionaste 3 analizadores BEL para '{concepto}'. No se pueden agregar más propios, pero se conserva lo que ya tenías guardado."
                            )
                            analizadores_propios = st.session_state["analizadores_propios_por_concepto"].get(concepto, [])
                            analizadores_por_concepto[concepto] = analizadores_del_concepto + analizadores_propios
                            continue

                        datos_guardados = st.session_state["analizadores_propios_por_concepto"].get(concepto, [])
                        if datos_guardados:
                            datos_normalizados = []
                            for item in datos_guardados:
                                datos_normalizados.append({
                                    "tipo": item.get("tipo", item.get("Analizador", "")),
                                    "marca": item.get("marca", ""),
                                    "modelo": item.get("modelo", ""),
                                    "serie": _resolver_serie_desde_fila(item)
                                })
                            df_propios_inicial = pd.DataFrame(datos_normalizados)
                        else:
                            df_propios_inicial = pd.DataFrame([
                                {"tipo": "", "marca": "", "modelo": "", "serie": ""}
                            ])

                        if len(df_propios_inicial.dropna(how="all")) > cupo_propios:
                            df_propios_inicial = df_propios_inicial.head(cupo_propios)

                        with st.form(key=f"form_{key_propios_editor}", clear_on_submit=False):
                            st.caption(
                                f"Completa hasta {cupo_propios} analizador(es) propio(s). Guarda los cambios antes de generar el paquete."
                            )
                            df_propios = st.data_editor(
                                df_propios_inicial,
                                key=key_propios_editor,
                                use_container_width=True,
                                hide_index=True,
                                num_rows="dynamic",
                                column_config={
                                    "tipo": st.column_config.TextColumn("Tipo de Analizador"),
                                    "marca": st.column_config.TextColumn("Marca"),
                                    "modelo": st.column_config.TextColumn("Modelo"),
                                    "serie": st.column_config.TextColumn("Numero de serie")
                                }
                            )
                            guardar_propios = st.form_submit_button(
                                "Guardar analizadores propios",
                                use_container_width=True
                            )

                        if guardar_propios:
                            nuevos_analizadores_propios = []
                            for _, fila in df_propios.fillna("").iterrows():
                                tipo = str(fila.get("tipo", "")).strip()
                                marca = str(fila.get("marca", "")).strip()
                                modelo = str(fila.get("modelo", "")).strip()
                                serie = _resolver_serie_desde_fila(fila)

                                if not any([tipo, marca, modelo, serie]):
                                    continue

                                nuevos_analizadores_propios.append({
                                    "tipo": tipo,
                                    "marca": marca,
                                    "modelo": modelo,
                                    "serie": serie,
                                    "sn": serie
                                })

                            if len(nuevos_analizadores_propios) > cupo_propios:
                                nuevos_analizadores_propios = nuevos_analizadores_propios[:cupo_propios]
                                st.error(
                                    f"Solo puedes usar 3 analizadores por tipo de equipo. Ya seleccionaste {len(seleccion)} de BEL, así que únicamente puedes guardar {cupo_propios} propio(s) para '{concepto}'."
                                )
                            else:
                                st.success("Información guardada correctamente.")

                            st.session_state["analizadores_propios_por_concepto"][concepto] = nuevos_analizadores_propios
                            analizadores_propios = nuevos_analizadores_propios
                    else:
                        if analizadores_propios:
                            st.info(
                                "Hay analizadores propios guardados para este concepto. Se seguirán usando al generar el paquete."
                            )
                            if st.button(
                                f"Borrar analizadores propios de {concepto}",
                                key=f"borrar_analizadores_propios_{clave_concepto}",
                                use_container_width=True
                            ):
                                st.session_state["analizadores_propios_por_concepto"][concepto] = []
                                analizadores_propios = []
                                st.success("Analizadores propios eliminados para este concepto.")
                                st.rerun()

                    if analizadores_propios:
                        st.caption(f"Analizadores propios registrados: {len(analizadores_propios)}")

                    st.markdown("**Fecha manual de mantenimiento para este tipo de equipo**")
                    key_fecha_manual = f"usar_fecha_manual_{clave_concepto}"
                    key_fecha = f"fecha_manual_{clave_concepto}"
                    usar_fecha_manual_concepto = st.checkbox(
                        f"Usar fecha manual para {concepto}",
                        key=key_fecha_manual,
                        value=concepto in st.session_state["fecha_mantenimiento_por_concepto"],
                        help="Si no se marca, las etiquetas de este tipo de equipo usarán la fecha actual."
                    )

                    if usar_fecha_manual_concepto:
                        fecha_guardada = st.session_state["fecha_mantenimiento_por_concepto"].get(concepto, datetime.now().date())
                        fecha_manual_concepto = st.date_input(
                            f"Fecha de mantenimiento para {concepto}",
                            value=fecha_guardada,
                            key=key_fecha,
                            help="Esta fecha se aplicará a las etiquetas de este tipo de equipo."
                        )
                        st.session_state["fecha_mantenimiento_por_concepto"][concepto] = fecha_manual_concepto
                    else:
                        st.session_state["fecha_mantenimiento_por_concepto"].pop(concepto, None)

                    analizadores_por_concepto[concepto] = analizadores_del_concepto + analizadores_propios

        st.session_state.periodicidad_por_concepto = periodicidades_por_concepto
        st.session_state.tiempo_mantenimiento_por_concepto = tiempos_por_concepto
        equipos_a_mantener["PERIODICIDAD"] = equipos_a_mantener["CONCEPTO"].map(
            periodicidades_por_concepto).fillna("Anual")
        equipos_a_mantener["TIEMPO MANTENIMIENTO"] = equipos_a_mantener["CONCEPTO"].map(
            tiempos_por_concepto).fillna("")

        st.session_state["analizadores_seleccionados"] = list(
            dict.fromkeys(analizador_seleccionado_por_concepto))
        st.session_state["analizadores_por_concepto"] = analizadores_por_concepto

        # ── Nombre del paquete ────────────────────────────────────────────────
        st.markdown("### 3. Configura la salida")
        nombre_carpeta = st.text_input(
            "Nombre del paquete",
            value=f"reporte_{datetime.now():%Y%m%d_%H%M%S}",
            help="Este será el nombre del archivo ZIP que se descargará."
        ).strip()
        if not nombre_carpeta:
            nombre_carpeta = f"reporte_{datetime.now():%Y%m%d_%H%M%S}"
        nombre_carpeta = re.sub(r"\W+", "_", nombre_carpeta).strip("_")

        # ── Generar ───────────────────────────────────────────────────────────
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
                    fecha_mantenimiento_base   = st.session_state.get("fecha_mantenimiento_por_concepto", {})
                )

                if exitos > 0 or hacer_etiquetas:
                    fecha_referencia_drive, periodo_mixto_drive = resolver_fecha_referencia_drive(
                        st.session_state.get("fecha_mantenimiento_por_concepto", {}),
                        datetime.now().date()
                    )
                    contenido_zip = buffer_zip.getvalue()
                    st.session_state.ultimo_paquete_zip_bytes = contenido_zip
                    st.session_state.ultimo_paquete_zip_nombre = f"{nombre_carpeta}.zip"
                    st.session_state.ultimo_paquete_drive_folder = formatear_nombre_carpeta_documentacion(
                        fecha_referencia_drive
                    )
                    st.session_state.ultimo_paquete_periodo = fecha_referencia_drive.isoformat()
                    st.session_state.ultimo_paquete_periodo_mixto = periodo_mixto_drive
                    st.session_state.ultimo_paquete_generado_en = datetime.now().isoformat()

                    st.download_button(
                        label               = "⬇️ Descargar Paquete (.zip)",
                        data                = contenido_zip,
                        file_name           = f"{nombre_carpeta}.zip",
                        mime                = "application/zip",
                        use_container_width = True
                    )
                    st.page_link(
                        "pages/4_Google Drive.py",
                        label="Guardar también en Google Drive",
                        use_container_width=True
                    )

                if errores:
                    with st.expander("Detalles de advertencias u omisiones"):
                        for err in errores:
                            st.warning(err)
    else:
        st.info("Selecciona uno o varios equipos de la tabla para continuar.")