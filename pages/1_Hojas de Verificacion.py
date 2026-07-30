# pages/Hojas de Verificacion.py
import re
import json
import zipfile
from io import BytesIO
import streamlit as st
import pandas as pd
import time
from datetime import datetime, date
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode, JsCode

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
@st.fragment(run_every="2s" if st.session_state.get("google_drive_auth_url") and not st.session_state.get("google_drive_credentials") else None)
def _seccion_drive(contenido_zip):
    if not st.session_state.get("google_drive_credentials"):
        _sincronizar_sesion_drive_desde_otra_pestana(mostrar_mensajes=False)

    st.caption("Paso 1: conecta tu cuenta de Drive")
    if st.session_state.get("google_drive_credentials"):
        usuario = st.session_state.get("google_drive_usuario", {})
        correo  = usuario.get("emailAddress", "")
        if correo:
            st.success(f"Sesión iniciada: {correo}")
        else:
            st.success("Sesión de Drive iniciada")
        if st.button("Cerrar sesión de Drive", use_container_width=True, key="btn_cerrar_sesion_drive"):
            st.session_state.google_drive_credentials = None
            st.session_state.google_drive_usuario = {}
            st.session_state.google_drive_auth_url = ""
            st.session_state.google_drive_login_state = ""
            st.session_state.google_oauth_state = ""
            st.session_state.google_oauth_flow_config = None
            st.session_state.ultimo_drive_folder_link = ""
            st.rerun()
    else:
        if st.button("1) Iniciar sesión en Drive", use_container_width=True):
            _iniciar_sesion_drive()
            st.rerun()

        auth_url = st.session_state.get("google_drive_auth_url", "")
        if auth_url:
            st.markdown(
                f"<a href='{auth_url}' target='_blank' rel='noopener noreferrer' style='display:block;text-align:center;padding:0.45rem 0.75rem;border:1px solid rgba(151,166,195,0.35);border-radius:0.5rem;text-decoration:none;'>Continuar con Google (nueva pestaña)</a>",
                unsafe_allow_html=True,
            )
            st.info("Autoriza en Google. Esta sección se actualizará sola.")

    st.caption("Paso 2: sube el paquete")
    if st.session_state.get("google_drive_credentials"):
        if st.button("☁️ Subir a Drive", use_container_width=True, type="primary"):
            _subir_a_drive(contenido_zip)
    else:
        if st.session_state.get("google_drive_auth_url"):
            st.info("Esperando autorización de Google para habilitar la subida...")
        else:
            st.info("Primero inicia sesión para habilitar la subida.")

def _contar_archivos_en_zip(contenido_zip):
    try:
        with zipfile.ZipFile(BytesIO(contenido_zip), "r") as archivo_zip:
            total = 0
            for info in archivo_zip.infolist():
                if not info.filename or info.is_dir() or info.filename.startswith("__MACOSX/"):
                    continue

                nombre = str(info.filename).strip().lower()
                if nombre.endswith((".pdf", ".xlsx")):
                    total += 1
                    continue

                contenido = archivo_zip.read(info.filename)
                if contenido[:5] == b"%PDF-":
                    total += 1
                    continue

                try:
                    with zipfile.ZipFile(BytesIO(contenido), "r") as candidato_excel:
                        nombres = set(candidato_excel.namelist())
                        if "[Content_Types].xml" in nombres and "xl/workbook.xml" in nombres:
                            total += 1
                except Exception:
                    pass

            return total
    except Exception:
        return 0

def _subir_a_drive(contenido_zip):
    """Sube el paquete usando credenciales ya guardadas en session_state."""
    from utils import google_drive

    periodo_iso    = st.session_state.get("ultimo_paquete_periodo", "")
    credenciales   = st.session_state.get("google_drive_credentials")

    try:
        fecha_periodo      = datetime.fromisoformat(periodo_iso).date() if periodo_iso else datetime.now().date()
    except ValueError:
        fecha_periodo      = datetime.now().date()

    ruta_drive_destino = google_drive.construir_ruta_documentacion(fecha_periodo)
    total_archivos_zip = _contar_archivos_en_zip(contenido_zip)
    if total_archivos_zip == 0:
        st.error("El paquete no contiene documentos PDF/XLSX para subir. Verifica que se hayan generado hojas o etiquetas.")
        return

    # Evita mostrar links de una subida previa si esta ejecución falla.
    st.session_state.ultimo_drive_folder_link = ""

    try:
        with st.spinner("Subiendo a Google Drive..."):
            service, service_sheets, credenciales_actualizadas = google_drive.construir_servicios_google(credenciales)
            carpetas       = google_drive.obtener_o_crear_ruta_carpetas(service, ruta_drive_destino)
            carpeta_destino = carpetas[-1]
            archivos       = google_drive.subir_zip_como_documentos(
                service,
                contenido_zip,
                folder_id     = carpeta_destino["id"],
                service_sheets = service_sheets,
            )
            st.session_state.google_drive_credentials = credenciales_actualizadas
            st.session_state.google_drive_usuario     = google_drive.obtener_usuario_conectado(service)
            guardar_token_local = getattr(google_drive, "guardar_token_oauth_local", None)
            if callable(guardar_token_local):
                guardar_token_local(st.session_state.google_drive_credentials, st.session_state.google_drive_usuario)

        if not archivos:
            st.error("No se subió ningún archivo porque no se encontraron documentos PDF/XLSX válidos en el paquete.")
            return

        st.success(f"✅ Se subieron {len(archivos)} documento(s) a {' / '.join(ruta_drive_destino)}")
        st.session_state.ultimo_paquete_drive_folder = " / ".join(ruta_drive_destino)
        st.session_state.ultimo_drive_folder_link = carpeta_destino.get("webViewLink", "")

        total_fallback = sum(1 for a in archivos if a.get("_checkbox_fallback"))
        if total_fallback > 0:
            st.warning(f"{total_fallback} archivo(s) se subieron como .xlsx sin conversión a Google Sheets.")

        if carpeta_destino.get("webViewLink"):
            st.link_button("📁 Abrir carpeta en Google Drive", carpeta_destino["webViewLink"],
                           use_container_width=True)

    except Exception as exc:
        from utils import google_drive as gd
        if gd.es_error_de_scopes_google(exc):
            st.session_state.google_drive_credentials = None
            st.session_state.google_drive_auth_url = ""
            limpiar_token_local = getattr(gd, "limpiar_token_oauth_local", None)
            if callable(limpiar_token_local):
                limpiar_token_local()
            st.error("Los permisos de Drive expiraron. Vuelve a presionar 'Subir a Drive' para reconectarte.")
        elif gd.es_error_api_sheets_deshabilitada(exc):
            st.error("Google Sheets API está deshabilitada en tu proyecto de Google Cloud.")
        else:
            st.error(f"No se pudo subir a Google Drive: {exc}")


def _finalizar_oauth_drive_si_regreso():
    """Completa OAuth cuando Google regresa con ?code=..."""
    from utils import google_drive

    def _pantalla_cerrar_pestana(mensaje):
        st.markdown(
            f"""
            <style>
            [data-testid="stAppViewContainer"] {{
                background: #000000;
            }}
            section[data-testid="stSidebar"] {{
                display: none;
            }}
            [data-testid="stHeader"] {{
                display: none;
            }}
            .oauth-cerrar-wrap {{
                min-height: 92vh;
                display: flex;
                align-items: center;
                justify-content: center;
                text-align: center;
                color: #f5f5f5;
                font-size: 32px;
                font-weight: 700;
                letter-spacing: 0.4px;
                padding: 2rem;
            }}
            .oauth-cerrar-sub {{
                margin-top: 0.8rem;
                font-size: 16px;
                font-weight: 400;
                color: #bdbdbd;
            }}
            </style>
            <div class="oauth-cerrar-wrap">
                <div>
                    <div>{mensaje}</div>
                    <div class="oauth-cerrar-sub">Puedes volver a la pestaña principal.</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    params = st.query_params
    oauth_error = params.get("error", "")
    code = params.get("code", "")
    state = params.get("state", "")

    if oauth_error:
        st.session_state.google_drive_auth_url = ""
        st.query_params.clear()
        _pantalla_cerrar_pestana("Cierre esta pestaña")
        return True

    if not code:
        return False

    expected_state = st.session_state.get("google_oauth_state", "")
    flow_config = st.session_state.get("google_oauth_flow_config")
    if not flow_config and state:
        obtener_flow_config = getattr(google_drive, "obtener_flow_config_oauth", None)
        flow_config = obtener_flow_config(state) if callable(obtener_flow_config) else None
        if flow_config:
            st.session_state.google_oauth_state = state
            st.session_state.google_oauth_flow_config = flow_config
    if expected_state and state and state != expected_state:
        st.session_state.google_drive_auth_url = ""
        st.query_params.clear()
        _pantalla_cerrar_pestana("Cierre esta pestaña")
        return True

    if not flow_config or not flow_config.get("code_verifier"):
        st.session_state.google_drive_auth_url = ""
        st.query_params.clear()
        _pantalla_cerrar_pestana("Cierre esta pestaña")
        return True

    try:
        credenciales = google_drive.intercambiar_codigo_por_credenciales(code, state, flow_config)
        service, _, credenciales_actualizadas = google_drive.construir_servicios_google(credenciales)
        usuario = google_drive.obtener_usuario_conectado(service)
        st.session_state.google_drive_credentials = credenciales_actualizadas
        st.session_state.google_drive_usuario = usuario
        guardar_resultado = getattr(google_drive, "guardar_resultado_oauth", None)
        if callable(guardar_resultado) and state:
            guardar_resultado(state, credenciales_actualizadas, usuario)
        guardar_token_local = getattr(google_drive, "guardar_token_oauth_local", None)
        if callable(guardar_token_local):
            guardar_token_local(credenciales_actualizadas, usuario)
        st.session_state.google_drive_auth_url = ""
        st.query_params.clear()
        _pantalla_cerrar_pestana("Cierre esta pestaña")
    except Exception as exc:
        guardar_error = getattr(google_drive, "guardar_error_oauth", None)
        if callable(guardar_error) and state:
            guardar_error(state, str(exc))
        st.session_state.google_drive_auth_url = ""
        st.query_params.clear()
        _pantalla_cerrar_pestana("Autenticación fallida. Cierre esta pestaña")

    return True


def _iniciar_sesion_drive():
    """Inicia OAuth de Google Drive y muestra enlace de continuación."""
    from utils import google_drive

    try:
        client_config, _ = google_drive.resolver_client_config_drive()
    except Exception as exc:
        st.error(f"Configuración OAuth no válida: {exc}")
        return

    if client_config is None:
        st.error("No se encontró configuración OAuth de Google Drive. Contacta al administrador.")
        return

    try:
        auth_url, _ = google_drive.autorizar_google_drive(client_config)
        st.session_state.google_drive_auth_url = auth_url
        st.session_state.google_drive_login_state = st.session_state.get("google_oauth_state", "")
    except ModuleNotFoundError as exc:
        modulo = str(getattr(exc, "name", "") or "dependencia requerida")
        st.error(f"Falta instalar '{modulo}'. Ejecuta: pip install -r requirements.txt")
    except Exception as exc:
        st.error(f"No fue posible conectar con Google Drive: {exc}")


def _sincronizar_sesion_drive_desde_otra_pestana(mostrar_mensajes=True):
    from utils import google_drive

    estado = st.session_state.get("google_drive_login_state", "") or st.session_state.get("google_oauth_state", "")
    if not estado:
        if mostrar_mensajes:
            st.warning("No hay una autenticación pendiente para sincronizar.")
        return False

    obtener_resultado = getattr(google_drive, "obtener_resultado_oauth", None)
    resultado = obtener_resultado(estado, consume=True) if callable(obtener_resultado) else None

    if not resultado:
        if mostrar_mensajes:
            st.warning("No hay una autenticación pendiente para sincronizar.")
        return False

    error_oauth = str(resultado.get("error") or "").strip()
    if error_oauth:
        st.session_state.google_drive_auth_url = ""
        st.session_state.google_drive_login_state = ""
        if mostrar_mensajes:
            st.error(f"No se pudo completar la autenticación de Google: {error_oauth}")
        return False

    credenciales = resultado.get("credenciales") or {}
    if not credenciales:
        return False

    st.session_state.google_drive_credentials = credenciales
    st.session_state.google_drive_usuario = resultado.get("usuario") or {}
    st.session_state.google_drive_auth_url = ""
    st.session_state.google_drive_login_state = ""
    if mostrar_mensajes:
        st.success("Sesión de Drive sincronizada. Ya puedes subir el paquete.")
    return True

def _actualizar_estado_drive_desde_oauth():
    # Flujo estable: intentar sincronizar resultado OAuth y, si no aparece,
    # rehidratar desde token local guardado.
    sincronizado = _sincronizar_sesion_drive_desde_otra_pestana(mostrar_mensajes=False)
    restaurado = _restaurar_sesion_drive_desde_token_local()
    if sincronizado or restaurado:
        st.success("Sesión de Drive detectada correctamente.")
        st.rerun()
    else:
        st.info("Aún no se detecta la autorización. Termina el proceso en Google y vuelve a intentar.")


def _detectar_sesion_drive_automatica():
    sincronizado = _sincronizar_sesion_drive_desde_otra_pestana(mostrar_mensajes=False)
    restaurado = _restaurar_sesion_drive_desde_token_local()
    return bool(sincronizado or restaurado)


def _activar_polling_drive_si_pendiente():
    if st.session_state.get("google_drive_credentials"):
        return

    if not st.session_state.get("google_drive_auth_url"):
        return

    fragment_api = getattr(st, "fragment", None)
    if callable(fragment_api):
        try:
            @fragment_api(run_every="2s")
            def _poll_drive_session():
                if _detectar_sesion_drive_automatica():
                    st.rerun()

            _poll_drive_session()
            return
        except TypeError:
            # Esta versión de Streamlit no soporta run_every en fragmentos.
            pass

    # Fallback universal: reintenta cada 2s con rerun completo del app.
    # Se detiene solo en cuanto haya credenciales (ver el return del inicio).
    if _detectar_sesion_drive_automatica():
        st.rerun()

    time.sleep(2)
    st.rerun()
def _restaurar_sesion_drive_desde_token_local():
    from utils import google_drive

    if st.session_state.get("google_drive_credentials"):
        return False

    cargar_token_local = getattr(google_drive, "cargar_token_oauth_local", None)
    if not callable(cargar_token_local):
        return False

    payload = cargar_token_local()
    if not payload:
        return False

    credenciales = payload.get("credenciales") or {}
    if not credenciales:
        return False

    st.session_state.google_drive_credentials = credenciales
    usuario = payload.get("usuario") or {}
    if usuario:
        st.session_state.google_drive_usuario = usuario
        return True

    try:
        service, credenciales_actualizadas = google_drive.construir_servicio_drive(credenciales)
        st.session_state.google_drive_usuario = google_drive.obtener_usuario_conectado(service)
        st.session_state.google_drive_credentials = credenciales_actualizadas
        guardar_token_local = getattr(google_drive, "guardar_token_oauth_local", None)
        if callable(guardar_token_local):
            guardar_token_local(credenciales_actualizadas, st.session_state.google_drive_usuario)
    except Exception:
        pass

    return True

MAX_ANALIZADORES = 3
COLUMNA_ID = "# ACTIVO"

# Paleta visual de la tabla AgGrid (ajustable en un solo lugar)
TABLA_COLORES = {
    "fondo": "#0E1117",
    "fondo_encabezado": "#1B1F2A",
    "texto_encabezado": "#A7B0C0",
    "texto_fila": "#F3F7FF",
    "texto_id": "#F3F7FF",
    "linea": "#2A3140",
    "linea_fuerte": "#353D4D",
    "hover_fila": "#151A24",
    "seleccion_fila": "#1E2B3E",
    "acento": "#23363D",
    "checkbox_borde": "#5D87B5",
}


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
        "ultimo_drive_folder_link":         "",
        "ultimo_paquete_periodo":           "",
        "ultimo_paquete_periodo_mixto":     False,
        "ultimo_paquete_generado_en":       "",
        "google_drive_auth_url":            "",
        "google_drive_login_state":         "",
        "filtro_concepto":                  [],
        "filtro_tipo_activo_display":       [],
        "filtro_marca":                     [],
        "filtro_activo":                    [],
        "filtro_ubicacion":                 [],
        "ui_filtro_tipo_activo_display":    [],
        "ui_filtro_marca":                  [],
        "ui_filtro_activo":                 [],
        "ui_filtro_ubicacion":              [],
        "_selector_reset_token":            0,
        "_selector_sincronizar_visual":     False,
        "_selector_limpiar_visual":         False,
        "_selector_ids_visibles":           [],
        "_selector_grid_key_rendered":      "",
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
    st.session_state.ui_filtro_tipo_activo_display = []
    st.session_state.ui_filtro_marca         = []
    st.session_state.ui_filtro_activo        = []
    st.session_state.ui_filtro_ubicacion     = []
    st.session_state.pop("_selector_cache_clave", None)
    st.session_state.pop("_selector_df_base", None)
    st.session_state["_selector_df_editor"] = None
    st.session_state["_selector_sincronizar_visual"] = False
    st.session_state["_selector_limpiar_visual"] = False
    st.session_state["_selector_ids_visibles"] = []


def _marcar_entrada_pagina_hojas():
    pagina_actual = "page_hojas_verificacion"
    pagina_previa = st.session_state.get("_pagina_actual", "")
    st.session_state["_pagina_actual"] = pagina_actual
    return pagina_previa != pagina_actual


def _restaurar_widgets_filtro_desde_estado(forzar=False):
    # Los widgets pueden perder su estado al navegar entre paginas;
    # estas claves no-widget conservan la seleccion original.
    if "ui_filtro_tipo_activo_display" not in st.session_state or (
        forzar and not st.session_state.get("ui_filtro_tipo_activo_display") and st.session_state.get("filtro_tipo_activo_display")
    ):
        st.session_state.ui_filtro_tipo_activo_display = list(st.session_state.get("filtro_tipo_activo_display", []))
    if "ui_filtro_marca" not in st.session_state or (
        forzar and not st.session_state.get("ui_filtro_marca") and st.session_state.get("filtro_marca")
    ):
        st.session_state.ui_filtro_marca = list(st.session_state.get("filtro_marca", []))
    if "ui_filtro_activo" not in st.session_state or (
        forzar and not st.session_state.get("ui_filtro_activo") and st.session_state.get("filtro_activo")
    ):
        st.session_state.ui_filtro_activo = list(st.session_state.get("filtro_activo", []))
    if "ui_filtro_ubicacion" not in st.session_state or (
        forzar and not st.session_state.get("ui_filtro_ubicacion") and st.session_state.get("filtro_ubicacion")
    ):
        st.session_state.ui_filtro_ubicacion = list(st.session_state.get("filtro_ubicacion", []))


def _sincronizar_filtros_desde_widgets():
    st.session_state.filtro_tipo_activo_display = list(st.session_state.get("ui_filtro_tipo_activo_display", []))
    st.session_state.filtro_marca = list(st.session_state.get("ui_filtro_marca", []))
    st.session_state.filtro_activo = list(st.session_state.get("ui_filtro_activo", []))
    st.session_state.filtro_ubicacion = list(st.session_state.get("ui_filtro_ubicacion", []))


def _clave(concepto):
    return re.sub(r"\W+", "_", concepto.strip().lower()).strip("_")


def _total_analizadores(concepto):
    bel     = [x for x in st.session_state.analizadores_bel_por_concepto.get(concepto, []) if x]
    propios = st.session_state.analizadores_propios_por_concepto.get(concepto, [])
    return len(bel) + len(propios)


# ── Inicio de la página ──────────────────────────────────────────────────────
inicializar_estado()
if _finalizar_oauth_drive_si_regreso():
    st.stop()
_restaurar_sesion_drive_desde_token_local()

if not st.session_state.get("google_drive_credentials"):
    _sincronizar_sesion_drive_desde_otra_pestana(mostrar_mensajes=False)

st.title("Generador de Hojas de Verificación")
st.caption("Filtra los equipos, selecciona los activos a trabajar y genera el paquete de hojas y etiquetas en un solo flujo.")
_activar_polling_drive_si_pendiente()
_entro_desde_otra_pagina = _marcar_entrada_pagina_hojas()
_restaurar_widgets_filtro_desde_estado(forzar=_entro_desde_otra_pagina)

# Al volver desde otra page, fuerza remontaje del grid para rehidratar
# visualmente checks y sombreado con la seleccion persistida.
if _entro_desde_otra_pagina:
    st.session_state["_selector_reset_token"] = st.session_state.get("_selector_reset_token", 0) + 1
    st.session_state["_selector_grid_key_rendered"] = ""

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
        _sincronizar_filtros_desde_widgets()
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

        sel_display = [v for v in st.session_state.ui_filtro_tipo_activo_display if v in mapa_display]
        if sel_display != st.session_state.ui_filtro_tipo_activo_display:
            st.session_state.ui_filtro_tipo_activo_display = sel_display

        f1c1, f1c2 = st.columns(2)
        with f1c1:
            st.multiselect(
                label="Tipo de activo",
                options=opciones_display,
                key="ui_filtro_tipo_activo_display",
                placeholder="Selecciona uno o varios tipos de activo",
                label_visibility="collapsed",
                format_func=lambda x: str(x).title(),
            )
            conceptos_filtrados = []
            for d in st.session_state.ui_filtro_tipo_activo_display:
                conceptos_filtrados.extend(mapa_display.get(d, []))
            st.session_state.filtro_tipo_activo_display = list(st.session_state.ui_filtro_tipo_activo_display)
            st.session_state.filtro_concepto = list(dict.fromkeys(conceptos_filtrados))

        with f1c2:
            st.multiselect(label="Marca", options=opciones_calc["MARCA"],
                           key="ui_filtro_marca", placeholder="Selecciona una o varias marcas",
                           label_visibility="collapsed")
            st.session_state.filtro_marca = list(st.session_state.ui_filtro_marca)

        f2c1, f2c2 = st.columns(2)
        with f2c1:
            st.multiselect(label="Activo", options=opciones_calc["# ACTIVO"],
                           key="ui_filtro_activo", placeholder="Selecciona uno o varios activos",
                           label_visibility="collapsed")
            st.session_state.filtro_activo = list(st.session_state.ui_filtro_activo)
        with f2c2:
            st.multiselect(label="Ubicación", options=opciones_calc["UBICACIÓN"],
                           key="ui_filtro_ubicacion", placeholder="Selecciona una o varias ubicaciones",
                           label_visibility="collapsed")
            st.session_state.filtro_ubicacion = list(st.session_state.ui_filtro_ubicacion)

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

    st.session_state["_selector_cache_clave"] = clave_filtros
    st.session_state["_selector_df_base"] = df_mostrar[columnas_mostrar].copy()


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

        # Remonta el grid cuando cambia el filtro para evitar que AgGrid recicle
        # seleccion por indice y marque filas nuevas por error.
        filtro_hash = abs(hash(clave_filtros))
        clave_df = f"selector_df_equipos_{st.session_state.get('_selector_reset_token', 0)}_{filtro_hash}"
        ultima_key_renderizada = st.session_state.get("_selector_grid_key_rendered", "")
        es_primer_render_de_esta_vista = ultima_key_renderizada != clave_df
        df_tabla = df_base_tabla.copy()
        df_tabla[COLUMNA_ID] = df_tabla[COLUMNA_ID].astype(str)
        ids_en_tabla = set(df_tabla[COLUMNA_ID].astype(str))

        pre_selected_rows = [
            idx for idx, activo in enumerate(df_tabla[COLUMNA_ID].tolist())
            if activo in ids_previos
        ]

        if es_primer_render_de_esta_vista:
            ids_preseleccionados_visibles = sorted(ids_previos & ids_en_tabla)
            js_ids = json.dumps(ids_preseleccionados_visibles)
            sincronizar_seleccion_js = JsCode(
                f"""
                function(params) {{
                    const wanted = new Set({js_ids});
                    params.api.forEachNode(function(node) {{
                        const rowId = String((node.data && node.data['# ACTIVO']) || '');
                        node.setSelected(wanted.has(rowId));
                    }});
                }}
                """
            )
        else:
            sincronizar_seleccion_js = None

        gb = GridOptionsBuilder.from_dataframe(df_tabla)
        gb.configure_default_column(
            sortable=True,
            filter=False,
            resizable=True,
            editable=False,
            minWidth=120,
            cellStyle={"textAlign": "left"},
        )
        gb.configure_selection(
            selection_mode="multiple",
            use_checkbox=True,
            pre_selected_rows=pre_selected_rows,
            rowMultiSelectWithClick=False,
            suppressRowDeselection=False,
        )
        gb.configure_grid_options(
            suppressRowClickSelection=True,
            rowSelection="multiple",
            rowHeight=38,
            headerHeight=36,
            animateRows=False,
            suppressCellFocus=True,
            domLayout="normal",
            getRowId=JsCode("function(params) { return String(params.data['# ACTIVO'] || ''); }"),
            onFirstDataRendered=sincronizar_seleccion_js,
        )
        gb.configure_column(
            COLUMNA_ID,
            checkboxSelection=True,
            headerCheckboxSelection=True,
            headerCheckboxSelectionFilteredOnly=True,
            pinned="left",
            width=190,
            cellStyle={"fontWeight": 600, "color": TABLA_COLORES["texto_id"], "textAlign": "left"},
        )

        if "CONCEPTO" in df_tabla.columns:
            gb.configure_column("CONCEPTO", width=270)
        if "MARCA" in df_tabla.columns:
            gb.configure_column("MARCA", width=170)
        if "MODELO" in df_tabla.columns:
            gb.configure_column("MODELO", width=145)
        if "UBICACIÓN" in df_tabla.columns:
            gb.configure_column("UBICACIÓN", width=170)
        if "SUB UBICACIÓN" in df_tabla.columns:
            gb.configure_column("SUB UBICACIÓN", width=175)

        css_tabla = {
            ".ag-theme-streamlit": {
                "--ag-background-color": f"{TABLA_COLORES['fondo']}",
                "--ag-foreground-color": f"{TABLA_COLORES['texto_fila']}",
                "--ag-header-background-color": f"{TABLA_COLORES['fondo_encabezado']}",
                "--ag-header-foreground-color": f"{TABLA_COLORES['texto_encabezado']}",
                "--ag-row-border-color": f"{TABLA_COLORES['linea']}",
                "--ag-border-color": f"{TABLA_COLORES['linea_fuerte']}",
                "--ag-selected-row-background-color": f"{TABLA_COLORES['seleccion_fila']}",
                "--ag-row-hover-color": f"{TABLA_COLORES['hover_fila']}",
                "--ag-odd-row-background-color": f"{TABLA_COLORES['fondo']}",
            },
            ".ag-root-wrapper": {
                "border": f"1px solid {TABLA_COLORES['linea_fuerte']}",
                "border-radius": "6px",
                "overflow": "hidden",
                "background-color": f"{TABLA_COLORES['fondo']} !important",
            },
            ".ag-root": {
                "background-color": f"{TABLA_COLORES['fondo']} !important",
            },
            ".ag-body": {
                "background-color": f"{TABLA_COLORES['fondo']} !important",
            },
            ".ag-header": {
                "background-color": f"{TABLA_COLORES['fondo_encabezado']} !important",
                "border-bottom": f"1px solid {TABLA_COLORES['linea_fuerte']}",
            },
            ".ag-header-cell": {
                "font-size": "13px",
                "font-weight": "600",
                "color": TABLA_COLORES["texto_encabezado"],
                "letter-spacing": "0.2px",
                "font-family": "'Source Sans Pro', sans-serif",
                "text-align": "left",
            },
            ".ag-header-cell-label": {
                "justify-content": "flex-start",
                "text-align": "left",
            },
            ".ag-row": {
                "background-color": f"{TABLA_COLORES['fondo']} !important",
                "color": TABLA_COLORES["texto_fila"],
                "border-bottom": f"1px solid {TABLA_COLORES['linea']}",
            },
            ".ag-row-hover": {
                "background-color": f"{TABLA_COLORES['hover_fila']} !important",
            },
            ".ag-row-selected": {
                "background-color": f"{TABLA_COLORES['seleccion_fila']} !important",
            },
            ".ag-center-cols-viewport": {
                "background-color": f"{TABLA_COLORES['fondo']} !important",
            },
            ".ag-center-cols-container": {
                "background-color": f"{TABLA_COLORES['fondo']} !important",
            },
            ".ag-body-viewport": {
                "background-color": f"{TABLA_COLORES['fondo']} !important",
            },
            ".ag-center-cols-clipper": {
                "background-color": f"{TABLA_COLORES['fondo']} !important",
            },
            ".ag-pinned-left-cols-viewport": {
                "background-color": f"{TABLA_COLORES['fondo']} !important",
            },
            ".ag-cell": {
                "font-size": "13px",
                "font-weight": "600",
                "font-family": "'Source Sans Pro', sans-serif",
                "border-right": f"1px solid {TABLA_COLORES['linea']}",
                "text-align": "left",
            },
            ".ag-header-cell, .ag-cell": {
                "padding-left": "8px",
                "padding-right": "8px",
                "line-height": "1.2",
            },
            ".ag-pinned-left-cols-container": {
                "border-right": f"1px solid {TABLA_COLORES['linea_fuerte']}",
            },
            ".ag-checkbox-input-wrapper": {
                "transform": "scale(0.95)",
                "border": f"1px solid {TABLA_COLORES['checkbox_borde']}",
                "border-radius": "3px",
                "background-color": "transparent",
            },
            ".ag-checkbox-input-wrapper.ag-checked": {
                "background-color": TABLA_COLORES["acento"],
                "border-color": TABLA_COLORES["acento"],
            },
        }

        grid_response = AgGrid(
            df_tabla,
            gridOptions=gb.build(),
            key=clave_df,
            height=404,
            fit_columns_on_grid_load=False,
            update_mode=GridUpdateMode.SELECTION_CHANGED,
            data_return_mode=DataReturnMode.AS_INPUT,
            allow_unsafe_jscode=True,
            theme="streamlit",
            custom_css=css_tabla,
            reload_data=False,
        )

        st.session_state["_selector_ids_visibles"] = sorted(ids_en_tabla)

        seleccionados_grid = set()
        data_bruta = grid_response.get("data")
        if isinstance(data_bruta, pd.DataFrame):
            data_grid = data_bruta.copy()
            if COLUMNA_ID in data_grid.columns:
                data_grid[COLUMNA_ID] = data_grid[COLUMNA_ID].astype(str).str.strip()
            if "_selectedRowNodeInfo" in data_grid.columns and COLUMNA_ID in data_grid.columns:
                mask_sel = data_grid["_selectedRowNodeInfo"].notna()
                seleccionados_grid = set(data_grid.loc[mask_sel, COLUMNA_ID])

        seleccion_bruta = grid_response.get("selected_rows")
        if isinstance(seleccion_bruta, pd.DataFrame):
            filas_seleccionadas = seleccion_bruta.to_dict("records")
        elif isinstance(seleccion_bruta, list):
            filas_seleccionadas = seleccion_bruta
        else:
            filas_seleccionadas = []

        st.session_state["_selector_grid_key_rendered"] = clave_df

        if not seleccionados_grid:
            seleccionados_grid = set(
                str(fila.get(COLUMNA_ID, "")).strip()
                for fila in filas_seleccionadas
                if str(fila.get(COLUMNA_ID, "")).strip()
            )

        # En el primer render tras cambiar filtros, AgGrid puede reportar vacio
        # antes de hidratar preseleccion; no borrar seleccion existente por eso.
        if es_primer_render_de_esta_vista and not seleccionados_grid:
            ids_finales = ids_previos
        else:
            ids_finales = (ids_previos - ids_en_tabla) | seleccionados_grid

        st.session_state.ids_equipos_seleccionados = sorted(ids_finales)

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
                total_archivos_zip = _contar_archivos_en_zip(contenido_zip)
                if total_archivos_zip > 0:
                    st.session_state.ultimo_paquete_zip_bytes     = contenido_zip
                    st.session_state.ultimo_paquete_zip_nombre    = f"{nombre_carpeta}.zip"
                    st.session_state.ultimo_paquete_drive_folder  = formatear_nombre_carpeta_documentacion(fecha_ref)
                    st.session_state.ultimo_paquete_periodo       = fecha_ref.isoformat()
                    st.session_state.ultimo_paquete_periodo_mixto = periodo_mixto
                    st.session_state.ultimo_paquete_generado_en   = datetime.now().isoformat()
                else:
                    st.session_state.ultimo_paquete_zip_bytes = b""
                    st.error("Se generó un paquete vacío. Revisa advertencias u omisiones antes de subir a Drive.")

            if errores:
                with st.expander("Detalles de advertencias u omisiones"):
                    for err in errores:
                        st.warning(err)

<<<<<<< HEAD
    # ── Botones de descarga y Drive (aparecen cuando hay paquete generado) ──
=======
    # Botones de descarga y Drive — aparecen cuando hay paquete generado
>>>>>>> 5b31ff5496ec7342564149291278006abbc07057
    contenido_zip = st.session_state.get("ultimo_paquete_zip_bytes", b"")
    nombre_zip    = st.session_state.get("ultimo_paquete_zip_nombre", "")

    if contenido_zip and nombre_zip:
<<<<<<< HEAD
            st.divider()
            col_descargar, col_drive = st.columns(2)

            with col_descargar:
                st.download_button(
                    label               = "⬇️ Descargar Paquete (.zip)",
                    data                = contenido_zip,
                    file_name           = nombre_zip,
                    mime                = "application/zip",
                    use_container_width = True,
                )

            with col_drive:
                _seccion_drive(contenido_zip)
=======
        st.divider()
        col_descargar, col_drive = st.columns(2)

        with col_descargar:
            st.download_button(
                label               = "⬇️ Descargar Paquete (.zip)",
                data                = contenido_zip,
                file_name           = nombre_zip,
                mime                = "application/zip",
                use_container_width = True,
            )

        with col_drive:
            if st.session_state.get("google_drive_credentials"):
                usuario = st.session_state.get("google_drive_usuario", {})
                correo  = usuario.get("emailAddress", "")
                if correo:
                    st.caption(f"☁️ Conectado como {correo}")
                if st.button("☁️ Subir a Drive", use_container_width=True, type="primary"):
                    _subir_a_drive(contenido_zip)
            else:
                if st.button("☁️ Subir a Drive", use_container_width=True):
                    _conectar_y_subir_a_drive(contenido_zip)
>>>>>>> 5b31ff5496ec7342564149291278006abbc07057
