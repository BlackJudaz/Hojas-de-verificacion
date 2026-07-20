from datetime import datetime

import streamlit as st

from utils import google_drive


def _formatear_periodo(periodo_iso):
    if not periodo_iso:
        return "No disponible"

    try:
        fecha = datetime.fromisoformat(periodo_iso).date()
    except ValueError:
        return "No disponible"

    return google_drive.formatear_nombre_carpeta_documentacion(fecha).replace("Documentación MP ", "")

def _formatear_fecha_hora(valor_iso):
    if not valor_iso:
        return "No disponible"

    try:
        return datetime.fromisoformat(valor_iso).strftime("%d/%m/%Y %H:%M:%S")
    except ValueError:
        return "No disponible"


def _resolver_fecha_periodo(periodo_iso):
    if not periodo_iso:
        return datetime.now().date()

    try:
        return datetime.fromisoformat(periodo_iso).date()
    except ValueError:
        return datetime.now().date()


def _tamano_legible(contenido):
    total_bytes = len(contenido or b"")
    if total_bytes < 1024:
        return f"{total_bytes} B"
    if total_bytes < 1024 * 1024:
        return f"{total_bytes / 1024:.1f} KB"
    return f"{total_bytes / (1024 * 1024):.2f} MB"


st.title("☁️ Google Drive")
st.info("Conecta tu cuenta de Google Drive para guardar el contenido del paquete como documentos separados, sin descargar nada localmente.")
st.divider()

if not hasattr(google_drive, "cargar_client_config_local") or not hasattr(google_drive, "resolver_client_config_drive"):
    st.error("La integración de Google Drive no terminó de cargarse correctamente. Reinicia la app para volver a intentarlo.")
    st.stop()

mensaje_flash = st.session_state.pop("google_drive_flash_ok", "")
if mensaje_flash:
    st.success(mensaje_flash)

contenido_zip = st.session_state.get("ultimo_paquete_zip_bytes", b"")
nombre_zip = st.session_state.get("ultimo_paquete_zip_nombre", "")
periodo_iso = st.session_state.get("ultimo_paquete_periodo", "")
nombre_carpeta_sugerido = st.session_state.get("ultimo_paquete_drive_folder", "")
periodo_mixto = st.session_state.get("ultimo_paquete_periodo_mixto", False)
fecha_generacion = st.session_state.get("ultimo_paquete_generado_en", "")
fecha_periodo = _resolver_fecha_periodo(periodo_iso)
ruta_drive_destino = google_drive.construir_ruta_documentacion(fecha_periodo)

if not contenido_zip or not nombre_zip:
    st.warning("Aún no hay un paquete ZIP disponible para subir. Primero genera un paquete en la página de hojas de verificación.")
    st.page_link("pages/1_Hojas de Verificacion.py", label="Ir a generar un paquete", use_container_width=True)
    st.stop()

with st.container(border=True):
    st.markdown("### 1. Último paquete detectado")
    col_1, col_2, col_3 = st.columns(3)
    col_1.metric("Archivo ZIP", nombre_zip)
    col_2.metric("Tamaño", _tamano_legible(contenido_zip))
    col_3.metric("Período sugerido", _formatear_periodo(periodo_iso))
    st.caption(f"Generado el {_formatear_fecha_hora(fecha_generacion)}")
    if periodo_mixto:
        st.warning("Se detectaron varias fechas manuales de mantenimiento dentro del mismo paquete. Revisa el nombre de la carpeta antes de subirlo a Drive.")

with st.container(border=True):
    st.markdown("### 2. Conexión con Google Drive")
    st.caption("Cada usuario conecta su propia cuenta desde el navegador. La configuración OAuth puede venir desde entorno, secretos de Streamlit o, como respaldo, desde un archivo local.")

    credenciales_drive = st.session_state.get("google_drive_credentials")
    usuario_drive = st.session_state.get("google_drive_usuario", {})
    client_config = None
    fuente_config_drive = "sin configurar"

    try:
        client_config, fuente_config_drive = google_drive.resolver_client_config_drive()
    except Exception as exc:
        st.error(f"La configuración OAuth de Google Drive no es válida: {exc}")

    config_disponible = client_config is not None
    if config_disponible:
        st.success("Es posible iniciar sesión y elegir su cuenta en el navegador.")
        st.caption(f"Fuente detectada: {fuente_config_drive}")
    else:
        st.warning(
            "No se encontró configuración OAuth. Define GOOGLE_OAUTH_CLIENT_JSON, GOOGLE_OAUTH_CLIENT_PATH o un secreto de Streamlit antes de conectar Google Drive."
        )

    col_conectar, col_desconectar = st.columns(2)
    with col_conectar:
        if st.button("Conectar mi cuenta de Google Drive", type="primary", use_container_width=True, disabled=not config_disponible):
            try:
                with st.spinner("Esperando autorización de Google Drive..."):
                    credenciales = google_drive.autorizar_google_drive(client_config)
                    service, credenciales_actualizadas = google_drive.construir_servicio_drive(credenciales)
                    usuario_conectado = google_drive.obtener_usuario_conectado(service)
                st.session_state.google_drive_credentials = credenciales_actualizadas
                st.session_state.google_drive_usuario = usuario_conectado
                st.session_state.google_drive_flash_ok = "Conexión con Google Drive completada para esta sesión."
                st.rerun()
            except Exception as exc:
                st.error(f"No fue posible conectar con Google Drive: {exc}")

    with col_desconectar:
        if st.button("Desconectar esta sesión", use_container_width=True, disabled=not bool(credenciales_drive)):
            st.session_state.google_drive_credentials = None
            st.session_state.google_drive_usuario = {}
            st.session_state.google_drive_flash_ok = "La cuenta de Google Drive se desconectó de esta sesión. Puedes conectar otra cuenta inmediatamente."
            st.rerun()

    if credenciales_drive:
        correo = usuario_drive.get("emailAddress", "")
        nombre = usuario_drive.get("displayName", "")
        if correo:
            st.caption(f"Estado actual: conectado como {nombre or correo} ({correo}).")
        else:
            st.caption("Estado actual: Google Drive conectado para esta sesión.")
    else:
        st.caption("Estado actual: sin conexión a Google Drive.")

with st.container(border=True):
    st.markdown("### 3. Guardar en Drive")
    st.write(f"Ruta destino: {' / '.join(ruta_drive_destino)}")
    if nombre_carpeta_sugerido:
        st.caption(f"Referencia anterior del paquete: {nombre_carpeta_sugerido}.")
    st.caption("Si ya existen archivos con el mismo nombre dentro de la carpeta de destino, se guardarán como 'archivo(1)', 'archivo(2)', etc.")

    if st.button("Subir último paquete a Google Drive", use_container_width=True, disabled=not bool(credenciales_drive), type="primary"):
        try:
            with st.spinner("Subiendo contenido a Google Drive..."):
                service, credenciales_actualizadas = google_drive.construir_servicio_drive(credenciales_drive)
                carpetas = google_drive.obtener_o_crear_ruta_carpetas(service, ruta_drive_destino)
                carpeta_destino = carpetas[-1]
                archivos = google_drive.subir_zip_como_documentos(service, contenido_zip, folder_id=carpeta_destino["id"])

                st.session_state.google_drive_credentials = credenciales_actualizadas
                st.session_state.google_drive_usuario = google_drive.obtener_usuario_conectado(service)

            st.success(
                f"Se subieron {len(archivos)} documento(s) a Google Drive en la ruta {' / '.join(ruta_drive_destino)}."
            )
            if carpeta_destino.get("webViewLink"):
                st.link_button("Abrir carpeta en Google Drive", carpeta_destino["webViewLink"], use_container_width=True)
        except Exception as exc:
            st.error(f"No se pudo subir el paquete a Google Drive: {exc}")