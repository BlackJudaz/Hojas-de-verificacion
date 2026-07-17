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


def _tamano_legible(contenido):
    total_bytes = len(contenido or b"")
    if total_bytes < 1024:
        return f"{total_bytes} B"
    if total_bytes < 1024 * 1024:
        return f"{total_bytes / 1024:.1f} KB"
    return f"{total_bytes / (1024 * 1024):.2f} MB"


st.title("☁️ Google Drive")
st.info("Conecta tu cuenta de Google Drive para guardar también el último paquete ZIP generado por la aplicación.")
st.divider()

if not hasattr(google_drive, "cargar_client_config_local") or not hasattr(google_drive, "obtener_ruta_client_config_drive"):
    st.error("La integración de Google Drive no terminó de cargarse correctamente. Reinicia la app para volver a intentarlo.")
    st.stop()

contenido_zip = st.session_state.get("ultimo_paquete_zip_bytes", b"")
nombre_zip = st.session_state.get("ultimo_paquete_zip_nombre", "")
periodo_iso = st.session_state.get("ultimo_paquete_periodo", "")
nombre_carpeta_sugerido = st.session_state.get("ultimo_paquete_drive_folder", "")
periodo_mixto = st.session_state.get("ultimo_paquete_periodo_mixto", False)
fecha_generacion = st.session_state.get("ultimo_paquete_generado_en", "")

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
    st.caption("Cada usuario conecta su propia cuenta desde el navegador. La sesión se usa solo mientras esa persona tenga abierta la app.")

    credenciales_drive = st.session_state.get("google_drive_credentials")
    usuario_drive = st.session_state.get("google_drive_usuario", {})
    ruta_config_drive = google_drive.obtener_ruta_client_config_drive()
    client_config = None

    try:
        client_config = google_drive.cargar_client_config_local()
    except Exception as exc:
        st.error(f"La configuración OAuth local no es válida: {exc}")

    config_disponible = client_config is not None
    if config_disponible:
        st.success("La configuración OAuth local está lista. Cada usuario podrá iniciar sesión y elegir su cuenta en el navegador.")
    else:
        st.warning(
            f"Falta el archivo local de configuración OAuth en '{ruta_config_drive}'. Debe configurarse una sola vez para esta app; después, cada usuario solo tendrá que iniciar sesión en el navegador."
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
                st.success("Conexión con Google Drive completada para esta sesión.")
            except Exception as exc:
                st.error(f"No fue posible conectar con Google Drive: {exc}")

    with col_desconectar:
        if st.button("Desconectar esta sesión", use_container_width=True, disabled=not bool(credenciales_drive)):
            st.session_state.google_drive_credentials = None
            st.session_state.google_drive_usuario = {}
            st.success("La cuenta de Google Drive se desconectó de esta sesión. Puedes conectar otra cuenta inmediatamente.")

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
    st.markdown("### 3. Guardar el ZIP en Drive")
    carpeta_drive = st.text_input(
        "Nombre de la carpeta en Drive",
        value=nombre_carpeta_sugerido,
        help="Si ya existe una carpeta con ese nombre, la app creará automáticamente una versión con sufijo como '(1)'."
    ).strip()

    if not carpeta_drive:
        carpeta_drive = nombre_carpeta_sugerido

    st.caption("Si dentro de la carpeta ya existe un ZIP con el mismo nombre, también se guardará con sufijo '(1)', '(2)', etc.")

    if st.button("Subir último paquete a Google Drive", use_container_width=True, disabled=not bool(credenciales_drive), type="primary"):
        try:
            with st.spinner("Creando carpeta y subiendo archivo a Google Drive..."):
                service, credenciales_actualizadas = google_drive.construir_servicio_drive(credenciales_drive)
                carpeta = google_drive.crear_carpeta_unica(service, carpeta_drive)
                archivo = google_drive.subir_archivo_bytes(service, nombre_zip, contenido_zip, folder_id=carpeta["id"])
                st.session_state.google_drive_credentials = credenciales_actualizadas
                st.session_state.google_drive_usuario = google_drive.obtener_usuario_conectado(service)

            st.success(
                f"Archivo subido correctamente a Google Drive en la carpeta '{carpeta['name']}' como '{archivo['name']}'."
            )
            if archivo.get("webViewLink"):
                st.link_button("Abrir archivo en Google Drive", archivo["webViewLink"], use_container_width=True)
        except Exception as exc:
            st.error(f"No se pudo subir el paquete a Google Drive: {exc}")