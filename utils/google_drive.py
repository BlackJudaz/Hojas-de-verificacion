import json
import os
from datetime import date, datetime
from io import BytesIO
from pathlib import Path


SCOPES = ["https://www.googleapis.com/auth/drive.file"]
_BASE_DIR = Path(__file__).resolve().parent.parent
_OAUTH_CLIENT_CONFIG_PATH = _BASE_DIR / "datos" / "google_oauth_client.json"

_MESES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
]


def _normalizar_fecha(valor, fecha_por_defecto):
    if valor is None:
        return fecha_por_defecto

    if isinstance(valor, datetime):
        return valor.date()

    if isinstance(valor, date):
        return valor

    if hasattr(valor, "date") and callable(getattr(valor, "date")):
        try:
            return valor.date()
        except Exception:
            return fecha_por_defecto

    if isinstance(valor, str):
        texto = valor.strip()
        if not texto:
            return fecha_por_defecto
        try:
            return datetime.fromisoformat(texto).date()
        except ValueError:
            return fecha_por_defecto

    return fecha_por_defecto


def resolver_fecha_referencia_drive(fechas_por_concepto=None, fecha_por_defecto=None):
    fecha_base = _normalizar_fecha(fecha_por_defecto, datetime.now().date())
    if not isinstance(fechas_por_concepto, dict) or not fechas_por_concepto:
        return fecha_base, False

    fechas = []
    for valor in fechas_por_concepto.values():
        fecha_normalizada = _normalizar_fecha(valor, None)
        if fecha_normalizada is not None:
            fechas.append(fecha_normalizada)

    if not fechas:
        return fecha_base, False

    periodos_unicos = sorted({(fecha.year, fecha.month) for fecha in fechas})
    periodo_seleccionado = max(periodos_unicos)
    fecha_resuelta = date(periodo_seleccionado[0], periodo_seleccionado[1], 1)
    return fecha_resuelta, len(periodos_unicos) > 1


def formatear_nombre_carpeta_documentacion(fecha_mantenimiento):
    fecha_normalizada = _normalizar_fecha(fecha_mantenimiento, datetime.now().date())
    mes = _MESES[fecha_normalizada.month - 1]
    return f"Documentación MP {mes} {fecha_normalizada.year}"


def cargar_client_config(desde_json):
    client_config = json.loads(desde_json)
    if not isinstance(client_config, dict) or not any(k in client_config for k in ("installed", "web")):
        raise ValueError("El archivo OAuth no tiene el formato esperado de Google.")
    return client_config


def _resolver_ruta_client_config_drive():
    # Primero usa la ruta canónica y, si no existe, tolera una doble extensión común en Windows.
    candidatos = [_OAUTH_CLIENT_CONFIG_PATH, _OAUTH_CLIENT_CONFIG_PATH.with_suffix(".json.json")]
    for ruta in candidatos:
        if ruta.exists() and ruta.is_file():
            return ruta
    return _OAUTH_CLIENT_CONFIG_PATH


def obtener_ruta_client_config_drive():
    return _resolver_ruta_client_config_drive()


def cargar_client_config_local():
    ruta_config = _resolver_ruta_client_config_drive()
    if not ruta_config.exists():
        return None

    with ruta_config.open("r", encoding="utf-8") as archivo:
        return cargar_client_config(archivo.read())


def autorizar_google_drive(client_config):
    from google_auth_oauthlib.flow import InstalledAppFlow

    flow = InstalledAppFlow.from_client_config(client_config, SCOPES)
    credentials = flow.run_local_server(
        port=0,
        open_browser=True,
        prompt="select_account consent",
        authorization_prompt_message="Abre este enlace para autorizar la conexión con Google Drive: {url}",
        success_message="La conexión con Google Drive fue autorizada. Puedes volver a la aplicación."
    )
    return json.loads(credentials.to_json())


def construir_servicio_drive(credentials_info):
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build

    credentials = Credentials.from_authorized_user_info(credentials_info, SCOPES)
    if credentials.expired and credentials.refresh_token:
        credentials.refresh(Request())
        credentials_info = json.loads(credentials.to_json())

    service = build("drive", "v3", credentials=credentials)
    return service, credentials_info


def obtener_usuario_conectado(service):
    try:
        info = service.about().get(fields="user(displayName,emailAddress)").execute()
        return info.get("user", {})
    except Exception:
        return {}


def _escape_drive_query(texto):
    return str(texto).replace("'", "\\'")


def _construir_nombre_versionado(nombre_base, indice):
    if indice == 0:
        return nombre_base

    raiz, extension = os.path.splitext(nombre_base)
    if extension:
        return f"{raiz} ({indice}){extension}"
    return f"{nombre_base} ({indice})"


def obtener_nombre_disponible(service, nombre_base, mime_type=None, parent_id=None):
    indice = 0
    while True:
        candidato = _construir_nombre_versionado(nombre_base, indice)
        filtros = [
            f"name = '{_escape_drive_query(candidato)}'",
            "trashed = false"
        ]
        if mime_type:
            filtros.append(f"mimeType = '{mime_type}'")
        if parent_id:
            filtros.append(f"'{_escape_drive_query(parent_id)}' in parents")

        resultado = service.files().list(
            q=" and ".join(filtros),
            fields="files(id, name)",
            pageSize=1,
            supportsAllDrives=False
        ).execute()
        if not resultado.get("files"):
            return candidato
        indice += 1


def crear_carpeta_unica(service, nombre_base, parent_id=None):
    mime_type = "application/vnd.google-apps.folder"
    nombre_final = obtener_nombre_disponible(service, nombre_base, mime_type=mime_type, parent_id=parent_id)
    metadata = {
        "name": nombre_final,
        "mimeType": mime_type,
    }
    if parent_id:
        metadata["parents"] = [parent_id]

    carpeta = service.files().create(body=metadata, fields="id, name").execute()
    return carpeta


def subir_archivo_bytes(service, nombre_archivo, contenido, folder_id=None, mime_type="application/zip"):
    from googleapiclient.http import MediaIoBaseUpload

    nombre_final = obtener_nombre_disponible(service, nombre_archivo, mime_type=mime_type, parent_id=folder_id)
    metadata = {
        "name": nombre_final,
    }
    if folder_id:
        metadata["parents"] = [folder_id]

    media = MediaIoBaseUpload(BytesIO(contenido), mimetype=mime_type, resumable=False)
    archivo = service.files().create(
        body=metadata,
        media_body=media,
        fields="id, name, webViewLink"
    ).execute()
    return archivo