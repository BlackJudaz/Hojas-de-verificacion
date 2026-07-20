import json
import mimetypes
import os
import socket
import time
from datetime import date, datetime
from io import BytesIO
from pathlib import Path, PurePosixPath
import zipfile


SCOPES = ["https://www.googleapis.com/auth/drive.file"]
_BASE_DIR = Path(__file__).resolve().parent.parent
_OAUTH_CLIENT_CONFIG_PATH = _BASE_DIR / "datos" / "google_oauth_client.json"
_OAUTH_TOKEN_PATH = _BASE_DIR / "datos" / "google_drive_token.json"
_ENV_OAUTH_CLIENT_JSON = "GOOGLE_OAUTH_CLIENT_JSON"
_ENV_OAUTH_CLIENT_PATH = "GOOGLE_OAUTH_CLIENT_PATH"

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


def construir_ruta_documentacion(fecha_mantenimiento=None):
    fecha_normalizada = _normalizar_fecha(fecha_mantenimiento, datetime.now().date())
    mes = _MESES[fecha_normalizada.month - 1]
    return ["Documentación MP", str(fecha_normalizada.year), mes]


def cargar_client_config(desde_json):
    client_config = json.loads(desde_json)
    if not isinstance(client_config, dict) or not any(k in client_config for k in ("installed", "web")):
        raise ValueError("El archivo OAuth no tiene el formato esperado de Google.")
    return client_config


def _obtener_secret_streamlit(clave):
    try:
        import streamlit as st
    except Exception:
        return None

    try:
        if clave in st.secrets:
            return st.secrets[clave]

        if "google_drive" in st.secrets:
            secretos_drive = st.secrets["google_drive"]
            if hasattr(secretos_drive, "get"):
                return secretos_drive.get(clave)
            if clave in secretos_drive:
                return secretos_drive[clave]
    except Exception:
        return None

    return None


def _resolver_ruta_externa_config(valor):
    texto = str(valor or "").strip()
    if not texto:
        return None
    return Path(os.path.expandvars(os.path.expanduser(texto)))


def _cargar_client_config_desde_archivo(ruta_config):
    if not ruta_config or not ruta_config.exists() or not ruta_config.is_file():
        return None

    with ruta_config.open("r", encoding="utf-8") as archivo:
        return cargar_client_config(archivo.read())


def _resolver_ruta_client_config_drive():
    # Primero usa la ruta canónica y, si no existe, tolera una doble extensión común en Windows.
    candidatos = [_OAUTH_CLIENT_CONFIG_PATH, _OAUTH_CLIENT_CONFIG_PATH.with_suffix(".json.json")]
    for ruta in candidatos:
        if ruta.exists() and ruta.is_file():
            return ruta
    return _OAUTH_CLIENT_CONFIG_PATH


def obtener_ruta_client_config_drive():
    ruta_externa = _resolver_ruta_externa_config(os.getenv(_ENV_OAUTH_CLIENT_PATH))
    if ruta_externa is not None:
        return ruta_externa

    ruta_secreto = _resolver_ruta_externa_config(_obtener_secret_streamlit("google_oauth_client_path"))
    if ruta_secreto is not None:
        return ruta_secreto

    return _resolver_ruta_client_config_drive()


def resolver_client_config_drive():
    json_entorno = str(os.getenv(_ENV_OAUTH_CLIENT_JSON, "")).strip()
    if json_entorno:
        return cargar_client_config(json_entorno), f"variable de entorno {_ENV_OAUTH_CLIENT_JSON}"

    json_secreto = _obtener_secret_streamlit("google_oauth_client_json")
    if json_secreto:
        return cargar_client_config(str(json_secreto)), "secreto de Streamlit google_oauth_client_json"

    ruta_externa = _resolver_ruta_externa_config(os.getenv(_ENV_OAUTH_CLIENT_PATH))
    if ruta_externa and ruta_externa.exists() and ruta_externa.is_file():
        return _cargar_client_config_desde_archivo(ruta_externa), f"variable de entorno {_ENV_OAUTH_CLIENT_PATH} ({ruta_externa})"

    ruta_secreto = _resolver_ruta_externa_config(_obtener_secret_streamlit("google_oauth_client_path"))
    if ruta_secreto and ruta_secreto.exists() and ruta_secreto.is_file():
        return _cargar_client_config_desde_archivo(ruta_secreto), f"secreto de Streamlit google_oauth_client_path ({ruta_secreto})"

    ruta_local = _resolver_ruta_client_config_drive()
    if ruta_local.exists() and ruta_local.is_file():
        return _cargar_client_config_desde_archivo(ruta_local), f"archivo local ({ruta_local})"

    return None, "sin configurar"


def cargar_client_config_local():
    client_config, _ = resolver_client_config_drive()
    return client_config


def autorizar_google_drive(client_config):
    from google_auth_oauthlib.flow import InstalledAppFlow

    flow = InstalledAppFlow.from_client_config(client_config, SCOPES)
    credentials = flow.run_local_server(
        port=0,
        open_browser=True,
        authorization_prompt_message="Abre este enlace para autorizar la conexión con Google Drive: {url}",
        success_message="Autorización completada. Puedes cerrar esta pestaña y volver a la aplicación."
    )
    return json.loads(credentials.to_json())

def construir_servicio_drive(credentials_info):
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build
    from google_auth_httplib2 import AuthorizedHttp
    import httplib2

    credentials = Credentials.from_authorized_user_info(credentials_info, SCOPES)
    if credentials.expired and credentials.refresh_token:
        credentials.refresh(Request())
        credentials_info = json.loads(credentials.to_json())

    http = AuthorizedHttp(credentials, http=httplib2.Http(timeout=180))
    service = build("drive", "v3", http=http, cache_discovery=False)
    return service, credentials_info


def obtener_usuario_conectado(service):
    try:
        info = service.about().get(fields="user(displayName,emailAddress)").execute()
        return info.get("user", {})
    except Exception:
        return {}


def _es_error_temporal_drive(exc):
    if isinstance(exc, (TimeoutError, socket.timeout, ConnectionError)):
        return True
    if isinstance(exc, OSError) and getattr(exc, "winerror", None) == 10060:
        return True

    texto = str(exc).lower()
    return any(patron in texto for patron in ["10060", "timed out", "timeout", "connection reset", "connection aborted"])


def _ejecutar_con_reintentos(request, intentos=4):
    ultimo_error = None
    for intento in range(intentos):
        try:
            return request.execute(num_retries=2)
        except Exception as exc:
            ultimo_error = exc
            if not _es_error_temporal_drive(exc) or intento == intentos - 1:
                raise
            time.sleep(2 + intento * 2)

    raise ultimo_error


def _escapar_valor_query_drive(valor):
    return str(valor).replace("\\", "\\\\").replace("'", "\\'")


def _buscar_archivos_en_carpeta(service, parent_id, nombre, mime_type=None):
    filtros = [
        "trashed = false",
        f"name = '{_escapar_valor_query_drive(nombre)}'",
        f"'{_escapar_valor_query_drive(parent_id)}' in parents",
    ]
    if mime_type:
        filtros.append(f"mimeType = '{_escapar_valor_query_drive(mime_type)}'")

    consulta = " and ".join(filtros)
    respuesta = _ejecutar_con_reintentos(service.files().list(
        q=consulta,
        spaces="drive",
        fields="files(id, name, mimeType, webViewLink)",
        pageSize=100
    ))
    return respuesta.get("files", [])


def obtener_o_crear_carpeta(service, nombre, parent_id=None):
    if parent_id:
        existentes = _buscar_archivos_en_carpeta(
            service,
            parent_id,
            nombre,
            mime_type="application/vnd.google-apps.folder"
        )
        if existentes:
            return existentes[0]

    metadata = {
        "name": nombre,
        "mimeType": "application/vnd.google-apps.folder",
    }
    if parent_id:
        metadata["parents"] = [parent_id]

    return _ejecutar_con_reintentos(service.files().create(body=metadata, fields="id, name, webViewLink"))


def obtener_o_crear_ruta_carpetas(service, nombres_carpeta, parent_id=None):
    carpetas = []
    carpeta_padre = parent_id
    for nombre in nombres_carpeta:
        carpeta = obtener_o_crear_carpeta(service, nombre, parent_id=carpeta_padre)
        carpetas.append(carpeta)
        carpeta_padre = carpeta["id"]
    return carpetas


def _separar_nombre_extension(nombre_archivo):
    base, extension = os.path.splitext(nombre_archivo)
    return base, extension


def resolver_nombre_unico_drive(service, nombre_archivo, folder_id):
    if not _buscar_archivos_en_carpeta(service, folder_id, nombre_archivo):
        return nombre_archivo

    base, extension = _separar_nombre_extension(nombre_archivo)
    indice = 1
    while True:
        candidato = f"{base}({indice}){extension}"
        if not _buscar_archivos_en_carpeta(service, folder_id, candidato):
            return candidato
        indice += 1


def crear_carpeta_unica(service, nombre_base, parent_id=None):
    metadata = {
        "name": nombre_base,
        "mimeType": "application/vnd.google-apps.folder",
    }
    if parent_id:
        metadata["parents"] = [parent_id]

    return _ejecutar_con_reintentos(service.files().create(body=metadata, fields="id, name, webViewLink"))


def subir_archivo_bytes(service, nombre_archivo, contenido, folder_id=None, mime_type="application/zip", usar_nombre_unico=True):
    from googleapiclient.http import MediaIoBaseUpload

    nombre_final = nombre_archivo
    if folder_id and usar_nombre_unico:
        nombre_final = resolver_nombre_unico_drive(service, nombre_archivo, folder_id)

    metadata = {
        "name": nombre_final,
    }
    if folder_id:
        metadata["parents"] = [folder_id]

    media = MediaIoBaseUpload(BytesIO(contenido), mimetype=mime_type, resumable=False)
    archivo = _ejecutar_con_reintentos(service.files().create(
        body=metadata,
        media_body=media,
        fields="id, name, webViewLink"
    ))
    return archivo


def _iterar_entradas_zip_validas(contenido_zip):
    with zipfile.ZipFile(BytesIO(contenido_zip), "r") as archivo_zip:
        info_list = [info for info in archivo_zip.infolist() if info.filename and not info.filename.startswith("__MACOSX/")]
        rutas_validas = []
        for info in info_list:
            ruta = PurePosixPath(info.filename)
            partes = [parte for parte in ruta.parts if parte not in ("", ".")]
            if not partes or any(parte == ".." for parte in partes):
                continue
            rutas_validas.append((info, partes))

        prefijo_comun = None
        if rutas_validas:
            primeras_partes = {partes[0] for _, partes in rutas_validas if partes}
            if len(primeras_partes) == 1:
                prefijo_comun = next(iter(primeras_partes))

        for info, partes in rutas_validas:
            if prefijo_comun and partes and partes[0] == prefijo_comun:
                partes = partes[1:]
            if not partes:
                continue
            yield archivo_zip, info, partes


def subir_zip_como_documentos(service, contenido_zip, folder_id):
    carpetas_cache = {(): folder_id}
    archivos_subidos = []

    for archivo_zip, info, partes in _iterar_entradas_zip_validas(contenido_zip):
        if info.is_dir():
            carpeta_actual = folder_id
            ruta_acumulada = []
            for parte in partes:
                ruta_acumulada.append(parte)
                clave = tuple(ruta_acumulada)
                if clave in carpetas_cache:
                    carpeta_actual = carpetas_cache[clave]
                    continue
                carpeta = obtener_o_crear_carpeta(service, parte, parent_id=carpeta_actual)
                carpeta_actual = carpeta["id"]
                carpetas_cache[clave] = carpeta_actual
            continue

        carpeta_actual = folder_id
        ruta_acumulada = []
        for parte in partes[:-1]:
            ruta_acumulada.append(parte)
            clave = tuple(ruta_acumulada)
            if clave not in carpetas_cache:
                carpeta = obtener_o_crear_carpeta(service, parte, parent_id=carpeta_actual)
                carpetas_cache[clave] = carpeta["id"]
            carpeta_actual = carpetas_cache[clave]

        nombre_archivo = partes[-1]
        mime_type = mimetypes.guess_type(nombre_archivo)[0] or "application/octet-stream"
        contenido = archivo_zip.read(info.filename)
        archivo = subir_archivo_bytes(
            service,
            nombre_archivo,
            contenido,
            folder_id=carpeta_actual,
            mime_type=mime_type,
            usar_nombre_unico=True
        )
        archivos_subidos.append(archivo)

    return archivos_subidos
