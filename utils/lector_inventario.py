# utils/lector_inventario.py
import pandas as pd
import re
import unicodedata
from html import unescape

COLUMNAS_INVENTARIO = {
    "# ACTIVO": [
        "# ACTIVO", "ACTIVO", "NO ACTIVO", "NO. ACTIVO", "N ACTIVO",
        "NUM ACTIVO", "NUM. ACTIVOS", "INVENTARIO", "INVENTARIO #"
    ],
    "CONCEPTO": [
        "CONCEPTO", "EQUIPO", "DESCRIPCION", "DESCRIPCIÓN", "ARTICULO", "ARTÍCULO",
        "PRODUCTO", "TIPO"
    ],
    "MARCA": [
        "MARCA", "BRAND"
    ],
    "MODELO": [
        "MODELO", "MODEL"
    ],
    "No. DE SERIE": [
        "NO DE SERIE", "NO. DE SERIE", "NUMERO DE SERIE", "NUM. DE SERIE",
        "N SERIE", "SERIE", "NUM SERIE", "NUM. DE SERIE", "SERIAL"
    ],
    "UBICACIÓN": [
        "UBICACION", "UBICACIÓN", "UBICACION 1", "UBICACIÓN 1", "AREA",
        "ÁREA", "DEPARTAMENTO", "LOCALIZACION", "LOCALIZACIÓN", "SITIO"
    ],
    "SUB UBICACIÓN": [
        "SUB UBICACION", "SUB UBICACIÓN", "SUBUBICACION", "SUB-UBICACIÓN",
        "UBICACION 2", "UBICACIÓN 2", "AREA SECUNDARIA", "SUBAREA", "SUB ÁREA"
    ]
}

COLUMNAS_REQUERIDAS = [
    "# ACTIVO", "CONCEPTO", "MARCA", "MODELO", "No. DE SERIE", "UBICACIÓN"
]


def normalizar_nombre_columna(nombre):
    nombre = str(nombre).strip().upper()
    nombre = unicodedata.normalize("NFKD", nombre)
    nombre = "".join(ch for ch in nombre if unicodedata.category(ch) != "Mn")
    nombre = re.sub(r"[^A-Z0-9]+", " ", nombre)
    return nombre.strip()


def nombre_coincide_con_patron(nombre_norm, patron):
    patron_norm = normalizar_nombre_columna(patron)
    if nombre_norm == patron_norm:
        return True
    if patron_norm in nombre_norm:
        return True
    return False


def detectar_objetivo_desde_valor(valor):
    valor_norm = normalizar_nombre_columna(valor)
    if not valor_norm:
        return None

    for nombre_objetivo, patrones in COLUMNAS_INVENTARIO.items():
        for patron in patrones:
            patron_norm = normalizar_nombre_columna(patron)
            if valor_norm == patron_norm:
                return nombre_objetivo
            if valor_norm.startswith(patron_norm) and len(valor_norm) <= len(patron_norm) + 10:
                return nombre_objetivo
            if patron_norm in valor_norm and len(valor_norm.split()) <= 4 and len(valor_norm) <= 28:
                return nombre_objetivo
    return None


def mapear_columnas_inventario(df):
    columnas = list(df.columns)
    normalizadas = {normalizar_nombre_columna(col): col for col in columnas}
    mapeo = {}

    for nombre_objetivo, patrones in COLUMNAS_INVENTARIO.items():
        for nombre_norm, columna_original in normalizadas.items():
            if any(nombre_coincide_con_patron(nombre_norm, patron) for patron in patrones):
                mapeo[nombre_objetivo] = columna_original
                break

    columnas_faltantes = [col for col in COLUMNAS_REQUERIDAS if col not in mapeo]
    if columnas_faltantes:
        return None, columnas_faltantes, mapeo

    df = df.rename(columns={original: target for target, original in mapeo.items()})

    if "SUB UBICACIÓN" not in df.columns:
        df["SUB UBICACIÓN"] = ""

    df = df.fillna("").astype(str).apply(lambda col: col.str.strip())
    df.columns = df.columns.str.strip()
    return df, [], mapeo


def buscar_valores_encabezado(fila):
    return [normalizar_nombre_columna(v) for v in fila.tolist() if pd.notna(v) and str(v).strip()]


def contar_coincidencias_encabezado(valores):
    matches = set()
    for valor in valores:
        for nombre_objetivo, patrones in COLUMNAS_INVENTARIO.items():
            if any(nombre_coincide_con_patron(valor, patron) for patron in patrones):
                matches.add(nombre_objetivo)
                break
    return len(matches), matches


def buscar_fila_encabezado(archivo, max_rows=60):
    try:
        if hasattr(archivo, "seek"):
            archivo.seek(0)

        df_encabezados = pd.read_excel(archivo, engine="openpyxl", header=None, nrows=max_rows, dtype=str)
        mejor_fila = None
        mejor_puntaje = 0
        for idx, fila in df_encabezados.iterrows():
            valores = buscar_valores_encabezado(fila)
            if not valores:
                continue

            puntaje, coincidencias = contar_coincidencias_encabezado(valores)
            if puntaje > mejor_puntaje:
                mejor_puntaje = puntaje
                mejor_fila = idx

        if mejor_puntaje >= 3:
            return mejor_fila
        return None
    except Exception:
        return None


def cargar_inventario(archivo):
    """
    Lee el archivo Excel del inventario, detecta columnas clave y devuelve un DataFrame limpio.
    """
    try:
        if hasattr(archivo, "seek"):
            archivo.seek(0)

        encabezado = buscar_fila_encabezado(archivo)
        intentos = []
        if encabezado is not None:
            intentos.append(encabezado)

        intentos.extend(i for i in range(0, 60) if i not in intentos)

        df = None
        df_mapeado = None
        error = None
        mapeo = {}

        for intento in intentos:
            if hasattr(archivo, "seek"):
                archivo.seek(0)
            if intento is None:
                df = pd.read_excel(archivo, engine="openpyxl", dtype=str)
            else:
                df = pd.read_excel(archivo, engine="openpyxl", header=intento, dtype=str)

            df.columns = df.columns.astype(str).str.strip()
            df = df.fillna("").astype(str).apply(lambda col: col.str.strip())

            df_mapeado, faltantes, mapeo = mapear_columnas_inventario(df)
            if df_mapeado is not None:
                return df_mapeado, None, mapeo

        # Si no encontramos una fila de encabezado válida, devolvemos el último error.
        columnas_encontradas = ", ".join(df.columns.tolist()) if df is not None else ""
        mensaje = ", ".join(faltantes) if df is not None else ""
        error = (
            f"No se pudieron detectar las columnas requeridas: {mensaje}. "
            f"Columnas encontradas: {columnas_encontradas}. "
            "Asegúrate de que la tabla de inventario tenga los encabezados correctos."
        )
        return None, error, {}
    except Exception as e:
        return None, f"Error al cargar el inventario: {e}", {}


def obtener_conceptos(df):
    """
    Retorna la lista de conceptos únicos del inventario, ordenados alfabéticamente.
    """
    return sorted(df["CONCEPTO"].dropna().unique().tolist())


def filtrar_por_concepto(df, concepto):
    """
    Filtra los equipos según el concepto seleccionado.
    """
    equipos = df[df["CONCEPTO"] == concepto].copy()
    return equipos.reset_index(drop=True)


def opciones_disponibles(df, campo):
    """
    Retorna los valores únicos disponibles en una columna, ordenados alfabéticamente.
    """
    if campo not in df.columns or df.empty:
        return []
    return sorted(df[campo].dropna().astype(str).unique().tolist())


def aplicar_filtros(df, filtros):
    """
    Aplica múltiples filtros al DataFrame.
    filtros: diccionario {columna: [valores seleccionados]}
    """
    resultado = df
    for columna, seleccion in filtros.items():
        if not seleccion:
            continue
        if columna == "# ACTIVO":
            resultado = resultado[resultado[columna].astype(str).isin(seleccion)]
        else:
            resultado = resultado[resultado[columna].isin(seleccion)]
    return resultado


def df_con_filtros(df, estados_filtros, excluir=None):
    """
    Aplica todos los filtros activos excepto el indicado en 'excluir'.
    Usado para actualizar dinámicamente las opciones de cada filtro.

    Parámetros:
        df: DataFrame completo del inventario
        estados_filtros: diccionario con los valores actuales de cada filtro
        excluir: key del filtro a ignorar (para que ese filtro muestre todas sus opciones)
    """
    campos = {
        "filtro_concepto": "CONCEPTO",
        "filtro_marca": "MARCA",
        "filtro_activo": "# ACTIVO",
        "filtro_ubicacion": "UBICACIÓN"
    }
    filtros = {}
    for key, columna in campos.items():
        if key == excluir:
            continue
        seleccion = estados_filtros.get(key)
        if seleccion:
            filtros[columna] = seleccion
    return aplicar_filtros(df, filtros)


def _limpiar_fragmento_html(fragmento):
    fragmento = re.sub(r"<br\s*/?>", " ", str(fragmento), flags=re.IGNORECASE)
    fragmento = re.sub(r"</(div|p|tr|td|li|span)>", " ", fragmento, flags=re.IGNORECASE)
    fragmento = re.sub(r"<[^>]+>", "", fragmento)
    fragmento = unescape(fragmento)
    return re.sub(r"\s+", " ", fragmento).strip()


def parsear_programacion_tinc(texto, html_texto=""):
    """
    Convierte el texto pegado de programación TINC en una tabla con folio e ID TINC.
    Espera bloques donde aparezca un folio SER... y un ID de activo AST... por registro.
    """
    if texto is None:
        return pd.DataFrame(columns=["folio", "url", "id_tinc", "bloque"])

    texto = str(texto).strip()
    if not texto:
        return pd.DataFrame(columns=["folio", "url", "id_tinc", "bloque"])

    texto = texto.replace("\r", "\n")
    html_texto = str(html_texto or "").strip()
    filas = []

    def _extraer_url_tinc(bloque, url_inicial=""):
        candidatos = []
        if url_inicial:
            candidatos.append(str(url_inicial).strip())

        candidatos.extend(re.findall(r"(?:https?|vscode-file|file)://[^\s)]+", str(bloque), flags=re.IGNORECASE))

        for candidato in candidatos:
            candidato = str(candidato).strip()
            if "app.cmmstinc.com" in candidato.lower():
                return candidato
        return ""

    patron_bloque = re.compile(
        r"\[(SER\d+)\]\(([^)]+)\)(.*?)(?=\[SER\d+\]\([^)]+\)|$)",
        flags=re.IGNORECASE | re.DOTALL,
    )

    patron_html = re.compile(
        r"<a[^>]*href=[\"']([^\"']+)[\"'][^>]*>\s*(SER\d+)\s*</a>(.*?)(?=<a[^>]*href=|$)",
        flags=re.IGNORECASE | re.DOTALL,
    )

    if html_texto:
        for match in patron_html.finditer(html_texto):
            folio = match.group(2).upper()
            url = _extraer_url_tinc(match.group(3), match.group(1))
            bloque = _limpiar_fragmento_html(folio + " " + match.group(3))
            id_match = re.search(r"AST\d+", bloque, flags=re.IGNORECASE)
            if not id_match:
                continue

            filas.append({
                "folio": folio,
                "url": url,
                "id_tinc": id_match.group(0).upper(),
                "bloque": bloque
            })

    if not filas:
        for match in patron_bloque.finditer(texto):
            folio = match.group(1).upper()
            url = _extraer_url_tinc(match.group(3), match.group(2))
            bloque = (match.group(1) + match.group(3)).strip()
            id_match = re.search(r"AST\d+", bloque, flags=re.IGNORECASE)
            if not id_match:
                continue

            filas.append({
                "folio": folio,
                "url": url,
                "id_tinc": id_match.group(0).upper(),
                "bloque": bloque
            })

    if not filas:
        texto_limpio = re.sub(r"\s+", " ", texto).strip()
        coincidencias_folio = list(re.finditer(r"SER\d+", texto_limpio, flags=re.IGNORECASE))
        for indice, coincidencia in enumerate(coincidencias_folio):
            inicio = coincidencia.start()
            fin = coincidencias_folio[indice + 1].start() if indice + 1 < len(coincidencias_folio) else len(texto_limpio)
            bloque = texto_limpio[inicio:fin].strip()
            folio = coincidencia.group(0).upper()
            url = _extraer_url_tinc(bloque)
            id_match = re.search(r"AST\d+", bloque, flags=re.IGNORECASE)
            if not id_match:
                continue

            filas.append({
                "folio": folio,
                "url": url,
                "id_tinc": id_match.group(0).upper(),
                "bloque": bloque
            })

    df = pd.DataFrame(filas, columns=["folio", "url", "id_tinc", "bloque"])
    if not df.empty:
        df = df.drop_duplicates(subset=["id_tinc"], keep="first").reset_index(drop=True)
    return df


def aplicar_programacion_tinc(df_inventario, df_programacion):
    """
    Agrega al inventario las columnas ID TINC y FOLIO TINC con base en la programación cargada.
    """
    if df_inventario is None:
        return None

    df = df_inventario.copy()
    if df_programacion is None or df_programacion.empty:
        if "ID TINC" not in df.columns:
            df["ID TINC"] = df.get("# ACTIVO", "")
        if "FOLIO TINC" not in df.columns:
            df["FOLIO TINC"] = ""
        if "URL TINC" not in df.columns:
            df["URL TINC"] = ""
        return df

    mapa_folios = (
        df_programacion.dropna(subset=["id_tinc", "folio"])
        .assign(id_tinc=lambda x: x["id_tinc"].astype(str).str.strip().str.upper())
        .assign(folio=lambda x: x["folio"].astype(str).str.strip().str.upper())
        .drop_duplicates(subset=["id_tinc"], keep="first")
        .set_index("id_tinc")["folio"]
        .to_dict()
    )
    mapa_urls = (
        df_programacion.dropna(subset=["id_tinc", "folio"])
        .assign(id_tinc=lambda x: x["id_tinc"].astype(str).str.strip().str.upper())
        .assign(url=lambda x: x["url"].astype(str).str.strip() if "url" in x.columns else "")
        .drop_duplicates(subset=["id_tinc"], keep="first")
        .set_index("id_tinc")["url"]
        .to_dict()
    )

    if "ID TINC" not in df.columns:
        df["ID TINC"] = df.get("# ACTIVO", "")
    else:
        df["ID TINC"] = df["ID TINC"].fillna("").astype(str).str.strip()
    if "URL TINC" not in df.columns:
        df["URL TINC"] = ""

    activos = df.get("# ACTIVO", pd.Series([""] * len(df), index=df.index)).fillna("").astype(str).str.strip().str.upper()
    ids_tinc = df["ID TINC"].fillna("").astype(str).str.strip().str.upper()
    df["FOLIO TINC"] = activos.map(mapa_folios).fillna("")
    df["URL TINC"] = activos.map(mapa_urls).fillna("")
    faltantes = df["FOLIO TINC"].eq("")
    if faltantes.any():
        df.loc[faltantes, "FOLIO TINC"] = ids_tinc[faltantes].map(mapa_folios).fillna("")
        df.loc[faltantes, "URL TINC"] = ids_tinc[faltantes].map(mapa_urls).fillna("")

    return df