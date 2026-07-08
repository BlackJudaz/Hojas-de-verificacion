import pandas as pd

def cargar_inventario(archivo):
    """
    Lee el archivo Excel del inventario y devuelve un DataFrame limpio.
    """
    try:
        df = pd.read_excel(
            archivo,
            dtype={"No. DE SERIE": str, "# ACTIVO": str}
        )
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        print(f"Error al cargar el inventario: {e}")
        return None


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