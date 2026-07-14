# utils/lector_analizadores.py
from pathlib import Path
import re

import pandas as pd

RUTA_ANALIZADORES = Path(__file__).resolve().parent.parent / "datos" / "analizadores_bel.xlsx"


def cargar_analizadores(ruta=RUTA_ANALIZADORES):
    """
    Carga la lista fija de analizadores desde el archivo Excel.
    """
    path = Path(ruta)
    if not path.exists():
        fallback = Path.cwd() / "datos" / "analizadores_bel.xlsx"
        if fallback.exists():
            path = fallback
        else:
            fallback2 = Path.cwd().parent / "datos" / "analizadores_bel.xlsx"
            if fallback2.exists():
                path = fallback2

    try:
        df = pd.read_excel(path, dtype=str, engine="openpyxl")
        df.columns = [str(c).strip() for c in df.columns]
        df = df.dropna(how="all")
        df = df.fillna("").astype(str).apply(lambda col: col.str.strip())
        return df
    except Exception as e:
        print(f"Error al cargar lista de analizadores desde {path}: {e}")
        return None


def obtener_analizadores(df):
    """
    Retorna la lista de tipos de analizadores disponible en la hoja.
    """
    if df is None or df.empty:
        return []
    return sorted(df["SIMULADOR / ANALIZADOR"].dropna().unique().tolist())


def obtener_analizadores_display(df):
    """
    Devuelve una lista de opciones legibles para cada analizador.
    """
    if df is None or df.empty:
        return []
    filas = df.drop_duplicates(
        subset=["SIMULADOR / ANALIZADOR", "MARCA", "MODELO", "NUM. DE SERIE"], keep="first"
    )
    return filas.apply(
        lambda row: f"{row['SIMULADOR / ANALIZADOR']} | {row['MARCA']} | {row['MODELO']} | {row['NUM. DE SERIE']}",
        axis=1
    ).tolist()


def buscar_analizadores_por_concepto(df, conceptos):
    """
    Busca analizadores cuyos nombres coincidan con palabras del concepto seleccionado.
    """
    if df is None or df.empty or not conceptos:
        return df

    palabras = set()
    for concepto in conceptos:
        for palabra in re.split(r"\W+", str(concepto).lower()):
            palabra = palabra.strip()
            if palabra:
                palabras.add(palabra)

    if not palabras:
        return df

    mask = df["SIMULADOR / ANALIZADOR"].fillna("").str.lower().apply(
        lambda texto: any(palabra in texto for palabra in palabras)
    )
    resultado = df[mask]
    return resultado if not resultado.empty else df


def filtrar_analizadores(df, tipo=None, marca=None, modelo=None, serie=None):
    """
    Filtra la lista de analizadores por tipo, marca, modelo o número de serie.
    """
    if df is None or df.empty:
        return df

    resultado = df
    if tipo:
        resultado = resultado[resultado["SIMULADOR / ANALIZADOR"] == tipo]
    if marca:
        resultado = resultado[resultado["MARCA"] == marca]
    if modelo:
        resultado = resultado[resultado["MODELO"] == modelo]
    if serie:
        serie_str = str(serie).strip()
        resultado = resultado[resultado["NUM. DE SERIE"].str.contains(serie_str, case=False, na=False)]
    return resultado.reset_index(drop=True)
