# utils/fechas.py
"""
Funciones de fecha compartidas entre gestor_plantillas y google_drive.
Centralizadas aquí para evitar duplicación.
"""
from datetime import date, datetime


_MESES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
]


def normalizar_fecha(valor, fecha_por_defecto):
    """
    Convierte distintos tipos de valor a un objeto date.
    Retorna fecha_por_defecto si no se puede convertir.
    """
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

    if isinstance(valor, dict):
        for clave in ("value", "date", "fecha", "selected"):
            if clave in valor:
                return normalizar_fecha(valor.get(clave), fecha_por_defecto)
        for item in valor.values():
            fecha_normalizada = normalizar_fecha(item, None)
            if fecha_normalizada is not None:
                return fecha_normalizada
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


def formato_mes_anio(fecha):
    """
    Convierte una fecha a string 'Mes Año'. Ej: 'Julio 2026'.
    """
    return f"{_MESES[fecha.month - 1]} {fecha.year}"


def calcular_siguiente_mantenimiento(fecha, periodicidad):
    """
    Calcula la fecha del próximo mantenimiento según la periodicidad.
    Periodicidades válidas: Bimestral, Cuatrimestral, Semestral, Anual.
    """
    periodicidad = str(periodicidad or "").strip().lower()

    if periodicidad == "bimestral":
        meses = 2
    elif periodicidad == "cuatrimestral":
        meses = 4
    elif periodicidad == "semestral":
        meses = 6
    else:
        return fecha.replace(year=fecha.year + 1)

    mes_nuevo = (fecha.month - 1 + meses) % 12 + 1
    anio_nuevo = fecha.year + ((fecha.month - 1 + meses) // 12)
    return fecha.replace(month=mes_nuevo, year=anio_nuevo)


def resolver_fecha_base(fecha_hoy, fecha_mantenimiento_base):
    """Resuelve la fecha base para etiquetas desde un valor simple."""
    return normalizar_fecha(fecha_mantenimiento_base, fecha_hoy)


def resolver_fecha_base_por_equipo(row, fecha_hoy, fecha_mantenimiento_base):
    """
    Resuelve la fecha base para un equipo específico.
    Si fecha_mantenimiento_base es un dict por concepto, busca el concepto del equipo.
    """
    if isinstance(fecha_mantenimiento_base, dict):
        concepto = str(row.get("CONCEPTO", "")).strip()
        fecha_concepto = fecha_mantenimiento_base.get(concepto)
        if fecha_concepto is not None:
            return normalizar_fecha(fecha_concepto, fecha_hoy)
        return fecha_hoy
    return resolver_fecha_base(fecha_hoy, fecha_mantenimiento_base)