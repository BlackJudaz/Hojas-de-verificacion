import json
import os
import re
import unicodedata
import zipfile
from io import BytesIO
from openpyxl import load_workbook
from datetime import date

RUTA_MAPEO = "datos/mapeo_plantillas.json"
RUTA_PLANTILLAS = "datos/hojas_de_verificacion.xlsx"
RUTA_REPORTES = "reportes/"


def cargar_mapeo():
    """
    Lee el archivo JSON de mapeo y lo retorna como diccionario.
    """
    try:
        with open(RUTA_MAPEO, "r", encoding="utf-8") as archivo:
            mapeo = json.load(archivo)
        return mapeo
    except Exception as e:
        print(f"Error al cargar mapeo: {e}")
        return None


def _normalizar_nombre_archivo(nombre):
    nombre = str(nombre).strip()
    nombre = unicodedata.normalize("NFKD", nombre)
    nombre = "".join(c for c in nombre if not unicodedata.combining(c))
    nombre = nombre.replace('"', "'")
    nombre = re.sub(r"[\/\\:\*\?<>|]", " ", nombre)
    nombre = re.sub(r"\s+", " ", nombre)
    return nombre.strip()


def obtener_pestana(concepto):
    """
    Busca qué pestaña de plantilla corresponde a un concepto.
    """
    mapeo = cargar_mapeo()
    if mapeo is None:
        return None
    return mapeo.get(concepto, None)


def _encontrar_fila_encabezado_analizadores(ws):
    encabezados_buscados = {"equipo", "marca", "modelo", "serie"}

    for fila in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=35):
        encontrados = set()
        for cell in fila:
            if isinstance(cell.value, str):
                texto = cell.value.strip().lower()
                if texto in encabezados_buscados:
                    encontrados.add(texto)
        if len(encontrados) >= 3:
            return fila[0].row

    if ws.max_row >= 44:
        return 44
    if ws.max_row >= 10:
        return 10
    return 1


def _obtener_celda_para_escribir(ws, fila, columna):
    for merged_range in ws.merged_cells.ranges:
        if (
            merged_range.min_row <= fila <= merged_range.max_row and
            merged_range.min_col <= columna <= merged_range.max_col
        ):
            return ws.cell(row=merged_range.min_row, column=merged_range.min_col)
    return ws.cell(row=fila, column=columna)


def _escribir_valor(ws, fila, columna, valor):
    celda = _obtener_celda_para_escribir(ws, fila, columna)
    celda.value = valor


def _llenar_analizadores(ws, analizadores):
    def _parse_value(value):
        if isinstance(value, dict):
            return {
                "tipo": value.get("tipo", value.get("Analizador", value.get("analizador", ""))),
                "marca": value.get("marca", ""),
                "modelo": value.get("modelo", ""),
                "serie": value.get("serie", "")
            }
        if not value:
            return {"tipo": "", "marca": "", "modelo": "", "serie": ""}
        partes = [parte.strip() for parte in str(value).split("|")]
        return {
            "tipo": partes[0] if len(partes) > 0 else "",
            "marca": partes[1] if len(partes) > 1 else "",
            "modelo": partes[2] if len(partes) > 2 else "",
            "serie": partes[3] if len(partes) > 3 else ""
        }

    if analizadores is None:
        analizadores = []
    elif isinstance(analizadores, (str, dict)):
        analizadores = [analizadores]

    encabezado = _encontrar_fila_encabezado_analizadores(ws)
    filas_inicio = encabezado + 1
    filas_a_escribir = max(len(analizadores), 6)

    for idx in range(filas_a_escribir):
        fila_actual = filas_inicio + idx
        if idx < len(analizadores):
            analizador = _parse_value(analizadores[idx])
            _escribir_valor(ws, fila_actual, 1, "--")
            _escribir_valor(ws, fila_actual, 2, analizador.get("tipo", ""))
            _escribir_valor(ws, fila_actual, 12, analizador.get("marca", ""))
            _escribir_valor(ws, fila_actual, 21, analizador.get("modelo", ""))
            _escribir_valor(ws, fila_actual, 30, analizador.get("serie", ""))
        else:
            _escribir_valor(ws, fila_actual, 1, None)
            _escribir_valor(ws, fila_actual, 2, None)
            _escribir_valor(ws, fila_actual, 12, None)
            _escribir_valor(ws, fila_actual, 21, None)
            _escribir_valor(ws, fila_actual, 30, None)


def _obtener_identificador_activo(equipo):
    for clave in ["# ACTIVO", "ID TINC", "id tinc", "ID", "id"]:
        if clave in equipo and equipo.get(clave):
            return str(equipo.get(clave)).strip()
    return "SIN_ACTIVO"


def _obtener_celdas_firma_por_pestana(pestana):
    """Devuelve la celda para ingeniero y la celda para jefe según la pestaña."""
    celdas = {
        "UNIDAD DE ELECTROCIRUGIA": ("B58", "Q58"),
        "CENTRIFUGA": ("B53", "Q53"),
        "BASCULA": ("B47", "Q47"),
        "CAMA": ("B51", "Q51"),
        "MESA DE EXPLORACION": ("B56", "Q56"),
        "SIERRA": ("B47", "Q47"),
        "MAQUINA DE ANESTESIA": ("B106", "Q106"),
        "DESFIBRILADOR": ("B79", "Q79"),
        "MONITOR DE SIGNOS VITALES": ("B56", "Q56"),
    }
    return celdas.get(str(pestana).strip().upper(), (None, None))


def _normalize_text(s):
    if s is None:
        return ""
    s = str(s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return s.strip().lower()


def _buscar_y_escribir_firmas_por_texto(ws, ingeniero, jefe):
    """Busca en la hoja los rótulos que indiquen firma de ingeniero o jefe
    y escribe el nombre en la celda justo arriba del rótulo encontrado.
    """
    try:
        for fila in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in fila:
                val = cell.value
                if not isinstance(val, str):
                    continue
                norm = _normalize_text(val)
                if "nombre" in norm and "ingeniero" in norm:
                    if cell.row > 1:
                        _escribir_valor(ws, cell.row - 1, cell.column, ingeniero or "")
                elif "nombre" in norm and ("jefe" in norm or "jefe de" in norm or "jefe servicio" in norm):
                    if cell.row > 1:
                        _escribir_valor(ws, cell.row - 1, cell.column, jefe or "")
    except Exception:
        # No detener el flujo si hay problema con búsqueda; caerá al mapeo estático
        pass


def _obtener_tipo_activo(equipo):
    return str(equipo.get("CONCEPTO", "SIN_TIPO_ACTIVO")).strip()


def generar_reporte(equipo, ingeniero, jefe, hospital, analizadores=None):
    """
    Toma la plantilla correspondiente al equipo, prelleana los campos fijos
    y guarda el reporte en la carpeta de reportes.
    """
    try:
        concepto = equipo.get("CONCEPTO", "")
        pestana = obtener_pestana(concepto)

        if pestana is None:
            return None, f"No hay plantilla asignada para '{concepto}'"

        wb = load_workbook(RUTA_PLANTILLAS)

        if pestana not in wb.sheetnames:
            return None, f"La pestaña '{pestana}' no existe en el archivo de plantillas"

        ws = wb[pestana]

        ws["G12"] = hospital
        ws["G13"] = equipo.get("UBICACIÓN", "")
        ws["G14"] = equipo.get("# ACTIVO", equipo.get("ID TINC", equipo.get("id tinc", "")))
        ws["AA12"] = equipo.get("MARCA", "")
        ws["AA13"] = equipo.get("MODELO", "")
        ws["AA14"] = equipo.get("No. DE SERIE", "")

        # Intentar escritura dinámica: buscar etiquetas en la hoja y escribir arriba
        _buscar_y_escribir_firmas_por_texto(ws, ingeniero, jefe)

        # Si la búsqueda dinámica no encontró celdas, usar mapeo estático como respaldo
        ing_cell, jefe_cell = _obtener_celdas_firma_por_pestana(pestana)
        if ing_cell:
            try:
                if not ws[ing_cell].value:
                    ws[ing_cell] = ingeniero or ""
            except Exception:
                pass
        if jefe_cell:
            try:
                if not ws[jefe_cell].value:
                    ws[jefe_cell] = jefe or ""
            except Exception:
                pass

        if analizadores:
            _llenar_analizadores(ws, analizadores)

        os.makedirs(RUTA_REPORTES, exist_ok=True)

        tipo_activo = _normalizar_nombre_archivo(_obtener_tipo_activo(equipo))
        identificador = _normalizar_nombre_archivo(_obtener_identificador_activo(equipo))
        # Nuevo formato: ID TINC TIPO ACTIVO (sin comillas)
        nombre_archivo = f"Lista de Verificación {identificador} {tipo_activo}.xlsx"
        ruta_destino = os.path.join(RUTA_REPORTES, nombre_archivo)

        hojas_a_eliminar = [h for h in wb.sheetnames if h != pestana]
        for hoja in hojas_a_eliminar:
            del wb[hoja]

        wb.save(ruta_destino)
        return ruta_destino, None

    except Exception as e:
        return None, f"Error al generar reporte: {e}"


def crear_paquete_reporte(equipos, nombre_carpeta, ingeniero, jefe=None, hospital=None, 
                           progress_bar=None, status_text=None,
                           hacer_hojas=True, hacer_etiquetas=True,
                           analizador=None,
                           analizadores_por_concepto=None,
                           analizadores_seleccionados=None,
                           fecha_mantenimiento_base=None):
    buffer_zip = BytesIO()
    errores    = []
    exitos     = 0

    if analizadores_por_concepto is None:
        analizadores_por_concepto = {}
    if analizadores_seleccionados is None:
        analizadores_seleccionados = []

    nombre_carpeta = _normalizar_nombre_archivo(nombre_carpeta) or "paquete"

    with zipfile.ZipFile(buffer_zip, "w", zipfile.ZIP_DEFLATED) as zip_file:
        zip_file.writestr(f"{nombre_carpeta}/", "")

        if hacer_hojas:
            for idx, (_, row) in enumerate(equipos.iterrows()):
                activo = row.get('# ACTIVO', 'SIN_ACTIVO')
                status_text.text(f"Procesando: {activo}")
                concepto = row.get("CONCEPTO", "SIN_CONCEPTO")
                pestana = obtener_pestana(concepto)

                if pestana is None:
                    errores.append(f"El equipo {activo} ({concepto}) no tiene plantilla.")
                    progress_bar.progress((idx + 1) / len(equipos))
                    continue

                analizadores_para_equipo = analizadores_por_concepto.get(concepto, [])
                ruta_excel, error = generar_reporte(
                    equipo        = row.to_dict(),
                    ingeniero     = ingeniero,
                    jefe          = jefe,
                    hospital      = hospital,
                    analizadores  = analizadores_para_equipo
                )

                if error or not ruta_excel or not os.path.exists(ruta_excel):
                    detalle = error or "No se generó archivo de reporte."
                    errores.append(f"Error en {activo}: {detalle}")
                else:
                    nombre_archivo_final  = os.path.basename(ruta_excel)
                    tipo_equipo           = _normalizar_nombre_archivo(str(concepto or "SIN_TIPO_ACTIVO"))
                    carpeta_equipo        = f"{nombre_carpeta}/{tipo_equipo}"
                    zip_file.writestr(f"{carpeta_equipo}/", "")
                    ruta_dentro_del_zip   = f"{carpeta_equipo}/{nombre_archivo_final}"
                    zip_file.write(ruta_excel, arcname=ruta_dentro_del_zip)
                    exitos += 1

                    if os.path.exists(ruta_excel):
                        os.remove(ruta_excel)

                progress_bar.progress((idx + 1) / len(equipos))

        if hacer_etiquetas:
            status_text.text("Generando etiquetas PDF...")
            ruta_pdf = crear_etiquetas_pdf(
                equipos,
                ingeniero,
                fecha_mantenimiento_base=fecha_mantenimiento_base
            )
            if os.path.exists(ruta_pdf):
                zip_file.write(ruta_pdf, arcname=f"{nombre_carpeta}/etiquetas_mantenimiento.pdf")
                os.remove(ruta_pdf)

        if analizador is not None:
            if hasattr(analizador, "to_dict"):
                analizador = analizador.to_dict()
            if isinstance(analizador, dict):
                detalles = [f"{campo}: {valor}" for campo, valor in analizador.items()]
                contenido = "\n".join(detalles)
                zip_file.writestr(f"{nombre_carpeta}/analizador_seleccionado.txt", contenido)

        if analizadores_seleccionados:
            contenido = "\n".join(analizadores_seleccionados)
            zip_file.writestr(f"{nombre_carpeta}/analizadores_utilizados.txt", contenido)

        status_text.empty()

    buffer_zip.seek(0)
    return buffer_zip, errores, exitos

def _formato_mes_anio(fecha):
    nombres_meses = [
        "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
        "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
    ]
    return f"{nombres_meses[fecha.month - 1]} {fecha.year}"


def _calcular_siguiente_mantenimiento(fecha, periodicidad):
    periodicidad = str(periodicidad or "").strip().lower()
    if periodicidad == "bimestral":
        return fecha.replace(month=((fecha.month - 1 + 2) % 12) + 1,
                             year=fecha.year + ((fecha.month - 1 + 2) // 12))
    if periodicidad == "cuatrimestral":
        return fecha.replace(month=((fecha.month - 1 + 4) % 12) + 1,
                             year=fecha.year + ((fecha.month - 1 + 4) // 12))
    if periodicidad == "semestral":
        return fecha.replace(month=((fecha.month - 1 + 6) % 12) + 1,
                             year=fecha.year + ((fecha.month - 1 + 6) // 12))
    if periodicidad == "anual":
        return fecha.replace(year=fecha.year + 1)
    return fecha.replace(year=fecha.year + 1)


def _resolver_fecha_base(fecha_hoy, fecha_mantenimiento_base):
    if fecha_mantenimiento_base is None:
        return fecha_hoy
    return fecha_mantenimiento_base


def crear_etiquetas_pdf(equipos, ingeniero=None, fecha_mantenimiento_base=None):
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import mm
    from reportlab.lib.utils import ImageReader
    from datetime import date

    RUTA_PDF = "reportes/etiquetas_mantenimiento.pdf"
    os.makedirs(RUTA_REPORTES, exist_ok=True)

    ancho_pagina, alto_pagina = letter

    ancho_etiqueta  = 69 * mm
    alto_etiqueta   = 43 * mm
    columnas        = 3
    filas           = 6
    margen_x        = 2 * mm
    margen_y        = 0.5 * mm  # reducir margen superior adicional para subir todo el contenido
    espacio_col     = 2 * mm
    espacio_fila    = 2 * mm
    alto_header     = 8 * mm
    alto_pie        = 4 * mm
    fuente_campos   = 6
    fuente_pie      = 5
    fuente_contacto = 4
    espacio_campos  = 3 * mm

    c = canvas.Canvas(RUTA_PDF, pagesize=letter)
    col_actual  = 0
    fila_actual = 0
    fecha_hoy_obj = date.today()

    for _, row in equipos.iterrows():
        x  = margen_x + col_actual * (ancho_etiqueta + espacio_col)
        y  = alto_pagina - margen_y - alto_etiqueta - fila_actual * (alto_etiqueta + espacio_fila)
        aw = ancho_etiqueta
        ah = alto_etiqueta

        # ── Borde exterior ───────────────────────────────────────────────────
        c.setLineWidth(0.5)
        c.setStrokeColorRGB(0.1, 0.1, 0.1)
        c.rect(x, y, aw, ah)

        # ── Pie azul ─────────────────────────────────────────────────────────
        ap = alto_pie
        c.setFillColorRGB(0.192, 0.509, 0.580)
        c.rect(x, y, aw, ap, fill=True, stroke=False)
        c.setFillColorRGB(1, 1, 1)
        c.setFont("Helvetica", fuente_pie)
        c.drawCentredString(x + aw/2, y + ap * 0.62,
                            "En caso de mal funcionamiento reporte esta unidad al")
        c.drawCentredString(x + aw/2, y + ap * 0.22,
                            "Departamento de Ingenieria Biomedica de su Hospital")

        # ── Sección fecha (formato previo) ───────────────────────────────────
        y_fecha = y + 1*mm
        c.setFillColorRGB(0, 0, 0)
        c.setFont("Helvetica", fuente_campos)
        c.drawString(x + 2*mm,  y_fecha + 4*mm, "Fecha:")
        c.line(x + 11*mm, y_fecha + 3.5*mm, x + 29*mm, y_fecha + 3.5*mm)
        periodicidad = str(row.get("PERIODICIDAD", "Anual")).strip()
        fecha_etiqueta = _resolver_fecha_base(fecha_hoy_obj, fecha_mantenimiento_base)
        fecha_actual = _formato_mes_anio(fecha_etiqueta)
        c.drawString(x + 11*mm, y_fecha + 4*mm, fecha_actual)

        # ── Próximo mantenimiento ───────────────────────────────────────────
        fecha_siguiente_date = _calcular_siguiente_mantenimiento(fecha_etiqueta, periodicidad)
        fecha_siguiente = _formato_mes_anio(fecha_siguiente_date)
        # Alinear 'Próximo' a la derecha manteniendo consistencia
        x_prox_label = x + aw - 38*mm
        x_prox_date  = x + aw - 22*mm
        c.drawString(x_prox_label, y_fecha + 4*mm, "Próximo:")
        c.line(x_prox_date - 8*mm, y_fecha + 3.5*mm, x + aw - 2*mm, y_fecha + 3.5*mm)
        c.drawString(x_prox_date, y_fecha + 4*mm, fecha_siguiente)

        # ── Encabezado: logo + contacto ──────────────────────────────────────
        ah_h   = alto_header
        y_head = y + ah - ah_h

        c.setFillColorRGB(1, 1, 1)
        c.rect(x + 0.5, y_head, aw - 1, ah_h - 0.5, fill=True, stroke=False)

        # Redibujar borde encima del relleno blanco
        c.setLineWidth(0.5)
        c.setStrokeColorRGB(0.1, 0.1, 0.1)
        c.rect(x, y, aw, ah)

        try:
            logo = ImageReader("image_5976e1.png")
            c.drawImage(logo, x + 1*mm, y_head + 1.5*mm,
                        width=22*mm, height=ah_h - 3*mm,
                        preserveAspectRatio=True, mask="auto")
        except Exception:
            pass

        c.setFillColorRGB(0, 0, 0)
        c.setFont("Helvetica", fuente_contacto)
        c.drawRightString(x + aw - 1*mm, y_head + ah_h * 0.75, "Contacto:")
        c.drawRightString(x + aw - 1*mm, y_head + ah_h * 0.45, "+52(33) 38328790")
        c.drawRightString(x + aw - 1*mm, y_head + ah_h * 0.15, "018000051015")

        
        # ── Campos del equipo ────────────────────────────────────────────────
        campos = [
            ("Equipo",         str(row.get("CONCEPTO", ""))),
            ("Marca",          str(row.get("MARCA", ""))),
            ("Modelo",         str(row.get("MODELO", ""))),
            ("No. Serie",      str(row.get("No. DE SERIE", ""))),
            ("No. Inventario", str(row.get("# ACTIVO", ""))),
            ("Area",           str(row.get("UBICACIÓN", ""))),
        ]

        # Reducir el espaciado entre renglones para que todo quepa en la etiqueta
        field_step = 3 * mm
        # Subir los campos varios milímetros para ajustar posición
        y_campos = y_head - 2*mm
        for i, (etiqueta, valor) in enumerate(campos):
            # Bajar 'Equipo' un renglón y usar espaciado uniforme entre campos
            if i == 0:
                yc = y_campos
            else:
                yc = y_campos - i * field_step
            ancho_label = c.stringWidth(f"{etiqueta}:", "Helvetica-Bold", fuente_campos)
            c.setFont("Helvetica-Bold", fuente_campos)
            c.setFillColorRGB(0, 0, 0)
            c.drawString(x + 2*mm, yc, f"{etiqueta}:")
            lx = x + 2*mm + ancho_label + 1*mm
            c.line(lx, yc - 0.8*mm, x + aw - 2*mm, yc - 0.8*mm)
            c.setFont("Helvetica", fuente_campos)
            c.drawString(lx + 1*mm, yc, valor[:30])

        # ── Firma del ingeniero (directamente debajo de 'Area') ──────────────
        # Calcular la posición de la última línea de campo (Area) usando el nuevo paso
        y_last_field = y_campos - (len(campos) - 1) * field_step
        # Ajustar la sección de firma: bajar un renglón respecto a la posición actual
        # pero subirla luego media línea para ajustar visualmente
        y_title = y_last_field - 2*mm - field_step/1 + (field_step/2)
        y_name  = y_title - 3*mm
        y_line  = y_name - 2*mm
        c.setFont("Helvetica", fuente_campos)
        c.drawCentredString(x + aw/2, y_title, "Mantenimiento Preventivo realizado por:")
        c.setFont("Helvetica-Bold", fuente_campos)
        nombre_ingeniero = ingeniero or ""
        c.drawCentredString(x + aw/2, y_name, nombre_ingeniero)
        c.setLineWidth(0.5)
        c.line(x + 3*mm, y_line, x + aw - 3*mm, y_line)

        # ── Avanzar posición ─────────────────────────────────────────────────
        col_actual += 1
        if col_actual >= columnas:
            col_actual  = 0
            fila_actual += 1
        if fila_actual >= filas:
            fila_actual = 0
            col_actual  = 0
            c.showPage()

    c.save()
    return RUTA_PDF