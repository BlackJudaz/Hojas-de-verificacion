import json
import os
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


def obtener_pestana(concepto):
    """
    Busca qué pestaña de plantilla corresponde a un concepto.
    """
    mapeo = cargar_mapeo()
    if mapeo is None:
        return None
    return mapeo.get(concepto, None)


def generar_reporte(equipo, ingeniero, hospital):
    """
    Toma la plantilla correspondiente al equipo, prelleana los campos fijos
    y guarda el reporte en la carpeta de reportes.
    """
    try:
        concepto = equipo["CONCEPTO"]
        pestana = obtener_pestana(concepto)

        if pestana is None:
            return None, f"No hay plantilla asignada para '{concepto}'"

        wb = load_workbook(RUTA_PLANTILLAS)

        if pestana not in wb.sheetnames:
            return None, f"La pestaña '{pestana}' no existe en el archivo de plantillas"

        ws = wb[pestana]

        ws["G12"] = hospital
        ws["G13"] = equipo.get("UBICACIÓN", "")
        ws["G14"] = equipo.get("# ACTIVO", "")
        ws["AA12"] = equipo.get("MARCA", "")
        ws["AA13"] = equipo.get("MODELO", "")
        ws["AA14"] = equipo.get("No. DE SERIE", "")

        os.makedirs(RUTA_REPORTES, exist_ok=True)

        fecha_hoy = date.today().strftime("%Y-%m-%d")
        activo = equipo.get("# ACTIVO", "SIN_ACTIVO")
        nombre_archivo = f"reporte_{activo}_{fecha_hoy}.xlsx"
        ruta_destino = os.path.join(RUTA_REPORTES, nombre_archivo)

        hojas_a_eliminar = [h for h in wb.sheetnames if h != pestana]
        for hoja in hojas_a_eliminar:
            del wb[hoja]

        wb.save(ruta_destino)
        return ruta_destino, None

    except Exception as e:
        return None, f"Error al generar reporte: {e}"


def crear_paquete_reporte(equipos, nombre_carpeta, ingeniero, hospital, 
                           progress_bar, status_text,
                           hacer_hojas=True, hacer_etiquetas=True):
    buffer_zip = BytesIO()
    errores    = []
    exitos     = 0

    with zipfile.ZipFile(buffer_zip, "w", zipfile.ZIP_DEFLATED) as zip_file:

        if hacer_hojas:
            for idx, (_, row) in enumerate(equipos.iterrows()):
                status_text.text(f"Procesando: {row['# ACTIVO']}")
                pestana = obtener_pestana(row["CONCEPTO"])

                if pestana is None:
                    errores.append(f"El equipo {row['# ACTIVO']} ({row['CONCEPTO']}) no tiene plantilla.")
                    progress_bar.progress((idx + 1) / len(equipos))
                    continue

                ruta_excel, error = generar_reporte(
                    equipo    = row.to_dict(),
                    ingeniero = ingeniero,
                    hospital  = hospital
                )

                if error:
                    errores.append(f"Error en {row['# ACTIVO']}: {error}")
                else:
                    nombre_archivo_final  = os.path.basename(ruta_excel)
                    tipo_equipo           = row["CONCEPTO"].strip()
                    ruta_dentro_del_zip   = f"{nombre_carpeta}/{tipo_equipo}/{nombre_archivo_final}"
                    zip_file.write(ruta_excel, arcname=ruta_dentro_del_zip)
                    exitos += 1

                    if os.path.exists(ruta_excel):
                        os.remove(ruta_excel)

                progress_bar.progress((idx + 1) / len(equipos))

        if hacer_etiquetas:
            status_text.text("Generando etiquetas PDF...")
            ruta_pdf = crear_etiquetas_pdf(equipos, ingeniero)
            if os.path.exists(ruta_pdf):
                zip_file.write(ruta_pdf, arcname=f"{nombre_carpeta}/etiquetas_mantenimiento.pdf")
                os.remove(ruta_pdf)

        status_text.empty()

    buffer_zip.seek(0)
    return buffer_zip, errores, exitos

def crear_etiquetas_pdf(equipos, ingeniero):
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
    margen_y        = 2 * mm
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
    fecha_hoy   = date.today().strftime("%d/%m/%Y")

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

        # ── Sección fecha / próximo ──────────────────────────────────────────
        y_fecha = y + 1*mm
        c.setFillColorRGB(0, 0, 0)
        c.setFont("Helvetica", fuente_campos)
        c.drawString(x + 2*mm,  y_fecha + 4*mm, "Fecha:")
        c.line(x + 11*mm, y_fecha + 3.5*mm, x + 29*mm, y_fecha + 3.5*mm)
        c.drawString(x + 11*mm, y_fecha + 4*mm, fecha_hoy)
        c.drawString(x + 31*mm, y_fecha + 4*mm, "Próximo:")
        c.line(x + 41*mm, y_fecha + 3.5*mm, x + aw - 2*mm, y_fecha + 3.5*mm)

        # ── Firma del ingeniero ──────────────────────────────────────────────
        y_firma = y_fecha + 8*mm
        c.setFont("Helvetica", fuente_campos)
        c.drawCentredString(x + aw/2, y_firma + 4*mm,
                            "Mantenimiento Preventivo realizado por:")
        c.line(x + 3*mm, y_firma, x + aw - 3*mm, y_firma)
        c.setFont("Helvetica-Bold", fuente_campos)
        c.drawCentredString(x + aw/2, y_firma + 0.5*mm,
                            f"{ingeniero}")

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

        y_campos = y_head - 4*mm
        for i, (etiqueta, valor) in enumerate(campos):
            yc          = y_campos - i * espacio_campos
            ancho_label = c.stringWidth(f"{etiqueta}:", "Helvetica-Bold", fuente_campos)
            c.setFont("Helvetica-Bold", fuente_campos)
            c.setFillColorRGB(0, 0, 0)
            c.drawString(x + 2*mm, yc, f"{etiqueta}:")
            lx = x + 2*mm + ancho_label + 1*mm
            c.line(lx, yc - 0.8*mm, x + aw - 2*mm, yc - 0.8*mm)
            c.setFont("Helvetica", fuente_campos)
            c.drawString(lx + 1*mm, yc, valor[:30])

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