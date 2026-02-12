from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet


def txt_a_pdf_tabla(path_txt: str, path_pdf: str):
    """
    Convierte un archivo TXT a PDF en formato de tabla.
    - path_txt: ruta del archivo TXT de entrada
    - path_pdf: ruta de salida del PDF
    """

    try:
        # Leer el archivo TXT
        with open(path_txt, "r", encoding="utf-8") as f:
            lineas = [linea.strip() for linea in f.readlines() if linea.strip()]

        # Convertir las líneas separadas por | en filas de tabla
        data = []
        for linea in lineas:
            if "|" in linea:
                columnas = [
                    col.strip() for col in linea.split("|") if col.strip() != ""
                ]
                data.append(columnas)

        if not data:
            raise Exception("El archivo TXT no contiene datos con estructura de tabla.")

        # Crear documento PDF
        doc = SimpleDocTemplate(path_pdf, pagesize=letter)
        estilo = getSampleStyleSheet()

        # Construir tabla
        tabla = Table(data)
        tabla.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, -1), 8),
                    ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
                    ("BACKGROUND", (0, 1), (-1, -1), colors.whitesmoke),
                    ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
                ]
            )
        )

        contenido = [
            Paragraph("Reporte de Validación de SOLPED", estilo["Heading2"]),
            tabla,
        ]

        doc.build(contenido)

        print(f"PDF generado correctamente en: {path_pdf}")
        return True

    except Exception as e:
        print(f"Error al generar PDF: {e}")
        return False
