import os
import pdfplumber
import re

# Plantillas de extracción definidas con patrones y un identificador único
plantillas_extraccion = {
    "tipo_1": {
        "identificador": r"PRIMA COMERCIAL",
        "prima": r"PRIMA COMERCIAL\s*:\s*([\d\.,]+)",
        "igv": r"IMPSTO\.GRAL\. A VENTAS\s*:\s*([\d\.,]+)",
        "importe_total": r"TOTAL A COBRAR(?:.*?)([\d\.,]+)"
    },
    "tipo_2": {
        "identificador": r"Prima",
        "prima": r"Prima\s*[:]?[\s]*([\d\.,]+)",
        "igv": r"IGV\s*[:]?[\s]*([\d\.,]+)|I\.G\.V\.\s*[:]?[\s]*([\d\.,]+)",
        "importe_total": r"Importe\s*Total\s*[:]?[\s]*([\d\.,]+)|TOTAL[\s]*([\d\.,]+)"
    }
}

# Ruta al archivo PDF
ruta_pdf = r'eps/vitapro-pe12/factura/PE12-SCTR-SALUD-11.23.pdf'

# Función para extraer texto del PDF
def extraer_texto_pdf(ruta_pdf):
    texto = ''
    with pdfplumber.open(ruta_pdf) as pdf:
        for pagina in pdf.pages:
            texto += pagina.extract_text()
    return texto

# Función para identificar el tipo de plantilla
def identificar_plantilla(texto, plantillas):
    for tipo, plantilla in plantillas.items():
        if re.search(plantilla["identificador"], texto):
            return tipo
    return None

# Función para extraer datos usando la plantilla correspondiente
def extraer_datos(texto, plantilla):
    datos = {}
    for campo, patron in plantilla.items():
        if campo != "identificador":
            match = re.search(patron, texto)
            if match:
                datos[campo] = match.group(1)
            else:
                datos[campo] = None
    return datos

# Extraer texto del PDF
texto_pdf = extraer_texto_pdf(ruta_pdf)
print("Texto extraído del PDF:\n", texto_pdf)

# Identificar la plantilla
tipo_plantilla = identificar_plantilla(texto_pdf, plantillas_extraccion)
print("Tipo de plantilla identificada:", tipo_plantilla)

if tipo_plantilla:
    # Extraer datos usando la plantilla identificada
    datos_extraidos = extraer_datos(texto_pdf, plantillas_extraccion[tipo_plantilla])
    print("Datos extraídos:", datos_extraidos)
else:
    print("No se pudo identificar la plantilla para el PDF proporcionado.")