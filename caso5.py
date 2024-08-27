import os
import pdfplumber
import pandas as pd
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
    },
    "tipo_f050": {
        "identificador": r"FACTURA ELECTRÓNICA",
        "prima": r"OP\. GRAVADA S/[\s]*([\d\.,]+)",
        "igv": r"I\.G\.V\. S/[\s]*([\d\.,]+)",
        "importe_total": r"IMPORTE TOTAL S/[\s]*([\d\.,]+)"
    }
}

# Función para identificar qué plantilla usar
def identificar_plantilla(texto):
    for nombre_plantilla, plantilla in plantillas_extraccion.items():
        if re.search(plantilla["identificador"], texto):
            return nombre_plantilla
    return None

# Función para extraer datos usando la plantilla identificada
def extraer_datos_con_plantilla(texto, plantilla):
    datos_extraidos = {}
    for key, pattern in plantilla.items():
        if key != "identificador":  # Saltar el identificador
            match = re.search(pattern, texto)
            if match:
                print(f"Match encontrado para {key}: {match.group(0)}")
                try:
                    datos_extraidos[key] = match.group(1).replace(',', '.').strip()
                except AttributeError as e:
                    print(f"Error al procesar el valor para {key}: {e}")
                    datos_extraidos[key] = None
            else:
                print(f"No se encontró un match para {key}")
                datos_extraidos[key] = None
                if key == "importe_total":
                    # Imprimir el contexto alrededor de "Importe Total" para depuración
                    contexto = re.search(r".{0,100}Importe Total.{0,100}", texto)
                    if contexto:
                        print(f"Contexto cercano a 'Importe Total': {contexto.group(0)}")
                    else:
                        print("No se encontró 'Importe Total' en el texto.")
    return datos_extraidos

# Función para verificar si el texto es seleccionable en todas las páginas del PDF
def es_pdf_sin_texto_seleccionable(pdf):
    for pagina in pdf.pages:
        texto = pagina.extract_text()
        if texto.strip():  # Si hay texto en alguna página, el PDF tiene texto seleccionable
            return False
    return True  # Si ninguna página tiene texto seleccionable, retorna True

# Función principal que recorre las carpetas y procesa los PDFs
def recorrer_carpetas_y_extraer_pdfs(ruta):
    pdfs_sin_texto_seleccionable = []  # Lista para almacenar los nombres de los PDFs problemáticos
    for root, dirs, files in os.walk(ruta):
        for file in files:
            if file.endswith('.pdf'):
                # Agregar la condición para omitir archivos que comiencen con "F050-"
                if file.startswith('F050-'):
                    continue
                pdf_path = os.path.join(root, file)
                print("#########################################################################################################")
                print(f"Procesando archivo: {file}")
                try:
                    with pdfplumber.open(pdf_path) as pdf:
                        if es_pdf_sin_texto_seleccionable(pdf):
                            print(f"No se pudo extraer texto de {file}.")
                            pdfs_sin_texto_seleccionable.append(file)
                            continue  # Saltar a la siguiente iteración si no se puede extraer texto
                        
                        all_text = ""
                        datos_extraidos = {}

                        # Procesar según el tipo de plantilla
                        first_page_text = pdf.pages[0].extract_text()
                        nombre_plantilla = identificar_plantilla(first_page_text)

                        if nombre_plantilla == "tipo_1":
                            # Procesar solo la primera página si es tipo_1
                            print("PDF identificado como tipo_1. Procesando solo la primera página.")
                            all_text = first_page_text
                        else:
                            # Procesar todas las páginas normalmente
                            for page in pdf.pages:
                                text = page.extract_text()
                                if text:
                                    all_text += text

                        # Imprimir el texto completo para depuración
                        print("Texto extraído del PDF:")
                        #print(all_text)  # Descomenta para ver el texto completo

                        # Identificar la plantilla adecuada
                        if nombre_plantilla:
                            print(f"Plantilla identificada: {nombre_plantilla}")
                            plantilla = plantillas_extraccion[nombre_plantilla]
                            datos_extraidos = extraer_datos_con_plantilla(all_text, plantilla)

                            # Imprimir los valores extraídos
                            print(f"Prima: {datos_extraidos.get('prima')}")
                            print(f"IGV: {datos_extraidos.get('igv')}")
                            print(f"Importe Total: {datos_extraidos.get('importe_total')}")

                            # Verificar si los datos son válidos y print OK o Fallo
                            if datos_extraidos.get('prima') and datos_extraidos.get('igv') and datos_extraidos.get('importe_total'):
                                print("Estado del archivo: OK")
                            else:
                                print("Estado del archivo: Fallo")

                            sociedad, codigo = buscar_sociedad_y_codigo(all_text, sociedades_df)
                            nuevo_nombre, proveedor, cod_proveedor, grupo_personal, imputacion, incluye = renombrar_pdf_y_validar(file, nuevos_nombres_df)
                            datos.append([file, sociedad, codigo, 
                                          datos_extraidos.get('prima'), 
                                          datos_extraidos.get('igv'), 
                                          datos_extraidos.get('importe_total'), 
                                          nuevo_nombre, proveedor, cod_proveedor, grupo_personal, imputacion, incluye])

                        else:
                            print(f"No se pudo identificar la plantilla para el archivo: {file}")

                except Exception as e:
                    print(f"Error al procesar el archivo {pdf_path}: {e}")

    # Imprimir los nombres de los PDFs que no tienen texto seleccionable
    if pdfs_sin_texto_seleccionable:
        print("\nLos siguientes archivos PDF no tienen texto seleccionable y podrían necesitar OCR:")
        for nombre_pdf in pdfs_sin_texto_seleccionable:
            print(nombre_pdf)
    else:
        print("\nTodos los archivos PDF tienen texto seleccionable.")

# Leer el archivo de rutas desde el Excel 
ruta_excel = r'C:\Users\user\Desktop\P_SCRIP\Em-ordenes-serv\rutas.xlsx'
rutas_df = pd.read_excel(ruta_excel, header=None)  # Cargar el Excel sin encabezados

# Obtener las rutas desde el Excel por indices
ruta_principal = rutas_df.iloc[1, 0]
ruta_descarga_excel = rutas_df.iloc[2, 0]  # Fila 3, Columna A (index 2, 0)

# Ruta del archivo Excel con las sociedades y nuevos nombres
ruta_sociedades =  r'C:\Users\user\Desktop\P_SCRIP\Em-ordenes-serv\Sociedades.xlsx'  
ruta_nuevos_nombres = r'C:\Users\user\Desktop\P_SCRIP\Em-ordenes-serv\Nuevos nombres.xlsx' 

# Leer los archivos Excel
try:
    sociedades_df = pd.read_excel(ruta_sociedades)  # Cargar el Excel de sociedades
    nuevos_nombres_df = pd.read_excel(ruta_nuevos_nombres)  # Cargar el Excel de nuevos nombres
except Exception as e:
    print(f"Error al leer los archivos Excel: {e}")
    exit()

# Lista para almacenar los datos extraídos
datos = []

def limpiar_texto(text):
    text = text.strip()  # Eliminar espacios al principio y final
    text = re.sub(r'[^A-Z0-9\s.,:/]', '', text)  # Eliminar caracteres especiales, mantener solo letras mayúsculas, números y espacios
    text = re.sub(r'\s+', ' ', text)  # Eliminar espacios múltiples
    return text

def buscar_sociedad_y_codigo(text, sociedades_df):
    text = limpiar_texto(text)
    for _, row in sociedades_df.iterrows():
        sociedad = limpiar_texto(row['SOCIEDAD'])
        if sociedad in text:
            return row['SOCIEDAD'], row['CÓDIGO']
    return "No identificado", "N/A"

def renombrar_pdf_y_validar(file_name, nuevos_nombres_df):
    if file_name.upper().startswith("F0"):
        nuevo_nombre = file_name.replace(".pdf", "")
    else:
        nombre_sin_codigo = file_name.split("-", 1)[-1]
        nombre_sin_codigo = limpiar_texto(nombre_sin_codigo.replace("pdf", "").strip())
        nuevo_nombre = re.sub(r'\b\d{2}\.\d{2}\b', '', nombre_sin_codigo).strip()

    nuevo_nombre = limpiar_texto(nuevo_nombre)

    for _, row in nuevos_nombres_df.iterrows():
        detalle = limpiar_texto(row['DETALLE'])

        if nuevo_nombre == detalle:
            proveedor = row['Proveedor']
            cod_proveedor = row['Cod. Proveedor']
            grupo_personal = row['Grupo de Personal']
            imputacion = row['I']
            incluye = row['Incluye']
            return nuevo_nombre, proveedor, cod_proveedor, grupo_personal, imputacion, incluye

    return nuevo_nombre, "N/A", "N/A", "N/A", "N/A", "N/A"

# Ejecutar el proceso
recorrer_carpetas_y_extraer_pdfs(ruta_principal)

# Crear un DataFrame con los datos extraídos y mostrarlo
df = pd.DataFrame(datos, columns=["Archivo", "Sociedad", "Código", "Prima", "IGV", "Importe Total", "DETALLE", "Proveedor", "Cod. Proveedor", "Grupo de Personal", "Imputacion", "Incluye"])
print(df)