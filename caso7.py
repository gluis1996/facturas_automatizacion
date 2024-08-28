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
                # print("#########################################################################################################")
                # print(f"Procesando archivo: {file}")
                nombre_sin_extension = os.path.splitext(file)[0]
                # print(nombre_sin_extension.split('-'))
                sociedad_1 = nombre_sin_extension.split('-')[0]
                detalle = ' '.join(nombre_sin_extension.split('-')[1:-1])
                # print(sociedad_1)
                # print(detalle)
                try:
                    with pdfplumber.open(pdf_path) as pdf:
                        if es_pdf_sin_texto_seleccionable(pdf):
                            # print(f"No se pudo extraer texto de {file}.")
                            pdfs_sin_texto_seleccionable.append(file)
                            continue  # Saltar a la siguiente iteración si no se puede extraer texto
                        
                        all_text = ""
                        datos_extraidos = {}

                        # Procesar según el tipo de plantilla
                        first_page_text = pdf.pages[0].extract_text()
                        nombre_plantilla = identificar_plantilla(first_page_text)

                        if nombre_plantilla == "tipo_1":
                            # Procesar solo la primera página si es tipo_1
                            # print("PDF identificado como tipo_1. Procesando solo la primera página.")
                            all_text = first_page_text
                        else:
                            # Procesar todas las páginas normalmente
                            for page in pdf.pages:
                                text = page.extract_text()
                                if text:
                                    all_text += text

                        # Imprimir el texto completo para depuración
                        # print("Texto extraído del PDF:")
                        #print(all_text)  # Descomenta para ver el texto completo

                        # Identificar la plantilla adecuada
                        if nombre_plantilla:
                            # print(f"Plantilla identificada: {nombre_plantilla}")
                            plantilla = plantillas_extraccion[nombre_plantilla]
                            datos_extraidos = extraer_datos_con_plantilla(all_text, plantilla)

                            # Imprimir los valores extraídos
                            # print(f"Prima: {datos_extraidos.get('prima')}")
                            # print(f"IGV: {datos_extraidos.get('igv')}")
                            # print(f"Importe Total: {datos_extraidos.get('importe_total')}")

                            # Verificar si los datos son válidos y print OK o Fallo
                            if datos_extraidos.get('prima') and datos_extraidos.get('igv') and datos_extraidos.get('importe_total'):
                                print("Estado del archivo: OK")
                            else:
                                print("Estado del archivo: Fallo")

                            # sociedad, codigo = buscar_sociedad_y_codigo(all_text, sociedades_df)
                            # nuevo_nombre, proveedor, cod_proveedor, grupo_personal, imputacion, incluye = renombrar_pdf_y_validar(file, nuevos_nombres_df)
                            datos.append([sociedad_1,  detalle,
                                        datos_extraidos.get('prima'), 
                                        datos_extraidos.get('igv'), 
                                        datos_extraidos.get('importe_total')])

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
ruta_excel = r'rutas.xlsx'
rutas_df = pd.read_excel(ruta_excel, header=None)  # Cargar el Excel sin encabezados

# Obtener las rutas desde el Excel por indices
ruta_principal = rutas_df.iloc[1, 0]
ruta_descarga_excel = rutas_df.iloc[2, 0]  # Fila 3, Columna A (index 2, 0)

# Ruta del archivo Excel con las sociedades y nuevos nombres
ruta_sociedades =  r'Sociedades.xlsx'  
ruta_nuevos_nombres = r'Nuevos nombres.xlsx' 

# Leer los archivos Excel
try:
    sociedades_df = pd.read_excel(ruta_sociedades)  # Cargar el Excel de sociedades
    nuevos_nombres_df = pd.read_excel(ruta_nuevos_nombres)  # Cargar el Excel de nuevos nombres
except Exception as e:
    print(f"Error al leer los archivos Excel: {e}")
    exit()

# Lista para almacenar los datos extraídos
datos = []

# Ejecutar el proceso
recorrer_carpetas_y_extraer_pdfs(ruta_principal)

# Crear un DataFrame con los datos extraídos y mostrarlo
df = pd.DataFrame(datos, columns=["Sociedad", "Detalle", "Prima", "IGV", "Importe Total"])
df1 = pd.read_excel('Nuevos nombres.xlsx')
grupo_articulo_varios = pd.read_excel('Grupo articulo y varios.xlsx')

print(df)
print(df1)


# Crear un DataFrame de nombrs nuevos

resultados =[]
# Formatear el campo "DISTINTIVO FACT" eliminando la primera parte y uniendo el resto con espacios
df1["FORMATTED_DISTINTIVO"] = df1["DISTINTIVO FACT"].str.split('-').str[1:].str.join(' ')

# Iterar sobre cada fila de 'df'
for index, row in df.iterrows():
    detalle = row["Detalle"].strip()  # Eliminar espacios en blanco al inicio y al final

    # Buscar coincidencias exactas de 'detalle' en 'df1'
    matches = df1[df1["FORMATTED_DISTINTIVO"].str.strip() == detalle]

    if not matches.empty:
        print(f"Detalle encontrado: {detalle}")
        # Repetir las coincidencias para cada ocurrencia en 'df'
        for _, match_row in matches.iterrows():
            # Crear un nuevo DataFrame con las coincidencias encontradas
            nuevo_df = pd.DataFrame({
                "Sociedad": [row["Sociedad"]],
                "DETALLE": [match_row["DETALLE"]],
                "Prima": [row["Prima"]],
                "IGV": [row["IGV"]],
                "Prima total": [row["Importe Total"]],
                "I": [match_row["I"]],
                "Grupo": [match_row["Grupo de Personal"]],                
                "Anticipo": [match_row["ANTICIPO"]],
                "Incluye": [match_row["Incluye"]],
                "Proveedor": [match_row["Proveedor"]],
                "Cod. Proveedor": [match_row["Cod. Proveedor"]],
                "Grupo Artículo": [match_row["Grupo Artículo"]],
                "N. SERVICIO": [match_row["N. SERVICIO"]]
            })
            resultados.append(nuevo_df)
    else:
        print(f"No se encontró detalle: {detalle}")

# Concatenar todos los resultados en un único DataFrame final
df_resultados = pd.concat(resultados, ignore_index=True)
print(df_resultados)




# Crear un DataFrame de grupo articulo y varios

resultados2 =[]
for index, row in df_resultados.iterrows():
    sociedad = row["Sociedad"].strip()  # Eliminar espacios en blanco al inicio y al final
    detalle = row["DETALLE"].strip()  # Eliminar espacios en blanco al inicio y al final
    
    # Buscar coincidencias exactas de 'detalle' en 'df1'
    # Corregir la condición de filtrado utilizando el operador `&` y paréntesis
    matches = grupo_articulo_varios[(grupo_articulo_varios["Codigo de Sociedad"] == sociedad) & 
                                    (grupo_articulo_varios["DETALLE"].str.strip().str.upper() == detalle.upper())]

    if not matches.empty:
        print(f"Grupo encontrado: {sociedad}")
        # Repetir las coincidencias para cada ocurrencia en 'df'
        for _, match_row in matches.iterrows():
            # Crear un nuevo DataFrame con las coincidencias encontradas
            nuevo_df = pd.DataFrame({
                "Sociedad": [row["Sociedad"]],
                "DETALLE": [row["DETALLE"]],
                "Prima": [row["Prima"]],
                "IGV": [row["IGV"]],
                "Prima total": [row["Prima total"]],
                "I": [row["I"]],
                "Grupo": [row["Grupo"]],
                "Anticipo": [row["Anticipo"]],
                "Incluye": [row["Incluye"]],
                "Proveedor": [row["Proveedor"]],
                "Cod. Proveedor": [row["Cod. Proveedor"]],
                "Grupo Artículo": [match_row["Grupo Artículo"]],
                "Codigo de Grupo Articulo": [match_row["Código de Grupo Artículo"]],
                "Centro": [match_row["Centro"]],
                "Codigo de Centro": [match_row["Código de Centro"]],
                "N. SERVICIO": [match_row["Servicio"]],
                "CECO": [match_row["CECO"]]
            })
            resultados2.append(nuevo_df)
    else:
        print(f"No se encontró grupo: {sociedad}, {detalle}")
        # Crear un DataFrame con campos vacíos
        nuevo_df_vacio = pd.DataFrame({
            "Sociedad": [row["Sociedad"]],
            "DETALLE": [row["DETALLE"]],
            "Prima": [row["Prima"]],
            "IGV": [row["IGV"]],
            "Prima total": [row["Prima total"]],
            "I": [row["I"]],
            "Grupo": [row["Grupo"]],
            "Anticipo": [row["Anticipo"]],
            "Incluye": [row["Incluye"]],
            "Proveedor": [row["Proveedor"]],
            "Cod. Proveedor": [""],  # Campo vacío
            "Grupo Artículo": [""],  # Campo vacío
            "Codigo de Grupo Articulo": [""],  # Campo vacío
            "Centro": [""],  # Campo vacío
            "Codigo de Centro": [""],  # Campo vacío
            "N. SERVICIO": [""],  # Campo vacío
            "CECO": [""]  # Campo vacío
        })
        resultados2.append(nuevo_df_vacio)
        print(f"No se encontró grupo: {sociedad} ", {detalle})

df_grupos_final= pd.concat(resultados2, ignore_index=True)
print(df_grupos_final)