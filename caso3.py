import os
import pdfplumber
import pandas as pd
import re

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
    text = re.sub(r'[^A-Z0-9\s]', '', text)  # Eliminar caracteres especiales, mantener solo letras mayúsculas, números y espacios
    text = re.sub(r'\s+', ' ', text)  # Eliminar espacios múltiples
    return text

def buscar_sociedad_y_codigo(text, sociedades_df):
    text = limpiar_texto(text)
    for _, row in sociedades_df.iterrows():
        sociedad = limpiar_texto(row['SOCIEDAD'])
        if sociedad in text:
            return row['SOCIEDAD'], row['CÓDIGO']
    return "No identificado", "N/A"

def obtener_valor_numerico(text, keyword):
    text = limpiar_texto(text)
    keyword = limpiar_texto(keyword)
    
    start_index = text.find(keyword)
    if start_index != -1:
        substring = text[start_index + len(keyword):].strip()
        lines = substring.splitlines()
        for line in lines:
            words = line.split()
            for word in words:
                if word.replace(',', '').replace('.', '').isdigit():
                    return word
    return None

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

prima_key = ["OP GRAVADA", "OP.GRAVADAS", "PRIMA TOTAL", "PRIMA COMERCIAL", "PRIMA", "Op. Gravadas"]
igv_key = ["IGV", "IMPSTO.GRAL. A VENTAS", "IGV", "IMPUESTO GENERAL A LAS VENTAS"]
total_key = ["IMPORTE TOTAL", "TOTAL A COBRAR", "TOTAL", 'Importe Total:']

def recorrer_carpetas_y_extraer_pdfs(ruta):
    for root, dirs, files in os.walk(ruta):
        for file in files:
            if file.endswith('.pdf'):
                pdf_path = os.path.join(root, file)
                print('#############################################################cls')
                print(f"Procesando archivo: {file}")
                try:
                    with pdfplumber.open(pdf_path) as pdf:
                        all_text = ""
                        prima = None
                        igv = None
                        importe_total = None

                        for page in pdf.pages:
                            text = page.extract_text()
                            if text:
                                all_text += text

                                if prima is None:
                                    for keyword in prima_key:
                                        prima = obtener_valor_numerico(text, keyword)
                                        if prima:
                                            break

                                if igv is None:
                                    for keyword in igv_key:
                                        igv = obtener_valor_numerico(text, keyword)
                                        if igv:
                                            break

                                if importe_total is None:
                                    for keyword in total_key:
                                        importe_total = obtener_valor_numerico(text, keyword)
                                        if importe_total:
                                            break

                        # Imprimir el texto completo para depuración
                        print('#############################################################cls')
                        print("Texto extraído del PDF:")
                        print(all_text)
                        
                        # Imprimir el nombre del archivo PDF sin extensión y los valores extraídos
                        nombre_sin_extension = os.path.splitext(file)[0]
                        print(f"Nombre del archivo: {nombre_sin_extension}")
                        print(f"Prima: {prima}")
                        print(f"IGV: {igv}")
                        print(f"Importe Total: {importe_total}")

                        # Verificar si los datos son válidos y print OK o Fallo
                        if prima and igv and importe_total:
                            print("Estado del archivo: OK")
                        else:
                            print("Estado del archivo: Fallo")
                        
                        if all_text:
                            sociedad, codigo = buscar_sociedad_y_codigo(all_text, sociedades_df)
                            nuevo_nombre, proveedor, cod_proveedor, grupo_personal, imputacion, incluye = renombrar_pdf_y_validar(file, nuevos_nombres_df)
                            datos.append([file, sociedad, codigo, prima, igv, importe_total, nuevo_nombre, proveedor, cod_proveedor, grupo_personal, imputacion, incluye])

                except Exception as e:
                    print(f"Error al procesar el archivo {pdf_path}: {e}")

recorrer_carpetas_y_extraer_pdfs(ruta_principal)

df = pd.DataFrame(datos, columns=["Archivo", "Sociedad", "Código", "Prima", "IGV", "Importe Total", "DETALLE", "Proveedor", "Cod. Proveedor", "Grupo de Personal", "Imputacion", "Incluye"])
print(df)

