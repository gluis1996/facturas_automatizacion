import os
import pdfplumber
import pandas as pd

# Leer el archivo de rutas desde el Excel proporcionado
ruta_excel = r'C:\Users\user\Desktop\P_SCRIP\Em-ordenes-serv\rutas.xlsx'  
rutas_df = pd.read_excel(ruta_excel, header=None)  # Cargar el Excel sin encabezados

# Obtener las rutas desde el Excel por indices
ruta_principal = rutas_df.iloc[1, 0]
ruta_descarga_excel = rutas_df.iloc[2, 0]  # Fila 3, Columna A (index 2, 0)

# Ruta directa al archivo Excel con las sociedades
ruta_sociedades = r'C:\Users\user\Desktop\P_SCRIP\Em-ordenes-serv\Sociedades.xlsx'  

# Leer el archivo Excel con las sociedades
sociedades_df = pd.read_excel(ruta_sociedades)  # Cargar el Excel de sociedades
# Lista para almacenar los datos extraídos
datos = []

def extract_sociedad_codigo_from_text(text, sociedades_df):
    text = text.lower()  # Convertir el texto a minúsculas
    for _, row in sociedades_df.iterrows():  # Iterar sobre cada fila del DataFrame de sociedades
        sociedad = row['SOCIEDAD'].lower()  # Convertir el nombre de la sociedad a minúsculas
        if sociedad in text:  # Verificar si el nombre de la sociedad está en el texto
            return row['SOCIEDAD'], row['CÓDIGO']  # Devolver la sociedad y su código si se encuentra
    return "No identificado", "N/A"  # Devolver valores por defecto si no se identifica la sociedad

def extract_value_from_text(text, keyword):
    text = text.lower()  # Convertir el texto a minúsculas
    keyword = keyword.lower()  # Convertir la palabra clave a minúsculas
    
    start_index = text.find(keyword)  # Buscar la posición de la palabra clave en el texto
    if start_index != -1:
        substring = text[start_index + len(keyword):].strip()  # Obtener el texto que sigue a la palabra clave
        lines = substring.splitlines()  # Dividir el texto en líneas
        for line in lines:
            words = line.split()  # Dividir la línea en palabras
            for word in words:
                if word.replace(',', '').replace('.', '').isdigit():  # Verificar si la palabra es un número
                    return word  # Devolver el primer número encontrado
    return None  # Devolver None si no se encuentra un número

# Palabras clave para buscar en los PDFs
prima_keywords = ["op. gravada","Op.gravadas", "prima total", "prima comercial", "prima"]
igv_keywords = ["i.g.v.", "impsto.gral. a ventas", "igv", "impuesto general a las ventas"]
total_keywords = ["importe total", "total a cobrar", "total"]

def recorrer_carpetas_y_extraer_pdfs(ruta):
    for root, dirs, files in os.walk(ruta):  # Recorrer las carpetas y archivos en la ruta dada
        for file in files:
            if file.endswith('.pdf'):  # Procesar solo archivos PDF
                pdf_path = os.path.join(root, file)  # Obtener la ruta completa del archivo PDF
                print(f"Procesando archivo: {file}")  # Imprimir el nombre del archivo PDF que se está procesando
                try:
                    with pdfplumber.open(pdf_path) as pdf:  # Abrir el PDF con pdfplumber
                        all_text = ""
                        prima = None
                        igv = None
                        importe_total = None

                        for page in pdf.pages:  # Iterar sobre cada página del PDF
                            text = page.extract_text()  # Extraer el texto de la página
                            if text:
                                all_text += text  # Acumular el texto de todas las páginas

                                if prima is None:  # Buscar la prima solo si aún no se ha encontrado
                                    for keyword in prima_keywords:
                                        prima = extract_value_from_text(text, keyword)
                                        if prima:
                                            break

                                if igv is None:  # Buscar el IGV solo si aún no se ha encontrado
                                    for keyword in igv_keywords:
                                        igv = extract_value_from_text(text, keyword)
                                        if igv:
                                            break

                                if importe_total is None:  # Buscar el importe total solo si aún no se ha encontrado
                                    for keyword in total_keywords:
                                        importe_total = extract_value_from_text(text, keyword)
                                        if importe_total:
                                            break

                        if all_text:
                            sociedad, codigo = extract_sociedad_codigo_from_text(all_text, sociedades_df)  # Extraer la sociedad y el código
                            if sociedad != "No identificado":
                                # Añadir los datos extraídos a la lista
                                datos.append([file, sociedad, codigo, prima, igv, importe_total])

                except Exception as e:
                    print(f"Error al procesar el archivo {pdf_path}: {e}")  # Manejo de errores

# Ejecutar la función para recorrer carpetas y extraer datos de los PDFs
recorrer_carpetas_y_extraer_pdfs(ruta_principal)

# Crear un DataFrame con los datos extraídos
df = pd.DataFrame(datos, columns=["Archivo", "Sociedad", "Código", "Prima", "IGV", "Importe Total"])

# Guardar el DataFrame en la ruta especificada para el archivo Excel
output_excel_path = os.path.join(ruta_descarga_excel, 'reto2.xlsx')
df.to_excel(output_excel_path, index=False)  # Guardar el DataFrame en un archivo Excel sin índices

print(f"Proceso completado. Los datos han sido guardados en {output_excel_path}.")  # Mensaje de confirmación
