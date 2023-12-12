import os
import pandas as pd

def unir_xlsx(ruta_directorio, nombre_archivo_salida):
    # Obtener la lista de archivos en el directorio
    archivos_xlsx = [archivo for archivo in os.listdir(ruta_directorio) if archivo.endswith('.xlsx')]

    # Verificar si hay archivos para procesar
    if not archivos_xlsx:
        print("No hay archivos .xlsx en el directorio proporcionado.")
        return

    # Leer la primera hoja de cada archivo y almacenarla en un DataFrame
    frames = []
    for archivo in archivos_xlsx:
        ruta_completa = os.path.join(ruta_directorio, archivo)
        df = pd.read_excel(ruta_completa, sheet_name=0)  # Lee la primera hoja

        # Eliminar la columna 'A'
        df = df.iloc[:, 1:]  # Elimina la primera columna

        # Eliminar filas vacías después de eliminar la columna 'A'
        df = df.dropna(how='all')

        # Agregar el DataFrame al listado
        frames.append(df)

    # Concatenar los DataFrames en uno solo
    resultado_final = pd.concat(frames, ignore_index=True)

    # Guardar el DataFrame combinado en un nuevo archivo Excel
    ruta_salida = os.path.join(ruta_directorio, nombre_archivo_salida)
    resultado_final.to_excel(ruta_salida, index=False)

    print(f"Archivos combinados y guardados en: {ruta_salida}")

# Especifica la ruta del directorio que contiene los archivos .xlsx y el nombre del archivo de salida
directorio = r"C:\Users\Usuario\Desktop\CYS\INVENTARIO_SUZUVAL\07-12-2023"
nombre_archivo_salida = R"C:\Users\Usuario\Desktop\CYS\INVENTARIO_SUZUVAL\07-12-2023\CONSOLIDADO\INVENTARIO_CONSOLIDADO_12-12.xlsx"

# Llama a la función para unir los archivos
unir_xlsx(directorio, nombre_archivo_salida)


