import os
import sys
import pandas as pd
from tkinter import Tk, filedialog

print("Iniciando script...")

# Oculta la ventana principal de Tkinter
root = Tk()
root.withdraw()
print("Ventana principal de Tkinter oculta.")

# Abre el diálogo para seleccionar carpeta
carpeta_csv = filedialog.askdirectory(title="Selecciona la carpeta con los CSV")
root.destroy()
print(f"Carpeta seleccionada: {carpeta_csv}")

if not carpeta_csv:
    print("No se seleccionó ninguna carpeta.")
    sys.exit()

archivo_excel_combinado = os.path.join(carpeta_csv, "resultados_combinados.xlsx")
archivo_excel_separado = os.path.join(carpeta_csv, "resultados_separados.xlsx")
print(f"Archivo Excel combinado destino: {archivo_excel_combinado}")
print(f"Archivo Excel separado destino: {archivo_excel_separado}")

csvs = sorted([f for f in os.listdir(carpeta_csv) if f.endswith('.csv')])
print(f"Archivos CSV encontrados: {csvs}")

if not csvs:
    print("No se encontraron archivos CSV en la carpeta seleccionada.")
    sys.exit()

# Para combinar todos los CSV en uno solo
dfs_combinados = []

# 1. Excel con cada archivo en una hoja distinta (nombres únicos)
with pd.ExcelWriter(archivo_excel_separado) as writer:
    nombres_hojas_usados = set()
    for i, csv in enumerate(csvs):
        nombre_base = os.path.splitext(csv)[0]
        # Trunca a 25 caracteres y añade un sufijo si es necesario
        nombre_hoja = nombre_base[:25]
        sufijo = 1
        nombre_final = nombre_hoja
        while nombre_final in nombres_hojas_usados:
            nombre_final = f"{nombre_hoja}_{sufijo}"
            sufijo += 1
        nombres_hojas_usados.add(nombre_final)

        ruta_csv = os.path.join(carpeta_csv, csv)
        print(f"Procesando archivo: {csv} (hoja: {nombre_final})")
        try:
            df = pd.read_csv(ruta_csv, sep=';', encoding='utf-8')
            print(f"Leído correctamente con UTF-8: {csv}")
        except UnicodeDecodeError:
            print(f"Error de codificación UTF-8 en {csv}, intentando con latin1...")
            df = pd.read_csv(ruta_csv, sep=';', encoding='latin1')
            print(f"Leído correctamente con latin1: {csv}")
        df.to_excel(writer, sheet_name=nombre_final, index=False)
        print(f"Guardado en Excel separado: {nombre_final}")
        df['__archivo_origen__'] = csv  # Añade columna con nombre de archivo
        dfs_combinados.append(df)
        print(f"Añadido a la lista de DataFrames combinados: {csv}")

# 2. Excel con todos los datos combinados en una sola hoja
if dfs_combinados:
    print("Combinando todos los DataFrames en uno solo para el Excel combinado...")
    df_final = pd.concat(dfs_combinados, ignore_index=True)
    with pd.ExcelWriter(archivo_excel_combinado) as writer:
        df_final.to_excel(writer, sheet_name="Combinado", index=False)
    print("Archivo Excel combinado guardado correctamente.")

print("Todos los CSV se han combinado en dos archivos Excel: uno con hojas separadas y otro con todos los datos combinados.")


