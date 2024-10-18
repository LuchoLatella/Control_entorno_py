import os
import pandas as pd

def convertir_excel_a_csv(archivo_excel, carpeta_salida):
    if not os.path.exists(carpeta_salida):
        os.makedirs(carpeta_salida)
    
    # Leer el archivo Excel
    xls = pd.ExcelFile(archivo_excel)
    
    # Iterar por cada hoja y convertirla a CSV
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        archivo_csv = os.path.join(carpeta_salida, f"{sheet_name}.csv")
        df.to_csv(archivo_csv, index=False)
