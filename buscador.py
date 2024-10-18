import os
import csv

def buscar_datos_en_csv(dato_busqueda, carpeta_csv, progress_bar):
    archivos_csv = [f for f in os.listdir(carpeta_csv) if f.endswith('.csv')]
    
    total_archivos = len(archivos_csv)
    progreso = 0
    
    for archivo in archivos_csv:
        archivo_ruta = os.path.join(carpeta_csv, archivo)
        
        with open(archivo_ruta, mode='r', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)
            
            for row in reader:
                # Si encontramos el dato en alguna columna de la fila
                if any(dato_busqueda.lower() in str(value).lower() for value in row.values()):
                    return f"Dato encontrado en la hoja: {archivo.split('.')[0]}"
        
        progreso += 1
        progress_bar['value'] = (progreso / total_archivos) * 100
        progress_bar.update()
    
    return "Dato no encontrado en ninguna hoja."
