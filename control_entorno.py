import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar
import os

# Función para convertir Excel a CSV
def convert_excel_to_csv(file_path, sheet_name):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    csv_path = f"{sheet_name}.csv"
    df.to_csv(csv_path, index=False)
    return csv_path

# Función para buscar un término en un archivo CSV
def search_in_csv(csv_path, search_term, relevant_columns):
    chunksize = 1000  # Procesar de a 1000 filas
    for chunk in pd.read_csv(csv_path, usecols=relevant_columns, chunksize=chunksize, dtype=str):
        # Verificar si alguna fila contiene el término de búsqueda
        found_row = chunk[chunk.astype(str).apply(lambda row: row.str.contains(search_term, case=False).any(), axis=1)]
        if not found_row.empty:
            return found_row
    return pd.DataFrame()  # Si no se encontró nada, devuelve un DataFrame vacío

# Función para mostrar los resultados en el cuadro de texto
def display_result(sheet_name, row, text_widget):
    text_widget.insert(tk.END, f"Dato encontrado en la hoja: {sheet_name}\n")
    if sheet_name == 'GCABA':
        result_text = f"CUIL: {row['CUIL'].values[0]}\nCARGO: {row['CARGO'].values[0]}\nAYN: {row['AYN'].values[0]}\n" \
                      f"COD_REP: {row['COD_REP'].values[0]}\nDESC_REP: {row['DESC_REP'].values[0]}\nMINISTERIO: {row['MINISTERIO'].values[0]}\nCAR_SIT_REV: {row['CAR_SIT_REV'].values[0]}\n"
    else:
        result_text = f"CUIL: {row['CUIL'].values[0]}\nCARGO: {row['CARGO'].values[0]}\nAYN: {row['AYN'].values[0]}\n" \
                      f"COD_REP: {row['COD_REP'].values[0]}\nDESC_REP: {row['DESC_REP'].values[0]}\nMINISTERIO: {row['MINISTERIO'].values[0]}\n"
    text_widget.insert(tk.END, result_text)

# Función para procesar el archivo y realizar la búsqueda
def process_file(file_path, search_term, text_widget, progress):
    # Limpiar el área de resultados
    text_widget.delete(1.0, tk.END)

    # Convertir hojas de Excel a CSV
    csv_gcaba = convert_excel_to_csv(file_path, 'GCABA')
    csv_pdc = convert_excel_to_csv(file_path, 'PDC')
    csv_ivc = convert_excel_to_csv(file_path, 'IVC')

    # Definir columnas relevantes por hoja
    columns_gcaba = ['CUIL', 'CARGO', 'AYN', 'COD_REP', 'DESC_REP', 'MINISTERIO', 'CAR_SIT_REV']
    columns_pdc_ivc = ['CUIL', 'CARGO', 'AYN', 'COD_REP', 'DESC_REP', 'MINISTERIO']

    # Lista de hojas y CSV asociados
    sheets_info = [
        ('GCABA', csv_gcaba, columns_gcaba),
        ('PDC', csv_pdc, columns_pdc_ivc),
        ('IVC', csv_ivc, columns_pdc_ivc)
    ]

    found = False  # Variable para verificar si se encontró el término

    # Iterar por cada hoja (CSV)
    for i, (sheet_name, csv_path, columns) in enumerate(sheets_info):
        # Actualizar la barra de progreso
        progress['value'] = (i + 1) * 33  # Suponiendo que hay 3 hojas, incrementamos en 33% por hoja
        root.update_idletasks()

        # Realizar la búsqueda en el archivo CSV
        result = search_in_csv(csv_path, search_term, columns)
        if not result.empty:
            found = True
            display_result(sheet_name, result, text_widget)
            break  # Si se encuentra en una hoja, detener la búsqueda

    if not found:
        text_widget.insert(tk.END, "Dato no encontrado en ninguna hoja.\n")

    # Finalizar la barra de progreso
    progress['value'] = 100
    root.update_idletasks()

# Función para seleccionar el archivo Excel
def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    return file_path

# Función para manejar el botón de búsqueda
def on_search():
    file_path = select_file()
    if not file_path:
        messagebox.showwarning("Archivo no seleccionado", "Por favor, selecciona un archivo Excel.")
        return

    search_term = entry_search.get().strip()
    if not search_term:
        messagebox.showwarning("Término no ingresado", "Por favor, ingresa un término de búsqueda.")
        return

    # Limpiar la barra de progreso antes de comenzar
    progress_bar['value'] = 0
    root.update_idletasks()

    # Iniciar el procesamiento y búsqueda
    process_file(file_path, search_term, text_output, progress_bar)

# Crear la interfaz gráfica
root = tk.Tk()
root.title("Búsqueda en Archivo Excel")
root.geometry("600x400")
root.configure(bg='#d3d3d3')

# Etiqueta y campo para ingresar el término de búsqueda
label_search = tk.Label(root, text="Ingresa el término de búsqueda (CUIL, DNI, AYN):", bg='#d3d3d3')
label_search.pack(pady=10)
entry_search = tk.Entry(root, width=50)
entry_search.pack(pady=5)

# Botón para iniciar la búsqueda
search_button = tk.Button(root, text="Buscar", command=on_search, bg='#808080', fg='white')
search_button.pack(pady=10)

# Barra de progreso
progress_bar = Progressbar(root, orient=tk.HORIZONTAL, length=400, mode='determinate')
progress_bar.pack(pady=10)

# Cuadro de texto para mostrar los resultados
text_output = tk.Text(root, height=10, width=70, bg='#f0f0f0')
text_output.pack(pady=20)

# Iniciar la ventana principal de Tkinter
root.mainloop()
