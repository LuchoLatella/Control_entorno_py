import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd

def load_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        file_label.config(text=f"Archivo seleccionado: {file_path}")
        return file_path
    return None

def search_data():
    search_term = search_entry.get().strip()
    if not search_term:
        messagebox.showwarning("Error", "Por favor, ingresa un término de búsqueda.")
        return
    
    file_path = load_file()
    if not file_path:
        return

    try:
        # Leer las tres hojas del archivo
        sheets = pd.read_excel(file_path, sheet_name=['GCABA', 'PDC', 'IVC'])

        # Variables para almacenar resultados
        found_in_sheet = None
        found_row = None
        
        # Progreso
        progress_var.set(0)
        progress_bar.update_idletasks()

        # Búsqueda en la hoja GCABA
        def search_in_sheet(sheet_name, df, columns):
            nonlocal found_in_sheet, found_row
            for _, row in df.iterrows():
                if row.astype(str).str.contains(search_term, case=False, na=False).any():
                    found_in_sheet = sheet_name
                    found_row = row[columns]
                    break
        
        # Buscar en las hojas
        search_in_sheet('GCABA', sheets['GCABA'], ['CUIL', 'CARGO', 'AYN', 'COD_REP', 'DESC_REP', 'MINISTERIO', 'CAR_SIT_REV'])
        progress_var.set(33)
        progress_bar.update_idletasks()

        if not found_in_sheet:
            search_in_sheet('PDC', sheets['PDC'], ['CUIL', 'CARGO', 'AYN', 'COD_REP', 'DESC_REP', 'MINISTERIO'])
            progress_var.set(66)
            progress_bar.update_idletasks()

        if not found_in_sheet:
            search_in_sheet('IVC', sheets['IVC'], ['CUIL', 'CARGO', 'AYN', 'COD_REP', 'DESC_REP', 'MINISTERIO'])
            progress_var.set(100)
            progress_bar.update_idletasks()

        # Mostrar resultados
        if found_in_sheet and found_row is not None:
            result_label.config(text=f"Dato encontrado en la hoja: {found_in_sheet}\n{found_row}")
            result_label.config(fg='green')
        else:
            result_label.config(text="Dato no encontrado en ninguna hoja.")
            result_label.config(fg='red')
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al procesar el archivo.\n{e}")

# Interfaz gráfica con Tkinter
root = tk.Tk()
root.title("Buscador en Excel")
root.geometry("600x400")
root.configure(bg="#f0f0f0")  # Fondo gris claro

# Configuración de estilo
style = ttk.Style()
style.theme_use('clam')
style.configure("TButton", background="#d3d3d3", foreground="black")
style.configure("TLabel", background="#f0f0f0", foreground="black", font=('Helvetica', 12))
style.configure("TProgressbar", troughcolor="#d3d3d3", background="green")

# Widgets
file_button = ttk.Button(root, text="Seleccionar archivo", command=load_file)
file_button.pack(pady=10)

file_label = ttk.Label(root, text="No se ha seleccionado ningún archivo.")
file_label.pack(pady=5)

search_label = ttk.Label(root, text="Ingresar término de búsqueda:")
search_label.pack(pady=5)

search_entry = ttk.Entry(root, width=40)
search_entry.pack(pady=5)

search_button = ttk.Button(root, text="Buscar", command=search_data)
search_button.pack(pady=10)

# Barra de progreso
progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100)
progress_bar.pack(pady=10, fill=tk.X)

# Etiqueta de resultados
result_label = ttk.Label(root, text="Resultado aparecerá aquí.", wraplength=500)
result_label.pack(pady=20)

root.mainloop()
