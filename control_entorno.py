import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import os
from conversor import convertir_excel_a_csv
from buscador import buscar_datos_en_csv

# Crear la ventana principal
root = tk.Tk()
root.title("Aplicación de Búsqueda en CSV")
root.geometry("500x400")
root.configure(bg="#f0f0f0")  # Fondo en gris claro

# Frame para cargar el archivo
frame_carga = tk.Frame(root, bg="#f0f0f0")
frame_carga.pack(pady=20)

label_carga = tk.Label(frame_carga, text="Subir archivo Excel:", bg="#f0f0f0")
label_carga.pack(side=tk.LEFT, padx=10)

def cargar_archivo_excel():
    archivo_excel = filedialog.askopenfilename(
        title="Seleccionar archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    
    if archivo_excel:
        convertir_excel_a_csv(archivo_excel, "biblioteca")
        messagebox.showinfo("Conversión Completa", "Los archivos CSV se guardaron en la carpeta 'biblioteca'.")

boton_cargar = tk.Button(frame_carga, text="Cargar Excel", command=cargar_archivo_excel)
boton_cargar.pack(side=tk.LEFT)

# Frame para la búsqueda
frame_busqueda = tk.Frame(root, bg="#f0f0f0")
frame_busqueda.pack(pady=20)

label_busqueda = tk.Label(frame_busqueda, text="Ingresar dato a buscar (CUIL/DNI/Apellido):", bg="#f0f0f0")
label_busqueda.pack()

entry_busqueda = tk.Entry(frame_busqueda, width=40)
entry_busqueda.pack(pady=10)

# Barra de progreso
progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
progress_bar.pack(pady=20)

def iniciar_busqueda():
    dato_busqueda = entry_busqueda.get()
    if not dato_busqueda:
        messagebox.showwarning("Dato vacío", "Por favor ingresa un dato para buscar.")
        return
    
    resultado = buscar_datos_en_csv(dato_busqueda, "biblioteca", progress_bar)
    messagebox.showinfo("Resultado de búsqueda", resultado)

# Botón para iniciar la búsqueda
boton_buscar = tk.Button(root, text="Buscar", command=iniciar_busqueda)
boton_buscar.pack(pady=10)

root.mainloop()
