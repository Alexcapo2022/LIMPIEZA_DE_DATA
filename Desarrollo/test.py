import pandas as pd
import tkinter as tk
from tkinter import messagebox

# Define la función que procesa el archivo CSV y muestra un mensaje en la interfaz gráfica.
def procesar_csv():
    ruta_csv = r'C:\Users\Alexander Cruz\OneDrive\Imágenes\Documentos\test.csv'
    
    try:
        dataframe = pd.read_csv(ruta_csv, sep=",", skiprows=0, header=0)
        dataframe = dataframe.dropna(axis=1, how='all')
        
        ruta_xlsx = r'C:\Users\Alexander Cruz\OneDrive\Imágenes\Documentos\test.xlsx'
        dataframe.to_excel(ruta_xlsx, index=False, header=True, engine='openpyxl')
        
        messagebox.showinfo("Éxito", f'Archivo {ruta_csv} procesado con éxito')
    except Exception as e:
        messagebox.showerror("Error", f'Error al procesar el archivo CSV: {str(e)}')

# Crear la ventana principal de la interfaz gráfica
root = tk.Tk()
root.title("Procesamiento de CSV")

# Crear un botón que llame a la función para procesar el archivo CSV
procesar_button = tk.Button(root, text="Procesar CSV", command=procesar_csv)
procesar_button.pack()

# Ejecutar el ciclo principal de la interfaz gráfica
root.mainloop()
