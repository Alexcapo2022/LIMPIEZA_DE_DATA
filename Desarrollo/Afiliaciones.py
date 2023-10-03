import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

# Define la función que procesa el archivo Excel y crea el nuevo archivo Excel con las columnas requeridas.
def procesar_excel():
    try:
        ruta_excel_original = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])

        # Comprueba si se seleccionó un archivo.
        if not ruta_excel_original:
            return

        # Carga el archivo Excel original en un DataFrame.
        dataframe_original = pd.read_excel(ruta_excel_original, engine='openpyxl')

        # Filtra y selecciona las columnas requeridas, incluyendo APELLIDOS y SEXO.
        columnas_requeridas = ["COD ALUM", "ESPECIALIDAD","CORREO1","Periodo"]
        dataframe_final = dataframe_original[columnas_requeridas]

        # Combina las columnas APE PAT y APE MAT en una sola columna APELLIDOS.
        dataframe_final['APELLIDOS'] = dataframe_final['APE PAT'] + " " + dataframe_final['APE MAT']

        # Crea un nuevo DataFrame con las columnas necesarias.
        columnas_nuevo = ["COD ALUM", "PROGRAMA ACADEMICO", "CONTACTO", "Periodo", "Grado", "Situacion", "Alumno", "Año admisión", "Ciclo admisión", "ID afiliación"]
        dataframe_nuevo = pd.DataFrame(columns=columnas_nuevo)

        # Solicita al usuario ingresar el nombre del archivo antes de guardarlo.
        nombre_archivo = simpledialog.askstring("Nombre del archivo", "Ingrese el nombre del archivo sin extensión:")

        if nombre_archivo:
            # Abre un cuadro de diálogo para que el usuario seleccione la carpeta donde se guardará el nuevo archivo.
            carpeta_destino = filedialog.askdirectory()

            # Define la ruta completa donde se guardará el nuevo archivo Excel con el nombre proporcionado por el usuario.
            ruta_excel_nuevo = f'{carpeta_destino}\\{nombre_archivo}.xlsx'

            # Guarda el DataFrame final en un nuevo archivo Excel.
            dataframe_nuevo.to_excel(ruta_excel_nuevo, index=False, engine='openpyxl', header=True)

            messagebox.showinfo("Éxito", f'Data importante extraída y guardada en "{ruta_excel_nuevo}"')
        else:
            messagebox.showinfo("Información", "No se ha especificado un nombre de archivo. La operación se ha cancelado.")

    except Exception as e:
        messagebox.showerror("Error", f'Error al procesar el archivo Excel: {str(e)}')

# Crear la ventana principal de la interfaz gráfica.
root = tk.Tk()
root.title("Procesamiento de Excel y Creación de Nuevo Excel")

# Crear un botón que llame a la función para procesar y crear el nuevo archivo Excel.
procesar_button = tk.Button(root, text="Seleccionar, Procesar y Crear Excel", command=procesar_excel)
procesar_button.pack()

# Ejecutar el ciclo principal de la interfaz gráfica.
root.mainloop()


