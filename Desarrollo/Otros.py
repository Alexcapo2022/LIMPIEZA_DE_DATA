import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

def procesar_excel_4():
    try:
        # Abre un cuadro de diálogo para seleccionar el archivo Excel original.
        ruta_excel_original = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
        
        # Comprueba si se seleccionó un archivo.
        if not ruta_excel_original:
            return
        
        # Carga el archivo Excel original en un DataFrame.
        dataframe_original = pd.read_excel(ruta_excel_original, engine='openpyxl')

        # Verifica si las columnas requeridas están presentes en el archivo de entrada.
        if not all(col in dataframe_original.columns for col in ["Código Ulima", "Nombre Completo", "Apellidos", "Correo electrónico"]):
            messagebox.showerror("Error", "El archivo de entrada no contiene todas las columnas requeridas.")
            return

        # Extrae los datos de las columnas requeridas.
        codigo_ulima = dataframe_original["Código Ulima"].astype(str)
        nombre_completo = dataframe_original["Nombre Completo"]
        apellidos = dataframe_original["Apellidos"]
        correo_electronico = dataframe_original["Correo electrónico"]

        # Crea un nuevo DataFrame con los datos extraídos.
        nuevo_dataframe = pd.DataFrame({
            "Código UL": codigo_ulima,
            "Nombre": nombre_completo,
            "Apellidos": apellidos,
            "Correo electrónico de contacto": correo_electronico
        })

        # Abre un cuadro de diálogo para que el usuario seleccione la carpeta donde se guardará el nuevo archivo.
        carpeta_destino = filedialog.askdirectory()

        # Solicita al usuario ingresar el nombre del archivo antes de guardarlo.
        nombre_archivo = simpledialog.askstring("Nombre del archivo", "Ingrese el nombre del archivo sin extensión:")

        if nombre_archivo:
            # Define la ruta completa donde se guardará el nuevo archivo Excel con el nombre proporcionado por el usuario.
            ruta_excel_nuevo = f'{carpeta_destino}\\{nombre_archivo}.xlsx'

            # Guarda el DataFrame final en un nuevo archivo Excel.
            nuevo_dataframe.to_excel(ruta_excel_nuevo, index=False, engine='openpyxl')

            messagebox.showinfo("Éxito", f'Datos procesados y guardados en "{ruta_excel_nuevo}"')
        else:
            messagebox.showinfo("Información", "No se ha especificado un nombre de archivo. La operación se ha cancelado.")

    except Exception as e:
        messagebox.showerror("Error", f'Error al procesar el archivo Excel: {str(e)}')

# # Crear la ventana principal de la interfaz gráfica.
# root = tk.Tk()
# root.title("Procesamiento de Excel y Creación de Nuevo Excel")

# # Crear un botón que llame a la función para procesar y crear el nuevo archivo Excel.
# procesar_button = tk.Button(root, text="Seleccionar, Copiar y Crear Excel", command=procesar_excel_4)
# procesar_button.pack()

# # Ejecutar el ciclo principal de la interfaz gráfica.
# root.mainloop()




 

            

        
