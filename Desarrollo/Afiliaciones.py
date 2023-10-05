import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

# Define la función que procesa el archivo Excel y crea el nuevo archivo Excel con las columnas requeridas.
def procesar_excel_2():
    try:
        # Abre un cuadro de diálogo para seleccionar el archivo Excel original.
        ruta_excel_original = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])

        # Comprueba si se seleccionó un archivo.
        if not ruta_excel_original:
            return

        # Carga el archivo Excel original en un DataFrame.
        dataframe_original = pd.read_excel(ruta_excel_original, engine='openpyxl')

        # Modifica los valores en la columna "ESPECIALIDAD" según las especificaciones dadas.
        dataframe_original["ESPECIALIDAD"] = dataframe_original["ESPECIALIDAD"].replace({
            "ING.CIVIL": "INGENIERIA CIVIL",
            "ING.INDUSTRIAL": "INGENIERIA INDUSTRIAL",
            "NEG. INTERNACIONALES": "NEGOCIOS INTERNACIONALES",
            "ING.SISTEMAS": "INGENIERIA DE SISTEMAS"
        })

        # Selecciona las columnas requeridas y crea un nuevo DataFrame con esas columnas.
        dataframe_final = dataframe_original[["COD ALUM", "ESPECIALIDAD", "CORREO1", "Periodo"]].copy()

        # Llena automáticamente las columnas "Grado" y "Situacion" con los valores especificados.
        dataframe_final["Grado"] = "PRE"
        dataframe_final["Situacion"] = "Alumno"
        dataframe_final["Alumno"] = "TRUE"

        # Llena las columnas "Año admisión" y "Ciclo admisión" basándose en el valor de la columna "Periodo".
        dataframe_final["Año admisión"] = dataframe_final["Periodo"].apply(lambda x: x.split('-')[0])
        dataframe_final["Ciclo admisión"] = dataframe_final["Periodo"].apply(lambda x: x.split('-')[1])

        # Crea la columna "ID afiliación" con los datos de las columnas "COD ALUM" y "ESPECIALIDAD".
        dataframe_final["ID afiliación"] = dataframe_final["COD ALUM"].astype(str) + "-" + dataframe_final["ESPECIALIDAD"]

        # Solicita al usuario ingresar el nombre del archivo antes de guardarlo.
        nombre_archivo = simpledialog.askstring("Nombre del archivo", "Ingrese el nombre del archivo sin extensión:")

        if nombre_archivo:
            # Abre un cuadro de diálogo para que el usuario seleccione la carpeta donde se guardará el nuevo archivo.
            carpeta_destino = filedialog.askdirectory()

            # Define la ruta completa donde se guardará el nuevo archivo Excel con el nombre proporcionado por el usuario.
            ruta_excel_nuevo = f'{carpeta_destino}\\{nombre_archivo}.xlsx'

            # Guarda el DataFrame final en un nuevo archivo Excel.
            dataframe_final.to_excel(ruta_excel_nuevo, index=False, engine='openpyxl', header=True, columns=["COD ALUM", "ESPECIALIDAD", "CORREO1", "Periodo", "Grado", "Situacion", "Alumno", "Año admisión", "Ciclo admisión", "ID afiliación"])

            messagebox.showinfo("Éxito", f'Data importante extraída y guardada en "{ruta_excel_nuevo}"')
        else:
            messagebox.showinfo("Información", "No se ha especificado un nombre de archivo. La operación se ha cancelado.")

    except Exception as e:
        messagebox.showerror("Error", f'Error al procesar el archivo Excel: {str(e)}')

# # Crear la ventana principal de la interfaz gráfica.
# root = tk.Tk()
# root.title("Procesamiento de Excel y Creación de Nuevo Excel")

# # Crear un botón que llame a la función para procesar y crear el nuevo archivo Excel.
# procesar_button = tk.Button(root, text="Seleccionar, Procesar y Crear Excel", command=procesar_excel_2)
# procesar_button.pack()

# # Ejecutar el ciclo principal de la interfaz gráfica.
# root.mainloop()






