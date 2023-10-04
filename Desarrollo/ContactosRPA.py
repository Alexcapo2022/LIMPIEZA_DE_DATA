import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import re
# Función para validar el formato del período ingresado por el usuario
def validar_periodo(periodo):
    patron = r'^\d{4}-(0|1|2)$'
    if re.match(patron, periodo):
        return True
    else:
        return False


# Define la función que procesa el archivo Excel y crea el nuevo archivo Excel con las columnas requeridas.
def procesar_excel():
    try:
        # Abre un cuadro de diálogo para seleccionar el archivo Excel original.
        ruta_excel_original = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])

        # Comprueba si se seleccionó un archivo.
        if not ruta_excel_original:
            return
        
        # Solicita al usuario ingresar el período (AAAA-P).
        periodo = simpledialog.askstring("Ingresar Período", "Ingrese el año y el período (AAAA-P) Ej: 2023-1")

        # Validar el formato del período ingresado por el usuario
        if not periodo or not validar_periodo(periodo):
            messagebox.showinfo("Información", "Formato de período inválido. Debe ser en el formato 'AAAA-P' (por ejemplo, '2023-1').")
            return

        # Carga el archivo Excel original en un DataFrame.
        dataframe_original = pd.read_excel(ruta_excel_original, engine='openpyxl')

        # Combina las columnas APE PAT y APE MAT en una sola columna APELLIDOS.
        dataframe_original['APELLIDOS'] = dataframe_original['APE PAT'] + " " + dataframe_original['APE MAT']

        # Reemplaza los valores "F" y "M" en la columna COD SEXO por "FEMENINO" y "MASCULINO".
        dataframe_original['COD SEXO'] = dataframe_original['COD SEXO'].replace({'F': 'FEMENINO', 'M': 'MASCULINO'})

        # Filtra y selecciona las columnas requeridas, incluyendo APELLIDOS y SEXO.
        columnas_requeridas = ["COD ALUM", "NOMBRE", "APELLIDOS", "COD SEXO", "TIPO DOC.IDEN.", "NRO DOC.IDEN.", 
                               "ESPECIALIDAD", "TELEFONO1", "TELEFONO2", "CORREO1"]
        dataframe_final = dataframe_original[columnas_requeridas]

        # Renombra la columna "COD SEXO" a "SEXO".
        dataframe_final.rename(columns={"COD SEXO": "SEXO"}, inplace=True)

        # Crea un nuevo DataFrame con las columnas necesarias, excluyendo la columna "COD SEXO".
        columnas_nuevo = ["COD ALUM", "APELLIDOS", "NOMBRE", "SEXO", "TIPO DOC.IDEN.", "NRO DOC.IDEN.", 
                           "ESPECIALIDAD", "TELEFONO1", "TELEFONO2", "CORREO1", "Correo electrónico secundario", 
                           "COD PS", "CORREO_UL", "Periodo", "Grado", "Situacion", "Nombre Favorito"]
        dataframe_nuevo = pd.DataFrame(columns=columnas_nuevo)

        # Combinar los DataFrames para llenar las similitudes y dejar los demás campos vacíos.
        resultado = pd.concat([dataframe_final, dataframe_nuevo], ignore_index=True)

        # Llena automáticamente las columnas "Grado" y "Situacion" con los valores especificados.
        resultado["Grado"].fillna("Pregrado", inplace=True)
        resultado["Situacion"].fillna("Alumno", inplace=True)

        # Copia los valores de la columna "COD PS" y "Nombre" al nuevo Excel.
        resultado["COD PS"] = dataframe_original["NRO DOC.IDEN."]
        resultado["Nombre Favorito"] = dataframe_original["NOMBRE"]

        # Llena la columna "Periodo" con el valor ingresado por el usuario.
        resultado["Periodo"] = periodo

        # Copia los valores de la columna "CORREO1" del Excel original a la columna "Correo electrónico secundario" del nuevo Excel.
        resultado["Correo electrónico secundario"] = dataframe_original["CORREO1"]

        # Crea la columna "CORREO1" utilizando la data de "COD ALUM" para generar los correos electrónicos en el formato adecuado.
        resultado["CORREO1"] = resultado["COD ALUM"].astype(str) + "@ALOE.ULIMA.EDU.PE"

        # Crea la columna "CORREO_UL" utilizando la data de "COD ALUM" para generar los correos electrónicos.
        resultado["CORREO_UL"] = resultado["COD ALUM"].astype(str) + "@ALOE.ULIMA.EDU.PE"

        # Abre un cuadro de diálogo para que el usuario seleccione la carpeta donde se guardará el nuevo archivo.
        carpeta_destino = filedialog.askdirectory()

        # Solicita al usuario ingresar el nombre del archivo antes de guardarlo.
        nombre_archivo = simpledialog.askstring("Nombre del archivo", "Ingrese el nombre del archivo sin extensión:")

        if nombre_archivo:
            # Define la ruta completa donde se guardará el nuevo archivo Excel con el nombre proporcionado por el usuario.
            ruta_excel_nuevo = f'{carpeta_destino}\\{nombre_archivo}.xlsx'

            # Guarda el DataFrame final en un nuevo archivo Excel.
            resultado.to_excel(ruta_excel_nuevo, index=False, engine='openpyxl')

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

