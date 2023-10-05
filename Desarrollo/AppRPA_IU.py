import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from PIL import Image, ImageTk
import ContactosRPA
import Afiliaciones
import Otros

class InterfazRPA:
    def __init__(self, root):
        self.root = root
        ruta_icono = r"C:\Users\Alexander Cruz\OneDrive\Escritorio\zoho\RPA IMPORTANCIONES\Desarrollo\Logo.ico"
        root.iconbitmap(ruta_icono)
        self.root.title("Limpieza de Datos Ulima")
        self.root.geometry("300x500")
        self.root.configure(bg="#FCFAF5")

        # Sección 1
        self.crear_seccion("Contactos", self.funcion_contactos)

        # Separación vertical
        separador = tk.Label(self.root, bg="#FCFAF5", height=2)
        separador.pack(pady=10)

        # Sección 2
        self.crear_seccion("Afiliaciones", self.funcion_afiliaciones)

        # Separación vertical
        separador = tk.Label(self.root, bg="#FCFAF5", height=2)
        separador.pack(pady=10)

        # Sección 3
        self.crear_seccion("Otros", self.funcion_otros)

        # Imagen
        # Reemplaza 'ruta_de_la_imagen.png' con la ruta real de tu imagen PNG
        ruta_imagen = r'C:\Users\Alexander Cruz\OneDrive\Escritorio\zoho\RPA IMPORTANCIONES\Desarrollo\ulima.png'
        imagen = Image.open(ruta_imagen)
        imagen = ImageTk.PhotoImage(imagen)
        self.label_imagen = tk.Label(self.root, image=imagen, bg="#FCFAF5")
        self.label_imagen.image = imagen
        self.label_imagen.pack(pady=20)

        # Nombre del desarrollador
        self.nombre_desarrollador = tk.Label(self.root, text="Desarrollado por: ACM", font=("Arial", 10), bg="#FCFAF5")
        self.nombre_desarrollador.pack(pady=10)

    def crear_seccion(self, titulo, funcion):
        # Título de la sección
        label_titulo = tk.Label(self.root, text=titulo, font=("Arial", 14, "bold"), bg="#FCFAF5")
        label_titulo.pack()

        # Botón con la función asociada
        boton = tk.Button(self.root, text="Procesar", font=("Arial", 10), command=funcion)
        boton.pack()

    # Funciones para los botones
    def funcion_contactos(self):
        ContactosRPA.procesar_excel()

    def funcion_afiliaciones(self):
        Afiliaciones.procesar_excel_2()

    def funcion_otros(self):
        Otros.procesar_excel_4()

# Crear la ventana principal de la interfaz gráfica
root = tk.Tk()
app = InterfazRPA(root)

# Ejecutar el ciclo principal de la interfaz gráfica
root.mainloop()


