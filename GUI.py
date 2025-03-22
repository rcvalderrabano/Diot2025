import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import webbrowser
import os

class AplicacionConfiguracion(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("DIOT 2025 Convertidor Excel a Txt")
        self.geometry("500x200")

        # Esta línea coloca la imagen que uno desea
        icono = tk.PhotoImage(file=r'SAT100X100.png')
        self.iconphoto(True, icono)

        # Variables de control
        self.filepath = None  # Para almacenar la ruta del archivo seleccionado

        # Crear widgets de la interfaz
        self.crear_widgets()

    def crear_widgets(self):
        # Botón para seleccionar archivo
        btn_select_file = tk.Button(self, text="Seleccionar Archivo Excel", command=self.select_file)
        btn_select_file.grid(row=0, column=0, padx=10, pady=5)

        # Campo de texto para mostrar la ruta del archivo
        self.entry_ruta = tk.Entry(self, width=50, state="readonly")
        self.entry_ruta.grid(row=0, column=1, padx=10, pady=5)

        # Botón para mostrar los datos en formato HTML
        btn_mostrar_html = tk.Button(self, text="Ver Tabla en HTML", command=self.mostrar_tabla_html)
        btn_mostrar_html.grid(row=2, column=0, padx=10, pady=5)

        # Botón para generar el archivo TXT
        btn_generar_txt = tk.Button(self, text="Generar archivo TXT", command=self.generar_archivo_txt)
        btn_generar_txt.grid(row=2, column=1, padx=10, pady=5)

    def select_file(self):
        """Permite seleccionar un archivo Excel y almacena su ruta."""
        filepath = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx;*.xls")])

        if filepath:
            self.filepath = filepath
            self.entry_ruta.config(state="normal")  # Habilitar campo de texto para actualizarlo
            self.entry_ruta.delete(0, tk.END)
            self.entry_ruta.insert(0, filepath)
            self.entry_ruta.config(state="readonly")  # Volver a deshabilitarlo
            messagebox.showinfo("Archivo Seleccionado", f"Se ha seleccionado el archivo:\n{filepath}")

    def mostrar_tabla_html(self):
        """Convierte el archivo Excel a HTML y lo abre en el navegador."""
        if not self.filepath:
            messagebox.showerror("Error", "Por favor, selecciona un archivo Excel primero.")
            return

        try:
            # Leer el archivo Excel
            df = pd.read_excel(self.filepath)
            df.columns = [col.strip() for col in df.columns]  # Eliminar espacios en los nombres de las columnas

            # Reemplazar NaN por cadenas vacías
            df = df.fillna("")

            # Convertir el DataFrame a HTML
            html = df.to_html(index=False)  # index=False evita mostrar la columna de índice

            logo_path = "ADVAN100X100.png"

            # Agregar enlace al archivo CSS
            html = f"""
            <html>
            <head>
                <link rel="stylesheet" type="text/css" href="styles.css">
            </head>
            <body>
                <div class="table-container">
                    <h2>DIOT 2025</h2>
                    <table>
                        <caption>
                            <img src="{logo_path}" alt="Logo Empresa" width="100" height="100">
                        </caption>
                        {html}  <!-- Aquí insertamos la tabla generada -->
                    </table>
                </div>
            </body>
            </html>
            """

            # Guardar el HTML en un archivo temporal
            html_filename = "tabla.html"
            with open(html_filename, "w", encoding="utf-8") as f:
                f.write(html)

            # Abrir el archivo HTML en el navegador predeterminado
            webbrowser.open(f"file://{os.path.abspath(html_filename)}")

            messagebox.showinfo("HTML Generado", "La tabla ha sido generada en formato HTML y se ha abierto en el navegador.")

        except Exception as e:
            messagebox.showerror("Error", f"Error al leer el archivo Excel o generar el HTML:\n{str(e)}")

    def generar_archivo_txt(self):
        """Genera un archivo TXT con los datos del archivo Excel, separados por pipes (|)."""
        if not self.filepath:
            messagebox.showerror("Error", "Por favor, selecciona un archivo Excel primero.")
            return

        try:
            # Leer el archivo Excel
            df = pd.read_excel(self.filepath)
            df.columns = [col.strip() for col in df.columns]  # Eliminar espacios en los nombres de las columnas

            # Omitir la primera fila (índices de columnas) al guardar en el archivo TXT
            df = df.iloc[1:]

            # Reemplazar NaN por cadenas vacías
            df = df.fillna("")

            # Crear el nombre del archivo TXT basado en el nombre del archivo Excel
            txt_filename = os.path.splitext(os.path.basename(self.filepath))[0] + ".txt"

            # Guardar el DataFrame en el archivo TXT separado por pipes
            with open(txt_filename, "w", encoding="utf-8") as f:
                # Escribir las filas, separadas por pipe
                for index, row in df.iterrows():
                    f.write("|".join(str(value) for value in row) + "\n")

            messagebox.showinfo("Archivo TXT Generado", f"El archivo TXT ha sido generado exitosamente:\n{txt_filename}")

        except Exception as e:
            messagebox.showerror("Error", f"Error al leer el archivo Excel o generar el archivo TXT:\n{str(e)}")

# Ejecutar la aplicación
if __name__ == "__main__":
    app = AplicacionConfiguracion()
    app.mainloop()

