import os
from tkinter import Label, Frame, filedialog, messagebox
from tkinterdnd2 import TkinterDnD
from PIL import Image
from pdf2image import convert_from_path
import pytesseract
import ttkbootstrap as ttk

# Obtener la ruta del directorio actual y concatenar 'output'
base_directory = os.path.dirname(os.path.abspath(__file__))
output_folder = os.path.join(base_directory, "output")

# Asegurarse de que el directorio de salida exista
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Definir la carpeta de salida
output_folder = os.path.join(base_directory, "output")

# Función para convertir PDF a imágenes y extraer texto
def extract_text_from_pdf(pdf_path):
    file_name = os.path.splitext(os.path.basename(pdf_path))[0]
    pages = convert_from_path(pdf_path, 300)
    for page_num, page in enumerate(pages):
        text = pytesseract.image_to_string(page)
        txt_filename = os.path.join(output_folder, f"{file_name}_pagina_{page_num + 1}.txt")
        with open(txt_filename, 'w', encoding='utf-8') as f:
            f.write(text)
        print(f"Texto de página {page_num + 1} guardado en {txt_filename}")

# Función para convertir PNG a texto
def extract_text_from_png(png_path):
    file_name = os.path.splitext(os.path.basename(png_path))[0]
    image = Image.open(png_path)
    text = pytesseract.image_to_string(image)
    txt_filename = os.path.join(output_folder, f"{file_name}.txt")
    with open(txt_filename, 'w', encoding='utf-8') as f:
        f.write(text)
    print(f"Texto de {file_name} guardado en {txt_filename}")

# Función para cargar archivos PDF
def load_files_pdf():
    file_paths = filedialog.askopenfilenames(filetypes=[("Archivos PDF", "*.pdf")])
    if file_paths:
        for pdf_path in file_paths:
            extract_text_from_pdf(pdf_path)
        messagebox.showinfo("Proceso completado", "Todos los archivos PDF se han procesado correctamente.")

# Función para cargar archivos PNG
def load_files_png():
    file_paths = filedialog.askopenfilenames(filetypes=[("Archivos PNG", "*.png")])
    if file_paths:
        for png_path in file_paths:
            extract_text_from_png(png_path)
        messagebox.showinfo("Proceso completado", "Todos los archivos PNG se han procesado correctamente.")

# Función para abrir la carpeta de salida
def open_output_folder():
    os.startfile(output_folder)

# Configuración de la interfaz gráfica
root = ttk.Window(themename="superhero")
root.title("Extractor de Texto OCR")
root.geometry("500x350")

# Etiqueta de encabezado
label = Label(root, text="Cargar archivos PDF o PNG para extraer texto", font=("Segoe UI", 14), bg="#343a40", fg="white")
label.pack(pady=20)

# Crear un marco para contener los botones
frame = Frame(root, bg="#343a40")
frame.pack(pady=10)

# Botón para cargar archivos PDF
pdf_button = ttk.Button(frame, text="Cargar PDFs", bootstyle="success", command=load_files_pdf)
pdf_button.grid(row=0, column=0, padx=10, pady=10)

# Botón para cargar archivos PNG
png_button = ttk.Button(frame, text="Cargar PNGs", bootstyle="info", command=load_files_png)
png_button.grid(row=0, column=1, padx=10, pady=10)

# Botón para abrir la carpeta de salida
output_button = ttk.Button(frame, text="Abrir carpeta de salida", bootstyle="light", command=open_output_folder)
output_button.grid(row=0, column=2, padx=10, pady=10)

# Iniciar la interfaz gráfica
root.mainloop()