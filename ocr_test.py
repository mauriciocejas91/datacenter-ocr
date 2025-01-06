import os
from tkinter import Tk, Button, Label, filedialog, messagebox, Frame
from tkinter import ttk
from tkinterdnd2 import TkinterDnD, DND_FILES
from PIL import Image
import pytesseract
from pdf2image import convert_from_path

# Obtener la ruta del directorio actual y concatenar 'output'
base_directory = os.path.dirname(os.path.abspath(__file__))
output_folder = os.path.join(base_directory, "output")

# Asegurarse de que el directorio de salida exista
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Función para convertir PDF a imágenes y extraer texto
def extract_text_from_pdf(pdf_path):
    # Obtener el nombre del archivo sin extensión
    file_name = os.path.splitext(os.path.basename(pdf_path))[0]
    
    # Convertir PDF a imágenes
    pages = convert_from_path(pdf_path, 300)
    
    # Procesar cada página
    for page_num, page in enumerate(pages):
        # Convertir la imagen a texto
        text = pytesseract.image_to_string(page)
        
        # Guardar el texto en un archivo .txt
        txt_filename = os.path.join(output_folder, f"{file_name}_pagina_{page_num + 1}.txt")
        with open(txt_filename, 'w', encoding='utf-8') as f:
            f.write(text)
        print(f"Texto de página {page_num + 1} guardado en {txt_filename}")

# Función para convertir PNG a texto
def extract_text_from_png(png_path):
    # Obtener el nombre del archivo sin extensión
    file_name = os.path.splitext(os.path.basename(png_path))[0]
    
    # Abrir la imagen PNG
    image = Image.open(png_path)
    
    # Extraer texto usando Tesseract
    text = pytesseract.image_to_string(image)
    
    # Guardar el texto en un archivo .txt
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

# Configuración de la interfaz gráfica
root = TkinterDnD.Tk()
root.title("Extractor de Texto OCR")
root.geometry("500x350")
root.config(bg="#2e2e2e")  # Dark background color

# Etiqueta de encabezado
label = Label(root, text="Cargar archivos PDF o PNG para extraer texto", font=("Segoe UI", 14), bg="#2e2e2e", fg="white")
label.pack(pady=20)

# Crear un marco para contener los botones
frame = Frame(root, bg="#2e2e2e")
frame.pack(pady=10)

# Estilo de botones moderno (con border-radius)
style = ttk.Style()
style.configure("Modern.TButton",
                font=("Segoe UI", 12),
                padding=10,
                relief="flat",
                background="#00CECC",
                foreground="#00CECC",
                focuscolor="none",
                borderwidth=0,
                anchor="center")

# Mapear el estado activo de los botones (hover/pressed)
style.map("Modern.TButton",
        background=[('active', '#45a049')],
        foreground=[('active', 'white')],
        relief=[('pressed', 'flat')])

# Botón para cargar archivos PDF
pdf_button = ttk.Button(frame, text="Cargar PDFs", style="Modern.TButton", command=load_files_pdf)
pdf_button.grid(row=0, column=0, padx=10, pady=10)

# Botón para cargar archivos PNG
png_button = ttk.Button(frame, text="Cargar PNGs", style="Modern.TButton", command=load_files_png)
png_button.grid(row=0, column=1, padx=10, pady=10)

# Iniciar la interfaz gráfica
root.mainloop()