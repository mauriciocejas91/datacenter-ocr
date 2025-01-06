import os
from tkinter import Label, Frame, filedialog, messagebox, Text
from tkinterdnd2 import TkinterDnD, DND_FILES  # Solo importar lo necesario
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

# Función para convertir PDF a imágenes y extraer texto
def extract_text_from_pdf(pdf_path):
    file_name = os.path.splitext(os.path.basename(pdf_path))[0]
    pages = convert_from_path(pdf_path, 300)
    
    # Mostrar progreso de las páginas
    progress_text.delete(1.0, "end")  # Limpiar texto previo
    progress_text.insert("end", f"Procesando archivo: {file_name}\n")
    progress_text.insert("end", f"Total de páginas: {len(pages)}\n")
    
    # Actualizar la barra de progreso
    progress_bar["value"] = 0
    progress_bar["maximum"] = len(pages)
    progress_bar.pack(pady=10)
    progress_text.pack(pady=10)
    root.update_idletasks()

    for page_num, page in enumerate(pages):
        text = pytesseract.image_to_string(page)
        txt_filename = os.path.join(output_folder, f"{file_name}_pagina_{page_num + 1}.txt")
        with open(txt_filename, 'w', encoding='utf-8') as f:
            f.write(text)
        
        # Actualizar barra de progreso
        progress_bar["value"] = page_num + 1
        root.update_idletasks()
        
        progress_text.insert("end", f"Página {page_num + 1} procesada\n")
        progress_text.yview_scroll(1, "units")  # Desplazar hacia abajo para ver nuevas líneas

    show_success_message(pdf_path)  # Mostrar mensaje de éxito después de procesar PDF

# Función para convertir PNG a texto
def extract_text_from_png(png_path):
    file_name = os.path.splitext(os.path.basename(png_path))[0]
    image = Image.open(png_path)
    text = pytesseract.image_to_string(image)
    txt_filename = os.path.join(output_folder, f"{file_name}.txt")
    with open(txt_filename, 'w', encoding='utf-8') as f:
        f.write(text)
    
    show_success_message(png_path)  # Mostrar mensaje de éxito después de procesar PNG

# Función para manejar archivos arrastrados y soltados
def on_drop(event):
    file_path = event.data
    # Verificar si el archivo es PDF o PNG
    if file_path.endswith('.pdf'):
        extract_text_from_pdf(file_path)
    elif file_path.endswith('.png'):
        extract_text_from_png(file_path)
    else:
        messagebox.showwarning("Tipo de archivo no soportado", "Por favor, arrastre archivos PDF o PNG.")

# Función para mostrar el mensaje de éxito con el botón de abrir carpeta
def show_success_message(file_path):
    file_name = os.path.basename(file_path)
    result = messagebox.askyesno("Proceso completado", 
                                f"El archivo '{file_name}' se ha procesado correctamente.\n¿Deseas abrir la carpeta de salida?")
    if result:
        open_output_folder()

# Función para abrir la carpeta de salida
def open_output_folder():
    os.startfile(output_folder)

# Función para cargar archivos PDF
def load_files_pdf():
    file_paths = filedialog.askopenfilenames(filetypes=[("Archivos PDF", "*.pdf")])
    if file_paths:
        for pdf_path in file_paths:
            extract_text_from_pdf(pdf_path)

# Función para cargar archivos PNG
def load_files_png():
    file_paths = filedialog.askopenfilenames(filetypes=[("Archivos PNG", "*.png")])
    if file_paths:
        for png_path in file_paths:
            extract_text_from_png(png_path)

# Configuración de la interfaz gráfica
root = TkinterDnD.Tk()  # Usamos TkinterDnD.Tk directamente
root.title("Extractor de Texto OCR")
root.geometry("500x350")

# Configuración de estilo y tema
style = ttk.Style()  # Creamos el objeto style
style.theme_use("superhero")  # Usamos el tema superhero

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

# Crear una barra de progreso
progress_bar = ttk.Progressbar(root, length=400, mode='determinate')

# Crear un widget Text para mostrar el progreso
progress_text = Text(root, height=6, width=50)

# Registrar la ventana para aceptar archivos arrastrados
root.drop_target_register(DND_FILES)
root.dnd_bind('<<Drop>>', on_drop)

# Iniciar la interfaz gráfica
root.mainloop()