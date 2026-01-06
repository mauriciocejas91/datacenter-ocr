import os
from tkinter import Label, Frame, filedialog, messagebox, Text
from tkinterdnd2 import TkinterDnD, DND_FILES
from PIL import Image
from pdf2image import convert_from_path
import ttkbootstrap as ttk
import torch
from transformers import TrOCRProcessor, VisionEncoderDecoderModel

# ----------------------
# Directories / Output
# ----------------------
base_directory = os.path.dirname(os.path.abspath(__file__))
output_folder = os.path.join(base_directory, "output")
os.makedirs(output_folder, exist_ok=True)

# ----------------------
# Load TrOCR Model
# ----------------------
device = "cuda" if torch.cuda.is_available() else "cpu"
processor = TrOCRProcessor.from_pretrained("microsoft/trocr-base-handwritten")
model = VisionEncoderDecoderModel.from_pretrained("microsoft/trocr-base-handwritten").to(device)
model.eval()

# ----------------------
# OCR Function
# ----------------------
def trocr_extract(image: Image.Image) -> str:
    """
    Extract text from an image using TrOCR and filter for 'Factura A' fields.
    """
    # Convert to RGB
    image = image.convert("RGB")

    # OCR inference
    pixel_values = processor(image, return_tensors="pt").pixel_values.to(device)
    generated_ids = model.generate(pixel_values)
    raw_text = processor.batch_decode(generated_ids, skip_special_tokens=True)[0]

    # Extract required fields
    fields = [
        "Fecha de Comprobante (*)",
        "Pto. de Venta (*)",
        "Nro. Comprobante (*)",
        "CUIT Remitente (*)",
        "CUIT Destinatario (*)",
        "Base Imponible (*)"
    ]

    # Prepare output dictionary
    output_lines = []
    for field in fields:
        # Simple extraction: look for lines containing the field or approximate match
        value = ""
        for line in raw_text.splitlines():
            if field.split()[0] in line:  # crude match on first word
                value = line.split(":")[-1].strip()
                break
        output_lines.append(f"{field}: {value}")

    return "\n".join(output_lines)

# ----------------------
# Image Preprocessing
# ----------------------
def preprocess_image(image: Image.Image) -> Image.Image:
    return image.convert("RGB")

# ----------------------
# OCR for PNG
# ----------------------
def extract_text_from_png(png_path):
    file_name = os.path.splitext(os.path.basename(png_path))[0]
    image = Image.open(png_path)
    enhanced_image = preprocess_image(image)
    text = trocr_extract(enhanced_image)
    txt_filename = os.path.join(output_folder, f"{file_name}.txt")
    with open(txt_filename, "w", encoding="utf-8") as f:
        f.write(text)
    show_success_message(png_path)

# ----------------------
# OCR for PDF
# ----------------------
def extract_text_from_pdf(pdf_path):
    file_name = os.path.splitext(os.path.basename(pdf_path))[0]
    pages = convert_from_path(pdf_path, 300)

    progress_text.delete(1.0, "end")
    progress_text.insert("end", f"Procesando archivo: {file_name}\n")
    progress_text.insert("end", f"Total de páginas: {len(pages)}\n")

    progress_bar["value"] = 0
    progress_bar["maximum"] = len(pages)
    progress_bar.pack(pady=10)
    progress_text.pack(pady=10)
    root.update_idletasks()

    for page_num, page in enumerate(pages):
        enhanced_page = preprocess_image(page)
        text = trocr_extract(enhanced_page)
        txt_filename = os.path.join(output_folder, f"{file_name}_pagina_{page_num + 1}.txt")
        with open(txt_filename, "w", encoding="utf-8") as f:
            f.write(text)

        progress_bar["value"] = page_num + 1
        root.update_idletasks()
        progress_text.insert("end", f"Página {page_num + 1} procesada\n")
        progress_text.yview_scroll(1, "units")

    show_success_message(pdf_path)

# ----------------------
# Drag & Drop Handler
# ----------------------
def on_drop(event):
    file_path = event.data.strip("{}")
    if file_path.endswith(".pdf"):
        extract_text_from_pdf(file_path)
    elif file_path.endswith(".png"):
        extract_text_from_png(file_path)
    else:
        messagebox.showwarning("Tipo de archivo no soportado", "Solo PDF o PNG.")

# ----------------------
# Success Message
# ----------------------
def show_success_message(file_path):
    file_name = os.path.basename(file_path)
    result = messagebox.askyesno(
        "Proceso completado",
        f"El archivo '{file_name}' se procesó correctamente.\n¿Deseas abrir la carpeta de salida?"
    )
    if result:
        open_output_folder()

def open_output_folder():
    os.startfile(output_folder)

# ----------------------
# File Dialog Handlers
# ----------------------
def load_files_pdf():
    file_paths = filedialog.askopenfilenames(filetypes=[("Archivos PDF", "*.pdf")])
    for pdf_path in file_paths:
        extract_text_from_pdf(pdf_path)

def load_files_png():
    file_paths = filedialog.askopenfilenames(filetypes=[("Archivos PNG", "*.png")])
    for png_path in file_paths:
        extract_text_from_png(png_path)

# ----------------------
# GUI
# ----------------------
root = TkinterDnD.Tk()
root.title("Factura A OCR Extractor")
root.geometry("500x350")

screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
window_width = 500
window_height = 350
position_top = int((screen_height / 2) - (window_height / 2))
position_right = int((screen_width / 2) - (window_width / 2))
root.geometry(f"{window_width}x{window_height}+{position_right}+{position_top}")

style = ttk.Style()
style.theme_use("superhero")

label = Label(
    root,
    text="Cargar PDFs o PNGs de Factura A para extraer campos",
    font=("Segoe UI", 14),
    bg="#343a40",
    fg="white"
)
label.pack(pady=20)

frame = Frame(root, bg="#343a40")
frame.pack(pady=10)

pdf_button = ttk.Button(frame, text="Cargar PDFs", bootstyle="success", command=load_files_pdf)
pdf_button.grid(row=0, column=0, padx=10, pady=10)

png_button = ttk.Button(frame, text="Cargar PNGs", bootstyle="info", command=load_files_png)
png_button.grid(row=0, column=1, padx=10, pady=10)

output_button = ttk.Button(frame, text="Abrir carpeta de salida", bootstyle="light", command=open_output_folder)
output_button.grid(row=0, column=2, padx=10, pady=10)

progress_bar = ttk.Progressbar(root, length=400, mode="determinate")
progress_text = Text(root, height=6, width=50)

root.drop_target_register(DND_FILES)
root.dnd_bind("<<Drop>>", on_drop)

root.mainloop()

