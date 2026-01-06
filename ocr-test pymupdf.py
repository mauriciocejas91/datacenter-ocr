import os
import re
import cv2
import numpy as np
import pytesseract
import fitz  # PyMuPDF
import ttkbootstrap as ttk
from tkinter import Label, Frame, filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES
from PIL import Image

# --- Configuracion ---
base_directory = os.path.dirname(os.path.abspath(__file__))
output_folder = os.path.join(base_directory, "output")

if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# --- Procesamiento imagen ---

def deskew_fast(image: Image.Image) -> Image.Image:
    """Buscar tilt mediante CV2."""
    cv_img = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
    gray = cv2.cvtColor(cv_img, cv2.COLOR_BGR2GRAY)
    
    # Invert and threshold to find text blocks
    gray = cv2.bitwise_not(gray)
    thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
    
    coords = np.column_stack(np.where(thresh > 0))
    if len(coords) == 0:
        return image

    angle = cv2.minAreaRect(coords)[-1]
    
    # Normalizar angulo
    if angle < -45:
        angle = -(90 + angle)
    else:
        angle = -angle

    if abs(angle) > 0.5:
        return image.rotate(angle, expand=True, resample=Image.BICUBIC, fillcolor=(255, 255, 255))
    
    return image

def preprocess_image(image):
    """Refined preprocessing for OCR accuracy."""
    # 1. Deskew
    image = deskew_fast(image)
    
    # 2. Denoising
    img_array = np.array(image)
    gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
    
    # 3. Thresholding for clean text
    binary = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2
    )
    
    # 4. Reduccion de ruido
    denoised = cv2.medianBlur(binary, 3)
    return Image.fromarray(denoised)

# --- Extraccion de datos---

def extraer_datos_factura(texto):
    datos = {
        "Fecha de Comprobante": "No encontrado",
        "Pto. de Venta": "No encontrado",
        "Nro. Comprobante": "No encontrado",
        "CUIT Remitente": "No encontrado",
        "CUIT Destinatario": "No encontrado",
        "Base Imponible": "No encontrado"
    }

    # 1. Fecha
    patrones_fecha = [
        r'(?:FECHA|Emisión|Fecha)(?:\s+de)?(?:\s+Emisión)?[:\s]*(\d{2}[/-]\d{2}[/-]\d{4})',
        r'(\d{2}[/-]\d{2}[/-]\d{4})'
    ]
    for patron in patrones_fecha:
        fecha_match = re.search(patron, texto, re.IGNORECASE)
        if fecha_match:
            datos["Fecha de Comprobante"] = fecha_match.group(1)
            break

    # 2. Punto de Venta y Número
    comp_unido = re.search(r'(\d{4,5})-(\d{8})', texto)
    if comp_unido:
        datos["Pto. de Venta"] = comp_unido.group(1).lstrip('0') or "0"
        datos["Nro. Comprobante"] = comp_unido.group(2)
    else:
        pv_match = re.search(r'(?:Punto de Venta|P.V.|P.Venta)[:\s]*(\d+)', texto, re.IGNORECASE)
        nro_match = re.search(r'(?:Comp\.?\s*Nro\.?|Comprobante\s*Nro\.?|Nro\.?\s*Comprobante)[:\s]*(\d+)', texto, re.IGNORECASE)
        if pv_match: datos["Pto. de Venta"] = pv_match.group(1).lstrip('0') or "0"
        if nro_match: datos["Nro. Comprobante"] = nro_match.group(1)

    # 3. CUITs
    cuits_encontrados = re.findall(r'(\d{2}-?\d{8}-?\d{1})', texto)
    cuits_limpios = []
    for c in [c.replace("-", "") for c in cuits_encontrados]:
        if c not in cuits_limpios:
            cuits_limpios.append(c)

    if len(cuits_limpios) >= 1: datos["CUIT Remitente"] = cuits_limpios[0]
    if len(cuits_limpios) >= 2: datos["CUIT Destinatario"] = cuits_limpios[1]

    # 4. Base Imponible
    patrones_base = [
        r'(?:Neto Gravado|Neto|Subtotal|Gravado).*?[:\$]?\s*([\d\.,]+)',
        r'TOTAL NETO.*?[:\$]?\s*([\d\.,]+)'
    ]
    for patron in patrones_base:
        match = re.search(patron, texto, re.IGNORECASE)
        if match:
            valor = match.group(1).strip()
            if len(valor.replace(",", "").replace(".", "")) > 4:
                datos["Base Imponible"] = valor
                break
    return datos

def procesar_y_guardar(file_path, text_completo, suffix=""):
    file_name = os.path.splitext(os.path.basename(file_path))[0]
    info = extraer_datos_factura(text_completo)
    output_path = os.path.join(output_folder, f"{file_name}{suffix}.txt")
    
    contenido = "=== DATOS EXTRAÍDOS ===\n"
    for campo, valor in info.items():
        contenido += f"{campo}: {valor}\n"
    contenido += "="*40 + "\n\n" + "--- TEXTO COMPLETO ---\n" + text_completo
    
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(contenido)

# --- Main Logic ---

def extract_text_from_pdf(pdf_path):
    """Procesar pagianas mediante PyMuPDF."""
    doc = fitz.open(pdf_path)
    progress_bar["maximum"] = len(doc)
    
    for i, page in enumerate(doc):
        # Renderizar a 300 DPI
        zoom = 300 / 72
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        
        # Convertir a PIL y preprocesar
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        enhanced = preprocess_image(img)
        
        # Siempre OCR
        text = pytesseract.image_to_string(enhanced)
        
        suffix = f"_pag_{i+1}" if len(doc) > 1 else ""
        procesar_y_guardar(pdf_path, text, suffix)
        
        progress_bar["value"] = i + 1
        root.update_idletasks()
    
    doc.close()
    show_success_message(pdf_path)

def extract_text_from_png(png_path):
    image = Image.open(png_path)
    enhanced = preprocess_image(image)
    text = pytesseract.image_to_string(enhanced)
    procesar_y_guardar(png_path, text)
    show_success_message(png_path)

# --- UI ---

def on_drop(event):
    path = event.data.strip('{}')
    if path.lower().endswith('.pdf'): extract_text_from_pdf(path)
    elif path.lower().endswith(('.png', '.jpg', '.jpeg')): extract_text_from_png(path)

def show_success_message(file_path):
    res = messagebox.askyesno("Proceso Exitoso", f"Se extrajeron los datos de:\n{os.path.basename(file_path)}\n\n¿Abrir carpeta?")
    if res: os.startfile(output_folder)

root = TkinterDnD.Tk()
root.title("DatacenterTDF | Extractor OCR inteligente")
root.geometry("550x380")
style = ttk.Style("superhero")

Label(root, text="Arrastre Facturas (PDF/PNG) aquí", font=("Segoe UI", 13), bg="#343a40", fg="white").pack(pady=25)

btn_frame = Frame(root, bg="#343a40")
btn_frame.pack(pady=10)
ttk.Button(btn_frame, text="Cargar PDF", command=lambda: [extract_text_from_pdf(p) for p in filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])]).grid(row=0, column=0, padx=5)
ttk.Button(btn_frame, text="Abrir carpeta de salida...", command=lambda: os.startfile(output_folder)).grid(row=0, column=1, padx=5)

progress_bar = ttk.Progressbar(root, length=400, mode='determinate')
progress_bar.pack(pady=30)

root.drop_target_register(DND_FILES)
root.dnd_bind('<<Drop>>', on_drop)
root.mainloop()