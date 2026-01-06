import os
import re
import cv2
import numpy as np
import pytesseract
import ttkbootstrap as ttk
from tkinter import Label, Frame, filedialog, messagebox, Text
from tkinterdnd2 import TkinterDnD, DND_FILES
from PIL import Image, ImageEnhance
from pdf2image import convert_from_path

# Configuración de rutas
base_directory = os.path.dirname(os.path.abspath(__file__))
output_folder = os.path.join(base_directory, "output")

if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Mejorar calidad de imagen

# Deskew usando Tesseract OSD

def deskew_with_tesseract(image: Image.Image) -> Image.Image:
    try:
        osd = pytesseract.image_to_osd(image, output_type=pytesseract.Output.DICT)
        angle = osd.get("rotate", 0)

        if angle != 0:
            return image.rotate(-angle, expand=True)

    except Exception:
        pass  # Si OSD falla, seguimos sin deskew

    return image


def preprocess_image(image):
    # 1. Deskew
    image = deskew_with_tesseract(image)

    # 2. Grayscale
    gray = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2GRAY)

    # 3. Binarización OTSU
    _, binary = cv2.threshold(
        gray, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU
    )

    # 4. Reducción de ruido
    denoised = cv2.medianBlur(binary, 3)

    # 5. Contraste
    enhanced_image = Image.fromarray(denoised)
    enhancer = ImageEnhance.Contrast(enhanced_image)

    return enhancer.enhance(2.0)

def extraer_datos_factura(texto):
    datos = {
        "Fecha de Comprobante": "No encontrado",
        "Pto. de Venta": "No encontrado",
        "Nro. Comprobante": "No encontrado",
        "CUIT Remitente": "No encontrado",
        "CUIT Destinatario": "No encontrado",
        "Base Imponible": "No encontrado"
    }

# 1. Fecha de Comprobante (Mejorado para detectar "Fecha de Emisión", "F. Emisión", etc.)
    # Este regex busca palabras clave, permite "de" opcionalmente y maneja varios separadores
    patrones_fecha = [
        # Busca: Fecha de Emisión: 01/12/2025 o FECHA: 29/11/2025
        r'(?:FECHA|Emisión|Fecha)(?:\s+de)?(?:\s+Emisión)?[:\s]*(\d{2}[/-]\d{2}[/-]\d{4})',
        # Fallback: Solo busca la primera fecha que parezca dd/mm/aaaa si lo anterior falla
        r'(\d{2}[/-]\d{2}[/-]\d{4})'
    ]
    
    for patron in patrones_fecha:
        fecha_match = re.search(patron, texto, re.IGNORECASE)
        if fecha_match:
            datos["Fecha de Comprobante"] = fecha_match.group(1)
            break

    # 2. Pto de Venta y Nro Comprobante (Mejorado para "Comp. Nro.")
    # Intenta primero el formato unido XXXXX-XXXXXXXX
    comp_unido = re.search(r'(\d{4,5})-(\d{8})', texto)
    if comp_unido:
        datos["Pto. de Venta"] = comp_unido.group(1).lstrip('0') or "0"
        datos["Nro. Comprobante"] = comp_unido.group(2)
    else:
        # Busca etiquetas separadas incluyendo "Comp. Nro."
        pv_match = re.search(r'(?:Punto de Venta|P.V.|P.Venta)[:\s]*(\d+)', texto, re.IGNORECASE)
        nro_match = re.search(r'(?:Comp\.?\s*Nro\.?|Comprobante\s*Nro\.?|Nro\.?\s*Comprobante)[:\s]*(\d+)', texto, re.IGNORECASE)
        
        if pv_match: datos["Pto. de Venta"] = pv_match.group(1).lstrip('0') or "0"
        if nro_match: datos["Nro. Comprobante"] = nro_match.group(1)

    # 3. Lógica de CUITs (Mejorada para evitar duplicados de Ingresos Brutos) 
    # Extraemos todos los CUITs (formato XX-XXXXXXXX-X o XXXXXXXXXXX)
    cuits_encontrados = re.findall(r'(\d{2}-?\d{8}-?\d{1})', texto)
    
    # Limpiamos los CUITs para que todos tengan el mismo formato (solo números) para comparar
    cuits_limpios = [c.replace("-", "") for c in cuits_encontrados]
    
    # Identificamos el CUIT que acompaña a "Ingresos Brutos" para ignorarlo si es repetido
    ing_brutos_match = re.search(r'(?:Ingresos Brutos|Ing\.? Brutos|IIBB)[:\s]*(\d{2}-?\d{8}-?\d{1}|\d{9,11})', texto, re.IGNORECASE)
    cuit_iibb = ing_brutos_match.group(1).replace("-", "") if ing_brutos_match else None

    # Lista para CUITs únicos y relevantes (que no sean el de IIBB del emisor)
    cuits_finales = []
    for c in cuits_limpios:
        # Si el CUIT ya está en la lista o es el de IIBB y ya tenemos el del Remitente, lo saltamos
        if c not in cuits_finales:
            # En facturas como el ejemplo 4, el CUIT emisor y el IIBB son el mismo 
            cuits_finales.append(c)

    if len(cuits_finales) >= 1:
        datos["CUIT Remitente"] = cuits_finales[0]
    if len(cuits_finales) >= 2:
        # El destinatario será el primer CUIT distinto al del emisor 
        datos["CUIT Destinatario"] = cuits_finales[1]

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
    """Genera el .txt con los datos extraídos arriba para fácil copiado."""
    file_name = os.path.splitext(os.path.basename(file_path))[0]
    info = extraer_datos_factura(text_completo)
    
    output_path = os.path.join(output_folder, f"{file_name}{suffix}.txt")
    
    contenido = "=== DATOS EXTRAÍDOS (COPIAR AQUÍ) ===\n"
    for campo, valor in info.items():
        contenido += f"{campo}: {valor}\n"
    contenido += "="*40 + "\n\n"
    contenido += "--- TEXTO COMPLETO DEL OCR ---\n"
    contenido += text_completo
    
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(contenido)

# --- Funciones de procesamiento de archivos ---

def extract_text_from_png(png_path):
    image = Image.open(png_path)
    enhanced = preprocess_image(image)
    text = pytesseract.image_to_string(enhanced)
    procesar_y_guardar(png_path, text)
    show_success_message(png_path)

def extract_text_from_pdf(pdf_path):
    pages = convert_from_path(pdf_path, 300)
    progress_bar["maximum"] = len(pages)
    
    for i, page in enumerate(pages):
        enhanced = preprocess_image(page)
        text = pytesseract.image_to_string(enhanced)
        suffix = f"_pag_{i+1}" if len(pages) > 1 else ""
        procesar_y_guardar(pdf_path, text, suffix)
        
        progress_bar["value"] = i + 1
        root.update_idletasks()
        
    show_success_message(pdf_path)

# --- Interfaz Gráfica (Mantenida y adaptada) ---

def on_drop(event):
    path = event.data.strip('{}')
    if path.lower().endswith('.pdf'): extract_text_from_pdf(path)
    elif path.lower().endswith(('.png', '.jpg', '.jpeg')): extract_text_from_png(path)

def show_success_message(file_path):
    res = messagebox.askyesno("Proceso Exitoso", f"Se extrajeron los datos de:\n{os.path.basename(file_path)}\n\n¿Abrir carpeta de salida?")
    if res: os.startfile(output_folder)

root = TkinterDnD.Tk()
root.title("DatacenterTDF | Extractor Inteligente")
root.geometry("550x400")
style = ttk.Style("superhero")

Label(root, text="Arrastre Facturas (PDF/PNG) aquí", font=("Segoe UI", 14), bg="#343a40", fg="white").pack(pady=20)

btn_frame = Frame(root, bg="#343a40")
btn_frame.pack(pady=10)
ttk.Button(btn_frame, text="Cargar PDFs", command=lambda: [extract_text_from_pdf(p) for p in filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])]).grid(row=0, column=0, padx=5)
ttk.Button(btn_frame, text="Abrir Salida", command=lambda: os.startfile(output_folder)).grid(row=0, column=1, padx=5)

progress_bar = ttk.Progressbar(root, length=400, mode='determinate')
progress_bar.pack(pady=20)

root.drop_target_register(DND_FILES)
root.dnd_bind('<<Drop>>', on_drop)
root.mainloop()