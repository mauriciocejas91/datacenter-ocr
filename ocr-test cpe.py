import os
import re
import cv2
import numpy as np
import pytesseract
import fitz  # PyMuPDF
import ttkbootstrap as ttk
from tkinter import Label, Frame, filedialog, messagebox, END
from tkinter.scrolledtext import ScrolledText
from tkinterdnd2 import TkinterDnD, DND_FILES
from PIL import Image

# --- Configuración ---
base_directory = os.path.dirname(os.path.abspath(__file__))
output_folder = os.path.join(base_directory, "output")

if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# --- Procesamiento imagen ---
def deskew_fast(image: Image.Image) -> Image.Image:
    cv_img = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
    gray = cv2.cvtColor(cv_img, cv2.COLOR_BGR2GRAY)
    gray = cv2.bitwise_not(gray)
    thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
    coords = np.column_stack(np.where(thresh > 0))
    if len(coords) == 0: return image
    angle = cv2.minAreaRect(coords)[-1]
    if angle < -45: angle = -(90 + angle)
    else: angle = -angle
    if abs(angle) > 0.5:
        return image.rotate(angle, expand=True, resample=Image.BICUBIC, fillcolor=(255, 255, 255))
    return image

def preprocess_image(image):
    image = deskew_fast(image)
    img_array = np.array(image)
    gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
    binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)
    denoised = cv2.medianBlur(binary, 3)
    return Image.fromarray(denoised)

# --- Extracción unificada de datos ---

def extraer_todo(texto):
    # Diccionario con todos los campos solicitados para asegurar su visibilidad en la UI
    datos = {
        "Tipo Documento": "Desconocido",
        "Fecha de Comprobante": "No encontrado",
        "CTG": "No encontrado",
        "Pto. de Venta": "No encontrado",
        "Nro. Comprobante": "No encontrado",
        "CUIT Remitente": "No encontrado",
        "CUIT Destinatario": "No encontrado",
        "CUIT Destino": "No encontrado",
        "Base Imponible / Tarifa": "No encontrado"
    }

    # Lógica para Carta de Porte Electrónica (CPE)
    if "Carta de Porte" in texto or "CPE" in texto or "CTG" in texto:
        datos["Tipo Documento"] = "Carta de Porte Electrónica"
        
        # 1. Fecha [cite: 11, 13]
        f = re.search(r'Fecha:\s*(\d{2}/\d{2}/\d{4})', texto)
        if f: datos["Fecha de Comprobante"] = f.group(1)

        # 2. CTG 
        ctg = re.search(r'CTG:\s*(\d+)', texto)
        if ctg: datos["CTG"] = ctg.group(1)

        # 3. Punto de Venta y Nro CPE [cite: 12, 14]
        cpe = re.search(r'(?:N° CPE|CPE)[:\s]*(\d{5})-(\d{8})', texto)
        if cpe:
            datos["Pto. de Venta"] = cpe.group(1)
            datos["Nro. Comprobante"] = cpe.group(2)

        # 4. CUITs específicos [cite: 10, 18, 20]
        # Remitente [cite: 10]
        r = re.search(r'(?:Titular Carta de Porte|Remitente Comercial Productor)[:\s]*(\d{11})', texto)
        if r: datos["CUIT Remitente"] = r.group(1)

        # Destinatario [cite: 18]
        dt = re.search(r'Destinatario[:\s]*(\d{11})', texto)
        if dt: datos["CUIT Destinatario"] = dt.group(1)

        # Destino 
        ds = re.search(r'Destino[:\s]*(\d{11})', texto)
        if ds: datos["CUIT Destino"] = ds.group(1)

        # 5. Tarifa [cite: 56]
        t = re.search(r'Tarifa:\s*(\d+)', texto)
        if t: datos["Base Imponible / Tarifa"] = t.group(1)

    else:
        # Lógica para Facturas Estándar (A, B, C)
        datos["Tipo Documento"] = "Factura / Comprobante"
        fecha = re.search(r'(?:FECHA|Emisión|Fecha)[:\s]*(\d{2}[/-]\d{2}[/-]\d{4})', texto, re.IGNORECASE)
        if fecha: datos["Fecha de Comprobante"] = fecha.group(1)

        comp = re.search(r'(\d{4,5})-(\d{8})', texto)
        if comp:
            datos["Pto. de Venta"] = comp.group(1).lstrip('0') or "0"
            datos["Nro. Comprobante"] = comp.group(2)

        cuits = list(dict.fromkeys(re.findall(r'(\d{2}-?\d{8}-?\d{1})', texto)))
        if len(cuits) >= 1: datos["CUIT Remitente"] = cuits[0].replace("-", "")
        if len(cuits) >= 2: datos["CUIT Destinatario"] = cuits[1].replace("-", "")

        base = re.search(r'(?:Neto Gravado|Neto|Subtotal).*?[:\$]?\s*([\d\.,]+)', texto, re.IGNORECASE)
        if base: datos["Base Imponible / Tarifa"] = base.group(1)

    return datos

# --- Gestión de Interfaz y Procesamiento ---

def actualizar_pantalla(datos):
    text_display.config(state='normal')
    text_display.delete(1.0, END)
    text_display.insert(END, f"--- {datos['Tipo Documento']} ---\n\n", "bold")
    for campo, valor in datos.items():
        if campo != "Tipo Documento":
            text_display.insert(END, f"{campo}: ", "bold")
            text_display.insert(END, f"{valor}\n")
    text_display.config(state='disabled')

def procesar_y_guardar(file_path, text_completo, suffix=""):
    file_name = os.path.splitext(os.path.basename(file_path))[0]
    info = extraer_todo(text_completo)
    actualizar_pantalla(info)
    
    output_path = os.path.join(output_folder, f"{file_name}{suffix}.txt")
    contenido = f"=== DATOS EXTRAÍDOS ({info['Tipo Documento']}) ===\n"
    for campo, valor in info.items():
        contenido += f"{campo}: {valor}\n"
    contenido += "\n" + "-"*40 + "\n--- TEXTO COMPLETO ---\n" + text_completo
    
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(contenido)

def extract_text_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    progress_bar["maximum"] = len(doc)
    for i, page in enumerate(doc):
        zoom = 300 / 72
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        text = pytesseract.image_to_string(preprocess_image(img))
        procesar_y_guardar(pdf_path, text, f"_pag_{i+1}" if len(doc) > 1 else "")
        progress_bar["value"] = i + 1
        root.update_idletasks()
    doc.close()
    messagebox.showinfo("Proceso Completo", f"Se procesó: {os.path.basename(pdf_path)}")

def extract_text_from_png(png_path):
    image = Image.open(png_path)
    text = pytesseract.image_to_string(preprocess_image(image))
    procesar_y_guardar(png_path, text)
    messagebox.showinfo("Proceso Completo", f"Se procesó: {os.path.basename(png_path)}")

def on_drop(event):
    path = event.data.strip('{}')
    if path.lower().endswith('.pdf'): extract_text_from_pdf(path)
    elif path.lower().endswith(('.png', '.jpg', '.jpeg')): extract_text_from_png(path)

# --- Configuración de la Ventana (UI) ---
root = TkinterDnD.Tk()
root.title("DatacenterTDF | Extractor Inteligente CPE y Facturas")
root.geometry("700x700")
style = ttk.Style("superhero")

# Cabecera
header = Frame(root, bg="#343a40")
header.pack(fill="x")
Label(header, text="Arrastre archivos o use los botones de abajo", font=("Segoe UI", 12), bg="#343a40", fg="white").pack(pady=20)

# Botones (RESTAURADOS)
btn_frame = Frame(root, bg="#343a40")
btn_frame.pack(fill="x", pady=5)
# Botón para cargar archivos manualmente
ttk.Button(btn_frame, text="Cargar Archivo", 
           command=lambda: [extract_text_from_pdf(p) if p.lower().endswith('.pdf') else extract_text_from_png(p) 
                          for p in filedialog.askopenfilenames()]).pack(side="left", padx=20, pady=10)
# Botón para abrir la carpeta de resultados
ttk.Button(btn_frame, text="Abrir Carpeta de Salida", 
           command=lambda: os.startfile(output_folder)).pack(side="right", padx=20, pady=10)

# Barra de progreso
progress_bar = ttk.Progressbar(root, length=600, mode='determinate')
progress_bar.pack(pady=15)

# Área de Visualización
results_frame = ttk.LabelFrame(root, text=" Información Extraída del Documento ", padding=15)
results_frame.pack(padx=20, pady=10, fill="both", expand=True)

text_display = ScrolledText(results_frame, font=("Consolas", 11), state='disabled', bg="#1e1e1e", fg="#00ff00")
text_display.pack(fill="both", expand=True)
text_display.tag_config("bold", foreground="white", font=("Consolas", 11, "bold"))

root.drop_target_register(DND_FILES)
root.dnd_bind('<<Drop>>', on_drop)
root.mainloop()