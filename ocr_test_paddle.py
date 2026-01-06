import os
import json
import re
import requests
import cv2
import numpy as np
from tkinter import Label, Frame, filedialog, messagebox, Text, Scrollbar, RIGHT, Y
from tkinterdnd2 import TkinterDnD, DND_FILES
from PIL import Image
from pdf2image import convert_from_path
import ttkbootstrap as ttk
from paddleocr import PaddleOCR

# ======================
# OLLAMA CONFIG
# ======================

OLLAMA_URL = "http://localhost:11434/api/generate"
OLLAMA_MODEL = "gemma3:4b"

PROMPT_TEMPLATE = """Output JSON only.
Use null if a field is missing.

Fields:
fecha_comprobante
pto_venta
nro_comprobante
cuit_remitente
cuit_destinatario
base_imponible

Text:
{ocr_text}
"""

# ======================
# OCR CONFIG (PaddleOCR)
# ======================

ocr = PaddleOCR(
    lang="es",
    use_textline_orientation=True
)

# ======================
# PATHS
# ======================

base_directory = os.path.dirname(os.path.abspath(__file__))
output_folder = os.path.join(base_directory, "output")

if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# ======================
# OCR FUNCTIONS
# ======================

def paddle_ocr_image(pil_image):
    img = cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2BGR)
    result = ocr.ocr(img)

    lines = []
    for block in result:
        for line in block:
            lines.append(line[1][0])

    return "\n".join(lines)

# ======================
# LLM FUNCTIONS
# ======================

def clean_json_from_llm(text):
    text = text.strip()

    if "```" in text:
        text = re.sub(r"```.*?```", lambda m: m.group(0).replace("```json", "").replace("```", ""), text, flags=re.S)

    match = re.search(r"\{.*\}", text, re.S)
    if not match:
        raise ValueError("No JSON object found")

    return match.group(0)

def extract_fields_with_llm(ocr_text):
    payload = {
        "model": OLLAMA_MODEL,
        "prompt": PROMPT_TEMPLATE.format(ocr_text=ocr_text),
        "stream": False
    }

    r = requests.post(OLLAMA_URL, json=payload, timeout=120)
    raw = r.json()["response"]
    cleaned = clean_json_from_llm(raw)

    return json.loads(cleaned)

# ======================
# DISPLAY
# ======================

def append_result(title, data):
    result_text.insert("end", f"\n--- {title} ---\n")
    for k, v in data.items():
        result_text.insert("end", f"{k.replace('_', ' ').title()}: {v}\n")
    result_text.yview_moveto(1)

# ======================
# FILE HANDLERS
# ======================

def extract_text_from_png(png_path):
    image = Image.open(png_path)
    text = paddle_ocr_image(image)

    data = extract_fields_with_llm(text)
    append_result(os.path.basename(png_path), data)

def extract_text_from_pdf(pdf_path):
    pages = convert_from_path(pdf_path, 300)

    for idx, page in enumerate(pages):
        ocr_text = paddle_ocr_image(page)
        data = extract_fields_with_llm(ocr_text)
        append_result(f"{os.path.basename(pdf_path)} - Página {idx + 1}", data)

# ======================
# UI CALLBACKS
# ======================

def on_drop(event):
    path = event.data.strip("{}")

    if path.lower().endswith(".pdf"):
        extract_text_from_pdf(path)
    elif path.lower().endswith(".png"):
        extract_text_from_png(path)
    else:
        messagebox.showwarning("Archivo no soportado", "Solo PDF o PNG")

def load_files_pdf():
    paths = filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])
    for p in paths:
        extract_text_from_pdf(p)

def load_files_png():
    paths = filedialog.askopenfilenames(filetypes=[("PNG", "*.png")])
    for p in paths:
        extract_text_from_png(p)

# ======================
# UI
# ======================

root = TkinterDnD.Tk()
root.title("DatacenterTDF | OCR Facturas")
root.geometry("700x500")

style = ttk.Style()
style.theme_use("superhero")

Label(
    root,
    text="Cargar PDF o PNG (OCR con PaddleOCR)",
    font=("Segoe UI", 14),
    bg="#343a40",
    fg="white"
).pack(pady=10)

frame = Frame(root, bg="#343a40")
frame.pack()

ttk.Button(frame, text="Cargar PDFs", bootstyle="success", command=load_files_pdf).grid(row=0, column=0, padx=5)
ttk.Button(frame, text="Cargar PNGs", bootstyle="info", command=load_files_png).grid(row=0, column=1, padx=5)

# Result box with scrollbar
result_frame = Frame(root)
result_frame.pack(fill="both", expand=True, pady=10)

scrollbar = Scrollbar(result_frame)
scrollbar.pack(side=RIGHT, fill=Y)

result_text = Text(result_frame, wrap="word", yscrollcommand=scrollbar.set)
result_text.pack(fill="both", expand=True)

scrollbar.config(command=result_text.yview)

root.drop_target_register(DND_FILES)
root.dnd_bind("<<Drop>>", on_drop)

root.mainloop()
