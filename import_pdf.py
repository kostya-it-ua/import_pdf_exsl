import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from pdf2image import convert_from_path
import pytesseract
import cv2
import numpy as np

# Налаштування шляху до Tesseract (для Windows)
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def preprocess_image(image):
    """Попередня обробка зображення для кращого розпізнавання"""
    gray = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2GRAY)  # Переводимо в чорно-біле
    _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)  # Бінаризація
    return thresh

def extract_text_from_scan(pdf_path):
    """Розпізнає текст зі сканованого PDF через OCR"""
    images = convert_from_path(pdf_path)  # Конвертуємо PDF у зображення
    extracted_text = []
    
    for img in images:
        processed_img = preprocess_image(img)  # Попередня обробка
        text = pytesseract.image_to_string(processed_img, lang="eng+rus")  # OCR
        extracted_text.append(text)

    return extracted_text

def pdf_to_excel(pdf_path, excel_path):
    """Перетворює PDF у Excel (включаючи скани)"""
    extracted_text = extract_text_from_scan(pdf_path)
    
    # Збереження тексту у DataFrame
    df = pd.DataFrame({"Text": extracted_text})
    df.to_excel(excel_path, index=False)

def select_files():
    root = tk.Tk()
    root.withdraw()
    
    pdf_path = filedialog.askopenfilename(title="Выберите PDF файл", filetypes=[("PDF files", "*.pdf")])
    if not pdf_path:
        return
    
    excel_path = filedialog.asksaveasfilename(title="Сохранить как", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not excel_path:
        return
    
    pdf_to_excel(pdf_path, excel_path)

if __name__ == "__main__":
    select_files()