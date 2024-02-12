import os
import PyPDF2
import time
from openpyxl import Workbook


PATH = "C:\\Users\\daszyma\\Desktop\\Sprzedaż\\Eksport wysyłka\\Faktury\\"

files = os.listdir(PATH)

wb = Workbook()
ws = wb.active

ws.append(["FV", "Waga", "Paczka"])

for file in files:
    if file[0] == "Z":
        pdf_dir = os.path.join(PATH, file)
        with open(pdf_dir, 'rb') as pdf_file:
            read_pdf = PyPDF2.PdfReader(pdf_file)
            number_of_pages = len(read_pdf.pages)
            page = read_pdf.pages[0]
            page_content = page.extract_text()
            if "VAT nr:" in page_content:
                start = page_content.find("VAT nr:") + len("VAT nr:") + 1
                end = start + 8
                fv = page_content[start:end]
            if "Waga Netto" in page_content:
                start = page_content.find("Waga Netto") + len("Waga Netto") + 1
                end = start + 11
                weight = page_content[start:end]
                print(weight)
            time.sleep(0.1)
            ws.append([f"00{fv}", weight, ""])

    elif file[0] == "9":
        pdf_dir = os.path.join(PATH, file)
        with open(pdf_dir, 'rb') as pdf_file:
            read_pdf = PyPDF2.PdfReader(pdf_file)
            number_of_pages = len(read_pdf.pages)
            page = read_pdf.pages[0]
            page_content = page.extract_text()
            if "Paczka:" in page_content:
                start = page_content.find("Paczka:") + len("Paczka:") + 11
                end = start + 6
                paczka = page_content[start:end]
            time.sleep(0.1)
            ws.append(["", "", paczka])


file_name = "fv_waga.xlsx"

wb.save(file_name)

os.system(file_name)
