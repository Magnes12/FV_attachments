import os
import PyPDF2
from openpyxl import Workbook


PATH = "C:\\Users\\daszyma\\Desktop\\Sprzedaż\\Eksport wysyłka\\Faktury"
files = os.listdir(PATH)
for file in files:
    if file[0] == "Z":
        pdf_dir = os.path.join(PATH, file)
        with open(pdf_dir, 'rb') as pdf_file:
            read_pdf = PyPDF2.PdfReader(pdf_file)
            number_of_pages = len(read_pdf.pages)
            page = read_pdf.pages[0]
            page_content = page.extract_text()
            if "VAT nr:" in page_content:
                start = page_content.find("VAT nr:") + len("VAT nr:")
                end = start + 9
                fv = page_content[start:end]
            if "Waga Netto" in page_content:
                start = page_content.find("Waga Netto") + len("Waga Netto") + 6
                end = start + 5
                weight = page_content[start:end]
                print(f'{fv} {weight}')
