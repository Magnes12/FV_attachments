import os
from pypdf import PdfReader
from openpyxl import Workbook

current_dir = os.path.dirname(os.path.abspath(__file__))
PATH = os.path.dirname(current_dir)

files = os.listdir(PATH)

wb = Workbook()
ws = wb.active

ws.append(["FV", "Waga", "Paczka"])

data_fv = []
data_weight = []
data_pack = []

for file in files:
    if file[0] == "Z":
        pdf_dir = os.path.join(PATH, file)
        with open(pdf_dir, 'rb') as pdf_file:
            read_pdf = PdfReader(pdf_file)
            for page in read_pdf.pages:
                text = page.extract_text()
                if "VAT nr:" in text:
                    start = text.find("VAT nr:") + len("VAT nr:") + 1
                    end = start + 8
                    vat_number = text[start:end]

                    data_fv.append(vat_number)

                if "Waga Netto" in text:
                    start = text.find("Waga Netto") + len("Waga Netto") + 4
                    end = start + 7
                    weight = text[start:end]

                    data_weight.append(weight)

    elif file[0] == "9":
        pdf_dir = os.path.join(PATH, file)
        with open(pdf_dir, 'rb') as pdf_file:
            read_pdf = PdfReader(pdf_file)
            for page in read_pdf.pages:
                text = page.extract_text()
                if "Paczka:" in text:
                    start = text.find("Paczka:") + len("Paczka:") + 11
                    end = start + 6
                    package = text[start:end]

                    data_pack.append(package)

for fv, weight, pack in zip(data_fv, data_weight, data_pack):
    ws.append([f"00{fv}", weight, pack])


file_name = "fv_waga.xlsx"

wb.save(file_name)

os.system(file_name)
