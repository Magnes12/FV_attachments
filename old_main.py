import os
import sys
import pdfplumber
from openpyxl import Workbook


dir = './invoices'

files = os.listdir(dir)

wb = Workbook()
ws = wb.active

ws.append(["FV", "Waga", "Paczka"])

data_fv = []
data_weight = []
data_pack = []

print("Znalezione pliki PDF")

for file in files:
    if file[0] == "Z":
        print(file)
        pdf_dir = os.path.join(dir, file)
        with pdfplumber.open(pdf_dir) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if "VAT nr:" in text:
                    start = text.find("VAT nr:") + len("VAT nr:") + 1
                    end = start + 8
                    vat_number = text[start:end]

                    data_fv.append(vat_number)

                if "Waga Netto" in text:
                    start = text.find("Waga Netto") + len("Waga Netto") + 2
                    end = start + 9
                    weight = text[start:end]
                    if "." in weight:
                        weight_clean = weight.replace(".", "")
                        data_weight.append(weight_clean.strip())
                    else:
                        data_weight.append(weight.strip())

    elif file[0] == "9":
        print(file)
        pdf_dir = os.path.join(dir, file)
        pdf_dir = os.path.join(dir, file)
        with pdfplumber.open(pdf_dir) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if "Paczka:" in text:
                    start = text.find("Paczka:") + len("Paczka:") + 11
                    end = start + 6
                    package = text[start:end]

                    data_pack.append(package.strip())

for fv, weight, pack in zip(data_fv, data_weight, data_pack):
    ws.append([f"00{fv}", weight, pack])


file_name = "fv_waga.xlsx"


print("Excel zapisany")
wb.save(file_name)

print("Otwieram plik excel")
os.startfile(file_name)

print("Zamykam program")
sys.exit()
