import os
import sys
from pypdf import PdfReader
from openpyxl import Workbook


dir = os.getcwd()

files = os.listdir(dir)

# wb = Workbook()
# ws = wb.active

# ws.append(["FV", "Waga", "Paczka"])

data_fv = []
data = []


print("Znalezione pliki PDF")

for file in files:
    if file[0] == "95":
        print(file)
        pdf_dir = os.path.join(dir, file)
        with open(pdf_dir, 'rb') as pdf_file:
            read_pdf = PdfReader(pdf_file)
            for page in read_pdf.pages:
                text = page.extract_text()
                if "Faktura VAT" in text:
                    start = text.find("Faktura VAT") + len("Faktura VAT") + 1
                    end = start + 8
                    vat_number = text[start:end]
                    print(vat_number)

                    data_fv.append(vat_number)

                if "8700" in text:
                    start = text.find("8700") + len("8700") + 4
                    end = start + 7
                    weight = text[start:end]

                    data.append(weight)

for fv, d in zip(data_fv, data):
    print(fv, d)
    # ws.append([f"00{fv}", weight, pack])


file_name = "fv_waga.xlsx"


print("Excel zapisany")
# wb.save(file_name)

print("Otwieram plik excel")
# os.system(file_name)

print("Zamykam program")
sys.exit()
