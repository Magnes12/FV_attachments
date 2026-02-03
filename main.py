import re
import os
import sys
import ctypes
import pdfplumber
import pygetwindow as gw
import time
import itertools
from openpyxl import Workbook


def force_window_height():
    # Give the system a split second to actually spawn the window
    time.sleep(0.1)

    # Get the active window (which is this console)
    win = gw.getActiveWindow()

    if win:
        # Get monitor work area height using ctypes (to avoid Taskbar)
        user32 = ctypes.windll.user32
        # SPI_GETWORKAREA = 48
        rect = ctypes.wintypes.RECT()
        user32.SystemParametersInfoW(48, 0, ctypes.byref(rect), 0)

        work_height = rect.bottom - rect.top

        # Move to top-left and resize
        # width=800, height=work_height
        win.moveTo(0, 0)
        win.resizeTo(800, work_height)


def print_header():
    """Print application header."""
    header = """
╔══════════════════════════════════════════════════════════════╗
║          PDF Parser - Ekstraktor Danych z Faktur             ║
╚══════════════════════════════════════════════════════════════╝
"""
    print(header)


def print_separator(char="─", length=62):
    """Print a separator line."""
    print(char * length)


def extract_vat_and_weight(pdf_path):
    """Extract VAT number and weight from Z-prefixed PDFs."""
    vat_number = None
    weight = None

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()

                # Extract VAT number
                if "VAT nr:" in text and vat_number is None:
                    start = text.find("VAT nr:") + len("VAT nr:") + 1
                    vat_number = text[start:start + 8].strip()

                # Extract weight
                if "Waga Netto" in text and weight is None:
                    start = text.find("Waga Netto") + len("Waga Netto")
                    weight_str = text[start:start + 15].strip()

                    clean_str = re.sub(r'[^\d,]', '', weight_str)

                    # Convert to float
                    try:
                        # Replace comma with dot for decimal point
                        weight = float(clean_str.replace(",", "."))
                    except ValueError:
                        # If conversion fails, keep as string
                        weight = clean_str

                if vat_number and weight:
                    break

    except Exception as e:
        print(f"  ✗ Błąd w pliku {os.path.basename(pdf_path)}: {e}")

    return vat_number, weight


def extract_package(pdf_path):
    """Extract package number from 9-prefixed PDFs."""
    package = None

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()

                # Handle both "Paczka:" and "P aczka:" (with space)
                pattern = r'P\s*aczka:\s*(\d+)'
                match = re.search(pattern, text)

                if match:
                    package = match.group(1).strip()[-6:]
                    break

    except Exception as e:
        print(f"  ✗ Błąd w pliku {os.path.basename(pdf_path)}: {e}")

    return package


def main():
    try:
        force_window_height()
        # Print header
        time.sleep(1)
        print_header()

        current_dir = os.getcwd()
        files = os.listdir(current_dir)

        # Separate files by type
        nine_files = sorted([f for f in files if f.startswith("9") and f.lower().endswith(".pdf")])
        z_files = sorted([f for f in files if f.startswith("Z") and f.lower().endswith(".pdf")])

        # Display found files
        print("\n📄 ZNALEZIONE PLIKI PDF")
        print_separator()

        col_width = 30

        print(f"\n  {'Faktury (9*)':<{col_width}} {'Załączniki (Z*)':<{col_width}}")

        # zip_longest handles cases where one list is longer than the other
        # fillvalue="" prevents printing 'None' for missing entries
        for f, z in itertools.zip_longest(nine_files, z_files, fillvalue=""):
            f_display = f"• {f}" if f else ""
            z_display = f"• {z}" if z else ""

            # Left-align the first column to the specified width
            print(f"    {f_display:<{col_width}} {z_display}")

        # --- WARNINGS ---
        if not nine_files or not z_files:
            print("\n⚠ UWAGA: Nie znaleziono wszystkich wymaganych plików!")
            if not nine_files:
                print("  - Brak faktur")
            if not z_files:
                print("  - Brak załączników")

        print("\n🔍 PRZETWARZANIE PLIKÓW")
        print_separator()

        data_pack = []
        data_fv = []
        data_weight = []

        col_width = 30

        print(f"\n  {'Faktury (9*)':<{col_width}} {'Załączniki (Z*)':<{col_width}}")

        # Using zip_longest to process files in parallel for the two-column view
        for f_file, z_file in itertools.zip_longest(nine_files, z_files, fillvalue=None):
            f_status = ""
            z_status = ""

            # 1. Process Invoice (Left Column)
            if f_file:
                pdf_path = os.path.join(current_dir, f_file)
                package = extract_package(pdf_path)
                if package:
                    data_pack.append(package)
                    f_status = f"✓  {f_file}"
                else:
                    f_status = f"✗  {f_file} (brak nr)"

            # 2. Process Attachment (Right Column)
            if z_file:
                pdf_path = os.path.join(current_dir, z_file)
                vat, weight = extract_vat_and_weight(pdf_path)
                if vat and weight:
                    data_fv.append(vat)
                    data_weight.append(weight)
                    z_status = f"✓  {z_file}"
                else:
                    z_status = f"✗  {z_file} (błąd)"

            # 3. Print both statuses in one line
            # {f_status:<{col_width}} aligns the first column to the left
            print(f"    {f_status:<{col_width}} {z_status}")

        # Check data consistency
        print("\n📊 PODSUMOWANIE DANYCH")
        print_separator()
        print(f"  Numery VAT:      {len(data_fv)}")
        print(f"  Wagi:            {len(data_weight)}")
        print(f"  Numery paczek:   {len(data_pack)}")

        if not (len(data_fv) == len(data_weight) == len(data_pack)):
            print("\n  ⚠ OSTRZEŻENIE: Niezgodna liczba danych!")
        else:
            print("  ✓ Wszystkie dane kompletne")

        print_separator()

        # Create Excel file
        print("\n💾 TWORZENIE PLIKU EXCEL")
        print_separator()

        wb = Workbook()
        ws = wb.active
        ws.title = "Dane Faktur"

        # Header
        ws.append(["FV", "Waga", "Paczka"])

        # Data
        row_count = 0
        for fv, weight, pack in zip(data_fv, data_weight, data_pack):
            ws.append([f"00{fv}", weight, pack])
            row_count += 1

        file_name = "fv_waga.xlsx"
        wb.save(file_name)

        print(f"  ✓ Plik zapisany: {file_name}")
        print(f"  ✓ Dodano wierszy: {row_count}")

        print_separator()

        # Final summary
        print("\n✅ ZAKOŃCZONO POMYŚLNIE")
        print(f"\n  📁 Plik wynikowy: {file_name}")
        print_separator()

        print("\n\nNaciśnij ENTER aby zakończyć i otworzyć plik...")
        input()

        print("📂 Otwieram plik Excel...")
        try:
            os.startfile(file_name)
        except Exception as e:
            print(f"❌ Błąd podczas otwierania pliku: {e}")

    except Exception as e:
        print(f"\n❌ BŁĄD KRYTYCZNY: {e}")
        print("\nNaciśnij ENTER aby zakończyć...")
        input()
        sys.exit(1)


if __name__ == "__main__":
    main()
