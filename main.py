import re
import os
import sys
import time
import ctypes
import subprocess
import pdfplumber
import pygetwindow as gw
from openpyxl import Workbook


def get_sumatra_path():
    """Get path to SumatraPDF.exe (works in .exe and dev)"""
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))

    sumatra_path = os.path.join(base_path, 'SumatraPDF-3.5.2-64.exe')

    if not os.path.exists(sumatra_path):
        raise FileNotFoundError(
            f"SumatraPDF.exe nie znaleziony!\n"
            f"Oczekiwana lokalizacja: {sumatra_path}\n"
            f"Pobierz z: https://www.sumatrapdfreader.org/download-free-pdf-viewer"
        )

    return sumatra_path


def force_window_height():
    """
    Set console window in 0,0 posiotion and resize it to 800 weight
    """
    time.sleep(0.1)
    win = gw.getActiveWindow()
    if win:
        user32 = ctypes.windll.user32
        rect = ctypes.wintypes.RECT()
        user32.SystemParametersInfoW(48, 0, ctypes.byref(rect), 0)
        work_height = rect.bottom - rect.top
        win.moveTo(0, 0)
        win.resizeTo(800, work_height)


def print_header():
    """
    Print application header.
    """
    header = """
    ╔══════════════════════════════════════════════════════════════╗
    ║          PDF Parser - Ekstraktor Danych z Załączników        ║
    ╚══════════════════════════════════════════════════════════════╝
    """
    print(header)


def print_separator(char="─", length=62):
    """
    Print separator in application

    :param char: character which is used to make separator
    :param length: length of separator made by characters
    """
    print(char * length)


def get_files_paths(current_dir):
    """
    Get Z* (attachments) and 009* (invoices) PDF files

    :param current_dir: Current working directory
    """
    files = os.listdir(current_dir)
    z_files = sorted([f for f in files if f.startswith("Z") and f.lower().endswith(".pdf")])
    nine_files = sorted([f for f in files if f.startswith("009") and f.lower().endswith(".pdf")])
    return z_files, nine_files


def print_founded_files(z_files, nine_files, col_width=30):
    """
    Print found PDF files

    :param z_files: PDF files starting with 'Z'
    :param nine_files: PDF files starting with '009'
    :param col_width: Columne width
    """

    print("\n📄 ZNALEZIONE PLIKI PDF")
    print_separator()
    print(f"\n  {'Załączniki (Z*)':<{col_width}}")
    for f in z_files:
        display = f"• {f}" if f else ""
        print(f"  {display:<{col_width}}")

    if nine_files:
        print(f"\n  {'Faktury (9*) - do wydruku':<{col_width}}")
        for f in nine_files:
            print(f"  • {f}")

    if not z_files:
        print("\n⚠ UWAGA: Nie znaleziono żadnych plików Z*!")
        print("\nNaciśnij ENTER aby zakończyć...")
        input()
        sys.exit(1)


def extract_vat_package_weight(pdf_path):
    """
    Extract VAT, package number, and weight from Z* PDFs

    :param pdf_path: Direct path to pdf file
    """
    vat_number = None
    package = None
    weight = None
    net_value = None
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()

                # Extract VAT number
                if "VAT nr:" in text and vat_number is None:
                    start = text.find("VAT nr:") + len("VAT nr:") + 1
                    vat_number = text[start:start + 10].strip()

                # Extract package number
                if "Nr paczki :" in text and package is None:
                    start = text.find("Nr paczki :") + len("Nr paczki :") + 1
                    package = text[start:start + 12].strip()

                # Extract weight
                if "Waga Netto" in text and weight is None:
                    start = text.find("Waga Netto") + len("Waga Netto")
                    weight_str = text[start:start + 15].strip()
                    clean_str = re.sub(r'[^\d,]', '', weight_str)
                    try:
                        weight = float(clean_str.replace(",", "."))
                    except ValueError:
                        weight = clean_str

                # Extract weight
                if "Wartosc Netto" in text and net_value is None:
                    start = text.find("Wartosc Netto") + len("Wartosc Netto")
                    net_value_str = text[start:start + 15].strip()
                    clean_str = re.sub(r'[^\d,]', '', net_value_str)
                    try:
                        net_value = float(clean_str.replace(",", "."))
                    except ValueError:
                        net_value = clean_str

                if vat_number and package and weight and net_value:
                    break
    except Exception as e:
        print(f"  ✗ Błąd w pliku {os.path.basename(pdf_path)}: {e}")
    return vat_number, package, weight, net_value


def processing_founded_files(files, current_dir, col_width=30):
    """Process Z* files and extract data"""
    print("\n\n🔍 PRZETWARZANIE PLIKÓW")
    print_separator()

    rows = []
    print(f"\n  {'Załączniki (Z*)':<{col_width}}")

    for file in files:
        pdf_path = os.path.join(current_dir, file)
        vat, package, weight, net_value = extract_vat_package_weight(pdf_path)

        missing = []
        if not vat:
            missing.append("VAT")
        if not package:
            missing.append("Paczka")
        if not weight:
            missing.append("Waga")
        if not net_value:
            missing.append("Wartości netto")

        if not missing:
            status = f"✓ {file}"
        else:
            missing_str = ", ".join(missing)
            status = f"✗ {file} (brak: {missing_str})"

        print(f"  {status}")
        rows.append((vat, weight, package, net_value))

    return rows


def summary(rows):
    """Print data summary"""
    print("\n\n📊 PODSUMOWANIE DANYCH")
    print_separator()

    count_vat = sum(1 for r in rows if r[0] is not None)
    count_weight = sum(1 for r in rows if r[1] is not None)
    count_package = sum(1 for r in rows if r[2] is not None)
    count_net_value = sum(1 for r in rows if r[3] is not None)
    total = len(rows)

    print(f"  Wiersze razem  : {total}")
    print(f"  Numery VAT     : {count_vat}/{total}")
    print(f"  Wagi           : {count_weight}/{total}")
    print(f"  Numery paczek  : {count_package}/{total}")
    print(f"  Wartości netto : {count_net_value}/{total}")

    missing = total - min(count_vat, count_weight, count_package)
    if missing:
        print(f"\n  ⚠ {missing} wiersze mają braki — komórki zostawione puste")
    else:
        print("  ✓ Wszystkie dane kompletne")

    print_separator()


def excel_create(rows):
    """Create Excel file with extracted data"""
    print("\n💾 TWORZENIE PLIKU EXCEL")
    print_separator()

    wb = Workbook()
    ws = wb.active
    ws.title = "Dane Faktur"

    ws.append(["FV", "Waga", "Paczka", "Wartość"])

    row_count = 0
    for vat, weight, package, net_value in rows:
        fv_cell = f"{vat}" if vat else None
        ws.append([fv_cell, weight, package, net_value])
        row_count += 1

    file_name = "fv_waga.xlsx"
    wb.save(file_name)

    print(f"  ✓ Plik zapisany : {file_name}")
    print(f"  ✓ Dodano wierszy: {row_count}")
    print_separator()

    return file_name


def print_invoices_sequential(invoice_files, current_dir):
    print("\n\n🖨️  DRUKOWANIE FAKTUR")
    print_separator()

    try:
        sumatra_exe = get_sumatra_path()
    except Exception as e:
        print(f"  ❌ {e}")
        return

    for idx, filename in enumerate(invoice_files, 1):
        pdf_path = os.path.join(current_dir, filename)
        print(f"  [{idx}/{len(invoice_files)}] Drukuję: {filename}...", end=" ", flush=True)

        try:
            result = subprocess.run(
                [sumatra_exe,
                 "-print-to-default",
                 "-print-settings", "fit",
                 "-exit-when-done",
                 "-reuse-instance",
                 pdf_path
                 ],
                capture_output=True,
                timeout=60
            )

            if result.returncode == 0:
                print("✓")
            else:
                print(f"✗ (Kod: {result.returncode})")

            time.sleep(12)

        except Exception as e:
            print(f"✗ Błąd: {e}")

    print_separator()
    print("  ✅ Wydruk zakończony")


def main():
    try:
        force_window_height()

        current_dir = os.getcwd()
        time.sleep(1)

        print_header()

        z_files, nine_files = get_files_paths(current_dir)
        print_founded_files(z_files, nine_files)

        rows = processing_founded_files(z_files, current_dir)
        summary(rows)

        file_name = excel_create(rows)

        print("\n✅ ZAKOŃCZONO POMYŚLNIE")
        print(f"\n  📁 Plik wynikowy: {file_name}")
        print_separator()

        print("\n\nNaciśnij ENTER aby otworzyć plik Excel...")
        input()

        print("📂 Otwieram plik Excel...")
        try:
            os.startfile(file_name)
        except Exception as e:
            print(f"❌ Błąd podczas otwierania pliku: {e}")

        # ═══════════════════════════════════════════════════════════════
        # PRINT INVOICES (009*) - Optional step
        # ═══════════════════════════════════════════════════════════════

        if nine_files:
            print("\n" + "═" * 62)
            print(f"\n📋 Znaleziono {len(nine_files)} faktur (009*)")

            response = input("\nCzy wydrukować faktury? [T/N]: ").strip().upper()

            if response in ('T', 'TAK', 'Y', 'YES'):
                print_invoices_sequential(nine_files, current_dir)
            else:
                print("\n  ⏭️  Pominięto drukowanie")
                print_separator()
        else:
            print("\n  ℹ️  Brak faktur (009*) do wydruku")

        print("\n\nNaciśnij ENTER aby zakończyć...")
        input()

    except Exception as e:
        print(f"\n❌ BŁĄD KRYTYCZNY: {e}")
        print("\nNaciśnij ENTER aby zakończyć...")
        input()
        sys.exit(1)


if __name__ == "__main__":
    main()
