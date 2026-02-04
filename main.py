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


def extract_package(pdf_path):
    """Extract package number from any PDF (primary: 9*, fallback: Z*)."""
    package = None
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                pattern = r'P\s*aczka:\s*(\d+)'
                match = re.search(pattern, text)
                if match:
                    package = match.group(1).strip()[-6:]
                    break
    except Exception as e:
        print(f"  ✗ Błąd w pliku {os.path.basename(pdf_path)}: {e}")
    return package


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
                    try:
                        weight = float(clean_str.replace(",", "."))
                    except ValueError:
                        weight = clean_str

                if vat_number and weight:
                    break
    except Exception as e:
        print(f"  ✗ Błąd w pliku {os.path.basename(pdf_path)}: {e}")
    return vat_number, weight


def main():
    try:
        force_window_height()

        time.sleep(1)
        print_header()

        current_dir = os.getcwd()
        files = os.listdir(current_dir)

        # Separate files by type
        nine_files = sorted([f for f in files if f.startswith("9") and f.lower().endswith(".pdf")])
        z_files    = sorted([f for f in files if f.startswith("Z") and f.lower().endswith(".pdf")])

        # ─── Display found files ────────────────────────────────────────
        print("\n📄 ZNALEZIONE PLIKI PDF")
        print_separator()

        col_width = 30
        print(f"\n  {'Faktury (9*)':<{col_width}} {'Załączniki (Z*)':<{col_width}}")

        for f, z in itertools.zip_longest(nine_files, z_files, fillvalue=""):
            f_display = f"• {f}" if f else ""
            z_display = f"• {z}" if z else ""
            print(f"  {f_display:<{col_width}} {z_display}")

        # ─── Warnings ───────────────────────────────────────────────────
        if not nine_files and not z_files:
            print("\n⚠ UWAGA: Nie znaleziono żadnych plików PDF!")
            print("\nNaciśnij ENTER aby zakończyć...")
            input()
            sys.exit(1)

        if not nine_files:
            print("\n⚠ UWAGA: Brak faktur (9*) — dane będą wyciągane tylko z załączników.")
        if not z_files:
            print("\n⚠ UWAGA: Brak załączników (Z*) — numery paczek będą szukane tylko w fakturach.")

        # ─── Processing ─────────────────────────────────────────────────
        print("\n\n🔍 PRZETWARZANIE PLIKÓW")
        print_separator()

        # Each row = one paired entry. Structure: (vat, weight, package)
        # Any of these can be None — that becomes an empty cell in Excel.
        rows = []

        print(f"\n  {'Faktury (9*)':<{col_width}} {'Załączniki (Z*)':<{col_width}}")

        for f_file, z_file in itertools.zip_longest(nine_files, z_files, fillvalue=None):
            f_status = ""
            z_status = ""

            vat = None
            weight = None
            package = None

            # ── 1. Process Invoice (9*) ─── extract package number ──────
            if f_file:
                pdf_path = os.path.join(current_dir, f_file)
                package = extract_package(pdf_path)
                f_status = f"✓ {f_file}" if package else f"✗ {f_file} (brak nr paczki)"

            # ── 2. Process Attachment (Z*) ─── extract VAT + weight ─────
            if z_file:
                pdf_path = os.path.join(current_dir, z_file)
                vat, weight = extract_vat_and_weight(pdf_path)
                z_status = f"✓ {z_file}" if (vat and weight) else f"✗ {z_file} (brak danych)"

            # ── 3. Fallback: no package from 9*? try Z* ─────────────────
            if package is None and z_file:
                pdf_path = os.path.join(current_dir, z_file)
                package = extract_package(pdf_path)
                if package:
                    # Update statuses to reflect the fallback
                    f_status = f"✗ {f_file} (brak nr)" if f_file else ""
                    z_status += " [paczka: fallback]"

            # ── 4. Print both statuses side by side ──────────────────────
            print(f"  {f_status:<{col_width}} {z_status}")

            # ── 5. Always append the row — None values → empty cells ─────
            rows.append((vat, weight, package))

        # ─── Summary ────────────────────────────────────────────────────
        print("\n\n📊 PODSUMOWANIE DANYCH")
        print_separator()

        count_vat = sum(1 for r in rows if r[0] is not None)
        count_weight = sum(1 for r in rows if r[1] is not None)
        count_package = sum(1 for r in rows if r[2] is not None)
        total = len(rows)

        print(f"  Wiersze razem  : {total}")
        print(f"  Numery VAT     : {count_vat}/{total}")
        print(f"  Wagi           : {count_weight}/{total}")
        print(f"  Numery paczek  : {count_package}/{total}")

        missing = total - min(count_vat, count_weight, count_package)
        if missing:
            print(f"\n  ⚠ {missing} wiersze mają braki — komórki zostawione puste")
        else:
            print("  ✓ Wszystkie dane kompletne")

        print_separator()

        # ─── Create Excel ───────────────────────────────────────────────
        print("\n💾 TWORZENIE PLIKU EXCEL")
        print_separator()

        wb = Workbook()
        ws = wb.active
        ws.title = "Dane Faktur"

        # Header
        ws.append(["FV", "Waga", "Paczka"])

        # Data — None stays as None → openpyxl writes empty cell
        row_count = 0
        for vat, weight, package in rows:
            fv_cell = f"00{vat}" if vat else None
            ws.append([fv_cell, weight, package])
            row_count += 1

        file_name = "fv_waga.xlsx"
        wb.save(file_name)

        print(f"  ✓ Plik zapisany : {file_name}")
        print(f"  ✓ Dodano wierszy: {row_count}")
        print_separator()

        # ─── Done ───────────────────────────────────────────────────────
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
