"""
PDF Serial Printer Script - Windows Edition
Prints PDF files from a directory in front/back order using Epson L130 (B&W).
2 PDF pages are merged side-by-side into one landscape sheet before printing.

FRONT: pages 1,2 | 5,6 | 9,10 | ...  (odd physical sheets)
BACK:  pages 3,4 | 7,8 | 11,12 | ... (even physical sheets)
       + a blank sheet appended if the PDF has an odd number of physical sheets,
         so the paper stack stays aligned for the next PDF.

Usage:
    python print_pdfs.py <directory> <front|back> [--printer "EPSON L130 Series"]
    python print_pdfs.py --list-printers

Requirements:
    pip install pypdf pywin32
    SumatraPDF (recommended): https://www.sumatrapdfreader.org/download-free-pdf-viewer
"""

import os
import sys
import math
import glob
import time
import shutil
import tempfile
import argparse
import subprocess

try:
    from pypdf import PdfReader, PdfWriter, PageObject, Transformation
except ImportError:
    print("ERROR: pypdf not found. Install it with:  pip install pypdf")
    sys.exit(1)

# ---------------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------------

PRINTER_NAME = "EPSON L130 Series"

SUMATRA_PATHS = [
    r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
    r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
    os.path.expanduser(r"~\AppData\Local\SumatraPDF\SumatraPDF.exe"),
]

# ---------------------------------------------------------------------------


def find_sumatra():
    for p in SUMATRA_PATHS:
        if os.path.isfile(p):
            return p
    return shutil.which("SumatraPDF")


def num_physical_sheets(total_pages):
    """Total physical sheets needed for a PDF (each sheet holds 2 pages)."""
    return math.ceil(total_pages / 2)


def get_front_pages(total_pages):
    """
    Pages printed on the FRONT side (odd-numbered physical sheets).
    Sheet 1 = pages 1,2 | Sheet 3 = pages 5,6 | Sheet 5 = pages 9,10 | ...
    """
    num_sheets = num_physical_sheets(total_pages)
    pages = []
    for sheet_idx in range(0, num_sheets, 2):   # 0-based: 0,2,4,...
        p1 = sheet_idx * 2 + 1
        p2 = p1 + 1
        pages.append(p1)
        if p2 <= total_pages:
            pages.append(p2)
    return pages


def get_back_pages(total_pages):
    """
    Pages printed on the BACK side (even-numbered physical sheets).
    Sheet 2 = pages 3,4 | Sheet 4 = pages 7,8 | Sheet 6 = pages 11,12 | ...
    """
    num_sheets = num_physical_sheets(total_pages)
    pages = []
    for sheet_idx in range(1, num_sheets, 2):   # 0-based: 1,3,5,...
        p1 = sheet_idx * 2 + 1
        p2 = p1 + 1
        pages.append(p1)
        if p2 <= total_pages:
            pages.append(p2)
    return pages


def needs_back_blank(total_pages):
    """
    Returns True if the back side needs a trailing blank sheet.

    A PDF needs a trailing blank on the back if it has an ODD number of
    physical sheets. Example:
      5 pages -> 3 physical sheets
      Front sheets: 1, 3        (2 sheets printed front)
      Back  sheets: 2           (1 sheet printed back)
      -> front count (2) != back count (1), so append 1 blank on back.

    In general: odd sheet count means front has one more sheet than back.
    """
    return num_physical_sheets(total_pages) % 2 == 1


def make_2up_sheet(left_page, right_page):
    """
    Merge two pages side-by-side into one landscape sheet.
    right_page may be None (blank right half).
    """
    src_w = float(left_page.mediabox.width)
    src_h = float(left_page.mediabox.height)
    sheet_w = src_w * 2
    sheet_h = src_h

    sheet = PageObject.create_blank_page(width=sheet_w, height=sheet_h)
    sheet.merge_transformed_page(left_page, Transformation().translate(0, 0))
    if right_page is not None:
        sheet.merge_transformed_page(right_page, Transformation().translate(src_w, 0))
    return sheet


def build_print_pdf(reader, pages, append_blank=False):
    """
    Build a temp PDF where each output page is a landscape 2-up sheet.
    If append_blank=True, one extra blank landscape sheet is added at the end.
    Returns path to the temp file.
    """
    writer = PdfWriter()

    # Get reference dimensions from the first page of this PDF
    ref_page = reader.pages[0]
    ref_w = float(ref_page.mediabox.width)
    ref_h = float(ref_page.mediabox.height)

    # Pair up selected pages into 2-up landscape sheets
    for i in range(0, len(pages), 2):
        left  = reader.pages[pages[i] - 1]
        right = reader.pages[pages[i + 1] - 1] if (i + 1) < len(pages) else None
        sheet = make_2up_sheet(left, right)
        writer.add_page(sheet)

    # Append a blank landscape sheet if needed
    if append_blank:
        blank = PageObject.create_blank_page(width=ref_w * 2, height=ref_h)
        writer.add_page(blank)
        print(f"    + Appended 1 blank sheet (odd sheet count — keeps stack aligned)")

    tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
    writer.write(tmp)
    tmp.close()
    return tmp.name


# ---------------------------------------------------------------------------
# PRINTING BACKENDS
# ---------------------------------------------------------------------------

def print_via_sumatra(pdf_path, printer, sumatra_exe):
    cmd = [
        sumatra_exe,
        "-print-to", printer,
        "-print-settings", "monochrome,fit",
        pdf_path,
    ]
    print(f"    [SumatraPDF] {' '.join(cmd)}")
    proc = subprocess.run(cmd, capture_output=True, text=True)
    time.sleep(2)
    if proc.returncode not in (0, 1):
        print(f"    WARNING: {proc.stderr.strip()}")
    else:
        print("    OK Job sent.")


def print_via_win32(pdf_path, printer):
    try:
        import win32api, win32print
    except ImportError:
        print("ERROR: pywin32 not found. Install:  pip install pywin32")
        sys.exit(1)
    original = win32print.GetDefaultPrinter()
    try:
        win32print.SetDefaultPrinter(printer)
        print(f"    [win32api] printto -> {printer}")
        win32api.ShellExecute(0, "printto", pdf_path, f'"{printer}"', ".", 0)
        time.sleep(5)
        print("    OK Job sent.")
    finally:
        win32print.SetDefaultPrinter(original)


def send_to_printer(pdf_path, printer, sumatra_exe):
    if sumatra_exe:
        print_via_sumatra(pdf_path, printer, sumatra_exe)
    else:
        print("    (SumatraPDF not found — using win32api fallback)")
        print_via_win32(pdf_path, printer)


# ---------------------------------------------------------------------------
# PER-FILE LOGIC
# ---------------------------------------------------------------------------

def print_pdf(pdf_path, side, printer, sumatra_exe):
    print(f"\n{'='*60}")
    print(f"File  : {os.path.basename(pdf_path)}")
    print(f"Side  : {side}")

    reader = PdfReader(pdf_path)
    total = len(reader.pages)
    num_sheets = num_physical_sheets(total)
    print(f"Pages : {total}  |  Physical sheets: {num_sheets}")

    if side == "front":
        pages = get_front_pages(total)
        append_blank = False   # front never needs a trailing blank
    else:
        pages = get_back_pages(total)
        append_blank = needs_back_blank(total)

    if not pages:
        # No back pages at all (1-page PDF) — still need blank if front printed one sheet
        if side == "back" and num_sheets >= 1:
            print("  No back pages — printing blank sheet to keep stack aligned.")
            ref = reader.pages[0]
            rw, rh = float(ref.mediabox.width), float(ref.mediabox.height)
            writer = PdfWriter()
            writer.add_page(PageObject.create_blank_page(width=rw * 2, height=rh))
            tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
            writer.write(tmp)
            tmp.close()
            try:
                send_to_printer(tmp.name, printer, sumatra_exe)
            finally:
                time.sleep(1)
                try: os.unlink(tmp.name)
                except: pass
        else:
            print("  Nothing to print for this side. Skipping.")
        return

    print(f"PDF pages selected : {pages}")
    output_sheets = math.ceil(len(pages) / 2) + (1 if append_blank else 0)
    print(f"Output sheets      : {output_sheets}" +
          (" (includes 1 blank)" if append_blank else ""))

    tmp_path = build_print_pdf(reader, pages, append_blank=append_blank)
    try:
        send_to_printer(tmp_path, printer, sumatra_exe)
    finally:
        time.sleep(1)
        try:
            os.unlink(tmp_path)
        except Exception:
            pass


# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------

def list_printers():
    try:
        import win32print
        printers = [p[2] for p in win32print.EnumPrinters(2)]
        print("Available printers:")
        for p in printers:
            print(f"  - {p}")
    except ImportError:
        print("Install pywin32 to list printers:  pip install pywin32")


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def main():
    if "--list-printers" in sys.argv:
        list_printers()
        sys.exit(0)

    parser = argparse.ArgumentParser(
        description="Print PDFs front/back with 2 pages per landscape sheet (Epson L130, B&W).",
        epilog="Run with --list-printers to find your exact printer name."
    )
    parser.add_argument("directory", help="Directory containing PDF files")
    parser.add_argument("side", choices=["front", "back"], help="Side to print")
    parser.add_argument("--printer", default=PRINTER_NAME,
                        help=f"Windows printer name (default: '{PRINTER_NAME}')")
    parser.add_argument("--list-printers", action="store_true",
                        help="List all available Windows printers and exit")
    args = parser.parse_args()

    directory = os.path.expanduser(args.directory)
    if not os.path.isdir(directory):
        print(f"Error: '{directory}' is not a valid directory.")
        sys.exit(1)

    pdf_files = sorted(glob.glob(os.path.join(directory, "*.pdf")))
    if not pdf_files:
        print(f"No PDF files found in '{directory}'.")
        sys.exit(0)

    sumatra_exe = find_sumatra()

    print(f"Found {len(pdf_files)} PDF file(s) in '{directory}'")
    print(f"Printer  : {args.printer}")
    print(f"Side     : {args.side}")
    print(f"Mode     : Black & White, 2 PDF pages per landscape sheet")
    print(f"Backend  : {'SumatraPDF @ ' + sumatra_exe if sumatra_exe else 'win32api fallback'}")

    for pdf_path in pdf_files:
        print_pdf(pdf_path, args.side, args.printer, sumatra_exe)

    print(f"\n{'='*60}")
    print("All done!")


if __name__ == "__main__":
    main()
