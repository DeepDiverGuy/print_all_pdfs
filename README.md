# print_all_pdfs

If you ever need to physically print all the PDFs in a directory, this Python script is exactly what you need!

It handles everything automatically: printing 2 PDF pages per landscape sheet, splitting front and back sides for manual duplex printing, and keeping the paper stack perfectly aligned across multiple PDFs — even when a PDF has an odd number of pages.

---

## Prerequisites

- **Windows** operating system
- **Python 3.10+** — [download here](https://www.python.org/downloads/)
- **SumatraPDF** (strongly recommended for best compatibility) — [download here](https://www.sumatrapdfreader.org/download-free-pdf-viewer)
- Required Python packages:
  ```
  pip install pypdf pywin32
  ```

> **Note:** If SumatraPDF is not found, the script falls back to `win32api`. In that case, you must manually configure your printer's default settings to use B&W and 1 page per sheet (since the 2-up layout is already baked into the PDF by the script).

---

## Setup

**1. Clone or download** `print_all_pdfs.py` into any folder.

**2. Install dependencies:**
```bash
pip install pypdf pywin32
```

**3. Set your printer name.** Open `print_all_pdfs.py` and edit line 41:
```python
PRINTER_NAME = "EPSON L130 Series"   # ← replace with your printer's exact name
```
Not sure of the name? Run:
```bash
python print_all_pdfs.py --list-printers
```

---

## Usage

```bash
python print_all_pdfs.py <directory> <front|back> [--printer "Printer Name"]
```

### Arguments

| Argument | Description |
|---|---|
| `directory` | Path to the folder containing your PDF files |
| `front` / `back` | Which side of the paper to print |
| `--printer` | (Optional) Override the printer name set in the script |
| `--list-printers` | List all available Windows printers and exit |

---

## How to print duplex (double-sided) manually

Since the Epson L130 is a single-sided printer, duplex printing is done in two passes:

**Step 1 — Print the front sides:**
```bash
python print_pdfs.py "C:\path\to\your\pdfs" front
```
Wait for all pages to finish printing.

**Step 2 — Flip the stack and feed it back into the printer.**

**Step 3 — Print the back sides:**
```bash
python print_all_pdfs.py "C:\path\to\your\pdfs" back
```

> The script processes PDFs in alphabetical order on both passes, so the pages will always line up correctly.

---

## How the page layout works

Each physical sheet holds **2 PDF pages side-by-side in landscape**:

```
┌─────────────┬─────────────┐
│   Page 1    │   Page 2    │  ← Sheet 1 (front)
└─────────────┴─────────────┘
┌─────────────┬─────────────┐
│   Page 3    │   Page 4    │  ← Sheet 1 (back)
└─────────────┴─────────────┘
┌─────────────┬─────────────┐
│   Page 5    │   Page 6    │  ← Sheet 2 (front)
└─────────────┴─────────────┘
```

| Side | Physical sheets printed | PDF pages |
|---|---|---|
| Front | Odd sheets (1, 3, 5, …) | 1,2 — 5,6 — 9,10 — … |
| Back  | Even sheets (2, 4, 6, …) | 3,4 — 7,8 — 11,12 — … |

---

## Blank page alignment (odd-page PDFs)

When a PDF has an **odd number of pages**, the front and back sheet counts don't match. For example, a 5-page PDF produces 3 physical sheets:

```
Front prints: Sheet 1 (pages 1,2)  +  Sheet 3 (page 5)   →  2 sheets
Back  prints: Sheet 2 (pages 3,4)                          →  1 sheet + 1 BLANK
```

The script automatically appends a **blank sheet** at the end of the back-side job for that PDF. This keeps the paper stack perfectly aligned so the next PDF's pages land on the correct sheets.

---

## Example

```
$ python print_all_pdfs.py "F:/Documents/Assignments" front

Found 3 PDF file(s) in 'F:/Documents/Assignments'
Printer  : EPSON L130 Series
Side     : front
Mode     : Black & White, 2 PDF pages per landscape sheet
Backend  : SumatraPDF @ C:\Program Files\SumatraPDF\SumatraPDF.exe

============================================================
File  : 01_lecture_notes.pdf
Side  : front
Pages : 5  |  Physical sheets: 3
PDF pages selected : [1, 2, 5]
Output sheets      : 2
    [SumatraPDF] ... monochrome,fit ...
    OK Job sent.

============================================================
File  : 02_assignment.pdf
...

============================================================
All done!
```
