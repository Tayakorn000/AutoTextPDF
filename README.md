# AutoTextPDF

A desktop application for automatically extracting order data from shipping-label PDFs and stamping product codes onto them — eliminating manual lookup and copy-paste from order management systems.

## Overview

AutoTextPDF parses Thai e-commerce shipping label PDFs (Shopee, Lazada, etc.), matches each line item against a product database, and writes the corresponding internal product codes directly onto the output PDF. A PySide6 GUI lets operators review, correct, and manage the product database in-session.

## Features

- **PDF extraction** — Pulls order lines (item name, variant, quantity) from multi-page shipping label PDFs via PyMuPDF
- **Fuzzy product matching** — Normalises Thai text and matches items against an Excel product database
- **PDF stamping** — Overlays product codes onto the original PDF pages using ReportLab (Thai font support on macOS and Windows)
- **Drag-and-drop interface** — Drop a PDF onto the window or use the file picker
- **In-app DB editor** — Double-click any row to add or edit product codes; multi-line codes supported
- **Live reload** — Reload the product database without restarting the app
- **Cross-platform** — Font paths auto-selected for macOS (`Thonburi`) and Windows (`Tahoma / Arial`)

## Requirements

- Python 3.10+
- PySide6
- PyMuPDF (`fitz`)
- pypdf
- reportlab
- pandas
- openpyxl

```bash
pip install PySide6 pymupdf pypdf reportlab pandas openpyxl
```

## Project Structure

```
AutoTextPDF/
├── main_gui.py       # PySide6 application — UI layout and event handling
├── processor.py      # PDF parsing, product matching, and PDF stamping logic
├── products_db.xlsx  # Product database (item name → product code mapping)
└── test_qt.py        # Quick Qt environment sanity check
```

## Usage

```bash
python main_gui.py
```

1. Drop a shipping label PDF onto the drop zone (or click to browse)
2. Review the extracted order table — unmatched items are highlighted
3. Double-click any row to assign or correct a product code and save it to the database
4. Click **Label PDFs** to generate the stamped output PDF alongside the original

## Database Format

`products_db.xlsx` contains two columns:

| item | variant_name | code |
|------|--------------|------|
| ชื่อสินค้า | ชื่อตัวเลือก | PROD-001 |

Import an existing Excel file via **File → Import DB** or edit rows directly in the app.
