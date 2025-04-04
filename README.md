# ScoreMe-Assignment
# PDF Table Extractor

This tool extracts tables from **system-generated PDFs** (with or without borders, including irregular-shaped tables) **without using Tabula, Camelot, or converting PDFs to images**. The extracted tables are stored in Excel format — one Excel file per PDF.

---

## Features

-  Detects **tables with borders**, **without borders**, and **irregular-shaped** tables.
-  No image conversion or third-party wrappers like Tabula or Camelot.
-  Uses **text-based PDF parsing** with `pdfplumber`.
-  Handles **multi-line cells**, **irregular rows**, and **merged cell patterns**.
-  Exports each detected table to a separate sheet in an Excel file.
-  Processes multiple PDFs in batch mode.
-  Simple CLI usage.

---

##  Input

- Folder containing one or more **system-generated PDF files** (not scanned/image-based).
- Each PDF may contain multiple tables with varying structures.

---

##  Output

- Excel file per PDF in the specified output folder.
- Each table is saved as a **separate sheet** in the Excel file.
- Sheet names indicate the page and order of tables found.

---

## Dependencies

Install required libraries using pip:

```bash
pip install pdfplumber pandas openpyxl

Tested with:
- Python  3.11.11
- pdfplumber 0.11.6
- pandas 2.2.3
- openpyxl 3.1.5

---


```
input_folder: Directory containing `.pdf` files.
output_folder: Where `.xlsx` output files will be saved.



