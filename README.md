# 📄 Excel Sheet to Paginated PDF Converter

This Python project converts each sheet of an Excel file into a **multi-page PDF**, with the **PDF named after the sheet name**. It automatically handles pagination based on a user-defined number of rows per page and removes all gridlines for a clean, print-ready layout.

---

## ✨ Features

- 🔄 Converts every sheet in an Excel workbook to a separate PDF
- 📄 Automatically paginates large sheets (e.g., 30 rows per page)
- 📁 PDF file is named after the Excel sheet name
- 🧼 Grid lines/borders removed for minimalistic design
- 💻 Works on both local machines and Google Colab

---

#note
This script converts only tabular data — charts, images, or Excel formatting are not included.

PDF layout is simple and clean with no internal cell borders.if you want internal cell borders rem ove that particular code

## 📦 Requirements

You can install all required libraries using:

```bash
pip install pandas matplotlib openpyxl
