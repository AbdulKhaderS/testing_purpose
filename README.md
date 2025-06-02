# Barcode Highlighter Tool 📦

A simple GUI-based tool written in Python that helps you identify and highlight duplicate barcodes from an Excel file (`NEW ITEMS SAM HUSSAIN.xlsx`) by comparing them against a master list of barcodes (`ALL BARCODES.xlsx`). Matching barcodes are highlighted in red with bold font and saved as a new Excel file.

---

## ✅ Features

- 🔍 Compares barcode values between two Excel files
- 🖌 Highlights matching barcodes in red with bold font
- 📁 Lets you select or drag-and-drop `.xlsx` files
- 💾 Saves output as a new Excel file: `NEW_ITEMS_HIGHLIGHTED.xlsx`
- 🚀 Fast lookup using Python sets for large datasets

---

## 🧰 Requirements

Before running this application, ensure you have the following installed:

```bash
pip install pandas openpyxl
