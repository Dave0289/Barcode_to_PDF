# Barcode_to_PDF
A Python application that reads an Excel file to generate barcodes and inserts them into specified locations in PDF files.
This tool utilizes pandas for Excel handling, barcode for barcode generation, and PyMuPDF for PDF manipulation.

Table of Contents

- Features
- Requirements
- Installation
- Usage
- GUI Overview
- License

Features

- Generate Code 128 barcodes from values in an Excel file.
- Insert generated barcodes into specified coordinates of PDF files.
- User-friendly GUI for selecting files and setting parameters.

Requirements

Before running the application, ensure you have the following installed:

- Python 3.9 or 3.10
- Required Python packages:
  - pandas
  - barcode
  - Pillow
  - PyMuPDF
  - Openpyxl
  - tkinter (included with Python)

You can install the required packages using pip:

```bash
pip install pandas python-barcode Pillow PyMuPDF Openpyxl
```

Installation

1. Clone this repository or download the script file PDFBARCODE.py.
2. Place the script in a suitable directory.
3. Ensure you have the required Python packages installed (see above).

Usage

1. Prepare your Excel file with the following columns:
   - File Name: The full paths of the PDF files where the barcodes will be inserted.
   - Barcode Info: The values for which barcodes will be generated.

   Example:

   | File Name                   | Barcode Info |
   |-----------------------------|--------------|
   | C:\path\to\document1.pdf    | 123456789012 |
   | C:\path\to\document2.pdf    | 987654321098 |

2. Place your PDF files in a directory that is accessible to the script.
3. Run the script:

```bash
python PDFBARCODE.py
```

4. Use the GUI to:
   - Enter the X and Y coordinates for barcode placement.
   - Specify the desired barcode height (in pixels).
   - Browse and select the directory containing the PDF files.
   - Browse and select the Excel file you prepared.

5. Click Process PDFs to generate and insert the barcodes into the PDFs. A success message will appear when all PDFs have been processed.

GUI Overview

- X Coordinate: Enter the horizontal position for barcode insertion (in pixels).
- Y Coordinate: Enter the vertical position for barcode insertion (in pixels).
- Barcode Height: Specify the height of the barcode (in pixels).
- PDF Directory: Choose the directory where your PDF files are located.
- Excel File: Choose the Excel file containing the barcode data.
- About: Click to view information about the application.
- Quit: Click to close the application.
- Process PDFs: Click to start processing the PDFs.

License

This project is licensed free to use. The developer [D.Lewis] shall be held harmless for any errors and/or omissions. Use at own risk.
