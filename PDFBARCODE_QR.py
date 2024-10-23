# D. Lewis PDF Barcode (QR Code) insertion utility
# Developed October 10, 2024 Revision 1.4
# Windows OS

import os
import pandas as pd
import qrcode  # For generating QR codes
from io import BytesIO
from PIL import Image
import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox

# Function to generate QR code image in memory from a string value
def generate_qrcode_in_memory(value):
    try:
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10,
            border=4,
        )
        qr.add_data(value)
        qr.make(fit=True)
        qr_image = qr.make_image(fill='black', back_color='white')

        # Convert QR code image to bytes for processing
        buffer = BytesIO()
        qr_image.save(buffer, format='PNG')
        buffer.seek(0)  # Reset buffer position
        return buffer
    except Exception as e:
        print(f"Error generating QR code for {value}: {e}")
        return None

# Function to process Excel, generate QR codes in-memory, and insert them into PDFs
def convert_excel_to_qrcodes_and_insert(excel_file, column_name, pdf_dir, x_coord, y_coord, qr_size):
    # Load the Excel file
    df = pd.read_excel(excel_file)

    # Check if the 'Barcode Info' (now QR Info) column exists
    if column_name not in df.columns:
        raise KeyError(f"Column '{column_name}' not found in Excel file. Available columns: {df.columns.tolist()}")

    # Iterate over the column values and generate QR codes
    for index, row in df.iterrows():
        value = row[column_name]
        pdf_file_name = row['File Name']

        # Ensure that the file name is a valid string
        if isinstance(pdf_file_name, str) and pd.notna(pdf_file_name):
            pdf_file_name = os.path.basename(pdf_file_name)  # Ensure only the file name is used

            # Make sure the value is a valid string
            if pd.notna(value):
                qrcode_buffer = generate_qrcode_in_memory(str(value))  # Generate QR code in memory

                if qrcode_buffer:
                    # Correctly join the directory and filename
                    pdf_path = os.path.join(pdf_dir, pdf_file_name)

                    # Check if the file exists before proceeding
                    if not os.path.exists(pdf_path):
                        print(f"PDF file not found: {pdf_path}")
                        continue

                    # Open the PDF
                    pdf_document = fitz.open(pdf_path)

                    # Insert QR code image on the first page
                    page = pdf_document[0]  # Modify this to select a specific page if needed

                    # Load the QR code image from buffer and get dimensions
                    qr_image = Image.open(qrcode_buffer)

                    # Resize the QR code image to the specified size
                    resized_qr_image = qr_image.resize((qr_size, qr_size))

                    # Save the in-memory image buffer to a temporary PNG file
                    temp_image_path = f"temp_qrcode_{index}.png"
                    resized_qr_image.save(temp_image_path)

                    # Define the image insertion rectangle on the PDF page
                    image_rect = fitz.Rect(x_coord, y_coord, x_coord + qr_size, y_coord + qr_size)

                    # Insert the resized QR code image into the PDF
                    page.insert_image(image_rect, filename=temp_image_path)

                    # Save the modified PDF to a new file using "QR_Code" in the filename
                    new_pdf_path = os.path.join(pdf_dir, f"QR_Code_{pdf_file_name}")
                    pdf_document.save(new_pdf_path)
                    pdf_document.close()

                    # Remove the temporary image file after use
                    os.remove(temp_image_path)

                    print(f"Saved modified file as {new_pdf_path}")
                else:
                    print(f"Failed to generate QR code for {value}")
            else:
                print(f"No valid value found for index {index}. Skipping.")
        else:
            print(f"Invalid or missing file name for index {index}. Skipping.")

# GUI Application
class PDFTextInserter:
    def __init__(self, master):
        self.master = master
        master.title("PDF QR Code Inserter")

        # X and Y coordinate labels and entries (shorter width)
        tk.Label(master, text="X Coordinate:").grid(row=0, column=0)
        self.x_entry = tk.Entry(master, width=10)  # Shortened width
        self.x_entry.grid(row=0, column=1)

        tk.Label(master, text="Y Coordinate:").grid(row=1, column=0)
        self.y_entry = tk.Entry(master, width=10)  # Shortened width
        self.y_entry.grid(row=1, column=1)

        # QR code size entry (shorter width)
        tk.Label(master, text="QR Code Size (in pixels):").grid(row=2, column=0)
        self.size_entry = tk.Entry(master, width=10)  # Shortened width
        self.size_entry.grid(row=2, column=1)

        # Directory selection for PDFs
        tk.Label(master, text="PDF Directory:").grid(row=3, column=0)
        self.pdf_dir_entry = tk.Entry(master, width=40)
        self.pdf_dir_entry.grid(row=3, column=1)
        self.browse_pdf_button = tk.Button(master, text="Browse", width=10, command=self.browse_pdf_directory)
        self.browse_pdf_button.grid(row=3, column=2)

        # Excel file selection with tooltip
        tk.Label(master, text="Excel File:").grid(row=4, column=0)
        self.excel_file_entry = tk.Entry(master, width=40)
        self.excel_file_entry.grid(row=4, column=1)
        self.excel_file_entry.bind("<Enter>", lambda e: self.show_tooltip(e, "Excel column names: File Name and Barcode Info"))
        self.browse_excel_button = tk.Button(master, text="Browse", width=10, command=self.browse_excel_file)
        self.browse_excel_button.grid(row=4, column=2)

        # About and Quit buttons placed directly below the Browse buttons
        self.about_button = tk.Button(master, text="About", width=10, command=self.show_about)
        self.about_button.grid(row=6, column=0)

        self.quit_button = tk.Button(master, text="Quit", width=10, command=master.quit)
        self.quit_button.grid(row=6, column=2)

        # Process PDFs button
        self.process_button = tk.Button(master, text="Process PDFs", width=10, command=self.process_pdfs)
        self.process_button.grid(row=6, column=1)

    def show_tooltip(self, event, text):
        # Create a tooltip window
        tooltip = tk.Toplevel(self.master)
        tooltip.wm_overrideredirect(True)
        tooltip.wm_geometry(f"+{event.x_root + 10}+{event.y_root + 10}")
        label = tk.Label(tooltip, text=text, background="yellow", relief="solid", borderwidth=1)
        label.pack()
        # Automatically close the tooltip after 2 seconds
        self.master.after(2000, tooltip.destroy)

    def browse_pdf_directory(self):
        directory = filedialog.askdirectory()
        self.pdf_dir_entry.delete(0, tk.END)
        self.pdf_dir_entry.insert(0, directory)

    def browse_excel_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        self.excel_file_entry.delete(0, tk.END)
        self.excel_file_entry.insert(0, file_path)

    def process_pdfs(self):
        pdf_directory = self.pdf_dir_entry.get()
        excel_file = self.excel_file_entry.get()

        # Validate inputs
        if not os.path.isdir(pdf_directory):
            messagebox.showerror("Error", "Invalid PDF directory.")
            return
        if not os.path.isfile(excel_file):
            messagebox.showerror("Error", "Invalid Excel file.")
            return

        try:
            x_coord = float(self.x_entry.get())
            y_coord = float(self.y_entry.get())
            qr_size = int(self.size_entry.get())  # QR code size in pixels
        except ValueError:
            messagebox.showerror("Error", "Please enter valid coordinates and size.")
            return

        # Call the function to generate QR codes and insert into PDFs
        try:
            convert_excel_to_qrcodes_and_insert(excel_file, 'Barcode Info', pdf_directory, x_coord, y_coord, qr_size)
            messagebox.showinfo("Success", "All PDFs processed.")
        except KeyError as e:
            messagebox.showerror("Error", str(e))

    # About button callback
    def show_about(self):
        messagebox.showinfo("About", "Developed by D. Lewis 2024 QR-Version 1.2")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFTextInserter(root)
    root.mainloop()

