# D. Lewis PDF Barcode insertion utility
# Developed October 10, 2024 Revision 1.4
# Windows OS

import os
import pandas as pd
import barcode
from barcode.writer import ImageWriter
import fitz  # PyMuPDF
from io import BytesIO
from PIL import Image
import tkinter as tk
from tkinter import filedialog, messagebox

# Function to generate barcode image in memory from a string value
def generate_barcode_in_memory(value):
    try:
        code128 = barcode.get('code128', value, writer=ImageWriter())
        buffer = BytesIO()  # Create an in-memory buffer
        code128.write(buffer)  # Write the barcode to the buffer
        buffer.seek(0)  # Reset buffer position
        return buffer
    except Exception as e:
        print(f"Error generating barcode for {value}: {e}")
        return None

# Function to process Excel, generate barcodes in-memory, and insert them into PDFs
def convert_excel_to_barcodes_and_insert(excel_file, column_name, pdf_dir, x_coord, y_coord, barcode_height):
    # Load the Excel file
    df = pd.read_excel(excel_file)

    # Check if the 'Barcode Info' column exists
    if column_name not in df.columns:
        raise KeyError(f"Column '{column_name}' not found in Excel file. Available columns: {df.columns.tolist()}")

    # Iterate over the column values and generate barcodes
    for index, row in df.iterrows():
        value = row[column_name]
        pdf_file_name = row['File Name']

        # Ensure that the file name is a valid string
        if isinstance(pdf_file_name, str) and pd.notna(pdf_file_name):
            pdf_file_name = os.path.basename(pdf_file_name)  # Ensure only the file name is used

            # Make sure the value is a valid string
            if pd.notna(value):
                barcode_buffer = generate_barcode_in_memory(str(value))  # Generate barcode in memory

                if barcode_buffer:
                    # Correctly join the directory and filename
                    pdf_path = os.path.join(pdf_dir, pdf_file_name)

                    # Check if the file exists before proceeding
                    if not os.path.exists(pdf_path):
                        print(f"PDF file not found: {pdf_path}")
                        continue

                    # Open the PDF
                    pdf_document = fitz.open(pdf_path)

                    # Insert barcode image on the first page
                    page = pdf_document[0]  # Modify this to select a specific page if needed

                    # Load the barcode image from buffer and get dimensions
                    barcode_image = Image.open(barcode_buffer)

                    # Rotate the image 90 degrees counterclockwise
                    rotated_barcode_image = barcode_image.rotate(90, expand=True)

                    # Resize the barcode image to match the specified height
                    img_width, img_height = rotated_barcode_image.size
                    aspect_ratio = img_width / img_height
                    new_width = int(aspect_ratio * barcode_height)
                    resized_barcode_image = rotated_barcode_image.resize((new_width, barcode_height))

                    # Save the in-memory image buffer to a temporary PNG file
                    temp_image_path = f"temp_barcode_{index}.png"
                    resized_barcode_image.save(temp_image_path)

                    # Define the image insertion rectangle on the PDF page
                    image_rect = fitz.Rect(x_coord, y_coord, x_coord + new_width, y_coord + barcode_height)

                    # Insert the resized and rotated barcode image into the PDF
                    page.insert_image(image_rect, filename=temp_image_path)

                    # Save the modified PDF to a new file using "Barcode" in the filename
                    new_pdf_path = os.path.join(pdf_dir, f"Barcode_{pdf_file_name}")
                    pdf_document.save(new_pdf_path)
                    pdf_document.close()

                    # Remove the temporary image file after use
                    os.remove(temp_image_path)

                    print(f"Saved modified file as {new_pdf_path}")
                else:
                    print(f"Failed to generate barcode for {value}")
            else:
                print(f"No valid value found for index {index}. Skipping.")
        else:
            print(f"Invalid or missing file name for index {index}. Skipping.")

# GUI Application
class PDFTextInserter:
    def __init__(self, master):
        self.master = master
        master.title("PDF Barcode Inserter")

        # X and Y coordinate labels and entries (shorter width)
        tk.Label(master, text="X Coordinate:").grid(row=0, column=0)
        self.x_entry = tk.Entry(master, width=10)  # Shortened width
        self.x_entry.grid(row=0, column=1)

        tk.Label(master, text="Y Coordinate:").grid(row=1, column=0)
        self.y_entry = tk.Entry(master, width=10)  # Shortened width
        self.y_entry.grid(row=1, column=1)

        # Barcode height entry (shorter width)
        tk.Label(master, text="Barcode Height (in pixels):").grid(row=2, column=0)
        self.height_entry = tk.Entry(master, width=10)  # Shortened width
        self.height_entry.grid(row=2, column=1)

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
            barcode_height = int(self.height_entry.get())  # Barcode height in pixels
        except ValueError:
            messagebox.showerror("Error", "Please enter valid coordinates and height.")
            return

        # Call the function to generate barcodes and insert into PDFs
        try:
            convert_excel_to_barcodes_and_insert(excel_file, 'Barcode Info', pdf_directory, x_coord, y_coord, barcode_height)
            messagebox.showinfo("Success", "All PDFs processed.")
        except KeyError as e:
            messagebox.showerror("Error", str(e))

    # About button callback
    def show_about(self):
        messagebox.showinfo("About", "Developed by D. Lewis 2024 Version 1.4")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFTextInserter(root)
    root.mainloop()
