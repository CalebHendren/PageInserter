import tkinter as tk
from tkinter import filedialog, messagebox
import PyPDF2
import os
import sys


def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller."""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def add_custom_page_every_other(sheet_pack_path, written_answer_sheet_path, output_pdf_path):
    try:
        with open(sheet_pack_path, 'rb') as sheet_pack_file, open(written_answer_sheet_path,
                                                                  'rb') as written_answer_file:
            reader_sheet_pack = PyPDF2.PdfReader(sheet_pack_file)
            reader_written_answer = PyPDF2.PdfReader(written_answer_file)
            written_answer_page = reader_written_answer.pages[0]
            writer = PyPDF2.PdfWriter()

            for page_number in range(len(reader_sheet_pack.pages)):
                writer.add_page(reader_sheet_pack.pages[page_number])
                writer.add_page(written_answer_page)

            with open(output_pdf_path, 'wb') as output_file:
                writer.write(output_file)

        messagebox.showinfo("Success", "PDF has been processed successfully!")
    except Exception as e:
        messagebox.showerror("Error", "Failed to process PDF: " + str(e))


def select_sheet_pack():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    entry_sheet_pack.delete(0, tk.END)
    entry_sheet_pack.insert(0, file_path)


def select_written_answer_sheet():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    entry_written_answer_sheet.delete(0, tk.END)
    entry_written_answer_sheet.insert(0, file_path)


def process_pdf():
    sheet_pack_path = entry_sheet_pack.get()
    written_answer_sheet_path = entry_written_answer_sheet.get()

    if sheet_pack_path and written_answer_sheet_path:
        output_pdf_path = os.path.splitext(sheet_pack_path)[0] + "_processed.pdf"
        add_custom_page_every_other(sheet_pack_path, written_answer_sheet_path, output_pdf_path)
    else:
        messagebox.showwarning("Warning", "Please select all files!")


root = tk.Tk()
root.title("PDF Page Inserter")

# Use the resource_path function to ensure correct path resolution for the icon
icon_path = resource_path('assets/pdf17.ico')
root.iconbitmap(icon_path)

tk.Label(root, text="Sheet Pack:").grid(row=0, column=0, padx=10, pady=10)
entry_sheet_pack = tk.Entry(root, width=50)
entry_sheet_pack.grid(row=0, column=1, padx=10, pady=10)
button_sheet_pack = tk.Button(root, text="Browse", command=select_sheet_pack)
button_sheet_pack.grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="Written Answer Sheet:").grid(row=1, column=0, padx=10, pady=10)
entry_written_answer_sheet = tk.Entry(root, width=50)
entry_written_answer_sheet.grid(row=1, column=1, padx=10, pady=10)
button_written_answer_sheet = tk.Button(root, text="Browse", command=select_written_answer_sheet)
button_written_answer_sheet.grid(row=1, column=2, padx=10, pady=10)

process_button = tk.Button(root, text="Process PDF", command=process_pdf)
process_button.grid(row=3, column=0, columnspan=3, padx=10, pady=10)

root.mainloop()
