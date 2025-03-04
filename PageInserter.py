import tkinter as tk
from tkinter import filedialog, messagebox
import PyPDF2
import os
import sys


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


class PDFProcessor:
    @staticmethod
    def add_custom_pages(sheet_pack_path, written_answer_sheet_path, output_pdf_path):
        try:
            with open(sheet_pack_path, 'rb') as sheet_pack_file, \
                    open(written_answer_sheet_path, 'rb') as written_answer_file:

                sheet_pack = PyPDF2.PdfReader(sheet_pack_file)
                written_answer = PyPDF2.PdfReader(written_answer_file)
                writer = PyPDF2.PdfWriter()

                written_pages = list(written_answer.pages)
                written_count = len(written_pages)
                add_blank = written_count % 2 == 0

                if add_blank and written_count > 0:
                    first_page = written_pages[0]
                    blank = PyPDF2.PageObject.create_blank_page(
                        width=first_page.mediabox.width,
                        height=first_page.mediabox.height
                    )

                for main_page in sheet_pack.pages:
                    writer.add_page(main_page)

                    for wp in written_pages:
                        writer.add_page(wp)

                    if add_blank and written_count > 0:
                        writer.add_page(blank)

                with open(output_pdf_path, 'wb') as output_file:
                    writer.write(output_file)

            messagebox.showinfo("Success", "PDF processed successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process PDF: {str(e)}")


class PDFApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Page Inserter")
        icon_path = resource_path('C:\\AppendAnswerSheet\\assets\\Icon.ico')
        self.root.iconbitmap(icon_path)
        self.setup_ui()

    def setup_ui(self):
        tk.Label(self.root, text="Sheet Pack:").grid(row=0, column=0, padx=10, pady=10)
        self.entry_sheet_pack = tk.Entry(self.root, width=50)
        self.entry_sheet_pack.grid(row=0, column=1, padx=10, pady=10)
        tk.Button(self.root, text="Browse", command=self.select_sheet_pack).grid(row=0, column=2, padx=10, pady=10)

        tk.Label(self.root, text="Written Answer Sheet:").grid(row=1, column=0, padx=10, pady=10)
        self.entry_written_answer = tk.Entry(self.root, width=50)
        self.entry_written_answer.grid(row=1, column=1, padx=10, pady=10)
        tk.Button(self.root, text="Browse", command=self.select_written_answer).grid(row=1, column=2, padx=10, pady=10)

        tk.Button(self.root, text="Process PDF", command=self.process_pdf).grid(row=3, column=0, columnspan=3, padx=10,
                                                                                pady=10)

    def select_sheet_pack(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        self.entry_sheet_pack.delete(0, tk.END)
        self.entry_sheet_pack.insert(0, file_path)

    def select_written_answer(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        self.entry_written_answer.delete(0, tk.END)
        self.entry_written_answer.insert(0, file_path)

    def process_pdf(self):
        sheet_pack = self.entry_sheet_pack.get()
        written_answer = self.entry_written_answer.get()

        if not sheet_pack or not written_answer:
            messagebox.showwarning("Warning", "Please select all files!")
            return

        output_path = os.path.splitext(sheet_pack)[0] + "_processed.pdf"
        PDFProcessor.add_custom_pages(sheet_pack, written_answer, output_path)


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFApp(root)
    root.mainloop()