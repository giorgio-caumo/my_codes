from PyPDF2 import PdfMerger, PdfReader
import os
from pathlib import Path

class PDFMerger:
    def __init__(self, main_pdf_path, attachment_path, output_path):
        self.main_pdf_path = main_pdf_path
        self.attachment_path = attachment_path
        self.output_path = output_path
        self.pdf_merger = PdfMerger()
        self.added_pages = 0

    def add_main_pdf(self):
        with open(self.main_pdf_path, 'rb') as main_pdf_file:
            self.pdf_merger.append(main_pdf_file)

    def collect_attachments(self):
        attachment_files = {}
        attachment_folders = os.listdir(self.attachment_path)

        for folder in attachment_folders:
            folder_path = os.path.join(self.attachment_path, folder)
            if os.path.isdir(folder_path):
                pdf_files = [os.path.join(folder_path, file) for file in os.listdir(folder_path) if
                             file.lower().endswith('.pdf')]
                attachment_files[folder] = pdf_files

        return attachment_files

    def add_insertion_points(self, insertion_points):
        for insertion_page in sorted(insertion_points.keys()):
            for pdf_to_insert in insertion_points[insertion_page]:
                with open(pdf_to_insert, 'rb') as pdf_to_append_file:
                    pdf_reader = PdfReader(pdf_to_append_file)
                    self.pdf_merger.merge(insertion_page + self.added_pages, pdf_reader)
                    self.added_pages += len(pdf_reader.pages)

    def write_output(self):
        with open(self.output_path, 'wb') as output_file:
            self.pdf_merger.write(output_file)


if __name__ == "__main__":

    THISDIR = Path(__file__).parent

    # Specify the directory where your PDF files are located
    attachment_path = THISDIR / "Attachments"

    # Specify the main PDF and the output path
    main_pdf = THISDIR / "print.pdf"
    output_pdf = THISDIR / "Merged.pdf"

    # Create an instance of PDFMerger
    pdf_merger = PDFMerger(main_pdf, attachment_path, output_pdf)

    # Collect attachment files
    attachment_files = pdf_merger.collect_attachments()

    # Specify insertion points based on your folders
    insertion_points = {
        11: attachment_files.get('Calculations', []),
        13: attachment_files.get('Attachment 1', []),
        15: attachment_files.get('Attachment 2', [])
    }

    # Merge the PDFs
    pdf_merger.add_main_pdf()
    pdf_merger.add_insertion_points(insertion_points)
    pdf_merger.write_output()
    print("PDF merged!")
