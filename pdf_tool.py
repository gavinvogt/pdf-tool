'''
File: pdf_tool.py
Author: Gavin Vogt
This program provides some utility for modifying PDF files
'''

try:
    from PyPDF2 import PdfFileReader, PdfFileWriter
except ModuleNotFoundError:
    print("Please install the `PyPDF2` module:")
    print(">> python -m pip install pypdf2")
    exit(1)
import os

INFO_STR = """The following commands are available for working with your PDF:

split
    - Splits the PDF into a new PDF for each individual page.

slice <page1> <page2> ...
slice <start>-<end> ...
    - Creates a new PDF containing the pages specified in the given order.
    - A combination of page numbers and ranges may be used.

list
    - Lists all .pdf files in the current directory.

help
    - Displays this dialog.

exit
quit
    - Exits the program."""

def process_command(pdf: PdfFileReader, pdf_name: str, command: str):
    if command == "help":
        print(INFO_STR)
    elif command == "list":
        list_pdfs()
    elif command == "split":
        split_pdf(pdf, pdf_name)
    else:
        words = command.split()
        if words[0] == "slice":
            try:
                slice_pdf(pdf, words[1:])
            except Exception:
                print("Failed to slice")
        else:
            print("Invalid command")

def list_pdfs():
    cwd = os.getcwd()
    for file_name in os.listdir(cwd):
        if os.path.isfile(os.path.join(cwd, file_name)):
            # Is a file
            ext = file_name.split(".")[-1]
            if ext.lower() == "pdf":
                # Found a PDF
                print("   ", file_name)

def split_pdf(pdf: PdfFileReader, pdf_name: str):
    for i in range(pdf.getNumPages()):
        writer = PdfFileWriter()
        writer.addPage(pdf.getPage(i))
        save_path = f'{pdf_name}_Page_{i+1}.pdf'
        with open(save_path, 'wb') as f:
            writer.write(f)
            print("Created file:", save_path)
    print("Split successful")

def slice_pdf(pdf: PdfFileReader, pages):
    page_nums = []
    for num in pages:
        page_range = num.split("-")
        if len(page_range) == 1:
            # Just a single page number
            page_nums.append(int(num))
        elif len(page_range) == 2:
            # Start - end range (inclusive)
            for i in range(int(page_range[0]), int(page_range[1]) + 1):
                page_nums.append(i)
        else:
            print("Invalid range:", num)
            return

    # Verify that the page numbers are valid
    num_pages = pdf.getNumPages()
    for page_num in page_nums:
        if not (1 <= page_num <= num_pages):
            print("Out of range: page", page_num)
            return

    # Create the slice
    file_name = input("Save as: ")
    file_name = pdf_without_extension(file_name) + ".pdf"
    writer = PdfFileWriter()
    for page_num in page_nums:
        writer.addPage(pdf.getPage(page_num - 1))
    with open(file_name, 'wb') as f:
        writer.write(f)
    print("Slice successful")

def pdf_without_extension(pdf_path):
    return os.path.splitext(pdf_path)[0]

def main():
    # Display PDF files in current directory
    print("PDF files in current directory:")
    list_pdfs()
    print()

    # Load the PDF
    pdf = None
    while pdf is None:
        pdf_path = input("PDF path: ")
        try:
            pdf = PdfFileReader(pdf_path)
        except Exception:
            pass
    num_pages = pdf.getNumPages()
    print("Loaded pdf:", pdf_path, f"({num_pages} pages)")

    # Take commands
    print("Enter 'help' to see available commands")
    pdf_name = pdf_without_extension(pdf_path)
    while True:
        command = input("\n>> ").strip().lower()
        if command in ("exit", "quit"):
            return
        process_command(pdf, pdf_name, command)

if __name__ == "__main__":
    main()
