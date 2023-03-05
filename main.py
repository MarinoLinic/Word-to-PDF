import os
from docx2pdf import convert

# specify the folder containing the docx files (current directory)
folder_path = "./Word"

# create a new directory for the PDF files
pdf_path = os.path.join(folder_path, "../PDF")
if not os.path.exists(pdf_path):
    os.makedirs(pdf_path)

# get a list of all docx files in the folder
docx_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith(".docx")]

# loop through the docx files and convert each one to PDF
for docx_file in docx_files:
    try:
        # get the file name without extension
        file_name = os.path.splitext(os.path.basename(docx_file))[0]

        # convert the docx file to a PDF with the same name
        pdf_file = os.path.join(pdf_path, f"{file_name}.pdf")
        convert(docx_file, pdf_file)
        print(f"Converted {docx_file} to {pdf_file}.")
    except Exception as e:
        print(f"Error converting {docx_file} to PDF: {e}")