import fitz  # PyMuPDF
import openpyxl
from openpyxl import Workbook
import os

# Function to extract hyperlinks from a PDF file
def extract_hyperlinks_from_pdf(pdf_file):
    doc = fitz.open(pdf_file)
    hyperlinks = []

    for page_num in range(len(doc)):
        page = doc[page_num]
        for link in page.get_links():
            uri = link.get('uri', None)
            if uri:
                # Find the clickable text (title)
                title = page.get_text("text", clip=link['from'])  # Extract text from the clickable area
                hyperlinks.append((title.strip(), uri))

    doc.close()
    return hyperlinks

# Function to write hyperlinks to an Excel file in the output folder
def write_to_excel(file_name, data, output_folder):
    workbook = Workbook()
    sheet = workbook.active

    # First row: file name
    sheet["A1"] = "File Name"
    sheet["B1"] = file_name

    # Second row: headers
    sheet["A2"] = "Title"
    sheet["B2"] = "Hyperlink"

    # From the third row: data (title and hyperlink)
    for idx, (title, hyperlink) in enumerate(data, start=3):
        sheet[f"A{idx}"] = title
        sheet[f"B{idx}"] = hyperlink

    # Ensure output folder exists
    os.makedirs(output_folder, exist_ok=True)

    # Save the Excel file in the output folder
    output_file_name = os.path.basename(file_name).replace(".pdf", ".xlsx")
    output_file_path = os.path.join(output_folder, output_file_name)
    workbook.save(output_file_path)

    print(f"Data written to {output_file_path}")

# Main function
def main():
    # Folder where the PDF file is located
    input_folder = "input"
    
    # Folder where the Excel file will be saved
    output_folder = "output"

    # List all files in the folder and select the first PDF file
    pdf_files = [f for f in os.listdir(input_folder) if f.endswith(".pdf")]

    if pdf_files:
        pdf_file_path = os.path.join(input_folder, pdf_files[0])  # Get the first PDF file in the folder
        
        # Extract hyperlinks
        hyperlinks = extract_hyperlinks_from_pdf(pdf_file_path)

        # Write data to Excel
        if hyperlinks:
            write_to_excel(pdf_file_path, hyperlinks, output_folder)
        else:
            print("No hyperlinks found in the PDF.")
    else:
        print("No PDF files found in the input folder.")

# Run the script
if __name__ == "__main__":
    main()
