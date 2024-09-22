import os
from docx import Document
import openpyxl
from openpyxl import Workbook

# Function to extract hyperlinks from a Word file (.docx)
def extract_hyperlinks_from_docx(docx_file):
    document = Document(docx_file)
    hyperlinks = []

    # Iterate through all paragraphs and runs to find hyperlinks
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            if run.text and run.hyperlink:  # Check if the run has a hyperlink
                title = run.text
                hyperlink = run.hyperlink.target
                hyperlinks.append((title.strip(), hyperlink))

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
    output_file_name = os.path.basename(file_name).replace(".docx", ".xlsx")
    output_file_path = os.path.join(output_folder, output_file_name)
    workbook.save(output_file_path)

    print(f"Data written to {output_file_path}")

# Main function
def main():
    # Folder where the Word file is located
    input_folder = "input"
    
    # Folder where the Excel file will be saved
    output_folder = "output"

    # List all files in the folder and select the first Word file
    docx_files = [f for f in os.listdir(input_folder) if f.endswith(".docx")]

    if docx_files:
        docx_file_path = os.path.join(input_folder, docx_files[0])  # Get the first Word file in the folder
        
        # Extract hyperlinks
        hyperlinks = extract_hyperlinks_from_docx(docx_file_path)

        # Write data to Excel
        if hyperlinks:
            write_to_excel(docx_file_path, hyperlinks, output_folder)
        else:
            print("No hyperlinks found in the Word document.")
    else:
        print("No Word (.docx) files found in the input folder.")

# Run the script
if __name__ == "__main__":
    main()
