import os
from my_package import docx_script, pptx_script, pdf_script, xlsx_script

# Main function to scan input folder and process files
def main():
    input_folder = "input"
    output_folder = "output"
    files = os.listdir(input_folder)

    for file in files:
        file_path = os.path.join(input_folder, file)
        if file.endswith(".xlsx"):
            hyperlinks = xlsx_script.extract_hyperlinks_from_xlsx(file_path)
            xlsx_script.write_to_excel(file_path, hyperlinks, output_folder)
        elif file.endswith(".docx"):
            hyperlinks = docx_script.extract_hyperlinks_from_docx(file_path)
            docx_script.write_to_excel(file_path, hyperlinks, output_folder)
        elif file.endswith(".pptx"):
            hyperlinks = pptx_script.extract_hyperlinks_from_pptx(file_path)
            pptx_script.write_to_excel(file_path, hyperlinks, output_folder)
        elif file.endswith(".pdf"):
            hyperlinks = pdf_script.extract_hyperlinks_from_pdf(file_path)
            pdf_script.write_to_excel(file_path, hyperlinks, output_folder)
        else:
            print(f"Unsupported file type: {file}")

# Run the script
if __name__ == "__main__":
    main()
