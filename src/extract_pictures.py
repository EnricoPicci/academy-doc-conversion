import zipfile
import os


def extract_images_from_docx(docx_path, output_folder):
    # Ensure the output folder exists
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Open the .docx file as a ZIP file
    with zipfile.ZipFile(docx_path, "r") as docx_zip:
        # List all contents
        for file_info in docx_zip.infolist():
            # Check if this file is in the 'word/media/' folder
            if file_info.filename.startswith("word/media/"):
                # Extract the file
                docx_zip.extract(file_info, output_folder)
