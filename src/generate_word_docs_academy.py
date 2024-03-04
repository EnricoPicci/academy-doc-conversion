import os
from extract_pictures import extract_images_from_docx
from convert_bin_images import convert_all_bin_to_image
from convert_word_docs import convert_word_doc


def generate_word_docs_academy(
    input_folder=None, output_folder=None, images_folder=None
):
    # if the input_folder does not exist print an error message and exit
    if not os.path.exists(input_folder):
        print(f"Error: {input_folder} does not exist")
        return

    if input_folder is None:
        input_folder = "."  # current directory
    if output_folder is None:
        output_folder = input_folder
    if images_folder is None:
        images_folder = os.path.join(output_folder, "images")

    # read all the .docx files in the input folder and in all its subfolders recursively
    # and, for each file, create a dictionary with the file name, the subdirectory of input_folder where the file resides,
    # the file path of the file and the folder where the images of the file will be stored
    files_info = []
    for root, dirs, files in os.walk(input_folder):
        for file in files:
            if file.endswith(".docx"):
                file_path = os.path.join(root, file)
                file_name = os.path.splitext(file)[0]
                # get the name of the subdirectory of input_folder where the file resides
                subdir = os.path.relpath(root, input_folder)
                # calculate the name of a folder for the output of the images of each document
                doc_images_folder = os.path.join(images_folder, subdir, file_name)
                # calculate the name of the folder where the new transformed document will be stored
                doc_output_folder = os.path.join(output_folder, subdir)
                files_info.append(
                    {
                        "file_name": file_name,
                        "subdir": subdir,
                        "file_path": file_path,
                        "doc_images_folder": doc_images_folder,
                        "doc_output_folder": doc_output_folder,
                    }
                )

    # for each file, extract the images from the .docx file, convert the .bin files to .png and convert the document to a new document
    for file_info in files_info:
        # extract the images from the .docx file
        extract_images_from_docx(file_info["file_path"], file_info["doc_images_folder"])

        # convert the .bin files to .png
        doc_media_folder = os.path.join(file_info["doc_images_folder"], "word/media")
        convert_all_bin_to_image(doc_media_folder)

        # convert the document to a new document
        converted_file_path = convert_word_doc(
            file_info["file_path"],
            file_info["doc_images_folder"],
            file_info["doc_output_folder"],
        )
        print(f"Converted file: {converted_file_path}")


# # Example usage:
# folder_path = "academy/input_2024_02_29"
# output_folder = "academy/output_2024_02_29"
# generate_word_docs_academy(folder_path, output_folder)
