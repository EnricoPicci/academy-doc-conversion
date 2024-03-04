import os
from src.extract_pictures import extract_images_from_docx


def test_extract_images_from_docx():
    # Path to the test .docx file and output directory
    docx_file = "test_data/example_to_convert.docx"
    output_dir = "tests/output"

    # Call the function
    extract_images_from_docx(docx_file, output_dir)

    image_folder = os.path.join(output_dir, "word", "media")

    # Check if the image file exists in the temporary directory
    assert os.path.isfile(os.path.join(image_folder, "image1.bin"))

    # check that the files extracted go from image1.bin to image33.bin
    for i in range(1, 34):
        assert os.path.isfile(os.path.join(image_folder, f"image{i}.bin"))
