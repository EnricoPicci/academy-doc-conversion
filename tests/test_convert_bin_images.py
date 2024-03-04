import os

from src.convert_bin_images import convert_all_bin_to_image


def test_extract_images_from_docx():
    # Path to the images folder and output directory
    images_folder = "test_data/bin_images"
    output_dir = "tests/output/bin_images"

    # count the number of files in the folder
    num_files = len(
        [
            f
            for f in os.listdir(images_folder)
            if os.path.isfile(os.path.join(images_folder, f))
        ]
    )
    # Call the function
    convert_all_bin_to_image(images_folder, output_dir)

    # Check if the image file exists in the temporary directory
    assert os.path.isfile(os.path.join(output_dir, "image1.png"))

    # check that the files extracted go from image1.bin to image33.bin
    for i in range(1, num_files + 1):
        assert os.path.isfile(os.path.join(images_folder, f"image{i}.bin"))
