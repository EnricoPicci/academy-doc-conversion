from PIL import Image
import os


def convert_bin_to_image(bin_path, output_path, output_format="PNG", verobose=False):
    try:
        # Attempt to open the .bin file as an image
        with Image.open(bin_path) as img:
            # Save the image in the desired format
            img.save(output_path, output_format)
            if verobose:
                print(f"Successfully converted and saved: {output_path}")
    except IOError:
        print(
            f"Cannot convert {bin_path}. Unsupported image format or not an image file."
        )


# convert all images in the folder
def convert_all_bin_to_image(folder_path, output_folder=None):
    if output_folder is None:
        output_folder = folder_path
    # Ensure the output folder exists
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # read all the .bin files in the folder
    for file in os.listdir(folder_path):
        if file.endswith(".bin"):
            bin_path = os.path.join(folder_path, file)
            output_path = os.path.join(output_folder, file.replace(".bin", ".png"))
            convert_bin_to_image(bin_path, output_path)
