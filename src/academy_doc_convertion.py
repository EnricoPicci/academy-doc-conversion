#!/usr/bin/env python3

import argparse

from generate_word_docs_academy import generate_word_docs_academy


def main():
    parser = argparse.ArgumentParser(
        description="Convert academy word documents to readable format."
    )

    # Add the arguments
    parser.add_argument(
        "--input-folder",
        type=str,
        help="path to the folder where the word documents to be converted are stored",
    )
    parser.add_argument(
        "--output-folder",
        type=str,
        help="path where the converted documents will be stored",
    )
    parser.add_argument(
        "--images-folder",
        type=str,
        help="path where the images extracted from the documents will be stored",
    )

    # Parse command line arguments
    args = parser.parse_args()

    generate_word_docs_academy(
        args.input_folder, args.output_folder, args.images_folder
    )


if __name__ == "__main__":
    main()

# Example usage:
# python3 src/academy_doc_convertion.py --input-folder "path_to_input" --output-folder "path_to_output" --images-folder "path_to_images"
