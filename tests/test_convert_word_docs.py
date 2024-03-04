import os
from docx import Document
from src.convert_bin_images import convert_all_bin_to_image
from src.convert_word_docs import get_relevant_paragraphs, get_slides_info
from src.extract_pictures import extract_images_from_docx


# Path to the test .docx file and output directory
docx_file = "test_data/example_to_convert.docx"
docx = Document(docx_file)


def test_get_relevant_paragraphs():
    docx = Document(docx_file)

    # Call the function
    paragraphs = get_relevant_paragraphs(docx)

    expected_num_of_paragraphs = 35
    expected_first_paragraph_text = "1 P&C NO MOTOR"
    expected_last_paragraph_text = "2.31 Exit Module"

    assert len(paragraphs) == expected_num_of_paragraphs
    assert paragraphs[0]["text"] == expected_first_paragraph_text
    assert paragraphs[-1]["text"] == expected_last_paragraph_text


def get_header_1_paragraphs():

    # Call the function
    paragraphs = get_relevant_paragraphs(docx)

    expected_num_of_h1_paragraphs = 2
    expected_first_paragraph_text = "1 P&C NO MOTOR"
    expected_last_paragraph_text = "2 P&C NO MOTOR Quote & Policy Issues"

    assert len(paragraphs) == expected_num_of_h1_paragraphs
    assert paragraphs[0]["text"] == expected_first_paragraph_text
    assert paragraphs[-1]["text"] == expected_last_paragraph_text


def test_get_slides_info():  # Call the function
    images_folder = "tests/output/images_for_test_get_slides_info"

    # prepare the images for the test to run
    extract_images_from_docx(docx_file, images_folder)
    convert_all_bin_to_image(images_folder)

    # call the function
    slides_info = get_slides_info(docx, images_folder)

    expected_num_of_slides = 33
    expected_first_slide_image_path = os.path.join(
        images_folder, "word", "media", "image1.png"
    )
    expected_last_slide_image_path = os.path.join(
        images_folder, "word", "media", f"image{expected_num_of_slides}.png"
    )
    expected_second_slide_notes_start = "This course lasts about one hour "
    expected_second_slide_notes_end = "a policy cancellation "
    expected_last_slide_notes = "Congratulations! You have completed the module."

    assert len(slides_info) == expected_num_of_slides
    assert slides_info[0]["image_path"] == expected_first_slide_image_path
    assert slides_info[-1]["image_path"] == expected_last_slide_image_path
    assert (
        slides_info[1]["slide_notes"][: len(expected_second_slide_notes_start)]
        == expected_second_slide_notes_start
    )
    assert (
        slides_info[1]["slide_notes"][-len(expected_second_slide_notes_end) :]
        == expected_second_slide_notes_end
    )
    assert slides_info[-1]["slide_notes"] == expected_last_slide_notes
