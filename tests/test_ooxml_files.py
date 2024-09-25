"""These tests are designed to test the tool's compatibility with all OOXML file formats.

Each test runs the 'bleach' on their respective file.

The 'bleached' files are then scanned for macros using the 'detect_macros' function

Valid files containing macros are restored to their original form after testing to ensure test repeatability.

All tests are written for and conducted using pytest.
"""
from os import listdir, remove, rename
from shutil import copyfile
from docubleach.bleach import detect_macros, remove_macros


test_dir = "tests/test_files/ooxml_files/"


def setup_module():
    for file in listdir(test_dir):
        copyfile(test_dir + file, test_dir + file + ".bak")


def teardown_module():
    for file in listdir(test_dir):
        if file.split(".")[-1] != "bak":
            remove(test_dir + file)

    for file in listdir(test_dir):
        if file.split(".")[-1] == "bak":
            rename(test_dir + file, test_dir + file[:-4])


def test_word_document():
    test_file = f"{test_dir}word_document.docx"

    remove_macros(test_file)

    assert detect_macros(test_file) is False


def test_word_document_with_macros():
    test_file = f"{test_dir}word_document_with_macros.docm"

    remove_macros(test_file)

    assert detect_macros(test_file) is False


def test_word_template():
    test_file = f"{test_dir}word_template.dotx"

    remove_macros(test_file)

    assert detect_macros(test_file) is False


def test_word_template_with_macros():
    test_file = f"{test_dir}word_template_with_macros.dotm"

    remove_macros(test_file)

    assert detect_macros(test_file) is False


def test_powerpoint_presentation():
    test_file = f"{test_dir}powerpoint_presentation.pptx"

    remove_macros(test_file)

    assert detect_macros(test_file) is False


def test_powerpoint_presentation_with_macros():
    test_file = f"{test_dir}powerpoint_presentation_with_macros.pptm"

    remove_macros(test_file)

    assert detect_macros(test_file) is False


def test_powerpoint_template():
    test_file = f"{test_dir}powerpoint_template.potx"

    remove_macros(test_file)

    assert detect_macros(test_file) is False


def test_powerpoint_template_with_macros():
    test_file = f"{test_dir}powerpoint_template_with_macros.potm"

    remove_macros(test_file)

    assert detect_macros(test_file) is False


def test_powerpoint_show():
    test_file = f"{test_dir}powerpoint_show.ppsx"

    remove_macros(test_file)

    assert detect_macros(test_file) is False


def test_powerpoint_show_with_macros():
    test_file = f"{test_dir}powerpoint_show_with_macros.ppsm"

    remove_macros(test_file)

    assert detect_macros(test_file) is False


def test_excel_spreadsheet():
    test_file = f"{test_dir}excel_spreadsheet.xlsx"

    remove_macros(test_file)

    assert detect_macros(test_file) is False


def test_excel_spreadsheet_with_macros():
    test_file = f"{test_dir}excel_spreadsheet_with_macros.xlsm"

    remove_macros(test_file)

    assert detect_macros(test_file) is False


def test_excel_template():
    test_file = f"{test_dir}excel_template.xltx"

    remove_macros(test_file)

    assert detect_macros(test_file) is False


def test_excel_template_with_macros():
    test_file = f"{test_dir}excel_template_with_macros.xltm"

    remove_macros(test_file)

    assert detect_macros(test_file) is False
