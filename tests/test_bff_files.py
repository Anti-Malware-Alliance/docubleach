"""These tests are designed to test the tool's compatibility with supported BFF file formats.

Each test runs the 'bleach' on their respective file.

The 'bleached' files are then scanned for macros using the 'detect_macros' function

Valid files containing macros are restored to their original form after testing to ensure test repeatability.

All tests are written for and conducted using pytest.
"""
from os import listdir, remove, rename
from shutil import copyfile

from docubleach.bleach import detect_macros, remove_macros


test_dir = "tests/test_files/bff_files/"


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


def test_legacy_word_document():
    test_file = f"{test_dir}legacy_word_document.doc"

    remove_macros(test_file)

    assert detect_macros(test_file) is False


def test_legacy_excel_spreadsheet():
    test_file = f"{test_dir}legacy_excel_spreadsheet.xls"

    remove_macros(test_file)

    assert detect_macros(test_file) is False
