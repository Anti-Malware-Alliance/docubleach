"""These tests check if the program can detect explicit and implicit hyperlinks within legacy office files

Each test runs the hyperlink detection function on each file

 The resulting list is then compared to a list of the actual hyperlinks in the file

Sets are used for expected and detected hyperlink lists as the order of the detected hyperlinks does not matter

Backups of the test files are made prior to testing so that the original files can be restored afterwards

All tests are written for and conducted using pytest.
"""

from docubleach.bleach import detect_bff_hyperlinks
from os import listdir, remove, rename
from shutil import copyfile


test_dir = "tests/test_files/bff_hyperlink_detection/"

actual_hyperlinks = {
        "https://anti-malware-alliance.org/",
        "https://patterbear.github.io/my-website",
        "https://github.com/Anti-Malware-Alliance"
}


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
    test_file = f"{test_dir}word_document.doc"

    detected_hyperlinks = set(detect_bff_hyperlinks(test_file))

    assert detected_hyperlinks == actual_hyperlinks


def test_excel_spreadsheet():
    test_file = f"{test_dir}excel_spreadsheet.xls"

    detected_hyperlinks = set(detect_bff_hyperlinks(test_file))

    assert detected_hyperlinks == actual_hyperlinks
