"""These tests are designed to check whether the program can detect explicit and implicit hyperlinks within office files

Each test unzips and runs the hyperlink detection on each file before re-zipping it, as per DocuBleach's usual procedure

The output of the hyperlink detection function is compared to a list of the actual hyperlinks in the file

Sets are used for expected and detected hyperlink lists as the order of the detected hyperlinks does not matter

Just in case the hyperlink detection alters the original files, backups are made prior to and restored after, testing

All tests are written for and conducted using pytest.
"""

from docubleach.bleach import unzip_file, rezip_file, detect_ooxml_hyperlinks
from os import listdir, remove, rename
from shutil import copyfile


test_dir = "tests/test_files/ooxml_hyperlink_detection/"

actual_hyperlinks = {
        "https://anti-malware-alliance.org",
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
    test_file = f"{test_dir}word_document.docx"

    unzip_file(test_file)
    detected_hyperlinks = set(detect_ooxml_hyperlinks(test_file))
    rezip_file(test_file)

    assert detected_hyperlinks == actual_hyperlinks


def test_word_document_with_macros():
    test_file = f"{test_dir}word_document_with_macros.docm"

    unzip_file(test_file)
    detected_hyperlinks = set(detect_ooxml_hyperlinks(test_file))
    rezip_file(test_file)

    assert detected_hyperlinks == actual_hyperlinks


def test_word_template():
    test_file = f"{test_dir}word_template.dotx"

    unzip_file(test_file)
    detected_hyperlinks = set(detect_ooxml_hyperlinks(test_file))
    rezip_file(test_file)

    assert detected_hyperlinks == actual_hyperlinks


def test_word_template_with_macros():
    test_file = f"{test_dir}word_template_with_macros.dotm"

    unzip_file(test_file)
    detected_hyperlinks = set(detect_ooxml_hyperlinks(test_file))
    rezip_file(test_file)

    assert detected_hyperlinks == actual_hyperlinks


def test_powerpoint_presentation():
    test_file = f"{test_dir}powerpoint_presentation.pptx"

    unzip_file(test_file)
    detected_hyperlinks = set(detect_ooxml_hyperlinks(test_file))
    rezip_file(test_file)

    assert detected_hyperlinks == actual_hyperlinks


def test_powerpoint_presentation_with_macros():
    test_file = f"{test_dir}powerpoint_presentation_with_macros.pptm"

    unzip_file(test_file)
    detected_hyperlinks = set(detect_ooxml_hyperlinks(test_file))
    rezip_file(test_file)

    assert detected_hyperlinks == actual_hyperlinks


def test_powerpoint_show():
    test_file = f"{test_dir}powerpoint_show.ppsx"

    unzip_file(test_file)
    detected_hyperlinks = set(detect_ooxml_hyperlinks(test_file))
    rezip_file(test_file)

    assert detected_hyperlinks == actual_hyperlinks


def test_powerpoint_show_with_macros():
    test_file = f"{test_dir}powerpoint_show_with_macros.ppsm"

    unzip_file(test_file)
    detected_hyperlinks = set(detect_ooxml_hyperlinks(test_file))
    rezip_file(test_file)

    assert detected_hyperlinks == actual_hyperlinks


def test_powerpoint_template():
    test_file = f"{test_dir}powerpoint_template.potx"

    unzip_file(test_file)
    detected_hyperlinks = set(detect_ooxml_hyperlinks(test_file))
    rezip_file(test_file)

    assert detected_hyperlinks == actual_hyperlinks


def test_powerpoint_template_with_macros():
    test_file = f"{test_dir}powerpoint_template_with_macros.potm"

    unzip_file(test_file)
    detected_hyperlinks = set(detect_ooxml_hyperlinks(test_file))
    rezip_file(test_file)

    assert detected_hyperlinks == actual_hyperlinks
