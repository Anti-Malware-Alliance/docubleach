"""These tests are designed to check file validation when various valid and invalid files are given as arguments.

Each test runs the bleaching function on their file and records the console output.

These are then compared against the correct outputs.

For the invalid file size test, a temporary file exceeding the file size limit is generated.

After output is recorded, it is subsequently deleted as to not occupy storage unnecessarily.

Valid files containing macros are restored to their original form after testing to ensure test repeatability.

This is because the purpose of these tests is to check file validation, not macro removal.

All tests are written for and conducted using pytest.
"""

from subprocess import check_output
from os import remove, rename, listdir
from shutil import copyfile


prog_dir = "docubleach/"
test_dir = "tests/test_files/"


def setup_module():
    for file in listdir(test_dir):
        copyfile(f"{test_dir}{file}", f"{test_dir}{file}.bak")


def teardown_module():
    for file in listdir(test_dir):
        if file[-4:] != '.bak':
            remove(f"{test_dir}{file}")
        else:
            rename(f"{test_dir}{file}", f"{test_dir}{file}"[:-4])


def test_valid_file_with_macros():
    output = check_output(["python", f"{prog_dir}bleach.py", f"{test_dir}valid_file_with_macros.docm"],
                          encoding='utf-8')

    assert output == ""


def test_word_document_with_macros():
    output = check_output(["python", f"{prog_dir}bleach.py", f"{test_dir}word_document_with_macros.docm", "-c"],
                          encoding='utf-8')

    assert output == "Macros detected and removed.\n"


def test_valid_file_with_macros_with_check():
    output = check_output(["python", f"{prog_dir}bleach.py", f"{test_dir}valid_file_with_macros_check.docm", "-c"],
                          encoding='utf-8')

    assert output == "Macros detected and removed.\n"


def test_valid_file_without_macros():
    output = check_output(["python", f"{prog_dir}bleach.py", f"{test_dir}valid_file_without_macros.docx"],
                          encoding='utf-8')

    assert output == ""


def test_invalid_file_type():
    output = check_output(["python", f"{prog_dir}bleach.py", f"{test_dir}invalid_file_type.txt"], encoding='utf-8')

    assert output == "Unsupported file format.\n"


def test_invalid_file_size():

    # Create temporary file exceeding 200MB limit
    with open(f"{test_dir}invalid_file_size.docx", "wb") as out:
        out.truncate(262144000)

    output = check_output(["python", f"{prog_dir}bleach.py", f"{test_dir}invalid_file_size.docx"], encoding='utf-8')

    remove(f"{test_dir}invalid_file_size.docx")

    assert output == "File exceeds size limit.\n"
