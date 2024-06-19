"""These tests are designed to test the tool's compatibility with all OOXML file formats.

Each test runs the bleaching function on their file and records the console output.

These are then compared against the correct outputs.

Valid files containing macros are restored to their original form after testing to ensure test repeatability.

All tests are written for and conducted using pytest.
"""
from subprocess import check_output
from os import remove, rename, listdir
from shutil import copyfile


prog_dir = "docubleach/"
test_dir = "tests\\test_files/"


def setup_module():
    for file in listdir(test_dir):
        copyfile(f"{test_dir}{file}", f"{test_dir}{file}.bak")


def teardown_module():
    for file in listdir(test_dir):
        if file[-4:] != '.bak':
            remove(f"{test_dir}{file}")
        else:
            rename(f"{test_dir}{file}", f"{test_dir}{file}"[:-4])


def test_word_document():
    output = check_output(f"python {prog_dir}bleach.py {test_dir}word_document.docx -c", encoding='utf-8')

    assert output == ""


def test_word_document_with_macros():
    output = check_output(f"python {prog_dir}bleach.py {test_dir}word_document_with_macros.docm -c",
                          encoding='utf-8')

    assert output == "Macros detected and removed.\n"


def test_word_template():
    output = check_output(f"python {prog_dir}bleach.py {test_dir}word_template.dotx -c", encoding='utf-8')

    assert output == ""


def test_word_template_with_macros():
    output = check_output(f"python {prog_dir}bleach.py {test_dir}word_template_with_macros.dotm -c",
                          encoding='utf-8')

    assert output == "Macros detected and removed.\n"


def test_powerpoint_presentation():
    output = check_output(f"python {prog_dir}bleach.py {test_dir}powerpoint_presentation.pptx -c",
                          encoding='utf-8')

    assert output == ""


def test_powerpoint_presentation_with_macros():
    output = check_output(f"python {prog_dir}bleach.py {test_dir}powerpoint_presentation_with_macros.pptm -c",
                          encoding='utf-8')

    assert output == "Macros detected and removed.\n"


def test_powerpoint_template():
    output = check_output(f"python {prog_dir}bleach.py {test_dir}powerpoint_template.potx -c",
                          encoding='utf-8')

    assert output == ""


def test_powerpoint_template_with_macros():
    output = check_output(f"python {prog_dir}bleach.py {test_dir}powerpoint_template_with_macros.potm -c",
                          encoding='utf-8')

    assert output == "Macros detected and removed.\n"


def test_powerpoint_show():
    output = check_output(f"python {prog_dir}bleach.py {test_dir}powerpoint_show.ppsx -c",
                          encoding='utf-8')

    assert output == ""


def test_powerpoint_show_with_macros():
    output = check_output(f"python {prog_dir}bleach.py {test_dir}powerpoint_show_with_macros.ppsm -c",
                          encoding='utf-8')

    assert output == "Macros detected and removed.\n"


def test_excel_spreadsheet():
    output = check_output(f"python {prog_dir}bleach.py {test_dir}excel_spreadsheet.xlsx -c",
                          encoding='utf-8')

    assert output == ""


def test_excel_spreadsheet_with_macros():
    output = check_output(f"python {prog_dir}bleach.py {test_dir}excel_spreadsheet_with_macros.xlsm -c",
                          encoding='utf-8')

    assert output == "Macros detected and removed.\n"


def test_excel_template():
    output = check_output(f"python {prog_dir}bleach.py {test_dir}excel_template.xltx -c", encoding='utf-8')

    assert output == ""


def test_excel_template_with_macros():
    output = check_output(f"python {prog_dir}bleach.py {test_dir}excel_template_with_macros.xltm -c",
                          encoding='utf-8')

    assert output == "Macros detected and removed.\n"
