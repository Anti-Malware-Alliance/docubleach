"""These tests are designed to test the tool's compatibility with all OOXML file formats.

Each test runs the bleaching function on their file and records the console output.

These are then compared against the correct outputs.

Valid files containing macros are restored to their original form after testing to ensure test repeatability.

All tests are written for and conducted using pytest.
"""
from os import remove, rename
from shutil import copyfile
from subprocess import check_output

prog_dir = "docubleach/"
test_dir = "tests/test_files/"


def test_word_document():
    output = check_output(f"python {prog_dir}bleach.py {test_dir}word_document.docx -c", encoding='utf-8')

    assert output == ""


def test_word_document_with_macros():
    copyfile(f"{test_dir}word_document_with_macros.docm", f"{test_dir}word_document_with_macros.docm.bak")

    output = check_output(f"python {prog_dir}bleach.py {test_dir}word_document_with_macros.docm -c",
                          encoding='utf-8')

    remove(f"{test_dir}word_document_with_macros.docm")
    rename(f"{test_dir}word_document_with_macros.docm.bak", f"{test_dir}word_document_with_macros.docm")

    assert output == "Macros detected and removed.\n"


def test_word_template():
    output = check_output(f"python {prog_dir}bleach.py {test_dir}word_template.dotx -c", encoding='utf-8')

    assert output == ""


def test_word_template_with_macros():
    copyfile(f"{test_dir}word_template_with_macros.dotm", f"{test_dir}word_template_with_macros.dotm.bak")

    output = check_output(f"python {prog_dir}bleach.py {test_dir}word_template_with_macros.dotm -c",
                          encoding='utf-8')

    remove(f"{test_dir}word_template_with_macros.dotm")
    rename(f"{test_dir}word_template_with_macros.dotm.bak", f"{test_dir}word_template_with_macros.dotm")

    assert output == "Macros detected and removed.\n"


def test_powerpoint_presentation():
    output = check_output(f"python {prog_dir}bleach.py {test_dir}powerpoint_presentation.pptx -c",
                          encoding='utf-8')

    assert output == ""


def test_powerpoint_presentation_with_macros():
    copyfile(f"{test_dir}powerpoint_presentation_with_macros.pptm",
             f"{test_dir}powerpoint_presentation_with_macros.pptm.bak")

    output = check_output(f"python {prog_dir}bleach.py {test_dir}powerpoint_presentation_with_macros.pptm -c",
                          encoding='utf-8')

    remove(f"{test_dir}powerpoint_presentation_with_macros.pptm")
    rename(f"{test_dir}powerpoint_presentation_with_macros.pptm.bak",
           f"{test_dir}powerpoint_presentation_with_macros.pptm")

    assert output == "Macros detected and removed.\n"


def test_powerpoint_template():
    output = check_output(f"python {prog_dir}bleach.py {test_dir}powerpoint_template.potx -c",
                          encoding='utf-8')

    assert output == ""


def test_powerpoint_template_with_macros():
    copyfile(f"{test_dir}powerpoint_template_with_macros.potm",
             f"{test_dir}powerpoint_template_with_macros.potm.bak")

    output = check_output(f"python {prog_dir}bleach.py {test_dir}powerpoint_template_with_macros.potm -c",
                          encoding='utf-8')

    remove(f"{test_dir}powerpoint_template_with_macros.potm")
    rename(f"{test_dir}powerpoint_template_with_macros.potm.bak",
           f"{test_dir}powerpoint_template_with_macros.potm")

    assert output == "Macros detected and removed.\n"


def test_powerpoint_show():
    output = check_output(f"python {prog_dir}bleach.py {test_dir}powerpoint_show.ppsx -c",
                          encoding='utf-8')

    assert output == ""


def test_powerpoint_show_with_macros():
    copyfile(f"{test_dir}powerpoint_show_with_macros.ppsm", f"{test_dir}powerpoint_show_with_macros.ppsm.bak")

    output = check_output(f"python {prog_dir}bleach.py {test_dir}powerpoint_show_with_macros.ppsm -c",
                          encoding='utf-8')

    remove(f"{test_dir}powerpoint_show_with_macros.ppsm")
    rename(f"{test_dir}powerpoint_show_with_macros.ppsm.bak", f"{test_dir}powerpoint_show_with_macros.ppsm")

    assert output == "Macros detected and removed.\n"


def test_excel_spreadsheet():
    output = check_output(f"python {prog_dir}bleach.py {test_dir}excel_spreadsheet.xlsx -c",
                          encoding='utf-8')

    assert output == ""


def test_excel_spreadsheet_with_macros():
    copyfile(f"{test_dir}excel_spreadsheet_with_macros.xlsm",
             f"{test_dir}excel_spreadsheet_with_macros.xlsm.bak")

    output = check_output(f"python {prog_dir}bleach.py {test_dir}excel_spreadsheet_with_macros.xlsm -c",
                          encoding='utf-8')

    remove(f"{test_dir}excel_spreadsheet_with_macros.xlsm")
    rename(f"{test_dir}excel_spreadsheet_with_macros.xlsm.bak", f"{test_dir}excel_spreadsheet_with_macros.xlsm")

    assert output == "Macros detected and removed.\n"


def test_excel_template():
    output = check_output(f"python {prog_dir}bleach.py {test_dir}excel_template.xltx -c", encoding='utf-8')

    assert output == ""


def test_excel_template_with_macros():
    copyfile(f"{test_dir}excel_template_with_macros.xltm", f"{test_dir}excel_template_with_macros.xltm.bak")

    output = check_output(f"python {prog_dir}bleach.py {test_dir}excel_template_with_macros.xltm -c",
                          encoding='utf-8')

    remove(f"{test_dir}excel_template_with_macros.xltm")
    rename(f"{test_dir}excel_template_with_macros.xltm.bak", f"{test_dir}excel_template_with_macros.xltm")

    assert output == "Macros detected and removed.\n"
