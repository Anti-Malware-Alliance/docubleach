"""These tests are designed to test the tool's compatibility with all OOXML file formats.

Each test runs the bleaching function on their file and records the console output.

These are then compared against the correct outputs.

Valid files containing macros are restored to their original form after testing to ensure test repeatability.

All tests are written for and conducted using pytest.
"""
from subprocess import check_output
from os import remove, rename
from shutil import copyfile


program_dir = "../ms_office_macro_bleach/"


def test_word_document():
    output = check_output(f"python {program_dir}bleach.py test_files/word_document.docx -c", encoding='utf-8')

    assert output == ""


def test_word_document_with_macros():
    copyfile("test_files/word_document_with_macros.docm", "test_files/word_document_with_macros.docm.bak")

    output = check_output(f"python {program_dir}bleach.py test_files/word_document_with_macros.docm -c", encoding='utf-8')

    remove("test_files/word_document_with_macros.docm")
    rename("test_files/word_document_with_macros.docm.bak", "test_files/word_document_with_macros.docm")

    assert output == "Macros detected and removed.\n"


def test_word_template():
    output = check_output(f"python {program_dir}bleach.py test_files/word_template.dotx -c", encoding='utf-8')

    assert output == ""


def test_word_template_with_macros():
    copyfile("test_files/word_template_with_macros.dotm", "test_files/word_template_with_macros.dotm.bak")

    output = check_output(f"python {program_dir}bleach.py test_files/word_template_with_macros.dotm -c", encoding='utf-8')

    remove("test_files/word_template_with_macros.dotm")
    rename("test_files/word_template_with_macros.dotm.bak", "test_files/word_template_with_macros.dotm")

    assert output == "Macros detected and removed.\n"


def test_powerpoint_presentation():
    output = check_output(f"python {program_dir}bleach.py test_files/powerpoint_presentation.pptx -c", encoding='utf-8')

    assert output == ""


def test_powerpoint_presentation_with_macros():
    copyfile("test_files/powerpoint_presentation_with_macros.pptm", "test_files/powerpoint_presentation_with_macros.pptm.bak")

    output = check_output(f"python {program_dir}bleach.py test_files/powerpoint_presentation_with_macros.pptm -c", encoding='utf-8')

    remove("test_files/powerpoint_presentation_with_macros.pptm")
    rename("test_files/powerpoint_presentation_with_macros.pptm.bak", "test_files/powerpoint_presentation_with_macros.pptm")

    assert output == "Macros detected and removed.\n"


def test_powerpoint_template():
    output = check_output(f"python {program_dir}bleach.py test_files/powerpoint_template.potx -c", encoding='utf-8')

    assert output == ""


def test_powerpoint_template_with_macros():
    copyfile("test_files/powerpoint_template_with_macros.potm", "test_files/powerpoint_template_with_macros.potm.bak")

    output = check_output(f"python {program_dir}bleach.py test_files/powerpoint_template_with_macros.potm -c", encoding='utf-8')

    remove("test_files/powerpoint_template_with_macros.potm")
    rename("test_files/powerpoint_template_with_macros.potm.bak", "test_files/powerpoint_template_with_macros.potm")

    assert output == "Macros detected and removed.\n"


def test_powerpoint_show():
    output = check_output(f"python {program_dir}bleach.py test_files/powerpoint_show.ppsx -c", encoding='utf-8')

    assert output == ""


def test_powerpoint_show_with_macros():
    copyfile("test_files/powerpoint_show_with_macros.ppsm", "test_files/powerpoint_show_with_macros.ppsm.bak")

    output = check_output(f"python {program_dir}bleach.py test_files/powerpoint_show_with_macros.ppsm -c", encoding='utf-8')

    remove("test_files/powerpoint_show_with_macros.ppsm")
    rename("test_files/powerpoint_show_with_macros.ppsm.bak", "test_files/powerpoint_show_with_macros.ppsm")

    assert output == "Macros detected and removed.\n"


def test_excel_spreadsheet():
    output = check_output(f"python {program_dir}bleach.py test_files/excel_spreadsheet.xlsx -c", encoding='utf-8')

    assert output == ""


def test_excel_spreadsheet_with_macros():
    copyfile("test_files/excel_spreadsheet_with_macros.xlsm", "test_files/excel_spreadsheet_with_macros.xlsm.bak")

    output = check_output(f"python {program_dir}bleach.py test_files/excel_spreadsheet_with_macros.xlsm -c", encoding='utf-8')

    remove("test_files/excel_spreadsheet_with_macros.xlsm")
    rename("test_files/excel_spreadsheet_with_macros.xlsm.bak", "test_files/excel_spreadsheet_with_macros.xlsm")

    assert output == "Macros detected and removed.\n"


def test_excel_template():
    output = check_output(f"python {program_dir}bleach.py test_files/excel_template.xltx -c", encoding='utf-8')

    assert output == ""


def test_excel_template_with_macros():
    copyfile("test_files/excel_template_with_macros.xltm", "test_files/excel_template_with_macros.xltm.bak")

    output = check_output(f"python {program_dir}bleach.py test_files/excel_template_with_macros.xltm -c", encoding='utf-8')

    remove("test_files/excel_template_with_macros.xltm")
    rename("test_files/excel_template_with_macros.xltm.bak", "test_files/excel_template_with_macros.xltm")

    assert output == "Macros detected and removed.\n"
