"""This module removes any and all macros/dynamic content from MS Office files.

VBA and OLE content in MS Office files can, and have sometimes been made to, act as vehicles for malware delivery.

Microsoft has previously attempted to protect users from macros by disabling them by default.

However, anybody is able to enable macros in an MS Office file before sending them on to a potential victim.

This module enables users to simply and safely remove any and all macros/dynamic content from MS Office files.

It converts the given file into a '.zip' archive, unzips it, and deletes the files containing macro data.

It then re-zips the unzipped archive and reverts it to its original file format.

It is part of a suite of programs developed by the AntiMalware Alliance.

Visit https://github.com/Anti-Malware-Alliance for more details about our organisation and projects.
"""
from sys import platform
from argparse import ArgumentParser
from os import rename, path, remove
from os.path import getsize
from zipfile import ZipFile
from shutil import make_archive, rmtree

if platform == "win32":
    from win32com import client

ooxml_formats = [
    "docx",
    "docm",
    "dotx",
    "dotm",
    "pptx",
    "pptm",
    "potx",
    "potm",
    "ppsx",
    "ppsm",
    "xlsx",
    "xlsm",
    "xltx",
    "xltm",
]

ooxml_macro_folders = {
    "do": "word",
    "pp": "ppt",
    "po": "ppt",
    "xl": "xl",
}

bff_formats = [
    "doc",
    "ppt",
    "xls",
]

FILESIZE_LIMIT = 209715200


def unzip_file(file):
    rename(file, file + ".zip")

    with ZipFile(file + ".zip", 'r') as zip_ref:
        zip_ref.extractall(file + "_temp")

    remove(file + ".zip")


def remove_macros(file, notify):
    file_type = file.split(".")[-1].lower()

    if file_type in ooxml_formats:
        unzip_file(file)
        remove_ooxml_macros(file, notify)
        rezip_file(file)

    if file_type in bff_formats:
        remove_bff_macros(file, notify)


def remove_ooxml_macros(file, notify):
    macros_found = False
    file_type = file.split(".")[-1].lower()

    macro_folder = ooxml_macro_folders.get(file_type[:2])

    if path.exists(file + f"_temp/{macro_folder}/vbaProject.bin"):
        remove(file + f"_temp/{macro_folder}/vbaProject.bin")
        macros_found = True

    if path.exists(file + f"_temp/{macro_folder}/vbaData.xml"):
        remove(file + f"_temp/{macro_folder}/vbaData.xml")
        macros_found = True

    if notify:
        if macros_found:
            print("Macros detected and removed.")


def rezip_file(file):
    make_archive(file, "zip", file + "_temp")
    rename(file + ".zip", file)
    rmtree(file + "_temp")


def remove_bff_macros(file, notify):
    file_type = file.split(".")[-1].lower()
    input_file = path.abspath(file)
    output_file = path.abspath(file + ".tmp")

    if file_type == "doc":
        app = client.Dispatch("Word.Application")
        app.Visible = False
        output_type = 12
        office_file = app.Documents.Open(input_file)
    elif file_type == "ppt":
        app = client.Dispatch("PowerPoint.Application")
        output_type = 24
        office_file = app.Presentations.Open(input_file, WithWindow=False)
    elif file_type == "xls":
        app = client.Dispatch("Excel.Application")
        app.Visible = False
        output_type = 51
        office_file = app.Workbooks.Open(input_file)
    else:
        return

    office_file.SaveAs(output_file, output_type)
    office_file.Close()
    app.Quit()

    remove(input_file)
    rename(output_file, input_file + 'x')


def validate_file(file):
    filetype = file.split(".")[-1].lower()

    if filetype in ooxml_formats or filetype in bff_formats:
        if getsize(file) < FILESIZE_LIMIT:
            return True
        else:
            print("File exceeds size limit.")
            return False
    else:
        print("Unsupported file format.")
        return False


def main():
    parser = ArgumentParser()
    parser.add_argument("file", help="file to be bleached")
    parser.add_argument("-c", help="notify if macros or potentially dangerous content is found", action="store_true")
    args = parser.parse_args()

    if validate_file(args.file):
        remove_macros(args.file, args.c)


if __name__ == "__main__":
    main()
