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
from argparse import ArgumentParser
from os import rename, path, remove
from os.path import getsize
from zipfile import ZipFile
from shutil import make_archive, rmtree


supported_formats = [
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


FILESIZE_LIMIT = 209715200


def unzip_file(file):
    rename(file, file + ".zip")

    with ZipFile(file + ".zip", 'r') as zip_ref:
        zip_ref.extractall(file + "_temp")

    remove(file + ".zip")


def remove_macros(file, notify):
    macros_found = False

    if path.exists(file + "_temp/word/vbaProject.bin"):
        remove(file + "_temp/word/vbaProject.bin")
        macros_found = True

    if path.exists(file + "_temp/word/vbaData.xml"):
        remove(file + "_temp/word/vbaData.xml")
        macros_found = True

    if notify:
        if macros_found:
            print("Macros detected and removed.")


def rezip_file(file):
    make_archive(file, "zip", file + "_temp")
    rename(file + ".zip", file)
    rmtree(file + "_temp")


def validate_file(file):
    filetype = file.split(".")[-1].lower()

    if filetype in supported_formats:
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
        unzip_file(args.file)
        remove_macros(args.file, args.c)
        rezip_file(args.file)


if __name__ == "__main__":
    main()
