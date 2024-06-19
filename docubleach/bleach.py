"""This module is designed to purge any and all macros and dynamic content from commonly used office formats.

VBA and OLE content in MS Office files can, and have sometimes been made to,
act as vehicles for malware delivery.

Microsoft has previously attempted to protect users from macros by disabling
them by default.

However, anybody is able to enable macros in an MS Office file before sending
them on to a potential victim.

This module enables users to simply and safely remove any and all
macros/dynamic content from MS Office files.

It is part of a suite of programs developed by the AntiMalware Alliance.

Visit https://github.com/Anti-Malware-Alliance for more details
about our organisation and projects.
"""
from argparse import ArgumentParser
from os import rename, path, remove
from os.path import getsize
from zipfile import ZipFile
from shutil import make_archive, rmtree
from olefile import OleFileIO

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
    "xls",
]

bff_macro_folders = [
    "VBA",
    "Macros",
    "_VBA_PROJECT_CUR",
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


def remove_bff_macros(file, notify):
    file_type = file.split(".")[-1].lower()
    macros_found = False

    if file_type == "doc" or file_type == "xls":
        streams = OleFileIO(file).listdir(streams=True)
        macro_streams = []

        for stream in streams:
            if stream[0] in bff_macro_folders:
                macro_streams.append(stream)

        ole = OleFileIO(file, write_mode=True)

        for macro_stream in macro_streams:
            macro_stream_size = ole.get_size(macro_stream)
            ole.write_stream(macro_stream, bytes(bytearray(macro_stream_size)))
        ole.close()

        if len(macro_streams) > 0:
            macros_found = True

    if file_type == "ppt":
        streams = OleFileIO(file).listdir(streams=True)
        # ppt logic here

    if notify and macros_found:
        print("Macros detected and removed.")


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

    if notify and macros_found:
        print("Macros detected and removed.")


def rezip_file(file):
    make_archive(file, "zip", file + "_temp")
    rename(file + ".zip", file)
    rmtree(file + "_temp")


def validate_file(file):
    filetype = file.split(".")[-1].lower()

    if filetype in ooxml_formats or filetype in bff_formats:
        """
        if getsize(file) < FILESIZE_LIMIT:
            return True
        else:
            print("File exceeds size limit.")
            return False"""
        return True
    else:
        print("Unsupported file format.")
        return False


def main():
    parser = ArgumentParser()
    parser.add_argument("file", help="file to be bleached")
    parser.add_argument("-c", help="notify if macros or potentially dangerous "
                                   "content is found", action="store_true")
    args = parser.parse_args()

    if validate_file(args.file):
        remove_macros(args.file, args.c)


if __name__ == "__main__":
    main()
