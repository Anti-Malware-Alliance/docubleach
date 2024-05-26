"""This module removes any and all macros/dynamic content from MS Office files.

VBA and OLE content in MS Office files can, and have sometimes been made to, act as vehicles for malware delivery.

Microsoft has previously attempted to protect users from macros by disabling them by default.

However, anybody is able to enable macros on an MS Office file before sending them on to a potential victim.

This module enables users to simply and safely remove any and all macros/dynamic content from MS Office files.

It is part of a suite of programs developed by the AntiMalware Alliance.

Visit https://github.com/Anti-Malware-Alliance for more details about our organisation and projects.
"""
from argparse import ArgumentParser
from os import rename, path, remove
from zipfile import ZipFile
from shutil import make_archive, rmtree


# List of supported file types
supported_formats = ["docx", "docm"]

# Unzip file function
# Converts file to a .zip archive and extracts it to file directory
def unzip_file(file):

    # Convert to zip archive
    rename(file, file + ".zip")

    # Extract contents into temporary folder
    with ZipFile(file + ".zip", 'r') as zip_ref:
        zip_ref.extractall(file + "_temp")

    # Delete original file
    remove(file + ".zip")


# Remove macros function
# Checks for macro files and deletes them if found
def remove_macros(file, notify):
    macros_found = False

    # Deletes macro binary file if found
    if path.exists(file + "_temp/word/vbaProject.bin"):
        remove(file + "_temp/word/vbaProject.bin")
        macros_found = True

    # Deletes macro XML file if found
    if path.exists(file + "_temp/word/vbaData.xml"):
        remove(file + "_temp/word/vbaData.xml")
        macros_found = True

    # Notifies user of macro status if '-c' flag is present
    if notify:
        if macros_found:
            print("Macros detected and removed.")
        else:
            print("No macros detected.")


# Rezip function
# Re-zips the unzipped file and restores its original file extension
def rezip_file(file):

    # Zip bleached folder
    make_archive(file, "zip", file + "_temp")

    # Convert back into original file format
    rename(file + ".zip", file)

    # Delete temporary folder
    rmtree(file + "_temp")


# Validate file format function
# Checks to see if file is supported
def validate_format(file):
    if file.split(".")[-1].lower() in supported_formats:
        return True
    else:
        return False


# Main function
# Initialises the argument parser and calls functions to bleach the file
def main():

    # Argument parser
    parser = ArgumentParser()
    parser.add_argument("file", help="file to be bleached")
    parser.add_argument("-c", help="notify if macros or potentially dangerous content is found", action="store_true")
    args = parser.parse_args()

    # Validate file format
    if validate_format(args.file):
        # Bleaching
        unzip_file(args.file)
        remove_macros(args.file, args.c)
        rezip_file(args.file)
    else:
        print("Unsupported file format.")


if __name__ == "__main__":
    main()
