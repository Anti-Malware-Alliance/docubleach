"""This module removes any and all macros/dynamic content from MS Office files.

It is part of a suite of programs developed by the AntiMalware Alliance.

Visit https://github.com/Anti-Malware-Alliance for more details about our organisation and projects.

Feel free to contact benjamin.mcgregor2002@gmail.com for any questions regarding this module.
"""
from argparse import ArgumentParser
from os import rename, path, remove
from zipfile import ZipFile


# Unzip file function
# Converts file to a .zip archive and extracts it to file directory
def unzip_file(file):

    # Convert to zip archive
    rename(file, file + '.zip')

    # Extract contents into temporary folder
    with ZipFile(file + ".zip", 'r') as zip_ref:
        zip_ref.extractall(file + "_temp")

    # Revert file name
    rename(file + '.zip', file)


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


# Argument parser
parser = ArgumentParser()
parser.add_argument("file", help="file to be bleached")
parser.add_argument("-c", help="notify if macros or potentially dangerous content is found", action="store_true")
args = parser.parse_args()

# Bleaching
unzip_file(args.file)
remove_macros(args.file, args.c)
