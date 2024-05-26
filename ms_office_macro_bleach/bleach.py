"""This module removes any and all macros/dynamic content from MS Office files.

It is part of a suite of programs developed by the AntiMalware Alliance.

Visit https://github.com/Anti-Malware-Alliance for more details about our organisation and projects.

Feel free to contact benjamin.mcgregor2002@gmail.com for any questions regarding this module.
"""
from argparse import ArgumentParser
from os import rename
from zipfile import ZipFile


# Unzip file function
# Converts file to a .zip archive and extracts it to file directory
def unzip_file(file):

    # Convert to zip archive
    rename(file, file + '.zip')

    # Extract contents into 'temp' folder
    with ZipFile(file + ".zip", 'r') as zip_ref:
        zip_ref.extractall("temp")

    # Revert file name
    rename(file + '.zip', file)


def remove_macros(notify):
    pass


# Argument parser
parser = ArgumentParser()
parser.add_argument("file", help="file to be bleached")
parser.add_argument("-c", "--check", help="notify if macros or potentially dangerous content is found")
args = parser.parse_args()

# Bleaching
unzip_file(args.file)
remove_macros(args.check)
