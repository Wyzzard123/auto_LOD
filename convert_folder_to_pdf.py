"""
Converts all .docx or .doc documents in a given folder to PDF in another folder.
"""

import sys
import os
import win32com.client
import re

wdFormatPDF = 17

# CHANGE THIS
# Folder with the word docs
IN_FOLDER = ("test_files")

# CHANGE THIS
# Folder to save the PDF docs to
OUT_FOLDER = ("pdfs")

# This code assumes that this python file is in the parent folder of the 
# IN_FOLDER and OUT_FOLDER
BASE_DIR = os.getcwd()

# Set this to True if you want to overwrite any already existing PDFs.
OVERWRITE_FILES = False

# Open microsoft word
word = win32com.client.Dispatch('Word.Application')

# Get a list of all folders.
in_files = list(os.scandir(IN_FOLDER))

for in_file in in_files:
    in_file_path = os.path.join(BASE_DIR, IN_FOLDER, in_file.name)

    # Check if the file is a .doc or .docx file. If it is not, print a message
    # stating that the file was skipped because it was not a doc or docx file.
    if re.match(r"(.+)(\.docx?)", in_file.name):
        out_file_name = re.match(r"(.+)(\.docx?)", in_file.name).group(1) + ".pdf"
        out_file_path = os.path.join(BASE_DIR, OUT_FOLDER, out_file_name)
    else:
        print(f"\nSKIPPED (not a word file):\n{in_file_path}")
    
    # Prevents you from overwriting existing files if OVERWRITE_FILES is False.
    if os.path.exists(out_file_path) and OVERWRITE_FILES is False:
        print(f"\nSKIPPED: {in_file_path}:\nFile already exists: {out_file_path} ")
    # Otherwise, convert the word file to a PDF.
    else:
        doc = word.Documents.Open(in_file_path)
        doc.SaveAs(out_file_path, FileFormat=wdFormatPDF)
        doc.Close()

word.Quit()