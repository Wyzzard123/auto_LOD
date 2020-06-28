"""
Autogenerates a List of Documents in docx format from a folder containing the appropriate documents.

Each document must be named <date in yyyy.mm.dd format> <time in 24h format (optional)> <Name of document>.<file_extension>

The table will have three columns:
1. Tab Number ('S/N') (list index + 1)
2. Date and Time (time for emails only) ('Date')
4. Name ('Description of Document')
"""

# TODO - Regex filter to capture the names of each

import docx
import os
import re
from collections import namedtuple
# import argparse

lod_files = []
LOD_File = namedtuple('LOD_File', 'index_number date time description extension') 
invalid_names = []


# Eg 2020.01.12 23.59 abcd.pdf OR
# Eg 2020.05.22 abcd.docx
file_name_pattern = r'^([0-9]{4}\.[0-9]{2}\.[0-9]{2})( )(([0-2][0-4]\.[0-5][0-9])( ))?(.+)((\.)(.+))$'

# CHANGE THESE TWO VARIABLES

# This is the folder where all your named files are located. You can set this to an absolute path.
lod_files_directory = './test_files'
# This is the name of the output file.
# You can also change this to an absolute directory path
word_doc_file_name = 'new_LOD.docx'


# TODO - Use argparse
# parser = argparse.ArgumentParser(description="Create a LOD file.")
# parser.add_argument('file directory', metavar='file_directory', type=str, help='A file directory to save the new file.', default='./pdf_files')
# parser.add_argument('file name', metavar='file_name', type=str, help='A file name or file path for the LOD file.', default='new_LOD.docx')
# print(parser.parse_args())

if not word_doc_file_name.endswith('.docx'):
    # If word doc file name does not have .docx at the end, add it.
    word_doc_file_name += ".docx"

# TODO - Sort differently. Currently follows folder sorting structure, which is by date, then time then alphabetical order. If there is no timestamp while other files have a timestamp on the same date, the file with no timestamp will appear at the end.
for index, entry in enumerate(os.scandir(lod_files_directory)):
    file_name = entry.name
    match = re.match(file_name_pattern, file_name)

    # If there is no regex name, prompt to rename the file and continue the loop
    if not match:
        invalid_names.append(entry.name)
        continue
    # Else, add each group to a tuple of format date, time, name
    elif match:
        index_number = index + 1
        # Extract from regex the relevant parameters
        date = match.group(1)
        time = match.group(4)
        description = match.group(6)

        # Saving this just in case
        extension = match.group(9)

        # Save a named tuple to the lod_files list
        lod_files.append(LOD_File(index_number, date, time, description, extension))

# no_of_files = len(lod_files)

# Alert user that the file names are wrongly formatted
if invalid_names:
    print("Rename the following files:")
    for invalid_name in invalid_names:
        print(invalid_name)

        
    
    lod_files.append(entry.name)


print(lod_files)

# Create a new  Microsoft Word document
lod_doc = docx.Document()

# Create a table with just headers first. Add rows as we iterate through loop.
table = lod_doc.add_table(1, 3)

# Make sure there are borders
table.style = 'TableGrid'

# Populate header row

heading_cells = table.rows[0].cells
heading_cells[0].text = 'S/N'
heading_cells[1].text = 'Date/Time'
heading_cells[2].text = 'Description'

# TODO - change this later
# # Either "yyyy.mm.dd" or "day month year"
# DATE_FORMAT = "day month year"

# # Either "24h" or "12h"
# TIME_FORMAT = "12h"

LIST_OF_MONTHS = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December"
]

for lod_file in lod_files:
    cells = table.add_row().cells

    # Adds a list number to S/N column
    cells[0].text = ""  # f"{lod_file.index_number}"
    cells[0].paragraphs[0].style = 'List Number'
    yyyymmdd_date = lod_file.date

    components_of_date = yyyymmdd_date.split(".")
    year = int(components_of_date[0])
    month = int(components_of_date[1])
    day = int(components_of_date[2])

    month_in_words = LIST_OF_MONTHS[month - 1]

    word_date = str(day) + " " + month_in_words + " " + str(year)

    if lod_file.time is not None:
        time_24h = lod_file.time
        components_of_time = time_24h.split(".")
        hour = int(components_of_time[0])
        
        # Minute is not changed to int because it would otherwise show things like "10.0 pm"
        minute = components_of_time[1]
        am_or_pm = ""
        if hour < 12:
            am_or_pm = "am"
            if hour != 0:
                hour_12h = hour
            elif hour == 0:
                hour_12h = 12
            
        elif int(hour) >= 12:
            am_or_pm = "pm"
            if hour != 12:
                hour_12h = hour - 12
            elif hour == 12:
                hour_12h = hour
        
        time_12h = f"{hour_12h}.{minute}{am_or_pm}"

    if lod_file.time is not None:
        cells[1].text = f"{word_date}\n{time_12h}"
        
    elif lod_file.time is None:
        cells[1].text = word_date
    cells[2].text = f"{lod_file.description}"

lod_doc.save(word_doc_file_name)

