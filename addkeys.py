#! python3
# a script to add a key column and simple key values to an xlsx document.

import os
import sys
import openpyxl

if len(sys.argv) < 2:
    print('Usage: python3 addkeys.py filename.xlsx')
    sys.exit()

filename = sys.argv[1]
path = os.path.join(os.getcwd(), filename)

print(path)

workbook = openpyxl.load_workbook(path)

key = 1

for sheet in workbook.worksheets:
    titlerow = sheet[1]
    keycol = -1
    for cell in titlerow:
        if not cell.value:
            keycol = cell.column
            break
    if keycol == -1:
        print("Error: no blank column")
        sys.exit()
    titlerow[keycol].value = 'Key'
    for row in sheet.iter_rows(min_row=2):
        if row[0]:
            row[keycol].value = key
            key += 1

split_filename = os.path.splitext(filename)
new_filename = split_filename[0] + '-keys' + split_filename[1]
workbook.save(new_filename)
print("Key values added.")
