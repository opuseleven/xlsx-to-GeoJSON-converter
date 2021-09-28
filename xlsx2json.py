import os
import sys
import openpyxl
from collections import OrderedDict
import json

if len(sys.argv) < 2:
    print('Usage: python3 xlsx2json.py filename.xlsx')
    sys.exit()

filename = sys.argv[1]
path = os.path.join(os.getcwd(), filename)

print(path)

workbook = openpyxl.load_workbook(path)

print('Converting to JSON...')

datalist = []

for sheet in workbook.worksheets:
    titlerow = sheet[1]
    colnames = []
    for cell in titlerow:
        colnames.append(cell.value)
    for row in sheet.iter_rows(min_row=2):
        data = OrderedDict()
        c = 0
        while c < len(colnames):
            data[colnames[c]] = row[c].value
            c += 1
        datalist.append(data)

j = json.dumps(datalist)

print('Writing JSON file...')

split_filename = os.path.splitext(filename)
new_filename = split_filename[0] + '.json'
with open(new_filename, 'w') as f:
    f.write(j)
