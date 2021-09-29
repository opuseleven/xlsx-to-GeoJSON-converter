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

def convertcoords(str):
    print(str)
    c = str.strip('[]').split(', ')
    print(c)
    lon = c[0]
    lat = c[1]
    coords = {
        "lon": lon,
        "lat": lat
    }
    print(coords)
    return coords

def skip():
    return

for sheet in workbook.worksheets:
    state = sheet.title
    titlerow = sheet[1]
    colnames = []
    for cell in titlerow:
        colnames.append(cell.value)
    for row in sheet.iter_rows(min_row=2):
        if row[0] == None:
            skip()
        elif row[0] == '    ':
            skip()
        else:
            data = OrderedDict()
            c = 0
            while c < len(colnames):
                if colnames[c] == 'Coordinates':
                    if row[c].value == None:
                        data[colnames[c]] = row[c].value
                    elif row[c].value == ' ':
                        data[colnames[c]] = row[c].value
                    else:
                        data[colnames[c]] = convertcoords(row[c].value)
                else:
                    data[colnames[c]] = row[c].value
                c += 1
            datalist.append(data)

for obj in datalist:
    if obj['Name'] == None:
        datalist.remove(obj)
    elif obj['Coordinates'] == False:
        datalist.remove(obj)
    elif obj['Address'] == False:
        datalist.remove(obj)

j = json.dumps(datalist)

print('Writing JSON file...')

split_filename = os.path.splitext(filename)
new_filename = split_filename[0] + '.json'
with open(new_filename, 'w') as f:
    f.write(j)

print("Done!")
