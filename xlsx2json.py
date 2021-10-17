import os
import sys
import openpyxl
from collections import OrderedDict
import json

if len(sys.argv) < 2:
    print('Usage: python3 xlsx2json.py filename.xlsx')
    sys.exit()

if len(sys.argv) > 3:
    print('Usage: python3 xlsx2json.py filename.xlsx')
    sys.exit()

filename = sys.argv[1]
path = os.path.join(os.getcwd(), filename)

print(path)

workbook = openpyxl.load_workbook(path)

print('Converting to GeoJSON...')

datalist = []

def convertcoords(str):
    c = str.strip('[]').split(', ')
    lon = c[0]
    lat = c[1]
    coords = {
        "lon": lon,
        "lat": lat
    }
    return coords

for sheet in workbook.worksheets:
    state = sheet.title
    titlerow = sheet[1]
    colnames = []
    counter = 0
    coordscol = -1
    for cell in titlerow:
        colnames.append(cell.value)
        if cell.value == 'Coordinates':
            coordscol = counter
        counter += 1
    if coordscol == -1:
        print("Error: Couldn't identify \"Coordinates\" column.")
        break
    for row in sheet.iter_rows(min_row=2):
        if row[0].value:
            if row[coordscol].value:
                if not row[coordscol].value.startswith(' '):
                    data = OrderedDict()
                    data['type'] = "Feature"
                    data['geometry'] = {
                    "type": "Point",
                    "coordinates": convertcoords(row[coordscol].value)
                    }
                    propsdata = OrderedDict()
                    c = 0
                    while c < len(colnames):
                        if c != coordscol:
                            propsdata[colnames[c]] = row[c].value
                        c += 1
                    data['properties'] = propsdata
                    datalist.append(data)

j = json.dumps(datalist)

print('Writing JSON file...')

split_filename = os.path.splitext(filename)
new_filename = split_filename[0] + '.json'
with open(new_filename, 'w') as f:
    f.write(j)

print("Done!")
