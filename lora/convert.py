import configparser
import xlwt
import csv
import os
from io import StringIO


config = configparser.ConfigParser()
config.read('convert.ini')
config = config['lora']

last_modified = None

while True:
    try:
        new_last_modified = os.stat(config['input']).st_mtime
    except Exception:
        continue
    if not last_modified or new_last_modified > last_modified:
        last_modified = new_last_modified + float(config['trigger']) / 1000
        with open(config['input'], 'rb') as csvfile:
            xlsout = []
            csvfile = StringIO(
                csvfile.read().decode(errors="ignore").replace('\x00', '')
            )
            workbook = xlwt.Workbook()
            worksheet = workbook.add_sheet('Mappe 1')
            vabreader = csv.reader(csvfile, delimiter=';')

            csvin = {}

            for i, row in enumerate(vabreader):
                row = [r.replace('?', '').strip() for r in row]
                if row:
                    csvin[row[0]] = row[1:]

            if config['mapping']:
                mapping = config['mapping'].split()
            else:
                mapping = csvin.keys()

            for i, key in enumerate(mapping):
                worksheet.write(i, 0, key)
                if key in csvin:
                    for j, cell in enumerate(csvin[key]):
                        worksheet.write(i, j+1, cell)
                else:
                    for j in range(0, len(list(csvin.values())[0])):
                        worksheet.write(i, j+1, '0')

            with open(config['output'], 'bw') as f:
                workbook.save(f)
