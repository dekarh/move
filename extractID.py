# -*- coding: utf-8 -*-
__author__ = 'Denis'

import sys, openpyxl

from move_slots import IN_IDS, IN_NAMES

if len(sys.argv) < 1:
    print('В командной строке не указан файл Excel')
    sys.exit()
elif not sys.argv[1].endswith('.xlsx'):
    print('В командной строке не указан файл Excel')
    sys.exit()
wb = openpyxl.load_workbook(filename=sys.argv[1], read_only=True)
sheet = wb[wb.sheetnames[0]]
if not sheet.max_row:
    print('Файл Excel некорректно сохранен OpenPyxl. Откройте и пересохраните его')
    sys.exit()
keys = {}
last_cell = 0
big_string = ''
for j, row in enumerate(sheet.rows):
    if j == 0:
        for k, cell in enumerate(row):  # Проверяем, чтобы был client_id
            if str(cell.value).upper() in IN_IDS:
                keys[IN_IDS[0]] = k
        if len(keys) > 0:
            for k, cell in enumerate(row):
                for n, name in enumerate(IN_NAMES):
                    if n == 0:
                        continue
                    if cell.value != None:
                        if str(cell.value).strip() != '':
                            last_cell = k
                            if str(cell.value).upper() == name:
                                keys[name] = k

        else:
            print('В файле ' + sys.argv[1] + 'отсутствует колонка с ID')
            sys.exit()
    else:
        big_string += "'" + row[keys[IN_IDS[0]]].value + "',"
big_string = big_string[:-1]
print(big_string)



