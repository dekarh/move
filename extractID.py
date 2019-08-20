# -*- coding: utf-8 -*-
__author__ = 'Denis'

import sys, openpyxl
from mysql.connector import MySQLConnection
from datetime import datetime

from move_slots import IN_IDS, IN_NAMES
from lib import read_config

if len(sys.argv) < 1:
    print(datetime.now().strftime("%H:%M:%S"), 'В командной строке не указан файл Excel')
    sys.exit()
elif not sys.argv[1].endswith('.xlsx'):
    print(datetime.now().strftime("%H:%M:%S"), 'В командной строке не указан файл Excel')
    sys.exit()
wb = openpyxl.load_workbook(filename=sys.argv[1], read_only=True)
sheet = wb[wb.sheetnames[0]]
if not sheet.max_row:
    print(datetime.now().strftime("%H:%M:%S"), 'Файл Excel некорректно сохранен OpenPyxl. Откройте и пересохраните его')
    sys.exit()
keys = {}
last_cell = 0
dbconfig = read_config(filename='move.ini', section='mysql')
dbconn = MySQLConnection(**dbconfig)
cursor = dbconn.cursor()
#big_string = ''
tuples_contracts = []
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
            print(datetime.now().strftime("%H:%M:%S"),'В файле ' + sys.argv[1] + 'отсутствует колонка с ID')
            sys.exit()
    else:
        #big_string += "'" + row[keys[IN_IDS[0]]].value + "',"
        tuples_contracts.append((row[keys[IN_IDS[0]]].value,))
        if j and not (j % 1000):
            print(datetime.now().strftime("%H:%M:%S"), j, 'из', sheet.max_row, int(100 * j / sheet.max_row), '%')
            cursor.executemany('UPDATE saturn_crm.contracts SET inserted_date = "2019-08-20", '
                               'status_callcenter_code = 0 WHERE client_id = %s', tuples_contracts)
            dbconn.commit()
            tuples_contracts = []

if len(tuples_contracts):
    cursor.executemany('UPDATE saturn_crm.contracts SET inserted_date = "2019-08-20", '
                       'status_callcenter_code = 0 WHERE client_id = %s', tuples_contracts)
    dbconn.commit()
print(datetime.now().strftime("%H:%M:%S"), sheet.max_row, 'из', sheet.max_row, '100 %')


#big_string = big_string[:-1]
#print(big_string)



