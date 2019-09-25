# Конвертация-фильтрация файла импорта ОПС для импорта в Еву

import sys
import os
import csv
import re

import openpyxl
from openpyxl import Workbook

RUS = []

RUS_MINUS = ['Имя',
'Имя_при_рождении',
'Фамилия',
'Фамилия_при_рождении']

RUS_SP_MINUS = ['Отчество',
'Отчество_при_рождении',
'Страна_рождения',
'Область_рождения',
'Район_рождения']

CIFR_MINUS = [
'Дата_рождения',
'Паспорт_дата',
'Паспорт_Код подразделения']

CIFR = ['Пол(0_мужской,1_женский)',
'Паспорт_серия',
'Паспорт_номер',
'Адрес_регистрации_Индекс',
'Адрес_проживания_Индекс',
'Мобильный_телефон',
'Телефон_родственников',
'Телефон_домашний']

FIELDS = {
    'Фамилия': 'p_surname',
    'Имя': 'p_name',
    'Отчество': 'p_lastname',
    'Фамилия_при_рождении': 'b_surname',
    'Имя_при_рождении': 'b_name',
    'Отчество_при_рождении': 'b_lastname',
    'Пол(0_мужской,1_женский)': 'gender',
    'Дата_рождения': 'b_date',
    'Страна_рождения': 'b_country',
    'Область_рождения': 'b_region',
    'Район_рождения': 'b_district',
    'Город_рождения': 'b_place',
    'Паспорт_серия': 'p_seria',
    'Паспорт_номер': 'p_number',
    'Паспорт_дата': 'p_date',
    'Паспорт_Кем выдан': 'p_police',
    'Паспорт_Код подразделения': 'p_police_code',
    'Адрес_регистрации_Индекс': 'p_postalcode',
    'Адрес_регистрации_Регион': 'p_region',
    'Адрес_регистрации_Тип_региона': 'p_region_type',
    'Адрес_регистрации_Район': 'p_district',
    'Адрес_регистрации_Тип_района': 'p_district_type',
    'Адрес_регистрации_Город': 'p_place',
    'Адрес_регистрации_Тип_города': 'p_place_type',
    'Адрес_регистрации_Населенный_пункт': 'p_subplace',
    'Адрес_регистрации_Тип_населенного_пункта': 'p_subplace_type',
    'Адрес_регистрации_Улица': 'p_street',
    'Адрес_регистрации_Тип_улицы': 'p_street_type',
    'Адрес_регистрации_Дом': 'p_building',
    'Адрес_регистрации_Корпус': 'p_corpus',
    'Адрес_регистрации_Квартира': 'p_flat',
    'Адрес_проживания_Индекс': 'd_postalcode',
    'Адрес_проживания_Регион': 'd_region',
    'Адрес_проживания_Тип_региона': 'd_region_type',
    'Адрес_проживания_Район': 'd_district',
    'Адрес_проживания_Тип_района': 'd_district_type',
    'Адрес_проживания_Город': 'd_place',
    'Адрес_проживания_Тип_города': 'd_place_type',
    'Адрес_проживания_Населенный_пункт': 'd_subplace',
    'Адрес_проживания_Тип_населенного_пункта': 'd_subplace_type',
    'Адрес_проживания_Улица': 'd_street',
    'Адрес_проживания_Тип_улицы': 'd_street_type',
    'Адрес_проживания_Дом': 'd_building',
    'Адрес_проживания_Корпус': 'd_corpus',
    'Адрес_проживания_Квартира': 'd_flat',
    'Мобильный_телефон': 'phone_personal_mobile',
    'Телефон_родственников': 'phone_relative_mobile',
    'Телефон_домашний': 'phone_home'
}

def translate(name):
    transtable = (
        ## Большие буквы
        (u"E", u"Е"),
        (u"T", u"Т"),
        (u"O", u"О"),
        (u"P", u"Р"),
        (u"A", u"А"),
        (u"D", u"Д"),
        (u"H", u"Н"),
        (u"K", u"К"),
        (u"X", u"Х"),
        (u"C", u"С"),
        (u"B", u"В"),
        (u"M", u"М"),
        ## Маленькие буквы
        (u"e", u"е"),
        (u"t", u"т"),
        (u"o", u"о"),
        (u"p", u"р"),
        (u"a", u"а"),
        (u"d", u"д"),
        (u"n", u"п"),
        (u"h", u"н"),
        (u"k", u"к"),
        (u"x", u"х"),
        (u"c", u"с"),
        (u"b", u"б"),
        (u"y", u"у"),
        (u"m", u"т"),
    )
    #перебираем символы в таблице и заменяем
    for symb_in, symb_out in transtable:
        name = name.replace(symb_in, symb_out)
    #возвращаем переменную
    return name

if __name__ == '__main__':
    if len(sys.argv) == 2:
        if os.path.exists(sys.argv[1]):
            all_files = os.listdir(sys.argv[1])
            for all_file in all_files:
                if all_file.endswith('.xlsx'):
                    wb = openpyxl.load_workbook(filename=os.path.join(sys.argv[1], all_file), read_only=True)
                    ws = wb[wb.sheetnames[0]]
                    with open(os.path.join(sys.argv[1], all_file[:-5] + '.csv'), 'w') as csv_file:
                        csv_writer = csv.DictWriter(csv_file, fieldnames=FIELDS.values())
                        csv_writer.writeheader()
                        col_names = []
                        for i, row in enumerate(ws):
                            if i == 0:
                                for j, cell in enumerate(row):
                                    col_names.append(cell.value)
                            else:
                                csv_row = {}
                                for j, cell in enumerate(row):
                                    if FIELDS.get(col_names[j], None):
                                        if str(type(cell.value)).find('str') > -1:
                                            tek = translate(cell.value).replace('11.11.1111', '1111-11-11')
                                            #tek = tek.replace('переехал','')
                                            if col_names[j] in RUS:
                                                csv_row[FIELDS[col_names[j]]] = re.sub(r'[^А-Яа-яЁё]', '', tek)
                                            elif col_names[j] in RUS_MINUS:
                                                csv_row[FIELDS[col_names[j]]] = re.sub(r'[^А-Яа-яЁё\-]', '', tek)
                                            elif col_names[j] in RUS_SP_MINUS:
                                                csv_row[FIELDS[col_names[j]]] = re.sub(r'[^А-Яа-яЁё\-\s]', '', tek)
                                            elif col_names[j] in CIFR:
                                                csv_row[FIELDS[col_names[j]]] = re.sub(r'[^0-9]', '', tek)
                                            elif col_names[j] in CIFR_MINUS:
                                                csv_row[FIELDS[col_names[j]]] = re.sub(r'[^0-9\-]', '', tek)
                                            else:
                                                csv_row[FIELDS[col_names[j]]] = re.sub(r'[^А-Яа-яЁё0-9\s\.\/\\\-]', '', tek)
                                            if csv_row[FIELDS[col_names[j]]] != tek:
                                                print('файл', all_file, 'строка:', i, 'столбец:', col_names[j],
                                                      csv_row[FIELDS[col_names[j]]], ' => ', tek)
                                csv_writer.writerow(csv_row)
        else:
            print('Не указана папка с файлами для преобразования')
    else:
        print('Не указана папка с файлами для преобразования')



