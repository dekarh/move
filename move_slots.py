# -*- coding: utf-8 -*-

from subprocess import Popen, PIPE
import os
import sys
import re
import string
import bz2
from string import digits
from random import random
from dateutil.parser import parse
from collections import OrderedDict

from datetime import datetime, timedelta, time
import time
import pytz
utc=pytz.UTC

import openpyxl
from openpyxl import Workbook

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QDate, QDateTime, QSize, Qt, QByteArray, QTimer, QUrl, QThread
from PyQt5.QtGui import QPixmap, QIcon
from PyQt5.QtWidgets import QTableWidgetItem, QMessageBox, QMainWindow, QWidget, QFrame, QFileDialog, QComboBox

from mysql.connector import MySQLConnection, Error

from move_win import Ui_Form

# import NormalizeFields as norm
from lib import read_config, l, s, fine_phone, format_phone

MANIPULATE_LABELS = ["-------------------------"
                     , "ФИО из поля"
                     , "-------------------------"
                     , "ФИО при рождении из поля"
                     , "-------------------------"
                     , "Регистрация -> Регион"
                     , "Регистрация -> Район"
                     , "Регистрация -> Город"
                     , "Регистрация -> Населенный_пункт"
                     , "Регистрация -> Улица"
                     , "-------------------------"
                     , "Проживание -> Регион"
                     , "Проживание -> Район"
                     , "Проживание -> Город"
                     , "Проживание -> Населенный_пункт"
                     , "Проживание -> Улица"
                     , "-------------------------"
                     , "Адрес регистрации из_поля"
                     , "-------------------------"
                     , "Адрес проживания из поля"
                     , "-------------------------"
                     , "Регион регистрации из номера"
                     , "-------------------------"
                     , "Регион проживания из номера"
                     , "-------------------------"
                     , "Cерия и Номер паспорта из поля"
#                     , "Пол_получить_из_ФИО"
#                     , "Пол_подставить_свои_значения"
                     ]

SNILS_LABEL = ["СНИЛС"]
FIO_LABELS = ["ФИО.Фамилия", "ФИО.Имя", "ФИО.Отчество"]
FIO_BIRTH_LABELS = ["ФИО_при_рождении.Фамилия", "ФИО_при_рождении.Имя", "ФИО_при_рождении.Отчество"]
FIO_SNILS_LABELS = ["Фамилия_по_СНИЛС", "Имя_по_СНИЛС", "Отчество_по_СНИЛС"]
GENDER_LABEL = ["Пол"]
DATE_BIRTH_LABEL = ["Дата_рождения"]
PLACE_BIRTH_LABELS = ["Место_рождения.Страна", "Место_рождения.Область", "Место_рождения.Район",
                      "Место_рождения.Город"]
PASSPORT_DATA_LABELS = ["Данные_паспорта.Серия", "Данные_паспорта.Номер", "Данные_паспорта.Дата_выдачи",
                        "Данные_паспорта.Кем_выдан", "Данные_паспорта.Код_подразделения"]
ADRESS_REG_LABELS = ["Адрес_регистрации.Индекс",
                     "Адрес_регистрации.Регион", "Адрес_регистрации.Тип_региона",
                     "Адрес_регистрации.Район", "Адрес_регистрации.Тип_района",
                     "Адрес_регистрации.Город", "Адрес_регистрации.Тип_города",
                     "Адрес_регистрации.Населенный_пункт", "Адрес_регистрации.Тип_населенного_пункта",
                     "Адрес_регистрации.Улица", "Адрес_регистрации.Тип_улицы",
                     "Адрес_регистрации.Дом",
                     "Адрес_регистрации.Корпус",
                     "Адрес_регистрации.Квартира"]

ADRESS_LIVE_LABELS = ["Адрес_проживания.Индекс",
                      "Адрес_проживания.Регион", "Адрес_проживания.Тип_региона",
                      "Адрес_проживания.Район", "Адрес_проживания.Тип_района",
                      "Адрес_проживания.Город", "Адрес_проживания.Тип_города",
                      "Адрес_проживания.Населенный_пункт", "Адрес_проживания.Тип_населенного_пункта",
                      "Адрес_проживания.Улица", "Адрес_проживания.Тип_улицы",
                      "Адрес_проживания.Дом",
                      "Адрес_проживания.Корпус",
                      "Адрес_проживания.Квартира"]

PHONES_LABELS = ["Телефон.Мобильный", "Телефон.Родственников", "Телефон.Домашний"]

TECH_LABELS = ["Агент_Ид", "Подписант_Ид", "Пред_Страховщик_Ид"]

#------------------------Отключил MANIPULATE_LABELS------------------------------------------------------------
# FIELDS_IN_RESULT_TABLE_COMPLETE = [SNILS_LABEL, FIO_LABELS, FIO_BIRTH_LABELS, GENDER_LABEL, DATE_BIRTH_LABEL,
#                                   PLACE_BIRTH_LABELS, PASSPORT_DATA_LABELS, ADRESS_REG_LABELS, ADRESS_LIVE_LABELS,
#                                   PHONES_LABELS, MANIPULATE_LABELS]
#------------------------Отключил MANIPULATE_LABELS------------------------------------------------------------

FIELDS_IN_RESULT_TABLE_COMPLETE = [SNILS_LABEL, FIO_LABELS, FIO_BIRTH_LABELS, FIO_SNILS_LABELS, GENDER_LABEL,
                                   DATE_BIRTH_LABEL, PLACE_BIRTH_LABELS, PASSPORT_DATA_LABELS, ADRESS_REG_LABELS,
                                   ADRESS_LIVE_LABELS, PHONES_LABELS, TECH_LABELS, MANIPULATE_LABELS]

HEAD_RESULT_EXCEL_FILE = ['СНИЛС',
                          'Фамилия', 'Имя', 'Отчество',
                          'Фамилия_при_рождении', 'Имя_при_рождении', 'Отчество_при_рождении',
                          'Фамилия_по_СНИЛС', 'Имя_по_СНИЛС', 'Отчество_по_СНИЛС',
                          'Пол(0_мужской,1_женский)',
                          'Дата_рождения',
                          'Страна_рождения', 'Область_рождения', 'Район_рождения', 'Город_рождения',
                          'Паспорт_серия', 'Паспорт_номер', 'Паспорт_дата', 'Паспорт_Кем выдан',
                          'Паспорт_Код подразделения',

                          'Адрес_регистрации_Индекс',
                          'Адрес_регистрации_Регион', 'Адрес_регистрации_Тип_региона',
                          'Адрес_регистрации_Район', 'Адрес_регистрации_Тип_района',
                          'Адрес_регистрации_Город', 'Адрес_регистрации_Тип_города',
                          'Адрес_регистрации_Населенный_пункт', 'Адрес_регистрации_Тип_населенного_пункта',
                          'Адрес_регистрации_Улица',
                          'Адрес_регистрации_Тип_улицы',
                          'Адрес_регистрации_Дом',
                          'Адрес_регистрации_Корпус',
                          'Адрес_регистрации_Квартира',

                          'Адрес_проживания_Индекс',
                          'Адрес_проживания_Регион', 'Адрес_проживания_Тип_региона',
                          'Адрес_проживания_Район', 'Адрес_проживания_Тип_района',
                          'Адрес_проживания_Город', 'Адрес_проживания_Тип_города',
                          'Адрес_проживания_Населенный_пункт', 'Адрес_проживания_Тип_населенного_пункта',
                          'Адрес_проживания_Улица', 'Адрес_проживания_Тип_улицы',
                          'Адрес_проживания_Дом',
                          'Адрес_проживания_Корпус',
                          'Адрес_проживания_Квартира',

                          'Мобильный_телефон', 'Телефон_родственников', 'Телефон_домашний',
                          'Агент_Ид', 'Подписант_Ид', 'Пред_Страховщик_Ид'
                          ]

LEN_SNILS = 11
LEN_PASSPORT_NOMER = 6
LEN_INDEX_NOMER = 6
LEN_PASSPORT_COD = 6
EAST_GENDER = ['кызы', 'оглы']

########################################################################################################################
# НУЛЕВЫЕ ЗНАЧЕНИЯ
NULL_VALUE = '\\N'  # НУЛЕВОЕ ЗНАЧЕНИЕ В ФАЙЛЕ
NEW_NULL_VALUE = ''  # НОВОЕ НУЛЕВОЕ ЗНАЧЕНИЕ

NEW_NULL_VALUE_FOR_DATE = '11.11.1111'
NEW_NULL_VALUE_FOR_SERIYA_PASSPORTA = '1111'
NEW_NULL_VALUE_FOR_NOMER_PASSPORTA = '111111'
NEW_NULL_VALUE_FOR_COD_PASSPORTA = '111-111'
NEW_NULL_VALUE_FOR_GENDER = '0'
NEW_NULL_VALUE_FOR_INDEX = '111111'
NEW_NULL_VALUE_FOR_ALL_TEXT = 'заполнить'
NEW_NULL_VALUE_FOR_HOME = 'заполнить'
########################################################################################################################
# ЗНАЧЕНИЕ ПРИ ОШИБКЕ
ERROR_VALUE = 'ERROR'
########################################################################################################################
# СОКРАЩЕНИЯ ТИПОВ В АДРЕСЕ
SPLIT_FIELD = ','
#SPLIT_FIELD = '_x0003_'  # Разделитель для адреса в одной строке (бывает '_x0003_')

#ORDER_FIELD = [13, 0, 8, 1, 9, 2, 10, 3, 11, 4, 12, 5, 6, 7] # См def FullAddress(get_values():
# ['Индекс', 'Регион', 'Тип_региона', 'Район', 'Тип_района', 'Город', 'Тип_города',
#  'Населенный_пункт', 'Тип_населенного_пункта', 'Улица', 'Тип_улицы', 'Дом', 'Корпус', 'Квартира']

REG_TYPES = ['обл', 'о', 'область', 'респ', 'республика', 'край', 'кр', 'ар', 'ао', 'авт']

DISTRICT_TYPES = ['р-н', 'р', 'район']

CITY_TYPES = ['г', 'гор', 'город']

NP_TYPES = ['пгт', 'пос', 'поселение', 'поселок', 'посёлок', 'п', 'рп', 'кп', 'к', 'пс', 'сс', 'смн', 'вл', 'дп',
            'нп', 'пст', 'ж/д_ст', 'с', 'м', 'д', 'дер', 'сл', 'ст', 'ст-ца', 'х', 'рзд', 'у', 'клх', 'свх', 'зим', 'мкр']

STREET_TYPES = ['аллея', 'а', 'бульвар', 'б-р', 'в/ч', 'городок', 'гск', 'кв-л', 'линия', 'наб', 'пер', 'переезд', 'пл',
                'пр-кт', 'проезд', 'тер', 'туп', 'ул', 'ш', ]

HOUSE_CUT_NAME = ['дом', 'д']
CORPUS_CUT_NAME = ['корп', 'корпус']
APARTMENT_CUT_NAME = ['кв']
########################################################################################################################
# ЗНАЧЕНИЕ В ПОЛЕ "ПОЛ" ИЗМЕНЯЕМ В ПРОЦЕССЕ
female_gender_value = 'Ж'
male_gender_value = 'М'
gender_length = 1
########################################################################################################################
# ЗАПОЛНЕНИЕ Агент_Ид, Подписант_Ид, Пред_Страховщик_Ид
#AGENT_ID = '10061'
#AGENT_ID = '9954'
#AGENT_ID = '9986'
#PODPISANT_ID = '208'
PREDSTRAH_ID = '1'
########################################################################################################################
# ИМЕНА ДЛЯ КЛЮЧЕЙ СЛОВАРЕЙ И ДЛЯ ПОРЯДКА ВЫВОД СЛОВАРЯ

FULL_ADRESS_LABELS = ['Индекс', 'Регион', 'Тип_региона', 'Район', 'Тип_района', 'Город', 'Тип_города',
                      'Населенный_пункт', 'Тип_населенного_пункта', 'Улица', 'Тип_улицы', 'Дом', 'Корпус', 'Квартира']

PASSPORT_LABELS = ['Серия', 'Номер', 'Дата_выдачи', 'Кем_выдан', 'Код_подразделения']

FIO_KEY_LABELS = ['Фамилия', 'Имя', 'Отчество']

BIRTH_PLACE_LABELS = ['Страна', 'Область', 'Район', 'Город']
########################################################################################################################
REGIONS = [
    "", "Адыгея респ.", "Башкортостан респ.", "Бурятия респ.", "Алтай респ.", "Дагестан респ.", "Ингушетия респ.",
    "Кабардино-Балкарская респ.", "Калмыкия респ.", "Карачаево-Черкесская респ.", "Карелия респ.", "Коми респ.",
    "Марий_Эл респ.", "Мордовия респ.", "Саха/Якутия/ респ.", "Северная_Осетия-Алания респ.", "Татарстан респ.",
    "Тыва респ.", "Удмуртская респ.", "Хакасия респ.", "Чеченская респ.", "Чувашская респ.", "Алтайский край",
    "Краснодарский край", "Красноярский край", "Приморский край", "Ставропольский край", "Хабаровский край",
    "Амурская обл.", "Архангельская обл.", "Астраханская обл.", "Белгородская обл.", "Брянская обл.",
    "Владимирская обл.", "Волгоградская обл.", "Вологодская обл.", "Воронежская обл.", "Ивановская обл.",
    "Иркутская обл.", "Калининградская обл.", "Калужская обл.", "Камчатский край", "Кемеровская обл.", "Кировская обл.",
    "Костромская обл.", "Курганская обл.", "Курская обл.", "Ленинградская обл.", "Липецкая обл.", "Магаданская обл.",
    "Московская обл.", "Мурманская обл.", "Нижегородская обл.", "Новгородская обл.", "Новосибирская обл.",
    "Омская обл.", "Оренбургская обл.", "Орловская обл.", "Пензенская обл.", "Пермский край", "Псковская обл.",
    "Ростовская обл.", "Рязанская обл.", "Самарская обл.", "Саратовская обл.", "Сахалинская обл.", "Свердловская обл.",
    "Смоленская обл.", "Тамбовская обл.", "Тверская обл.", "Томская обл.", "Тульская обл.", "Тюменская обл.",
    "Ульяновская обл.", "Челябинская обл.", "Забайкальский край", "Ярославская обл.", "Москва г.", "Санкт-Петербург г.",
    "Еврейская авт.обл.", "Агинский_Бурятский авт.округ", "Коми-Пермяцкий авт.округ", "Корякский авт.округ",
    "Ненецкий авт.округ", "Таймырский_(Долгано-Ненецкий) авт.округ", "Усть-Ордынский_Бурятский авт.округ",
    "Ханты-Мансийский/Югра/ авт.округ", "Чукотский авт.округ", "Эвенкийский авт.округ", "Ямало-Ненецкий авт.округ",
    "","Крым респ.", "Севастополь г.","","","","","","","Байконур г." ]
########################################################################################################################
# True - заменять СНИЛС на некорректный, неиспользованный в Сатурн

GENERATE_SNILS = False

########################################################################################################################

IN_IDS = ['ID','ИД_КЛИЕНТА','CLIENT_ID']
IN_SNILS = ['СНИЛС', 'СТРАХОВОЙ_НОМЕР', 'СТРАХОВОЙ НОМЕР', 'NUMBER','СТРАХОВОЙНОМЕР']
IN_NAMES = ['ID', 'СНИЛС', 'СТРАХОВОЙ_НОМЕР', 'СТРАХОВОЙ НОМЕР', 'NUMBER', 'ФАМИЛИЯ', 'ИМЯ', 'ОТЧЕСТВО', 'ФИО']

DIR4MOVE = '/home/da3/Move/'
DIR4IMPORT = '/home/da3/CheckLoad/'
DIR4CFGIMPORT = '/home/da3/CheckLoad/cfg/'
DIR4PCHECK = '/home/da3/PasportChecks/'

class MainWindowSlots(Ui_Form):   # Определяем функции, которые будем вызывать в слотах

    def setupUi(self, form):
        Ui_Form.setupUi(self,form)
        self.passports = {}
        self.tableWidget.setColumnCount(2)
        self.tableWidget.setHorizontalHeaderLabels(('Результат', 'Исходник'))
        for j in range(self.tableWidget.columnCount()):
            self.tableWidget.setColumnWidth(j, 220)
        self.tableWidget.setRowCount(0)
        self.dbconfig = read_config(filename='move.ini', section='mysql')
        self.dbconfig_pasp = read_config(filename='move.ini', section='pasport')
        self.MoveImportPasport = 1
        self.signer_ids = []
        self.signer_names = {}
        self.fond_ids = []
        self.fond_names = {}
        self.agent_ids = []
        self.agent_names = {}
        self.clients_ids = []
        self.fond_touched = False
        self.fonds_str = '1'
        self.agent_touched = False
        self.signer_touched = False
        self.cfg_file_touched = False
        self.cfg_file_loaded = False
        self.cfg_file_names = {}
        self.file_touched = False
        self.file_loaded = False
        self.table_loaded = False
        self.file_names = {}
        self.file_name = ''
        self.tab_names = {}
        self.table = []
        self.twParsingResult.hide()
        self.cmbGenderType.addItems(['М или Ж', 'Муж. или Жен.', 'Мужской или Женский'])
        self.refresh()
        return

    def refresh(self):
        if self.fond_touched:                                   # Запомнили позиции combobox'ов
            fond_id = self.fond_ids[self.cmbFond.currentIndex()]
        if self.agent_touched:
            agent_id = self.agent_ids[self.cmbAgent.currentIndex()]
        if self.signer_touched:
            signer_id = self.signer_ids[self.cmbSigner.currentIndex()]
        if self.cfg_file_loaded:
            cfg_file_name = self.cmbCfgFile.currentText()
        if self.file_loaded:
            file_name = self.cmbFile.currentText()
            tab_name = self.cmbTab.currentText()
        self.selectAction()

        dbconn = MySQLConnection(**self.dbconfig)
        cursor = dbconn.cursor()
        if self.leAgent.text().strip():
            sql = "SELECT CONCAT_WS(' ', code, '-', user_surname, user_name, user_lastname, '-', position_id), code " \
                  "FROM offices_staff WHERE user_fired = 0 AND " \
                  "CONCAT_WS(' ', code, '-', user_surname, user_name, user_lastname, user_lastname) LIKE %s"
            cursor.execute(sql, ('%' + self.leAgent.text() + '%',))
        else:
            sql = "SELECT CONCAT_WS(' ', code, '-', user_surname, user_name, user_lastname, '-', position_id), code " \
                  "FROM offices_staff WHERE user_fired = 0"
            cursor.execute(sql)
        rows = cursor.fetchall()
        agents = []
        self.agent_names = {}
        self.agent_ids = []
        for i, row in enumerate(rows):
            agents.append(row[0])
            self.agent_names[row[1]] = row[0]
            self.agent_ids.append(row[1])
        self.cmbAgent.clear()
        self.cmbAgent.addItems(agents)
        cursor = dbconn.cursor()
        if self.leFond.text().strip():
            sql = "SELECT CONCAT_WS(' ', id, '-', name), id FROM subdomains " \
                  "WHERE CONCAT_WS(' ', id, '-', name) LIKE %s AND id IN (2,6,8,11,12,13,14)"
            cursor.execute(sql, ('%' + self.leFond.text() + '%',))
        else:
            sql = "SELECT CONCAT_WS(' ', id, '-', name), id FROM subdomains WHERE id IN (2,6,8,11,12,13,14)"
            cursor.execute(sql)
        rows = cursor.fetchall()
        fonds = []
        self.fond_ids = []
        self.fond_names = {}
        self.fonds_str = '1'
        for i, row in enumerate(rows):
            fonds.append(row[0])
            self.fond_names[row[1]] = row[0]
            self.fond_ids.append(row[1])
            if i == 0:
                self.fonds_str = str(row[1])
                continue
            self.fonds_str += ','+ str(row[1])
        self.cmbFond.clear()
        self.cmbFond.addItems(fonds)

        cursor = dbconn.cursor()
        if self.leSigner.text().strip():
            if self.fond_touched:
                sql = "SELECT CONCAT_WS(' ', id, '-', signer_surname, signer_name, signer_lastname), id " \
                      "FROM signers " \
                      "WHERE CONCAT_WS(' ', id, '-', signer_surname, signer_name, signer_lastname) LIKE %s " \
                      "AND subdomain_id = %s"
                cursor.execute(sql, (self.fond_ids[self.cmbFond.currentIndex()],))
            else:
                sql = "SELECT CONCAT_WS(' ', id, '-', signer_surname, signer_name, signer_lastname), id " \
                      "FROM signers " \
                      "WHERE CONCAT_WS(' ', id, '-', signer_surname, signer_name, signer_lastname) LIKE %s " \
                      "AND subdomain_id IN (2,6,8,11,12,13,14) AND subdomain_id IN (" + self.fonds_str + ")"
                cursor.execute(sql, ('%' + self.leSigner.text() + '%',))
        else:
            if self.fond_touched:
                sql = "SELECT CONCAT_WS(' ', id, '-', signer_surname, signer_name, signer_lastname), id " \
                      "FROM signers WHERE subdomain_id = %s"
                cursor.execute(sql, (self.fond_ids[self.cmbFond.currentIndex()],))
            else:
                sql = "SELECT CONCAT_WS(' ', id, '-', signer_surname, signer_name, signer_lastname), id " \
                      "FROM signers WHERE subdomain_id IN (" + self.fonds_str + ") " \
                      "AND subdomain_id IN (2,6,8,11,12,13,14)"
                cursor.execute(sql)
        rows = cursor.fetchall()
        self.signer_ids = []
        self.signer_names = {}
        signers = []
        for i, row in enumerate(rows):
            self.signer_names[row[1]] = row[0]
            self.signer_ids.append(row[1])
            signers.append(row[0])
        self.cmbSigner.clear()
        self.cmbSigner.addItems(signers)

        rows = os.listdir(DIR4CFGIMPORT)                        # список конфигов
        cfg_files = []
        for row in rows:
            if row.find('.xlsx') > -1:
                cfg_files.append(row)
        self.cfg_file_names = {}
        for i, cfg_file in enumerate(cfg_files):
            self.cfg_file_names[cfg_file] = i
        self.cmbCfgFile.clear()
        self.cmbCfgFile.addItems(cfg_files)

        try:                                                    # список файлов загрузки
            files = os.listdir(path=self.leDir.text())
            for i, file in enumerate(files):
                self.file_names[file] = i
            self.cmbFile.clear()
            self.cmbFile.addItems(files)
            try:
                if not self.file_loaded:
                    self.twAllExcels.setColumnCount(0)
                    self.twAllExcels.setRowCount(0)
                    self.file_name = ''
                else:
                    self.cmbFile.setCurrentIndex(self.file_names[file_name])
            except ValueError:
                self.file_loaded = False
                self.twAllExcels.setColumnCount(0)
                self.twAllExcels.setRowCount(0)
                self.file_name = ''
            else:
                if self.file_loaded:
                    if self.cmbFile.currentText()[len(self.cmbFile.currentText()) - 5:] == '.xlsx':
                        self.wb = openpyxl.load_workbook(filename=self.leDir.text() + self.cmbFile.currentText(),
                                                         read_only=True)
                        tabs = self.wb.sheetnames
                        for i, tab in enumerate(tabs):
                            self.tab_names[tab] = i
                        self.cmbTab.clear()
                        self.cmbTab.addItems(tabs)
        except OSError:
            self.errMessage('Нет такого файла')

        try:                                            # Восстанавиваем позиции combobox'ов
            if self.fond_touched:
                self.cmbFond.setCurrentIndex(self.fond_ids.index(fond_id))
        except ValueError:
            self.fond_touched = False
        try:
            if self.agent_touched:
                self.cmbAgent.setCurrentIndex(self.agent_ids.index(agent_id))
        except ValueError:
            self.agent_touched = False
        try:
            if self.signer_touched:
                self.cmbSigner.setCurrentIndex(self.signer_ids.index(signer_id))
        except ValueError:
            self.signer_touched = False
        try:
            if self.file_loaded:
                self.cmbTab.setCurrentIndex(self.tab_names[tab_name])
        except ValueError:
            self.file_loaded = False
            self.twAllExcels.setColumnCount(0)
            self.twAllExcels.setRowCount(0)
            self.file_name = ''
        try:
            if self.cfg_file_loaded:
                self.cmbCfgFile.setCurrentIndex(self.cfg_file_names[cfg_file_name])
        except ValueError:
            self.cfg_file_loaded = False
        return

    def click_pbRefresh(self):
        self.refresh()
        return

    def click_pbMove(self):
        if not self.file_touched:                               # Проверяем достаточность данных
            self.frFile.setStyleSheet("QFrame{background-image: url(./x.png)}")
            return
        if self.leAgent.isEnabled() and not self.agent_touched:
            self.frAgent.setStyleSheet("QFrame{background-image: url(./x.png)}")
            return
        if self.leFond.isEnabled() and not self.fond_touched:
            self.frFond.setStyleSheet("QFrame{background-image: url(./x.png)}")
            return
        if self.leSigner.isEnabled() and not self.signer_touched:
            self.frSigner.setStyleSheet("QFrame{background-image: url(./x.png)}")
            return
                                                                # Создаем файл с исходными данными и логом
        wb_log = openpyxl.Workbook(write_only=True)

        ws_log = wb_log.create_sheet('Лог')
        ws_log.append([datetime.now().strftime("%H:%M:%S"), ' Начинаем'])

        log_name = DIR4MOVE + datetime.now().strftime('%Y-%m-%d_%H-%M')
        if self.fond_touched:
            log_name += 'ф' + str(self.fond_ids[self.cmbFond.currentIndex()])
        if self.agent_touched:
            log_name += 'а' + str(self.agent_ids[self.cmbAgent.currentIndex()])
        log_name += '.xlsx'

        all_clients_ids = "'" + self.clients_ids[0] + "'"       # Проверка на дубли clients
        for i, client_id in enumerate(self.clients_ids):
            if i == 0:
                continue
            all_clients_ids += ",'" + client_id + "'"
        sql = "SELECT cl.client_id FROM clients AS cl WHERE cl.client_id IN (" + all_clients_ids + \
              ") GROUP BY cl.client_id HAVING COUNT(cl.client_id) > 1 ORDER BY cl.client_id DESC"
        dbconn = MySQLConnection(**self.dbconfig)
        cursor = dbconn.cursor()
        cursor.execute(sql)
        rows = cursor.fetchall()
        exit_because_doubles = False
        if len(rows) > 0:
            exit_because_doubles = True
            ws_log.append([datetime.now().strftime("%H:%M:%S"), ' Дубли в clients'])
            ws_clients = wb_log.create_sheet('Дубли в clients')
            for row in rows:
                ws_clients.append(row[0])
        else:
            ws_log.append([datetime.now().strftime("%H:%M:%S"), ' В clients нет дублей'])
                                                                # Проверка на дубли contracts
        sql = "SELECT co.client_id FROM contracts AS co WHERE co.client_id IN (" + all_clients_ids + \
              ") GROUP BY co.client_id HAVING COUNT(co.client_id) > 1 ORDER BY co.client_id DESC"
        dbconn = MySQLConnection(**self.dbconfig)
        cursor = dbconn.cursor()
        cursor.execute(sql)
        rows = cursor.fetchall()
        if len(rows) > 0:
            exit_because_doubles = True
            ws_log.append([datetime.now().strftime("%H:%M:%S"), ' Дубли в contracts'])
            ws_contracts = wb_log.create_sheet('Дубли в contracts')
            for row in rows:
                ws_contracts.append(row[0])
        else:
            ws_log.append([datetime.now().strftime("%H:%M:%S"), ' В contracts нет дублей'])

                # Проверка на дубли исходной таблицы
        doubles_in_input = list(set([x for x in self.clients_ids if self.clients_ids.count(x) > 1]))
        if len(doubles_in_input) > 0:
            ws_log.append([datetime.now().strftime("%H:%M:%S"), ' Дубли в исходной таблице'])
            ws_input_doubles = wb_log.create_sheet('Дубли в исходной таблице')
            for row in doubles_in_input:
                ws_input_doubles.append(row)
        else:
            ws_log.append([datetime.now().strftime("%H:%M:%S"), ' В исходной таблице нет дублей'])

        if exit_because_doubles:                          # Если дубли в clients или contracts - ничего не переносим
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'Аварийное завершение - дублирование записей'])
            return

        ws_log.append([datetime.now().strftime("%H:%M:%S"), 'Дублируем исходную excel таблицу в этот файл'])
        ws_input = wb_log.create_sheet('Исходная таблица')
        for table_row in self.table:
            row = []
            for cell in table_row:
                row.append(cell)
            ws_input.append(row)

        ws_log.append([datetime.now().strftime("%H:%M:%S"), 'Бэкап исходного состояния БД создан'])
        ws_backup = wb_log.create_sheet('бэкап БД')
        dbconn = MySQLConnection(**self.dbconfig)
        cursor = dbconn.cursor()
        sql = "SELECT cl.*, co.* FROM clients AS cl LEFT JOIN contracts AS co " \
              "ON (cl.client_id = co.client_id) WHERE cl.client_id IN (" + all_clients_ids + ")"
        cursor.execute(sql)
        dbrows = cursor.fetchall()
        ws_backup.append(cursor.column_names)
        for dbrow in dbrows:
            row = []
            for dbcell in dbrow:
                row.append(dbcell)
            ws_backup.append(row)

        ws_log.append([datetime.now().strftime("%H:%M:%S"), ' Состояние программы:'])
        ws_log.append([datetime.now().strftime("%H:%M:%S"), 'файл ', self.file_name])
        if self.leFond.isEnabled():
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'фонд', self.cmbFond.currentText()])
        else:
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'фонд', 'не выбран'])
        if self.leAgent.isEnabled():
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'агент', self.cmbAgent.currentText()])
        else:
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'агент', 'не выбран'])
        if self.leSigner.isEnabled():
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'подписант', self.cmbSigner.currentText()])
        else:
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'подписант', 'не выбран'])
        if self.chbClientOnly.isChecked():
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'Перенести только клиента', 'выбрано'])
        else:
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'Перенести только клиента', 'не выбрано'])
        if self.chbSocium.isChecked():
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'Сбросить номер Социума', 'выбрано'])
        else:
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'Сбросить номер Социума', 'не выбрано'])
        if self.chbSuff.isChecked():
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'Суффикс', self.leSuff.text()])
        else:
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'Суффикс', 'не выбрано'])
        if self.chbOurStat.isChecked():
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'Сбросить внутренние статусы', 'выбрано'])
        else:
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'Сбросить внутренние статусы', 'не выбрано'])
        if self.chbFondStat.isChecked():
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'Сбросить статусы Фонда', 'выбрано'])
        else:
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'Сбросить статусы Фонда', 'не выбрано'])
        if self.chbArhivON.isChecked():
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'Поставить флаг "Архивный"', 'выбрано'])
        else:
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'Поставить флаг "Архивный"', 'не выбрано'])
        if self.chbArhivOFF.isChecked():
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'Убрать флаг "Архивный"', 'выбрано'])
        else:
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'Убрать флаг "Архивный"', 'не выбрано'])

        ws_log.append([datetime.now().strftime("%H:%M:%S"), ' Формируем запросы:'])
        sql_cl = 'UPDATE clients AS cl SET'
        sql_co = 'UPDATE contracts AS co SET'
        if self.leAgent.isEnabled():
            sql_cl += ' cl.inserted_user_code = %s'
            sql_co += ' co.inserted_code = %s, co.agent_code = %s'
        if self.leFond.isEnabled():
            if sql_cl[len(sql_cl) - 2:] == '%s':
                sql_cl += ','
            sql_cl += ' cl.subdomain_id = %s'
        if self.chbArhivON.isChecked() or self.chbArhivOFF.isChecked():
            if sql_cl[len(sql_cl) - 2:] == '%s':
                sql_cl += ','
            sql_cl += ' cl.archived = %s'
        if self.leSigner.isEnabled():
            if sql_co[len(sql_co) - 2:] == '%s':
                sql_co += ','
            sql_co += ' co.signer_id = %s'
        if self.chbSocium.isChecked():
            if sql_co[len(sql_co) - 2:] == '%s':
                sql_co += ','
            sql_co += ' co.socium_contract_number = %s'
        if self.chbFondStat.isChecked():
            if sql_co[len(sql_co) - 2:] == '%s':
                sql_co += ','
            sql_co += ' co.external_status_code = %s, co.external_status_secure_code = %s,' \
                      ' co.external_status_callcenter_code = %s'
        if self.chbOurStat.isChecked():
            if sql_co[len(sql_co) - 2:] == '%s':
                sql_co += ','
            sql_co += ' co.status_code = %s, co.status_secure_code = %s, co.status_callcenter_code = %s'
        if self.chbSuff.isChecked():
            if sql_co[len(sql_co) - 2:] == '%s':
                sql_co += ','
            sql_co += ' co.partner_remote_id = %s'
        if sql_cl[len(sql_cl) - 3:] == 'SET':
            self.leSQLcl.setText('')
        else:
            self.leSQLcl.setText(sql_cl + ' WHERE cl.client_id = %s')
        if sql_co[len(sql_co) - 3:] == 'SET' or self.chbClientOnly.isChecked():
            self.leSQLco.setText('')
        else:
            self.leSQLco.setText(sql_co + ' WHERE co.client_id = %s')

        tuples_clients = []                                     # Формируем переменные для запросов
        tuples_contracts = []
        for i, client_id in enumerate(self.clients_ids):
            tuple_client = tuple()
            tuple_contract = tuple()
            if self.leAgent.isEnabled():
                tuple_client += (self.agent_ids[self.cmbAgent.currentIndex()],)
                tuple_contract += (self.agent_ids[self.cmbAgent.currentIndex()],
                                   self.agent_ids[self.cmbAgent.currentIndex()])
            if self.leFond.isEnabled():
                tuple_client += (self.fond_ids[self.cmbFond.currentIndex()],)
            if self.chbArhivON.isChecked():
                tuple_client += (1,)
            if self.chbArhivOFF.isChecked():
                tuple_client += (0,)
            if self.leSigner.isEnabled():
                tuple_contract += (self.signer_ids[self.cmbSigner.currentIndex()],)
            if self.chbSocium.isChecked():
                tuple_contract += (None,)
            if self.chbFondStat.isChecked():
                tuple_contract += (0,0,0)
            if self.chbOurStat.isChecked():
                tuple_contract += (0,0,0)
            if self.chbSuff.isChecked():
                tuple_contract += (self.leSuff.text(),)
            tuple_contract += (client_id,)
            tuple_client += (client_id,)
            tuples_clients.append(tuple_client)
            tuples_contracts.append(tuple_contract)
        if self.leSQLcl.text():
            ws_log.append([datetime.now().strftime("%H:%M:%S"), self.leSQLcl.text()])
            ws_log.append([datetime.now().strftime("%H:%M:%S")] + list(tuples_clients[0]))
            dbconn = MySQLConnection(**self.dbconfig)
            cursor = dbconn.cursor()
            cursor.executemany(self.leSQLcl.text(),tuples_clients)
            dbconn.commit()
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'update отработал'])
        else:
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'clients - нет запроса'])
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'clients - нет данных'])
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'clients - запрос не исполнен'])
        if self.leSQLco.text():
            ws_log.append([datetime.now().strftime("%H:%M:%S"), self.leSQLco.text()])
            ws_log.append([datetime.now().strftime("%H:%M:%S")] + list(tuples_contracts[0]))
            dbconn = MySQLConnection(**self.dbconfig)
            cursor = dbconn.cursor()
            cursor.executemany(self.leSQLco.text(),tuples_contracts)
            dbconn.commit()
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'update отработал'])
        else:
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'contracts - нет запроса'])
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'contracts - нет данных'])
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'contracts - запрос не исполнен'])

        wb_log.save(log_name)
        q=0

    def change_leAgent(self):
        if self.agent_touched:
            agent_id = self.agent_ids[self.cmbAgent.currentIndex()]
        dbconn = MySQLConnection(**self.dbconfig)
        cursor = dbconn.cursor()
        if self.leAgent.text().strip():
            sql = "SELECT CONCAT_WS(' ', code, '-', user_surname, user_name, user_lastname, '-', position_id), code " \
                  "FROM offices_staff WHERE user_fired = 0 AND " \
                  "CONCAT_WS(' ', code, '-', user_surname, user_name, user_lastname, user_lastname) LIKE %s"
            cursor.execute(sql, ('%' + self.leAgent.text() + '%',))
        else:
            sql = "SELECT CONCAT_WS(' ', code, '-', user_surname, user_name, user_lastname, '-', position_id), code " \
                  "FROM offices_staff WHERE user_fired = 0"
            cursor.execute(sql)
        rows = cursor.fetchall()
        agents = []
        self.agent_names = {}
        self.agent_ids = []
        for i, row in enumerate(rows):
            agents.append(row[0])
            self.agent_names[row[1]] = row[0]
            self.agent_ids.append(row[1])
        self.cmbAgent.clear()
        self.cmbAgent.addItems(agents)
        try:
            if self.agent_touched:
                self.cmbAgent.setCurrentIndex(self.agent_ids.index(agent_id))
        except ValueError:
            self.agent_touched = False

    def click_pbAgent(self):
        self.frAgent.setStyleSheet("QFrame{background-image: }")
        self.leAgent.setEnabled(not self.leAgent.isEnabled())
        self.cmbAgent.setEnabled(not self.cmbAgent.isEnabled())
        return

    def click_pbFond(self):
        self.frFond.setStyleSheet("QFrame{background-image: }")
        self.leFond.setEnabled(not self.leFond.isEnabled())
        self.cmbFond.setEnabled(not self.cmbFond.isEnabled())
        return

    def click_pbSigner(self):
        self.frSigner.setStyleSheet("QFrame{background-image: }")
        self.leSigner.setEnabled(not self.leSigner.isEnabled())
        self.cmbSigner.setEnabled(not self.cmbSigner.isEnabled())
        return

    def set_cmbAgent(self):
        self.agent_touched = True
        return

    def set_cmbFond(self):
        self.fond_touched = True
        dbconn = MySQLConnection(**self.dbconfig)
        cursor = dbconn.cursor()
        if self.leSigner.text().strip():
            if self.fond_touched:
                sql = "SELECT CONCAT_WS(' ', id, '-', signer_surname, signer_name, signer_lastname), id " \
                      "FROM signers " \
                      "WHERE CONCAT_WS(' ', id, '-', signer_surname, signer_name, signer_lastname) LIKE %s " \
                      "AND subdomain_id = %s"
                cursor.execute(sql, (self.fond_ids[self.cmbFond.currentIndex()],))
            else:
                sql = "SELECT CONCAT_WS(' ', id, '-', signer_surname, signer_name, signer_lastname), id " \
                      "FROM signers " \
                      "WHERE CONCAT_WS(' ', id, '-', signer_surname, signer_name, signer_lastname) LIKE %s " \
                      "AND subdomain_id IN (2,6,8,11,12,13,14) AND subdomain_id IN (" + self.fonds_str + ")"
                cursor.execute(sql, ('%' + self.leSigner.text() + '%',))
        else:
            if self.fond_touched:
                sql = "SELECT CONCAT_WS(' ', id, '-', signer_surname, signer_name, signer_lastname), id " \
                      "FROM signers WHERE subdomain_id = %s"
                cursor.execute(sql, (self.fond_ids[self.cmbFond.currentIndex()],))
            else:
                sql = "SELECT CONCAT_WS(' ', id, '-', signer_surname, signer_name, signer_lastname), id " \
                      "FROM signers WHERE subdomain_id IN (" + self.fonds_str + ") " \
                      "AND subdomain_id IN (2,6,8,11,12,13,14)"
                cursor.execute(sql)
        rows = cursor.fetchall()
        self.signer_ids = []
        self.signer_names = {}
        signers = []
        for i, row in enumerate(rows):
            self.signer_names[row[1]] = row[0]
            self.signer_ids.append(row[1])
            signers.append(row[0])
        self.cmbSigner.clear()
        self.cmbSigner.addItems(signers)
        return

    def set_cmbSigner(self):
        self.signer_touched = True
        return

    def set_cmbCfgFile(self):
        self.frCfgFile.setStyleSheet("QFrame{background-image: }")
        self.cfg_file_touched = True
        self.cfg_file_loaded = False

    def click_clbXAgent(self):
        self.agent_touched = False
        self.leAgent.setText('')
        return

    def click_clbXSigner(self):
        self.signer_touched = False
        self.leSigner.setText('')
        return

    def click_clbXFond(self):
        self.fond_touched = False
        self.leFond.setText('')
        dbconn = MySQLConnection(**self.dbconfig)
        cursor = dbconn.cursor()
        if self.leSigner.text().strip():
            if self.fond_touched:
                sql = "SELECT CONCAT_WS(' ', id, '-', signer_surname, signer_name, signer_lastname), id " \
                      "FROM signers " \
                      "WHERE CONCAT_WS(' ', id, '-', signer_surname, signer_name, signer_lastname) LIKE %s " \
                      "AND subdomain_id = %s"
                cursor.execute(sql, (self.fond_ids[self.cmbFond.currentIndex()],))
            else:
                sql = "SELECT CONCAT_WS(' ', id, '-', signer_surname, signer_name, signer_lastname), id " \
                      "FROM signers " \
                      "WHERE CONCAT_WS(' ', id, '-', signer_surname, signer_name, signer_lastname) LIKE %s " \
                      "AND subdomain_id IN (2,6,8,11,12,13,14) AND subdomain_id IN (" + self.fonds_str + ")"
                cursor.execute(sql, ('%' + self.leSigner.text() + '%',))
        else:
            if self.fond_touched:
                sql = "SELECT CONCAT_WS(' ', id, '-', signer_surname, signer_name, signer_lastname), id " \
                      "FROM signers WHERE subdomain_id = %s"
                cursor.execute(sql, (self.fond_ids[self.cmbFond.currentIndex()],))
            else:
                sql = "SELECT CONCAT_WS(' ', id, '-', signer_surname, signer_name, signer_lastname), id " \
                      "FROM signers WHERE subdomain_id IN (" + self.fonds_str + ") " \
                      "AND subdomain_id IN (2,6,8,11,12,13,14)"
                cursor.execute(sql)
        rows = cursor.fetchall()
        self.signer_ids = []
        self.signer_names = {}
        signers = []
        for i, row in enumerate(rows):
            self.signer_names[row[1]] = row[0]
            self.signer_ids.append(row[1])
            signers.append(row[0])
        self.cmbSigner.clear()
        self.cmbSigner.addItems(signers)
        return

    def change_leDir(self):
        try:
            files = os.listdir(path=self.leDir.text())
        except OSError:
            return
        self.cmbFile.clear()
        self.cmbFile.addItems(files)
        self.twAllExcels.setColumnCount(0)
        self.twAllExcels.setRowCount(0)
        self.file_touched = False
        self.file_loaded = False
        self.file_name = ''
        return

    def click_chbDateFrom(self):
        if self.deDateFrom.isEnabled():
            self.deDateFrom.setEnabled(False)
        else:
            self.deDateFrom.setEnabled(True)

    def click_chbDateTo(self):
        if self.deDateTo.isEnabled():
            self.deDateTo.setEnabled(False)
        else:
            self.deDateTo.setEnabled(True)


    def set_cmbFile(self):
        if self.leDir.text()[len(self.leDir.text()) - 1:] != '/':
            self.leDir.setText(self.leDir.text() + '/')
        if self.cmbFile.currentText()[len(self.cmbFile.currentText()) - 5:] == '.xlsx':
            self.wb = openpyxl.load_workbook(filename=self.leDir.text() + self.cmbFile.currentText(),
                                        read_only=True)
            self.file_name = self.leDir.text() + self.cmbFile.currentText()
            self.cmbTab.clear()
            self.cmbTab.addItems(self.wb.sheetnames)
        else:
            return

    def set_cmbTab(self):
        self.frFile.setStyleSheet("QFrame{background-image: }")
        self.file_touched = True
        if self.MoveImportPasport == 1:
            self.load4move()
        elif self.MoveImportPasport == 2:
            self.load4import()
        return

    def selectAction(self):
        if self.MoveImportPasport == 1:
            self.frMove.show()
            self.frMoveInf.show()
            self.frImport.hide()
            self.frImportInf.hide()
            self.frPasport.hide()
            self.frPasportInf.hide()
            self.twParsingResult.hide()
        elif self.MoveImportPasport == 2:
            self.frImport.show()
            self.frImportInf.show()
            self.frMove.hide()
            self.frMoveInf.hide()
            self.frPasport.hide()
            self.frPasportInf.hide()
            self.twParsingResult.show()
        else:
            self.frPasport.show()
            self.frPasportInf.show()
            self.frMove.hide()
            self.frMoveInf.hide()
            self.frImport.hide()
            self.frImportInf.hide()
            self.twParsingResult.hide()

    def click_clbMove(self):
        self.MoveImportPasport = 2
        self.selectAction()

    def click_clbImport(self):
        self.MoveImportPasport = 3
        self.selectAction()

    def click_clbPasport(self):
        self.MoveImportPasport = 1
        self.selectAction()

    def load4move(self):
        self.sheet = self.wb[self.wb.sheetnames[self.cmbTab.currentIndex()]]
        if not self.sheet.max_row:
            self.errMessage('Файл Excel некорректно сохранен OpenPyxl. Откройте и пересохраните его')
            return
        keys = {}
        last_cell = 0
        for j, row in enumerate(self.sheet.rows):
            if j == 0:
                for k, cell in enumerate(row):  # Проверяем, чтобы был СНИЛС
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
                    self.errMessage('В файле ' + self.cmbFile.currentText() + ' на вкладке ' + self.cmbTab.currentText() +
                                    ' отсутствует колонка с ID')
                    return
            elif j == 1:
                for k, cell in enumerate(row):
                    if cell.value != None:
                        if str(cell.value).strip() != '':
                            if k > last_cell:
                                last_cell = k

        self.twAllExcels.setColumnCount(len(keys))                 # Устанавливаем кол-во колонок
        self.twAllExcels.setRowCount(len(list(self.sheet.rows)) - 1)   # Кол-во строк из таблицы
        self.clients_ids = []
        for j, row in enumerate(self.sheet.rows):
            if j == 0:
                continue
            for k, key in enumerate(keys):
                self.twAllExcels.setItem(j - 1, k, QTableWidgetItem(str(row[keys[key]].value)))
                if k == 0:
                    self.clients_ids.append(row[keys[key]].value)

        # Устанавливаем заголовки таблицы
        self.twAllExcels.setHorizontalHeaderLabels(list(keys))
        # Устанавливаем выравнивание на заголовки
        self.twAllExcels.horizontalHeaderItem(0).setTextAlignment(Qt.AlignCenter)
        # делаем ресайз колонок по содержимому
        self.twAllExcels.resizeColumnsToContents()
        self.file_loaded = True
        self.table_loaded = False
        if self.leDir.text()[len(self.leDir.text()) - 1:] != '/':
            self.leDir.setText(self.leDir.text() + '/')
        self.file_name = self.leDir.text() + self.cmbFile.currentText() + "!'" + self.cmbTab.currentText() + "'"
        self.table = []
        for i, row in enumerate(self.sheet.rows):
            table_row = []
            for j, cell in enumerate(row):
                if j > last_cell:
                    break
                table_row.append(cell.value)
            self.table.append(table_row)
        return

# !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! ПРОВЕРКА ПАСПОРТА !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

    def click_pbIdGenerate4PasportCheck(self):
        self.refresh()
        self.frPasport.setStyleSheet("QFrame{background-image: }")
                                                                                # Формируем запрос
        sql_cl = 'SELECT client_id, p_seria, p_number, number, ' \
                 'p_surname, p_name, p_lastname FROM clients AS cl WHERE'
        if self.leAgent.isEnabled():
            sql_cl += ' cl.inserted_user_code = %s'
        if self.leFond.isEnabled():
            if sql_cl[len(sql_cl) - 2:] == '%s':
                sql_cl += ' AND'
            sql_cl += ' cl.subdomain_id = %s'
        if self.chbDateFrom.isChecked():
            if sql_cl[len(sql_cl) - 2:] == '%s':
                sql_cl += ' AND'
            sql_cl += ' cl.inserted_date >= %s'
        if self.chbDateTo.isChecked():
            if sql_cl[len(sql_cl) - 2:] == '%s':
                sql_cl += ' AND'
            sql_cl += ' cl.inserted_date <= %s'
        if sql_cl[len(sql_cl) - 5:] == 'WHERE':
            self.errMessage('Нет ни одного ЗЛ')
            return
        tuple_client = tuple()                                  # Формируем переменные для запросов
        if self.leAgent.isEnabled():
            tuple_client += (self.agent_ids[self.cmbAgent.currentIndex()],)
        if self.leFond.isEnabled():
            tuple_client += (self.fond_ids[self.cmbFond.currentIndex()],)
        if self.chbDateFrom.isChecked():
            tuple_client += (self.deDateFrom.dateTime().toPyDateTime(),)
        if self.chbDateTo.isChecked():
            tuple_client += (self.deDateTo.dateTime().toPyDateTime() + timedelta(hours=23,minutes=59),)
        dbconn = MySQLConnection(**self.dbconfig)
        cursor = dbconn.cursor()
        cursor.execute(sql_cl, tuple_client)
        rows = cursor.fetchall()

        if len(rows):
            self.twAllExcels.setColumnCount(len(rows[0]))                 # Устанавливаем кол-во колонок
            self.twAllExcels.setRowCount(len(rows))   # Кол-во строк из таблицы
            self.clients_ids = []
            for j, row in enumerate(rows):
                for k, cell in enumerate(row):
                    self.twAllExcels.setItem(j, k, QTableWidgetItem(str(cell)))
                    if k == 0:
                        self.clients_ids.append(str(cell))

            # Устанавливаем заголовки таблицы
            self.twAllExcels.setHorizontalHeaderLabels(cursor.column_names)
            # Устанавливаем выравнивание на заголовки
            self.twAllExcels.horizontalHeaderItem(0).setTextAlignment(Qt.AlignCenter)
            # делаем ресайз колонок по содержимому
            self.twAllExcels.resizeColumnsToContents()
            self.table = rows
            self.table_loaded = True
            self.file_loaded = False
        else:
            self.errMessage('Нет ЗЛ по данному запросу')

    def click_pbPasportCheck(self):
        if self.leAgent.isEnabled() and not self.agent_touched: # Проверяем достаточность данных
            self.frAgent.setStyleSheet("QFrame{background-image: url(./x.png)}")
            return
        if self.leFond.isEnabled() and not self.fond_touched:
            self.frFond.setStyleSheet("QFrame{background-image: url(./x.png)}")
            return
        if not self.table_loaded:
            self.frPasport.setStyleSheet("QFrame{background-image: url(./x.png)}")
            return
                                                                # Создаем файл с исходными данными и логом
        wb_log = openpyxl.Workbook(write_only=True)

        ws_log = wb_log.create_sheet('Лог')
        ws_log.append([datetime.now().strftime("%H:%M:%S"), ' Начинаем'])

        log_name = DIR4PCHECK + datetime.now().strftime('%Y-%m-%d_%H-%M')
        if self.fond_touched:
            log_name += 'ф' + str(self.fond_ids[self.cmbFond.currentIndex()])
        if self.agent_touched:
            log_name += 'а' + str(self.agent_ids[self.cmbAgent.currentIndex()])
        log_name += '.xlsx'

        #ws_log.append([datetime.now().strftime("%H:%M:%S"), 'Дублируем исходную excel таблицу в этот файл'])
        #ws_input = wb_log.create_sheet('Исходная таблица')
        #for table_row in self.table:
        #    row = []
        #    for cell in table_row:
        #        row.append(cell)
        #    ws_input.append(row)

        ws_log.append([datetime.now().strftime("%H:%M:%S"), ' Состояние программы:'])
        if self.leFond.isEnabled():
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'фонд', self.cmbFond.currentText()])
        else:
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'фонд', 'не выбран'])
        if self.leAgent.isEnabled():
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'агент', self.cmbAgent.currentText()])
        else:
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'агент', 'не выбран'])
        if self.chbDateFrom.isChecked():
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'Выборка ОТ', self.deDateFrom.date().toPyDate()])
        else:
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'Выборка ОТ', 'не выбрана'])
        if self.chbDateTo.isChecked():
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'Выборка ДО', self.deDateTo.date().toPyDate()])
        else:
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'Выборка ДО', 'не выбрана'])

        #ws_log.append([datetime.now().strftime("%H:%M:%S"), 'Бэкап исходного состояния БД создан'])
        #all_clients_ids = "'" + self.clients_ids[0] + "'"       # Проверка на дубли clients
        #for i, client_id in enumerate(self.clients_ids):
        #    if i == 0:
        #        continue
        #    all_clients_ids += ",'" + client_id + "'"
        #ws_backup = wb_log.create_sheet('бэкап БД')
        #dbconn = MySQLConnection(**self.dbconfig)
        #cursor = dbconn.cursor()
        #sql = "SELECT cl.*, co.* FROM clients AS cl LEFT JOIN contracts AS co " \
        #      "ON (cl.client_id = co.client_id) WHERE cl.client_id IN (" + all_clients_ids + ")"
        #cursor.execute(sql)
        #dbrows = cursor.fetchall()
        #ws_backup.append(cursor.column_names)
        #for dbrow in dbrows:
        #    row = []
        #    for dbcell in dbrow:
        #        row.append(dbcell)
        #    ws_backup.append(row)

        p = Popen(['ls', '-l', '--time-style=long-iso', 'list_of_expired_passports.csv.bz2'], stdout=PIPE)
        out, err = p.communicate()
        has_passports = False
        if not p.returncode and out:
            if datetime.now() - timedelta(days=3) < datetime.strptime(out.decode('utf-8').split(' ')[5] +
                                                    ' ' + out.decode('utf-8').split(' ')[6], '%Y-%m-%d %H:%M'):
                has_passports = True

        if not has_passports:
            try:
                os.remove('list_of_expired_passports.csv.bz2')
            except:
                q = 0
            try:
                os.remove('list_of_expired_passports.csv')
            except:
                q = 0

            i = 0                                   # Если база паспортов с ГУМВД устаревшая - скачиваем
            ok = 1
            while ok != 0 and i < 10:
                p = Popen(['wget', '-q', '-c', '-t5',
                           'https://guvm.mvd.ru/upload/expired-passports/list_of_expired_passports.csv.bz2'], stdout=PIPE)
                out, err = p.communicate()
                ok = p.returncode
                i += 1
            if i >= 10:
                print(datetime.now().strftime("%d.%m.%Y %H:%M:%S"), ' Не скачивается, наверное погода нелетная :)')
                return

            all_files = os.listdir(path=".")        # Распаковываем все bzip2 в директории
            for i, all_file in enumerate(all_files):
                if all_file.endswith(".bz2"):
                    with open(all_file.replace('.bz2', ''), 'wb') as new_file, bz2.BZ2File(all_file, 'rb') as file:
                        for data in iter(lambda: file.read(100 * 1024), b''):
                            new_file.write(data)

        has_files = False                       # Проверяем есть ли .csv
        all_files = os.listdir(path=".")
        for all_file in all_files:
            if all_file.endswith(".csv"):
                has_files = True
                new_csv = all_file
        if not has_files:
            print(datetime.now().strftime("%H:%M:%S"), ' В скачанном архиве нет .csv')
            try:
                os.remove('list_of_expired_passports.csv.bz2')
            except:
                q = 0
                return

        if not len(self.passports):
            self.progressBar.setMaximum(118000000)
            with open("list_of_expired_passports.csv","rt") as file_passports:
                for i,line in enumerate(file_passports):
                    if i:
                        if not i%1000000:
                            self.progressBar.setValue(i)
                        self.passports[l(line)] = 1

        self.progressBar.setMaximum(len(self.table)-1)
        ws_pasport = wb_log.create_sheet('Проверка паспортов')
        ws_pasport.append(['ID', 'Серия', 'Номер', 'СНИЛС', 'Фамилия', 'Имя', 'Отчество', 'Проверка паспорта'])  # добавляем первую строку xlsx
        dbconn_saturn = MySQLConnection(**self.dbconfig)
        bad_passport_ids = []
        for j, row in enumerate(self.table):                            # Проверяем паспорта из таблицы
            check = l(row[1])*1000000 + l(row[2])
            try:
                q = self.passports[check]
                rez = 'плохой'
                if self.chbSetStatusInSaturn.isChecked():
                    bad_passport_ids.append((row[0],))
            except KeyError:
                rez = 'OK'

#            if check in self.passports:
#                rez = 'плохой'
#                if self.chbSetStatusInSaturn.isChecked():
#                    bad_passport_ids.append((row[0],))
#            else:
#                rez = 'OK'
#--------------------------------------------------------------------------
#            try:
#                q = self.passports.index(l(row[1])*1000000 + l(row[2]))
#                rez = 'плохой'
#                if self.chbSetStatusInSaturn.isChecked():
#                    bad_passport_ids.append((row[0],))
#            except ValueError:
#                rez = 'OK'
#---------------------------------------------------------------------------
#            for passport in self.passports:
#                if l(row[1]) == l(passport)// 1000000 and l(row[2]) == l(passport)  % 1000000:
#                    rez = 'плохой'
#                    if self.chbSetStatusInSaturn.isChecked():
#                        bad_passport_ids.append((row[0],))
#                    break
            ws_pasport.append([row[0], row[1], row[2], row[3], row[4], row[5], row[6], rez])
            if not j%100:
                self.progressBar.setValue(j)
            if len(bad_passport_ids) and not len(bad_passport_ids)%1000:
                write_cursor = dbconn_saturn.cursor()
                write_cursor.executemany('UPDATE contracts AS co SET co.status_secure_code = 6 WHERE co.client_id = %s',
                                         bad_passport_ids)
                dbconn_saturn.commit()
                bad_passport_ids = []
        if len(bad_passport_ids):
            write_cursor = dbconn_saturn.cursor()
            write_cursor.executemany('UPDATE contracts AS co SET co.status_secure_code = 6 WHERE co.client_id = %s',
                                     bad_passport_ids)
            dbconn_saturn.commit()
        self.progressBar.setValue(0)



        qq = """
        self.progressBar.setMaximum(len(self.table)-1)
        ws_pasport = wb_log.create_sheet('Проверка паспортов')
        ws_pasport.append(['ID', 'Серия', 'Номер', 'СНИЛС', 'Фамилия', 'Имя', 'Отчество', 'Проверка паспорта'])  # добавляем первую строку xlsx
        dbconn_pasp = MySQLConnection(**self.dbconfig_pasp)
        dbconn_saturn = MySQLConnection(**self.dbconfig)
        for j, row in enumerate(self.table):                            # Проверяем паспорта из таблицы
            rez = 'OK'
            read_cursor = dbconn_pasp.cursor()
            read_cursor.execute('SELECT p_seria, p_number FROM passport_greylist WHERE p_seria = %s AND p_number = %s',
                                (l(row[1]), l(row[2])))
            row_msg = read_cursor.fetchall()
            if len(row_msg) > 0:
                if self.chbSetStatusInSaturn.isChecked():
                    write_cursor = dbconn_saturn.cursor()
                    write_cursor.execute('UPDATE contracts AS co SET co.status_secure_code = 6 WHERE co.client_id = %s',
                                         (row[0],))
                rez = 'плохой'
            else:
                rez = 'ОК'
            ws_pasport.append([row[0], row[1], row[2], row[3], row[4], row[5], row[6], rez])
            self.progressBar.setValue(j)
        dbconn_saturn.commit()
        """

        wb_log.save(log_name)

    #!!!!!!!!!!!!!!!!!!!!! IMPORT !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

    def click_pbSaveCfgFile(self):
        wb_cfg = openpyxl.Workbook(write_only=True)
        ws_cfg = wb_cfg.create_sheet('Конфиг')
        for i in range(self.tableWidget.rowCount()):
            ws_cfg.append([self.tableWidget.cellWidget(i,0).currentIndex(), self.tableWidget.cellWidget(i,1).currentIndex()])
        if self.file_loaded:
            wb_cfg.save(DIR4CFGIMPORT + self.cmbFile.currentText() + '.xlsx')
        rows = os.listdir(DIR4CFGIMPORT)                        # обновляем список конфигов
        cfg_files = []
        for row in rows:
            if row.find('.xlsx') > -1:
                cfg_files.append(row)
        self.cfg_file_names = {}
        for i, cfg_file in enumerate(cfg_files):
            self.cfg_file_names[cfg_file] = i
        self.cmbCfgFile.clear()
        self.cmbCfgFile.addItems(cfg_files)


    def click_pbRefreshImport(self):
        self.refresh()
        self.load4import()

    def updateProgressBar(self, val):
        self.progressBar.setValue(val)

    def load4import(self):  # Заполняем табличку
        if not self.file_touched:                               # Проверяем достаточность данных
            self.frFile.setStyleSheet("QFrame{background-image: url(./x.png)}")
            return
        if not self.cfg_file_touched:
            self.frCfgFile.setStyleSheet("QFrame{background-image: url(./x.png)}")
            return

        self.progressBar.setMaximum(0)
        self.sheet = self.wb[self.wb.sheetnames[self.cmbTab.currentIndex()]]
        if not self.sheet.max_row:
            self.errMessage('Файл Excel некорректно сохранен OpenPyxl. Откройте и пересохраните его')
            return
        head = []
        for i, row in enumerate(self.sheet.rows):
            if i == 0:
                for cell in row:
                    head.append(cell.value)
                break
        self.progressBar.setMaximum(self.sheet.max_row - 1)

        if not self.cfg_file_loaded:
            for j in range(self.tableWidget.rowCount()):
                self.tableWidget.removeRow(0)
            conf_mass = []
            mass = []
            wb_conf = openpyxl.load_workbook(filename=DIR4CFGIMPORT + self.cmbCfgFile.currentText(),
                                             read_only=True)
            sheet_conf = wb_conf[wb_conf.sheetnames[0]]
            try:
                for row in sheet_conf:
                    conf_mas = []
                    for j, cell in enumerate(row):
                        if j < 2:
                            conf_mas.append(int(cell.value))
                    conf_mass.append(conf_mas)
            except ValueError:
                self.cfg_file_loaded = False
            else:
                self.cfg_file_loaded = True

            for i in FIELDS_IN_RESULT_TABLE_COMPLETE:
                for name in i:
                    mass.append(name)

            for i, j in enumerate(mass):
                if i < len(conf_mass) and self.cfg_file_loaded:
                    #                if len(conf_mass[i]) < len(mass):
                    if len(conf_mass[i]) > 1:
                        combo1index = conf_mass[i][0]
                        combo2index = conf_mass[i][1]
                    else:
                        combo1index = i  # по умолчанию
                        combo2index = i
                else:
                    combo1index = i  # по умолчанию
                    combo2index = i
                                                        # Заполняем комбобоксы в таблице
                try:  # combo2index - индекс выбора второго QCombobox
                    if head is None:
                        pass
                except AttributeError:
                    return
                self.tableWidget.insertRow(self.tableWidget.rowCount().real)
                items = []

                self.combobox_table_result = QComboBox()
                # self.combobox_table_result.setMaxVisibleItems(15)
                for row in FIELDS_IN_RESULT_TABLE_COMPLETE:
                    for name in row:
                        self.combobox_table_result.addItem(name)
                items.append(self.combobox_table_result)
                if combo1index != -1:
                    self.combobox_table_result.setCurrentIndex(combo1index)  # combobox_table_result - первый комбобокс
                name_combobox_table_result = "combobox_table_result_{0}".format(self.tableWidget.rowCount() - 1)
                self.combobox_table_result.setObjectName(name_combobox_table_result)

                self.combobox_table_from = QComboBox()
                # self.combobox_table_from.setMaxVisibleItems(15)
                for name in head:
                    name = str(name)
                    self.combobox_table_from.addItem(name)
                items.append(self.combobox_table_from)
                if combo2index != -1:
                    self.combobox_table_from.setCurrentIndex(combo2index)  # combobox_table_from - второй комбобокс
                name_combobox_table_from = "combobox_table_from_{0}".format(self.tableWidget.rowCount() - 1)
                self.combobox_table_from.setObjectName(name_combobox_table_from)

                for n, i in enumerate(items):
                    self.tableWidget.setCellWidget(self.tableWidget.rowCount() - 1, n, i)
            self.tableWidget.resizeColumnsToContents()

        keys = {}
        last_cell = 0
        for j, row in enumerate(self.sheet.rows):
            if j == 0:
                for k, cell in enumerate(row):  # Проверяем, чтобы был СНИЛС
                    if str(cell.value).strip().upper() in IN_SNILS:
                        keys[IN_SNILS[0]] = k
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
                    self.errMessage('В файле ' + self.cmbFile.currentText() + ' на вкладке ' + self.cmbTab.currentText() +
                                    ' отсутствует колонка со СНИЛС')
                    return
            elif j == 1:
                for k, cell in enumerate(row):
                    if cell.value != None:
                        if str(cell.value).strip() != '':
                            if k > last_cell:
                                last_cell = k

        self.clients_snils = []                                             # Добавляем СНИЛСы
        for i, row in enumerate(self.sheet):
            if i == 1:
                continue
            self.clients_snils.append(l(row[keys[IN_SNILS[0]]].value))

        self.twAllExcels.setColumnCount(last_cell + 1)                      # Отображаем исходную таблицу
        self.twAllExcels.setRowCount(len(list(self.sheet.rows)) - 1)
        self.clients_ids = []
        headers = []
        for j, row in enumerate(self.sheet.rows):
            for k, cell in enumerate(row):
                if j == 0:
                    headers.append(s(cell.value))
                else:
                    self.twAllExcels.setItem(j - 1, k, QTableWidgetItem(s(cell.value)))

        self.twAllExcels.setHorizontalHeaderLabels(headers)             # Устанавливаем заголовки таблицы
        # Устанавливаем выравнивание на заголовки
        self.twAllExcels.horizontalHeaderItem(0).setTextAlignment(Qt.AlignCenter)
        # делаем ресайз колонок по содержимому
        self.twAllExcels.resizeColumnsToContents()
        self.file_loaded = True
        self.table_loaded = False

        for i, row in enumerate(self.sheet.rows):
            table_row = []
            for j, cell in enumerate(row):
                if j > last_cell:
                    break
                table_row.append(cell.value)
            self.table.append(table_row)
        self.previewImport()

    def errMessage(self, err_text):  ## Method to open a message box
        infoBox = QMessageBox()  ##Message Box that doesn't run
        infoBox.setIcon(QMessageBox.Warning)
        infoBox.setText(err_text)
        #        infoBox.setInformativeText("Informative Text")
        infoBox.setWindowTitle(datetime.strftime(datetime.now(), "%H:%M:%S") + ' Ошибка: ')
        #        infoBox.setDetailedText("Detailed Text")
        #        infoBox.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        infoBox.setStandardButtons(QMessageBox.Ok)
        #        infoBox.setEscapeButton(QMessageBox.Close)
        infoBox.exec_()


    def click_pbImport(self):
        if not self.file_loaded:                               # Проверяем достаточность данных
            self.frFile.setStyleSheet("QFrame{background-image: url(./x.png)}")
            return
        if not self.cfg_file_loaded:
            self.frCfgFile.setStyleSheet("QFrame{background-image: url(./x.png)}")
            return
        if self.leAgent.isEnabled() and not self.agent_touched:
            self.frAgent.setStyleSheet("QFrame{background-image: url(./x.png)}")
            return
        if self.leFond.isEnabled() and not self.fond_touched:
            self.frFond.setStyleSheet("QFrame{background-image: url(./x.png)}")
            return
        if self.leSigner.isEnabled() and not self.signer_touched:
            self.frSigner.setStyleSheet("QFrame{background-image: url(./x.png)}")
            return

        if self.cmbGenderType == 0:
            female_gender_value = 'Ж'
            male_gender_value = 'М'
            gender_length = 1
        elif self.cmbGenderType == 1:
            female_gender_value = 'Ж'
            male_gender_value = 'М'
            gender_length = 1
        else:
            female_gender_value = 'Ж'
            male_gender_value = 'М'
            gender_length = 1

            # Создаем директорию
        if datetime.now().strftime("%Y-%m-%d") not in os.listdir(DIR4IMPORT):
            os.mkdir(DIR4IMPORT + datetime.now().strftime("%Y-%m-%d"))
        dir_name = DIR4IMPORT + datetime.now().strftime("%Y-%m-%d") + '/' + \
                   datetime.now().strftime('%H-%M_')
        if self.fond_touched:
            dir_name += 'ф' + str(self.fond_ids[self.cmbFond.currentIndex()])
        if self.agent_touched:
            dir_name += 'а' + str(self.agent_ids[self.cmbAgent.currentIndex()])
        dir_name += '/'
        os.mkdir(dir_name)
        self.file_name = dir_name + self.cmbFile.currentText()
                                                                        # Создаем файл с исходными данными и логом
        log_name = dir_name + self.cmbFile.currentText() + '_' + self.cmbTab.currentText() + '.xlsx'
        wb_log = openpyxl.Workbook(write_only=True)
        ws_log = wb_log.create_sheet('Лог')
        ws_log.append([datetime.now().strftime("%H:%M:%S"), ' Начинаем'])

                                                                        # Проверка на дубли исходной таблицы
        doubles_in_input = list(set([x for x in self.clients_snils if self.clients_snils.count(x) > 1]))
        if len(doubles_in_input):
            ws_log.append([datetime.now().strftime("%H:%M:%S"), ' Дубли в СНИЛС в исходной таблице'])
            ws_input_doubles = wb_log.create_sheet('Дубли в СНИЛС в исходной таблице')
            ws_input_doubles.append(['ID'])
            for row in doubles_in_input:
                ws_input_doubles.append([normalize_snils(row)])
        ws_log.append([datetime.now().strftime("%H:%M:%S"), ' Состояние программы:'])
        ws_log.append([datetime.now().strftime("%H:%M:%S"), 'Исходный файл ', self.file_name])
        ws_log.append([datetime.now().strftime("%H:%M:%S"), 'Конфигурационный файл ', self.cmbCfgFile.currentText()])
        if self.leFond.isEnabled():
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'фонд', self.cmbFond.currentText()])
        else:
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'фонд', 'не выбран'])
        if self.leAgent.isEnabled():
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'агент', self.cmbAgent.currentText()])
        else:
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'агент', 'не выбран'])
        if self.leSigner.isEnabled():
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'подписант', self.cmbSigner.currentText()])
        else:
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'подписант', 'не выбран'])

        ws_log.append([datetime.now().strftime("%H:%M:%S"), 'Дублируем исходную excel таблицу в файл лога'])
        ws_input = wb_log.create_sheet('Исходная таблица')
        for table_row in self.table:
            row = []
            for cell in table_row:
                row.append(cell)
            ws_input.append(row)

        self.updateProgressBar(0)
        self.pbImport.setEnabled(False)

        self.workerThread = WorkerThread(sheet=self.sheet, tableWidget=self.tableWidget,
                                         fname=self.file_name, agent=self.agent_ids[self.cmbAgent.currentIndex()],
                                         signer=self.signer_ids[self.cmbSigner.currentIndex()])  # <<<<<<<<<<<<<<<<<<<<<<<<<запускаем подпроцесс
        self.workerThread.progress_value.connect(self.updateProgressBar)
        self.workerThread.start()
        self.updateProgressBar(0)
        self.pbImport.setEnabled(True)
        ws_log.append([datetime.now().strftime("%H:%M:%S"), 'Импорт отработал, файл(ы) создан(ы)'])
        wb_log.save(log_name)


    def previewImport(self):
        err_from_log = {}
        self.twParsingResult.setColumnCount(0)
        self.twParsingResult.setRowCount(0)
        self.twParsingResult.setColumnCount(len(HEAD_RESULT_EXCEL_FILE))
        self.twParsingResult.setHorizontalHeaderLabels(HEAD_RESULT_EXCEL_FILE)  # добавляем заголовки
        self.twParsingResult.setRowCount(3)
        maxParsingResult = 0
        for num_row, row in enumerate(self.sheet.rows):
            if num_row == 0:
                continue
            result_row = {}
            passport = Passport()
            phone = Phone()

            for num_item in range(self.tableWidget.rowCount()):
                item0 = self.tableWidget.cellWidget(num_item, 0).currentIndex()
                item1 = self.tableWidget.cellWidget(num_item, 1).currentIndex()
                label0 = self.tableWidget.cellWidget(num_item, 0).currentText()
                label1 = self.tableWidget.cellWidget(num_item, 1).currentText()

                row_item = str(row[item1].value)  # Если преобразовывать все в стринг, то только тут
                if row_item == 'None' or row_item == '2001-01-00' or row_item == '2001-01-00 00:00:00' \
                        or row_item == 'null' or row_item == 'NULL' \
                        or row_item == '\\N' or row_item == '\\n' \
                        or row_item == 'заполнить' or row_item == '00.00.0000' \
                        or row_item == '0000-00-00' or row_item == 'ERROR' \
                        or row_item == '=#ССЫЛ!' or row_item == '#ССЫЛ!' \
                        or row_item == '=#REF!' or row_item == '#REF!' or row_item == '-':
                    row_item = ''
                elif row_item == '0' and label0 != 'Пол':
                    row_item = ''

                if label0 in MANIPULATE_LABELS:

                    if label0 in [MANIPULATE_LABELS[1], MANIPULATE_LABELS[3]]:
                        FIO = field2fio(row_item)
                        if label0 == MANIPULATE_LABELS[1]:
                            lab = FIO_LABELS
                        elif label0 == MANIPULATE_LABELS[3]:
                            lab = FIO_BIRTH_LABELS
                        if row_item == '':
                            for j in range(len(lab)):
                                result_row[lab[j]] = NEW_NULL_VALUE_FOR_ALL_TEXT
                        else:
                            for j in range(len(FIO)):
                                result_row[lab[j]] = FIO[j]
                        continue

                    # ------------------------------------------------------- Убрал класс Gender --------------------------------------
                    #                    elif label0 == "Пол_получить_из_ФИО":
                    #                        gender = Gender(row_item)
                    #                        result_row[GENDER_LABEL[0]] = gender.get_value()

                    #                    elif label0 == "Пол_подставить_свои_значения":
                    #                        gender = Gender(FIO[2], gender_field_exists=True, gender=row_item) ## !!!!!!!!!!!!!!
                    #                        result_row[GENDER_LABEL[0]] = gender.get_value()
                    # ------------------------------------------------------- Убрал класс Gender --------------------------------------
                    # Регистрация -> Регион
                    elif label0 == MANIPULATE_LABELS[5]:
                        addr = field2addr(row_item)
                        lab = [ADRESS_REG_LABELS[1], ADRESS_REG_LABELS[2]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
                    # Регистрация -> Район
                    elif label0 == MANIPULATE_LABELS[6]:
                        addr = field2addr(row_item)
                        lab = [ADRESS_REG_LABELS[3], ADRESS_REG_LABELS[4]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
                    # Регистрация -> Город
                    elif label0 == MANIPULATE_LABELS[7]:
                        addr = field2addr(row_item)
                        lab = [ADRESS_REG_LABELS[5], ADRESS_REG_LABELS[6]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
                    # Регистрация -> Населенный_пункт
                    elif label0 == MANIPULATE_LABELS[8]:
                        addr = field2addr(row_item)
                        lab = [ADRESS_REG_LABELS[7], ADRESS_REG_LABELS[8]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
                    # Регистрация -> Улица
                    elif label0 == MANIPULATE_LABELS[9]:
                        addr = field2addr(row_item)
                        lab = [ADRESS_REG_LABELS[9], ADRESS_REG_LABELS[10]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
                    # ADRESS_LIVE_LABELS
                    # Проживание -> Регион
                    elif label0 == MANIPULATE_LABELS[11]:
                        addr = field2addr(row_item)
                        lab = [ADRESS_LIVE_LABELS[1], ADRESS_LIVE_LABELS[2]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
                    # Проживание -> Район
                    elif label0 == MANIPULATE_LABELS[12]:
                        addr = field2addr(row_item)
                        lab = [ADRESS_LIVE_LABELS[3], ADRESS_LIVE_LABELS[4]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
                    # Проживание -> Город
                    elif label0 == MANIPULATE_LABELS[13]:
                        addr = field2addr(row_item)
                        lab = [ADRESS_LIVE_LABELS[5], ADRESS_LIVE_LABELS[6]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
                    # Проживание -> Населенный_пункт
                    elif label0 == MANIPULATE_LABELS[14]:
                        addr = field2addr(row_item)
                        lab = [ADRESS_LIVE_LABELS[7], ADRESS_LIVE_LABELS[8]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
                    # Проживание -> Улица
                    elif label0 == MANIPULATE_LABELS[15]:
                        addr = field2addr(row_item)
                        lab = [ADRESS_LIVE_LABELS[9], ADRESS_LIVE_LABELS[10]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
                    # Адрес регистрации из_поля
                    elif label0 == MANIPULATE_LABELS[17]:
                        result_row[ADRESS_REG_LABELS[0]] = '111111'
                        adress_reg = FullAdress(row_item)
                        #                        qr = ''
                        for z, cell in enumerate(adress_reg.get_values()):
                            result_row[ADRESS_REG_LABELS[z]] = cell
                        #                            qr += cell + ' '
                        #                        print(qr)
                        n = [char for char in result_row[ADRESS_REG_LABELS[0]] if char in string.digits]
                        if len(n) != 6:
                            result_row[ADRESS_REG_LABELS[0]] = '111111'

                    # Адрес проживания из поля
                    elif label0 == MANIPULATE_LABELS[19]:
                        result_row[ADRESS_LIVE_LABELS[0]] = '111111'
                        adress_zhit = FullAdress(row_item)
                        #                        qr = ''
                        for z, cell in enumerate(adress_zhit.get_values()):
                            result_row[ADRESS_LIVE_LABELS[z]] = cell
                        #                            qr += cell + ' '
                        #                        print(qr)
                        n = [char for char in result_row[ADRESS_LIVE_LABELS[0]] if char in string.digits]
                        if len(n) != 6:
                            result_row[ADRESS_LIVE_LABELS[0]] = '111111'

                    # Регион регистрации из номера
                    elif label0 == MANIPULATE_LABELS[21]:
                        addr = field2addr(REGIONS[intl(row_item)])
                        lab = [ADRESS_REG_LABELS[1], ADRESS_REG_LABELS[2]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]

                    # Регион проживания из номера
                    elif label0 == MANIPULATE_LABELS[23]:
                        addr = field2addr(REGIONS[intl(row_item)])
                        lab = [ADRESS_LIVE_LABELS[1], ADRESS_LIVE_LABELS[2]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]

                    # Серия и номер паспорта из поля
                    elif label0 == MANIPULATE_LABELS[25]:
                        addr = field2sernum(row_item)
                        lab = [PASSPORT_DATA_LABELS[0], PASSPORT_DATA_LABELS[1]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]

                elif label0 == '-------------------------':
                    continue
                elif label0 in SNILS_LABEL:
                    if GENERATE_SNILS:
                        dbconfig = read_config(filename='move.ini', section='mysql')
                        dbconn = MySQLConnection(**dbconfig)
                        count_snils = 1
                        cached_snils = 0
                        while count_snils > 0:
                            self.start_snils -= 1
                            checksum_snils = self.checksum(self.start_snils)
                            for i in range(0, 99):
                                if i != checksum_snils:
                                    full_snils = self.start_snils * 100 + i
                                    dbcursor = dbconn.cursor()
                                    dbcursor.execute('SELECT `number` FROM clients WHERE `number` = %s',
                                                     (full_snils,))
                                    dbchk = dbcursor.fetchall()
                                    if len(dbchk) == 0:
                                        cached_snils = full_snils
                                        count_snils -= 1
                        dbconn.close()
                        result_row[label0] = normalize_snils(cached_snils)
                        self.ws_comp.append([row_item, cached_snils])
                    else:
                        result_row[label0] = normalize_snils(row_item)
                elif label0 in PLACE_BIRTH_LABELS:
                    result_row[label0] = row_item
                elif label0 in PASSPORT_DATA_LABELS:
                    if PASSPORT_DATA_LABELS.index(label0) == 0:
                        result_row[label0] = normalize_seria(row_item)
                    elif PASSPORT_DATA_LABELS.index(label0) == 1:
                        result_row[label0] = normalize_index(row_item)
                    elif PASSPORT_DATA_LABELS.index(label0) == 2:
                        result_row[label0] = normalize_date(row_item)
                    elif PASSPORT_DATA_LABELS.index(label0) == 3:
                        result_row[label0] = normalize_text(row_item)
                    elif PASSPORT_DATA_LABELS.index(label0) == 4:
                        result_row[label0] = normalize_index(row_item)

                elif label0 in PHONES_LABELS:
                    if PHONES_LABELS.index(label0) == 0:
                        phone.tel_mob = row_item
                    elif PHONES_LABELS.index(label0) == 1:
                        phone.tel_rod = row_item
                    elif PHONES_LABELS.index(label0) == 2:
                        phone.tel_dom = row_item
                elif label0 in DATE_BIRTH_LABEL:
                    result_row[label0] = normalize_date(row_item)
                elif label0 in GENDER_LABEL:
                    result_row[label0] = normalize_gender(row_item)
                elif label0 == ADRESS_REG_LABELS[0] or label0 == ADRESS_LIVE_LABELS[0]:
                    result_row[label0] = normalize_index(row_item)
                elif label0 in ADRESS_REG_LABELS[11]:
                    result_row[label0] = normalize_home(row_item)
                elif label0 in ADRESS_LIVE_LABELS[11]:
                    result_row[label0] = normalize_home(row_item)
                elif label0 in ADRESS_REG_LABELS:
                    result_row[label0] = row_item
                elif label0 in ADRESS_LIVE_LABELS:
                    result_row[label0] = row_item
                elif label0 in TECH_LABELS:
                    if label0 == TECH_LABELS[0]:
                        result_row[label0] = self.agent_ids[self.cmbAgent.currentIndex()]
                    elif label0 == TECH_LABELS[1]:
                        result_row[label0] = self.signer_ids[self.cmbSigner.currentIndex()]
                    elif label0 == TECH_LABELS[2]:
                        result_row[label0] = PREDSTRAH_ID
                else:
                    result_row[label0] = normalize_text(row_item)

#            for num, z in enumerate(passport.get_values()):
#                result_row[PASSPORT_DATA_LABELS[num]] = z
            for num, z in enumerate(phone.get_values()):
                result_row[PHONES_LABELS[num]] = z

            LABELS = [SNILS_LABEL, FIO_LABELS, FIO_BIRTH_LABELS, FIO_SNILS_LABELS, GENDER_LABEL, DATE_BIRTH_LABEL,
                      PLACE_BIRTH_LABELS, PASSPORT_DATA_LABELS, ADRESS_REG_LABELS, ADRESS_LIVE_LABELS,
                      PHONES_LABELS, TECH_LABELS]
            mass = []
            for label_group in LABELS:
                for label in label_group:
                    mass.append(label)
            yum = True
            yum_phone0 = -1
            yum_phone1 = -1
            yum_phone2 = -1
            for num, cell in enumerate(mass):
                if cell in result_row:
                    mass[num] = result_row[cell]  # заполняем mass, чтобы его добавить как строку в xlsx
                    if cell == PHONES_LABELS[0]:
                        if mass[num] == ERROR_VALUE:
                            mass[num] = ''
                        yum_phone0 = num
                    elif cell == PHONES_LABELS[1]:
                        if mass[num] == ERROR_VALUE:
                            mass[num] = ''
                        yum_phone1 = num
                    elif cell == PHONES_LABELS[2]:
                        if mass[num] == ERROR_VALUE:
                            mass[num] = ''
                        yum_phone2 = num
                    elif mass[num] == ERROR_VALUE:
                        yum = False
                else:
                    mass[num] = ''  # заполняем mass, чтобы его добавить как строку в xlsx

            if mass[yum_phone0] == mass[yum_phone1] and mass[yum_phone0] != '':  # стираем дублирующиеся телефоны
                mass[yum_phone1] = ''
            if mass[yum_phone1] == mass[yum_phone2] and mass[yum_phone1] != '':
                mass[yum_phone2] = ''
            if mass[yum_phone0] == mass[yum_phone2] and mass[yum_phone0] != '':
                mass[yum_phone2] = ''

            if mass[yum_phone0] == '' and mass[yum_phone1] == '' and mass[yum_phone2] == '':
                yum = False  # если нет ни одного телефона - ошибка

            if yum and err_from_log.get(num_row + 1) == None:
                for ind, cell in enumerate(mass):
                    self.twParsingResult.setItem(maxParsingResult, ind, QTableWidgetItem(str(cell)))
                maxParsingResult += 1
                #ws.append(mass)
                #print(num_row, result_row['ФИО.Фамилия'], result_row['ФИО.Имя'], result_row['ФИО.Отчество'])
            else:
                mass.append(num_row + 1)
                mass.append(err_from_log.get(num_row + 1))
                #ws_err.append(mass)
                print('Ошибка:', num_row, result_row['ФИО.Фамилия'], result_row['ФИО.Имя'], result_row['ФИО.Отчество'])

            if maxParsingResult > 4:
                break
#        self.twParsingResult.horizontalHeaderItem(0).setTextAlignment(Qt.AlignCenter)
        # делаем ресайз колонок по содержимому
        self.twParsingResult.resizeColumnsToContents()

class WorkerThread(QThread):
    progress_value = QtCore.pyqtSignal(int)

    def __init__(self, tableWidget, sheet, fname, agent, signer, parent=None, ):
        super(WorkerThread, self).__init__(parent)
        self.tableWidget = tableWidget
        self.sheet = sheet
        self.fname = fname
        self.agent_id = agent
        self.signer_id = signer
        if GENERATE_SNILS:
            dbconfig = read_config(filename='move.ini', section='mysql')
            dbconn = MySQLConnection(**dbconfig)
            dbcursor = dbconn.cursor()
            dbcursor.execute('SELECT min(`number`) FROM  clients WHERE `number` > 99000000000 and subdomain_id = 2;')
            dbrows = dbcursor.fetchall()
            dbconn.close()
            self.start_snils = int('{0:011d}'.format(dbrows[0][0])[:-2])  # 9 цифр неправильного СНИЛСа с которого уменьшаем
            self.wb_comp = Workbook(write_only=True)
            self.ws_comp = self.wb_comp.create_sheet('Лист1')
            self.ws_comp.append(['Реальный СНИЛС', 'Псевдо-СНИЛС'])  # добавляем первую строку xlsx
        else:
            self.start_snils = 0

    def checksum(self, snils_dig):  # Вычисляем 2 последних цифры СНИЛС по первым 9-ти
        def snils_csum(sn):
            k = range(9, 0, -1)
            pairs = zip(k, [int(x) for x in sn.replace('-', '').replace(' ', '')])
            return sum([k * v for k, v in pairs])
        snils = '{0:09d}'.format(snils_dig)
        csum = snils_csum(snils)
        while csum > 101:
            csum %= 101
        if csum in (100, 101):
            csum = 0
        return csum

    def run(self):
        self.start_process()

    def start_process(self):
        lname = self.fname[0:self.fname.rfind('xlsx')]+ 'log'
        err_from_log = {}
        use_log = False
        try:
            log_file = open(lname,'rt',encoding='utf-8')
            log_file_string = log_file.read()
            first_sq = 0
            next_str = 0
            dub_toch = 0
            last_sq = 1
            n_str_w_err = ''
            text_err = ''
            for nx in range(len(log_file_string)):
                if log_file_string[nx] == ':':
                    dub_toch = nx
                if log_file_string[nx] == '\n':
                    next_str = nx
                if log_file_string[nx] == '#' or nx == len(log_file_string) - 1:
                    first_sq = last_sq
                    last_sq = nx
                    if dub_toch > 0:
                        n_str_w_err = int(log_file_string[first_sq+1:dub_toch])
                        text_err = log_file_string[dub_toch + 3:next_str]
                        err_from_log[n_str_w_err] = text_err
            use_log = True
        except:
            use_log = False

        cname = self.fname[0:self.fname.rfind('.xlsx')]+ '_new.cfg'
        conf_file = open(cname,'wt',encoding='utf-8')
        for i in range(self.tableWidget.rowCount()):
            conf_file.write(str(self.tableWidget.cellWidget(i,0).currentIndex()) + ' ' +
                            str(self.tableWidget.cellWidget(i,1).currentIndex()) + '\n')
        conf_file.close()

        wb_err = Workbook(write_only=True)
        ws_err = wb_err.create_sheet('Ошибки')
        ws_err.append(HEAD_RESULT_EXCEL_FILE)                                         # добавляем первую строку xlsx
        wb = Workbook(write_only=True)
        ws = wb.create_sheet('Лист1')
        ws.append(HEAD_RESULT_EXCEL_FILE)                                             # добавляем первую строку xlsx

        # --------------------------------------- Заменил первую строку xls файла---------------------------------------
        #        result_file_columns = [SNILS_LABEL, FIO_LABELS, FIO_BIRTH_LABELS, GENDER_LABEL, DATE_BIRTH_LABEL,
        #                            PLACE_BIRTH_LABELS, PASSPORT_DATA_LABELS, ADRESS_REG_LABELS, ADRESS_LIVE_LABELS,
        #                            PHONES_LABELS]

        #        listmerge = lambda result_file_columns: [col for label in result_file_columns for col in label] # заполняем первую строку xlsx
        #        head_result_file = listmerge(result_file_columns)


        #        ws.append(head_result_file)                                             # добавляем первую строку xlsx
        # --------------------------------------- Заменил первую строку xls файла ---------------------------------------

        file_number = 1
        for num_row, row in enumerate(self.sheet.rows):
            self.progress_value.emit(num_row + 1)  # отрисовываем ProgresBar
            if num_row == 0:
                continue
            result_row = {}
            passport = Passport()
            phone = Phone()

            for num_item in range(self.tableWidget.rowCount()):
                item0 = self.tableWidget.cellWidget(num_item, 0).currentIndex()
                item1 = self.tableWidget.cellWidget(num_item, 1).currentIndex()
                label0 = self.tableWidget.cellWidget(num_item, 0).currentText()
                label1 = self.tableWidget.cellWidget(num_item, 1).currentText()

                row_item = str(row[item1].value)                         #Если преобразовывать все в стринг, то только тут
                if row_item == 'None' or row_item == '2001-01-00' or row_item == '2001-01-00 00:00:00' \
                                      or  row_item == 'null' or  row_item == 'NULL' \
                                      or  row_item == '\\N' or  row_item == '\\n' \
                                      or  row_item == 'заполнить' or row_item == '00.00.0000'\
                                      or row_item == '0000-00-00' or row_item == 'ERROR' \
                                      or row_item == '=#ССЫЛ!' or row_item == '#ССЫЛ!'\
                                      or row_item == '=#REF!' or row_item == '#REF!' or row_item == '-':
                    row_item = ''
                elif row_item == '0' and label0 != 'Пол':
                    row_item = ''

                if label0 in MANIPULATE_LABELS:

                    if label0 in [MANIPULATE_LABELS[1], MANIPULATE_LABELS[3]]:
                        FIO = field2fio(row_item)
                        if label0 == MANIPULATE_LABELS[1]:
                            lab = FIO_LABELS
                        elif label0 == MANIPULATE_LABELS[3]:
                            lab = FIO_BIRTH_LABELS
                        if row_item == '':
                            for j in range(len(lab)):
                                result_row[lab[j]] = NEW_NULL_VALUE_FOR_ALL_TEXT
                        else:
                            for j in range(len(FIO)):
                                result_row[lab[j]] = FIO[j]
                        continue

#------------------------------------------------------- Убрал класс Gender --------------------------------------
#                    elif label0 == "Пол_получить_из_ФИО":
#                        gender = Gender(row_item)
#                        result_row[GENDER_LABEL[0]] = gender.get_value()

#                    elif label0 == "Пол_подставить_свои_значения":
#                        gender = Gender(FIO[2], gender_field_exists=True, gender=row_item) ## !!!!!!!!!!!!!!
#                        result_row[GENDER_LABEL[0]] = gender.get_value()
#------------------------------------------------------- Убрал класс Gender --------------------------------------
# Регистрация -> Регион
                    elif label0 == MANIPULATE_LABELS[5]:
                        addr = field2addr(row_item)
                        lab = [ADRESS_REG_LABELS[1], ADRESS_REG_LABELS[2]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
# Регистрация -> Район
                    elif label0 == MANIPULATE_LABELS[6]:
                        addr = field2addr(row_item)
                        lab = [ADRESS_REG_LABELS[3], ADRESS_REG_LABELS[4]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
# Регистрация -> Город
                    elif label0 == MANIPULATE_LABELS[7]:
                        addr = field2addr(row_item)
                        lab = [ADRESS_REG_LABELS[5], ADRESS_REG_LABELS[6]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
# Регистрация -> Населенный_пункт
                    elif label0 == MANIPULATE_LABELS[8]:
                        addr = field2addr(row_item)
                        lab = [ADRESS_REG_LABELS[7], ADRESS_REG_LABELS[8]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
# Регистрация -> Улица
                    elif label0 == MANIPULATE_LABELS[9]:
                        addr = field2addr(row_item)
                        lab = [ADRESS_REG_LABELS[9], ADRESS_REG_LABELS[10]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
# ADRESS_LIVE_LABELS
# Проживание -> Регион
                    elif label0 == MANIPULATE_LABELS[11]:
                        addr = field2addr(row_item)
                        lab = [ADRESS_LIVE_LABELS[1], ADRESS_LIVE_LABELS[2]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
# Проживание -> Район
                    elif label0 == MANIPULATE_LABELS[12]:
                        addr = field2addr(row_item)
                        lab = [ADRESS_LIVE_LABELS[3], ADRESS_LIVE_LABELS[4]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
# Проживание -> Город
                    elif label0 == MANIPULATE_LABELS[13]:
                        addr = field2addr(row_item)
                        lab = [ADRESS_LIVE_LABELS[5], ADRESS_LIVE_LABELS[6]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
# Проживание -> Населенный_пункт
                    elif label0 == MANIPULATE_LABELS[14]:
                        addr = field2addr(row_item)
                        lab = [ADRESS_LIVE_LABELS[7], ADRESS_LIVE_LABELS[8]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
# Проживание -> Улица
                    elif label0 == MANIPULATE_LABELS[15]:
                        addr = field2addr(row_item)
                        lab = [ADRESS_LIVE_LABELS[9], ADRESS_LIVE_LABELS[10]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
#Адрес регистрации из_поля
                    elif label0 == MANIPULATE_LABELS[17]:
                        result_row[ADRESS_REG_LABELS[0]] = '111111'
                        adress_reg = FullAdress(row_item)
#                        qr = ''
                        for z, cell in enumerate(adress_reg.get_values()):
                            result_row[ADRESS_REG_LABELS[z]] = cell
#                            qr += cell + ' '
#                        print(qr)
                        n = [char for char in result_row[ADRESS_REG_LABELS[0]] if char in string.digits]
                        if len(n) != 6:
                            result_row[ADRESS_REG_LABELS[0]] = '111111'

# Адрес проживания из поля
                    elif label0 == MANIPULATE_LABELS[19]:
                        result_row[ADRESS_LIVE_LABELS[0]] = '111111'
                        adress_zhit = FullAdress(row_item)
#                        qr = ''
                        for z, cell in enumerate(adress_zhit.get_values()):
                            result_row[ADRESS_LIVE_LABELS[z]] = cell
#                            qr += cell + ' '
#                        print(qr)
                        n = [char for char in result_row[ADRESS_LIVE_LABELS[0]] if char in string.digits]
                        if len(n) != 6:
                            result_row[ADRESS_LIVE_LABELS[0]] = '111111'

# Регион регистрации из номера
                    elif label0 == MANIPULATE_LABELS[21]:
                        addr = field2addr(REGIONS[intl(row_item)])
                        lab = [ADRESS_REG_LABELS[1], ADRESS_REG_LABELS[2]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]

# Регион проживания из номера
                    elif label0 == MANIPULATE_LABELS[23]:
                        addr = field2addr(REGIONS[intl(row_item)])
                        lab = [ADRESS_LIVE_LABELS[1], ADRESS_LIVE_LABELS[2]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
# Серия и номер паспорта из поля
                    elif label0 == MANIPULATE_LABELS[25]:
                        addr = field2sernum(row_item)
                        lab = [PASSPORT_DATA_LABELS[0],PASSPORT_DATA_LABELS[1]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]

                elif label0 == '-------------------------':
                    continue
                elif label0 in SNILS_LABEL:
                    if GENERATE_SNILS:
                        dbconfig = read_config(filename='move.ini', section='mysql')
                        dbconn = MySQLConnection(**dbconfig)
                        count_snils = 1
                        cached_snils = 0
                        while count_snils > 0:
                            self.start_snils -= 1
                            checksum_snils = self.checksum(self.start_snils)
                            for i in range(0, 99):
                                if i != checksum_snils:
                                    full_snils = self.start_snils * 100 + i
                                    dbcursor = dbconn.cursor()
                                    dbcursor.execute('SELECT `number` FROM clients WHERE `number` = %s', (full_snils,))
                                    dbchk = dbcursor.fetchall()
                                    if len(dbchk) == 0:
                                        cached_snils = full_snils
                                        count_snils -= 1
                        dbconn.close()
                        result_row[label0] = normalize_snils(cached_snils)
                        self.ws_comp.append([row_item, cached_snils])
                    else:
                        result_row[label0] = normalize_snils(row_item)
                elif label0 in PLACE_BIRTH_LABELS:
                    result_row[label0] = row_item
                elif label0 in PASSPORT_DATA_LABELS:
                    if PASSPORT_DATA_LABELS.index(label0) == 0:
                        result_row[label0] = normalize_seria(row_item)
                    elif PASSPORT_DATA_LABELS.index(label0) == 1:
                        result_row[label0] = normalize_index(row_item)
                    elif PASSPORT_DATA_LABELS.index(label0) == 2:
                        result_row[label0] = normalize_date(row_item)
                    elif PASSPORT_DATA_LABELS.index(label0) == 3:
                        result_row[label0] = normalize_text(row_item)
                    elif PASSPORT_DATA_LABELS.index(label0) == 4:
                        result_row[label0] = normalize_index(row_item)

                elif label0 in PHONES_LABELS:
                    if PHONES_LABELS.index(label0) == 0:
                        phone.tel_mob = row_item
                    elif PHONES_LABELS.index(label0) == 1:
                        phone.tel_rod = row_item
                    elif PHONES_LABELS.index(label0) == 2:
                        phone.tel_dom = row_item
                elif label0 in DATE_BIRTH_LABEL:
                    result_row[label0] = normalize_date(row_item)
                elif label0 in GENDER_LABEL:
                    result_row[label0] = normalize_gender(row_item)
                elif label0 == ADRESS_REG_LABELS[0] or label0 == ADRESS_LIVE_LABELS[0]:
                    result_row[label0] = normalize_index(row_item)
                elif label0 in ADRESS_REG_LABELS[11]:
                    result_row[label0] = normalize_home(row_item)
                elif label0 in ADRESS_LIVE_LABELS[11]:
                    result_row[label0] = normalize_home(row_item)
                elif label0 in ADRESS_REG_LABELS:
                    result_row[label0] = row_item
                elif label0 in ADRESS_LIVE_LABELS:
                    result_row[label0] = row_item
                elif label0 in TECH_LABELS:
                    if label0 == TECH_LABELS[0]:
                        result_row[label0] = self.agent_id
                    elif label0 == TECH_LABELS[1]:
                        result_row[label0] = self.signer_id
                    elif label0 == TECH_LABELS[2]:
                        result_row[label0] = PREDSTRAH_ID
                else:
                    result_row[label0] = normalize_text(row_item)

#            for num, z in enumerate(passport.get_values()):
#                result_row[PASSPORT_DATA_LABELS[num]] = z
            for num, z in enumerate(phone.get_values()):
                result_row[PHONES_LABELS[num]] = z

            LABELS = [SNILS_LABEL, FIO_LABELS, FIO_BIRTH_LABELS, FIO_SNILS_LABELS, GENDER_LABEL, DATE_BIRTH_LABEL,
                      PLACE_BIRTH_LABELS, PASSPORT_DATA_LABELS, ADRESS_REG_LABELS, ADRESS_LIVE_LABELS,
                      PHONES_LABELS, TECH_LABELS]
            mass = []
            for label_group in LABELS:
                for label in label_group:
                    mass.append(label)
            yum = True
            yum_phone0 = -1
            yum_phone1 = -1
            yum_phone2 = -1
            for num, cell in enumerate(mass):
                if cell in result_row:
                    mass[num] = result_row[cell]                # заполняем mass, чтобы его добавить как строку в xlsx
                    if cell == PHONES_LABELS[0]:
                        if mass[num] == ERROR_VALUE:
                            mass[num] = ''
                        yum_phone0 = num
                    elif cell == PHONES_LABELS[1]:
                        if mass[num] == ERROR_VALUE:
                            mass[num] = ''
                        yum_phone1 = num
                    elif cell == PHONES_LABELS[2]:
                        if mass[num] == ERROR_VALUE:
                            mass[num] = ''
                        yum_phone2 = num
                    elif mass[num] == ERROR_VALUE:
                        yum = False
                else:
                    mass[num] = ''                # заполняем mass, чтобы его добавить как строку в xlsx

#            yam = 0
#            if len(phone.tel_mob) > 0:
#                yam = int(phone.tel_mob)
#            if len(phone.tel_rod) > 0:
#                yam = yam + int(phone.tel_rod)
#            if len(phone.tel_dom) > 0:
#                yam = yam + int(phone.tel_dom)

            if mass[yum_phone0] == mass[yum_phone1] and mass[yum_phone0] !='':      # стираем дублирующиеся телефоны
                mass[yum_phone1] = ''
            if mass[yum_phone1] == mass[yum_phone2] and mass[yum_phone1] !='':
                mass[yum_phone2] = ''
            if mass[yum_phone0] == mass[yum_phone2] and mass[yum_phone0] !='':
                mass[yum_phone2] = ''

            if mass[yum_phone0] == '' and mass[yum_phone1] == '' and mass[yum_phone2] == '':
                yum = False                                                  # если нет ни одного телефона - ошибка


            if yum and err_from_log.get(num_row + 1) == None:
                ws.append(mass)
#                print(num_row, result_row['ФИО.Фамилия'], result_row['ФИО.Имя'], result_row['ФИО.Отчество'])
            else:
                mass.append(num_row + 1)
                mass.append(err_from_log.get(num_row + 1))
                ws_err.append(mass)
#                print(num_row, result_row['ФИО.Фамилия'], result_row['ФИО.Имя'], result_row['ФИО.Отчество'])
            if num_row % 10000 == 0:                # режем по 10000
                f = self.fname.replace(self.fname.split('/')[-1], '{0:02d}'.format(file_number) + '_' +
                                       self.fname.split('/')[-1])
                wb.save(f)
                wb = Workbook(write_only=True)
                ws = wb.create_sheet('Лист1')
                ws.append(HEAD_RESULT_EXCEL_FILE)  # добавляем первую строку xlsx
                file_number += 1

        f = self.fname.replace(self.fname.split('/')[-1], '{0:02d}'.format(file_number) + '_'+ self.fname.split('/')[-1])
        wb.save(f)
        f = self.fname.replace(self.fname.split('/')[-1], 'err'.format(file_number) + self.fname.split('/')[-1])
        wb_err.save(f)
        if use_log:
            log_file.close()
        if GENERATE_SNILS:
            self.wb_comp.save(self.fname.replace(self.fname.split('/')[-1], 'com'.format(file_number)
                                                 + self.fname.split('/')[-1]))


class BaseClass:

    def __setattr__(self, name, value):
        if isinstance(value, (int, str)):
            self.__dict__[name] = str(value).strip()
        else:
            self.__dict__[name] = value


def normalize(*args):
    result = []
    for arg in args:
        if arg == NULL_VALUE:
            result.append(NEW_NULL_VALUE)
        else:
            result.append(str(arg).strip())
    return result


def normalize_snils(snils):
    snils = str(snils).strip()
    snilsX = ''
    if snils != NULL_VALUE and snils != '' and isinstance(snils, str):
        try:
            for cc in snils:
                if cc in string.digits:
                    snilsX = snilsX+cc
            if len(snilsX) < LEN_SNILS:
                snilsX = '0' * (LEN_SNILS - len(snilsX)) + snilsX
            elif len(snilsX) == LEN_SNILS:
                pass
            else:
                return ERROR_VALUE
            return snilsX
        except TypeError:
            return ERROR_VALUE
    else:
        return ERROR_VALUE


def field2fio(field):
    if len(field) > 0 and field != NULL_VALUE:
        first_name, second_name, third_name = '', '', ''
        field = field.strip().replace('  ',' ').replace('  ',' ').split(' ')
        for i, word in enumerate(field):
            if i == 0:
                first_name = field[i]
            elif i == 1:
                second_name = field[i]
            else:
                third_name += field[i] + ' '
        if len(third_name) > 0:
            third_name = third_name[:-1]
        return first_name, second_name, third_name
    else:
        return NEW_NULL_VALUE_FOR_ALL_TEXT

def field2addr(field):
    addr_name, addr_type = '', ''
    if len(field) > 0 and field != NULL_VALUE:
        new_field = ''
        for i, ch in enumerate(field):          # убираем точки и запятые
            if ch == '.' or ch == ',':
                ch = ''
            new_field = new_field + ch
        field = new_field.strip().split(' ')
        TYPES = [REG_TYPES, DISTRICT_TYPES, CITY_TYPES, NP_TYPES, STREET_TYPES]
        for i, word in enumerate(field):
            addr_type_vrem = ''
            for label_group in TYPES:
                for label in label_group:
                    if word.lower() == label.lower():
                        addr_type_vrem = label
            if addr_type_vrem == '':
                addr_name = addr_name + ' ' + word
            else:
                addr_type = addr_type_vrem
    return addr_name, addr_type

#class Gender(BaseClass):
#    def __init__(self, third_name='', gender_field_exists=False, gender=''):
#        self.female_gender_value = female_gender_value
#        self.male_gender_value = male_gender_value
#        self.third_name = str(third_name).strip()
#        self.gender_field_exists = gender_field_exists
#        self.gender = gender.strip()

#    def gender_from_fio(self):
#        if self.third_name == '':
#            return ERROR_VALUE
#        third_name = self.third_name
#        third_name = third_name.split(' ')
#        if len(third_name) == 1:
#            if ''.join(third_name[0][-3:]).lower() == 'вна':
#                gender = '0'
#            elif ''.join(third_name[0][-3:]).lower() == 'вич':
#                gender = '1'
#            else:
#                gender = ERROR_VALUE
#        elif len(third_name) == 2:
#            if third_name[-1].lower() in EAST_GENDER[0]:  # женщина
#                gender = '0'
#            elif third_name[-1].lower() in EAST_GENDER[1]:  # мужчина
#                gender = '1'
#            else:
#                gender = ERROR_VALUE
#        else:
#            gender = ERROR_VALUE
#        return gender

#    def set_gender_value(self, male_value, female_value):
#        self.female_gender_value = female_value.lower()
#        self.male_gender_value = male_value.lower()

#    def get_gender_value(self):
#        return self.female_gender_value, self.male_gender_value

#    def normalize_gender(self):
#        gender = self.gender
#        gender = gender.lower()
#        if gender == self.female_gender_value:
#            return '0'
#        elif gender == self.male_gender_value:
#            return '1'
#        else:
#            return self.gender_from_fio()

#    def get_value(self):
#        if self.gender_field_exists:
#            return self.normalize_gender()
#        else:
#            return self.gender_from_fio()

def normalize_gender(gender):
    gender = str(gender).strip()
    if gender =='':
        return NEW_NULL_VALUE_FOR_GENDER
    elif len(gender) > 1 and (gender.strip().upper()[:gender_length] != female_gender_value and
                              gender.strip().upper()[:gender_length] != male_gender_value):
        return NEW_NULL_VALUE_FOR_GENDER
    else:
        if gender.strip().upper()[:gender_length] == female_gender_value:
            return '1'
        else:
            return '0'

def normalize_text(tx):
    tx = str(tx).strip()
    if len(tx) <= 1:
        return NEW_NULL_VALUE_FOR_ALL_TEXT
    else:
        return tx

def normalize_date(date):
    date = str(date)
    try:
        result = re.findall(r'\b(\d{4}|\d{2})[\.:-](\d{2})[\.:-](\d{4}|\d{2})\b', date)
        if len(result) > 0:
            if result[0] == NULL_VALUE:
                return NEW_NULL_VALUE_FOR_DATE
            if len(result[0][0]) == 4:
                result[0] = result[0][::-1]
            elif len(result[0][2]) == 2:
                if result[0][2] < 20:
                    result[0][2] = '20' + result[0][2]
                elif result[0][2] > 20:
                    result[0][2] = '19' + result[0][2]
            return '.'.join(result[0])
        else:
            return NEW_NULL_VALUE_FOR_DATE
    except Exception as ee:
        return  NEW_NULL_VALUE_FOR_DATE


# print(normalize_date('01.09.2003'))

# normalize место рождения


class Passport(BaseClass):
    def __init__(self, seriya='', nomer='', date='', who='', cod=''):
        self.seriya = str(seriya).strip()
        self.nomer = str(nomer).strip()
        self.date = normalize_date(date)
        self.who = normalize_text(who)
        self.cod = str(cod).strip()

    def __setattr__(self, name, value):
        if name == 'date':
            self.__dict__[name] = normalize_date(value)
        else:
            if isinstance(value, (int, str)):
                self.__dict__[name] = str(value)
            else:
                self.__dict__[name] = value

    def normalize_seriya(self):
        if self.seriya != NULL_VALUE and self.seriya != '' and isinstance(self.seriya, str):
            try:
                self.seriya = ''.join([char for char in self.seriya if char in string.digits])
                if len(self.seriya) == 3:
                    self.seriya = '0' + self.seriya
                elif len(self.seriya) == 4:
                    pass
                else:
                    return NEW_NULL_VALUE_FOR_SERIYA_PASSPORTA
#                   return ERROR_VALUE
#                self.seriya = self.seriya[:2] + ' ' + self.seriya[2:]         # Тот самый пробел между ## ##
                return self.seriya
            except TypeError:
                return NEW_NULL_VALUE_FOR_SERIYA_PASSPORTA
        else:
            return NEW_NULL_VALUE_FOR_SERIYA_PASSPORTA

    def normalize_nomer(self):
        if self.nomer != NULL_VALUE and self.nomer != '' and isinstance(self.seriya, str):
            try:
                self.nomer = ''.join([char for char in self.nomer if char in string.digits])
                if len(self.nomer) < LEN_PASSPORT_NOMER and len(self.nomer) > 0 :
                    self.nomer = '0' * (LEN_PASSPORT_NOMER - len(self.nomer)) + self.nomer
                elif len(self.nomer) == LEN_PASSPORT_NOMER:
                    pass
                else:
                    return NEW_NULL_VALUE_FOR_NOMER_PASSPORTA
#                    return ERROR_VALUE
                return self.nomer
            except TypeError:
                return NEW_NULL_VALUE_FOR_NOMER_PASSPORTA
        else:
            return NEW_NULL_VALUE_FOR_NOMER_PASSPORTA


    def normalize_who(self):
        if self.who == NULL_VALUE:
            self.who = NEW_NULL_VALUE
        return self.who

    def normalize_cod(self):
        if self.cod != NULL_VALUE and self.cod != '':
            self.cod = ''.join([char for char in self.cod if char in string.digits])
            if len(self.cod) < LEN_PASSPORT_COD and len(self.cod) > 0:
                self.cod = '0' * (LEN_PASSPORT_COD - len(self.cod)) + self.cod
            elif len(self.cod) == LEN_PASSPORT_COD:
                pass
            else:
                return NEW_NULL_VALUE_FOR_COD_PASSPORTA
#                return ERROR_VALUE
            self.cod = self.cod[:3] + '-' + self.cod[3:]
            return self.cod
        else:
            return NEW_NULL_VALUE_FOR_COD_PASSPORTA

    def get_values(self):
        return self.normalize_seriya(), self.normalize_nomer(), self.date, self.normalize_who(), self.normalize_cod()


def normalize_index(index):
    index = str(index).strip()
    if index != NULL_VALUE and index != '' and isinstance(index, str):
        try:
            index = ''.join([char for char in index if char in string.digits])
            if len(index) < LEN_INDEX_NOMER:
                index = '0' * (LEN_INDEX_NOMER - len(index)) + index
            elif len(index) == LEN_INDEX_NOMER:
                pass
            else:
                return NEW_NULL_VALUE_FOR_INDEX
#                return ERROR_VALUE
            return index
        except TypeError:
            return NEW_NULL_VALUE_FOR_INDEX
    else:
        return NEW_NULL_VALUE_FOR_INDEX

def normalize_home(tx):
        tx = str(tx).strip()
        numbers = True
        for i in range(len(tx)):
            if tx[i] not in string.digits:
                numbers = False
        if len(tx) < 1:
            return tx
        elif len(tx) > 10:
            return NEW_NULL_VALUE_FOR_HOME
        elif numbers:
            if int(tx) > 1500:
                return NEW_NULL_VALUE_FOR_HOME
            else:
                return tx
        else:
            return tx


class FullAdress(BaseClass):
    def __init__(self, field=''):
        self.field = str(field)
        self.full_adress = []
        self.FULL_ADRESS_DICT = {}
        for label in FULL_ADRESS_LABELS:
            self.FULL_ADRESS_DICT[label] = ''
        self.iter_types = [DISTRICT_TYPES, CITY_TYPES, NP_TYPES, STREET_TYPES, HOUSE_CUT_NAME, CORPUS_CUT_NAME, APARTMENT_CUT_NAME]

    def normalize_adress(self):
        if len(self.field) != 0 and self.field != NULL_VALUE:
            self.field = self.field.lower()
            values = self.field.split(SPLIT_FIELD)
            for i, word in enumerate(values):
                n = []
                word = word.strip()
                if i == 0:
                    n = [char for char in word if char in string.digits]
                    if len(n) != 6:
                        return NEW_NULL_VALUE_FOR_INDEX
                    self.FULL_ADRESS_DICT[FULL_ADRESS_LABELS[0]] = ''.join(n)
                    continue
                elif i == 1:
                    self.FULL_ADRESS_DICT[FULL_ADRESS_LABELS[1]] = ' '.join(word.split(' ')[:-1])
                    self.FULL_ADRESS_DICT[FULL_ADRESS_LABELS[2]] = word.split(' ')[-1]
                    continue
                else:
                    for j, types in enumerate(self.iter_types):
                        if j < 4:
                            if word.split(' ')[-1] in types:
                                self.FULL_ADRESS_DICT[FULL_ADRESS_LABELS[3 + 2 * j]] = ' '.join(
                                    word.split(' ')[:-1])
                                self.FULL_ADRESS_DICT[FULL_ADRESS_LABELS[3 + 2 * j + 1]] = word.split(' ')[-1]
                        elif j >= 4:
                            for type in types:
                                if word.find(type) != (-1):
                                    word = word.replace(type, '').replace('.', '')
                                    self.FULL_ADRESS_DICT[FULL_ADRESS_LABELS[11 + j - 4]] = word
            return self.FULL_ADRESS_DICT
        else:
            if self.field == NULL_VALUE:
                return NEW_NULL_VALUE
            else:
                return NEW_NULL_VALUE
                #                return ERROR_VALUE

    def create_output_list(self):                   # Создаем список вывода
        if self.field != '':
            FULL_ADRESS_DICT = self.normalize_adress()
        for label in FULL_ADRESS_LABELS:
            self.full_adress.append(self.FULL_ADRESS_DICT[label].upper())
        return self.full_adress

    def get_values(self):  # Когда адрес 414000, г. Астрахань, ул. Такая, д. Т...
        output_list = []
        for elem in self.create_output_list():
            output_list.append(elem.strip())
        return output_list

    a = """
    def get_values(self):                           # Когда все поля по раздельности...
        output_list = []
        if len(self.field) != 0 and self.field != NULL_VALUE:
            self.field = self.field.lower()
            values = self.field.split(SPLIT_FIELD)
            for i, nn in enumerate(ORDER_FIELD):
                if nn < len(values):
                    output_list.append(values[nn])
            return output_list
        else:
            if self.field == NULL_VALUE:
                return NEW_NULL_VALUE
            else:
                return NEW_NULL_VALUE


    # def __call__(self, *args, **kwargs):
    #     return self.create_output_list()


# f = FullAdress('123592, Москва г, строгинский бульвар, д. 26, корпус 2, кв. 425')
# print(f.get_values())
    """

class Phone(BaseClass):
    def __init__(self, tel_mob='', tel_rod='', tel_dom=''):
        self.tel_mob = str(tel_mob).strip()
        self.tel_rod = str(tel_rod).strip()
        self.tel_dom = str(tel_dom).strip()

    def normalize_tel_number(self, tel):
        tel = tel.strip()
        if tel == '' or tel == NULL_VALUE:
            return ERROR_VALUE
        tel = str(tel).strip()
        tel = ''.join([char for char in tel if char in string.digits])
        if len(tel) == 11:
            if tel[0] in ['7', '8', '9']:
                tel = '7' + tel[1:]
            else:
                return ERROR_VALUE
        elif len(tel) == 10:
            tel = '7' + tel
        else:
            return ERROR_VALUE
        return tel

    def poryadoc(self, *tels):
        tels = sorted(tels)
#        tels.reverse()
        return list(tels)

    def get_values(self):
        self.tel_mob = self.normalize_tel_number(self.tel_mob)
        self.tel_rod = self.normalize_tel_number(self.tel_rod)
        self.tel_dom = self.normalize_tel_number(self.tel_dom)
#        if self.tel_rod == self.tel_mob:
#            self.tel_rod = ''
#        self.tel_dom = self.normalize_tel_number(self.tel_dom)
#        if self.tel_dom == self.tel_mob or self.tel_dom == self.tel_rod:
#            self.tel_rod = ''
        return self.poryadoc(self.tel_mob, self.tel_rod, self.tel_dom)

# p = Phone()
# p.tel_dom = 89040964007
# p.tel_rod = 89257349331
# p.tel_mob = 'dd'
# print(p.get_values())

def intl(a):               # белиберду в цифры или 0
    try:
        if a != None:
            a = str(a).strip()
            if  a != '':
                a = ''.join([char for char in a if char in string.digits])
                if len(a) > 0:
                    return int(a)
                else:
                    return 0
        return 0
    except TypeError:
        return 0

def field2sernum(field):
    if len(field) > 0 and field != NULL_VALUE and l(field) > 999999 and len('{0:010d}'.format(l(field))) == 10:
        new_field = '{0:010d}'.format(l(field))
        seria = new_field[:4]
        number = new_field[4:]
        return seria, number
    else:
        return '1111', '111111'

def normalize_seria(field):
    if len(field) > 0 and field != NULL_VALUE and l(field) > 0 and len('{0:04d}'.format(l(field))) == 4:
        return '{0:04d}'.format(l(field))
    else:
        return '1111'


