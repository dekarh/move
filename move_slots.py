# -*- coding: utf-8 -*-

from subprocess import Popen, PIPE
import os
import sys
import string
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

import NormalizeFields as norm
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
#                     , "-------------------------"
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



IN_IDS = ['ID','ИД_Клиента','client_id']
IN_SNILS = ['СНИЛС', 'Страховой_номер', 'number']
IN_NAMES = ['ID', 'СНИЛС', 'Страховой_номер', 'number', 'Фамилия', 'Имя', 'Отчество']

DIR4MOVE = '/home/da3/Move/'
DIR4IMPORT = '/home/da3/CheckLoad/'
DIR4CFGIMPORT = '/home/da3/CheckLoad/cfg/'

class MainWindowSlots(Ui_Form):   # Определяем функции, которые будем вызывать в слотах

    def setupUi(self, form):
        Ui_Form.setupUi(self,form)
        self.tableWidget.setColumnCount(2)
        self.tableWidget.setHorizontalHeaderLabels(('Результат', 'Исходник'))
        for j in range(self.tableWidget.columnCount()):
            self.tableWidget.setColumnWidth(j, 220)
        self.tableWidget.setRowCount(0)
        self.dbconfig = read_config(filename='move.ini', section='mysql')
        self.MoveImport = True
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
        self.file_names = {}
        self.file_name = ''
        self.tab_names = {}
        self.table = []
        self.twParsingResult.hide()
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
        if self.MoveImport:
            self.frImport.hide()
            self.frMove.show()
        else:
            self.frImport.show()
            self.frMove.hide()

        dbconn = MySQLConnection(**self.dbconfig)
        cursor = dbconn.cursor()
        if self.leAgent.text().strip():
            sql = "SELECT CONCAT_WS(' ', code, '-', user_surname, user_name, user_lastname, '-', position_id), code " \
                  "FROM saturn_crm.offices_staff WHERE user_fired = 0 AND " \
                  "CONCAT_WS(' ', code, '-', user_surname, user_name, user_lastname, user_lastname) LIKE %s"
            cursor.execute(sql, ('%' + self.leAgent.text() + '%',))
        else:
            sql = "SELECT CONCAT_WS(' ', code, '-', user_surname, user_name, user_lastname, '-', position_id), code " \
                  "FROM saturn_crm.offices_staff WHERE user_fired = 0"
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
            sql = "SELECT CONCAT_WS(' ', id, '-', name), id FROM saturn_crm.subdomains " \
                  "WHERE CONCAT_WS(' ', id, '-', name) LIKE %s AND id IN (2,6,8,11,12,13)"
            cursor.execute(sql, ('%' + self.leFond.text() + '%',))
        else:
            sql = "SELECT CONCAT_WS(' ', id, '-', name), id FROM saturn_crm.subdomains WHERE id IN (2,6,8,11,12,13)"
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
                      "FROM saturn_crm.signers " \
                      "WHERE CONCAT_WS(' ', id, '-', signer_surname, signer_name, signer_lastname) LIKE %s " \
                      "AND subdomain_id = %s"
                cursor.execute(sql, (self.fond_ids[self.cmbFond.currentIndex()],))
            else:
                sql = "SELECT CONCAT_WS(' ', id, '-', signer_surname, signer_name, signer_lastname), id " \
                      "FROM saturn_crm.signers " \
                      "WHERE CONCAT_WS(' ', id, '-', signer_surname, signer_name, signer_lastname) LIKE %s " \
                      "AND subdomain_id IN (2,6,8,11,12,13) AND subdomain_id IN (" + self.fonds_str + ")"
                cursor.execute(sql, ('%' + self.leSigner.text() + '%',))
        else:
            if self.fond_touched:
                sql = "SELECT CONCAT_WS(' ', id, '-', signer_surname, signer_name, signer_lastname), id " \
                      "FROM saturn_crm.signers WHERE subdomain_id = %s"
                cursor.execute(sql, (self.fond_ids[self.cmbFond.currentIndex()],))
            else:
                sql = "SELECT CONCAT_WS(' ', id, '-', signer_surname, signer_name, signer_lastname), id " \
                      "FROM saturn_crm.signers WHERE subdomain_id IN (" + self.fonds_str + ") " \
                      "AND subdomain_id IN (2,6,8,11,12,13)"
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
                    self.file_loaded = False
                    self.file_touched = False
                    self.file_name = ''
                else:
                    self.cmbFile.setCurrentIndex(self.file_names[file_name])
            except ValueError:
                self.cfg_file_loaded = False
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
            self.cfg_file_loaded = False
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
        sql = "SELECT cl.client_id FROM saturn_crm.clients AS cl WHERE cl.client_id IN (" + all_clients_ids + \
              ") GROUP BY cl.client_id HAVING COUNT(cl.client_id) > 1 ORDER BY cl.client_id DESC"
        dbconn = MySQLConnection(**self.dbconfig)
        cursor = dbconn.cursor()
        cursor.execute(sql)
        rows = cursor.fetchall()
        exit_because_doubles = False
        if len(rows) > 0:
            exit_because_doubles = True
            ws_log.append([datetime.now().strftime("%H:%M:%S"), ' Дубли в saturn_crm.clients'])
            ws_clients = wb_log.create_sheet('Дубли в saturn_crm.clients')
            for row in rows:
                ws_clients.append(row[0])
        else:
            ws_log.append([datetime.now().strftime("%H:%M:%S"), ' В saturn_crm.clients нет дублей'])
                                                                # Проверка на дубли contracts
        sql = "SELECT co.client_id FROM saturn_crm.contracts AS co WHERE co.client_id IN (" + all_clients_ids + \
              ") GROUP BY co.client_id HAVING COUNT(co.client_id) > 1 ORDER BY co.client_id DESC"
        dbconn = MySQLConnection(**self.dbconfig)
        cursor = dbconn.cursor()
        cursor.execute(sql)
        rows = cursor.fetchall()
        if len(rows) > 0:
            exit_because_doubles = True
            ws_log.append([datetime.now().strftime("%H:%M:%S"), ' Дубли в saturn_crm.contracts'])
            ws_contracts = wb_log.create_sheet('Дубли в saturn_crm.contracts')
            for row in rows:
                ws_contracts.append(row[0])
        else:
            ws_log.append([datetime.now().strftime("%H:%M:%S"), ' В saturn_crm.contracts нет дублей'])

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
        sql = "SELECT cl.*, co.* FROM saturn_crm.clients AS cl LEFT JOIN saturn_crm.contracts AS co " \
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
        if self.chbArhiv.isChecked():
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'Поставить флаг "Архивный"', 'выбрано'])
        else:
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'Поставить флаг "Архивный"', 'не выбрано'])

        ws_log.append([datetime.now().strftime("%H:%M:%S"), ' Формируем запросы:'])
        sql_cl = 'UPDATE saturn_crm.clients AS cl SET'
        sql_co = 'UPDATE saturn_crm.contracts AS co SET'
        if self.leAgent.isEnabled():
            sql_cl += ' cl.inserted_user_code = %s'
            sql_co += ' co.inserted_code = %s, co.agent_code = %s'
        if self.leFond.isEnabled():
            if sql_cl[len(sql_cl) - 2:] == '%s':
                sql_cl += ','
            sql_cl += ' cl.subdomain_id = %s'
        if self.chbArhiv.isChecked():
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
            if self.chbArhiv.isChecked():
                tuple_client += (1,)
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
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'saturn_crm.clients - нет запроса'])
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'saturn_crm.clients - нет данных'])
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'saturn_crm.clients - запрос не исполнен'])
        if self.leSQLco.text():
            ws_log.append([datetime.now().strftime("%H:%M:%S"), self.leSQLco.text()])
            ws_log.append([datetime.now().strftime("%H:%M:%S")] + list(tuples_contracts[0]))
            dbconn = MySQLConnection(**self.dbconfig)
            cursor = dbconn.cursor()
            cursor.executemany(self.leSQLco.text(),tuples_contracts)
            dbconn.commit()
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'update отработал'])
        else:
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'saturn_crm.contracts - нет запроса'])
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'saturn_crm.contracts - нет данных'])
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'saturn_crm.contracts - запрос не исполнен'])

        wb_log.save(log_name)
        q=0

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
                      "FROM saturn_crm.signers " \
                      "WHERE CONCAT_WS(' ', id, '-', signer_surname, signer_name, signer_lastname) LIKE %s " \
                      "AND subdomain_id = %s"
                cursor.execute(sql, (self.fond_ids[self.cmbFond.currentIndex()],))
            else:
                sql = "SELECT CONCAT_WS(' ', id, '-', signer_surname, signer_name, signer_lastname), id " \
                      "FROM saturn_crm.signers " \
                      "WHERE CONCAT_WS(' ', id, '-', signer_surname, signer_name, signer_lastname) LIKE %s " \
                      "AND subdomain_id IN (2,6,8,11,12,13) AND subdomain_id IN (" + self.fonds_str + ")"
                cursor.execute(sql, ('%' + self.leSigner.text() + '%',))
        else:
            if self.fond_touched:
                sql = "SELECT CONCAT_WS(' ', id, '-', signer_surname, signer_name, signer_lastname), id " \
                      "FROM saturn_crm.signers WHERE subdomain_id = %s"
                cursor.execute(sql, (self.fond_ids[self.cmbFond.currentIndex()],))
            else:
                sql = "SELECT CONCAT_WS(' ', id, '-', signer_surname, signer_name, signer_lastname), id " \
                      "FROM saturn_crm.signers WHERE subdomain_id IN (" + self.fonds_str + ") " \
                      "AND subdomain_id IN (2,6,8,11,12,13)"
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
                      "FROM saturn_crm.signers " \
                      "WHERE CONCAT_WS(' ', id, '-', signer_surname, signer_name, signer_lastname) LIKE %s " \
                      "AND subdomain_id = %s"
                cursor.execute(sql, (self.fond_ids[self.cmbFond.currentIndex()],))
            else:
                sql = "SELECT CONCAT_WS(' ', id, '-', signer_surname, signer_name, signer_lastname), id " \
                      "FROM saturn_crm.signers " \
                      "WHERE CONCAT_WS(' ', id, '-', signer_surname, signer_name, signer_lastname) LIKE %s " \
                      "AND subdomain_id IN (2,6,8,11,12,13) AND subdomain_id IN (" + self.fonds_str + ")"
                cursor.execute(sql, ('%' + self.leSigner.text() + '%',))
        else:
            if self.fond_touched:
                sql = "SELECT CONCAT_WS(' ', id, '-', signer_surname, signer_name, signer_lastname), id " \
                      "FROM saturn_crm.signers WHERE subdomain_id = %s"
                cursor.execute(sql, (self.fond_ids[self.cmbFond.currentIndex()],))
            else:
                sql = "SELECT CONCAT_WS(' ', id, '-', signer_surname, signer_name, signer_lastname), id " \
                      "FROM saturn_crm.signers WHERE subdomain_id IN (" + self.fonds_str + ") " \
                      "AND subdomain_id IN (2,6,8,11,12,13)"
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
        if self.MoveImport:
            self.load4move()
        else:
            self.load4import()
        return

    def click_clbMove(self):
        self.MoveImport = False
        self.frImport.show()
        self.twParsingResult.show()
        self.frMove.hide()

    def click_clbImport(self):
        self.MoveImport = True
        self.frImport.hide()
        self.twParsingResult.hide()
        self.frMove.show()

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
                    if cell.value in IN_IDS:
                        keys[IN_IDS[0]] = k
                if len(keys) > 0:
                    for k, cell in enumerate(row):
                        for n, name in enumerate(IN_NAMES):
                            if n == 0:
                                continue
                            if cell.value != None:
                                if str(cell.value).strip() != '':
                                    last_cell = k
                                    if cell.value == name:
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

#!!!!!!!!!!!!!!!!!!!!! IMPORT !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

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
                    if cell.value in IN_SNILS:
                        keys[IN_SNILS[0]] = k
                if len(keys) > 0:
                    for k, cell in enumerate(row):
                        for n, name in enumerate(IN_NAMES):
                            if n == 0:
                                continue
                            if cell.value != None:
                                if str(cell.value).strip() != '':
                                    last_cell = k
                                    if cell.value == name:
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
            for row in doubles_in_input:
                ws_input_doubles.append(row)

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
        i10l = 0
        i10 = 0
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
            passport = norm.Passport()
            phone = norm.Phone()

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
                        FIO = norm.field2fio(row_item)
                        if label0 == MANIPULATE_LABELS[1]:
                            lab = FIO_LABELS
                        elif label0 == MANIPULATE_LABELS[3]:
                            lab = FIO_BIRTH_LABELS
                        if row_item == '':
                            for j in range(len(lab)):
                                result_row[lab[j]] = norm.NEW_NULL_VALUE_FOR_ALL_TEXT
                        else:
                            for j in range(len(FIO)):
                                result_row[lab[j]] = FIO[j]
                        continue

                    # ------------------------------------------------------- Убрал класс Gender --------------------------------------
                    #                    elif label0 == "Пол_получить_из_ФИО":
                    #                        gender = norm.Gender(row_item)
                    #                        result_row[GENDER_LABEL[0]] = gender.get_value()

                    #                    elif label0 == "Пол_подставить_свои_значения":
                    #                        gender = norm.Gender(FIO[2], gender_field_exists=True, gender=row_item) ## !!!!!!!!!!!!!!
                    #                        result_row[GENDER_LABEL[0]] = gender.get_value()
                    # ------------------------------------------------------- Убрал класс Gender --------------------------------------
                    # Регистрация -> Регион
                    elif label0 == MANIPULATE_LABELS[5]:
                        addr = norm.field2addr(row_item)
                        lab = [ADRESS_REG_LABELS[1], ADRESS_REG_LABELS[2]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
                    # Регистрация -> Район
                    elif label0 == MANIPULATE_LABELS[6]:
                        addr = norm.field2addr(row_item)
                        lab = [ADRESS_REG_LABELS[3], ADRESS_REG_LABELS[4]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
                    # Регистрация -> Город
                    elif label0 == MANIPULATE_LABELS[7]:
                        addr = norm.field2addr(row_item)
                        lab = [ADRESS_REG_LABELS[5], ADRESS_REG_LABELS[6]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
                    # Регистрация -> Населенный_пункт
                    elif label0 == MANIPULATE_LABELS[8]:
                        addr = norm.field2addr(row_item)
                        lab = [ADRESS_REG_LABELS[7], ADRESS_REG_LABELS[8]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
                    # Регистрация -> Улица
                    elif label0 == MANIPULATE_LABELS[9]:
                        addr = norm.field2addr(row_item)
                        lab = [ADRESS_REG_LABELS[9], ADRESS_REG_LABELS[10]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
                    # ADRESS_LIVE_LABELS
                    # Проживание -> Регион
                    elif label0 == MANIPULATE_LABELS[11]:
                        addr = norm.field2addr(row_item)
                        lab = [ADRESS_LIVE_LABELS[1], ADRESS_LIVE_LABELS[2]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
                    # Проживание -> Район
                    elif label0 == MANIPULATE_LABELS[12]:
                        addr = norm.field2addr(row_item)
                        lab = [ADRESS_LIVE_LABELS[3], ADRESS_LIVE_LABELS[4]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
                    # Проживание -> Город
                    elif label0 == MANIPULATE_LABELS[13]:
                        addr = norm.field2addr(row_item)
                        lab = [ADRESS_LIVE_LABELS[5], ADRESS_LIVE_LABELS[6]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
                    # Проживание -> Населенный_пункт
                    elif label0 == MANIPULATE_LABELS[14]:
                        addr = norm.field2addr(row_item)
                        lab = [ADRESS_LIVE_LABELS[7], ADRESS_LIVE_LABELS[8]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
                    # Проживание -> Улица
                    elif label0 == MANIPULATE_LABELS[15]:
                        addr = norm.field2addr(row_item)
                        lab = [ADRESS_LIVE_LABELS[9], ADRESS_LIVE_LABELS[10]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
                    # Адрес регистрации из_поля
                    elif label0 == MANIPULATE_LABELS[17]:
                        result_row[ADRESS_REG_LABELS[0]] = '111111'
                        adress_reg = norm.FullAdress(row_item)
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
                        adress_zhit = norm.FullAdress(row_item)
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
                        addr = norm.field2addr(norm.REGIONS[norm.intl(row_item)])
                        lab = [ADRESS_REG_LABELS[1], ADRESS_REG_LABELS[2]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]

                    # Регион проживания из номера
                    elif label0 == MANIPULATE_LABELS[23]:
                        addr = norm.field2addr(norm.REGIONS[norm.intl(row_item)])
                        lab = [ADRESS_LIVE_LABELS[1], ADRESS_LIVE_LABELS[2]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]

                elif label0 == '-------------------------':
                    continue
                elif label0 in SNILS_LABEL:
                    if norm.GENERATE_SNILS:
                        dbconfig = read_config(filename='NormXLS.ini', section='main_mysql')
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
                        result_row[label0] = norm.normalize_snils(cached_snils)
                        self.ws_comp.append([row_item, cached_snils])
                    else:
                        result_row[label0] = norm.normalize_snils(row_item)
                elif label0 in PLACE_BIRTH_LABELS:
                    result_row[label0] = row_item
                elif label0 in PASSPORT_DATA_LABELS:
                    if PASSPORT_DATA_LABELS.index(label0) == 0:
                        passport.seriya = row_item
                    elif PASSPORT_DATA_LABELS.index(label0) == 1:
                        passport.nomer = row_item
                    elif PASSPORT_DATA_LABELS.index(label0) == 2:
                        passport.date = row_item
                    elif PASSPORT_DATA_LABELS.index(label0) == 3:
                        passport.who = norm.normalize_text(row_item)
                    elif PASSPORT_DATA_LABELS.index(label0) == 4:
                        passport.cod = row_item

                elif label0 in PHONES_LABELS:
                    if PHONES_LABELS.index(label0) == 0:
                        phone.tel_mob = row_item
                    elif PHONES_LABELS.index(label0) == 1:
                        phone.tel_rod = row_item
                    elif PHONES_LABELS.index(label0) == 2:
                        phone.tel_dom = row_item
                elif label0 in DATE_BIRTH_LABEL:
                    result_row[label0] = norm.normalize_date(row_item)
                elif label0 in GENDER_LABEL:
                    result_row[label0] = norm.normalize_gender(row_item)
                elif label0 == ADRESS_REG_LABELS[0] or label0 == ADRESS_LIVE_LABELS[0]:
                    result_row[label0] = norm.normalize_index(row_item)
                elif label0 in ADRESS_REG_LABELS[11]:
                    result_row[label0] = norm.normalize_home(row_item)
                elif label0 in ADRESS_LIVE_LABELS[11]:
                    result_row[label0] = norm.normalize_home(row_item)
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
                        result_row[label0] = norm.PREDSTRAH_ID
                else:
                    result_row[label0] = norm.normalize_text(row_item)

            for num, z in enumerate(passport.get_values()):
                result_row[PASSPORT_DATA_LABELS[num]] = z
            for num, z in enumerate(phone.get_values()):
                result_row[PHONES_LABELS[num]] = z

            LABELS = [SNILS_LABEL, FIO_LABELS, FIO_BIRTH_LABELS, FIO_SNILS_LABELS, GENDER_LABEL, DATE_BIRTH_LABEL,
                      PLACE_BIRTH_LABELS, PASSPORT_DATA_LABELS, ADRESS_REG_LABELS, ADRESS_LIVE_LABELS,
                      PHONES_LABELS, TECH_LABELS]
            mass = []
            for l in LABELS:
                for ll in l:
                    mass.append(ll)
            yum = True
            yum_phone0 = -1
            yum_phone1 = -1
            yum_phone2 = -1
            for num, cell in enumerate(mass):
                if cell in result_row:
                    mass[num] = result_row[cell]  # заполняем mass, чтобы его добавить как строку в xlsx
                    if cell == PHONES_LABELS[0]:
                        if mass[num] == norm.ERROR_VALUE:
                            mass[num] = ''
                        yum_phone0 = num
                    elif cell == PHONES_LABELS[1]:
                        if mass[num] == norm.ERROR_VALUE:
                            mass[num] = ''
                        yum_phone1 = num
                    elif cell == PHONES_LABELS[2]:
                        if mass[num] == norm.ERROR_VALUE:
                            mass[num] = ''
                        yum_phone2 = num
                    elif mass[num] == norm.ERROR_VALUE:
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
                    self.twParsingResult.setItem(maxParsingResult, ind, QTableWidgetItem(cell))
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
        if norm.GENERATE_SNILS:
            dbconfig = read_config(filename='NormXLS.ini', section='main_mysql')
            dbconn = MySQLConnection(**dbconfig)
            dbcursor = dbconn.cursor()
            dbcursor.execute('SELECT min(`number`) FROM  saturn_crm.clients WHERE `number` > 99000000000 and subdomain_id = 2;')
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
        i10l = 0
        i10 = 0
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

        for num_row, row in enumerate(self.sheet.rows):
            self.progress_value.emit(num_row + 1)  # отрисовываем ProgresBar
            if num_row == 0:
                continue
            i10 = int(num_row / 10005)
#--------------------------------------- С этим if не добавляло первую строку ----------------------------------
#            if num_row == 0:
#                continue
#--------------------------------------- С этим if не добавляло первую строку ----------------------------------

            result_row = {}

            passport = norm.Passport()
            phone = norm.Phone()

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
                        FIO = norm.field2fio(row_item)
                        if label0 == MANIPULATE_LABELS[1]:
                            lab = FIO_LABELS
                        elif label0 == MANIPULATE_LABELS[3]:
                            lab = FIO_BIRTH_LABELS
                        if row_item == '':
                            for j in range(len(lab)):
                                result_row[lab[j]] = norm.NEW_NULL_VALUE_FOR_ALL_TEXT
                        else:
                            for j in range(len(FIO)):
                                result_row[lab[j]] = FIO[j]
                        continue

#------------------------------------------------------- Убрал класс Gender --------------------------------------
#                    elif label0 == "Пол_получить_из_ФИО":
#                        gender = norm.Gender(row_item)
#                        result_row[GENDER_LABEL[0]] = gender.get_value()

#                    elif label0 == "Пол_подставить_свои_значения":
#                        gender = norm.Gender(FIO[2], gender_field_exists=True, gender=row_item) ## !!!!!!!!!!!!!!
#                        result_row[GENDER_LABEL[0]] = gender.get_value()
#------------------------------------------------------- Убрал класс Gender --------------------------------------
# Регистрация -> Регион
                    elif label0 == MANIPULATE_LABELS[5]:
                        addr = norm.field2addr(row_item)
                        lab = [ADRESS_REG_LABELS[1], ADRESS_REG_LABELS[2]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
# Регистрация -> Район
                    elif label0 == MANIPULATE_LABELS[6]:
                        addr = norm.field2addr(row_item)
                        lab = [ADRESS_REG_LABELS[3], ADRESS_REG_LABELS[4]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
# Регистрация -> Город
                    elif label0 == MANIPULATE_LABELS[7]:
                        addr = norm.field2addr(row_item)
                        lab = [ADRESS_REG_LABELS[5], ADRESS_REG_LABELS[6]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
# Регистрация -> Населенный_пункт
                    elif label0 == MANIPULATE_LABELS[8]:
                        addr = norm.field2addr(row_item)
                        lab = [ADRESS_REG_LABELS[7], ADRESS_REG_LABELS[8]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
# Регистрация -> Улица
                    elif label0 == MANIPULATE_LABELS[9]:
                        addr = norm.field2addr(row_item)
                        lab = [ADRESS_REG_LABELS[9], ADRESS_REG_LABELS[10]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
# ADRESS_LIVE_LABELS
# Проживание -> Регион
                    elif label0 == MANIPULATE_LABELS[11]:
                        addr = norm.field2addr(row_item)
                        lab = [ADRESS_LIVE_LABELS[1], ADRESS_LIVE_LABELS[2]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
# Проживание -> Район
                    elif label0 == MANIPULATE_LABELS[12]:
                        addr = norm.field2addr(row_item)
                        lab = [ADRESS_LIVE_LABELS[3], ADRESS_LIVE_LABELS[4]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
# Проживание -> Город
                    elif label0 == MANIPULATE_LABELS[13]:
                        addr = norm.field2addr(row_item)
                        lab = [ADRESS_LIVE_LABELS[5], ADRESS_LIVE_LABELS[6]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
# Проживание -> Населенный_пункт
                    elif label0 == MANIPULATE_LABELS[14]:
                        addr = norm.field2addr(row_item)
                        lab = [ADRESS_LIVE_LABELS[7], ADRESS_LIVE_LABELS[8]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
# Проживание -> Улица
                    elif label0 == MANIPULATE_LABELS[15]:
                        addr = norm.field2addr(row_item)
                        lab = [ADRESS_LIVE_LABELS[9], ADRESS_LIVE_LABELS[10]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]
#Адрес регистрации из_поля
                    elif label0 == MANIPULATE_LABELS[17]:
                        result_row[ADRESS_REG_LABELS[0]] = '111111'
                        adress_reg = norm.FullAdress(row_item)
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
                        adress_zhit = norm.FullAdress(row_item)
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
                        addr = norm.field2addr(norm.REGIONS[norm.intl(row_item)])
                        lab = [ADRESS_REG_LABELS[1], ADRESS_REG_LABELS[2]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]

# Регион проживания из номера
                    elif label0 == MANIPULATE_LABELS[23]:
                        addr = norm.field2addr(norm.REGIONS[norm.intl(row_item)])
                        lab = [ADRESS_LIVE_LABELS[1], ADRESS_LIVE_LABELS[2]]
                        for j in range(len(addr)):
                            result_row[lab[j]] = addr[j]

                elif label0 == '-------------------------':
                    continue
                elif label0 in SNILS_LABEL:
                    if norm.GENERATE_SNILS:
                        dbconfig = read_config(filename='NormXLS.ini', section='main_mysql')
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
                        result_row[label0] = norm.normalize_snils(cached_snils)
                        self.ws_comp.append([row_item, cached_snils])
                    else:
                        result_row[label0] = norm.normalize_snils(row_item)
                elif label0 in PLACE_BIRTH_LABELS:
                    result_row[label0] = row_item
                elif label0 in PASSPORT_DATA_LABELS:
                    if PASSPORT_DATA_LABELS.index(label0) == 0:
                        passport.seriya = row_item
                    elif PASSPORT_DATA_LABELS.index(label0) == 1:
                        passport.nomer = row_item
                    elif PASSPORT_DATA_LABELS.index(label0) == 2:
                        passport.date = row_item
                    elif PASSPORT_DATA_LABELS.index(label0) == 3:
                        passport.who = norm.normalize_text(row_item)
                    elif PASSPORT_DATA_LABELS.index(label0) == 4:
                        passport.cod = row_item

                elif label0 in PHONES_LABELS:
                    if PHONES_LABELS.index(label0) == 0:
                        phone.tel_mob = row_item
                    elif PHONES_LABELS.index(label0) == 1:
                        phone.tel_rod = row_item
                    elif PHONES_LABELS.index(label0) == 2:
                        phone.tel_dom = row_item
                elif label0 in DATE_BIRTH_LABEL:
                    result_row[label0] = norm.normalize_date(row_item)
                elif label0 in GENDER_LABEL:
                    result_row[label0] = norm.normalize_gender(row_item)
                elif label0 == ADRESS_REG_LABELS[0] or label0 == ADRESS_LIVE_LABELS[0]:
                    result_row[label0] = norm.normalize_index(row_item)
                elif label0 in ADRESS_REG_LABELS[11]:
                    result_row[label0] = norm.normalize_home(row_item)
                elif label0 in ADRESS_LIVE_LABELS[11]:
                    result_row[label0] = norm.normalize_home(row_item)
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
                        result_row[label0] = norm.PREDSTRAH_ID
                else:
                    result_row[label0] = norm.normalize_text(row_item)

            for num, z in enumerate(passport.get_values()):
                result_row[PASSPORT_DATA_LABELS[num]] = z
            for num, z in enumerate(phone.get_values()):
                result_row[PHONES_LABELS[num]] = z

            LABELS = [SNILS_LABEL, FIO_LABELS, FIO_BIRTH_LABELS, FIO_SNILS_LABELS, GENDER_LABEL, DATE_BIRTH_LABEL,
                      PLACE_BIRTH_LABELS, PASSPORT_DATA_LABELS, ADRESS_REG_LABELS, ADRESS_LIVE_LABELS,
                      PHONES_LABELS, TECH_LABELS]
            mass = []
            for l in LABELS:
                for ll in l:
                    mass.append(ll)
            yum = True
            yum_phone0 = -1
            yum_phone1 = -1
            yum_phone2 = -1
            for num, cell in enumerate(mass):
                if cell in result_row:
                    mass[num] = result_row[cell]                # заполняем mass, чтобы его добавить как строку в xlsx
                    if cell == PHONES_LABELS[0]:
                        if mass[num] == norm.ERROR_VALUE:
                            mass[num] = ''
                        yum_phone0 = num
                    elif cell == PHONES_LABELS[1]:
                        if mass[num] == norm.ERROR_VALUE:
                            mass[num] = ''
                        yum_phone1 = num
                    elif cell == PHONES_LABELS[2]:
                        if mass[num] == norm.ERROR_VALUE:
                            mass[num] = ''
                        yum_phone2 = num
                    elif mass[num] == norm.ERROR_VALUE:
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

            if i10 > i10l:                                  # режем по 10000
                i10l = i10
                f = self.fname.replace(self.fname.split('/')[-1], '{0:02d}'.format(i10) + '_' + self.fname.split('/')[-1])
                wb.save(f)
                wb = Workbook(write_only=True)
                ws = wb.create_sheet('Лист1')
                ws.append(HEAD_RESULT_EXCEL_FILE)  # добавляем первую строку xlsx

        f = self.fname.replace(self.fname.split('/')[-1], '{0:02d}'.format(i10+1) + '_'+ self.fname.split('/')[-1])
        wb.save(f)
        f = self.fname.replace(self.fname.split('/')[-1], 'err'.format(i10+1) + self.fname.split('/')[-1])
        wb_err.save(f)
        if use_log:
            log_file.close()
        if norm.GENERATE_SNILS:
            self.wb_comp.save(self.fname.replace(self.fname.split('/')[-1], 'com'.format(i10+1)
                                                 + self.fname.split('/')[-1]))

