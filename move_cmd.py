# Запуск из командной строки (генерируется в move.py)

import openpyxl
import sys
import argparse

IN_IDS = ['ID','ИД_КЛИЕНТА','CLIENT_ID']
IN_SNILS = ['СНИЛС', 'СТРАХОВОЙ_НОМЕР', 'СТРАХОВОЙ НОМЕР', 'NUMBER','СТРАХОВОЙНОМЕР']
IN_NAMES = ['ID', 'СНИЛС', 'СТРАХОВОЙ_НОМЕР', 'СТРАХОВОЙ НОМЕР', 'NUMBER', 'ФАМИЛИЯ', 'ИМЯ', 'ОТЧЕСТВО', 'ФИО']


def createParser ():
    parser = argparse.ArgumentParser()
    parser.add_argument ('-sheetName')
    parser.add_argument ('-leFond')
    parser.add_argument ('-leAgent')
    parser.add_argument ('-leSigner')
    parser.add_argument ('-chbClientOnly')
    parser.add_argument ('-chbSocium')
    parser.add_argument ('-chbSuff')
    parser.add_argument ('-chbOurStat')
    parser.add_argument ('-chbFondStat')
    parser.add_argument ('-chbArhivON')
    parser.add_argument ('-chbArhivOFF')
    parser.add_argument ('-chbNoDubPhonePartner')


class my(object):
    def __init__(self, namespace):
        self.args = namespace

    def one(self):
        wb = openpyxl.load_workbook(sys.argv[1])
        if not self.args.sheetName:
            self.args.sheetName = wb.sheetnames[0]
        self.sheet = wb[self.args.sheetName]
        if not self.sheet.max_row:
            print('Файл Excel некорректно сохранен OpenPyxl. Откройте и пересохраните его')
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
                    print('В файле ' + sys.argv[1] + ' на вкладке ' + wb.sheetnames[0] + ' отсутствует колонка с ID')
                    return
            elif j == 1:
                for k, cell in enumerate(row):
                    if cell.value != None:
                        if str(cell.value).strip() != '':
                            if k > last_cell:
                                last_cell = k

        self.table = []
        for i, row in enumerate(self.sheet.rows):
            table_row = []
            for j, cell in enumerate(row):
                if j > last_cell:
                    break
                table_row.append(cell.value)
            self.table.append(table_row)


        if not self.file_touched:  # Проверяем достаточность данных
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

        if not self.chbNoBackup.isChecked():
            all_clients_ids = "'" + self.clients_ids[0] + "'"  # Проверка на дубли clients
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
                    ws_clients.append([row[0]])
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
                    ws_contracts.append([row[0]])
            else:
                ws_log.append([datetime.now().strftime("%H:%M:%S"), ' В contracts нет дублей'])
            if exit_because_doubles:  # Если дубли в clients или contracts - ничего не переносим
                ws_log.append([datetime.now().strftime("%H:%M:%S"), 'Аварийное завершение - дублирование записей'])
                return

        # Проверка на дубли исходной таблицы
        doubles_in_input = list(set([x for x in self.clients_ids if self.clients_ids.count(x) > 1]))
        if len(doubles_in_input) > 0:
            ws_log.append([datetime.now().strftime("%H:%M:%S"), ' Дубли в исходной таблице'])
            ws_input_doubles = wb_log.create_sheet('Дубли в исходной таблице')
            for row in doubles_in_input:
                ws_input_doubles.append(row)
        else:
            ws_log.append([datetime.now().strftime("%H:%M:%S"), ' В исходной таблице нет дублей'])

        if not self.chbNoBackup.isChecked():
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
        if self.chbNoDubPhonePartner.isChecked():
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'Без дублей телефонов у партнера', 'выбрано'])
        else:
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'Без дублей телефонов у партнера', 'не выбрано'])

        # Список телефонов у партнера в фонде в который переносим
        if self.chbNoDubPhonePartner.isChecked() and self.leAgent.isEnabled():
            dbconn = MySQLConnection(**self.dbconfig)
            cursor = dbconn.cursor()
            cursor.execute('SELECT partner_code FROM offices_staff WHERE code = %s',
                           (self.agent_ids[self.cmbAgent.currentIndex()],))
            partner = cursor.fetchall()
            if self.partner != partner[0][0]:
                self.partner = partner[0][0]
                phones = []
                cursor = dbconn.cursor()
                sql_tel = 'SELECT phone_personal_mobile FROM clients AS cl LEFT JOIN offices_staff AS os ' \
                          'ON cl.inserted_user_code = os.code WHERE os.partner_code = %s'
                if self.leFond.isEnabled():
                    cursor.execute(sql_tel + ' AND cl.subdomain_id = %s', (partner[0][0],
                                                                           self.fond_ids[self.cmbFond.currentIndex()]))
                else:
                    cursor.execute(sql_tel, (partner[0][0],))
                phones_sql = cursor.fetchall()
                self.progressBar.setMaximum(len(phones_sql) - 1)
                for i, phone_sql in enumerate(phones_sql):
                    if not (i % 10000):
                        self.progressBar.setValue(i)
                    if phone_sql[0] and phone_sql[0] not in phones:
                        phones.append(phone_sql[0])
                    # if i > 10000:
                    #    break
                self.phones = phones
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

        tuples_clients = []  # Формируем переменные для запросов
        tuples_contracts = []
        self.progressBar.setMaximum(len(self.clients_ids) - 1)
        dbconn = MySQLConnection(**self.dbconfig)
        cursor = dbconn.cursor()
        cursor_phones = dbconn.cursor()
        i_tek = 0
        if len(self.phones):
            ws_phones_doubles = wb_log.create_sheet('Дубли телефонов у партнера')
            ws_phones_doubles.append(['client_id', 'СНИЛС', 'телефон'])
        for i, client_id in enumerate(self.clients_ids):
            if len(self.phones):
                cursor_phones.execute('SELECT phone_personal_mobile, number FROM clients AS cl WHERE cl.client_id = %s',
                                      (client_id,))
                rows = cursor_phones.fetchall()
                if rows[0][0] in self.phones:
                    ws_phones_doubles.append([client_id, rows[0][1], rows[0][0]])
                    continue
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
                tuple_contract += (0, 0, 0)
            if self.chbOurStat.isChecked():
                tuple_contract += (0, 0, 0)
            if self.chbSuff.isChecked():
                tuple_contract += (self.leSuff.text(),)
            tuple_contract += (client_id,)
            tuple_client += (client_id,)
            tuples_clients.append(tuple_client)
            tuples_contracts.append(tuple_contract)
            if i_tek and not (i_tek % 1000):
                self.progressBar.setValue(i)
                if self.leSQLcl.text():
                    cursor.executemany(self.leSQLcl.text(), tuples_clients)
                    dbconn.commit()
                    tuples_clients = []
                if self.leSQLco.text():
                    cursor.executemany(self.leSQLco.text(), tuples_contracts)
                    dbconn.commit()
                    tuples_contracts = []
            i_tek += 1
        if self.leSQLcl.text():
            ws_log.append([datetime.now().strftime("%H:%M:%S"), self.leSQLcl.text()])
            ws_log.append([datetime.now().strftime("%H:%M:%S")] + list(tuples_clients[0]))
            if len(tuples_clients):
                cursor.executemany(self.leSQLcl.text(), tuples_clients)
                dbconn.commit()
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'update clients отработал'])
        else:
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'clients - нет запроса'])
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'clients - нет данных'])
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'clients - запрос не исполнен'])
        if self.leSQLco.text():
            ws_log.append([datetime.now().strftime("%H:%M:%S"), self.leSQLco.text()])
            ws_log.append([datetime.now().strftime("%H:%M:%S")] + list(tuples_contracts[0]))
            if len(tuples_contracts):
                cursor.executemany(self.leSQLco.text(), tuples_contracts)
                dbconn.commit()
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'update contracts отработал'])
        else:
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'contracts - нет запроса'])
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'contracts - нет данных'])
            ws_log.append([datetime.now().strftime("%H:%M:%S"), 'contracts - запрос не исполнен'])
        wb_log.save(log_name)

if __name__ == '__main__':
    parser = createParser()
    all = my(parser.parse_args(sys.argv[1:]))
    all.one()

