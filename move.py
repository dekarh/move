# -*- coding: utf-8 -*-
__author__ = 'Denis'

import sys

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import (QApplication, QWidget, QMainWindow, QFileDialog, QMessageBox, QTableWidgetItem, QComboBox)
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QDate

from move_slots import MainWindowSlots

class MainWindow(MainWindowSlots):

    # При инициализации класса нам необходимо выполнить некоторые операции
    def __init__(self, form):
        # Сконфигурировать интерфейс методом из базового класса Ui_Form
        self.setupUi(form)
        # Подключить созданные нами слоты к виджетам
        self.connect_slots()

    # Подключаем слоты к виджетам (для каждого действия, которое надо обработать - свой слот)
    def connect_slots(self):
        self.pbRefresh.clicked.connect(self.click_pbRefresh)
        self.pbRefreshImport.clicked.connect(self.click_pbRefreshImport)
        self.pbIdGenerate4PasportCheck.clicked.connect(self.click_pbIdGenerate4PasportCheck)
        self.pbMove.clicked.connect(self.click_pbMove)
        self.pbAgent.clicked.connect(self.click_pbAgent)
        self.pbFond.clicked.connect(self.click_pbFond)
        self.pbSigner.clicked.connect(self.click_pbSigner)
        self.pbSaveCfgFile.clicked.connect(self.click_pbSaveCfgFile)
        self.cmbAgent.activated[str].connect(self.set_cmbAgent)
        self.cmbFond.activated[str].connect(self.set_cmbFond)
        self.cmbSigner.activated[str].connect(self.set_cmbSigner)
        self.clbXAgent.clicked.connect(self.click_clbXAgent)
        self.clbXSigner.clicked.connect(self.click_clbXSigner)
        self.clbXFond.clicked.connect(self.click_clbXFond)
        self.cmbFile.activated[str].connect(self.set_cmbFile)
        self.cmbTab.activated[str].connect(self.set_cmbTab)
        self.cmbCfgFile.activated[str].connect(self.set_cmbCfgFile)
        self.leDir.textChanged[str].connect(self.change_leDir)
        self.leAgent.textChanged[str].connect(self.change_leAgent)
        #self.deCalendar.dateChanged.connect(self.change_deCalendar)
        #self.twAllExcels.customContextMenuRequested.connect(self.click_label_3)
        self.clbImport.clicked.connect(self.click_clbImport)
        self.clbMove.clicked.connect(self.click_clbMove)
        self.clbPasport.clicked.connect(self.click_clbPasport)
        self.pbImport.clicked.connect(self.click_pbImport)
        self.pbPasportCheck.clicked.connect(self.click_pbPasportCheck)
        self.chbDateFrom.clicked.connect(self.click_chbDateFrom)
        self.chbDateTo.clicked.connect(self.click_chbDateTo)
        return None

if __name__ == '__main__':
    # Создаём экземпляр приложения
    app = QApplication(sys.argv)
    # Создаём базовое окно, в котором будет отображаться наш UI
    window = QWidget()
    # Создаём экземпляр нашего UI
    ui = MainWindow(window)
    # Отображаем окно
    window.show()
    # Обрабатываем нажатие на кнопку окна "Закрыть"
    sys.exit(app.exec_())




