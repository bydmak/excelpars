import os
from pprint import pprint
import openpyxl
import json
import re
from PyQt5.QtCore import QSize, Qt
from PyQt5.QtGui import QStandardItem
from PyQt5.QtWidgets import QApplication, QMainWindow, QGridLayout, QWidget, QTableWidget, QTableWidgetItem, QHeaderView
from PyQt5 import QtWidgets
from PyQt5 import QtGui
from PyQt5 import QtCore
from mydesign import Ui_MainWindow  # импорт нашего сгенерированного файла
import sys

class mywindow(QtWidgets.QMainWindow):
    def __init__(self):# Обязательно нужно вызвать метод супер класса
        QMainWindow.__init__(self)

        self.setMinimumSize(QSize(1366, 768))             # Устанавливаем размеры
        self.setWindowTitle("2C")    # Устанавливаем заголовок окна
        central_widget = QWidget(self)                  # Создаём центральный виджет
        self.setCentralWidget(central_widget)  # Устанавливаем центральный виджет
        self.pushButton = QtWidgets.QPushButton(self)
        self.pushButton_2 = QtWidgets.QPushButton(self)
        self.pushButton_3 = QtWidgets.QPushButton(self)
        self.pushButton_4 = QtWidgets.QPushButton(self)
        self.pushButton_5 = QtWidgets.QPushButton(self)
        self.textEdit_2 = QtWidgets.QTextEdit(self)
        self.textEdit = QtWidgets.QTextEdit(self)
        self.textEdit_3 = QtWidgets.QTextEdit(self)
        self.textEdit_4 = QtWidgets.QTextEdit(self)


        self.pushButton.setGeometry(QtCore.QRect(30, 30, 201, 41))
        self.pushButton_2.setGeometry(QtCore.QRect(240, 10, 171, 31))
        self.pushButton_3.setGeometry(QtCore.QRect(240, 60, 171, 31))
        self.pushButton_4.setGeometry(QtCore.QRect(610, 60, 171, 31))
        self.pushButton_5.setGeometry(QtCore.QRect(610, 10, 171, 31))

        font = QtGui.QFont()
        font.setPointSize(16)
        font1 = QtGui.QFont()
        font1.setPointSize(10)
        self.pushButton.setFont(font)
        self.pushButton.setStyleSheet("background-color: yellow;\n"
                                      "border-radius: 20%")


        self.pushButton.setText("Выбрать таблицу")
        self.pushButton_2.setText("Поиск")
        self.textEdit.setPlaceholderText("Введите Инвентарный номер")
        self.pushButton_3.setText("Поиск")
        self.textEdit_2.setPlaceholderText("Введите Основное средство")
        self.textEdit_3.setPlaceholderText("Введите ФИО")
        self.pushButton_4.setText("Поиск по ФИО")
        self.pushButton_5.setText("Поиск по Отделу")
        self.textEdit_4.setPlaceholderText("Введите Отдел")

        self.textEdit.setGeometry(QtCore.QRect(410, 10, 171, 31))
        self.textEdit_2.setGeometry(QtCore.QRect(410, 60, 171, 31))
        self.textEdit_3.setGeometry(QtCore.QRect(780, 60, 171, 31))
        self.textEdit_4.setGeometry(QtCore.QRect(780, 10, 171, 31))

        #self.pushButton.clicked.connect(self.zapisvtable)
        self.pushButton_2.clicked.connect(self.chtetieisjson)
        self.pushButton_3.clicked.connect(self.chtetieisjson1)
        self.pushButton_4.clicked.connect(self.PoiskFIO)
        self.pushButton_5.clicked.connect(self.PoiskOtdel)


        grid_layout = QGridLayout()             # Создаём QGridLayout
        central_widget.setLayout(grid_layout)   # Устанавливаем данное размещение в центральный виджет

        with open('data_file.json', encoding='windows-1251') as json_file:
            self.data = json.load(json_file)
        self.table = QTableWidget(self)  # Создаём таблицу
        self.table.setColumnCount(17)     # Устанавливаем  колонки
        self.table.setRowCount(len(self.data['Numenklatures']))        # строки в таблицу
        self.table.setHorizontalHeaderLabels(
            [
                "Категория",
                "ФИО ответственного",
                "Отдел",
                "Основное средство",
                "Инвентарный номер",
                "ОКОФ",
                "Амортизационная группа",
                "Способ начисления амортизации",
                "Дата ввода в эксплуатацию",
                "Состояние",
                "Срок полезного использования",
                "Мес. норма износа",
                "Износ",
                "Балансовая стоимость",
                "Количество",
                "Сумма амортизации",
                "Остаточная стоимость"])

        with open('data_file.json', encoding='windows-1251') as json_file:
            data = json.load(json_file)
            i = 0
            for p in data['Numenklatures']:
                stolb1 = str(p['Категория'])
                stolb2 = str(p['ФИО ответственного'])
                stolb3 = str(p['Отдел'])
                stolb4 = str(p['Основное средство'])
                stolb5 = str(p['Инвентарный номер'])
                stolb6 = str(p['ОКОФ'])
                stolb7 = str(p['Амортизационная группа'])
                stolb8 = str(p['Способ начисления амортизации'])
                stolb9 = str(p['Дата ввода в эксплуатацию'])
                stolb10 = str(p['Состояние'])
                stolb11 = str(p['Срок полезного использования'])
                stolb12 = str(p['Мес. норма износа'])
                stolb13 = str(p['Износ'])
                stolb14 = str(p['Балансовая стоимость'])
                stolb15 = str(p['Количество'])
                stolb16 = str(p['Сумма амортизации'])
                stolb17 = str(p['Остаточная стоимость'])
                self.table.setItem(i, 0, QTableWidgetItem(stolb1))
                self.table.setItem(i, 1, QTableWidgetItem(stolb2))
                self.table.setItem(i, 2, QTableWidgetItem(stolb3))
                self.table.setItem(i, 3, QTableWidgetItem(stolb4))
                self.table.setItem(i, 4, QTableWidgetItem(stolb5))
                self.table.setItem(i, 5, QTableWidgetItem(stolb6))
                self.table.setItem(i, 6, QTableWidgetItem(stolb7))
                self.table.setItem(i, 7, QTableWidgetItem(stolb8))
                self.table.setItem(i, 8, QTableWidgetItem(stolb9))
                self.table.setItem(i, 9, QTableWidgetItem(stolb10))
                self.table.setItem(i, 10, QTableWidgetItem(stolb11))
                self.table.setItem(i, 11, QTableWidgetItem(stolb12))
                self.table.setItem(i, 12, QTableWidgetItem(stolb13))
                self.table.setItem(i, 13, QTableWidgetItem(stolb14))
                self.table.setItem(i, 14, QTableWidgetItem(stolb15))
                self.table.setItem(i, 15, QTableWidgetItem(stolb16))
                self.table.setItem(i, 16, QTableWidgetItem(stolb17))
                i = i + 1

        # делаем ресайз колонок по содержимому
        self.table.resizeColumnsToContents()

        grid_layout.addWidget(self.table, 0, 0)   # Добавляем таблицу в сетку
        grid_layout.setContentsMargins(30, 100, 30, 30)

    def buttonClicked(self):
        file, _ = QtWidgets.QFileDialog.getOpenFileName(self,
                                                        'Open File',
                                                        './',
                                                        'py Files (*.xlsx)')
        print(type(file))
        excel_file = openpyxl.load_workbook(file)
        namelist = excel_file.sheetnames  # ['TDSheet'] - получение название листа
        employees_sheet = excel_file[str(namelist[0])]
        currently_active_sheet = excel_file.active
        print(employees_sheet["A" + '50'].value)
        if not file:
            return

    def chtetieisjson(self):
        maxcolumuns = len(self.data['Numenklatures'])
        self.table.setRowCount(maxcolumuns)
        textboxValue = self.textEdit.toPlainText()
        with open('data_file.json', encoding='windows-1251') as json_file:
            data = json.load(json_file)
            i = 0
            for p in data['Numenklatures']:
                if p['Инвентарный номер'] == textboxValue:
                    self.table.setItem(i, 0, QTableWidgetItem(str(p['Категория'])))
                    self.table.setItem(i, 1, QTableWidgetItem(str(p['ФИО ответственного'])))
                    self.table.setItem(i, 2, QTableWidgetItem(str(p['Отдел'])))
                    self.table.setItem(i, 3, QTableWidgetItem(str(p['Основное средство'])))
                    self.table.setItem(i, 4, QTableWidgetItem(str(p['Инвентарный номер'])))
                    self.table.setItem(i, 5, QTableWidgetItem(str(p['ОКОФ'])))
                    self.table.setItem(i, 6, QTableWidgetItem(str(p['Амортизационная группа'])))
                    self.table.setItem(i, 7, QTableWidgetItem(str(p['Способ начисления амортизации'])))
                    self.table.setItem(i, 8, QTableWidgetItem(str(p['Дата ввода в эксплуатацию'])))
                    self.table.setItem(i, 9, QTableWidgetItem(str(p['Состояние'])))
                    self.table.setItem(i, 10, QTableWidgetItem(str(p['Срок полезного использования'])))
                    self.table.setItem(i, 11, QTableWidgetItem(str(p['Мес. норма износа'])))
                    self.table.setItem(i, 12, QTableWidgetItem(str(p['Износ'])))
                    self.table.setItem(i, 13, QTableWidgetItem(str(p['Балансовая стоимость'])))
                    self.table.setItem(i, 14, QTableWidgetItem(str(p['Количество'])))
                    self.table.setItem(i, 15, QTableWidgetItem(str(p['Сумма амортизации'])))
                    self.table.setItem(i, 16, QTableWidgetItem(str(p['Остаточная стоимость'])))
                    i = i + 1
            while self.table.rowCount() > i:
                    self.table.removeRow(i)
            textboxValue = 0

    def chtetieisjson1(self):
        maxcolumuns = len(self.data['Numenklatures'])
        self.table.setRowCount(maxcolumuns)
        textboxValue = self.textEdit_2.toPlainText()
        with open('data_file.json', encoding='windows-1251') as json_file:
            data = json.load(json_file)
            i = 0
            for p in data['Numenklatures']:
                if p['Основное средство'] == textboxValue:
                    self.table.setItem(i, 0, QTableWidgetItem(str(p['Категория'])))
                    self.table.setItem(i, 1, QTableWidgetItem(str(p['ФИО ответственного'])))
                    self.table.setItem(i, 2, QTableWidgetItem(str(p['Отдел'])))
                    self.table.setItem(i, 3, QTableWidgetItem(str(p['Основное средство'])))
                    self.table.setItem(i, 4, QTableWidgetItem(str(p['Инвентарный номер'])))
                    self.table.setItem(i, 5, QTableWidgetItem(str(p['ОКОФ'])))
                    self.table.setItem(i, 6, QTableWidgetItem(str(p['Амортизационная группа'])))
                    self.table.setItem(i, 7, QTableWidgetItem(str(p['Способ начисления амортизации'])))
                    self.table.setItem(i, 8, QTableWidgetItem(str(p['Дата ввода в эксплуатацию'])))
                    self.table.setItem(i, 9, QTableWidgetItem(str(p['Состояние'])))
                    self.table.setItem(i, 10, QTableWidgetItem(str(p['Срок полезного использования'])))
                    self.table.setItem(i, 11, QTableWidgetItem(str(p['Мес. норма износа'])))
                    self.table.setItem(i, 12, QTableWidgetItem(str(p['Износ'])))
                    self.table.setItem(i, 13, QTableWidgetItem(str(p['Балансовая стоимость'])))
                    self.table.setItem(i, 14, QTableWidgetItem(str(p['Количество'])))
                    self.table.setItem(i, 15, QTableWidgetItem(str(p['Сумма амортизации'])))
                    self.table.setItem(i, 16, QTableWidgetItem(str(p['Остаточная стоимость'])))
                    i = i + 1
        while self.table.rowCount() > i:
            self.table.removeRow(i)
            textboxValue = 0

    def PoiskFIO(self):
        maxcolumuns = len(self.data['Numenklatures'])
        self.table.setRowCount(maxcolumuns)
        textboxValue = self.textEdit_3.toPlainText()
        with open('data_file.json', encoding='windows-1251') as json_file:
            data = json.load(json_file)
            i = 0
            for p in data['Numenklatures']:
                if p['ФИО ответственного'] == textboxValue:
                    self.table.setItem(i, 0, QTableWidgetItem(str(p['Категория'])))
                    self.table.setItem(i, 1, QTableWidgetItem(str(p['ФИО ответственного'])))
                    self.table.setItem(i, 2, QTableWidgetItem(str(p['Отдел'])))
                    self.table.setItem(i, 3, QTableWidgetItem(str(p['Основное средство'])))
                    self.table.setItem(i, 4, QTableWidgetItem(str(p['Инвентарный номер'])))
                    self.table.setItem(i, 5, QTableWidgetItem(str(p['ОКОФ'])))
                    self.table.setItem(i, 6, QTableWidgetItem(str(p['Амортизационная группа'])))
                    self.table.setItem(i, 7, QTableWidgetItem(str(p['Способ начисления амортизации'])))
                    self.table.setItem(i, 8, QTableWidgetItem(str(p['Дата ввода в эксплуатацию'])))
                    self.table.setItem(i, 9, QTableWidgetItem(str(p['Состояние'])))
                    self.table.setItem(i, 10, QTableWidgetItem(str(p['Срок полезного использования'])))
                    self.table.setItem(i, 11, QTableWidgetItem(str(p['Мес. норма износа'])))
                    self.table.setItem(i, 12, QTableWidgetItem(str(p['Износ'])))
                    self.table.setItem(i, 13, QTableWidgetItem(str(p['Балансовая стоимость'])))
                    self.table.setItem(i, 14, QTableWidgetItem(str(p['Количество'])))
                    self.table.setItem(i, 15, QTableWidgetItem(str(p['Сумма амортизации'])))
                    self.table.setItem(i, 16, QTableWidgetItem(str(p['Остаточная стоимость'])))
                    i = i + 1
        while self.table.rowCount() > i:
            self.table.removeRow(i)
            textboxValue = 0

    def PoiskOtdel(self):
        maxcolumuns = len(self.data['Numenklatures'])
        self.table.setRowCount(maxcolumuns)
        textboxValue = self.textEdit_4.toPlainText()
        with open('data_file.json', encoding='windows-1251') as json_file:
            data = json.load(json_file)
            i = 0
            for p in data['Numenklatures']:
                if p['Отдел'] == textboxValue:
                    self.table.setItem(i, 0, QTableWidgetItem(str(p['Категория'])))
                    self.table.setItem(i, 1, QTableWidgetItem(str(p['ФИО ответственного'])))
                    self.table.setItem(i, 2, QTableWidgetItem(str(p['Отдел'])))
                    self.table.setItem(i, 3, QTableWidgetItem(str(p['Основное средство'])))
                    self.table.setItem(i, 4, QTableWidgetItem(str(p['Инвентарный номер'])))
                    self.table.setItem(i, 5, QTableWidgetItem(str(p['ОКОФ'])))
                    self.table.setItem(i, 6, QTableWidgetItem(str(p['Амортизационная группа'])))
                    self.table.setItem(i, 7, QTableWidgetItem(str(p['Способ начисления амортизации'])))
                    self.table.setItem(i, 8, QTableWidgetItem(str(p['Дата ввода в эксплуатацию'])))
                    self.table.setItem(i, 9, QTableWidgetItem(str(p['Состояние'])))
                    self.table.setItem(i, 10, QTableWidgetItem(str(p['Срок полезного использования'])))
                    self.table.setItem(i, 11, QTableWidgetItem(str(p['Мес. норма износа'])))
                    self.table.setItem(i, 12, QTableWidgetItem(str(p['Износ'])))
                    self.table.setItem(i, 13, QTableWidgetItem(str(p['Балансовая стоимость'])))
                    self.table.setItem(i, 14, QTableWidgetItem(str(p['Количество'])))
                    self.table.setItem(i, 15, QTableWidgetItem(str(p['Сумма амортизации'])))
                    self.table.setItem(i, 16, QTableWidgetItem(str(p['Остаточная стоимость'])))
                    i = i + 1
        while self.table.rowCount() > i:
            self.table.removeRow(i)
            textboxValue = 0

app = QtWidgets.QApplication([])
app.setWindowIcon(QtGui.QIcon('icon1.png'))
application = mywindow()
application.setWindowIcon(QtGui.QIcon('icon1.png'))
application.show()
sys.exit(app.exec())

#pyinstaller -F --noconsole -i icon1.ico test.py