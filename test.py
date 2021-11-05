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
        self.textEdit_2 = QtWidgets.QTextEdit(self)
        self.textEdit = QtWidgets.QTextEdit(self)


        self.pushButton.setGeometry(QtCore.QRect(30, 30, 201, 41))
        self.pushButton_2.setGeometry(QtCore.QRect(420, 10, 171, 31))
        self.pushButton_3.setGeometry(QtCore.QRect(420, 60, 171, 31))

        font = QtGui.QFont()
        font.setPointSize(16)
        font1 = QtGui.QFont()
        font1.setPointSize(10)
        self.pushButton.setFont(font)
        self.pushButton.setStyleSheet("background-color: yellow;\n"
                                      "border-radius: 20%")

        self.pushButton.setObjectName("pushButton")
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_3.setObjectName("pushButton_3")

        self.pushButton.setText("Выбрать таблицу")
        self.pushButton_2.setText("Поиск")
        self.textEdit.setPlaceholderText("Введите Инвентарный номер")
        self.pushButton_3.setText("Поиск")
        self.textEdit_2.setPlaceholderText("Введите Основное средство")

        self.textEdit.setGeometry(QtCore.QRect(610, 10, 171, 31))
        self.textEdit.setStyleSheet("")
        self.textEdit.setCursorWidth(0)
        self.textEdit.setObjectName("textEdit")
        self.textEdit_2.setGeometry(QtCore.QRect(610, 60, 171, 31))
        self.textEdit_2.setStyleSheet("")
        self.textEdit_2.setCursorWidth(0)
        self.textEdit_2.setObjectName("textEdit_2")

        #self.pushButton.clicked.connect(self.zapisvtable)
        self.pushButton_2.clicked.connect(self.chtetieisjson)
        self.pushButton_3.clicked.connect(self.chtetieisjson1)


        grid_layout = QGridLayout()             # Создаём QGridLayout
        central_widget.setLayout(grid_layout)   # Устанавливаем данное размещение в центральный виджет

        # От сюда нужно получить сначала 101.32 или типо того, затем получить количество номенклатур входящих в неё, далее такому количеству по порядку изменить первое поле на это название.
        # С отделом и ФИО: Берем ячейку с номенклатурой и берем ячейку слева выше и проверяем ее на число, если это не число то записываем к этой номеклатуре эту клетку и над ней клетку, если число то идем выше, пока не найдем не число, если над не числом находится число, то номенклатуре добавляем прочерк.

        with open('data_file.json', encoding='windows-1251') as json_file:
            self.data = json.load(json_file)
        self.table = QTableWidget(self)  # Создаём таблицу
        self.table.setColumnCount(18)     # Устанавливаем  колонки
        self.table.setRowCount(len(self.data['Numenklatures']))        # строки в таблицу
        self.table.setHorizontalHeaderLabels(
            [
                "Категория",
                "Номер",
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
        # Устанавливаем заголовки таблицы
        #table.setHorizontalHeaderLabels([NameStolb[1], NameStolb[2], NameStolb[3], NameStolb[4], NameStolb[5], NameStolb[6], NameStolb[7], NameStolb[8], NameStolb[9], NameStolb[10], NameStolb[11], NameStolb[12], NameStolb[13], NameStolb[14], NameStolb[15], NameStolb[16], NameStolb[17], NameStolb[18]])
        # # Устанавливаем всплывающие подсказки на заголовки
        # table.horizontalHeaderItem(0).setToolTip("Column 1 ")
        # table.horizontalHeaderItem(1).setToolTip("Column 2 ")
        # table.horizontalHeaderItem(2).setToolTip("Column 3 ")

        # Устанавливаем выравнивание на заголовки
        #table.horizontalHeaderItem(0).setTextAlignment(Qt.AlignHCenter)
        #table.horizontalHeaderItem(1).setTextAlignment(Qt.AlignHCenter)
        #table.horizontalHeaderItem(2).setTextAlignment(Qt.AlignHCenter)
        # заполняем первую строку
        # for i in range(len(listName)):
        #     for j in range(18):
        #         table.setItem(i, j, QTableWidgetItem(listName[i][j+1]))

        with open('data_file.json', encoding='windows-1251') as json_file:
            data = json.load(json_file)
            i = 0
            for p in data['Numenklatures']:
                stolb1 = str(p['Категория'])
                stolb2 = str(p['Номер'])
                stolb3 = str(p['ФИО ответственного'])
                stolb4 = str(p['Отдел'])
                stolb5 = str(p['Основное средство'])
                stolb6 = str(p['Инвентарный номер'])
                stolb7 = str(p['ОКОФ'])
                stolb8 = str(p['Амортизационная группа'])
                stolb9 = str(p['Способ начисления амортизации'])
                stolb10 = str(p['Дата ввода в эксплуатацию'])
                stolb11 = str(p['Состояние'])
                stolb12 = str(p['Срок полезного использования'])
                stolb13 = str(p['Мес. норма износа'])
                stolb14 = str(p['Износ'])
                stolb15 = str(p['Балансовая стоимость'])
                stolb16 = str(p['Количество'])
                stolb17 = str(p['Сумма амортизации'])
                stolb18 = str(p['Остаточная стоимость'])
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
                self.table.setItem(i, 17, QTableWidgetItem(stolb18))
                i = i + 1
            # for p in data['Numenklatures']:
            #     if p['Инвентарный номер'] == textboxValue:
            #         print('ID: ' + p['ID'])
            # col1 = 0
            # for i in data['Numenklatures']:
            #     for j in range(18):
            #         table.setItem(i, j, QTableWidgetItem(i[col1][j+1]))
            #         col1 = col1 + 1

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
                    self.table.setItem(i, 1, QTableWidgetItem(str(p['Номер'])))
                    self.table.setItem(i, 2, QTableWidgetItem(str(p['ФИО ответственного'])))
                    self.table.setItem(i, 3, QTableWidgetItem(str(p['Отдел'])))
                    self.table.setItem(i, 4, QTableWidgetItem(str(p['Основное средство'])))
                    self.table.setItem(i, 5, QTableWidgetItem(str(p['Инвентарный номер'])))
                    self.table.setItem(i, 6, QTableWidgetItem(str(p['ОКОФ'])))
                    self.table.setItem(i, 7, QTableWidgetItem(str(p['Амортизационная группа'])))
                    self.table.setItem(i, 8, QTableWidgetItem(str(p['Способ начисления амортизации'])))
                    self.table.setItem(i, 9, QTableWidgetItem(str(p['Дата ввода в эксплуатацию'])))
                    self.table.setItem(i, 10, QTableWidgetItem(str(p['Состояние'])))
                    self.table.setItem(i, 11, QTableWidgetItem(str(p['Срок полезного использования'])))
                    self.table.setItem(i, 12, QTableWidgetItem(str(p['Мес. норма износа'])))
                    self.table.setItem(i, 13, QTableWidgetItem(str(p['Износ'])))
                    self.table.setItem(i, 14, QTableWidgetItem(str(p['Балансовая стоимость'])))
                    self.table.setItem(i, 15, QTableWidgetItem(str(p['Количество'])))
                    self.table.setItem(i, 16, QTableWidgetItem(str(p['Сумма амортизации'])))
                    self.table.setItem(i, 17, QTableWidgetItem(str(p['Остаточная стоимость'])))
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
                    self.table.setItem(i, 1, QTableWidgetItem(str(p['Номер'])))
                    self.table.setItem(i, 2, QTableWidgetItem(str(p['ФИО ответственного'])))
                    self.table.setItem(i, 3, QTableWidgetItem(str(p['Отдел'])))
                    self.table.setItem(i, 4, QTableWidgetItem(str(p['Основное средство'])))
                    self.table.setItem(i, 5, QTableWidgetItem(str(p['Инвентарный номер'])))
                    self.table.setItem(i, 6, QTableWidgetItem(str(p['ОКОФ'])))
                    self.table.setItem(i, 7, QTableWidgetItem(str(p['Амортизационная группа'])))
                    self.table.setItem(i, 8, QTableWidgetItem(str(p['Способ начисления амортизации'])))
                    self.table.setItem(i, 9, QTableWidgetItem(str(p['Дата ввода в эксплуатацию'])))
                    self.table.setItem(i, 10, QTableWidgetItem(str(p['Состояние'])))
                    self.table.setItem(i, 11, QTableWidgetItem(str(p['Срок полезного использования'])))
                    self.table.setItem(i, 12, QTableWidgetItem(str(p['Мес. норма износа'])))
                    self.table.setItem(i, 13, QTableWidgetItem(str(p['Износ'])))
                    self.table.setItem(i, 14, QTableWidgetItem(str(p['Балансовая стоимость'])))
                    self.table.setItem(i, 15, QTableWidgetItem(str(p['Количество'])))
                    self.table.setItem(i, 16, QTableWidgetItem(str(p['Сумма амортизации'])))
                    self.table.setItem(i, 17, QTableWidgetItem(str(p['Остаточная стоимость'])))
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