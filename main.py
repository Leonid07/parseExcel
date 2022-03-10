import openpyxl
from PyQt5.QtWidgets import *
from PyQt5 import QtCore, QtGui
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from array import *
from PyQt5.uic import loadUi
import sys

class Main(QDialog):
    def __init__(self):
        super(Main, self).__init__()
        loadUi('GUI.ui', self)
        self.setWindowTitle("Parsing файлов в формате xlsx")
        self.setWindowIcon(QtGui.QIcon('7422413_exel_spreadsheet_sheets_table_icon.png'))
        self.resolution.clicked.connect(self.resolution_)
        self.btn_upload_data.clicked.connect(self.upload_data_from_file)
        # Загрузка файла с расширение xlsx

    def upload_data_from_file(self):
        global path_to_file
        path_to_file = QFileDialog.getOpenFileName(self, 'Открыть файл', '', "Text Files (*.xlsx)")[0]
        # Реализация парсинга xlsx файлов при помощи библиотеки openpyxl

    def resolution_(self):
        try:
            self.plainTextEdit.clear()
            book = openpyxl.open(path_to_file, read_only=True)
            sheet = book.active
            for row in sheet.iter_rows():
                for cell in row:
                    print(cell.value, end=(' '))
                    self.plainTextEdit.insertPlainText(str(cell.value) + " ")
                self.plainTextEdit.insertPlainText("\n")
                self.plainTextEdit.insertPlainText("")
                print("")
        except:
            self.plainTextEdit.clear()
            self.plainTextEdit.appendPlainText("Не указан путь к файлу")


app = QApplication(sys.argv)
window = Main()
window.show()
sys.exit(app.exec_())
