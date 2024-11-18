import pandas
import sqlite3
import openpyxl
import sys
import os

from PyQt6 import uic
from PyQt6.QtGui import QAction
from PyQt6.QtWidgets import QApplication, QMainWindow, QWidget, QFileDialog, QTableWidget, QTableWidgetItem, QRadioButton, QVBoxLayout, QCheckBox


class VCardGenerator(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi("ui/main_window.ui", self)

        self.load_database: QAction = self.load_database
        self.load_database.triggered.connect(self.action_with_files)

        self.load_database: QAction = self.save_database
        self.save_database.triggered.connect(self.action_with_files)

        self.button_all_groups: QRadioButton = self.button_all_groups
        self.button_all_groups.clicked.connect(self.select_all_groups)
        self.layout_groups: QVBoxLayout = self.layout_groups
        self.group_check_boxes = []
        self.count_groups = 0

        self.button_all_teachers: QRadioButton = self.button_all_teachers
        self.button_all_teachers.clicked.connect(self.select_all_teachers)
        self.layout_teachers: QVBoxLayout = self.layout_teachers
        self.teacher_check_boxes = []
        self.count_teachers = 0

        self.tableWidget: QTableWidget = self.tableWidget
        self.titles = []
        self.data = []
        self.index_of_groups = -1
        self.index_of_teachers = -1
        self.groups = []
        self.teachers = []
        self.query = '''
        SELECT * FROM database
        '''

        self.second_form = SavingDatabase()

    def action_with_files(self):
        if self.sender().text() == "Загрузить БД":
            file_name = QFileDialog.getOpenFileName(self, 'Выберите базу данных', '', 'База данных (*.xlsx)')[0]
            excel_file = pandas.read_excel(file_name)
            self.titles = list(excel_file)
            try:
                os.remove("saves/save.sqlite")
            finally:
                con = sqlite3.connect("saves/save.sqlite")
                excel_file.to_sql(name="database", con=con, index=False)
                con.close()
                self.reloading_table()
                self.create_check_boxes()
                self.statusBar().showMessage("БД успешно загружена")
        elif self.sender().text() == "Выгрузить БД":
            self.second_form.show()

    def reloading_table(self):
        con = sqlite3.connect("saves/save.sqlite")
        cursor = con.cursor()
        self.data = cursor.execute(self.query).fetchall()
        self.tableWidget.setColumnCount(len(self.data[0]))
        self.tableWidget.setHorizontalHeaderLabels(self.titles)
        self.tableWidget.setRowCount(len(self.data))
        for i in range(len(self.data)):
            for j in range(len(self.data[0])):
                self.tableWidget.setItem(i, j, QTableWidgetItem(str(self.data[i][j])))
        self.tableWidget.resizeColumnsToContents()
        con.close()

    def create_check_boxes(self):
        self.finding_indexes()
        self.groups = self.create_values_of_kinds(self.index_of_groups)
        self.teachers = self.create_values_of_kinds(self.index_of_teachers)
        self.count_groups = len(self.groups)
        self.count_teachers = len(self.teachers)
        self.group_check_boxes = [QCheckBox()] * self.count_groups
        self.teacher_check_boxes = [QCheckBox()] * self.count_teachers
        for i in range(self.count_groups):
            group_check_box = QCheckBox(self.groups[i])
            self.group_check_boxes[i] = group_check_box
            group_check_box.clicked.connect(self.remake_query)
            self.layout_groups.addWidget(group_check_box)
        for i in range(self.count_teachers):
            teacher_check_box = QCheckBox(self.teachers[i])
            self.teacher_check_boxes[i] = teacher_check_box
            teacher_check_box.clicked.connect(self.remake_query)
            self.layout_teachers.addWidget(teacher_check_box)

    def finding_indexes(self):
        for i in range(len(self.titles)):
            if "групп" in self.titles[i].lower():
                self.index_of_groups = i
            elif "преподавател" in self.titles[i].lower():
                self.index_of_teachers = i

    def create_values_of_kinds(self, index_of_something):
        kinds = set()
        for item in self.data:
            kinds.add(item[index_of_something])
        kinds.remove(None)
        return sorted(list(kinds))

    def select_all_groups(self):
        if self.button_all_groups.isChecked():
            value = True
        else:
            value = False
        for i in range(self.count_groups):
            self.group_check_boxes[i].setChecked(value)

    def select_all_teachers(self):
        if self.button_all_teachers.isChecked():
            value = True
        else:
            value = False
        for i in range(self.count_teachers):
            self.teacher_check_boxes[i].setChecked(value)

    def remake_query(self):
        self.query = f'''
        SELECT * FROM database
        '''
        self.reloading_table()


class SavingDatabase(QWidget):
    def __init__(self):
        super().__init__()
        uic.loadUi("ui/save_window.ui", self)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = VCardGenerator()
    ex.show()
    sys.exit(app.exec())
