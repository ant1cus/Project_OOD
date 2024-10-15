import json
import os
import pathlib
import default_window

from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QLineEdit, QDialog, QButtonGroup, QLabel, QSizePolicy, QPushButton, QFileDialog, QComboBox,\
    QDoubleSpinBox
from PyQt5.QtCore import QDir
from rewrite_settings import rewrite


class Button(QLineEdit):

    def __init__(self, parent):
        super(Button, self).__init__(parent)

        self.setAcceptDrops(True)

    def dragEnterEvent(self, e):

        if e.mimeData().hasUrls():
            e.accept()
        else:
            super(Button, self).dragEnterEvent(e)

    def dragMoveEvent(self, e):

        super(Button, self).dragMoveEvent(e)

    def dropEvent(self, e):

        if e.mimeData().hasUrls():
            for url in e.mimeData().urls():
                self.setText(os.path.normcase(url.toLocalFile()))
                e.accept()
        else:
            super(Button, self).dropEvent(e)


class DefaultWindow(QDialog, default_window.Ui_Dialog):  # Настройки по умолчанию
    def __init__(self, parent, path, name_list):
        super().__init__()
        self.setupUi(self)
        self.path_for_default = path
        self.parent = parent
        self.name_list = name_list
        self.name_box = [self.groupBox_checked, self.groupBox_parcing, self.groupBox_exctracting,
                         self.groupBox_gen_pemi, self.groupBox_gen_hfe, self.groupBox_gen_hfi,
                         self.groupBox_application, self.groupBox_gen_LF, self.groupBox_cotinuous_spectrum,
                         self.groupBox_number_instance, self.groupBox_finding_file]
        self.name_grid = [self.gridLayout_checked, self.gridLayout_parcer, self.gridLayout_exctract,
                          self.gridLayout_gen_pemi, self.gridLayout_HFE, self.gridLayout_HFI,
                          self.gridLayout_application, self.gridLayout_gen_LF, self.gridLayout_cotinuous_spectrum,
                          self.gridLayout_number_instance, self.gridLayout_finding_file]
        with open(pathlib.Path(self.path_for_default, 'Настройки.txt'), "r", encoding='utf-8-sig') as f:  # Открываем
            dict_load = json.load(f)  # Загружаем данные
            self.data = dict_load['widget_settings']
            self.tab_order = dict_load['gui_settings']['tab_order']
            self.tab_visible = dict_load['gui_settings']['tab_visible']
            self.version = dict_load['version']['actual_version']
        self.buttongroup_add = QButtonGroup()
        self.buttongroup_add.buttonClicked[int].connect(self.add_button_clicked)
        self.buttongroup_clear = QButtonGroup()
        self.buttongroup_clear.buttonClicked[int].connect(self.clear_button_clicked)
        self.buttongroup_open = QButtonGroup()
        self.buttongroup_open.buttonClicked[int].connect(self.open_button_clicked)
        self.pushButton_ok.clicked.connect(self.accept)  # Принять
        self.pushButton_cancel.clicked.connect(lambda: self.close())  # Отмена
        self.line = {}  # Для имен
        self.name = {}  # Для значений
        self.combo = {}  # Для комбобоксов
        self.button = {}  # Для кнопки «изменить»
        self.button_clear = {}  # Для кнопки «очистить»
        self.button_open = {}  # Для кнопки «открыть»
        for i, el in enumerate(self.name_list):  # Заполняем
            frame = False
            grid = False
            for j, n in enumerate(['checked', 'parser', 'extract', 'gen_pemi', 'HFE', 'HFI', 'application', 'LF',
                                   'CC', 'NI', 'FF']):
                if n in el.partition('-')[0]:
                    frame = self.name_box[j]
                    grid = self.name_grid[j]
                    break
            self.line[i] = QLabel(frame)  # Помещаем в фрейм
            self.line[i].setText(self.name_list[el][0])  # Название элемента
            self.line[i].setFont(QFont("Times", 12, QFont.Light))  # Шрифт, размер
            self.line[i].setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)  # Размеры виджета
            self.line[i].setFixedWidth(250)
            self.line[i].setDisabled(True)  # Делаем неактивным, чтобы нельзя было просто так редактировать
            grid.addWidget(self.line[i], i, 0)  # Добавляем виджет
            if 'checkBox' not in el and 'groupBox' not in el:
                self.button[i] = QPushButton("Изменить", frame)  # Создаем кнопку
                self.button[i].setFont(QFont("Times", 12, QFont.Light))  # Размер шрифта
                self.button[i].setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)  # Размеры виджета
                self.buttongroup_add.addButton(self.button[i], i)  # Добавляем в группу
                grid.addWidget(self.button[i], i, 1)  # Добавляем в фрейм по месту

                self.button_clear[i] = QPushButton("Очистить", frame)  # Создаем кнопку
                self.button_clear[i].setFont(QFont("Times", 12, QFont.Light))  # Размер шрифта
                self.button_clear[i].setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)  # Размеры виджета
                self.buttongroup_clear.addButton(self.button_clear[i], i)  # Добавляем в группу
                grid.addWidget(self.button_clear[i], i, 2)  # Добавляем в фрейм по месту
                if 'spinbox' in el.lower():
                    self.name[i] = QDoubleSpinBox()
                    if el in self.data:
                        self.name[i].setValue(float(self.data[el]))
                    self.name[i].setSingleStep(0.1)
                    self.name[i].setDecimals(1)
                else:
                    self.name[i] = Button(frame)  # Помещаем в фрейм
                    self.name[i].setStyleSheet("QLineEdit {"
                                               "border-style: solid;"
                                               "}")
                    if el in self.data:
                        self.name[i].setText(self.data[el])
                self.name[i].setFont(QFont("Times", 12, QFont.Light))  # Шрифт, размер
                self.name[i].setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)  # Размеры виджета
                self.name[i].setDisabled(True)  # Неактивный
                grid.addWidget(self.name[i], i, 3)  # Помещаем в фрейм
                if 'path_' in el:
                    self.button_open[i] = QPushButton("Открыть", frame)  # Создаем кнопку
                    self.button_open[i].setFont(QFont("Times", 12, QFont.Light))  # Размер шрифта
                    self.button_open[i].setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)  # Размеры виджета
                    self.button_open[i].setDisabled(True)  # Неактивный
                    self.buttongroup_open.addButton(self.button_open[i], i)  # Добавляем в группу
                    grid.addWidget(self.button_open[i], i, 4)  # Добавляем в фрейм по месту
            else:
                self.combo[i] = QComboBox(frame)  # Помещаем в фрейм
                self.combo[i].addItems(['Включён', 'Выключен'])
                self.combo[i].setCurrentIndex(0) if el in self.data and self.data[el] \
                    else self.combo[i].setCurrentIndex(1)
                self.combo[i].setFont(QFont("Times", 12, QFont.Light))  # Шрифт, размер
                self.combo[i].setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)  # Размеры виджета
                grid.addWidget(self.combo[i], i, 3)  # Помещаем в фрейм

    def open_button_clicked(self, num):  # Для кнопки открыть
        value = self.line[num].text()
        for key in self.name_list:
            if value == self.name_list[key][0]:
                if 'folder' in key:
                    directory = QFileDialog.getExistingDirectory(self, "Открыть папку", QDir.currentPath())
                else:
                    directory = QFileDialog.getOpenFileName(self, "Открыть файл", QDir.currentPath())
                if directory and isinstance(directory, tuple):
                    if directory[0]:
                        self.name[num].setText(directory[0])
                elif directory and isinstance(directory, str):
                    self.name[num].setText(directory)
                break

    def add_button_clicked(self, number):  # Если кликнули по кнопке
        self.name[number].setEnabled(True)  # Делаем активным для изменения
        if number in self.button_open:
            self.button_open[number].setEnabled(True)  # Неактивный
        self.name[number].setStyleSheet("QLineEdit {"
                                        "border-style: solid;"
                                        "border-width: 1px;"
                                        "border-color: black; "
                                        "}")

    def clear_button_clicked(self, number):
        self.name[number].clear()

    def accept(self):  # Если нажали кнопку принять
        for i, el in enumerate(self.name_list):  # Пробегаем значения
            if 'checkBox' in el or 'groupBox' in el:
                self.data[el] = True if self.combo[i].currentIndex() == 0 else False
            else:
                if self.name[i].isEnabled():  # Если виджет активный (означает потенциальное изменение)
                    if self.name[i].text():  # Если внутри виджета есть текст, то помещаем внутрь базы
                        self.data[el] = self.name[i].text()
                    else:  # Если нет текста, то удаляем значение
                        self.data[el] = None
        data_insert = {"widget_settings": self.data,
                       "gui_settings":
                           {"tab_order": self.tab_order,
                            "tab_visible": self.tab_visible
                            },
                       "version": {"actual_version": self.version}
                       }
        rewrite(self.path_for_default, data_insert, widget=True)
        self.close()  # Закрываем

    def closeEvent(self, event):
        os.chdir(pathlib.Path.cwd())
        if self.sender() and self.sender().text() == 'Принять':
            event.accept()
            # Открываем и загружаем данные
            with open(pathlib.Path(self.path_for_default, 'Настройки.txt'), "r", encoding='utf-8-sig') as f:
                data = json.load(f)['widget_settings']
            self.parent.default_date(data)
            self.parent.show()
        else:
            event.accept()
            self.parent.show()
