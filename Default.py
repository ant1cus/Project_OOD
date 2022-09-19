import json
import os
import pathlib
import default_window

from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QLineEdit, QDialog, QButtonGroup, QLabel, QSizePolicy, QPushButton


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
    def __init__(self, parent, path):
        super().__init__()
        self.setupUi(self)
        self.path_for_default = path
        self.parent = parent
        # Имена
        self.name_list = {'checked-path_check': 'Путь к файлам', 'checked-table_number': 'Номер таблицы',
                          'checked-stationary_FSB': 'Стац. ФСБ', 'checked-carry_FSB': 'Воз. ФСБ',
                          'checked-wear_FSB': 'Нос. ФСБ', 'checked-r1_FSB': 'r1 ФСБ', 'checked-r1s_FSB': 'r1` ФСБ',
                          'checked-stationary_FSTEK': 'Стац. ФСТЭК', 'checked-carry_FSTEK': 'Воз. ФСТЭК',
                          'checked-wear_FSTEK': 'Нос. ФСТЭК', 'checked-r1_FSTEK': 'r1 ФСТЭК',
                          'parser-path_parser': 'Путь к файлам',
                          'extract-path_original_extract': 'Путь к файлам', 'extract-conclusion': 'ФИО заключение',
                          'extract-protocol': 'ФИО протокол', 'extract-prescription': 'ФИО предписание',
                          'gen_pemi-path_original_file': 'Путь к исходникам',
                          'gen_pemi-path_finish_folder': 'Папка для генерации',
                          'gen_pemi-path_freq_restrict': 'Файл ограничений',
                          'gen_pemi-complect_quant_pemi': 'Количество комплектов',
                          'gen_pemi-complect_number_pemi': 'Номера комплектов',
                          'HFE-path_file_HFE': 'Путь к файлам', 'HFE-complect_quant_HFE': 'Количество комплектов',
                          'HFE-frequency': 'Частота', 'HFE-level': 'Уровень',
                          'HFI-path_file_HFI': 'Путь к файлам', 'HFI-complect_quant_HFI': 'Количество комплектов',
                          'HFI-imposition_freq': 'Частота навязывания', 'application-path_example': 'Путь к файлу',
                          'application-path_finish_folder_example': 'Конечная папка',
                          'application-number_position': 'Номер позиции',
                          'application-quantity_document': 'Количество комплектов'}
        self.name_box = [self.groupBox_checked, self.groupBox_parcing, self.groupBox_exctracting,
                         self.groupBox_gen_pemi, self.groupBox_gen_hfe, self.groupBox_gen_hfi,
                         self.groupBox_application]
        self.name_grid = [self.gridLayout_checked, self.gridLayout_parcer, self.gridLayout_exctract,
                          self.gridLayout_gen_pemi, self.gridLayout_HFE, self.gridLayout_HFI,
                          self.gridLayout_application]
        with open(pathlib.Path(self.path_for_default, 'Настройки.txt'), "r", encoding='utf-8-sig') as f:  # Открываем
            dict_load = json.load(f)  # Загружаем данные
            self.data = dict_load['widget_settings']
            self.tab_order = dict_load['gui_settings']['tab_order']
            self.tab_visible = dict_load['gui_settings']['tab_visible']
        self.buttongroup_add = QButtonGroup()
        self.buttongroup_add.buttonClicked[int].connect(self.add_button_clicked)
        self.pushButton_ok.clicked.connect(self.accept)  # Принять
        self.pushButton_cancel.clicked.connect(lambda: self.close())  # Отмена
        self.line = {}  # Для имен
        self.name = {}  # Для значений
        self.button = {}  # Для кнопки «изменить»
        for i, el in enumerate(self.name_list):  # Заполняем
            frame = False
            grid = False
            for j, n in enumerate(['checked', 'parser', 'extract', 'gen_pemi', 'HFE', 'HFI', 'application']):
                if n in el.partition('-')[0]:
                    frame = self.name_box[j]
                    grid = self.name_grid[j]
                    break
            self.line[i] = QLabel(frame)  # Помещаем в фрейм
            self.line[i].setText(self.name_list[el])  # Название элемента
            self.line[i].setFont(QFont("Times", 12, QFont.Light))  # Шрифт, размер
            self.line[i].setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)  # Размеры виджета
            self.line[i].setFixedWidth(220)
            self.line[i].setDisabled(True)  # Делаем неактивным, чтобы нельзя было просто так редактировать
            grid.addWidget(self.line[i], i, 0)  # Добавляем виджет
            self.name[i] = Button(frame)  # Помещаем в фрейм
            if el in self.data:
                self.name[i].setText(self.data[el])
            self.name[i].setFont(QFont("Times", 12, QFont.Light))  # Шрифт, размер
            self.name[i].setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)  # Размеры виджета
            self.name[i].setStyleSheet("QLineEdit {"
                                       "border-style: solid;"
                                       "}")
            self.name[i].setDisabled(True)  # Неактивный
            grid.addWidget(self.name[i], i, 1)  # Помещаем в фрейм
            self.button[i] = QPushButton("Изменить", frame)  # Создаем кнопку
            self.button[i].setFont(QFont("Times", 12, QFont.Light))  # Размер шрифта
            self.button[i].setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)  # Размеры виджета
            self.buttongroup_add.addButton(self.button[i], i)  # Добавляем в группу
            grid.addWidget(self.button[i], i, 2)  # Добавляем в фрейм по месту

    def add_button_clicked(self, number):  # Если кликнули по кнопке
        self.name[number].setEnabled(True)  # Делаем активным для изменения
        self.name[number].setStyleSheet("QLineEdit {"
                                        "border-style: solid;"
                                        "border-width: 1px;"
                                        "border-color: black; "
                                        "}")

    def accept(self):  # Если нажали кнопку принять
        for i, el in enumerate(self.name_list):  # Пробегаем значения
            if self.name[i].isEnabled():  # Если виджет активный (означает потенциальное изменение)
                if self.name[i].text():  # Если внутри виджета есть текст, то помещаем внутрь базы
                    self.data[el] = self.name[i].text()
                else:  # Если нет текста, то удаляем значение
                    self.data[el] = None
        with open(pathlib.Path(self.path_for_default, 'Настройки.txt'), 'w', encoding='utf-8-sig') as f:  # Пишем в файл
            data_insert = {"widget_settings": self.data,
                           "gui_settings":
                               {"tab_order": self.tab_order}
                           }
            json.dump(data_insert, f, ensure_ascii=False, sort_keys=True, indent=4)
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
