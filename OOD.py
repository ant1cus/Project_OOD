import datetime
import json
import os
import pathlib
import queue
import sys

from PyQt5.QtGui import QSyntaxHighlighter, QTextCharFormat, QColor

import Main
import logging
import about
from PyQt5.QtCore import QTranslator, QLocale, QLibraryInfo, QDir
from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog, QMessageBox, QDialog
from checked import (checked_zone_checked, checked_file_parcing, checked_generation_pemi,
                     checked_delete_header_footer, checked_hfe_generation, checked_hfi_generation,
                     checked_application_data, checked_lf_data, checked_generation_cc)
from rewrite_settings import rewrite
from Default import DefaultWindow
from Zone_Check import ZoneChecked
from File_Parcing import FileParcing
from Generation_Files import GenerationFile
from Delete_Header_Footer import DeleteHeaderFooter
from HFE_Generation import HFEGeneration
from HFI_Generation import HFIGeneration
from CopyApplication import GenerateCopyApplication
from LowFrequency_dispertion import LFGeneration
from ContinuousSpectrum import GenerationFileCC


class AboutWindow(QDialog, about.Ui_Dialog):  # Для отображения информации
    def __init__(self):
        super().__init__()
        self.setupUi(self)


def about():  # Открываем окно с описанием
    window_add = AboutWindow()
    window_add.exec_()


class SyntaxHighlighter(QSyntaxHighlighter):
    def __init__(self, parent):
        super(SyntaxHighlighter, self).__init__(parent)
        self._highlight_lines = {}

    def highlight_line(self, line, fmt):
        if isinstance(line, int) and line >= 0 and isinstance(fmt, QTextCharFormat):
            self._highlight_lines[line] = fmt
            tb = self.document().findBlockByNumber(line)
            self.rehighlightBlock(tb)

    def clear_highlight(self):
        self._highlight_lines = {}
        self.rehighlight()

    def highlightBlock(self, text):
        line = self.currentBlock().blockNumber()
        fmt = self._highlight_lines.get(line)
        if fmt is not None:
            self.setFormat(0, len(text), fmt)


def highlighter(plain_text_edit):
    _highlighter = SyntaxHighlighter(plain_text_edit.document())
    fmt = QTextCharFormat()
    fmt.setBackground(QColor("#E1E1E1"))
    _highlighter.clear_highlight()
    for i in range(len(plain_text_edit.toPlainText().split('\n'))):
        if i % 2 == 0:
            _highlighter.highlight_line(i, fmt)


class MainWindow(QMainWindow, Main.Ui_MainWindow):  # Главное окно

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setupUi(self)
        self.queue = queue.Queue(maxsize=1)
        filename = str(datetime.date.today()) + '_logs.log'
        os.makedirs(pathlib.Path('logs'), exist_ok=True)
        filemode = 'a' if pathlib.Path('logs', filename).is_file() else 'w'
        logging.basicConfig(filename=pathlib.Path('logs', filename),
                            level=logging.DEBUG,
                            filemode=filemode,
                            format="%(asctime)s - %(levelname)s - %(funcName)s: %(lineno)d - %(message)s")
        self.pushButton_open_folder_zone_check.clicked.connect((lambda: self.browse(self.lineEdit_path_check)))
        self.pushButton_open_folder_parser.clicked.connect((lambda: self.browse(self.lineEdit_path_parser)))
        self.pushButton_open_folder_original_exctract.clicked.connect((lambda:
                                                                       self.browse(self.lineEdit_path_start_extract)))
        self.pushButton_open_folder_start_pemi.clicked.connect((lambda: self.browse(self.lineEdit_path_start_pemi)))
        self.pushButton_open_folder_finish_pemi.clicked.connect((lambda: self.browse(self.lineEdit_path_finish_pemi)))
        self.pushButton_open_file_freq_restrict.clicked.connect((lambda: self.browse(self.lineEdit_path_freq_restrict)))
        self.pushButton_open_file_HFE.clicked.connect((lambda: self.browse(self.lineEdit_path_file_HFE)))
        self.pushButton_open_file_HFI.clicked.connect((lambda: self.browse(self.lineEdit_path_file_HFI)))
        self.pushButton_open_folder_example.clicked.connect((lambda: self.browse(self.lineEdit_path_start_example)))
        self.pushButton_open_finish_folder_example.clicked.connect((lambda:
                                                                    self.browse(self.lineEdit_path_finish_example)))
        self.pushButton_open_folder_start_lf.clicked.connect((lambda: self.browse(self.lineEdit_path_start_folder_lf)))
        self.pushButton_open_folder_finish_lf.clicked.connect((lambda:
                                                               self.browse(self.lineEdit_path_finish_folder_lf)))
        self.pushButton_open_file_excel_lf.clicked.connect((lambda: self.browse(self.lineEdit_path_file_excel_lf)))
        self.pushButton_open_folder_start_cc.clicked.connect((lambda: self.browse(self.lineEdit_path_folder_start_cc)))
        self.pushButton_open_folder_finish_cc.clicked.connect((lambda:
                                                               self.browse(self.lineEdit_path_folder_finish_cc)))
        self.groupBox_FSB.clicked.connect(self.group_box_change_state)
        self.groupBox_FSTEK.clicked.connect(self.group_box_change_state)
        self.pushButton_stop.clicked.connect(self.pause_thread)
        self.pushButton_check.clicked.connect(self.checked_zone)
        self.pushButton_parser.clicked.connect(self.parcing_file)
        self.pushButton_generation_pemi.clicked.connect(self.generate_pemi)
        self.pushButton_generation_exctract.clicked.connect(self.delete_header_footer)
        self.pushButton_generation_HFE.clicked.connect(self.generate_hfe)
        self.pushButton_generation_HFI.clicked.connect(self.generate_hfi)
        self.pushButton_create_application.clicked.connect(self.copy_application)
        self.pushButton_start_insert_lf.clicked.connect(self.generate_lf)
        self.pushButton_ss_start.clicked.connect(self.generate_cc)
        self.action_settings_default.triggered.connect(self.default_settings)
        self.menu_about.aboutToShow.connect(about)
        self.action_zone_checked.triggered.connect(self.add_tab)
        self.action_parser.triggered.connect(self.add_tab)
        self.action_extract.triggered.connect(self.add_tab)
        self.action_gen_application.triggered.connect(self.add_tab)
        self.action_gen_pemi.triggered.connect(self.add_tab)
        self.action_gen_HFE.triggered.connect(self.add_tab)
        self.action_gen_HFI.triggered.connect(self.add_tab)
        self.action_gen_LF.triggered.connect(self.add_tab)
        self.action_gen_cc.triggered.connect(self.add_tab)
        self.tabWidget.tabBar().tabMoved.connect(self.tab_)
        self.tabWidget.tabBarClicked.connect(self.tab_click)
        self.tabWidget.tabCloseRequested.connect(lambda index: self.tabWidget.removeTab(index))
        self.start_index = False
        self.start_name = False
        self.default_path = pathlib.Path.cwd()  # Путь для файла настроек
        # Имена в файле
        self.name_list = {'checked-path_folder_check': ['Путь к дир. с файлами', self.lineEdit_path_check],
                          'checked-table_number': ['Номер таблицы', self.lineEdit_table_number],
                          'checked-checkBox_first_table': ['Только 1 таб.', self.checkBox_first_table],
                          'checked-groupBox_FSB': ['Проверка ФСБ', self.groupBox_FSB],
                          'checked-checkBox_win_lin': ['Windows + Linux', self.checkBox_win_lin],
                          'checked-stationary_FSB': ['Стац. ФСБ', self.lineEdit_stationary_FSB],
                          'checked-carry_FSB': ['Воз. ФСБ', self.lineEdit_carry_FSB],
                          'checked-wear_FSB': ['Нос. ФСБ', self.lineEdit_wear_FSB],
                          'checked-r1_FSB': ['r1 ФСБ', self.lineEdit_r1_FSB],
                          'checked-r1s_FSB': ['r1` ФСБ', self.lineEdit_r1s_FSB],
                          'checked-groupBox_FSTEK': ['Проверка ФСТЭК', self.groupBox_FSTEK],
                          'checked-stationary_FSTEK': ['Стац. ФСТЭК', self.lineEdit_stationary_FSTEK],
                          'checked-carry_FSTEK': ['Воз. ФСТЭК', self.lineEdit_carry_FSTEK],
                          'checked-wear_FSTEK': ['Нос. ФСТЭК', self.lineEdit_wear_FSTEK],
                          'checked-r1_FSTEK': ['r1 ФСТЭК', self.lineEdit_r1_FSTEK],
                          'parser-path_folder_parser': ['Путь к дир. с файлами', self.lineEdit_path_parser],
                          'parser-checkBox_group_parcing': ['Пакетный парсинг', self.checkBox_group_parcing],
                          'parser-checkBox_no_freq_limit': ['Без ограничения частот', self.checkBox_no_freq_limit],
                          'extract-path_folder_start_extract': ['Путь к дир. с файлами',
                                                                self.lineEdit_path_start_extract],
                          'extract-groupBox_value_for_extract': ['Значения для выписки',
                                                                 self.groupBox_value_for_extract],
                          'extract-conclusion_post': ['Должность заключение', self.lineEdit_conclusion_post],
                          'extract-conclusion_name': ['ФИО заключение', self.lineEdit_conclusion_name],
                          'extract-protocol_post': ['Должность протокол', self.lineEdit_protocol_post],
                          'extract-protocol_name': ['ФИО протокол', self.lineEdit_protocol_name],
                          'extract-prescription_post': ['Должность предписание', self.lineEdit_prescription_post],
                          'extract-prescription_name': ['ФИО предписание', self.lineEdit_prescription_name],
                          'gen_pemi-path_folder_start': ['Путь к дир. с исходниками',
                                                         self.lineEdit_path_start_pemi],
                          'gen_pemi-path_folder_finish': ['Путь к дир. для генерации',
                                                          self.lineEdit_path_finish_pemi],
                          'gen_pemi-checkBox_freq_restrict': ['Файл ограничения частот', self.checkBox_freq_restrict],
                          'gen_pemi-path_file_freq_restrict': ['Путь к файлу ограничений',
                                                               self.lineEdit_path_freq_restrict],
                          'gen_pemi-checkBox_no_excel_generation': ['Не генерировать excel',
                                                                    self.checkBox_no_excel_generation],
                          'gen_pemi-checkBox_no_limit_freq_gen': ['Без ограничения знач.',
                                                                  self.checkBox_no_limit_freq_gen],
                          'gen_pemi-checkBox_3db_difference': ['Разница 3 дБ', self.checkBox_3db_difference],
                          'gen_pemi-complect_quant_pemi': ['Количество комплектов', self.lineEdit_complect_quant_pemi],
                          'gen_pemi-complect_number_pemi': ['Номера комплектов', self.lineEdit_complect_number_pemi],
                          'HFE-path_file_HFE': ['Путь к дир. с файлами', self.lineEdit_path_file_HFE],
                          'HFE-complect_quant_HFE': ['Количество комплектов', self.lineEdit_complect_quant_HFE],
                          'HFE-checkBox_required_values_HFE': ['Значения вручную', self.checkBox_required_values_HFE],
                          'HFE-frequency': ['Частота', self.lineEdit_frequency],
                          'HFE-level': ['Уровень', self.lineEdit_level],
                          'HFI-path_file_HFI': ['Путь к дир. с файлами', self.lineEdit_path_file_HFI],
                          'HFI-complect_quant_HFI': ['Количество комплектов', self.lineEdit_complect_quant_HFE],
                          'HFI-checkBox_imposition_freq': ['Ручной ввод частоты', self.checkBox_imposition_freq],
                          'HFI-imposition_freq': ['Частота навязывания', self.lineEdit_imposition_freq],
                          'HFI-checkBox_power_supply': ['Питание', self.checkBox_power_supply],
                          'HFI-checkBox_symetrical': ['Симметричка', self.checkBox_symetrical],
                          'HFI-checkBox_asymetriacal': ['Не симметричка', self.checkBox_asymetriacal],
                          'application-path_file_example': ['Путь к файлу', self.lineEdit_path_start_example],
                          'application-path_folder_finish_example': ['Путь к конечной дир.',
                                                                     self.lineEdit_path_finish_example],
                          'application-number_position': ['Номер позиции', self.lineEdit_number_position],
                          'application-quantity_document': ['Количество комплектов', self.lineEdit_quantity_document],
                          'LF-path_folder_start': ['Путь к начальной дир.', self.lineEdit_path_start_folder_lf],
                          'LF-path_folder_finish': ['Путь к конечной дир.', self.lineEdit_path_finish_folder_lf],
                          'LF-path_file_excel': ['Путь к файлу генератору', self.lineEdit_path_file_excel_lf],
                          'CC-path_folder_start': ['Путь к файлам спектра', self.lineEdit_path_folder_start_cc],
                          'CC-path_folder_finish': ['Путь к конечной папке', self.lineEdit_path_folder_finish_cc]}
        # Грузим значения по умолчанию
        self.name_tab = {"tab_zone_checked": "Проверка зон", "tab_parser": "Парсер txt",
                         "tab_exctract": "Обезличивание", "tab_gen_application": "Генератор приложений",
                         "tab_gen_pemi": "Генератор ПЭМИ", "tab_gen_HFE": "Генератор ВЧО",
                         "tab_gen_HFI": "Генератор ВЧН", 'tab_gen_LF': 'Генератор НЧ',
                         "tab_continuous_spectrum": 'Сплошной спектр'}
        self.name_action = {"tab_zone_checked": self.action_zone_checked, "tab_parser": self.action_parser,
                            "tab_exctract": self.action_extract, "tab_gen_application": self.action_gen_application,
                            "tab_gen_pemi": self.action_gen_pemi, "tab_gen_HFE": self.action_gen_HFE,
                            "tab_gen_HFI": self.action_gen_HFI, 'tab_gen_LF': self.action_gen_LF,
                            "tab_continuous_spectrum": self.action_gen_cc}
        try:
            with open(pathlib.Path(pathlib.Path.cwd(), 'Настройки.txt'), "r", encoding='utf-8-sig') as f:
                dict_load = json.load(f)
                self.data = dict_load['widget_settings']
                self.tab_order = dict_load['gui_settings']['tab_order']
                self.tab_visible = dict_load['gui_settings']['tab_visible']
        except FileNotFoundError:
            with open(pathlib.Path(pathlib.Path.cwd(), 'Настройки.txt'), "w", encoding='utf-8-sig') as f:
                data_insert = {"widget_settings": {},
                               "gui_settings":
                                   {"tab_order": {'0': "tab_zone_checked", '1': "tab_parser", '2': "tab_exctract",
                                                  '3': "tab_gen_application", '4': "tab_gen_pemi", '5': "tab_gen_HFE",
                                                  '6': "tab_gen_HFI", '7': "tab_gen_LF",
                                                  '8': "tab_continuous_spectrum"},
                                    "tab_visible": {"tab_zone_checked": True, "tab_parser": True, "tab_exctract": True,
                                                    "tab_gen_application": True, "tab_gen_pemi": True,
                                                    "tab_gen_HFE": True, "tab_gen_HFI": True, "tab_gen_LF": True,
                                                    "tab_continuous_spectrum": True}
                                    }
                               }
                json.dump(data_insert, f, ensure_ascii=False, sort_keys=True, indent=4)
                self.data = {}
                self.tab_order = data_insert['gui_settings']['tab_order']
                self.tab_visible = data_insert['gui_settings']['tab_visible']

        self.tab_for_paint = {}
        for tab in range(0, self.tabWidget.tabBar().count()):
            if self.tabWidget.widget(tab).objectName() not in self.tab_order.values():
                self.tab_order[str(len(self.tab_order))] = self.tabWidget.widget(tab).objectName()
                rewrite(self.default_path, self.tab_order, visible='tab_order')
                self.tab_visible[str(self.tabWidget.widget(tab).objectName())] = True
                rewrite(self.default_path, self.tab_visible, visible='tab_visible')
            self.tab_for_paint[self.tabWidget.widget(tab).objectName()] = self.tabWidget.widget(tab)
        self.tabWidget.clear()
        for tab in self.tab_order:
            if self.tab_visible[self.tab_order[tab]]:
                self.name_action[self.tab_order[tab]].setChecked(True)
                self.tabWidget.addTab(self.tab_for_paint[self.tab_order[tab]], self.name_tab[self.tab_order[tab]])
        self.tabWidget.tabBar().setCurrentIndex(0)
        # Линии для заполнения
        # self.line = [self.lineEdit_path_check, self.lineEdit_table_number, self.lineEdit_stationary_FSB,
        #              self.lineEdit_carry_FSB, self.lineEdit_wear_FSB, self.lineEdit_r1_FSB, self.lineEdit_r1s_FSB,
        #              self.lineEdit_stationary_FSTEK, self.lineEdit_carry_FSTEK, self.lineEdit_wear_FSTEK,
        #              self.lineEdit_r1_FSTEK, self.lineEdit_path_parser, self.lineEdit_path_original_extract,
        #              self.lineEdit_conclusion_post, self.lineEdit_conclusion_name, self.lineEdit_protocol_post,
        #              self.lineEdit_protocol_name, self.lineEdit_prescription_post, self.lineEdit_prescription_name,
        #              self.lineEdit_path_original_file, self.lineEdit_path_finish_folder,
        #              self.lineEdit_path_freq_restrict, self.lineEdit_complect_quant_pemi,
        #              self.lineEdit_complect_number_pemi, self.lineEdit_path_file_HFE,
        #              self.lineEdit_complect_quant_HFE, self.lineEdit_frequency, self.lineEdit_level,
        #              self.lineEdit_path_file_HFI, self.lineEdit_complect_quant_HFI, self.lineEdit_imposition_freq,
        #              self.lineEdit_path_example, self.lineEdit_path_finish_folder_example,
        #              self.lineEdit_number_position, self.lineEdit_quantity_document,
        #              self.lineEdit_path_start_folder_lf, self.lineEdit_path_finish_folder_lf,
        #              self.lineEdit_path_file_excel_lf]
        self.default_date(self.data)

    def tab_(self, index):
        # print(self.tabWidget.currentIndex(), self.tabWidget.currentWidget().objectName())
        for tab in self.tab_order.items():
            if tab[1] == self.start_name and tab[1] == self.tabWidget.currentWidget().objectName():
                self.tab_order[str(index)], self.tab_order[tab[0]] = self.tab_order[tab[0]], self.tab_order[str(index)]
                break
            elif tab[1] == self.tabWidget.currentWidget().objectName():
                self.tab_order[str(index)], self.tab_order[tab[0]] = self.tab_order[tab[0]], self.tab_order[str(index)]
                break
        rewrite(self.default_path, self.tab_order, order='tab_order')

    def tab_click(self, index):
        self.start_name = self.tab_order[str(index)]

    def add_tab(self):
        name_open_tab = {self.tabWidget.widget(ind).objectName(): ind for ind
                         in range(0, self.tabWidget.tabBar().count())}
        for el in self.name_tab:
            if self.name_tab[el] == self.sender().text():
                if self.name_action[el].isChecked():
                    if el not in name_open_tab:
                        self.tabWidget.addTab(self.tab_for_paint[el], self.name_tab[el])
                    if self.tab_visible[el] is False:
                        self.tab_visible[el] = True
                        rewrite(self.default_path, self.tab_visible, visible='tab_visible')
                else:
                    if self.tab_visible[el]:
                        self.tab_visible[el] = False
                        rewrite(self.default_path, self.tab_visible, visible='tab_visible')

    def default_date(self, incoming_data):
        for element in self.name_list:
            if element in incoming_data:
                if 'checkBox' in element or 'groupBox' in element:
                    self.name_list[element][1].setChecked(True) if incoming_data[element] \
                        else self.name_list[element][1].setChecked(False)
                else:
                    self.name_list[element][1].setText(incoming_data[element])
                # self.line[i].setText(d[el])  # Помещаем значение

    def default_settings(self):  # Запускаем окно с настройками по умолчанию.
        self.close()
        window_add = DefaultWindow(self, self.default_path, self.name_list)
        window_add.show()

    def group_box_change_state(self, state):
        if self.sender() == self.groupBox_FSTEK and state:
            self.groupBox_FSTEK.setChecked(True)
            self.groupBox_FSB.setChecked(False)
        elif self.sender() == self.groupBox_FSTEK and state is False:
            self.groupBox_FSTEK.setChecked(False)
        elif self.sender() == self.groupBox_FSB and state:
            self.groupBox_FSTEK.setChecked(False)
            self.groupBox_FSB.setChecked(True)
        elif self.sender() == self.groupBox_FSB and state is False:
            self.groupBox_FSB.setChecked(False)

    def browse(self, line_edit):  # Для кнопки открыть
        directory = None
        if 'folder' in self.sender().objectName():  # Если необходимо открыть директорию
            directory = QFileDialog.getExistingDirectory(self, "Открыть папку", QDir.currentPath())
        else:  # Если необходимо открыть файл
            directory = QFileDialog.getOpenFileName(self, "Открыть", QDir.currentPath())
        if directory and isinstance(directory, tuple):
            if directory[0]:
                line_edit.setText(directory[0])
        elif directory and isinstance(directory, str):
            line_edit.setText(directory)

    def copy_application(self):
        application = checked_application_data(self.lineEdit_path_start_example, self.lineEdit_path_finish_example,
                                               self.lineEdit_number_position, self.lineEdit_quantity_document)
        if isinstance(application, list):
            self.on_message_changed(application[0], application[1])
            return
        # Если всё прошло запускаем поток
        application['logging'], application['queue'] = logging, self.queue
        application['default_path'] = self.default_path
        self.thread = GenerateCopyApplication(application)
        self.thread.progress.connect(self.progressBar.setValue)
        self.thread.status.connect(self.statusBar().showMessage)
        self.thread.messageChanged.connect(self.on_message_changed)
        self.thread.errors.connect(self.errors)
        self.thread.start()

    def generate_pemi(self):
        generate = checked_generation_pemi(self.lineEdit_path_start_pemi, self.lineEdit_path_finish_pemi,
                                           self.lineEdit_complect_number_pemi, self.lineEdit_complect_quant_pemi,
                                           self.checkBox_freq_restrict.isChecked(), self.lineEdit_path_freq_restrict)
        no_freq_lim = self.checkBox_no_limit_freq_gen.isChecked()
        no_excel_file = self.checkBox_no_excel_generation.isChecked()
        db_differeence = self.checkBox_3db_difference.isChecked()
        if isinstance(generate, list):
            self.on_message_changed(generate[0], generate[1])
            return
        # Если всё прошло запускаем поток
        generate['logging'], generate['queue'] = logging, self.queue
        generate['no_freq_lim'], generate['no_excel_file'] = no_freq_lim, no_excel_file
        generate['3db_difference'], generate['default_path'] = db_differeence, self.default_path
        self.thread = GenerationFile(generate)
        self.thread.progress.connect(self.progressBar.setValue)
        self.thread.status.connect(self.statusBar().showMessage)
        self.thread.messageChanged.connect(self.on_message_changed)
        self.thread.errors.connect(self.errors)
        self.thread.start()

    def generate_hfe(self):
        generate = checked_hfe_generation(self.lineEdit_path_file_HFE, self.lineEdit_complect_quant_HFE,
                                          self.checkBox_required_values_HFE, self.lineEdit_frequency,
                                          self.lineEdit_level)
        if isinstance(generate, list):
            self.on_message_changed(generate[0], generate[1])
            return
        # Если всё прошло запускаем поток
        generate['logging'], generate['queue'] = logging, self.queue
        generate['default_path'] = self.default_path
        self.thread = HFEGeneration(generate)
        self.thread.progress.connect(self.progressBar.setValue)
        self.thread.status.connect(self.statusBar().showMessage)
        self.thread.messageChanged.connect(self.on_message_changed)
        self.thread.errors.connect(self.errors)
        self.thread.start()

    def generate_hfi(self):
        generate = checked_hfi_generation(self.lineEdit_path_file_HFI, self.lineEdit_imposition_freq,
                                          self.lineEdit_complect_quant_HFI,
                                          [self.checkBox_power_supply.isChecked(),
                                           self.checkBox_symetrical.isChecked(),
                                           self.checkBox_asymetriacal.isChecked()])
        if isinstance(generate, list):
            self.on_message_changed(generate[0], generate[1])
            return
        # Если всё прошло запускаем поток
        generate['logging'], generate['queue'] = logging, self.queue
        generate['default_path'] = self.default_path
        self.thread = HFIGeneration(generate)
        self.thread.progress.connect(self.progressBar.setValue)
        self.thread.status.connect(self.statusBar().showMessage)
        self.thread.messageChanged.connect(self.on_message_changed)
        self.thread.errors.connect(self.errors)
        self.thread.start()

    def generate_lf(self):
        generate = checked_lf_data(self.lineEdit_path_start_folder_lf, self.lineEdit_path_finish_folder_lf,
                                   self.lineEdit_path_file_excel_lf)
        if isinstance(generate, list):
            self.on_message_changed(generate[0], generate[1])
            return
        # Если всё прошло запускаем поток
        generate['logging'], generate['queue'] = logging, self.queue
        generate['default_path'] = self.default_path
        self.thread = LFGeneration(generate)
        self.thread.progress.connect(self.progressBar.setValue)
        self.thread.status.connect(self.statusBar().showMessage)
        self.thread.messageChanged.connect(self.on_message_changed)
        self.thread.errors.connect(self.errors)
        self.thread.start()

    def parcing_file(self):
        self.plainTextEdit_succsess_order.clear()
        self.groupBox_succsess_order.setStyleSheet("")
        self.plainTextEdit_error_order.clear()
        self.groupBox_error_order.setStyleSheet("")
        self.plainTextEdit_errors.clear()
        self.groupBox_errors.setStyleSheet("")
        group_file = self.checkBox_group_parcing.isChecked()
        no_freq_lim = self.checkBox_no_freq_limit.isChecked()
        folder = checked_file_parcing(self.lineEdit_path_parser, group_file)
        if isinstance(folder, list):
            self.on_message_changed(folder[0], folder[1])
            return
        # Если всё прошло запускаем поток
        folder['group_file'], folder['no_freq_lim'] = group_file, no_freq_lim
        folder['logging'], folder['queue'] = logging, self.queue
        folder['default_path'] = self.default_path
        self.thread = FileParcing(folder)
        self.thread.progress.connect(self.progressBar.setValue)
        self.thread.status.connect(self.statusBar().showMessage)
        self.thread.errors.connect(self.errors)
        self.thread.messageChanged.connect(self.on_message_changed)
        self.thread.start()

    def errors(self):
        text = self.queue.get_nowait()
        if 'Прошедшие заказы:' in text:
            self.groupBox_succsess_order.setStyleSheet('''
            QGroupBox {border: 0.5px solid;
            border-radius: 5px;
             border-color: green;
              padding:10px 0px 0px 0px;}''')
            self.plainTextEdit_succsess_order.insertPlainText(text['Прошедшие заказы:'])
        elif 'Заказы с ошибками:' in text:
            self.groupBox_error_order.setStyleSheet('''
            QGroupBox {border: 0.5px solid;
            border-radius: 5px;
             border-color: red;
              padding:10px 0px 0px 0px;}''')
            self.plainTextEdit_error_order.insertPlainText(text['Заказы с ошибками:'])
        elif 'errors' in text:
            self.groupBox_errors.setStyleSheet('''
            QGroupBox {border: 0.5px solid;
            border-radius: 5px;
             border-color: yellow;
              padding:10px 0px 0px 0px;}''')
            self.plainTextEdit_errors.insertPlainText('\n'.join(text['errors']))
            highlighter(self.plainTextEdit_errors)
        elif 'errors_gen' in text:
            self.on_message_changed('УПС!', '\n'.join(text['errors_gen']))

    def checked_zone(self):
        department = self.groupBox_FSB.isChecked()
        win_lin = self.checkBox_win_lin.isChecked()
        one_table = self.checkBox_first_table.isChecked()
        zone = [self.lineEdit_stationary_FSB, self.lineEdit_carry_FSB, self.lineEdit_wear_FSB, self.lineEdit_r1_FSB,
                self.lineEdit_r1s_FSB] \
            if department else [self.lineEdit_stationary_FSTEK, self.lineEdit_carry_FSTEK,
                                self.lineEdit_wear_FSTEK, self.lineEdit_r1_FSTEK]
        zone_all = checked_zone_checked(self.lineEdit_path_check, self.lineEdit_table_number, zone)
        if isinstance(zone_all, list):
            self.on_message_changed(zone_all[0], zone_all[1])
            return
        # Если всё прошло запускаем поток
        if self.checkBox_win_lin.isChecked():
            zone = {i + 5: zone_all[i] for i in zone_all}
            zone_all = {**zone_all, **zone}
        zone = {'path_check': self.lineEdit_path_check.text().strip(),
                'table_number': self.lineEdit_table_number.text().strip(), 'department': department, 'win_lin': win_lin,
                'zone_all': zone_all, 'one_table': one_table, 'logging': logging, 'queue': self.queue,
                'default_path': self.default_path}
        self.thread = ZoneChecked(zone)
        self.thread.progress.connect(self.progressBar.setValue)
        self.thread.status.connect(self.statusBar().showMessage)
        self.thread.messageChanged.connect(self.on_message_changed)
        self.thread.start()

    def delete_header_footer(self):
        output = checked_delete_header_footer(self.lineEdit_path_start_extract, self.lineEdit_conclusion_post,
                                              self.lineEdit_conclusion_name, self.lineEdit_protocol_post,
                                              self.lineEdit_protocol_name, self.lineEdit_prescription_post,
                                              self.lineEdit_prescription_name)
        if isinstance(output, list):
            self.on_message_changed(output[0], output[1])
            return
        # Если всё прошло запускаем поток
        output['logging'], output['queue'] = logging, self.queue
        output['default_path'] = self.default_path
        self.thread = DeleteHeaderFooter(output)
        self.thread.progress.connect(self.progressBar.setValue)
        self.thread.status.connect(self.statusBar().showMessage)
        self.thread.messageChanged.connect(self.on_message_changed)
        self.thread.start()

    def generate_cc(self):
        generate = checked_generation_cc(self.lineEdit_path_folder_start_cc, self.lineEdit_path_folder_finish_cc,
                                         self.lineEdit_complect_number_cc, self.lineEdit_complect_quantity_cc)
        if isinstance(generate, list):
            self.on_message_changed(generate[0], generate[1])
            return
        # Если всё прошло запускаем поток
        generate['logging'], generate['queue'] = logging, self.queue
        generate['default_path'] = self.default_path
        self.thread = GenerationFileCC(generate)
        self.thread.progress.connect(self.progressBar.setValue)
        self.thread.status.connect(self.statusBar().showMessage)
        self.thread.messageChanged.connect(self.on_message_changed)
        self.thread.errors.connect(self.errors)
        self.thread.start()

    def pause_thread(self):
        if self.queue.empty():
            self.statusBar().showMessage(self.statusBar().currentMessage() + ' (прерывание процесса, подождите...)')
            self.queue.put(True)

    def on_message_changed(self, title, description):
        if title == 'УПС!':
            QMessageBox.critical(self, title, description)
        elif title == 'Внимание!':
            QMessageBox.warning(self, title, description)
        elif title == 'Вопрос?':
            self.statusBar().clearMessage()
            ans = QMessageBox.question(self, title, description, QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if ans == QMessageBox.No:
                self.thread.queue.put(True)
            else:
                self.thread.queue.put(False)
            self.thread.event.set()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    translator = QTranslator(app)
    locale = QLocale.system().name()
    path = QLibraryInfo.location(QLibraryInfo.TranslationsPath)
    translator.load('qtbase_%s' % locale.partition('_')[0], path)
    app.installTranslator(translator)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
