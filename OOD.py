import datetime
import json
import os
import pathlib
import queue
import sys
import traceback

from PyQt5.QtGui import QSyntaxHighlighter, QTextCharFormat, QColor, QIcon

import Main
import logging
import about
from PyQt5.QtCore import QTranslator, QLocale, QLibraryInfo, QDir
from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog, QMessageBox, QDialog
from checked import (checked_zone_checked, checked_file_parcing, checked_generation_pemi,
                     checked_delete_header_footer, checked_hfe_generation, checked_hfi_generation,
                     checked_application_data, checked_lf_data, checked_generation_cc, checked_number_instance,
                     checked_find_files, checked_lf_pemi)
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
from Number_Instance import ChangeNumberInstance
from Find_Files import FindingFiles


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
        self.actual_version = '3.2.1'
        self.queue = queue.Queue(maxsize=1)
        self.pushButton_open_folder_zone_check.clicked.connect((lambda: self.browse(self.lineEdit_path_check)))
        self.pushButton_open_folder_parser.clicked.connect((lambda: self.browse(self.lineEdit_path_parser)))
        self.pushButton_open_folder_original_exctract.clicked.connect((lambda:
                                                                       self.browse(self.lineEdit_path_start_extract)))
        self.pushButton_open_folder_start_pemi.clicked.connect((lambda: self.browse(self.lineEdit_path_start_pemi)))
        self.pushButton_open_folder_finish_pemi.clicked.connect((lambda: self.browse(self.lineEdit_path_finish_pemi)))
        self.pushButton_open_file_freq_restrict.clicked.connect((lambda: self.browse(self.lineEdit_path_freq_restrict)))
        self.pushButton_open_folder_HFE.clicked.connect((lambda: self.browse(self.lineEdit_path_folder_HFE)))
        self.pushButton_open_folder_HFI.clicked.connect((lambda: self.browse(self.lineEdit_path_folder_HFI)))
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
        self.pushButton_open_folder_old_number_instance.clicked.connect(
            (lambda: self.browse(self.lineEdit_path_folder_old_number_instance)))
        self.pushButton_open_folder_new_number_instance.clicked.connect(
            (lambda: self.browse(self.lineEdit_path_folder_new_number_instance)))
        self.pushButton_open_file_unloading_find.clicked.connect((lambda:
                                                                  self.browse(self.lineEdit_path_file_unloading_find)))
        self.pushButton_open_folder_start_find.clicked.connect((lambda:
                                                                self.browse(self.lineEdit_path_folder_start_find)))
        self.pushButton_open_folder_finish_find.clicked.connect((lambda:
                                                                 self.browse(self.lineEdit_path_folder_finish_find)))
        self.pushButton_open_folder_start_lf_pemi.clicked.connect((lambda:
                                                                   self.browse(self.lineEdit_path_folder_start_lf_pemi)))
        self.pushButton_open_folder_finish_lf_pemi.clicked.connect((lambda:
                                                                    self.browse(self.lineEdit_path_folder_finish_lf_pemi)))
        self.groupBox_FSB.clicked.connect(self.group_box_change_state)
        self.groupBox_FSTEK.clicked.connect(self.group_box_change_state)
        self.pushButton_check.clicked.connect(self.checked_zone)
        self.pushButton_parser.clicked.connect(self.parcing_file)
        self.pushButton_generation_pemi.clicked.connect(self.generate_pemi)
        self.pushButton_generation_exctract.clicked.connect(self.delete_header_footer)
        self.pushButton_generation_HFE.clicked.connect(self.generate_hfe)
        self.pushButton_generation_HFI.clicked.connect(self.generate_hfi)
        self.pushButton_create_application.clicked.connect(self.copy_application)
        self.pushButton_start_insert_lf.clicked.connect(self.generate_lf)
        self.pushButton_ss_start.clicked.connect(self.generate_cc)
        self.pushButton_number_instance.clicked.connect(self.change_number_instance)
        self.pushButton_start_find.clicked.connect(self.finding_files)
        self.pushButton_start_lf_pemi.clicked.connect(self.gen_lf_pemi)
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
        self.action_number_instance.triggered.connect(self.add_tab)
        self.action_finding_file.triggered.connect(self.add_tab)
        self.action_gen_lf_pemi.triggered.connect(self.add_tab)
        self.tabWidget.tabBar().tabMoved.connect(self.tab_)
        self.tabWidget.tabBarClicked.connect(self.tab_click)
        self.tabWidget.tabCloseRequested.connect(lambda index: self.tabWidget.removeTab(index))
        self.start_index = False
        self.start_name = False
        self.default_path = pathlib.Path.cwd()  # Путь для файла настроек
        self.setWindowIcon(QIcon(str(pathlib.Path(self.default_path, 'icons', 'logo.png'))))
        # Имена в файле
        self.name_list = {'checked-path_folder_check': ['Папка с файлами', self.lineEdit_path_check],
                          'checked-table_number': ['Номер таблицы', self.lineEdit_table_number],
                          'checked-checkBox_first_table': ['Только 1 таб.', self.checkBox_first_table],
                          'checked-groupBox_FSB': ['Проверка ФСБ', self.groupBox_FSB],
                          'checked-checkBox_win_lin': ['Windows + Linux', self.checkBox_win_lin],
                          'checked-checkBox_extend_report': ['Отчёт по режимам', self.checkBox_extend_report],
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
                          'parser-path_folder_parser': ['Папка с файлами', self.lineEdit_path_parser],
                          'parser-checkBox_group_parcing': ['Пакетный парсинг', self.checkBox_group_parcing],
                          'parser-checkBox_no_freq_limit': ['Без ограничения частот', self.checkBox_no_freq_limit],
                          'parser-checkBox_12_sectors': ['12 секторов', self.checkBox_12_sectors],
                          'extract-path_folder_start_extract': ['Папка с файлами',
                                                                self.lineEdit_path_start_extract],
                          'extract-conclusion': ['Заключение', self.lineEdit_conclusion],
                          'extract-protocol': ['Протокол', self.lineEdit_protocol],
                          'extract-prescription': ['Предписание', self.lineEdit_prescription],
                          'extract-checkBox_project_prescription': ['Выписка из проекта предписания',
                                                                    self.checkBox_project_prescription],
                          'extract-checkBox_director': ['Директор', self.checkBox_director],
                          'extract-old_director': ['Директор кого ищем', self.lineEdit_old_director],
                          'extract-new_director': ['Директор на кого меняем', self.lineEdit_new_director],
                          'extract-checkBox_margin': ['Отступы', self.checkBox_margin],
                          'extract-doubleSpinBox_left': ['Левый', self.doubleSpinBox_left_margin],
                          'extract-doubleSpinBox_top': ['Верхний', self.doubleSpinBox_top_margin],
                          'extract-doubleSpinBox_right': ['Правый', self.doubleSpinBox_right_margin],
                          'extract-doubleSpinBox_bottom': ['Нижний', self.doubleSpinBox_bottom_margin],
                          'gen_pemi-path_folder_start': ['Папка с исходниками',
                                                         self.lineEdit_path_start_pemi],
                          'gen_pemi-path_folder_finish': ['Папка для генерации',
                                                          self.lineEdit_path_finish_pemi],
                          'gen_pemi-checkBox_freq_restrict': ['Файл ограничения частот', self.checkBox_freq_restrict],
                          'gen_pemi-path_file_freq_restrict': ['Файл ограничений',
                                                               self.lineEdit_path_freq_restrict],
                          'gen_pemi-checkBox_no_excel_generation': ['Не генерировать excel',
                                                                    self.checkBox_no_excel_generation],
                          'gen_pemi-checkBox_no_limit_freq_gen': ['Без ограничения знач.',
                                                                  self.checkBox_no_limit_freq_gen],
                          'gen_pemi-checkBox_3db_difference': ['Разница 3 дБ', self.checkBox_3db_difference],
                          'gen_pemi-set_quant_pemi': ['Количество комплектов', self.lineEdit_complect_quant_pemi],
                          'gen_pemi-set_number_pemi': ['Номера комплектов', self.lineEdit_complect_number_pemi],
                          'HFE-path_folder_HFE': ['Папка с файлами', self.lineEdit_path_folder_HFE],
                          'HFE-set_quant_HFE': ['Количество комплектов', self.lineEdit_complect_quant_HFE],
                          'HFE-groupBox_required_values_HFE': ['Значения вручную', self.groupBox_required_values_HFE],
                          'HFE-frequency': ['Частота', self.lineEdit_frequency],
                          'HFE-level': ['Уровень', self.lineEdit_level],
                          'HFI-path_folder_HFI': ['Папка с файлами', self.lineEdit_path_folder_HFI],
                          'HFI-set_quant_HFI': ['Количество комплектов', self.lineEdit_complect_quant_HFI],
                          'HFI-checkBox_imposition_freq': ['Ручной ввод частоты', self.checkBox_imposition_freq],
                          'HFI-imposition_freq': ['Частота навязывания', self.lineEdit_imposition_freq],
                          'HFI-checkBox_power_supply': ['Питание', self.checkBox_power_supply],
                          'HFI-checkBox_symmetrical': ['Симметричка', self.checkBox_symetrical],
                          'HFI-checkBox_asymmetriacal': ['Несимметричка', self.checkBox_asymetriacal],
                          'application-path_file_example': ['Файл', self.lineEdit_path_start_example],
                          'application-path_folder_finish_example': ['Конечная папка',
                                                                     self.lineEdit_path_finish_example],
                          'application-number_position': ['Номер позиции', self.lineEdit_number_position],
                          'application-quantity_document': ['Количество комплектов', self.lineEdit_quantity_document],
                          'LF-path_folder_start': ['Начальная папка', self.lineEdit_path_start_folder_lf],
                          'LF-path_folder_finish': ['Конечная папка', self.lineEdit_path_finish_folder_lf],
                          'LF-path_file_excel': ['Файл генератору', self.lineEdit_path_file_excel_lf],
                          'CC-path_folder_start': ['Файлы спектра', self.lineEdit_path_folder_start_cc],
                          'CC-path_folder_finish': ['Конечная папка', self.lineEdit_path_folder_finish_cc],
                          'CC-checkBox_cc_frequency': ['Конечная частота', self.checkBox_cc_frequency],
                          'CC-set_frequency': ['Конечная частота (МГц)', self.lineEdit_frequency_cc],
                          'CC-set_numbers': ['Номера комплектов', self.lineEdit_set_number_cc],
                          'CC-checkBox_cc_txt': ['Генерировать только txt', self.checkBox_cc_txt],
                          'CC-checkBox_cc_dispersion': ['Включить разброс', self.checkBox_cc_dispersion],
                          'CC-dispersion': ['Разброс значений (%)', self.lineEdit_cc_dispersion],
                          'NI-path_folder_old_number_instance': ['Начальная папка',
                                                                 self.lineEdit_path_folder_old_number_instance],
                          'NI-path_folder_new_number_instance': ['Конечная папка',
                                                                 self.lineEdit_path_folder_new_number_instance],
                          'NI-number_instance': ['Номера экземпляров', self.lineEdit_number_instance],
                          'FF-path_unloading_file': ['Файл выгрузки', self.lineEdit_path_file_unloading_find],
                          'FF-path_start_folder': ['Папка с файлами', self.lineEdit_path_folder_start_find],
                          'FF-path_finish_folder': ['Конечная папка', self.lineEdit_path_folder_finish_find],
                          'lf_pemi-path_folder_start_lf_pemi': ['Папка с исходниками',
                                                                self.lineEdit_path_folder_start_lf_pemi],
                          'lf_pemi-path_folder_finish_lf_pemi': ['Конечная папка',
                                                                 self.lineEdit_path_folder_finish_lf_pemi],
                          'lf_pemi-set_number': ['Номера комплектов', self.lineEdit_set_number_lf_pemi],
                          'lf_pemi-checkBox_lf_pemi_values_spread': ['Включить разброс',
                                                                     self.checkBox_lf_pemi_values_spread],
                          'lf_pemi-values_spread': ['Разброс значений', self.lineEdit_values_spread_lf_pemi],
                          }
        # Грузим значения по умолчанию
        self.name_tab = {"tab_zone_checked": "Проверка зон", "tab_parser": "Парсер txt",
                         "tab_exctract": "Обезличивание", "tab_gen_application": "Генератор приложений",
                         "tab_gen_pemi": "Генератор ПЭМИ", "tab_gen_HFE": "Генератор ВЧО",
                         "tab_gen_HFI": "Генератор ВЧН", "tab_gen_LF": "Генератор НЧ",
                         "tab_continuous_spectrum": "Сплошной спектр", "tab_number_instance": "Номера экземпляра",
                         "tab_finding_files": "Поиск файлов", "tab_gen_lf_pemi": "Генератор НЧ ПЭМИ"}
        self.name_action = {"tab_zone_checked": self.action_zone_checked, "tab_parser": self.action_parser,
                            "tab_exctract": self.action_extract, "tab_gen_application": self.action_gen_application,
                            "tab_gen_pemi": self.action_gen_pemi, "tab_gen_HFE": self.action_gen_HFE,
                            "tab_gen_HFI": self.action_gen_HFI, 'tab_gen_LF': self.action_gen_LF,
                            "tab_continuous_spectrum": self.action_gen_cc,
                            "tab_number_instance": self.action_number_instance,
                            "tab_finding_files": self.action_finding_file,
                            "tab_gen_lf_pemi": self.action_gen_lf_pemi}
        try:
            with open(pathlib.Path(pathlib.Path.cwd(), 'Настройки.txt'), "r", encoding='utf-8-sig') as f:
                dict_load = json.load(f)
                self.data = dict_load['widget_settings']
                self.tab_order = dict_load['gui_settings']['tab_order']
                self.tab_visible = dict_load['gui_settings']['tab_visible']
                if 'version' not in dict_load.keys():
                    dict_load["version"] = {'actual_version': ''}
                self.version = dict_load["version"]['actual_version']

        except FileNotFoundError:
            with open(pathlib.Path(pathlib.Path.cwd(), 'Настройки.txt'), "w", encoding='utf-8-sig') as f:
                dict_load = {"widget_settings": {},
                             "gui_settings":
                                 {"tab_order": {'0': "tab_zone_checked", '1': "tab_parser", '2': "tab_exctract",
                                                '3': "tab_gen_application", '4': "tab_gen_pemi", '5': "tab_gen_HFE",
                                                '6': "tab_gen_HFI", '7': "tab_gen_LF",
                                                '8': "tab_continuous_spectrum", '9': "tab_number_instance",
                                                '10': "tab_finding_files", '11': "tab_gen_lf_pemi"},
                                  "tab_visible": {"tab_zone_checked": True, "tab_parser": True, "tab_exctract": True,
                                                  "tab_gen_application": True, "tab_gen_pemi": True,
                                                  "tab_gen_HFE": True, "tab_gen_HFI": True, "tab_gen_LF": True,
                                                  "tab_continuous_spectrum": True, "tab_number_instance": True,
                                                  "tab_finding_files": True, "tab_gen_lf_pemi": True}
                                  },
                             "version": {"actual_version": ''}
                             }
                json.dump(dict_load, f, ensure_ascii=False, sort_keys=True, indent=4)
                self.data = dict_load['widget_settings']
                self.tab_order = dict_load['gui_settings']['tab_order']
                self.tab_visible = dict_load['gui_settings']['tab_visible']
                self.version = dict_load['version']['actual_version']

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
        self.default_date(self.data)
        if self.actual_version != self.version:
            dict_load['version'] = {'actual_version': self.actual_version}
            rewrite(self.default_path, dict_load, widget=True)
            about()
        # Для каждого потока свой лог. Потом сливаем в один и удаляем
        self.logging_dict = {}
        # Для сдвига окна при появлении
        self.thread_dict = {'zone_checked': {}, 'continuous_spectrum': {}, 'delete_header_footer': {},
                            'copy_application': {}, 'change_number_instance': {}, 'generate_lf': {},
                            'generate_hfi': {}, 'generate_hfe': {}, 'finding_files': {}, 'generate_pemi': {},
                            'parcing_file': {}, 'gen_lf_pemi': {}}

    def logging_file(self, name):
        filename_now = str(datetime.datetime.today().timestamp()) + '_logs.log'
        filename_all = str(datetime.date.today()) + '_logs.log'
        os.makedirs(pathlib.Path('logs', name), exist_ok=True)
        self.logging_dict[filename_now] = logging.getLogger(filename_now)
        self.logging_dict[filename_now].setLevel(logging.DEBUG)
        name_log = logging.FileHandler(pathlib.Path('logs', name, filename_now))
        basic_format = logging.Formatter("%(asctime)s - %(levelname)s - %(funcName)s: %(lineno)d - %(message)s")
        name_log.setFormatter(basic_format)
        self.logging_dict[filename_now].addHandler(name_log)
        return [filename_now, filename_all]

    def finished_thread(self, method, thread='', name_all='', name_now=''):
        if thread:
            file_all = pathlib.Path('logs', method, self.thread_dict[method][thread]['filename_all'])
            file_now = pathlib.Path('logs', method, self.thread_dict[method][thread]['filename_now'])
        else:
            file_all, file_now = pathlib.Path(name_all), pathlib.Path(name_now)
        filemode = 'a' if file_all.is_file() else 'w'
        with open(file_now, mode='r') as f:
            file_data = f.readlines()
        logging.shutdown()
        os.remove(file_now)
        self.logging_dict.pop(file_now.name)
        with open(file_all, mode=filemode) as f:
            f.write(''.join(file_data))
        if thread:
            self.thread_dict[method].pop(thread, None)

    def tab_(self, index):
        for tab in self.tab_order.items():
            if tab[1] == self.start_name and tab[1] == self.tabWidget.currentWidget().objectName():
                self.tab_order[str(index)], self.tab_order[tab[0]] = self.tab_order[tab[0]], self.tab_order[str(index)]
                break
            elif tab[1] == self.tabWidget.currentWidget().objectName():
                self.tab_order[str(index)], self.tab_order[tab[0]] = self.tab_order[tab[0]], self.tab_order[str(index)]
                break
        rewrite(self.default_path, self.tab_order, order='tab_order')

    def tab_click(self, index):
        try:
            self.start_name = self.tab_order[str(index)]
        except KeyError:
            pass

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
                elif 'doubleSpinBox' in element:
                    self.name_list[element][1].setValue(float(incoming_data[element]))
                else:
                    self.name_list[element][1].setText(incoming_data[element])

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
        file_name = self.logging_file('copy_application')
        self.logging_dict[file_name[0]].info('----------------Запускаем copy_application----------------')
        self.logging_dict[file_name[0]].info('Проверка данных')
        try:
            application = checked_application_data(self.lineEdit_path_start_example, self.lineEdit_path_finish_example,
                                                   self.lineEdit_number_position, self.lineEdit_quantity_document)
            if isinstance(application, list):
                self.logging_dict[file_name[0]].warning(application[1])
                self.logging_dict[file_name[0]].warning('Ошибки в заполнении формы, программа не запущена в работу')
                self.on_message_changed(application[0], application[1])
                self.finished_thread('copy_application',
                                     name_all=str(pathlib.Path('logs', 'copy_application', file_name[1])),
                                     name_now=str(pathlib.Path('logs', 'copy_application', file_name[0])))
                return
            # Если всё прошло запускаем поток
            self.logging_dict[file_name[0]].info('Запуск на выполнение')
            application['logging'], application['queue'] = self.logging_dict[file_name[0]], self.queue
            application['default_path'] = self.default_path
            application['move'] = len(self.thread_dict['copy_application'])
            self.thread = GenerateCopyApplication(application)
            self.thread.status_finish.connect(self.finished_thread)
            self.thread.status.connect(self.statusBar().showMessage)
            self.thread.start()
            self.thread_dict['copy_application'][str(self.thread)] = {'filename_all': file_name[1],
                                                                      'filename_now': file_name[0]}
        except BaseException as exception:
            self.logging_dict[file_name[0]].error('Ошибка при старте copy_application')
            self.logging_dict[file_name[0]].error(exception)
            self.logging_dict[file_name[0]].error(traceback.format_exc())
            self.on_message_changed('УПС!', 'Неизвестная ошибка, обратитесь к разработчику')
            self.finished_thread('copy_application',
                                 name_all=str(pathlib.Path('logs', 'copy_application', file_name[1])),
                                 name_now=str(pathlib.Path('logs', 'copy_application', file_name[0])))
            return

    def generate_pemi(self):
        file_name = self.logging_file('generate_pemi')
        self.logging_dict[file_name[0]].info('----------------Запускаем generate_pemi----------------')
        self.logging_dict[file_name[0]].info('Проверка данных')
        try:
            generate = checked_generation_pemi(self.lineEdit_path_start_pemi, self.lineEdit_path_finish_pemi,
                                               self.lineEdit_complect_number_pemi, self.lineEdit_complect_quant_pemi,
                                               self.checkBox_freq_restrict.isChecked(),
                                               self.lineEdit_path_freq_restrict)
            no_freq_lim = self.checkBox_no_limit_freq_gen.isChecked()
            no_excel_file = self.checkBox_no_excel_generation.isChecked()
            db_difference = self.checkBox_3db_difference.isChecked()
            if isinstance(generate, list):
                self.logging_dict[file_name[0]].warning(generate[1])
                self.logging_dict[file_name[0]].warning('Ошибки в заполнении формы, программа не запущена в работу')
                self.on_message_changed(generate[0], generate[1])
                self.finished_thread('generate_pemi',
                                     name_all=str(pathlib.Path('logs', 'generate_pemi', file_name[1])),
                                     name_now=str(pathlib.Path('logs', 'generate_pemi', file_name[0])))
                return
            # Если всё прошло запускаем поток
            self.logging_dict[file_name[0]].info('Запуск на выполнение')
            generate['logging'], generate['queue'] = self.logging_dict[file_name[0]], self.queue
            generate['no_freq_lim'], generate['no_excel_file'] = no_freq_lim, no_excel_file
            generate['3db_difference'], generate['default_path'] = db_difference, self.default_path
            generate['move'] = len(self.thread_dict['generate_pemi'])
            self.thread = GenerationFile(generate)
            self.thread.status_finish.connect(self.finished_thread)
            self.thread.status.connect(self.statusBar().showMessage)
            self.thread.start()
            self.thread_dict['generate_pemi'][str(self.thread)] = {'filename_all': file_name[1],
                                                                   'filename_now': file_name[0]}
        except BaseException as exception:
            self.logging_dict[file_name[0]].error('Ошибка при старте generate_pemi')
            self.logging_dict[file_name[0]].error(exception)
            self.logging_dict[file_name[0]].error(traceback.format_exc())
            self.on_message_changed('УПС!', 'Неизвестная ошибка, обратитесь к разработчику')
            self.finished_thread('generate_pemi',
                                 name_all=str(pathlib.Path('logs', 'generate_pemi', file_name[1])),
                                 name_now=str(pathlib.Path('logs', 'generate_pemi', file_name[0])))
            return

    def generate_hfe(self):
        file_name = self.logging_file('generate_hfe')
        self.logging_dict[file_name[0]].info('----------------Запускаем generate_hfe----------------')
        self.logging_dict[file_name[0]].info('Проверка данных')
        try:
            generate = checked_hfe_generation(self.lineEdit_path_folder_HFE, self.lineEdit_complect_quant_HFE,
                                              self.groupBox_required_values_HFE, self.lineEdit_frequency,
                                              self.lineEdit_level)
            if isinstance(generate, list):
                self.logging_dict[file_name[0]].warning(generate[1])
                self.logging_dict[file_name[0]].warning('Ошибки в заполнении формы, программа не запущена в работу')
                self.on_message_changed(generate[0], generate[1])
                self.finished_thread('generate_hfe',
                                     name_all=str(pathlib.Path('logs', 'generate_hfe', file_name[1])),
                                     name_now=str(pathlib.Path('logs', 'generate_hfe', file_name[0])))
                return
            # Если всё прошло запускаем поток
            self.logging_dict[file_name[0]].info('Запуск на выполнение')
            generate['logging'], generate['queue'] = self.logging_dict[file_name[0]], self.queue
            generate['default_path'], generate['move'] = self.default_path, len(self.thread_dict['generate_hfe'])
            self.thread = HFEGeneration(generate)
            self.thread.status_finish.connect(self.finished_thread)
            self.thread.status.connect(self.statusBar().showMessage)
            self.thread.start()
            self.thread_dict['generate_hfe'][str(self.thread)] = {'filename_all': file_name[1],
                                                                  'filename_now': file_name[0]}
        except BaseException as exception:
            self.logging_dict[file_name[0]].error('Ошибка при старте generate_hfe')
            self.logging_dict[file_name[0]].error(exception)
            self.logging_dict[file_name[0]].error(traceback.format_exc())
            self.on_message_changed('УПС!', 'Неизвестная ошибка, обратитесь к разработчику')
            self.finished_thread('generate_hfe',
                                 name_all=str(pathlib.Path('logs', 'generate_hfe', file_name[1])),
                                 name_now=str(pathlib.Path('logs', 'generate_hfe', file_name[0])))
            return

    def generate_hfi(self):
        file_name = self.logging_file('generate_hfi')
        self.logging_dict[file_name[0]].info('----------------Запускаем generate_hfi----------------')
        self.logging_dict[file_name[0]].info('Проверка данных')
        try:
            generate = checked_hfi_generation(self.lineEdit_path_folder_HFI, self.lineEdit_imposition_freq,
                                              self.lineEdit_complect_quant_HFI,
                                              [self.checkBox_power_supply.isChecked(),
                                               self.checkBox_symetrical.isChecked(),
                                               self.checkBox_asymetriacal.isChecked()])
            if isinstance(generate, list):
                self.logging_dict[file_name[0]].warning(generate[1])
                self.logging_dict[file_name[0]].warning('Ошибки в заполнении формы, программа не запущена в работу')
                self.on_message_changed(generate[0], generate[1])
                self.finished_thread('generate_hfi',
                                     name_all=str(pathlib.Path('logs', 'generate_hfi', file_name[1])),
                                     name_now=str(pathlib.Path('logs', 'generate_hfi', file_name[0])))
                return
            # Если всё прошло запускаем поток
            self.logging_dict[file_name[0]].info('Запуск на выполнение')
            generate['logging'], generate['queue'] = self.logging_dict[file_name[0]], self.queue
            generate['default_path'], generate['move'] = self.default_path, len(self.thread_dict['generate_hfi'])
            self.thread = HFIGeneration(generate)
            self.thread.status_finish.connect(self.finished_thread)
            self.thread.status.connect(self.statusBar().showMessage)
            self.thread.start()
            self.thread_dict['generate_hfi'][str(self.thread)] = {'filename_all': file_name[1],
                                                                  'filename_now': file_name[0]}
        except BaseException as exception:
            self.logging_dict[file_name[0]].error('Ошибка при старте generate_hfi')
            self.logging_dict[file_name[0]].error(exception)
            self.logging_dict[file_name[0]].error(traceback.format_exc())
            self.on_message_changed('УПС!', 'Неизвестная ошибка, обратитесь к разработчику')
            self.finished_thread('generate_hfi',
                                 name_all=str(pathlib.Path('logs', 'generate_hfi', file_name[1])),
                                 name_now=str(pathlib.Path('logs', 'generate_hfi', file_name[0])))
            return

    def generate_lf(self):
        file_name = self.logging_file('generate_lf')
        self.logging_dict[file_name[0]].info('----------------Запускаем generate_lf----------------')
        self.logging_dict[file_name[0]].info('Проверка данных')
        try:
            generate = checked_lf_data(self.lineEdit_path_start_folder_lf, self.lineEdit_path_finish_folder_lf,
                                       self.lineEdit_path_file_excel_lf)
            if isinstance(generate, list):
                self.logging_dict[file_name[0]].warning(generate[1])
                self.logging_dict[file_name[0]].warning('Ошибки в заполнении формы, программа не запущена в работу')
                self.on_message_changed(generate[0], generate[1])
                self.finished_thread('generate_lf',
                                     name_all=str(pathlib.Path('logs', 'generate_lf', file_name[1])),
                                     name_now=str(pathlib.Path('logs', 'generate_lf', file_name[0])))
                return
            # Если всё прошло запускаем поток
            self.logging_dict[file_name[0]].info('Запуск на выполнение')
            generate['logging'], generate['queue'] = self.logging_dict[file_name[0]], self.queue
            generate['default_path'], generate['move'] = self.default_path, len(self.thread_dict['generate_lf'])
            self.thread = LFGeneration(generate)
            self.thread.status_finish.connect(self.finished_thread)
            self.thread.status.connect(self.statusBar().showMessage)
            self.thread.start()
            self.thread_dict['generate_lf'][str(self.thread)] = {'filename_all': file_name[1],
                                                                 'filename_now': file_name[0]}
        except BaseException as exception:
            self.logging_dict[file_name[0]].error('Ошибка при старте generate_lf')
            self.logging_dict[file_name[0]].error(exception)
            self.logging_dict[file_name[0]].error(traceback.format_exc())
            self.on_message_changed('УПС!', 'Неизвестная ошибка, обратитесь к разработчику')
            self.finished_thread('generate_lf',
                                 name_all=str(pathlib.Path('logs', 'generate_lf', file_name[1])),
                                 name_now=str(pathlib.Path('logs', 'generate_lf', file_name[0])))
            return

    def parcing_file(self):
        file_name = self.logging_file('parcing_file')
        self.logging_dict[file_name[0]].info('----------------Запускаем parcing_file----------------')
        self.logging_dict[file_name[0]].info('Проверка данных')
        try:
            self.plainTextEdit_succsess_order.clear()
            self.groupBox_succsess_order.setStyleSheet("")
            self.plainTextEdit_error_order.clear()
            self.groupBox_error_order.setStyleSheet("")
            self.plainTextEdit_errors.clear()
            self.groupBox_errors.setStyleSheet("")
            group_file = self.checkBox_group_parcing.isChecked()
            no_freq_lim = self.checkBox_no_freq_limit.isChecked()
            twelve_sectors = self.checkBox_12_sectors.isChecked()
            folder = checked_file_parcing(self.lineEdit_path_parser, group_file)
            if isinstance(folder, list):
                self.logging_dict[file_name[0]].warning(folder[1])
                self.logging_dict[file_name[0]].warning('Ошибки в заполнении формы, программа не запущена в работу')
                self.on_message_changed(folder[0], folder[1])
                self.finished_thread('parcing_file',
                                     name_all=str(pathlib.Path('logs', 'parcing_file', file_name[1])),
                                     name_now=str(pathlib.Path('logs', 'parcing_file', file_name[0])))
                return
            # Если всё прошло запускаем поток
            self.logging_dict[file_name[0]].info('Запуск на выполнение')
            folder['group_file'], folder['no_freq_lim'], folder['12_sec'] = group_file, no_freq_lim, twelve_sectors
            folder['logging'], folder['queue'] = self.logging_dict[file_name[0]], self.queue
            folder['default_path'], folder['move'] = self.default_path, len(self.thread_dict['parcing_file'])
            self.thread = FileParcing(folder)
            self.thread.status_finish.connect(self.finished_thread)
            self.thread.status.connect(self.statusBar().showMessage)
            self.thread.errors.connect(self.errors)
            self.thread.start()
            self.thread_dict['parcing_file'][str(self.thread)] = {'filename_all': file_name[1],
                                                                  'filename_now': file_name[0]}
        except BaseException as exception:
            self.logging_dict[file_name[0]].error('Ошибка при старте parcing_file')
            self.logging_dict[file_name[0]].error(exception)
            self.logging_dict[file_name[0]].error(traceback.format_exc())
            self.on_message_changed('УПС!', 'Неизвестная ошибка, обратитесь к разработчику')
            self.finished_thread('parcing_file',
                                 name_all=str(pathlib.Path('logs', 'parcing_file', file_name[1])),
                                 name_now=str(pathlib.Path('logs', 'parcing_file', file_name[0])))
            return

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
        file_name = self.logging_file('zone_checked')
        self.logging_dict[file_name[0]].info('----------------Запускаем zone_checked----------------')
        self.logging_dict[file_name[0]].info('Проверка данных')
        try:
            department = self.groupBox_FSB.isChecked()
            win_lin = self.checkBox_win_lin.isChecked()
            one_table = self.checkBox_first_table.isChecked()
            zone = [self.lineEdit_stationary_FSB, self.lineEdit_carry_FSB, self.lineEdit_wear_FSB, self.lineEdit_r1_FSB,
                    self.lineEdit_r1s_FSB] \
                if department else [self.lineEdit_stationary_FSTEK, self.lineEdit_carry_FSTEK,
                                    self.lineEdit_wear_FSTEK, self.lineEdit_r1_FSTEK]
            zone_all = checked_zone_checked(self.lineEdit_path_check, self.lineEdit_table_number, zone)
            if isinstance(zone_all, list):
                self.logging_dict[file_name[0]].warning(zone_all[1])
                self.logging_dict[file_name[0]].warning('Ошибки в заполнении формы, программа не запущена в работу')
                self.on_message_changed(zone_all[0], zone_all[1])
                self.finished_thread('zone_checked',
                                     name_all=str(pathlib.Path('logs', 'zone_checked', file_name[1])),
                                     name_now=str(pathlib.Path('logs', 'zone_checked', file_name[0])))
                return
            # Если всё прошло запускаем поток
            self.logging_dict[file_name[0]].info('Запуск на выполнение')
            zone = {'path_check': self.lineEdit_path_check.text().strip(),
                    'table_number': self.lineEdit_table_number.text().strip(), 'department': department,
                    'win_lin': win_lin, 'zone_all': zone_all, 'one_table': one_table,
                    'logging': self.logging_dict[file_name[0]],
                    'queue': self.queue, 'default_path': self.default_path,
                    'move': len(self.thread_dict['zone_checked']),
                    'extend_report': self.checkBox_extend_report.isChecked()}
            self.thread = ZoneChecked(zone)
            self.thread.status_finish.connect(self.finished_thread)
            self.thread.status.connect(self.statusBar().showMessage)
            self.thread.start()
            self.thread_dict['zone_checked'][str(self.thread)] = {'filename_all': file_name[1],
                                                                  'filename_now': file_name[0]}
        except BaseException as exception:
            self.logging_dict[file_name[0]].error('Ошибка при старте zone_checked')
            self.logging_dict[file_name[0]].error(exception)
            self.logging_dict[file_name[0]].error(traceback.format_exc())
            self.on_message_changed('УПС!', 'Неизвестная ошибка, обратитесь к разработчику')
            self.finished_thread('zone_checked',
                                 name_all=str(pathlib.Path('logs', 'zone_checked', file_name[1])),
                                 name_now=str(pathlib.Path('logs', 'zone_checked', file_name[0])))
            return

    def delete_header_footer(self):
        file_name = self.logging_file('delete_header_footer')
        self.logging_dict[file_name[0]].info('----------------Запускаем delete_header_footer----------------')
        self.logging_dict[file_name[0]].info('Проверка данных')
        try:
            output = checked_delete_header_footer(self.lineEdit_path_start_extract, self.checkBox_director,
                                                  self.lineEdit_old_director, self.lineEdit_new_director,
                                                  self.checkBox_margin, self.doubleSpinBox_left_margin,
                                                  self.doubleSpinBox_top_margin, self.doubleSpinBox_right_margin,
                                                  self.doubleSpinBox_bottom_margin)
            if isinstance(output, list):
                self.logging_dict[file_name[0]].warning(output[1])
                self.logging_dict[file_name[0]].warning('Ошибки в заполнении формы, программа не запущена в работу')
                self.on_message_changed(output[0], output[1])
                self.finished_thread('delete_header_footer',
                                     name_all=str(pathlib.Path('logs', 'delete_header_footer', file_name[1])),
                                     name_now=str(pathlib.Path('logs', 'delete_header_footer', file_name[0])))
                return
            # Если всё прошло запускаем поток
            self.logging_dict[file_name[0]].info('Запуск на выполнение')
            output['logging'], output['queue'] = self.logging_dict[file_name[0]], self.queue
            output['project_prescription'] = self.checkBox_project_prescription.isChecked()
            output['default_path'] = self.default_path
            output['conclusion'] = self.lineEdit_conclusion.text()
            output['protocol'] = self.lineEdit_protocol.text()
            output['prescription'] = self.lineEdit_prescription.text()
            output['move'] = len(self.thread_dict['delete_header_footer'])
            self.thread = DeleteHeaderFooter(output)
            self.thread.status_finish.connect(self.finished_thread)
            self.thread.status.connect(self.statusBar().showMessage)
            self.thread.start()
            self.thread_dict['delete_header_footer'][str(self.thread)] = {'filename_all': file_name[1],
                                                                          'filename_now': file_name[0]}
        except BaseException as exception:
            self.logging_dict[file_name[0]].error('Ошибка при старте delete_header_footer')
            self.logging_dict[file_name[0]].error(exception)
            self.logging_dict[file_name[0]].error(traceback.format_exc())
            self.on_message_changed('УПС!', 'Неизвестная ошибка, обратитесь к разработчику')
            self.finished_thread('delete_header_footer',
                                 name_all=str(pathlib.Path('logs', 'delete_header_footer', file_name[1])),
                                 name_now=str(pathlib.Path('logs', 'delete_header_footer', file_name[0])))
            return

    def generate_cc(self):
        file_name = self.logging_file('continuous_spectrum')
        self.logging_dict[file_name[0]].info('----------------Запускаем continuous_spectrum----------------')
        self.logging_dict[file_name[0]].info('Проверка данных')
        try:
            generate = checked_generation_cc(self.lineEdit_path_folder_start_cc, self.lineEdit_path_folder_finish_cc,
                                             self.lineEdit_set_number_cc, self.checkBox_cc_frequency,
                                             self.lineEdit_frequency_cc, self.checkBox_cc_txt,
                                             self.checkBox_cc_dispersion, self.lineEdit_cc_dispersion)
            if isinstance(generate, list):
                self.logging_dict[file_name[0]].warning(generate[1])
                self.logging_dict[file_name[0]].warning('Ошибки в заполнении формы, программа не запущена в работу')
                self.on_message_changed(generate[0], generate[1])
                self.finished_thread('continuous_spectrum',
                                     name_all=str(pathlib.Path('logs', 'continuous_spectrum', file_name[1])),
                                     name_now=str(pathlib.Path('logs', 'continuous_spectrum', file_name[0])))
                return
            # Если всё прошло запускаем поток
            self.logging_dict[file_name[0]].info('Запуск на выполнение')
            generate['logging'], generate['queue'] = self.logging_dict[file_name[0]], self.queue
            generate['default_path'], generate['move'] = self.default_path, len(self.thread_dict['continuous_spectrum'])
            self.thread = GenerationFileCC(generate)
            self.thread.status_finish.connect(self.finished_thread)
            self.thread.status.connect(self.statusBar().showMessage)
            self.thread.start()
            self.thread_dict['continuous_spectrum'][str(self.thread)] = {'filename_all': file_name[1],
                                                                         'filename_now': file_name[0]}
        except BaseException as exception:
            self.logging_dict[file_name[0]].error('Ошибка при старте continuous_spectrum')
            self.logging_dict[file_name[0]].error(exception)
            self.logging_dict[file_name[0]].error(traceback.format_exc())
            self.on_message_changed('УПС!', 'Неизвестная ошибка, обратитесь к разработчику')
            self.finished_thread('continuous_spectrum',
                                 name_all=str(pathlib.Path('logs', 'continuous_spectrum', file_name[1])),
                                 name_now=str(pathlib.Path('logs', 'continuous_spectrum', file_name[0])))
            return

    def change_number_instance(self):
        file_name = self.logging_file('change_number_instance')
        self.logging_dict[file_name[0]].info('----------------Запускаем change_number_instance----------------')
        self.logging_dict[file_name[0]].info('Проверка данных')
        try:
            incoming = checked_number_instance(self.lineEdit_path_folder_old_number_instance,
                                               self.lineEdit_path_folder_new_number_instance,
                                               self.lineEdit_number_instance)
            if isinstance(incoming, list):
                self.logging_dict[file_name[0]].warning(incoming[1])
                self.logging_dict[file_name[0]].warning('Ошибки в заполнении формы, программа не запущена в работу')
                self.on_message_changed(incoming[0], incoming[1])
                self.finished_thread('change_number_instance',
                                     name_all=str(pathlib.Path('logs', 'change_number_instance', file_name[1])),
                                     name_now=str(pathlib.Path('logs', 'change_number_instance', file_name[0])))
                return
            # Если всё прошло запускаем поток
            self.logging_dict[file_name[0]].info('Запуск на выполнение')
            incoming['logging'], incoming['queue'] = self.logging_dict[file_name[0]], self.queue
            incoming['default_path'] = self.default_path
            incoming['move'] = len(self.thread_dict['change_number_instance'])
            self.thread = ChangeNumberInstance(incoming)
            self.thread.status_finish.connect(self.finished_thread)
            self.thread.status.connect(self.statusBar().showMessage)
            self.thread.start()
            self.thread_dict['change_number_instance'][str(self.thread)] = {'filename_all': file_name[1],
                                                                            'filename_now': file_name[0]}
        except BaseException as exception:
            self.logging_dict[file_name[0]].error('Ошибка при старте change_number_instance')
            self.logging_dict[file_name[0]].error(exception)
            self.logging_dict[file_name[0]].error(traceback.format_exc())
            self.on_message_changed('УПС!', 'Неизвестная ошибка, обратитесь к разработчику')
            self.finished_thread('change_number_instance',
                                 name_all=str(pathlib.Path('logs', 'change_number_instance', file_name[1])),
                                 name_now=str(pathlib.Path('logs', 'change_number_instance', file_name[0])))
            return

    def finding_files(self):
        file_name = self.logging_file('finding_files')
        self.logging_dict[file_name[0]].info('----------------Запускаем finding_files----------------')
        self.logging_dict[file_name[0]].info('Проверка данных')
        try:
            find_files = checked_find_files(self.lineEdit_path_file_unloading_find,
                                            self.lineEdit_path_folder_start_find, self.lineEdit_path_folder_finish_find)
            if isinstance(find_files, list):
                self.logging_dict[file_name[0]].warning(find_files[1])
                self.logging_dict[file_name[0]].warning('Ошибки в заполнении формы, программа не запущена в работу')
                self.on_message_changed(find_files[0], find_files[1])
                self.finished_thread('finding_files',
                                     name_all=str(pathlib.Path('logs', 'finding_files', file_name[1])),
                                     name_now=str(pathlib.Path('logs', 'finding_files', file_name[0])))
                return
            # Если всё прошло запускаем поток
            self.logging_dict[file_name[0]].info('Запуск на выполнение')
            find_files['logging'], find_files['queue'] = self.logging_dict[file_name[0]], self.queue
            find_files['default_path'] = self.default_path
            find_files['move'] = len(self.thread_dict['finding_files'])
            self.thread = FindingFiles(find_files)
            self.thread.status_finish.connect(self.finished_thread)
            self.thread.status.connect(self.statusBar().showMessage)
            self.thread.start()
            self.thread_dict['finding_files'][str(self.thread)] = {'filename_all': file_name[1],
                                                                   'filename_now': file_name[0]}
        except BaseException as exception:
            self.logging_dict[file_name[0]].error('Ошибка при старте finding_files')
            self.logging_dict[file_name[0]].error(exception)
            self.logging_dict[file_name[0]].error(traceback.format_exc())
            self.on_message_changed('УПС!', 'Неизвестная ошибка, обратитесь к разработчику')
            self.finished_thread('finding_files',
                                 name_all=str(pathlib.Path('logs', 'finding_files', file_name[1])),
                                 name_now=str(pathlib.Path('logs', 'finding_files', file_name[0])))
            return

    def gen_lf_pemi(self):
        file_name = self.logging_file('gen_lf_pemi')
        self.logging_dict[file_name[0]].info('----------------Запускаем gen_lf_pemi----------------')
        self.logging_dict[file_name[0]].info('Проверка данных')
        try:
            incoming_data = checked_lf_pemi(self.lineEdit_path_folder_start_lf_pemi,
                                            self.lineEdit_path_folder_finish_lf_pemi,
                                            self.lineEdit_set_number_lf_pemi,
                                            self.checkBox_lf_pemi_values_spread,
                                            self.lineEdit_set_number_lf_pemi)
            if isinstance(incoming_data, list):
                self.logging_dict[file_name[0]].warning(incoming_data[1])
                self.logging_dict[file_name[0]].warning('Ошибки в заполнении формы, программа не запущена в работу')
                self.on_message_changed(incoming_data[0], incoming_data[1])
                self.finished_thread('gen_lf_pemi',
                                     name_all=str(pathlib.Path('logs', 'gen_lf_pemi', file_name[1])),
                                     name_now=str(pathlib.Path('logs', 'gen_lf_pemi', file_name[0])))
                return
            # Если всё прошло запускаем поток
            self.logging_dict[file_name[0]].info('Запуск на выполнение')
            incoming_data['logging'], incoming_data['queue'] = self.logging_dict[file_name[0]], self.queue
            incoming_data['default_path'] = self.default_path
            incoming_data['move'] = len(self.thread_dict['gen_lf_pemi'])
            self.thread = FindingFiles(incoming_data)
            self.thread.status_finish.connect(self.finished_thread)
            self.thread.status.connect(self.statusBar().showMessage)
            self.thread.start()
            self.thread_dict['gen_lf_pemi'][str(self.thread)] = {'filename_all': file_name[1],
                                                                 'filename_now': file_name[0]}
        except BaseException as exception:
            self.logging_dict[file_name[0]].error('Ошибка при старте gen_lf_pemi')
            self.logging_dict[file_name[0]].error(exception)
            self.logging_dict[file_name[0]].error(traceback.format_exc())
            self.on_message_changed('УПС!', 'Неизвестная ошибка, обратитесь к разработчику')
            self.finished_thread('gen_lf_pemi',
                                 name_all=str(pathlib.Path('logs', 'gen_lf_pemi', file_name[1])),
                                 name_now=str(pathlib.Path('logs', 'gen_lf_pemi', file_name[0])))
            return

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
