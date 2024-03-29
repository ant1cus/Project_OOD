import datetime
import json
import os
import pathlib
import queue
import sys
import traceback

from PyQt5.QtGui import QSyntaxHighlighter, QTextCharFormat, QColor

import Main
import logging
import about
from PyQt5.QtCore import QTranslator, QLocale, QLibraryInfo, QDir
from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog, QMessageBox, QDialog
from checked import (checked_zone_checked, checked_file_parcing, checked_generation_pemi,
                     checked_delete_header_footer, checked_hfe_generation, checked_hfi_generation,
                     checked_application_data, checked_lf_data, checked_generation_cc, checked_number_instance)
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
        self.pushButton_open_folder_old_number_instance.clicked.connect(
            (lambda: self.browse(self.lineEdit_path_folder_old_number_instance)))
        self.pushButton_open_folder_new_number_instance.clicked.connect(
            (lambda: self.browse(self.lineEdit_path_folder_new_number_instance)))
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
        self.pushButton_number_instance.clicked.connect(self.change_number_instance)
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
                          'parser-checkBox_12_sectors': ['12 секторов', self.checkBox_12_sectors],
                          'extract-path_folder_start_extract': ['Путь к дир. с файлами',
                                                                self.lineEdit_path_start_extract],
                          'extract-groupBox_value_for_extract': ['Значения для выписки',
                                                                 self.groupBox_value_for_extract],
                          'extract-conclusion': ['Заключение', self.lineEdit_conclusion],
                          'extract-protocol': ['Протокол', self.lineEdit_protocol],
                          'extract-prescription': ['Предписание', self.lineEdit_prescription],
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
                          'gen_pemi-set_quant_pemi': ['Количество комплектов', self.lineEdit_complect_quant_pemi],
                          'gen_pemi-set_number_pemi': ['Номера комплектов', self.lineEdit_complect_number_pemi],
                          'HFE-path_file_HFE': ['Путь к дир. с файлами', self.lineEdit_path_file_HFE],
                          'HFE-set_quant_HFE': ['Количество комплектов', self.lineEdit_complect_quant_HFE],
                          'HFE-checkBox_required_values_HFE': ['Значения вручную', self.checkBox_required_values_HFE],
                          'HFE-frequency': ['Частота', self.lineEdit_frequency],
                          'HFE-level': ['Уровень', self.lineEdit_level],
                          'HFI-path_file_HFI': ['Путь к дир. с файлами', self.lineEdit_path_file_HFI],
                          'HFI-set_quant_HFI': ['Количество комплектов', self.lineEdit_complect_quant_HFE],
                          'HFI-checkBox_imposition_freq': ['Ручной ввод частоты', self.checkBox_imposition_freq],
                          'HFI-imposition_freq': ['Частота навязывания', self.lineEdit_imposition_freq],
                          'HFI-checkBox_power_supply': ['Питание', self.checkBox_power_supply],
                          'HFI-checkBox_symmetrical': ['Симметричка', self.checkBox_symetrical],
                          'HFI-checkBox_asymmetriacal': ['Не симметричка', self.checkBox_asymetriacal],
                          'application-path_file_example': ['Путь к файлу', self.lineEdit_path_start_example],
                          'application-path_folder_finish_example': ['Путь к конечной дир.',
                                                                     self.lineEdit_path_finish_example],
                          'application-number_position': ['Номер позиции', self.lineEdit_number_position],
                          'application-quantity_document': ['Количество комплектов', self.lineEdit_quantity_document],
                          'LF-path_folder_start': ['Путь к начальной дир.', self.lineEdit_path_start_folder_lf],
                          'LF-path_folder_finish': ['Путь к конечной дир.', self.lineEdit_path_finish_folder_lf],
                          'LF-path_file_excel': ['Путь к файлу генератору', self.lineEdit_path_file_excel_lf],
                          'CC-path_folder_start': ['Путь к файлам спектра', self.lineEdit_path_folder_start_cc],
                          'CC-path_folder_finish': ['Путь к конечной папке', self.lineEdit_path_folder_finish_cc],
                          'CC-checkBox_cc_frequency': ['Конечная частота', self.checkBox_cc_frequency],
                          'CC-set_frequency': ['Конечная частота (МГц)', self.lineEdit_frequency_cc],
                          'CC-set_numbers': ['Номера комплектов', self.lineEdit_set_number_cc],
                          'CC-checkBox_cc_txt': ['Генерировать только txt', self.checkBox_cc_txt],
                          'CC-checkBox_cc_dispersion': ['Включить разброс значений', self.checkBox_cc_dispersion],
                          'CC-dispersion': ['Разброс значений (%)', self.lineEdit_cc_dispersion],
                          'NI-path_folder_old_number_instance': ['Путь к начальной дир.',
                                                                 self.lineEdit_path_folder_old_number_instance],
                          'NI-path_folder_new_number_instance': ['Путь к конечной дир.',
                                                                 self.lineEdit_path_folder_new_number_instance],
                          'NI-number_instance': ['Номера экземпляров', self.lineEdit_number_instance]}
        # Грузим значения по умолчанию
        self.name_tab = {"tab_zone_checked": "Проверка зон", "tab_parser": "Парсер txt",
                         "tab_exctract": "Обезличивание", "tab_gen_application": "Генератор приложений",
                         "tab_gen_pemi": "Генератор ПЭМИ", "tab_gen_HFE": "Генератор ВЧО",
                         "tab_gen_HFI": "Генератор ВЧН", "tab_gen_LF": "Генератор НЧ",
                         "tab_continuous_spectrum": "Сплошной спектр", "tab_number_instance": "Номера экземпляра"}
        self.name_action = {"tab_zone_checked": self.action_zone_checked, "tab_parser": self.action_parser,
                            "tab_exctract": self.action_extract, "tab_gen_application": self.action_gen_application,
                            "tab_gen_pemi": self.action_gen_pemi, "tab_gen_HFE": self.action_gen_HFE,
                            "tab_gen_HFI": self.action_gen_HFI, 'tab_gen_LF': self.action_gen_LF,
                            "tab_continuous_spectrum": self.action_gen_cc,
                            "tab_number_instance": self.action_number_instance}
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
                                                  '8': "tab_continuous_spectrum", '9': "tab_number_instance"},
                                    "tab_visible": {"tab_zone_checked": True, "tab_parser": True, "tab_exctract": True,
                                                    "tab_gen_application": True, "tab_gen_pemi": True,
                                                    "tab_gen_HFE": True, "tab_gen_HFI": True, "tab_gen_LF": True,
                                                    "tab_continuous_spectrum": True, "tab_number_instance": True}
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
        self.default_date(self.data)

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
        try:
            logging.info('----------------Запускаем copy_application----------------')
            logging.info('Проверка данных')
            application = checked_application_data(self.lineEdit_path_start_example, self.lineEdit_path_finish_example,
                                                   self.lineEdit_number_position, self.lineEdit_quantity_document)
            if isinstance(application, list):
                logging.info('Обнаружены ошибки данных')
                self.on_message_changed(application[0], application[1])
                return
            # Если всё прошло запускаем поток
            logging.info('Запуск на выполнение')
            application['logging'], application['queue'] = logging, self.queue
            application['default_path'] = self.default_path
            self.thread = GenerateCopyApplication(application)
            self.thread.progress.connect(self.progressBar.setValue)
            self.thread.status.connect(self.statusBar().showMessage)
            self.thread.messageChanged.connect(self.on_message_changed)
            self.thread.errors.connect(self.errors)
            self.thread.start()
        except BaseException as exception:
            logging.error('Ошибка copy_application')
            logging.error(exception)
            logging.error(traceback.format_exc())

    def generate_pemi(self):
        try:
            logging.info('----------------Запускаем generate_pemi----------------')
            logging.info('Проверка данных')
            generate = checked_generation_pemi(self.lineEdit_path_start_pemi, self.lineEdit_path_finish_pemi,
                                               self.lineEdit_complect_number_pemi, self.lineEdit_complect_quant_pemi,
                                               self.checkBox_freq_restrict.isChecked(),
                                               self.lineEdit_path_freq_restrict)
            no_freq_lim = self.checkBox_no_limit_freq_gen.isChecked()
            no_excel_file = self.checkBox_no_excel_generation.isChecked()
            db_differeence = self.checkBox_3db_difference.isChecked()
            if isinstance(generate, list):
                logging.info('Обнаружены ошибки данных')
                self.on_message_changed(generate[0], generate[1])
                return
            # Если всё прошло запускаем поток
            logging.info('Запуск на выполнение')
            generate['logging'], generate['queue'] = logging, self.queue
            generate['no_freq_lim'], generate['no_excel_file'] = no_freq_lim, no_excel_file
            generate['3db_difference'], generate['default_path'] = db_differeence, self.default_path
            self.thread = GenerationFile(generate)
            self.thread.progress.connect(self.progressBar.setValue)
            self.thread.status.connect(self.statusBar().showMessage)
            self.thread.messageChanged.connect(self.on_message_changed)
            self.thread.errors.connect(self.errors)
            self.thread.start()
        except BaseException as exception:
            logging.error('Ошибка generate_pemi')
            logging.error(exception)
            logging.error(traceback.format_exc())

    def generate_hfe(self):
        try:
            logging.info('----------------Запускаем generate_hfe----------------')
            logging.info('Проверка данных')
            generate = checked_hfe_generation(self.lineEdit_path_file_HFE, self.lineEdit_complect_quant_HFE,
                                              self.checkBox_required_values_HFE, self.lineEdit_frequency,
                                              self.lineEdit_level)
            if isinstance(generate, list):
                logging.info('Обнаружены ошибки данных')
                self.on_message_changed(generate[0], generate[1])
                return
            # Если всё прошло запускаем поток
            logging.info('Запуск на выполнение')
            generate['logging'], generate['queue'] = logging, self.queue
            generate['default_path'] = self.default_path
            self.thread = HFEGeneration(generate)
            self.thread.progress.connect(self.progressBar.setValue)
            self.thread.status.connect(self.statusBar().showMessage)
            self.thread.messageChanged.connect(self.on_message_changed)
            self.thread.errors.connect(self.errors)
            self.thread.start()
        except BaseException as exception:
            logging.error('Ошибка generate_hfe')
            logging.error(exception)
            logging.error(traceback.format_exc())

    def generate_hfi(self):
        try:
            logging.info('----------------Запускаем generate_hfi----------------')
            logging.info('Проверка данных')
            generate = checked_hfi_generation(self.lineEdit_path_file_HFI, self.lineEdit_imposition_freq,
                                              self.lineEdit_complect_quant_HFI,
                                              [self.checkBox_power_supply.isChecked(),
                                               self.checkBox_symetrical.isChecked(),
                                               self.checkBox_asymetriacal.isChecked()])
            if isinstance(generate, list):
                logging.info('Обнаружены ошибки данных')
                self.on_message_changed(generate[0], generate[1])
                return
            # Если всё прошло запускаем поток
            logging.info('Запуск на выполнение')
            generate['logging'], generate['queue'] = logging, self.queue
            generate['default_path'] = self.default_path
            self.thread = HFIGeneration(generate)
            self.thread.progress.connect(self.progressBar.setValue)
            self.thread.status.connect(self.statusBar().showMessage)
            self.thread.messageChanged.connect(self.on_message_changed)
            self.thread.errors.connect(self.errors)
            self.thread.start()
        except BaseException as exception:
            logging.error('Ошибка generate_hfi')
            logging.error(exception)
            logging.error(traceback.format_exc())

    def generate_lf(self):
        try:
            logging.info('----------------Запускаем generate_lf----------------')
            logging.info('Проверка данных')
            generate = checked_lf_data(self.lineEdit_path_start_folder_lf, self.lineEdit_path_finish_folder_lf,
                                       self.lineEdit_path_file_excel_lf)
            if isinstance(generate, list):
                logging.info('Обнаружены ошибки данных')
                self.on_message_changed(generate[0], generate[1])
                return
            # Если всё прошло запускаем поток
            logging.info('Запуск на выполнение')
            generate['logging'], generate['queue'] = logging, self.queue
            generate['default_path'] = self.default_path
            self.thread = LFGeneration(generate)
            self.thread.progress.connect(self.progressBar.setValue)
            self.thread.status.connect(self.statusBar().showMessage)
            self.thread.messageChanged.connect(self.on_message_changed)
            self.thread.errors.connect(self.errors)
            self.thread.start()
        except BaseException as exception:
            logging.error('Ошибка generate_lf')
            logging.error(exception)
            logging.error(traceback.format_exc())

    def parcing_file(self):
        try:
            logging.info('----------------Запускаем parcing_file----------------')
            logging.info('Проверка данных')
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
                logging.info('Обнаружены ошибки данных')
                self.on_message_changed(folder[0], folder[1])
                return
            # Если всё прошло запускаем поток
            logging.info('Запуск на выполнение')
            folder['group_file'], folder['no_freq_lim'], folder['12_sec'] = group_file, no_freq_lim, twelve_sectors
            folder['logging'], folder['queue'] = logging, self.queue
            folder['default_path'] = self.default_path
            self.thread = FileParcing(folder)
            self.thread.progress.connect(self.progressBar.setValue)
            self.thread.status.connect(self.statusBar().showMessage)
            self.thread.errors.connect(self.errors)
            self.thread.messageChanged.connect(self.on_message_changed)
            self.thread.start()
        except BaseException as exception:
            logging.error('Ошибка parcing_file')
            logging.error(exception)
            logging.error(traceback.format_exc())

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
        try:
            logging.info('----------------Запускаем checked_zone----------------')
            logging.info('Проверка данных')
            department = self.groupBox_FSB.isChecked()
            win_lin = self.checkBox_win_lin.isChecked()
            one_table = self.checkBox_first_table.isChecked()
            zone = [self.lineEdit_stationary_FSB, self.lineEdit_carry_FSB, self.lineEdit_wear_FSB, self.lineEdit_r1_FSB,
                    self.lineEdit_r1s_FSB] \
                if department else [self.lineEdit_stationary_FSTEK, self.lineEdit_carry_FSTEK,
                                    self.lineEdit_wear_FSTEK, self.lineEdit_r1_FSTEK]
            zone_all = checked_zone_checked(self.lineEdit_path_check, self.lineEdit_table_number, zone)
            if isinstance(zone_all, list):
                logging.info('Обнаружены ошибки данных')
                self.on_message_changed(zone_all[0], zone_all[1])
                return
            # Если всё прошло запускаем поток
            logging.info('Запуск на выполнение')
            if self.checkBox_win_lin.isChecked():
                zone = {i + 5: zone_all[i] for i in zone_all}
                zone_all = {**zone_all, **zone}
            zone = {'path_check': self.lineEdit_path_check.text().strip(),
                    'table_number': self.lineEdit_table_number.text().strip(), 'department': department,
                    'win_lin': win_lin, 'zone_all': zone_all, 'one_table': one_table, 'logging': logging,
                    'queue': self.queue, 'default_path': self.default_path}
            self.thread = ZoneChecked(zone)
            self.thread.progress.connect(self.progressBar.setValue)
            self.thread.status.connect(self.statusBar().showMessage)
            self.thread.messageChanged.connect(self.on_message_changed)
            self.thread.start()
        except BaseException as exception:
            logging.error('Ошибка checked_zone')
            logging.error(exception)
            logging.error(traceback.format_exc())

    def delete_header_footer(self):
        try:
            logging.info('----------------Запускаем delete_header_footer----------------')
            logging.info('Проверка данных')
            output = checked_delete_header_footer(self.lineEdit_path_start_extract)
            if isinstance(output, list):
                logging.info('Обнаружены ошибки данных')
                self.on_message_changed(output[0], output[1])
                return
            # Если всё прошло запускаем поток
            logging.info('Запуск на выполнение')
            output['logging'], output['queue'] = logging, self.queue
            output['default_path'] = self.default_path
            output['conclusion'] = self.lineEdit_conclusion.text()
            output['protocol'] = self.lineEdit_protocol.text()
            output['prescription'] = self.lineEdit_prescription.text()
            self.thread = DeleteHeaderFooter(output)
            self.thread.progress.connect(self.progressBar.setValue)
            self.thread.status.connect(self.statusBar().showMessage)
            self.thread.messageChanged.connect(self.on_message_changed)
            self.thread.start()
        except BaseException as exception:
            logging.error('Ошибка delete_header_footer')
            logging.error(exception)
            logging.error(traceback.format_exc())

    def generate_cc(self):
        try:
            logging.info('----------------Запускаем generate_cc----------------')
            logging.info('Проверка данных')
            generate = checked_generation_cc(self.lineEdit_path_folder_start_cc, self.lineEdit_path_folder_finish_cc,
                                             self.lineEdit_set_number_cc, self.checkBox_cc_frequency,
                                             self.lineEdit_frequency_cc, self.checkBox_cc_txt,
                                             self.checkBox_cc_dispersion, self.lineEdit_cc_dispersion)
            if isinstance(generate, list):
                logging.info('Обнаружены ошибки данных')
                self.on_message_changed(generate[0], generate[1])
                return
            # Если всё прошло запускаем поток
            logging.info('Запуск на выполнение')
            generate['logging'], generate['queue'] = logging, self.queue
            generate['default_path'] = self.default_path
            self.thread = GenerationFileCC(generate)
            self.thread.progress.connect(self.progressBar.setValue)
            self.thread.status.connect(self.statusBar().showMessage)
            self.thread.messageChanged.connect(self.on_message_changed)
            self.thread.errors.connect(self.errors)
            self.thread.start()
        except BaseException as exception:
            logging.error('Ошибка generate_cc')
            logging.error(exception)
            logging.error(traceback.format_exc())

    def change_number_instance(self):
        try:
            logging.info('----------------Запускаем change_number_instance----------------')
            logging.info('Проверка данных')
            incoming = checked_number_instance(self.lineEdit_path_folder_old_number_instance,
                                               self.lineEdit_path_folder_new_number_instance,
                                               self.lineEdit_number_instance)
            if isinstance(incoming, list):
                logging.info('Обнаружены ошибки данных')
                self.on_message_changed(incoming[0], incoming[1])
                return
            # Если всё прошло запускаем поток
            logging.info('Запуск на выполнение')
            incoming['logging'], incoming['queue'] = logging, self.queue
            incoming['default_path'] = self.default_path
            self.thread = ChangeNumberInstance(incoming)
            self.thread.progress.connect(self.progressBar.setValue)
            self.thread.status.connect(self.statusBar().showMessage)
            self.thread.messageChanged.connect(self.on_message_changed)
            self.thread.errors.connect(self.errors)
            self.thread.start()
        except BaseException as exception:
            logging.error('Ошибка change_number_instance')
            logging.error(exception)
            logging.error(traceback.format_exc())

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
