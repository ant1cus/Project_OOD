import json
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
                     checked_application_data, checked_lf_data)
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
        self.q = queue.Queue(maxsize=1)
        logging.basicConfig(filename="my_log.log",
                            level=logging.DEBUG,
                            filemode="w",
                            format="%(asctime)s - %(levelname)s - %(funcName)s: %(lineno)d - %(message)s")
        self.pushButton_open_zone_check.clicked.connect((lambda: self.browse(1)))
        self.pushButton_open_parser.clicked.connect((lambda: self.browse(2)))
        self.pushButton_open_original_file.clicked.connect((lambda: self.browse(3)))
        self.pushButton_open_finish_folder.clicked.connect((lambda: self.browse(4)))
        self.pushButton_open_freq_restrict.clicked.connect((lambda: self.browse(5)))
        self.pushButton_open_file_HFE.clicked.connect((lambda: self.browse(6)))
        self.pushButton_open_file_HFI.clicked.connect((lambda: self.browse(7)))
        self.pushButton_open_original_exctract.clicked.connect((lambda: self.browse(8)))
        self.pushButton_open_example.clicked.connect((lambda: self.browse(9)))
        self.pushButton_open_finish_folder_example.clicked.connect((lambda: self.browse(10)))
        self.pushButton_open_start_folder_lf.clicked.connect((lambda: self.browse(11)))
        self.pushButton_open_finish_folder_lf.clicked.connect((lambda: self.browse(12)))
        self.pushButton_open_file_excel_lf.clicked.connect((lambda: self.browse(13)))
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
        self.tabWidget.tabBar().tabMoved.connect(self.tab_)
        self.tabWidget.tabBarClicked.connect(self.tab_click)
        self.tabWidget.tabCloseRequested.connect(lambda index: self.tabWidget.removeTab(index))
        self.start_name = False
        self.path_for_default = pathlib.Path.cwd()  # Путь для файла настроек
        # Имена в файле
        self.name_list = {'checked-path_check': 'Путь к файлам', 'checked-table_number': 'Номер таблицы',
                          'checked-stationary_FSB': 'Стац. ФСБ', 'checked-carry_FSB': 'Воз. ФСБ',
                          'checked-wear_FSB': 'Нос. ФСБ', 'checked-r1_FSB': 'r1 ФСБ', 'checked-r1s_FSB': 'r1` ФСБ',
                          'checked-stationary_FSTEK': 'Стац. ФСТЭК', 'checked-carry_FSTEK': 'Воз. ФСТЭК',
                          'checked-wear_FSTEK': 'Нос. ФСТЭК', 'checked-r1_FSTEK': 'r1 ФСТЭК',
                          'parser-path_parser': 'Путь к файлам',
                          'extract-path_original_extract': 'Путь к файлам',
                          'extract-conclusion_post': 'Должность заключение',
                          'extract-conclusion_name': 'ФИО заключение',
                          'extract-protocol_post': 'Должность протокол', 'extract-protocol_name': 'ФИО протокол',
                          'extract-prescription_post': 'Должность предписание',
                          'extract-prescription_name': 'ФИО предписание',
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
                          'application-quantity_document': 'Количество комплектов',
                          'LF-path_start_folder': 'Путь к начальной папке',
                          'LF-path_finish_folder': 'Путь к конечной папке',
                          'LF-path_file_excel': 'Путь к файлу генератору'}
        # Грузим значения по умолчанию
        self.name_tab = {"tab_zone_checked": "Проверка зон", "tab_parser": "Парсер txt",
                         "tab_exctract": "Обезличивание", "tab_gen_application": "Генератор приложений",
                         "tab_gen_pemi": "Генератор ПЭМИ", "tab_gen_HFE": "Генератор ВЧО",
                         "tab_gen_HFI": "Генератор ВЧН", 'tab_gen_LF': 'Генератор НЧ'}
        self.name_action = {"tab_zone_checked": self.action_zone_checked, "tab_parser": self.action_parser,
                            "tab_exctract": self.action_extract, "tab_gen_application": self.action_gen_application,
                            "tab_gen_pemi": self.action_gen_pemi, "tab_gen_HFE": self.action_gen_HFE,
                            "tab_gen_HFI": self.action_gen_HFI, 'tab_gen_LF': self.action_gen_LF}
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
                                                  '6': "tab_gen_HFI", '7': "tab_gen_LF"},
                                    "tab_visible": {"tab_zone_checked": True, "tab_parser": True, "tab_exctract": True,
                                                    "tab_gen_application": True, "tab_gen_pemi": True,
                                                    "tab_gen_HFE": True, "tab_gen_HFI": True, "tab_gen_LF": True}
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
                rewrite(self.path_for_default, self.tab_order, visible='tab_order')
                self.tab_visible[str(self.tabWidget.widget(tab).objectName())] = True
                rewrite(self.path_for_default, self.tab_visible, visible='tab_visible')
            self.tab_for_paint[self.tabWidget.widget(tab).objectName()] = self.tabWidget.widget(tab)
        self.tabWidget.clear()
        for tab in self.tab_order:
            if self.tab_visible[self.tab_order[tab]]:
                self.name_action[self.tab_order[tab]].setChecked(True)
                self.tabWidget.addTab(self.tab_for_paint[self.tab_order[tab]], self.name_tab[self.tab_order[tab]])
        self.tabWidget.tabBar().setCurrentIndex(0)
        # Линии для заполнения
        self.line = [self.lineEdit_path_check, self.lineEdit_table_number, self.lineEdit_stationary_FSB,
                     self.lineEdit_carry_FSB, self.lineEdit_wear_FSB, self.lineEdit_r1_FSB, self.lineEdit_r1s_FSB,
                     self.lineEdit_stationary_FSTEK, self.lineEdit_carry_FSTEK, self.lineEdit_wear_FSTEK,
                     self.lineEdit_r1_FSTEK, self.lineEdit_path_parser, self.lineEdit_path_original_extract,
                     self.lineEdit_conclusion_post, self.lineEdit_conclusion_name, self.lineEdit_protocol_post,
                     self.lineEdit_protocol_name, self.lineEdit_prescription_post, self.lineEdit_prescription_name,
                     self.lineEdit_path_example, self.lineEdit_path_finish_folder_example,
                     self.lineEdit_number_position, self.lineEdit_quantity_document,
                     self.lineEdit_path_original_file, self.lineEdit_path_finish_folder,
                     self.lineEdit_path_freq_restrict, self.lineEdit_complect_quant_pemi,
                     self.lineEdit_complect_number_pemi, self.lineEdit_path_file_HFE,
                     self.lineEdit_complect_quant_HFE, self.lineEdit_frequency, self.lineEdit_level,
                     self.lineEdit_path_file_HFI, self.lineEdit_complect_quant_HFI, self.lineEdit_imposition_freq,
                     self.lineEdit_path_start_folder_lf, self.lineEdit_path_finish_folder_lf,
                     self.lineEdit_path_file_excel_lf]
        self.default_date(self.data)

    def tab_(self, index):
        for tab in self.tab_order.items():
            if tab[1] == self.start_name:
                self.tab_order[str(index)], self.tab_order[tab[0]] = self.tab_order[tab[0]], self.tab_order[str(index)]
                break
        rewrite(self.path_for_default, self.tab_order, order='tab_order')

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
                        rewrite(self.path_for_default, self.tab_visible, visible='tab_visible')
                else:
                    if self.tab_visible[el]:
                        self.tab_visible[el] = False
                        rewrite(self.path_for_default, self.tab_visible, visible='tab_visible')

    def default_date(self, d):
        for i, el in enumerate(self.name_list):
            if el in d:
                self.line[i].setText(d[el])  # Помещаем значение

    def default_settings(self):  # Запускаем окно с настройками по умолчанию.
        self.close()
        window_add = DefaultWindow(self, self.path_for_default)
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

    def browse(self, num):  # Для кнопки открыть
        directory = None
        if num in [5, 9, 13]:  # Если необходимо открыть файл
            directory = QFileDialog.getOpenFileName(self, "Find Files", QDir.currentPath())
        elif num in [1, 2, 3, 4, 6, 7, 8, 10, 11, 12]:  # Если необходимо открыть директорию
            directory = QFileDialog.getExistingDirectory(self, "Find Files", QDir.currentPath())
        # Список линий
        line = [self.lineEdit_path_check, self.lineEdit_path_parser, self.lineEdit_path_original_file,
                self.lineEdit_path_finish_folder, self.lineEdit_path_freq_restrict, self.lineEdit_path_file_HFE,
                self.lineEdit_path_file_HFI, self.lineEdit_path_original_extract,
                self.lineEdit_path_example, self.lineEdit_path_finish_folder_example,
                self.lineEdit_path_start_folder_lf, self.lineEdit_path_finish_folder_lf,
                self.lineEdit_path_file_excel_lf]
        if directory:  # Если нажать кнопку отркыть в диалоге выбора
            if num in [5, 9]:  # Если файлы
                if directory[0]:  # Если есть файл, чтобы не очищалось поле
                    line[num - 1].setText(directory[0])
            else:  # Если директории
                line[num - 1].setText(directory)

    def copy_application(self):
        application = checked_application_data(self.lineEdit_path_example, self.lineEdit_path_finish_folder_example,
                                               self.lineEdit_number_position, self.lineEdit_quantity_document)
        if isinstance(application, list):
            self.on_message_changed(application[0], application[1])
            return
        # Если всё прошло запускаем поток
        application['logging'], application['q'] = logging, self.q
        self.thread = GenerateCopyApplication(application)
        self.thread.progress.connect(self.progressBar.setValue)
        self.thread.status.connect(self.show_mess)
        self.thread.messageChanged.connect(self.on_message_changed)
        self.thread.errors.connect(self.errors)
        self.thread.start()

    def generate_pemi(self):
        generate = checked_generation_pemi(self.lineEdit_path_original_file, self.lineEdit_path_finish_folder,
                                           self.lineEdit_complect_number_pemi, self.lineEdit_complect_quant_pemi,
                                           self.checkBox_freq_restrict.isChecked(), self.lineEdit_path_freq_restrict)
        if isinstance(generate, list):
            self.on_message_changed(generate[0], generate[1])
            return
        # Если всё прошло запускаем поток
        generate['logging'], generate['q'] = logging, self.q
        self.thread = GenerationFile(generate)
        self.thread.progress.connect(self.progressBar.setValue)
        self.thread.status.connect(self.show_mess)
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
        generate['logging'], generate['q'] = logging, self.q
        self.thread = HFEGeneration(generate)
        self.thread.progress.connect(self.progressBar.setValue)
        self.thread.status.connect(self.show_mess)
        self.thread.messageChanged.connect(self.on_message_changed)
        self.thread.errors.connect(self.errors)
        self.thread.start()

    def generate_hfi(self):
        generate = checked_hfi_generation(self.lineEdit_path_file_HFI, self.lineEdit_imposition_freq,
                                          self.lineEdit_complect_quant_HFI,
                                          [self.checkBox_power_supply.isChecked(),
                                           self.checkBox_symmetrical.isChecked(),
                                           self.checkBox_asymetriacal.isChecked()])
        if isinstance(generate, list):
            self.on_message_changed(generate[0], generate[1])
            return
        # Если всё прошло запускаем поток
        generate['logging'], generate['q'] = logging, self.q
        self.thread = HFIGeneration(generate)
        self.thread.progress.connect(self.progressBar.setValue)
        self.thread.status.connect(self.show_mess)
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
        generate['logging'], generate['q'] = logging, self.q
        self.thread = LFGeneration(generate)
        self.thread.progress.connect(self.progressBar.setValue)
        self.thread.status.connect(self.show_mess)
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
        folder['logging'], folder['q'] = logging, self.q
        self.thread = FileParcing(folder)
        self.thread.progress.connect(self.progressBar.setValue)
        self.thread.status.connect(self.show_mess)
        self.thread.errors.connect(self.errors)
        self.thread.messageChanged.connect(self.on_message_changed)
        self.thread.start()

    def errors(self):
        text = self.q.get_nowait()
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
                'table_number': self.lineEdit_table_number.text().strip(), 'department': department,
                'win_lin': win_lin, 'zone_all': zone_all, 'one_table': one_table, 'logging': logging, 'q': self.q}
        self.thread = ZoneChecked(zone)
        self.thread.progress.connect(self.progressBar.setValue)
        self.thread.status.connect(self.show_mess)
        self.thread.messageChanged.connect(self.on_message_changed)
        self.thread.start()

    def delete_header_footer(self):
        output = checked_delete_header_footer(self.lineEdit_path_original_extract, self.lineEdit_conclusion_post,
                                              self.lineEdit_conclusion_name, self.lineEdit_protocol_post,
                                              self.lineEdit_protocol_name, self.lineEdit_prescription_post,
                                              self.lineEdit_prescription_name)
        if isinstance(output, list):
            self.on_message_changed(output[0], output[1])
            return
        # Если всё прошло запускаем поток
        output['logging'], output['q'], output['default_path'] = logging, self.q, self.path_for_default
        self.thread = DeleteHeaderFooter(output)
        self.thread.progress.connect(self.progressBar.setValue)
        self.thread.status.connect(self.show_mess)
        self.thread.messageChanged.connect(self.on_message_changed)
        self.thread.start()

    def pause_thread(self):
        if self.q.empty():
            self.statusBar().showMessage(self.statusBar().currentMessage() + ' (прерывание процесса, подождите...)')
            self.q.put(True)

    def show_mess(self, value):  # Вывод значения в статус бар
        self.statusBar().showMessage(value)

    def on_message_changed(self, title, description):
        if title == 'УПС!':
            QMessageBox.critical(self, title, description)
        elif title == 'Внимание!':
            QMessageBox.warning(self, title, description)
        elif title == 'Вопрос?':
            self.statusBar().clearMessage()
            ans = QMessageBox.question(self, title, description, QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if ans == QMessageBox.No:
                self.thread.q.put(True)
            else:
                self.thread.q.put(False)
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
