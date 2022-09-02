import json
import pathlib
import queue
import sys

from PyQt5.QtGui import QSyntaxHighlighter, QTextCharFormat, QColor

import Main
import logging
from PyQt5.QtCore import (QTranslator, QLocale, QLibraryInfo, QDir)
from PyQt5.QtWidgets import (QMainWindow, QApplication, QFileDialog, QMessageBox, QDialog)
from checked import (checked_zone_checked, checked_file_parcing, check_generation_data, checked_delete_header_footer,
                     checked_hfe_generation, checked_hfi_generation, check_application_data)
import about
from Default import DefaultWindow
from Zone_Check import ZoneChecked
from File_Parcing import FileParcing
from Generation_Files import GenerationFile
from Delete_Header_Footer import DeleteHeaderFooter
from HFE_Generation import HFEGeneration
from HFI_Generation import HFIGeneration
from CopyApplication import GenerateCopyApplication


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
        self._highlight_lines = dict()

    def highlight_line(self, line, fmt):
        if isinstance(line, int) and line >= 0 and isinstance(fmt, QTextCharFormat):
            self._highlight_lines[line] = fmt
            tb = self.document().findBlockByNumber(line)
            self.rehighlightBlock(tb)

    def clear_highlight(self):
        self._highlight_lines = dict()
        self.rehighlight()

    def highlightBlock(self, text):
        line = self.currentBlock().blockNumber()
        fmt = self._highlight_lines.get(line)
        if fmt is not None:
            self.setFormat(0, len(text), fmt)


def highlighter(plainTextEdit):
    _highlighter = SyntaxHighlighter(plainTextEdit.document())
    fmt = QTextCharFormat()
    fmt.setBackground(QColor("#E1E1E1"))
    _highlighter.clear_highlight()
    for i in range(len(plainTextEdit.toPlainText().split('\n'))):
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
        self.action_settings_default.triggered.connect(self.default_settings)
        self.menu_about.aboutToShow.connect(about)

        self.path_for_default = pathlib.Path.cwd()  # Путь для файла настроек
        # Имена в файле
        self.name_list = {'checked-path_check': 'Путь к файлам', 'checked-table_number': 'Номер таблицы',
                          'checked-stationary_FSB': 'Стац. ФСБ', 'checked-carry_FSB': 'Воз. ФСБ',
                          'checked-wear_FSB': 'Нос. ФСБ', 'checked-r1_FSB': 'r1 ФСБ', 'checked-r1s_FSB': 'r1` ФСБ',
                          'checked-stationary_FSTEK': 'Стац. ФСТЭК', 'checked-carry_FSTEK': 'Воз. ФСТЭК',
                          'checked-wear_FSTEK': 'Нос. ФСТЭК', 'checked-r1_FSTEK': 'r1 ФСТЭК',
                          'parser-path_parser': 'Путь к файлам',
                          'extract-path_original_extract': 'Путь к файлам', 'extract-conclusion': 'ФИО заключение',
                          'extract-protocol': 'ФИО протокол', 'extract-prescription': 'ФИО предписание',
                          'gen_pemi-path_original_file': 'Путь к начальным файлам',
                          'gen_pemi-path_finish_folder': 'Путь к конечным файлам',
                          'gen_pemi-path_freq_restrict': 'Файл ограничений',
                          'gen_pemi-complect_quant_pemi': 'Количество комплектов',
                          'gen_pemi-complect_number_pemi': 'Номера комплектов',
                          'HFE-path_file_HFE': 'Путь к файлам', 'HFE-complect_quant_HFE': 'Количество комплектов',
                          'HFE-frequency': 'Частота', 'HFE-level': 'Уровень',
                          'HFI-path_file_HFI': 'Путь к файлам', 'HFI-complect_quant_HFI': 'Количество комплектов',
                          'HFI-imposition_freq': 'Частота навязывания'}
        # Грузим значения по умолчанию
        try:
            with open(pathlib.Path(pathlib.Path.cwd(), 'Настройки.txt'), "r", encoding='utf-8-sig') as f:
                data = json.load(f)
        except FileNotFoundError:
            with open(pathlib.Path(pathlib.Path.cwd(), 'Настройки.txt'), "w", encoding='utf-8-sig') as f:
                json.dump({}, f, ensure_ascii=False, sort_keys=True, indent=4)
                data = {}

        # Линии для заполнения
        self.line = [self.lineEdit_path_check, self.lineEdit_table_number, self.lineEdit_stationary_FSB,
                     self.lineEdit_carry_FSB, self.lineEdit_wear_FSB, self.lineEdit_r1_FSB, self.lineEdit_r1s_FSB,
                     self.lineEdit_stationary_FSTEK, self.lineEdit_carry_FSTEK, self.lineEdit_wear_FSTEK,
                     self.lineEdit_r1_FSTEK, self.lineEdit_path_parser, self.lineEdit_path_original_extract,
                     self.lineEdit_conclusion, self.lineEdit_protocol, self.lineEdit_prescription,
                     self.lineEdit_path_original_file, self.lineEdit_path_finish_folder,
                     self.lineEdit_path_freq_restrict, self.lineEdit_complect_quant_pemi,
                     self.lineEdit_complect_number_pemi, self.lineEdit_path_file_HFE,
                     self.lineEdit_complect_quant_HFE, self.lineEdit_frequency, self.lineEdit_level,
                     self.lineEdit_path_file_HFI, self.lineEdit_complect_quant_HFI, self.lineEdit_imposition_freq]
        self.default_date(data)

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
        if num in [5, 9]:  # Если необходимо открыть файл
            directory = QFileDialog.getOpenFileName(self, "Find Files", QDir.currentPath())
        elif num in [1, 2, 3, 4, 6, 7, 8, 10]:  # Если необходимо открыть директорию
            directory = QFileDialog.getExistingDirectory(self, "Find Files", QDir.currentPath())
        # Список линий
        line = [self.lineEdit_path_check, self.lineEdit_path_parser, self.lineEdit_path_original_file,
                self.lineEdit_path_finish_folder, self.lineEdit_path_freq_restrict, self.lineEdit_path_file_HFE,
                self.lineEdit_path_file_HFI, self.lineEdit_path_original_extract,
                self.lineEdit_path_example, self.lineEdit_path_finish_folder_example]
        if directory:  # Если нажать кнопку отркыть в диалоге выбора
            if num in [5, 9]:  # Если файлы
                if directory[0]:  # Если есть файл, чтобы не очищалось поле
                    line[num - 1].setText(directory[0])
            else:  # Если директории
                line[num - 1].setText(directory)

    def copy_application(self):
        application = check_application_data(self.lineEdit_path_example, self.lineEdit_path_finish_folder_example,
                                             self.lineEdit_number_position, self.lineEdit_quantity_document)
        if type(application) == list:
            self.on_message_changed(application[0], application[1])
            return
        else:  # Если всё прошло запускаем поток
            application['logging'], application['q'] = logging, self.q
            self.thread = GenerateCopyApplication(application)
            self.thread.progress.connect(self.progressBar.setValue)
            self.thread.status.connect(self.show_mess)
            self.thread.messageChanged.connect(self.on_message_changed)
            self.thread.errors.connect(self.errors)
            self.thread.start()

    def generate_pemi(self):
        generate = check_generation_data(self.lineEdit_path_original_file, self.lineEdit_path_finish_folder,
                                         self.lineEdit_complect_number_pemi, self.lineEdit_complect_quant_pemi,
                                         self.checkBox_freq_restrict.isChecked(), self.lineEdit_path_freq_restrict)
        if type(generate) == list:
            self.on_message_changed(generate[0], generate[1])
            return
        else:  # Если всё прошло запускаем поток
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
        if type(generate) == list:
            self.on_message_changed(generate[0], generate[1])
            return
        else:  # Если всё прошло запускаем поток
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
        if type(generate) == list:
            self.on_message_changed(generate[0], generate[1])
            return
        else:  # Если всё прошло запускаем поток
            generate['logging'], generate['q'] = logging, self.q
            self.thread = HFIGeneration(generate)
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
        group_file = True if self.checkBox_group_parcing.isChecked() else False
        folder = checked_file_parcing(self.lineEdit_path_parser, group_file)
        if type(folder) == list:
            self.on_message_changed(folder[0], folder[1])
            return
        else:  # Если всё прошло запускаем поток
            folder['group_file'], folder['logging'], folder['q'] = group_file, logging, self.q
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
        department = True if self.groupBox_FSB.isChecked() else False
        win_lin = True if self.checkBox_win_lin.isChecked() else False
        one_table = True if self.checkBox_first_table.isChecked() else False
        zone = [self.lineEdit_stationary_FSB, self.lineEdit_carry_FSB, self.lineEdit_wear_FSB, self.lineEdit_r1_FSB,
                self.lineEdit_r1s_FSB] \
            if department else [self.lineEdit_stationary_FSTEK, self.lineEdit_carry_FSTEK,
                                self.lineEdit_wear_FSTEK, self.lineEdit_r1_FSTEK]
        zone_all = checked_zone_checked(self.lineEdit_path_check, self.lineEdit_table_number, zone)
        if type(zone_all) == list:
            self.on_message_changed(zone_all[0], zone_all[1])
            return
        else:  # Если всё прошло запускаем поток
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
        output = checked_delete_header_footer(self.lineEdit_path_original_extract, self.lineEdit_conclusion,
                                              self.lineEdit_protocol, self.lineEdit_prescription)
        if type(output) == list:
            self.on_message_changed(output[0], output[1])
            return
        else:  # Если всё прошло запускаем поток
            output['logging'], output['q'] = logging, self.q
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
