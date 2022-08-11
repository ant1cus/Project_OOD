import queue
import re
import sys
import os

from PyQt5.QtGui import QSyntaxHighlighter, QTextCharFormat, QColor

import Main
import logging
from PyQt5.QtCore import (QTranslator, QLocale, QLibraryInfo, QDir)
from PyQt5.QtWidgets import (QMainWindow, QApplication, QFileDialog, QMessageBox)
from checked import checked_zone_checked, file_parcing_checked, check_generation_data
from Zone_Check import ZoneChecked
from File_Parcing import FileParcing
from Generation_Files import GenerationFile


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
        self.pushButton_open_file_vcho.clicked.connect((lambda: self.browse(6)))
        self.pushButton_open_file_vchn.clicked.connect((lambda: self.browse(7)))
        self.groupBox_FSB.clicked.connect(self.group_box_change_state)
        self.groupBox_FSTEK.clicked.connect(self.group_box_change_state)
        self.pushButton_stop.clicked.connect(self.pause_thread)
        self.pushButton_check.clicked.connect(self.checked_zone)
        self.pushButton_parser.clicked.connect(self.parcing_file)
        self.pushButton_generation_pemi.clicked.connect(self.generate_pemi)

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
        if num in [5]:  # Если необходимо открыть файл
            directory = QFileDialog.getOpenFileName(self, "Find Files", QDir.currentPath())
        elif num in [1, 2, 3, 4, 6, 7]:  # Если необходимо открыть директорию
            directory = QFileDialog.getExistingDirectory(self, "Find Files", QDir.currentPath())
        # Список линий
        line = [self.lineEdit_path_check, self.lineEdit_path_parser, self.lineEdit_path_original_file,
                self.lineEdit_path_finish_folder, self.lineEdit_path_freq_restrict, self.lineEdit_path_file_vcho,
                self.lineEdit_path_file_vchn]
        if directory:  # Если нажать кнопку отркыть в диалоге выбора
            if num in [5]:  # Если файлы
                if directory[0]:  # Если есть файл, чтобы не очищалось поле
                    line[num - 1].setText(directory[0])
            else:  # Если директории
                line[num - 1].setText(directory)

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
            self.thread.start()

    def parcing_file(self):
        self.plainTextEdit_succsess_order.clear()
        self.groupBox_succsess_order.setStyleSheet("")
        self.plainTextEdit_error_order.clear()
        self.groupBox_error_order.setStyleSheet("")
        self.plainTextEdit_errors.clear()
        self.groupBox_errors.setStyleSheet("")
        group_file = True if self.checkBox_group_parcing.isChecked() else False
        folder = file_parcing_checked(self.lineEdit_path_parser, group_file)
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
        if 'errors' in text:
            self.groupBox_errors.setStyleSheet('''
            QGroupBox {border: 0.5px solid;
            border-radius: 5px;
             border-color: yellow;
              padding:10px 0px 0px 0px;}''')
            self.plainTextEdit_errors.insertPlainText('\n'.join(text['errors']))
            highlighter(self.plainTextEdit_errors)

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
            self.thread = ZoneChecked([self.lineEdit_path_check.text().strip(),
                                       self.lineEdit_table_number.text().strip(),
                                       department, win_lin, zone_all, one_table, logging, self.q])
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
