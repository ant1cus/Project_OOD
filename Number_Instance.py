import os
import pathlib
import re
import shutil
import threading
import traceback

from PyQt5.QtCore import QThread, pyqtSignal
from docx.shared import Pt
import docx


class ChangeNumberInstance(QThread):
    progress = pyqtSignal(int)  # Сигнал для progress bar
    status = pyqtSignal(str)  # Сигнал для статус бара
    messageChanged = pyqtSignal(str, str)
    errors = pyqtSignal()

    def __init__(self, incoming_data):  # Список переданных элементов.
        QThread.__init__(self)
        self.path_old = incoming_data['path_old']
        self.path_new = incoming_data['path_new']
        self.set_number = incoming_data['set_number']
        self.logging = incoming_data['logging']
        self.queue = incoming_data['queue']
        self.default_path = incoming_data['default_path']
        self.event = threading.Event()

    def run(self):
        try:
            current_progress = 0
            self.logging.info('Начинаем менять номера комплектов')
            self.status.emit('Старт')
            self.progress.emit(current_progress)
            percent = 100/(len(os.listdir(self.path_old))*len(self.set_number))
            for number_folder in self.set_number:
                os.mkdir(pathlib.Path(self.path_new, str(number_folder) + ' экземпляр'))
                for doc in os.listdir(self.path_old):
                    if doc.endswith('.docx'):
                        shutil.copy2(pathlib.Path(self.path_old, doc),
                                     pathlib.Path(self.path_new, str(number_folder) + ' экземпляр'))
                        doc_2 = docx.Document(pathlib.Path(self.path_new, str(number_folder) + ' экземпляр', doc))
                        for p_2 in doc_2.sections[0].first_page_header.paragraphs:
                            if re.findall(r'№1', p_2.text):
                                text = re.sub(r'№1', '№' + str(number_folder), p_2.text)
                                p_2.text = text
                                for run in p_2.runs:
                                    run.font.size = Pt(11)
                                    run.font.name = 'Times New Roman'
                                break
                        for p_2 in doc_2.sections[len(doc_2.sections) - 1].first_page_footer.paragraphs:
                            if re.findall(r'Отп. 1 экз. в адрес', p_2.text):
                                text = re.sub(r'Отп. 1 экз. в адрес', 'Отп. ' + str(number_folder) + ' экз. в адрес',
                                              p_2.text)
                                p_2.text = text
                                for run in p_2.runs:
                                    run.font.size = Pt(11)
                                    run.font.name = 'Times New Roman'
                                break
                        doc_2.save(pathlib.Path(self.path_new, str(number_folder) + ' экземпляр', doc))  # Сохраняем
                    current_progress += percent
                    self.progress.emit(int(current_progress))
            self.logging.info("Конец работы программы")
            self.progress.emit(100)
            self.status.emit('Готово')
            os.chdir(self.default_path)
            return
        except BaseException as es:
            self.logging.error(es)
            self.logging.error(traceback.format_exc())
            self.progress.emit(0)
            self.status.emit('Ошибка!')
            os.chdir(self.default_path)
            return

