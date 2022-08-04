import os
import threading

import docx
import openpyxl
import re
import traceback
from natsort import natsorted
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from PyQt5.QtCore import QThread, pyqtSignal


class FileParcing(QThread):
    progress = pyqtSignal(int)  # Сигнал для прогресс бара
    status = pyqtSignal(str)  # Сигнал для статус бара
    messageChanged = pyqtSignal(str, str)

    def __init__(self, output):  # Список переданных элементов.
        QThread.__init__(self)
        self.path = output[0]
        self.group_file = output[1]
        self.logging = output[2]
        self.q = output[3]
        self.event = threading.Event()

    def run(self):
        progress = 0
        self.logging.info("Начинаем")
        self.status.emit('Старт')
        self.progress.emit(progress)