import copy
import os
import re
import shutil
import threading
import traceback
import zipfile

from PyQt5.QtCore import QThread, pyqtSignal
from docx.shared import Pt
from lxml import etree
import docx
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT


class DeleteHeaderFooter(QThread):
    progress = pyqtSignal(int)  # Сигнал для прогресс бара
    status = pyqtSignal(str)  # Сигнал для статус бара
    messageChanged = pyqtSignal(str, str)
    errors = pyqtSignal()

    def __init__(self, incoming_data):  # Список переданных элементов.
        QThread.__init__(self)
        self.path = incoming_data['path']
        self.name_executor = incoming_data['name_executor']
        self.logging = incoming_data['logging']
        self.q = incoming_data['q']
        self.event = threading.Event()

    def run(self):
        try:
            current_progress = 0
            self.logging.info('Начинаем обезличивать документы')
            self.status.emit('Старт')
            self.progress.emit(current_progress)
            percent = 100/len(os.listdir(self.path))
            for element in os.listdir(self.path):
                self.logging.info('Обезличиваем документ ' + element)
                self.status.emit('Обезличиваем документ ' + element)
                temp_docx = os.path.join(self.path + '\\', element)
                temp_zip = os.path.join(self.path + '\\', element + ".zip")
                temp_folder = os.path.join(self.path + '\\', "template")
                if os.path.exists(temp_zip):
                    shutil.rmtree(temp_zip)
                if os.path.exists(temp_folder):
                    shutil.rmtree(temp_folder)
                if os.path.exists(self.path + '\\zip'):
                    shutil.rmtree(self.path + '\\zip')
                os.rename(temp_docx, temp_zip)
                os.mkdir(self.path + '\\zip')
                self.logging.info('Извлекаем архив')
                with zipfile.ZipFile(temp_zip) as my_document:
                    my_document.extractall(temp_folder)
                pages_xml = os.path.join(temp_folder, "word", "document.xml")
                tree = etree.parse(pages_xml)
                root = tree.getroot()
                j = 0
                self.logging.info('Получаем тэг для копирования')
                for i, e in reversed(list(enumerate(root[0]))):
                    if e.tag.rpartition('}')[2] != 'sectPr' and len(e):
                        j = i
                        break
                header = copy.deepcopy(root[0][j][0])
                self.logging.info('Собираем и удаляем архив')
                os.remove(temp_zip)
                shutil.make_archive(temp_zip.replace(".zip", ""), 'zip', temp_folder)
                os.rename(temp_zip, temp_docx)  # rename zip file to docx
                while True:
                    try:
                        shutil.rmtree(temp_folder)
                        shutil.rmtree(self.path + '\\zip')
                        break
                    except OSError as es:
                        self.logging.error(es)
                        self.logging.error(traceback.format_exc())
                        self.logging.info('Ошибка с удалением, пробуем ещё раз')
                self.logging.info('Работаем с ворд документом')
                doc = docx.Document(self.path + '\\' + element)  # Открываем

                def paragraph_del(par):
                    par.text = None
                    p = par._element
                    p.getparent().remove(p)
                    p._p = p._element = None

                for paragraph in doc.sections[0].first_page_header.paragraphs:
                    if 'Экз.' in paragraph.text:
                        paragraph_del(paragraph)
                        break
                    paragraph_del(paragraph)
                # doc.sections[0].first_page_header.paragraphs[0].text = None
                # p = doc.sections[0].first_page_header.paragraphs[0]._element
                # p.getparent().remove(p)
                # p._p = p._element = None
                number = doc.sections[0].first_page_footer.paragraphs[0].text
                doc.sections[0].first_page_footer.paragraphs[0].text = None
                doc.sections[0].footer.paragraphs[0].text = None
                if doc.sections[len(doc.sections) - 1].different_first_page_header_footer:
                    for paragraph in doc.sections[len(doc.sections) - 1].first_page_footer.paragraphs:
                        if re.findall(r'\d{2}\.\d{2}\.\d{4}', paragraph.text):
                            date = re.findall(r'\d{2}\.\d{2}\.\d{4}', paragraph.text)
                        if 'Б/ч' in paragraph.text:
                            paragraph_del(paragraph)
                            break
                        paragraph_del(paragraph)
                    doc.sections[len(doc.sections) - 1].first_page_header.is_linked_to_previous = False  # Хидер
                    doc.sections[len(doc.sections) - 1].first_page_footer.is_linked_to_previous = False  # Футер
                else:
                    for paragraph in doc.sections[len(doc.sections) - 1].footer.paragraphs:
                        if re.findall(r'\d{2}\.\d{2}\.\d{4}', paragraph.text):
                            date = re.findall(r'\d{2}\.\d{2}\.\d{4}', paragraph.text)
                        if 'Б/ч' in paragraph.text:
                            paragraph_del(paragraph)
                            break
                        paragraph_del(paragraph)
                    # date = re.findall(r'\d{2}\.\d{2}\.\d{4}',
                    #                   doc.sections[len(doc.sections) - 1].footer.paragraphs[0].text)
                # doc.sections[len(doc.sections) - 1].first_page_footer.paragraphs[0].text = None
                doc.save(self.path + '\\' + element)
                doc = docx.Document(self.path + '\\' + element)  # Открываем
                self.logging.info('Переименовываем заголовок')
                name = element.rpartition('.')[0]
                if 'Заключение СП' in name:
                    name = re.sub('Заключение СП', 'Заключение', name)
                for i, paragraph in enumerate(doc.paragraphs):
                    if re.findall(r'«\d{2}»\s\w+\s\d{4}\s[г]\.', paragraph.text):
                        paragraph.text = re.sub(r'«\d{2}»\s\w+\s\d{4}\s[г]\.', 'date', paragraph.text)
                    if re.findall(name.partition(' ')[0].upper(), paragraph.text):
                        if paragraph.runs[0].font.size is None:
                            size = paragraph.style.font.size.pt
                        else:
                            size = paragraph.runs[0].font.size.pt
                        name_doc = name.partition(' ')[0] if name.partition(' ')[0] else name.partition('.')[0]
                        if name_doc.lower() == 'протокол':
                            name_doc += 'а'
                        else:
                            name_doc = list(name_doc)
                            name_doc[-1] = 'я'
                            name_doc = ''.join(name_doc)
                        paragraph.text = re.sub(name.partition(' ')[0].upper(),
                                                'ВЫПИСКА ИЗ ' + name_doc.upper(),
                                                paragraph.text)
                        doc.paragraphs[i + 1].insert_paragraph_before('\tУч. № ' + number + ' от ' + date[0])
                        for j in [i, i + 1]:
                            doc.paragraphs[j].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                            for run in doc.paragraphs[j].runs:
                                run.font.bold = True
                                run.font.size = Pt(size)
                        break
                if 'Заключение' in name:
                    name_executor = self.name_executor[0]
                elif 'Протокол' in name:
                    name_executor = self.name_executor[1]
                else:
                    name_executor = self.name_executor[2]
                for e in reversed(list(doc.paragraphs)):
                    if 'специальных' in e.text:
                        text = e.text.strip()
                        if e.runs[0].font.size is None:
                            size = e.style.font.size.pt
                        else:
                            size = e.runs[0].font.size.pt
                        name_person = text.rpartition(' ')[2]
                        name_person = text[len(text) - (5 + len(name_person)):len(text)]
                        e.add_run('\n\nВыписка верна\n\n' + re.sub(name_person, name_executor, text.strip()))
                        for run in e.runs:
                            run.font.size = Pt(size)
                        break
                doc.save(self.path + '\\' + element)
                doc = docx.Document(self.path + '\\' + element)  # Открываем
                self.logging.info('Удаляем лишние параграфы в конце документа')
                sectPrs = doc._element.xpath(".//w:pPr/w:sectPr")
                for sectPr in sectPrs:
                    sectPr.getparent().remove(sectPr)
                for para in doc.paragraphs[len(doc.paragraphs):0:-1]:
                    if len(para.text) in [0, 1]:
                        p = para._element
                        p.getparent().remove(p)
                        p._p = p._element = None
                    else:
                        break
                doc.save(self.path + '\\' + element)
                self.logging.info('Извлекаем архив и добавляем шапку')
                os.rename(temp_docx, temp_zip)
                os.mkdir(self.path + '\\zip')
                with zipfile.ZipFile(temp_zip) as my_document:
                    my_document.extractall(temp_folder)
                pages_xml = os.path.join(temp_folder, "word", "document.xml")
                tree = etree.parse(pages_xml)
                root = tree.getroot()
                root[0][-1].append(header)
                tree.write(os.path.join(temp_folder, "word", "document.xml"), encoding="UTF-8", xml_declaration=True)
                self.logging.info('Удаляем архив')
                os.remove(temp_zip)
                shutil.make_archive(temp_zip.replace(".zip", ""), 'zip', temp_folder)
                os.rename(temp_zip, temp_docx)  # rename zip file to docx
                shutil.rmtree(temp_folder)
                shutil.rmtree(self.path + '\\zip')
                doc = docx.Document(self.path + '\\' + element)  # Открываем
                self.logging.info('Удаляем лишние параграфы в шапке')
                for para in reversed(list(doc.sections[0].first_page_header.paragraphs)):
                    if len(para.text) in [0, 1]:
                        p = para._element
                        p.getparent().remove(p)
                        p._p = p._element = None
                    else:
                        break
                doc.save(self.path + '\\' + element)
                current_progress += percent
                self.progress.emit(current_progress)
            self.logging.info("Конец работы программы")
            self.progress.emit(100)
            self.status.emit('Готово')
            os.chdir('C:\\')
            return
        except BaseException as es:
            self.logging.error(es)
            self.logging.error(traceback.format_exc())
            self.progress.emit(0)
            self.status.emit('Ошибка!')
            os.chdir('C:\\')
            return

