import os
import re
import pathlib

import psutil
from openpyxl import load_workbook


def check(n, e):
    f = True
    for el in e:
        if n == el:
            f = False
            return f
    return f


def checked_zone_checked(line_edit_path_check, line_edit_table_number, zone):

    for proc in psutil.process_iter():
        if proc.name() == 'WINWORD.EXE':
            return ['УПС!', 'Закройте все файлы Word!']
    path = line_edit_path_check.text().strip()
    if not path:
        return ['УПС!', 'Путь к проверяемым документам пуст']
    if os.path.isdir(path):
        pass
    else:
        return ['УПС!', 'Указанный путь к проверяемым документам не является директорией']
    table = line_edit_table_number.text().strip()
    if not table:
        return ['УПС!', 'Не указан номер таблицы']
    for i in table:
        if check(i, ('1', '2', '3', '4', '5', '6', '7', '8', '9', '0')):
            return ['УПС!', 'Есть лишние символы в номере таблицы']
    zone_out = {i: el.text().replace(',', '.') for i, el in enumerate(zone) if el.text()}
    err_msg = ('Есть лишние символы в ограничении по стационарным антеннам',
               'Есть лишние символы в ограничении по возимым антеннам',
               'Есть лишние символы в ограничении по носимым антеннам',
               'Есть лишние символы в ограничении r1',
               'Есть лишние символы в ограничении r1`')
    if zone_out:
        for i in zone_out:
            for j in zone_out[i]:
                if check(j, ('1', '2', '3', '4', '5', '6', '7', '8', '9', '0', '.')):
                    return ['УПС!', err_msg[i]]
    else:
        return ['УПС!', 'Не указано ни одно ограничение для проверки']
    return zone_out


def checked_file_parcing(dir_path, group_file):
    def folder_checked(p):
        errors = []
        txt_files = list(filter(lambda x: x.endswith('.txt'), os.listdir(p)))
        excel_files = [x for x in os.listdir(p) if x.endswith('.xlsx')]
        if 'Описание.txt' not in txt_files and 'описание.txt' not in txt_files:
            errors.append('Нет файла с описанием режимов (' + p + ')')
        else:
            with open(p + '\\Описание.txt', mode='r') as f:
                lines = f.readlines()
                for line in lines:
                    if re.findall(r'\s', line.rstrip('\n')):
                        errors.append('Пробелы в названии режимов (' + p + ', ' + line.rstrip('\n') + ')')
        return {'errors': errors, 'len': len(excel_files)}
    # Выбираем путь для исходников.
    path = dir_path.text().strip()
    if not path:
        return ['УПС!', 'Не указан путь к исходной папке']
    elif os.path.isfile(path):
        return ['УПС!', 'Указанный путь к исходным файлам не является директорией']
    else:
        folders = [i for i in os.listdir(path) if os.path.isdir(path + '\\' + i) and i != 'txt']
        if group_file is False and folders:
            return ['УПС!', 'В директории для парсинга присутствуют папки']
        elif group_file and folders is False:
            return ['УПС!', 'В директории для парсинга нет папок для преобразования']
    error = []
    progress = 0
    if group_file:
        for folder in os.listdir(path):
            if os.path.isdir(path + '\\' + folder):
                err = folder_checked(path + '\\' + folder)
                progress += err['len']
                if err['errors']:
                    error.append(err)
    else:
        err = folder_checked(path)
        error, progress = err['errors'], err['len']
    return ['УПС!', '\n'.join(error)] if error else {'path': path, 'progress': progress}


def checked_generation_pemi(source_file, output_file, complect_number, complect_quantity, freq_restrict,
                            freq_restrict_path):

    source = source_file.text().strip()
    if not source:
        return ['УПС!', 'Путь к исходным файлам пуст']
    if os.path.isfile(source):
        return ['УПС!', 'Указанный путь к исходным файлам не является директорией']
    output = output_file.text().strip()
    if not output:
        return ['УПС!', 'Путь к создаваемым файлам пуст']
    if os.path.isfile(output):
        return ['УПС!', 'Указанный путь к создаваемым файлам не является директорией']
    if len([True for el in pathlib.Path(source).iterdir() if el.is_dir()]) > 1:
            return ['УПС!', 'В указанной директории слишком много папок']
    complect_num = complect_number.text().strip()
    if not complect_num:
        return ['УПС!', 'Не указаны номера комплектов']
    for i in complect_num:
        if check(i, ('1', '2', '3', '4', '5', '6', '7', '8', '9', '0', ' ', '-', ',', '.')):
            return ['УПС!', 'Есть лишние символы в номерах комплектов']
    complect_num = complect_num.replace(' ', '').replace(',', '.')
    if complect_num[0] == '.' or complect_num[0] == '-':
        return ['УПС!', 'Первый символ введён не верно']
    if complect_num[-1] == '.' or complect_num[-1] == '-':
        return ['УПС!', 'Последний символ введён не верно']
    for i in range(len(complect_num)):
        if complect_num[i] == '.' or complect_num[i] == '-':
            if complect_num[i + 1] == '.' or complect_num[i + 1] == '-':
                return ['УПС!', 'Два разделителя номеров подряд']
    complect = []
    for element in complect_num.split('.'):
        if '-' in element:
            num1, num2 = int(element.partition('-')[0]), int(element.partition('-')[2])
            if num1 >= num2:
                return ['УПС!', 'Диапазон номеров комплектов указан не верно']
            else:
                for el in range(num1, num2 + 1):
                    complect.append(el)
        else:
            complect.append(element)
    complect.sort()
    if len(complect) != len(set(complect)):
        return ['УПС!', 'Есть повторения в номерах комплектов']
    complect_quant = complect_quantity.text()
    if not complect_quant:
        return ['УПС!', 'Не указано количество комплектов']
    for i in complect_quant:
        if check(i, ('1', '2', '3', '4', '5', '6', '7', '8', '9', '0')):
            return ['УПС!', 'Не правильно указано количество комплектов']
    if len(complect) != int(complect_quant):
        return ['УПС!', 'Указанные номера не совпадают с количеством генерируемых комплектов']
    txt_files = list(filter(lambda x: x.endswith('.txt'), os.listdir(source)))
    if len(txt_files) != 1:
        return ['УПС!', 'В папке больше одного или нет txt файла']
    name_mode = ''
    for file in sorted(txt_files):
        if os.stat(source + '\\' + file).st_size == 0:
            return ['УПС!', 'Файл с описанием режимов пуст']
        else:
            try:
                with open(str(pathlib.Path(source, file)), mode='r', encoding='utf-8') as f:
                    name_mode = f.readlines()
                    if ' ' in name_mode:
                        return ['УПС!', 'Пробелы в названии режимов в файле «Описание»']
                    else:
                        name_mode = [line.rstrip().lower() for line in name_mode if line]
            except UnicodeDecodeError:
                with open(str(pathlib.Path(source, file)), mode='r') as f:
                    name_mode = f.readlines()
                    if ' ' in name_mode:
                        return ['УПС!', 'Пробелы в названии режимов в файле «Описание»']
                    else:
                        name_mode = [line.rstrip().lower() for line in name_mode if line]
    error = []
    for file_exel in sorted(list(filter(lambda x: x.endswith('.xlsx'), os.listdir(output)))):
        if int(file_exel[:-5]) in complect:
            error.append(file_exel)
    if error:
        return ['УПС!', 'В указанной папке уже есть такие исходники:\n' + '\n'.join(error)]
    error = []
    for file_exel in sorted(list(filter(lambda x: x.endswith('.xlsx'), os.listdir(source)))):
        wb = load_workbook(str(pathlib.Path(source, file_exel)))  # Откроем книгу.
        name = wb.sheetnames  # Список листов.
        for name_exel in name:
            if name_exel.lower() not in name_mode:
                error.append('Названия режимов в файле ' + file_exel + ' не совпадают с описанием.')
                break
        wb.close()
    if error:
        return ['УПС!', '\n'.join(error)]
    # Добавление для разницы!
    restrict_file = False
    if freq_restrict:
        restrict_file = freq_restrict_path.text().strip()
        if not restrict_file:
            return ['УПС!', 'Путь к файлу с ограничениями пуст']
        if os.path.isfile(restrict_file):
            if restrict_file.endswith('.txt'):
                if os.stat(restrict_file).st_size == 0:
                    return ['УПС!', 'Файл для комплектов пуст']
            else:
                return ['УПС!', 'Файл для комплектов не txt']
        else:
            return ['УПС!', 'Указана директория в файле для ограничений']
    return {'source': source, 'output': output, 'complect': complect, 'complect_quant': complect_quant,
            'name_mode': name_mode, 'restrict_file': restrict_file}


def checked_delete_header_footer(path, conclusion_post, conclusion_name, protocol_post, protocol_name,
                                 prescription_post, prescription_name):
    source = path.text().strip()
    if not source:
        return ['УПС!', 'Путь к исходным файлам пуст']
    if os.path.isfile(source):
        return ['УПС!', 'Указанный путь к исходным файлам не является директорией']
    name_file = [False, False, False]
    error = []
    for file in os.listdir(source):
        if file.endswith('.doc'):
            error.append(file)
        if 'заключение' in file.lower():
            name_file[0] = True
        elif 'протокол' in file.lower():
            name_file[1] = True
        elif 'предписание' in file.lower():
            name_file[2] = True
    if error:
        return ['УПС!', 'Файлы старого формата:\n' + '\n'.join(error)]
    post_concl, post_prot, post_pre = conclusion_post.text().strip(), protocol_post.text().strip(), prescription_post.text().strip()
    name_concl, name_prot, name_pre = conclusion_name.text().strip(), protocol_name.text().strip(), prescription_name.text().strip()
    name_rus = ['заключении', 'протоколе', 'предписании']
    error = []
    if any([name_concl, name_prot, name_pre]):
        for name, el_name, el_post in zip(name_file, [name_concl, name_prot, name_pre],
                                          [post_concl, post_prot, post_pre]):
            if name:
                if re.match(r'[А-Я]\.[А-Я]\.\s[А-Я][а-я]+', el_name) is None:
                    error.append('Имя ' + el_name + ' в ' + name_rus[[name_concl, name_prot, name_pre].index(el_name)]
                                 + ' написано неверно (Шаблон имени «И.И. Иванов»)')
                if len(el_post) == 0:
                    error.append('Не указана должность в ' + name_rus[[post_concl, post_prot, post_pre].index(el_post)])
        if error:
            return ['УПС!', '\n'.join(error)]
        return {'path': source, 'post_executor': [post_concl, post_prot, post_pre],
                'name_executor': [name_concl, name_prot, name_pre]}
    else:
        return ['УПС!', 'Не указано ни одно имя']


def checked_hfe_generation(path, complect, req_val, frequency, value):
    output = path.text().strip()
    if not output:
        return ['УПС!', 'Не указан путь к конечной папке']
    if os.path.isfile(output):
        return ['УПС!', 'Указанный путь не является директорией']
    if req_val.isChecked():
        freq = frequency.text().strip()
        val = value.text().strip()
    else:
        freq = 100
        val = 100
    quantity = complect.text().strip()
    variables = [freq, val, quantity]
    errors1 = ['Не указана частота', 'Не указан уровень', 'Не указано количество комплектов']
    errors2 = ['Частота указана с ошибкой', 'Уровень указан с ошибкой', 'Количество комплектов указано с ошибкой']
    for i in range(0, 3, 1):
        if not variables[i]:
            return ['УПС!', errors1[i]]
        for j in variables[i]:
            if check(j, ('1', '2', '3', '4', '5', '6', '7', '8', '9', '0')):
                return ['УПС!', errors2[i]]
    return {'path': output, 'quantity': int(quantity), 'freq': freq, 'val': val}


def checked_hfi_generation(path, frequency, quantity, mode):
    output = path.text().strip()
    if not output:
        return ['УПС!', 'Не указан путь к конечной папке']
    if os.path.isfile(output):
        return ['УПС!', 'Указанный путь не является директорией']
    variables = [frequency.text().strip(), quantity.text().strip()]
    errors1 = ['Не указана частота навязывания', 'Не указано количество комплектов']
    errors2 = ['Частота навязывания указана с ошибкой', 'Количество комплектов указано с ошибкой']
    for i in range(0, 2, 1):
        if not variables[i]:
            return ['УПС!', errors1[i]]
        for j in variables[i]:
            if check(j, ('1', '2', '3', '4', '5', '6', '7', '8', '9', '0')):
                return ['УПС!', errors2[i]]
    if any(mode):
        return {'path': output, 'quantity': int(quantity), 'freq': variables[0], 'val': variables[1], 'mode': mode}
    else:
        return ['УПС!', 'Не выбран не один режим']


def checked_application_data(file_path, folder_path, number, quant):
    file = file_path.text().strip()
    if not file:
        return ['УПС!', 'Не указан путь к файлу для генерации']
    if os.path.isdir(file):
        return ['УПС!', 'Указанный путь не является файлом']
    else:
        if file.endswith('.docx'):
            pass
        else:
            return ['УПС!', 'Указанный файл неверного формата (необходим docx)']
    folder = folder_path.text().strip()
    if not folder:
        return ['УПС!', 'Не указан путь к конечной папке для генерации']
    if os.path.isfile(folder):
        return ['УПС!', 'Указанный путь не является директорией']
    position_num = number.text().strip()
    if not position_num:
        return ['УПС!', 'Не указан номер позиции']
    for i in position_num:
        if check(i, ('1', '2', '3', '4', '5', '6', '7', '8', '9', '0')):
            return ['УПС!', 'Не правильно указан номер позиции']
    quantity = quant.text().strip()
    if not quantity:
        return ['УПС!', 'Не указано количество комплектов']
    for i in quantity:
        if check(i, ('1', '2', '3', '4', '5', '6', '7', '8', '9', '0')):
            return ['УПС!', 'Не правильно указано количество комплектов']
    return {'file': file, 'path': folder, 'quantity': int(quantity), 'position_num': int(position_num)}


def checked_lf_data(source_folder, output_folder, excel_file):
    source = source_folder.text().strip()
    if not source:
        return ['УПС!', 'Путь к исходным файлам пуст']
    if os.path.isfile(source):
        return ['УПС!', 'Указанный путь к исходным файлам не является директорией']
    output = output_folder.text().strip()
    if not output:
        return ['УПС!', 'Путь к создаваемым файлам пуст']
    if os.path.isfile(output):
        return ['УПС!', 'Указанный путь к создаваемым файлам не является директорией']
    file = excel_file.text().strip()
    if not file:
        return ['УПС!', 'Путь к файлу генератору пуст']
    else:
        if file.endswith('.xlsx'):
            return {'source': source, 'output': output, 'excel': file}
        else:
            return ['УПС!', 'Указанный файл недопустимого формата (необходимо .xlsx)']


def checked_generation_cc(start_folder, finish_folder, set_number, checkbox_frequency, frequency, checkbox_txt):

    source = start_folder.text().strip()
    if not source:
        return ['УПС!', 'Путь к исходным файлам пуст']
    if os.path.isfile(source):
        return ['УПС!', 'Указанный путь к исходным файлам не является директорией']
    output = finish_folder.text().strip()
    if not output:
        return ['УПС!', 'Путь к создаваемым файлам пуст']
    if os.path.isfile(output):
        return ['УПС!', 'Указанный путь к создаваемым файлам не является директорией']
    if len([True for el in pathlib.Path(source).iterdir() if el.is_dir()]) > 1:
            return ['УПС!', 'В указанной директории слишком много папок']
    set_num = set_number.text().strip()
    if set_num:
        for i in set_num:
            if check(i, ('1', '2', '3', '4', '5', '6', '7', '8', '9', '0', ' ', '-', ',', '.')):
                return ['УПС!', 'Есть лишние символы в номерах комплектов']
        set_num = set_num.replace(' ', '').replace(',', '.')
        if set_num[0] == '.' or set_num[0] == '-':
            return ['УПС!', 'Первый символ введён не верно']
        if set_num[-1] == '.' or set_num[-1] == '-':
            return ['УПС!', 'Последний символ введён не верно']
        for i in range(len(set_num)):
            if set_num[i] == '.' or set_num[i] == '-':
                if set_num[i + 1] == '.' or set_num[i + 1] == '-':
                    return ['УПС!', 'Два разделителя номеров подряд']
        set_list = []
        for element in set_num.split('.'):
            if '-' in element:
                num1, num2 = int(element.partition('-')[0]), int(element.partition('-')[2])
                if num1 >= num2:
                    return ['УПС!', 'Диапазон номеров комплектов указан не верно']
                else:
                    for el in range(num1, num2 + 1):
                        set_list.append(el)
            else:
                set_list.append(element)
        set_list.sort()
        if len(set_list) != len(set(set_list)):
            return ['УПС!', 'Есть повторения в номерах комплектов']
        if len(set_list) == 1 and set_list[0] == '0':
            set_num = False
        else:
            set_num = set_list
    freq = frequency.text() if checkbox_frequency.isChecked() else ''
    if freq:
        for i in freq:
            if check(i, ('1', '2', '3', '4', '5', '6', '7', '8', '9', '0', '.')):
                return ['УПС!', 'Не правильно указана частота для ограничения (дробный разделитель - ".")']
    txt = checkbox_txt.isChecked()
    return {'source': source, 'output': output, 'set': set_num, 'frequency': freq, 'only_txt': txt}
