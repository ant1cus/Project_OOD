import os
import re
import traceback
from pathlib import Path

import pandas as pd
import numpy as np


def file_parcing(path, logging, line_doing, now_doc, all_doc, line_progress, progress, per, cp,
                 default_path, event, window_check, twelve_sectors=False, no_freq_lim=False, difference_3=False,
                 two_percent=False, video_check=False, del_frq_check=False, del_frq=0):
    try:
        name_dir = Path(path).name
        pat = ['_ЦП', '.m', '.v']  # список ключевых слов для поиска в ЦП
        mode = []
        errors = []
        errors_continue = []  # Для того, чтобы при ошибках в 2% распарсить файлы
        logging.info(f"Читаем txt, сохраняем режимы, ищем подходящие excel в {name_dir}")
        for file in Path(path).glob('*.txt'):
            try:
                with open(file, mode='r', encoding="utf-8-sig") as f:
                    logging.info("Кодировка utf-8-sig")
                    mode_1 = f.readlines()
                    mode_1 = [line.rstrip() for line in mode_1]
            except UnicodeDecodeError:
                with open(file, mode='r') as f:
                    logging.info("Другая кодировка")
                    mode_1 = f.readlines()
                    mode_1 = [line.rstrip() for line in mode_1]
            mode = [x for x in mode_1 if x]
        for file in Path(path).glob('*.xlsx'):
            try:
                event.wait()
                if window_check.stop_threading:
                    return {'status': 'cancel', 'text': '', 'data': {}}
                if '~' in file.name:
                    continue
                now_doc += 1
                line_doing.emit(f"Проверяем названия рабочих листов в документе {file.name} ({now_doc} из {all_doc})")
                pat_rez = ['_ЦП', '.m', '.v']
                logging.info("Открываем книгу")
                new_book = {}
                book = pd.read_excel(file, sheet_name=None, header=None)
                for enum, name_list in enumerate(book.keys()):
                    x = name_list
                    if re.search(r'_ЦП', name_list) or re.search(r'\.m', name_list) or re.search(r'\.v', name_list):
                        logging.info(f"Нашли название {name_list}")
                        rez = []
                        for y in pat:  # прогоняем список
                            logging.info("Ищем совпадение в нашем списке")
                            if y == '.v':
                                replace = re.findall(r'.v\d', x)
                                if replace:
                                    y = pat_rez[2] = replace[0]
                            rez.append(1) if x.find(y) != -1 else rez.append(-1)  # добавляем заметки для ключевых слов
                            logging.info("Изменяем название")
                            x = x.replace(y, '')  # оставляем только название режима
                        for i in range(0, 3):
                            x = x + pat_rez[i] if rez[i] == 1 else x  # добавляем необходимые ключевые слова
                    new_book[x] = book[name_list]
                name = [x for x in new_book.keys()]
                logging.info("Проверяем на совпадение названий с файлом описания")
                if name != mode:  # проверяем названия на соответствия
                    logging.info("Названия не совпадают")
                    output = f"В заказе {name_dir} названия режимов в исходнике {file} не совпадают с описанием: "
                    for i_out, name_isx in enumerate(name):
                        if mode.count(name_isx) == 0:
                            output += f"{i_out}) режим {name_isx}; "
                    errors.append(output.strip(' '))
                line_doing.emit(f"Проверяем режимы в {file} на ошибки ({now_doc} из {all_doc})")
                logging.info("Проверяем документы на наличие ошибок")
                for sheet in name:  # Загоняем в txt.
                    if sheet.lower() == 'описание':
                        continue
                    df = new_book[sheet]
                    if twelve_sectors:
                        alphabet = [chr(i) for i in range(65, 90)]
                        df = df.fillna(0.0000001)
                        for column in df.columns:
                            data = df[column]
                            try:  # Блок try для отлова текста в значениях
                                if data.astype(float).all():
                                    continue
                                else:
                                    errors.append(f"В заказе {Path(path).name} в исходнике {file} в режиме {sheet} в"
                                                  f" колонке {column + 1} неведомая штука (не преобразовывается ни в"
                                                  f" строку, ни в значение)!")
                            except ValueError:
                                for i, row in enumerate(df[column]):
                                    if type(row) == str:
                                        errors.append(f"В заказе {Path(path).name} в исходнике {file} в режиме {sheet}"
                                                      f" в ячейке {alphabet[column] + str(i + 1)} есть текстовое"
                                                      f" значение!")
                        continue
                    df = df.fillna(False)
                    if df.shape[0] == 1 and df.shape[1] == 1:  # для отлова листов с надписью «не обнаружено»
                        continue
                    two_percent_check = []
                    if del_frq_check:
                        df = df.drop(df[df[0] > del_frq].index.to_list())
                    if two_percent and re.findall(r'dp', sheet, re.I):
                        df_rep = df.replace(False, 0)
                        two_percent_check = (np.where(
                            df_rep[1] != 0, ((df_rep[1] - df_rep[2]) * 100 / df_rep[1]) < 2, False)).tolist()
                    for i in range(df.shape[0]):
                        # Здесь может не попасть в i или брать с названиями
                        frq, s, n = round(df.iloc[i, 0], 4), df.iloc[i, 1], df.iloc[i, 2]
                        if s is False and n is False:
                            continue
                        if isinstance(frq, str):
                            errors.append(f"В заказе «{name_dir}» в исходнике {file.name} в режиме {sheet} в строке "
                                          f"{i + 1} записано текстовое значение!")
                        if s is False:
                            errors.append(f"В заказе «{name_dir}» в исходнике {file.name} в режиме {sheet} на частоте "
                                          f"{frq} есть значение шума, но нет сигнала!")
                        if isinstance(s, str):
                            errors.append(f"В заказе «{name_dir}» в исходнике {file.name} в режиме {sheet} на частоте "
                                          f"{frq} сигнал указан как текстовое значение")
                        if n is False:
                            errors.append(f"В заказе «{name_dir}» в исходнике {file.name} в режиме {sheet} на частоте "
                                          f"{frq} есть значение сигнала, но нет шума!")
                        if isinstance(n, str):
                            errors.append(f"В заказе «{name_dir}» в исходнике {file.name} в режиме {sheet} на частоте "
                                          f"{frq} шум указан как текстовое значение")
                        if isinstance(s, str) or isinstance(n, str) or no_freq_lim:
                            continue
                        if s < n:
                            errors.append(f"В заказе «{name_dir}» в исходнике {file.name} в режиме {sheet} на частоте "
                                          f"{frq} значения шума больше сигнала!")
                        if s == n:
                            errors.append(f"В заказе «{name_dir}» в исходнике {file.name} в режиме {sheet} на частоте "
                                          f"{frq} одинаковые значения сигнала и шума!")
                        if abs(s-n) > 60:
                            errors.append(f"В заказе «{name_dir}» в исходнике {file.name} в режиме {sheet} на частоте "
                                          f"{frq} слишком большая разница между сигналом и шумом!")
                        # Дополнительные проверки (кроме ограничений)
                        if video_check and re.findall('video', sheet, re.I):
                            continue
                        if difference_3 and abs(s-n) < 3.1:
                            errors.append(f"В заказе «{name_dir}» в исходнике {file.name} в режиме {sheet} на частоте "
                                          f"{frq} разница между сигналом и шумом меньше 3.1!")
                        if two_percent and any(two_percent_check):
                            err_str = f"В заказе «{name_dir}» в исходнике {file.name} в режиме {sheet}" \
                                      f" не пройдена проверка 2%!"
                            if err_str not in errors_continue:
                                errors_continue.append(err_str)
                cp += per
                line_progress.emit(f'Выполнено {int(cp)} %')
                progress.emit(int(cp))
            except BaseException as ex:
                logging.error(f"Ошибка при проверке файла {file.name}: {ex}")
                errors.append(f"Ошибка при проверке файла {file.name}, обратитесь к разработчику")

        if errors:
            return {'status': 'errors', 'text': 'Ошибки в файлах',
                    'data': {'errors': errors, 'errors_continue': errors_continue, 'cp': cp, 'now_doc': now_doc}}
        for file in Path(path).glob('*.xlsx'):
            try:
                event.wait()
                if window_check.stop_threading:
                    return {'status': 'cancel', 'text': '', 'data': {}}
                df_for_write = {}
                now_doc += 1
                line_doing.emit(f'Создаем txt файлы для документа {file.name} ({now_doc} из {all_doc})')
                book = pd.read_excel(file, sheet_name=None, header=None)
                sheets = book.keys()
                if os.path.exists(Path(path, 'txt', file.stem)) is False:
                    os.makedirs(Path(path, 'txt', file.stem))
                for sheet in sheets:
                    name_first = sheet.partition('.')[0]
                    name_dot = '.'
                    name_second = sheet.partition('.')[2]
                    if re.findall(r'_lin|_linux', sheet, re.I):
                        name_sheet = name_first.upper()
                    elif re.findall(r'_win|_windows', sheet, re.I):
                        name_sheet = name_first.lower()
                    else:
                        name_sheet = name_first.upper()
                    name_sheet = name_sheet + name_dot + name_second.lower() if name_second else name_sheet
                    df = book[sheet]
                    if df.empty or isinstance(df.iloc[0, 0], str):
                        with open(Path(path, 'txt', file.stem, name_sheet + '.txt'), 'w'):
                            pass
                        continue
                    if del_frq_check:
                        df = df.drop(df[df[0] > del_frq].index.to_list())
                        df_for_write[sheet] = df
                    if sheet.lower() != 'описание':
                        if twelve_sectors is False:
                            df = df.drop(df.columns[[i for i in df.columns.tolist() if i > 2]], axis=1)
                            if [0, 1, 2] in df.columns.tolist():
                                df = df[[0, 1, 2]]
                            df = df.dropna()
                        else:
                            df = df.fillna(0)
                    df = df.round(4)
                    df.to_csv(Path(path, 'txt', file.stem, name_sheet + '.txt'),
                              index=None, sep='\t', mode='w', header=None)
                if del_frq_check:
                    with pd.ExcelWriter(file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        for element in df_for_write:
                            df_for_write[element].to_excel(writer, sheet_name=element, index=False, header=False)
                cp += per
                line_progress.emit(f'Выполнено {int(cp)} %')
                progress.emit(int(cp))
            except BaseException as ex:
                logging.error(f"Ошибка при парсинге файла {file.name}: {ex}")
                errors.append(f"Ошибка при парсинге файла {file.name}, обратитесь к разработчику")
        return {'status': 'errors' if errors or errors_continue else 'success',
                'text': 'Выполнено с ошибками' if errors or errors_continue else 'Успешно',
                'data': {'errors': errors, 'errors_continue': errors_continue, 'cp': cp, 'now_doc': now_doc}}
    # Подумать что тут с исключениями
    except BaseException as es:
        return {'status': 'exception', 'text': 'Функция «file_parcing» завершилась с ошибкой',
                'data': {'trace': traceback.format_exc(), 'exception': es}}
