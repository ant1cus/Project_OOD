import pathlib
import json


def rewrite(path, data, widget=False, visible=False, order=False):
    if widget:
        dict_load = data
    else:
        name = visible if visible else order
        with open(pathlib.Path(path, 'Настройки.txt'), "r", encoding='utf-8-sig') as f:  # Открываем
            dict_load = json.load(f)  # Загружаем данные
            dict_load['gui_settings'][name] = data
    with open(pathlib.Path(path, 'Настройки.txt'), 'w', encoding='utf-8-sig') as f:  # Пишем в файл
        json.dump(dict_load, f, ensure_ascii=False, sort_keys=True, indent=4)
