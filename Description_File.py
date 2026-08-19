import check_description

from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QDialog, QSizePolicy, QCheckBox


class CheckDescription(QDialog, check_description.Ui_Dialog_description):  # Настройки по умолчанию
    def __init__(self, parent, all_mode):
        super().__init__()
        self.setupUi(self)
        self.all_mode = all_mode
        self.parent = parent
        self.pushButton_description_save.clicked.connect(self.accept)  # Сохранить
        self.pushButton_description_close.clicked.connect(lambda: self.close())  # Без сохранения
        self.pushButton_description_check_all.clicked.connect(lambda: self.all_change(True))
        self.pushButton_description_reset_all.clicked.connect(lambda: self.all_change(False))
        self.check_box = {}  # Для чекбоксов
        check_box_row = 0
        for row in self.all_mode.itertuples():  # Заполняем
            name = 'Мощность ' + row.name if 'power' in row.mode_and_sys[0].lower() or\
                                             'pwr' in row.mode_and_sys[0].lower() else row.name
            self.check_box[row.Index] = QCheckBox(name)  # Помещаем в фрейм
            self.check_box[row.Index].setFont(QFont("Times", 12, QFont.Light))  # Шрифт, размер
            self.check_box[row.Index].setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)  # Размеры виджета
            self.check_box[row.Index].setFixedWidth(1050)
            self.check_box[row.Index].setChecked(row.in_report)
            self.gridLayout_description_check_box.addWidget(self.check_box[row.Index], check_box_row, 0)
            check_box_row += 1

    def all_change(self, state):
        for element in self.check_box:
            self.check_box[element].setChecked(state)

    def accept(self):  # Если нажали кнопку принять
        for element in self.check_box:
            self.all_mode.loc[element, 'in_report'] = self.check_box[element].isChecked()
        self.close()  # Закрываем

    def closeEvent(self, event):
        if self.sender() and self.sender().text() == 'Сохранить':
            event.accept()
            self.parent.all_mode = self.all_mode
            self.parent.show()
        else:
            event.accept()
            self.parent.show()
