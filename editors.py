import pandas as pd
from PIL import ImageFont

font = ImageFont.truetype('T-FLEXA.ttf', 12)


width_of_symbols = {'A': 7.0, 'B': 6.0, 'C': 5.0, 'D': 6.0,
                    'E': 5.0, 'F': 5.0, 'G': 6.0, 'H': 6.0,
                    'I': 2.0, 'J': 5.0, 'K': 6.0, 'L': 5.0,
                    'M': 7.0, 'N': 6.0, 'O': 6.0, 'P': 6.0,
                    'Q': 7.0, 'R': 6.0, 'S': 6.0, 'T': 6.0,
                    'U': 6.0, 'V': 7.0, 'W': 9.0, 'X': 7.0,
                    'Y': 7.0, 'Z': 6.0, '0': 6.0, '1': 4.0,
                    '2': 6.0, '3': 5.0, '4': 6.0, '5': 5.0,
                    '6': 6.0, '7': 6.0, '8': 6.0, '9': 6.0,
                    '-': 5.0, ' ': 4.0, ',': 2.0}



cat_names_plural = {'A': "Устройства",
                    'BF': "Телефоны",
                    'BH': "Датчики Холла",
                    'BM': "Микрофоны",
                    'C': "Конденсаторы",
                    'D': "Микросхемы",
                    'FU': "Предохранители",
                    'FA': "Предохранители",
                    'FP': "Термопредахранители",
                    'F': "Разрядники",
                    'GB': "Батареи",
                    'G': "Генераторы",
                    'H': "Устройства индикации",
                    'K': "Реле",
                    'L': "Индуктивности",
                    'R': "Резисторы",
                    'RK': "Терморезисторы",
                    'RP': "Потенциометры",
                    'RU': "Варисторы",
                    'S': "Переключатели",
                    'T': "Трансформаторы",
                    'VD': "Диоды",
                    'VS': "Тиристоры",
                    'VT': "Транзисторы",
                    'WA': "Антенны",
                    'X': "Разъемы",
                    'Z': "Фильтры",
                    'ZQ': "Фильтры кварцевые",
                    'U': "Оптопары"}  # Список наименованый групп компонентов во мн. ч.


cat_names_singular = {'A': "Устройство",
                      'BF': "Телефон",
                      'BH': "Датчик Холла",
                      'BM': "Микрофон",
                      'C': "Конденсатор",
                      'D': "Микросхема",
                      'FU': "Предохранитель",
                      'FA': "Предохранитель",
                      'FP': "Термопредахранитель",
                      'F': "Разрядник",
                      'GB': "Батарея",
                      'G': "Генератор",
                      'H': "Устройство индикации",
                      'K': "Реле",
                      'L': "Индуктивность",
                      'R': "Резистор",
                      'RK': "Терморезистор",
                      'RP': "Потенциометр",
                      'RU': "Варистор",
                      'S': "Переключатель",
                      'T': "Трансформатор",
                      'VD': "Диод",
                      'VS': "Тиристор",
                      'VT': "Транзистор",
                      'WA': "Антенна",
                      'X': "Разъем",
                      'Z': "Фильтр",
                      'ZQ': "Фильтр кварцевый",
                      'U': "Оптопара"}  # Список наименований групп компонентов в ед ч.

class Element:
    def __init__(self, ser: pd.Series = None):
        if ser is not None:
            self.ser = ser
            self.designator = ser['Designator']
            self.char = self.designator[0]
            self.rem = ser['Rem']
            self.body = ser['Корпус']
            self.tke = ser['TKE']
            self.value = ser['Value']
            self.tolerance = ser['Tolerance']
            self.pv = ser['Power/Voltage']
            self.manuf = ser['Manufacturer'].split(',')[0]
            self.man_part_num = ser['ManufacturerPartNumber']
            self.quantity = ser['Quantity']
            self.notice = ser['Примечание']
            self.module = ""

            self.name = self._make_name()

            if '0' not in self.body and '1' not in self.body and self.body != '':
                self.body = "«" + self.body + "»"

            if '"' in self.tke:
                output = list(self.tke)
                output[-1] = "»"
                output[-3] = "«"
                self.tke = ''.join(output)

    def __str__(self):
        return f"{self.name} [{self.quantity}] [{self.module}]"

    def __repr__(self):
        return f"{self.name} [{self.quantity}] [{self.module}]"

    def _make_name(self):
        # Для R и C Номинал и Погрешность переносятся вместе, поэтому сохраняем их в одну строку
        val_tol = self.value + " " + self.tolerance
        name = ''

        # Формирование наименований
        if self.char == 'C':
            if self.manuf.find("ТУ", 0) != -1:
                if self.tke != '':
                    name = f"{self.rem} {self.pv} - {val_tol} - {self.tke} {self.manuf}"
                else:
                    name = f"{self.rem} {self.pv} - {val_tol} {self.manuf}"
            else:
                if self.tke != '':
                    name = f"{self.rem} {self.body} {self.tke}  - {self.pv} - {val_tol}, {self.manuf}"
                else:
                    name = f"{self.rem} {self.body} - {self.pv} - {val_tol}, {self.manuf}"
        if self.char == 'R':
            if self.manuf.find("ТУ", 0) != -1:
                if self.tke != '':
                    name = f"{self.rem} - {self.pv} - {val_tol} - {self.tke} {self.manuf}"
                else:
                    name = f"{self.rem} - {self.pv} - {val_tol} {self.manuf}"
            else:
                name = f"{self.rem} {self.body} - {val_tol}, {self.manuf}"
        # Для прочих компонентов
        if self.char != 'C' and self.char != 'R':
            if self.manuf.find("ТУ", 0) != -1:
                name = f"{self.man_part_num} {self.manuf}"
            else:
                name = f"{self.man_part_num}, {self.manuf}"
            if self.man_part_num == '':
                if self.manuf.find("ТУ", 0) != -1:
                    name = f"{self.value} {self.manuf}"
                else:
                    name = f"{self.value}, {self.manuf}"
        return name

    def split_designators(self, shift_threshold: int):
        """
        Переносит позиции элементов, если те не влезают в строку шаблона

        :param shift_threshold: Порог количества символов для переноса
        :return: Лист позиций. Каждый новый элемент на новой строке
        """
        #shift_threshold = 53
        splitted_desig = self.designator.split(", ")
        for i in range(len(splitted_desig)):
            next_i = i + 1
            # Проверка существования следующей позиции
            if next_i >= len(splitted_desig):
                break
            # Пока длина позиций не превышает порог, добавляем следующую позицию
            # -2 для вставки запятой в конце строки
            while font.getlength(splitted_desig[i] + ", " + splitted_desig[next_i]) <= shift_threshold - 2:
                splitted_desig[i] += ", " + splitted_desig[next_i]
                # Удаляем добавленную позицию из списка всех позиций
                splitted_desig.remove(splitted_desig[next_i])
                # Проверка существования следующей позиции

                if next_i >= len(splitted_desig):
                    break

            if next_i < len(splitted_desig):
                splitted_desig[i] += ","

        self.designator = splitted_desig

    def split_name(self, shift_threshold, cat_name='', one_man=''):
        val_tol = self.value + " " + self.tolerance
        splitted_name = []

        if self.manuf.find("ТУ", 0) != -1 and self.man_part_num != "":
            new_split = self.man_part_num.split()

            i = 0
            while i + 1 < len(new_split):
                new_split[i] = new_split[i].replace('“', "«")
                new_split[i] = new_split[i].replace('”', "»")
                while new_split[i + 1] == '-' or new_split[i + 1] == 'В' or \
                        ("Ф" in new_split[i + 1] or "Гн" in new_split[i + 1] or "Ом" in new_split[i + 1]) or \
                        "%" in new_split[i + 1]:
                    new_split[i] = f"{new_split[i]} {new_split[i + 1]}"
                    del (new_split[i + 1])
                    if i + 1 == len(new_split):
                        break
                else:
                    i += 1
            splitted_name = new_split + [f"{self.manuf}"]
        elif self.char == 'C':
            if self.manuf.find("ТУ", 0) != -1 or (self.name is list and self.name.find("ТУ", 0) != -1):
                if self.tke != '':
                    splitted_name = [f"{self.rem}", f"{self.pv} -", f"{val_tol} -", f"{self.tke}",
                                     f"{self.manuf}"]
                else:
                    splitted_name = [f"{self.rem}", f"{self.pv} -", f"{val_tol}", f"{self.manuf}"]
            else:
                if self.tke != '':
                    splitted_name = [f"{self.rem}", f"{self.body}", f"{self.tke} -", f"{self.pv} -", f"{val_tol},",
                                     f"{self.manuf}"]
                else:
                    splitted_name = [f"{self.rem}", f"{self.body} -", f"{self.pv} -", f"{val_tol},", f"{self.manuf}"]
        elif self.char == 'R':
            if self.manuf.find("ТУ", 0) != -1 or (self.name is list and self.name.find("ТУ", 0) != -1):
                if self.tke != '':
                    splitted_name = [f"{self.rem} -", f"{self.pv} -", f"{val_tol} -", f"{self.tke}", f"{self.manuf}"]
                else:
                    splitted_name = [f"{self.rem} -", f"{self.pv} -", f"{val_tol} -", f"{self.manuf}"]
            else:
                splitted_name = [f"{self.rem}", f"{self.body} -", f"{val_tol},", f"{self.manuf}"]
        # Для прочих компонентов
        elif self.char != 'C' and self.char != 'R':
            if self.manuf.find("ТУ", 0) != -1:
                splitted_name = [f"{self.man_part_num}", f"{self.manuf}"]
            else:
                splitted_name = [f"{self.man_part_num},", f"{self.manuf}"]
            if self.man_part_num == '':
                if self.manuf.find("ТУ", 0) != -1:
                    splitted_name = [f"{self.value}", f"{self.manuf}"]
                else:
                    splitted_name = [f"{self.value},", f"{self.manuf}"]

        if cat_name != '':
            splitted_name.insert(0, cat_name)

        if one_man != '':
            splitted_name.insert(-1, cat_name)

        i = 0 # Индекс для добавления части наименования в уже существующую строку
        name = ['']
        for name_part in splitted_name:
            # Если длина итоговой строки длиннее порога, то создаем новый перенос
            if len(f"{name[-1]} {name_part}") > shift_threshold:
                if name[-1][-1] == '-':
                    name_part = f"- {name_part}"
                name.append(name_part)
                i += 1
                continue
            name[i] += f" {name_part}"

        for index, name_part in enumerate(name):
            name[index] = name[index].lstrip().rstrip()
            if self.manuf == '':
                name[index] = name[index][:-1]

        self.name = name

    def split_man(self, shift_threshold):
        if self.manuf != '':
            i = 0
            new_man = ['']
            for man_part in self.manuf.split():
                # Если длина итоговой строки длиннее порога, то создаем новый перенос
                if len(f"{new_man[-1]} {man_part}") > shift_threshold:
                    if new_man[-1][-1] == '-':
                        man_part = f"- {man_part}"
                    new_man.append(man_part)
                    i += 1
                    continue
                new_man[i] += f" {man_part}"
            self.manuf = new_man

    def split_notice(self, shift_threshold):
        """
        Разбивает текст примечания на строки, не превышающие заданную ширину в пикселях
        :param shift_threshold: Максимальная ширина строки в пикселях
        """
        if not self.notice or not self.notice.strip():
            self.notice = []
            return

        words = self.notice.split()
        result = []
        current_line = ""

        for word in words:
            # Формируем новую строку
            if current_line == "":
                test_line = word
            else:
                test_line = current_line + " " + word

            # Проверяем ширину в пикселях
            if font.getlength(test_line) <= shift_threshold:
                current_line = test_line
            else:
                if current_line:
                    result.append(current_line)
                current_line = word  # начинаем новую строку с текущего слова

        # Добавляем последнюю строку
        if current_line:
            result.append(current_line)

        self.notice = result



def create_name(df: pd.DataFrame, index: int, shift_treshold: int):
    """
    Формирует наименование компонента и переносит его, если тот не будет влезат в рамки шаблона

    :param df: Словарь с сортированными по обозначениям компонентами
    :param index: Индекс редактирумеого элемента
    :param shift_treshold: Порог длины строки, привышая который будет производиться перенос наимнования
    :return: Готовое наименование элемента, [name2, name3] - переносы
    """
    desig = df['Designator'][index]
    rem = df['Rem'][index]
    body = df['Корпус'][index]
    tke = df['TKE'][index]
    val = df['Value'][index]
    pv = df['Power/Voltage'][index]
    tol = df['Tolerance'][index]
    man = df['Manufacturer'][index]
    manpnb = df['ManufacturerPartNumber'][index]
    # Для R и C Номинал и Погрешность переносятся вместе, поэтому сохраняю их в одну строку
    val_tol = val + " " + tol

    name = ''
    name2 = ''
    name3 = ''

    if '0' not in body and '1' not in body and body != '':
        body = "«" + body + "»"

    if '"' in tke:
        output = list(tke)
        output[-1] = "»"
        output[-3] = "«"
        tke = ''.join(output)

    # Формирование наименований
    if desig.find("C", 0) != -1:
        if man.find("ТУ", 0) != -1:
            if tke != '':
                # Костыль для некоторых корпусов
                if '0' not in body and 'М' not in body:
                    body = ''
                if len(f"{rem} {body} - {pv} - {val_tol}") > shift_treshold:
                    name = f"{rem} {body} - {pv} -"
                    name2 = f"- {val_tol} - {tke} {man}"
                    if len(name2) > shift_treshold:
                        name2 = f"- {val_tol} - {tke}"
                        name3 = f"{man}"
                elif len(f"{rem} {body} - {pv} - {val_tol} - {tke}") > shift_treshold:
                    name = f"{rem} {body} - {pv} - {val_tol} -"
                    name2 = f"- {tke} {man}"
                elif len(f"{rem} {body} {pv} - {val_tol} - {tke} {man}") > shift_treshold:
                    name = f"{rem} {body} {pv} - {val_tol} - {tke}"
                    name2 = f"{man}"
                else:
                    name = f"{rem} {body} - {pv} - {val_tol} - {tke} {man}"
            else:
                if len(f"{rem} - {pv} - {val_tol}") > shift_treshold:
                    name = f"{rem} - {pv} -"
                    name2 = f"- {val_tol} {man}"
                elif len(f"{rem} - {pv} - {val_tol} {man}") > shift_treshold:
                    name = f"{rem} - {pv} - {val_tol}"
                    name2 = f"{man}"
                else:
                    name = f"{rem} {body} - {pv} - {val_tol} {man}"
        else:
            if tke != '':
                if len(f"{rem} {body} {tke} - {pv}") > shift_treshold:
                    name = f"{rem} {body} {tke} -"
                    name2 = f"- {pv} - {val_tol}, {man}"
                elif len(f"{rem} {body} {tke} - {pv} - {val_tol}") > shift_treshold:
                    name = f"{rem} {body} {tke}  - {pv} -"
                    name2 = f"- {val_tol}, {man}"
                elif len(f"{rem} {body} {tke}  - {pv} - {val_tol}, {man}") > shift_treshold:
                    name = f"{rem} {body} {tke}  - {pv} - {val_tol},"
                    name2 = f"{man}"
                else:
                    name = f"{rem} {body} {tke}  - {pv} - {val_tol}, {man}"
            else:
                if len(f"{rem} {body} - {pv} - {val_tol}") > shift_treshold:
                    name = f"{rem} {body} - {pv} -"
                    name2 = f"- {val_tol}, {man}"
                elif len(f"{rem} {body} - {pv} - {val_tol}, {man}") > shift_treshold:
                    name = f"{rem} {body} - {pv} - {val_tol},"
                    name2 = f"{man}"
                else:
                    name = f"{rem} {body} - {pv} - {val_tol}, {man}"
    elif desig.find("R", 0) != -1:
        if man.find("ТУ", 0) != -1:
            if tke != '':
                if len(f"{rem} - {pv} - {val_tol}") > shift_treshold:
                    name = f"{rem} - {pv} -"
                    name2 = f"{val_tol} - {tke} {man}"
                    if len(name2) > shift_treshold:
                        name2 = f"- {val_tol} - {tke}"
                        name3 = f"{man}"
                elif len(f"{rem} - {pv} - {val_tol} - {tke}") > shift_treshold:
                    name = f"{rem} - {pv} - {val_tol} -"
                    name2 = f"- {tke} {man}"
                elif len(f"{rem} - {pv} - {val_tol} - {tke} {man}") > shift_treshold:
                    name = f"{rem} - {pv} - {val_tol} - {tke}"
                    name2 = f"{man}"
                else:
                    name = f"{rem} - {pv} - {val_tol} - {tke} {man}"
            else:
                if len(f"{rem} - {pv} - {val_tol}") > shift_treshold:
                    name = f"{rem} - {pv} -"
                    name2 = f"- {val_tol} - {tke} {man}"
                elif len(f"{rem} - {pv} - {val_tol} {tke}") > shift_treshold:
                    name = f"{rem} - {pv} - {val_tol}"
                    name2 = f"- {tke} {man}"
                elif len(f"{rem} - {pv} - {val_tol} {tke} {man}") > shift_treshold:
                    name = f"{rem} - {pv} - {val_tol} - {tke}"
                    name2 = f"{man}"
                else:
                    name = f"{rem} - {pv} - {val_tol} - {tke} {man}"
        else:
            if len(f"{rem} {body}") > shift_treshold:
                name = f"{rem} {body}"
                name2 = f"- {val_tol}, {man}"
            elif len(f"{rem} {body} - {val_tol}") > shift_treshold:
                name = f"{rem} {body} -"
                name2 = f"- {val_tol}, {man}"
            elif len(f"{rem} {body} - {val_tol}, {man}") > shift_treshold:
                name = f"{rem} {body} - {val_tol},"
                name2 = f"{man}"
            else:
                name = f"{rem} {body} - {val_tol}, {man}"
    # Для прочих компонентов
    if desig.find("C", 0) == -1 and desig.find("R", 0) == -1:
        if man.find("ТУ", 0) != -1:
            if len(f"{manpnb} {man}") > shift_treshold:
                name = f"{manpnb}"
                name2 = f"{man}"
            else:
                name = f"{manpnb} {man}"
        else:
            if len(f"{manpnb}, {man}") > shift_treshold:
                name = f"{manpnb},"
                name2 = f"{man}"
            else:
                name = f"{manpnb}, {man}"
        if manpnb == '':
            if man.find("ТУ", 0) != -1:
                name = f"{val} {man}"
            else:
                name = f"{val}, {man}"
    return [name, name2, name3]


def create_names_vp(df: pd.DataFrame, index: int, shift_threshold: int):
    """
    Формирует наименование компонента и переносит его, если тот не будет влезат в рамки шаблона

    :param df: Словарь с сортированными по обозначениям компонентами
    :param index: Индекс редактирумеого элемента
    :param shift_threshold: Порог длины строки, превышая который будет производиться перенос наимнования
    :return: Готовое наименование элемента, [name2, name3] - переносы
    """
    desig = df['Designator'][index]
    rem = df['Rem'][index]
    body = df['Корпус'][index]
    tke = df['TKE'][index]
    val = df['Value'][index]
    pv = df['Power/Voltage'][index]
    tol = df['Tolerance'][index]
    man = df['Manufacturer'][index]
    manpnb = df['ManufacturerPartNumber'][index]
    # Для R и C Номинал и Погрешность переносятся вместе, поэтому сохраняю их в одну строку
    val_tol = val + " " + tol

    name = ''
    name2 = ''

    if '0' not in body \
            and '1' not in body \
            and body != '':
        body = "\"" + body + "\""

    # Формирование наименований
    if desig.find("C", 0) != -1:
        if man.find("ТУ", 0) != -1:
            if tke != '':
                # Костыль для некоторых корпусов
                if '0' not in body and 'М' not in body:
                    body = ''
                if len(f"{rem} {body} - {pv}") > shift_threshold:
                    name = f"{rem} {body} -"
                    name2 = f"- {pv} - {val_tol} - {tke}"
                elif len(f"{rem} {body} - {pv} - {val_tol}") > shift_threshold:
                    name = f"{rem} {body} - {pv} -"
                    name2 = f"- {val_tol} - {tke}"
                elif len(f"{rem} {body} - {pv} - {val_tol} - {tke}") > shift_threshold:
                    name = f"{rem} {body} - {pv} - {val_tol} -"
                    name2 = f"- {tke}"
                else:
                    name = f"{rem} {body} - {pv} - {val_tol} - {tke}"
            else:
                if len(f"{rem} {body} - {pv}") > shift_threshold:
                    name = f"{rem} {body} -"
                    name2 = f"- {pv} - {val_tol}"
                elif len(f"{rem} {body} - {pv} - {val_tol}") > shift_threshold:
                    name = f"{rem} {body} - {pv} -"
                    name2 = f"- {val_tol}"
                else:
                    name = f"{rem} {body} - {pv} - {val_tol}"
        else:
            if tke != '':
                if len(f"{rem} {body} {tke} - {pv}") > shift_threshold:
                    name = f"{rem} {body} {tke} -"
                    name2 = f"- {pv} - {val_tol}"
                elif len(f"{rem} {body} {tke} - {pv} - {val_tol}") > shift_threshold:
                    name = f"{rem} {body} {tke}  - {pv} -"
                    name2 = f"- {val_tol}"
                else:
                    name = f"{rem} {body} {tke}  - {pv} - {val_tol}"
            else:
                if len(f"{rem} {body} - {pv} - {val_tol}") > shift_threshold:
                    name = f"{rem} {body} - {pv} -"
                    name2 = f"- {val_tol}"
                else:
                    name = f"{rem} {body} - {pv} - {val_tol}"
    elif desig.find("R", 0) != -1:
        if man.find("ТУ", 0) != -1:
            if tke != '':
                if len(f"{rem} - {pv} - {val_tol}") > shift_threshold:
                    name = f"{rem} - {pv} -"
                    name2 = f"- {val_tol} - {tke}"
                elif len(f"{rem} - {pv} - {val_tol} - {tke}") > shift_threshold:
                    name = f"{rem} - {pv} - {val_tol} -"
                    name2 = f"- {tke}"
                else:
                    name = f"{rem} - {pv} - {val_tol} - {tke}"
            else:
                if len(f"{rem} - {pv} - {val_tol}") > shift_threshold:
                    name = f"{rem} - {pv} -"
                    name2 = f"- {val_tol} - {tke}"
                elif len(f"{rem} - {pv} - {val_tol} {tke}") > shift_threshold:
                    name = f"{rem} - {pv} - {val_tol}"
                    name2 = f"- {tke}"
                else:
                    name = f"{rem} - {pv} - {val_tol} - {tke}"
        else:
            if len(f"{rem} {body}") > shift_threshold:
                name = f"{rem} {body}"
                name2 = f"- {val_tol}"
            elif len(f"{rem} {body} - {val_tol}") > shift_threshold:
                name = f"{rem} {body} -"
                name2 = f"- {val_tol}"
            else:
                name = f"{rem} {body} - {val_tol}"
    # Для прочих компонентов
    if desig.find("C", 0) == -1 and desig.find("R", 0) == -1:
        if man.find("ТУ", 0) != -1:
            name = f"{manpnb}"
        else:
            name = f"{manpnb}"
        if manpnb == '':
            name = f"{val}"
    return [name, name2]
