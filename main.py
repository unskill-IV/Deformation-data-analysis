# from matplotlib import pyplot as plt
from openpyxl import load_workbook, Workbook


def load_data(data_name):
    """
    Загружает xlsx файлы с данными трёх датчиков (T1, T2, T3) для дальнейшей обработки.

    :param data_name: Название серии измерений (н-р, 3-В4_deformation_2018_6MHz_sensor_3).
    :type data_name: str

    :return: 3 объекта типа Workbook, соответствующие датчикам Τ1, Τ2, Τ3.
    :rtype: tuple[Workbook, Workbook, Workbook]
    """

    wb_1 = load_workbook(f'./measurements/{data_name}/{data_name} (Т1).xlsx', data_only=True, read_only=True)
    wb_2 = load_workbook(f'./measurements/{data_name}/{data_name} (Т2).xlsx', data_only=True, read_only=True)
    wb_3 = load_workbook(f'./measurements/{data_name}/{data_name} (Т3).xlsx', data_only=True, read_only=True)

    return wb_1, wb_2, wb_3


def get_rows_count(book, sheet_name):
    """
    Возвращает количество строк в листе.

    :param book: Workbook, в котором находится нужный лист
    :type book: Workbook

    :param sheet_name: Название листа
    :type sheet_name: str

    :rtype: int
    """

    sheet = book[sheet_name]

    return len(list(sheet.rows))


def str_to_float_if_needed(val):
    """
    Иногда Excel выдаёт данные в формате str вместо float. Функция решает эту проблему.\n
    Также функция округляет значения до 4-х знаков после запятой.

    Пример:
    ' 18,9452 ' -> 18.9452

    :param val: Входные данные
    :type val: str | float

    :rtype: float
    """

    if isinstance(val, str):
        val = float(val[1:-1].replace(',', '.'))

    return round(val, 4)


def get_values(book, sheet_name):
    """
    Выгружает пару данных с листа в 2 массива, игнорируя ошибочные ячейки.\n
    При считывании используется функция str_to_float_if_needed.\n
    Считанные значения округляются до 4-х знаков после запятой.

    :param book: Workbook, в котором находятся данные.
    :type book: Workbook

    :param sheet_name: Название листа, с которого считываются данные.
    :type sheet_name: str

    :rtype: tuple[list, list]
    """

    rc = get_rows_count(book, sheet_name)
    sheet = book[sheet_name]
    val_1 = []
    val_2 = []

    for r in range(2, rc+1):
        if sheet.cell(row=r, column=2).value != 'ERROR!':
            val_1.append(str_to_float_if_needed(sheet.cell(row=r, column=2).value))
            val_2.append(round(sheet.cell(row=r, column=3).value, 4))

    return val_1, val_2


def mean_value(vals):
    """
    Находит среднее значение серии. Результат округляется до 4-х знаков после запятой.

    :param vals: Массив с данными формата float.
    :type vals: list[float]
    :rtype: float
    """

    mv = round(sum(vals) / len(vals), 4)
    return mv


def error_finder(vals_1, vals_2):
    """
    Определяет константу сдвига значений. Результат округляется до 2-х знаков после запятой.\n
    Выделяет пары с большой ошибкой, для оставшихся значений находит среднее - база, от которой находится сдвиг, находит средний сдвиг выделенных измерений.

    :param vals_1: Массив значений 1
    :type vals_1: list[float]

    :param vals_2: Массив значений 2
    :type vals_2: list[float]

    :rtype: float

    .. note::

        Если абсолютный сдвиг значения оказывается больше 0.35, то это значение игнорируется.
    """

    errindexs = []

    for l in range(len(vals_1)):
        dv = vals_1[l] - vals_2[l] * 2

        if abs(dv) > 0.1:
            errindexs.append(l)

    if not errindexs:
        return 0.0
    else:
        vls_1 = []
        vls_2 = []

        for l in range(len(vals_1)):
            if l not in errindexs:
                vls_1.append(vals_1[l])
                vls_2.append(vals_2[l])

        mean_1 = mean_value(vls_1)
        mean_2 = mean_value(vls_2)
        errs = []

        for v in errindexs:
            err_1 = round(vals_1[v] - mean_1, 2)
            err_2 = round(vals_2[v] - mean_2, 2)

            if err_1 <= 0.35:
                errs.append(abs(err_1))

            if err_2 <= 0.35:
                errs.append(abs(err_2))

        err = round(mean_value(errs), 2)

        return err


def error_correction(val_1, val_2, err):
    """
    Корректирует значения на заданную константу по следующим правилам:\n
    Если ошибка между .1 и .4 -> значение 1 - константа;\n
    Если ошибка больше .4 -> значение 2 + константа;\n
    Если ошибка между -.4 и -.1 -> значение 1, значение 2 - константа;\n
    Если ошибка меньше -.4 -> значение 2 - константа.

    :param val_1: Значение 1
    :type val_1: float

    :param val_2: Значение 2
    :type val_2: float

    :param err: Константа сдвига
    :type err: float

    :return: Пару значений в порядке Значение 1, Значение 2
    :rtype: tuple[float, float]

    .. note::

        Ошибка всё ещё может остаться большой, в таких случаях можно повторно использовать корректировку или игнорировать такие значения.
    """

    dval = round(val_1 - val_2 * 2, 4)

    if abs(dval) >= 0.1:
        if 0.1 <= dval <= 0.4:
            val_1 -= err
        elif -0.4 <= dval <= -0.1:
            val_1 -= err
            val_2 -= err
        elif dval > 0.4:
            val_2 += err
        elif dval < -0.4:
            val_2 -= err

    return round(val_1, 4), round(val_2, 4)


file = input('Enter measurement name: ')
db = load_data(file)

wb = Workbook()

for i in range(3):
    n = 0
    wb['Sheet'].cell(row=((i * 3) + 2), column=1, value=f'T{(i + 1)}')
    wb['Sheet'].cell(row=((i * 3) + 2), column=2, value='dt')
    wb['Sheet'].cell(row=((i * 3) + 3), column=2, value='mt')
    wb['Sheet'].cell(row=((i * 3) + 4), column=2, value='UV|TV')

    shs = db[i].sheetnames

    for sh in shs:
        tn, t2 = get_values(db[i], sh)
        ta = list(map(lambda x: x * 2, t2))
        xc = error_finder(tn, t2)
        dt = []
        tv = len(tn)

        ntn = [0.0] * tv
        nt2 = [0.0] * tv
        nta = [0.0] * tv

        for j in range(tv):
            ntn[j], nt2[j] = error_correction(tn[j], t2[j], xc)
            nta[j] = nt2[j] * 2
            dtj = round(ntn[j] - nta[j], 4)
            dt.append(ntn[j] - nta[j])

        while max(list(map(lambda x: abs(x), dt))) > 0.1:
            for j in range(len(dt)):
                if abs(dt[j]) > 0.1:
                    ntn.pop(j)
                    nt2.pop(j)
                    nta.pop(j)
                    dt.pop(j)
                    break

        ut = len(ntn)
        mt = list(map(lambda x: round(x / 2, 4), ntn))
        mdt = mean_value(dt)
        mmt = mean_value(mt)

        n += 1
        wb['Sheet'].cell(row=1, column=(n + 2), value=sh)
        wb['Sheet'].cell(row=((i * 3) + 2), column=(n + 2), value=mdt)
        wb['Sheet'].cell(row=((i * 3) + 3), column=(n + 2), value=mmt)
        wb['Sheet'].cell(row=((i * 3) + 4), column=(n + 2), value=f'{ut}|{tv}')

wb.save(f'./results/{file}/{file}.xlsx')
