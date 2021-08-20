from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook


def resolve_name_file(dir_excel, type_ivk):  # функция получения адерса файла из ИВК ВУ для парсинга

    # проверка типа входных данных ИВК ВУ П или Ц
    if type_ivk == "P":
        name_ivk = " Кубань пирамида.xlsx"
    if type_ivk == "C":
        name_ivk = " Кубань Ц.xlsx"

    current_datetime = datetime.now()
    month = current_datetime.month
    day = current_datetime.day

    # првоверка на длину месяца если меньше 2 знаков то впереди добавить 0
    if month < 10:
        month = "0" + str(month)
    else:
        month = str(month)

    # првоверка на длину дня если меньше 2 знаков то впереди добавить 0
    if day < 10:
        day = "0" + str(day)
    else:
        day = str(day)

    # собираем название файла
    file_name = str(current_datetime.year) + month + day + name_ivk

    # определеить домашню директорию
    home_dir = Path.home()
    # собираем адрес до файла
    path = Path(home_dir, dir_excel, file_name)
    path_name = str(path)

    return path_name


#
def open_excel(name_file):  # функция получения сырого массива из файла

    wb = load_workbook(filename=name_file, read_only=True)
    ws = wb['Лист1']
    # массив для данных из excel
    mass_excel = []
    i = 0
    for row in ws.rows:
        # добавить строку в массив
        mass_excel.append([])
        for cell in row:
            # добавить содержимое ячейки в массив
            mass_excel[i].append(cell.value)
        i += 1
    # закрытие документа
    wb.close()

    # вернуть массив
    return mass_excel
