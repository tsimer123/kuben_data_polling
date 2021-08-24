from datetime import datetime
from func import resolve_name_file, open_excel, first_del_empty_str, del_old_pu, del_alien_pu, off_status_pu, \
    del_name_feeder, rename_empty_feeder, save_data_excel_in_wb, count_pu_in_feeder

dir_excel = "data_excel"
type_ivk = "P"
start_building = datetime(2020, 6, 1)
list_type_pu = ['СЕ', 'Ри']
str_type_off_status = 'Не учит.'
list_del_name_feeder = ['Фидер ', 'ВЛ-10 кВ ']

resolve_name = resolve_name_file(dir_excel, type_ivk)

data_excel = open_excel(resolve_name)

data_excel = first_del_empty_str(data_excel)

data_excel = del_old_pu(data_excel, start_building)

data_excel = del_alien_pu(data_excel, list_type_pu)

data_excel = off_status_pu(data_excel, str_type_off_status)

data_excel = rename_empty_feeder(data_excel)

data_excel = del_name_feeder(data_excel, list_del_name_feeder)

feeder_count_pu = count_pu_in_feeder(data_excel)

#data_excel = save_data_excel_in_wb(data_excel)

print("1")
