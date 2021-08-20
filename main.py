from func import resolve_name_file, open_excel

dir_excel = "data_excel"
type_ivk = "P"

resolve_name = resolve_name_file(dir_excel, type_ivk)

Data_exel = open_excel(resolve_name)




print("1")
