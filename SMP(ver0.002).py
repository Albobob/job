from openpyxl import load_workbook
from docx import Document

# *****************МЕНЮ*************************************
file_name = 'sprvk.xlsx'  # EXEL Файл где хронится Форма 2
sheet = 'Unit'  # ЛИСТ где СМП\
nz = [1, 5, 6]
first_year = ''
last_year = ''
# *****************СЛОВАРЬ*************************************

nz_dict = {1: 'Брюшной тиф', 2: 'Паратифы А,В,С и неуточненный', 3: 'Бактерионосители брюшного тифа, паратифов',
           4: 'Холера', 5: 'Вибриононосители холеры', 6: 'Другие сальмонеллезные инфекции',
           7: 'из них вызванные: сальмонеллами группы B', 8: 'сальмонеллами группы C', 9: 'сальмонеллами группы Д',
           10: 'Бактериальная дизентерия (шигеллез)', 11: 'из нее бактериологически подтвержденная',
           12: 'из нее вызванная: шигеллами Зонне', 13: 'шигеллами Флекснера', 14: 'Бактерионосители дизентерии',
           15: 'ОКИ установленной этиологии и ПТИ'}

wb_val = load_workbook(filename=f'{file_name}', data_only=True, )  # Open file
sheet_val = wb_val[f'{sheet}']  # Open sheet

abc = ['c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l']  # Столбци таблицы

rus = 3  # Строка с РФ (в Exel)
dis = 4  # Строка с Округом (в Exel)
reg = 5  # Строка с Регионом (в Exel)

russian = []  # Список показателей заболеваемости от нозологии № 1 (С3:L3) до number_nz
district = []  # Список показателей заболеваемости от нозологии № 1 (С4:L4) до number_nz
region = []  # Список показателей заболеваемости от нозологии № 1 (С5:L5) до number_nz
for number_nz in nz:
    cell_rus = (rus * int(number_nz)) + 1
    cell_dis = (dis * int(number_nz)) + 1
    cell_reg = (reg * int(number_nz)) + 1

    while rus < cell_rus or dis < cell_dis or reg < cell_reg:
        for i in abc:
            ru = sheet_val[f'{i}{rus}'].value
            russian.append(ru)
            di = sheet_val[f'{i}{dis}'].value
            district.append(di)
            re = sheet_val[f'{i}{reg}'].value
            region.append(re)
        rus += 5
        dis += 5
        reg += 5
    if int(number_nz) == 1:
        c = 0
        l = 10
    else:
        c = 10 * int(number_nz - 1)  # Индекс значения заболеваемости первого года (first_year)
        l = 10 * int(number_nz)  # Индекс значения заболеваемости последнего года (last_year)\

    string_russian = russian[c:l]
    string_district = district[c:l]
    string_region = region[c:l]

    # print(c)
    # print(l)
    print(f'НОЗОЛОГИЯ # {number_nz}')
    print(string_russian)
    # print(len(string_russian))

    print(string_district)
    # print(len(string_district))

    print(string_region)
    # print(len(string_region))
