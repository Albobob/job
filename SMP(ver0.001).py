from openpyxl import load_workbook

# Menu
file_name = 'sprvk.xlsx'
sheet = 'Unit'  # ЛИСТ где СМП
number_nz = '29'
# Menu


wb_val = load_workbook(filename=f'{file_name}', data_only=True, )  # Open file

sheet_val = wb_val[f'{sheet}']  # Open sheet
abc = ['c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l']  # Выбираю строку со значениями
new_list = []  # Список со всеми занчениями заболеваемости региона
alb = 5
cell_alb = (alb * int(number_nz) + 1)
while alb < cell_alb:
    for i in abc:
        # print(alb)
        # print(i)
        c5_val = sheet_val[f'{i}{alb}'].value
        new_list.append(c5_val)

    alb += 5
print(f' Все НОЗОЛОГИИ: {new_list}')
# print(len(new_list))    # кол-во ячеек
nz_list = []
cl_5 = 0
lc_5 = 9

while cl_5 <= (int(number_nz) * 10) - 10 and lc_5 <= (int(number_nz) * 10) - 1:
    string = new_list[cl_5:lc_5]
    nz_list.append(string)

    cl_5 += 10
    lc_5 += 10
print(f'Нозологии без последнего года: {nz_list}')
# print(len(nz_list))   # кол-во листов
string_nz = nz_list[int(number_nz) - 1]

print(f'Нозология №{number_nz}: {string_nz}')
string_nz_copy = string_nz.copy()


for i in string_nz_copy:
    if i == 0:
        null = string_nz_copy.index(0)
        string_nz_copy.pop(null)  # Удаляем ноль
maximum = max(string_nz_copy)
minimum = min(string_nz_copy)
indx_max = string_nz_copy.index(maximum)
string_nz_copy.pop(indx_max)  # Удаляем Максимальное число
indx_min = string_nz_copy.index(minimum)
string_nz_copy.pop(indx_min)  # Удаляем минимальное число

smp = sum(string_nz_copy) / len(string_nz_copy)

print(f'{smp} СМП нозологии №{number_nz}')

