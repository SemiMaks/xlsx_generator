import xlsxwriter
from survey import year, month, score_day, start_day, work_day
# from filter import

workbook = xlsxwriter.Workbook('table.xlsx')
worksheet = workbook.add_worksheet()
merge_format = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': '#e3e3e3'})

name = 'чек лист'
stage = 'Состояние'
temper_out = 't на улице'
temper_in = 't внутри'

time_p = 'Время замера: '
comment = 'Примечания: '
tavro = 'Подпись: '

# Шапка таблицы
worksheet.merge_range('A1:B3', name.upper(), merge_format)  # слитно 2х2
worksheet.merge_range('C1:C3', stage, merge_format)  # слитно 2х2
worksheet.merge_range('D1:X1', year, merge_format)
worksheet.merge_range('D2:X2', month.upper(), merge_format)

h = 2  # номер отсчёта строки
v = 2  # номер отсчёта столбца
c = start_day - 1  # начало месяца
i = 0  # счётчик

while i != score_day - 1:
    h += 1  # номер строки
    v = 2  # номер столбца
    c += 1  # начало отсчёта дней
    i += 1
    worksheet.write(v, h, c, merge_format)

date_list = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10',
             '11', '12', '13', '14', '15', '16', '17', '18', '19', '20',
             '21', '22', '23', '24', '25', '26', '27', '28', '29', '30', '31']
if work_day == '1':
    new_d = date_list[start_day - 1:]
    print(new_d)

elif work_day == '2':
    new_d = date_list[start_day - 1:]

elif work_day == '3':
    new_d = date_list[start_day - 1:]

elif work_day == '4':
    new_d = date_list[start_day - 1:]

elif work_day == '5':
    new_d = date_list[start_day - 1:]

else:
    print('Недопустимое число!')
print('Выбран день: ', work_day)
# print(new_d)


# Столбец слева (заголовки)
worksheet.merge_range('A4:B4', temper_out, merge_format)  # слитно 2х2 температура уличная
worksheet.merge_range('A5:B5', temper_in, merge_format)  # слитно 2х2 температура в помещении

option = ['t вход', 't выход', 't уставка', 'режим']

# worksheet.merge_range('A6:A9', 'Блок №1', merge_format)
# worksheet.write('B6', op1, merge_format)
# worksheet.write('B7', op2, merge_format)
# worksheet.write('B8', op3, merge_format)
# worksheet.write('B9', op4, merge_format)

worksheet.merge_range('A6:A9', 'Блок №1', merge_format)
worksheet.merge_range('A10:A13', 'Блок №2', merge_format)
worksheet.merge_range('A14:A17', 'Блок №3', merge_format)
worksheet.merge_range('A18:A21', 'Блок №4', merge_format)

i = 0
v = 5
h = 1
for i in range(4):
    for opt in option:
        worksheet.write(v, h, opt, merge_format)
        v += 1

worksheet.merge_range('A22:B23', time_p, merge_format)
worksheet.merge_range('A24:B25', comment, merge_format)
worksheet.merge_range('A26:B27', tavro, merge_format)

# worksheet.insert_image('A29', 'fintech.png')

workbook.close()
