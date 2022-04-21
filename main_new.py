'''Вторая версия генератора таблиц'''

import xlsxwriter

'''Получение данных от пользователя.'''
print('Для генерации таблицы внесите необходимые данные.')
year = input('Введите год: ')  # текущий год
month = input('Введите название  месяца: ')  # текущий месяц
score_day = int(input('Количество дней в месяце: '))  # общее количество дней в мясяце
start_day = int(input('С какого числа начинаем отсчет?: '))  # с какого числа начинаются рабочие дни

workbook = xlsxwriter.Workbook('table_1.xlsx')
worksheet = workbook.add_worksheet()

'''Формат ячеек'''
merge_format = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': '#e0ffff'})

merge_format_light = workbook.add_format({
    'border': 1,
    'fg_color': '#ffffff',
})

'''Назначение переменных заголовков'''
name = 'Чек лист'
stage = 'Состояние'
temper_out = 't на улице'
temper_in = 't внутри'
time_p = 'Время замера: '
comment = 'Примечания: '
tavro = 'Подпись: '

'''Отрисовываем и заполняем шапку таблицы'''
worksheet.merge_range('A1:C3', name.upper(), merge_format)  # слитно 2х2
worksheet.merge_range('D1:E3', stage, merge_format)  # слитно 2х2
worksheet.merge_range('F1:T2', year, merge_format)
worksheet.merge_range('F3:T3', month.upper(), merge_format)

'''Отрисовываем и заполняем столбец слева (заголовки)'''
worksheet.merge_range('A4:C5', temper_out, merge_format)
worksheet.merge_range('A6:C7', temper_in, merge_format)

worksheet.merge_range('A8:A15', 'Блок №1', merge_format)
worksheet.merge_range('A16:A23', 'Блок №2', merge_format)
worksheet.merge_range('A24:A31', 'Блок №3', merge_format)
worksheet.merge_range('A32:A39', 'Блок №4', merge_format)

# '''В цикле заполняем 4 одинаковых блока таблицы'''
option = ['t вход', 't выход', 't уставка', 'режим']
# i = 0
# v = 5
# h = 1
# for i in range(4):
#     for opt in option:
#         worksheet.write(v, h, opt, merge_format)
#         v += 1

# Блок 1
worksheet.merge_range('B8:C9', option[0], merge_format)
worksheet.merge_range('B10:C11', option[1], merge_format)
worksheet.merge_range('B12:C13', option[2], merge_format)
worksheet.merge_range('B14:C15', option[3], merge_format)

# Блок 2
worksheet.merge_range('B16:C17', option[0], merge_format)
worksheet.merge_range('B18:C19', option[1], merge_format)
worksheet.merge_range('B20:C21', option[2], merge_format)
worksheet.merge_range('B22:C23', option[3], merge_format)

# Блок 3
worksheet.merge_range('B24:C25', option[0], merge_format)
worksheet.merge_range('B26:C27', option[1], merge_format)
worksheet.merge_range('B28:C29', option[2], merge_format)
worksheet.merge_range('B30:C31', option[3], merge_format)

# Блок 4
worksheet.merge_range('B32:C33', option[0], merge_format)
worksheet.merge_range('B34:C35', option[1], merge_format)
worksheet.merge_range('B36:C37', option[2], merge_format)
worksheet.merge_range('B38:C39', option[3], merge_format)

'''Отрисовываем и заполняем футер таблицы'''
worksheet.merge_range('A40:C41', time_p, merge_format)
worksheet.merge_range('A42:C43', tavro, merge_format)
worksheet.merge_range('A44:C45', comment, merge_format)

'''Генерация пустых ячеек таблицы'''
row = 3
col = 2
blanc = ''

h = 2  # номер отсчёта строки
v = 2  # номер отсчёта столбца
c = start_day - 1  # начало месяца
i = 0  # счётчик
end_cell = 0
t = 0

# '''Генерируем пустые ячейки таблицы'''
# if score_day == 30:
#     t = 15
#     end_cell = 18
# elif score_day == 31:
#     t = 16
#     end_cell = 19
#
# for i in range(3, 27):
#     for j in range(2, end_cell):
#         worksheet.write(i, j, blanc, merge_format_light)
#
# '''Заполняем числа месяца'''
# while i != t:
#     v = 2  # номер строки
#     h += 1  # номер столбца
#     c += 1  # начало отсчёта дней
#     i += 1
#     worksheet.write(v, h, c, merge_format)
#     if c >= score_day:
#         break
#     elif c >= start_day + 15:
#         break

# list_col = ['D4:E5', 'D6:E7', 'D8:E9', 'D10:E11', 'D12:E13', 'D14:E15',
#             'D16:E17', 'D18:E19', 'D20:E21', 'D22:E23', 'D24:E25', 'D26:E27',
#             'D28:E29', 'D30:E31', 'D32:E33', 'D34:E35', 'D36:E37', 'D38:E39',
#             'D40:E41', 'D42:E43', 'D44:E45']

dict_col = {'1': 'D4:E5', '2': 'D6:E7', '3': 'D8:E9', '4': 'D10:E11', '5': 'D12:E13', '6': 'D14:E15',
            '7': 'D16:E17', '8': 'D18:E19', '9': 'D20:E21', '10': 'D22:E23', '11': 'D24:E25', '12': 'D26:E27',
            '13': 'D28:E29', '14': 'D30:E31', '15': 'D32:E33', '16': 'D34:E35', '17': 'D36:E37', '18': 'D38:E39',
            '19': 'D40:E41', '20': 'D42:E43', '21': 'D44:E45'}

i = 1
for i in dict_col:
    # q = dict_col[i]
    worksheet.merge_range(dict_col[i], blanc, merge_format_light)




# worksheet.write(5, 5, blanc, merge_format_light)

# worksheet.merge_range('D4:E5', blanc, merge_format_light)

workbook.close()
