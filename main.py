import xlsxwriter

'''Получение данных от пользователя.'''
print('Для генерации таблицы внесите необходимые данные.')
year = input('Введите год: ')  # текущий год
month = input('Введите название  месяца: ')  # текущий месяц
score_day = int(input('Количество дней в месяце: '))  # общее количество дней в мясяце
start_day = int(input('С какого числа начинаем отсчет?: '))  # с какого числа начинаются рабочие дни

workbook = xlsxwriter.Workbook('table.xlsx')
worksheet = workbook.add_worksheet()

# Формат ячеек
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

# Назначение переменных заголовков
name = 'Чек лист'
stage = 'Состояние'
temper_out = 't на улице'
temper_in = 't внутри'
time_p = 'Время замера: '
comment = 'Примечания: '
tavro = 'Подпись: '

# Отрисовываем и заполняем шапку таблицы
worksheet.merge_range('A1:B3', name.upper(), merge_format)  # слитно 2х2
worksheet.merge_range('C1:C3', stage, merge_format)  # слитно 2х2
worksheet.merge_range('D1:S1', year, merge_format)
worksheet.merge_range('D2:S2', month.upper(), merge_format)

# Отрисовываем и заполняем столбец слева (заголовки)
worksheet.merge_range('A4:B4', temper_out, merge_format)  # слитно 2х2 температура уличная
worksheet.merge_range('A5:B5', temper_in, merge_format)  # слитно 2х2 температура в помещении
worksheet.merge_range('A6:A9', 'Блок №1', merge_format)
worksheet.merge_range('A10:A13', 'Блок №2', merge_format)
worksheet.merge_range('A14:A17', 'Блок №3', merge_format)
worksheet.merge_range('A18:A21', 'Блок №4', merge_format)

# В цикле заполняем 4 одинаковых блока таблицы
option = ['t вход', 't выход', 't уставка', 'режим']
i = 0
v = 5
h = 1
for i in range(4):
    for opt in option:
        worksheet.write(v, h, opt, merge_format)
        v += 1

# Отрисовываем и заполняем футер таблицы
worksheet.merge_range('A22:B22', time_p, merge_format)
worksheet.merge_range('A23:B23', tavro, merge_format)
worksheet.merge_range('A24:B27', comment, merge_format)

# Генерация пустых полей таблицы
row = 3
col = 2
blanc = ''

h = 2  # номер отсчёта строки
v = 2  # номер отсчёта столбца
c = start_day - 1  # начало месяца
i = 0  # счётчик

# Генерируем пустые ячейки таблицы
if score_day == 30:
    t = 15
    end_cell = 18
elif score_day == 31:
    t = 16
    end_cell = 19

for i in range(3, 27):
    for j in range(2, end_cell):
        worksheet.write(i, j, blanc, merge_format_light)

while i != t:
    v = 2  # номер строки
    h += 1  # номер столбца
    c += 1  # начало отсчёта дней
    i += 1
    worksheet.write(v, h, c, merge_format)
    if c >= score_day:
        break
    elif c >= start_day + 15:
        break

workbook.close()
