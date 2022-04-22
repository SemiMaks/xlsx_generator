import xlsxwriter

'''Получение данных от пользователя.'''
print()
line = '-' * 50
print(line)
print('Для генерации таблицы внесите необходимые данные.')
print()
year = input('Введите год: ')  # текущий год
month = input('Введите название  месяца: ')  # текущий месяц

try:
    day_months = int(input('Полное количество дней в месяце: '))
    score_day = int(input('Часть месяца (1 - первая половина, 2 - вторая): '))  # часть месяца
except Exception as err:
    print(err)

start_day = 1
print('Год:', year, ', месяц:', month, ', выбор части:', score_day)
print(line)

name_file = str(year) + '-' + month.lower() + '-' + str(score_day) + '.xlsx'
workbook = xlsxwriter.Workbook(name_file)
worksheet = workbook.add_worksheet()

try:
    '''Формат ячеек'''
    '''для текста заголовков'''
    merge_format = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#e0ffff'})

    '''для отрисовки ячеек'''
    merge_format_light = workbook.add_format({
        'border': 1,
        'fg_color': '#ffffff',
    })

    '''Назначение переменных заголовков'''
    name = 'чек лист'
    stage = 'статус'
    temper_out = 't на улице'
    temper_in = 't внутри'
    time_p = 'время замера: '
    comment = 'примечания: '
    tavro = 'подпись: '
    long_year = ''  # переменная для определения длины заголовка Год
    long_month = ''  # переменная для определения длины заголовка Месяц
    t = 0  # количество столбцов для генерации пустых ячеек
    k = 0  # конечное число чисел месяца

    # Устанавливаем условия для отрисовки таблицы
    if score_day == 1:
        month = month + '-' + '1/2'
        long_year = 'D1:R1'
        long_month = 'D2:R2'
        t = 15
        k = 15
    elif score_day == 2:
        month = month + '-' + '2/2'
        if day_months == 31:
            long_year = 'D1:S1'
            long_month = 'D2:S2'
            t = 16
            k = 31
        elif day_months == 30:
            long_year = 'D1:R1'
            long_month = 'D2:R2'
            t = 15
            k = 30
        elif day_months == 29:
            long_year = 'D1:Q1'
            long_month = 'D2:Q2'
            t = 14
            k = 29
        elif day_months == 28:
            long_year = 'D1:P1'
            long_month = 'D2:P2'
            t = 13
            k = 28
    else:
        print('Недопустимое число!')

    '''Отрисовываем и заполняем шапку таблицы'''
    worksheet.merge_range('A1:B3', name.upper(), merge_format)
    worksheet.merge_range('C1:C3', stage.upper(), merge_format)
    worksheet.merge_range(long_year, year, merge_format)
    worksheet.merge_range(long_month, month.upper(), merge_format)

    '''Отрисовываем и заполняем столбец слева (заголовки)'''
    worksheet.merge_range('A4:B4', temper_out, merge_format)
    worksheet.merge_range('A5:B5', temper_in, merge_format)
    worksheet.merge_range('A6:A9', 'Блок №1', merge_format)
    worksheet.merge_range('A10:A13', 'Блок №2', merge_format)
    worksheet.merge_range('A14:A17', 'Блок №3', merge_format)
    worksheet.merge_range('A18:A21', 'Блок №4', merge_format)

    '''В цикле заполняем 4 одинаковых блока таблицы'''
    option = ['t вход', 't выход', 't уставка', 'режим']
    i = 0
    v = 5
    h = 1
    for i in range(4):
        for opt in option:
            worksheet.write(v, h, opt, merge_format)
            v += 1

    '''Отрисовываем и заполняем низ таблицы'''
    worksheet.merge_range('A22:B22', time_p.upper(), merge_format)
    worksheet.merge_range('A23:B23', tavro.upper(), merge_format)
    worksheet.merge_range('A24:B24', comment.upper(), merge_format)

    '''Генерация пустых ячеек таблицы'''
    row = 3
    col = 2
    blanc = ''

    h = 2  # номер отсчёта строки
    v = 2  # номер отсчёта столбца
    c = start_day - 1  # начало месяца
    i = 0  # счётчик
    end_cell = 0

    '''Заполняем числа месяца'''
    if score_day == 1:
        m = 0
        for m in range(0, k):
            v = 2  # номер столбца
            h += 1  # номер строки
            m += 1
            worksheet.write(v, h, m, merge_format)
    elif score_day == 2:
        m = 0
        for m in range(15, k):
            v = 2  # номер столбца
            h += 1  # номер строки
            m += 1
            worksheet.write(v, h, m, merge_format)

    '''Генерируем пустые ячейки таблицы'''
    print(t)
    t = t + 3  # коррекция числа столбцов
    for i in range(3, 30):
        for j in range(2, t):
            worksheet.write(i, j, blanc, merge_format_light)
except Exception as err:
    print(err)

workbook.close()
