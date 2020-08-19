import gspread
from datetime import datetime, timedelta


#  Подключаемся к google sheet
gc = gspread.service_account(filename='creds.json')
sh = gc.open('test Статистика support 2020')
months = {
    8: 'Август',
    9: 'Сентябрь',
    10: 'Октябрь',
    11: 'Ноябрь',
    12: 'Декабрь'
}

#  Смотрим текущий месяц
month = months.get(datetime.now().month)
#  Выбираем вкладку текущего месяца
worksheet = sh.worksheet(month)
#  Находим столбец, где вчерашний день
yesterday = datetime.strftime((datetime.now() - timedelta(days=1)), "%d.%m.%Y")
cell = worksheet.find(yesterday)
#  cell.row - строка cell.col - столбец. Заполняем данные

stat = [3, 27, 25, 23, 2, 31, 24, 23, 50]

#  Создано support
worksheet.update_cell((cell.row + 1), cell.col, stat[0])
#  Создано Customer Help Desk
worksheet.update_cell((cell.row + 2), cell.col, stat[1])
#  Создано Сustomer Help Desk Vezet
worksheet.update_cell((cell.row + 3), cell.col, stat[2])
#  Создано Support Kassir
worksheet.update_cell((cell.row + 4), cell.col, stat[3])
#  Закрыто SUPPORT
worksheet.update_cell((cell.row + 6), cell.col, stat[4])
#  Закрыто Сustomer Help Desk
worksheet.update_cell((cell.row + 7), cell.col, stat[5])
#  Закрыто Сustomer Help Desk Vezet
worksheet.update_cell((cell.row + 8), cell.col, stat[6])
#  Закрыто support Kassir
worksheet.update_cell((cell.row + 9), cell.col, stat[7])
#  Осталось в задачнике на конец суток
worksheet.update_cell((cell.row + 12), cell.col, stat[8])