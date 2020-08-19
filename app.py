from config import youtrack_api_token
from youtrack.connection import Connection as YouTrack
import gspread
from datetime import datetime, timedelta


stat = []


def yt_stat():

    yt = YouTrack('https://youtrack.bisys.ru', token=youtrack_api_token)

    support_created_yesterday = yt.getNumberOfIssues(filter='проект: {!W Support yt@bisys.ru} создана: Вчера')
    stat.append(support_created_yesterday)
    customer_helpdesk_created_yesterday = yt.getNumberOfIssues(filter='проект: {Customer Help Desk} создана: Вчера')
    stat.append(customer_helpdesk_created_yesterday)
    customer_helpdesk_vezet_created_yesterday = yt.getNumberOfIssues(filter='проект: {Customer Help Desk Vezet, vezet@ckassa.ru} создана: Вчера')
    stat.append(customer_helpdesk_vezet_created_yesterday)
    support_kassir_created_yesterday = yt.getNumberOfIssues(filter='проект: {!W Support Kassir yt-kassir@bisys.ru} создана: Вчера')
    stat.append(support_kassir_created_yesterday)

    support_closed_yesterday = yt.getNumberOfIssues(filter='проект: SUPPORT дата завершения: Вчера')
    stat.append(support_closed_yesterday)
    customer_helpdesk_closed_yesterday = yt.getNumberOfIssues(filter='проект: {Customer Help Desk} дата завершения: Вчера')
    stat.append(customer_helpdesk_closed_yesterday)
    customer_helpdesk_vezet_closed_yesterday = yt.getNumberOfIssues(filter='проект: {Customer Help Desk Vezet, vezet@ckassa.ru} дата завершения: Вчера')
    stat.append(customer_helpdesk_vezet_closed_yesterday)
    support_kassir_closed_yesterday = yt.getNumberOfIssues(filter='проект: {!W Support Kassir yt-kassir@bisys.ru} дата завершения: Вчера')
    stat.append(support_kassir_closed_yesterday)

    issules_waiting_summary = yt.getNumberOfIssues(filter='проект: SUPPORT , {Customer Help Desk} , {Customer Help Desk Vezet, vezet@ckassa.ru} , KASSA #Незавершенная')
    stat.append(issules_waiting_summary)

    print(f'Данные собраны:{stat}')


def record_data():
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


if __name__ == '__main__':
    yt_stat()
    record_data()
