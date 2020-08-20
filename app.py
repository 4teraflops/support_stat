from config import youtrack_api_token
from youtrack.connection import Connection as YouTrack
import gspread
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import os
from config import email, passwd
import time

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

    print(f'Данные YT собраны:{stat}')


def c2d_stat():
    login_url = 'https://web.chat2desk.com/'
    cd_dir_path = 'src' + os.sep + 'chromedriver'

    chrome_options = Options()  # задаем параметры запуска драйвера, чтоб не крашился (когда работает во много потоков)
    chrome_options.add_argument('--headless')  # скрывать окна хрома
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument("--window-size=1980,1024")  # Включил на весь экран, чтоб срендерилась кновка "вход"
    driver = webdriver.Chrome(cd_dir_path, chrome_options=chrome_options)  # указал где брать гугл хром драйвер
    driver.implicitly_wait(30)  # неявное ожидание драйвера
    wait = WebDriverWait(driver, 3)  # Задал переменную, чтоб настроить явное ожидание элемента (сек)

    #  Логин
    driver.get(login_url)
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/header/div[1]/div[1]/a[1]'))).click()  # Нашел форму ввода логина
    login_field = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div[1]/div/form/input')))
    login_field.send_keys(f'{email}')
    passwd_filed = wait.until(
        EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div[1]/div/form/div[1]/input')))
    passwd_filed.send_keys(f'{passwd}')
    time.sleep(1)
    passwd_filed.send_keys(Keys.ENTER)

    #  Открываем отчет 1 для ТА
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/section/div[1]/div/div/div[1]/ul/li[2]/a'))).click()
    wait.until(
        EC.element_to_be_clickable((By.XPATH, '/html/body/section/div[1]/div/div/div[1]/ul/li[2]/ul/li[3]/a'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/section/section/section/div/div/section/div/div/div/div/div[3]/div[2]/div[3]/div[2]/div/div[1]/button'))).click()

    #  Генерим отчет
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/section/section/section/div/div/section/div/div/div/div[3]/div[1]/div[1]/button[3]'))).click()
    time.sleep(1)
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/section/section/section/div/div/section/div/div/div/div[3]/div[3]/button'))).click()

    #  Собираем данные
    appreal_sum = 0
    #  Вытащил все значения из столбца
    appreal_count = wait.until(EC.visibility_of_any_elements_located(
        (By.XPATH, '/html/body/section/section/section/div/div/section/div/div/div/div[4]/table/tbody/tr/td[3]')))
    #  ссуммировал их между собой
    for count in appreal_count:
        appreal_sum += int(count.text)
    stat.append(appreal_sum)

    #  Теперь идем в custom_reports_section
    wait.until(
        EC.element_to_be_clickable((By.XPATH, '/html/body/section/div[1]/div/div/div[1]/ul/li[2]/ul/li[3]/a'))).click()

    #  Выбираем отчет
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/section/section/section/div/div/section/div/div/div/div/div[3]/div[2]/div[1]/div[2]/div/div[2]/button'))).click()

    #  Генерим отчет
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/section/section/section/div/div/section/div/div/div/div[3]/div[1]/div[1]/button[3]'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH,
                                           '/html/body/section/section/section/div/div/section/div/div/div/div[3]/div[2]/div/div/div/div/div'))).click()
    time.sleep(1)
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div[2]/ul/li[1]'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/section/section/section/div/div/section/div/div/div/div[3]/div[3]/button'))).click()

    #  Собираем данные
    raiting_average = wait.until(EC.element_to_be_clickable(
        (By.XPATH, '/html/body/section/section/section/div/div/section/div/div/div/div[4]/table/tbody/tr[1]/td[2]')))
    stat.append(raiting_average.text)

    #  Теперь идем в custom_reports_section
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/section/div[1]/div/div/div[1]/ul/li[2]/ul/li[3]/a'))).click()

    #  Выбираем отчет reply_time
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/section/section/section/div/div/section/div/div/div/div/div[3]/div[2]/div[1]/div[3]/div/div/button'))).click()

    #  Генерим отчет
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/section/section/section/div/div/section/div/div/div/div[3]/div[1]/div[1]/button[3]'))).click()
    time.sleep(1)
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/section/section/section/div/div/section/div/div/div/div[3]/div[3]/button'))).click()

    #  Забираем данные
    total_reply_time = wait.until(EC.element_to_be_clickable((By.XPATH,
                                                              '/html/body/section/section/section/div/div/section/div/div/div/div[4]/table/tbody/tr[4]/td[3]'))).text

    stat.append(total_reply_time)

    print(f'Данные из C2D добавлены: {stat}')


def record_data():
    #  Подключаемся к google sheet
    gc = gspread.service_account(filename='creds.json')
    sh = gc.open('Статистика support 2020')
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
    #  Обработано диалогов в открытых линиях
    worksheet.update_cell((cell.row + 13), cell.col, stat[9])
    #  Средняя оценка в чатах
    worksheet.update_cell((cell.row + 14), cell.col, stat[10])
    #  Среднее время реакции
    worksheet.update_cell((cell.row + 19), cell.col, stat[11])


if __name__ == '__main__':
    yt_stat()
    c2d_stat()
    record_data()
