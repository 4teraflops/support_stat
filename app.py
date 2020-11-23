from config import youtrack_api_token, tg_webhook_url, admin_id
from youtrack.connection import Connection as YouTrack
from youtrack import YouTrackException
import gspread
from bs4 import BeautifulSoup
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
import requests
import json
from loguru import logger


stat = {'yt': [], 'c2d': {}, 'calls': {}}
"""Не работает, если будешь собирать стату за последний день месяца"""
logger.add(f'log/{__name__}.log', format='{time} {level} {message}', level='DEBUG', rotation='10 MB', compression='zip')


def yt_stat():

    yt = YouTrack('https://youtrack.bisys.ru', token=youtrack_api_token)

    try:
        support_created_yesterday = yt.getNumberOfIssues(filter='проект: {!W Support yt@bisys.ru} создана: Вчера')
        stat['yt'].append(support_created_yesterday)
        customer_helpdesk_created_yesterday = yt.getNumberOfIssues(filter='проект: {Customer Help Desk} создана: Вчера')
        stat['yt'].append(customer_helpdesk_created_yesterday)
        customer_helpdesk_vezet_created_yesterday = yt.getNumberOfIssues(filter='проект: {Customer Help Desk Vezet, vezet@ckassa.ru} создана: Вчера')
        stat['yt'].append(customer_helpdesk_vezet_created_yesterday)
        support_kassir_created_yesterday = yt.getNumberOfIssues(filter='проект: {!W Support Kassir yt-kassir@bisys.ru} создана: Вчера')
        stat['yt'].append(support_kassir_created_yesterday)
        support_closed_yesterday = yt.getNumberOfIssues(filter='проект: SUPPORT дата завершения: Вчера')
        stat['yt'].append(support_closed_yesterday)
        customer_helpdesk_closed_yesterday = yt.getNumberOfIssues(filter='проект: {Customer Help Desk} дата завершения: Вчера')
        stat['yt'].append(customer_helpdesk_closed_yesterday)
        customer_helpdesk_vezet_closed_yesterday = yt.getNumberOfIssues(filter='проект: {Customer Help Desk Vezet, vezet@ckassa.ru} дата завершения: Вчера')
        stat['yt'].append(customer_helpdesk_vezet_closed_yesterday)
        support_kassir_closed_yesterday = yt.getNumberOfIssues(filter='проект: {!W Support Kassir yt-kassir@bisys.ru} дата завершения: Вчера')
        stat['yt'].append(support_kassir_closed_yesterday)
        issules_waiting_summary = yt.getNumberOfIssues(filter='проект: SUPPORT , {Customer Help Desk} , {Customer Help Desk Vezet, vezet@ckassa.ru} , KASSA #Незавершенная')
        stat['yt'].append(issules_waiting_summary)
    except TimeoutException as e:
        logger.error(f'Timeout Exception: {e}')
    except YouTrackException as e:
        logger.debug(f'YouTrackException: {e}')
    except Exception as e:
        logger.error(f'Other Exception: {e}')
    if stat['yt']:
        logger.debug(f'Данные YT собраны:{stat}')
    else:
        logger.debug("Can't record YT data")


def c2d_stat():

    '''Задумано так, что статистика будет собираться после 19 часов, иначе данные поменяются и соберутся некорректно'''

    login_url = 'https://web.chat2desk.com/'
    cd_dir_path = 'src' + os.sep + 'chromedriver'

    chrome_options = Options()  # задаем параметры запуска драйвера
    chrome_options.add_argument('--headless')  # скрывать окна хрома
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument("--window-size=1980,1024")  # Включил на весь экран, чтоб срендерилась кновка "вход"
    driver = webdriver.Chrome(cd_dir_path, chrome_options=chrome_options)
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
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/section/section/section/div/div/section/div/div/div/div/div[1]/div[2]/div[1]/button/span[1]'))).click()

    #  Генерим отчет (теперь отчет генерится сам)
    #wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/section/section/section/div/div/section/div/div/div/div[3]/div[1]/div[1]/button[3]'))).click()
    #time.sleep(1)
    #wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/section/section/section/div/div/section/div/div/div/div[3]/div[3]/button'))).click()

    #  Собираем данные
    appreal_sum = 0
    #  Вытащил все значения из столбца
    appreal_count = wait.until(EC.visibility_of_any_elements_located(
        (By.XPATH, '/html/body/section/section/section/div/div/section/div/div/div/div[4]/table/tbody/tr/td[3]')))
    #  Ссумировал их между собой
    for count in appreal_count:
        appreal_sum += int(count.text)
    stat['c2d']['appreal_sum'] = appreal_sum
    #logger.info(f'appreal_sum: {appreal_sum}')

    #  Теперь идем в custom_reports_section
    wait.until(
        EC.element_to_be_clickable((By.XPATH, '/html/body/section/div[1]/div/div/div[1]/ul/li[2]/ul/li[3]/a'))).click()

    #  Выбираем отчет
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/section/section/section/div/div/section/div/div/div/div/div[1]/div[2]/div[3]/button/span[1]/span'))).click()

    #  Генерим отчет (теперь отчет генерится сам)
    #wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/section/section/section/div/div/section/div/div/div/div[3]/div[1]/div[1]/button[3]'))).click()
    #wait.until(EC.element_to_be_clickable((By.XPATH,
    #                                       '/html/body/section/section/section/div/div/section/div/div/div/div[3]/div[2]/div/div/div/div/div'))).click()
    #time.sleep(1)
    #wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div[2]/ul/li[1]'))).click()
    #wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/section/section/section/div/div/section/div/div/div/div[3]/div[3]/button'))).click()

    #  Собираем данные
    raiting_average = wait.until(EC.element_to_be_clickable(
        (By.XPATH, '/html/body/section/section/section/div/div/section/div/div/div/div[4]/table/tbody/tr[1]/td[2]'))).text
    #logger.info(f'raiting_average: {raiting_average.text}')
    stat['c2d']['raiting_average'] = raiting_average
    #logger.info(f'stat: {stat}')
    #  Теперь идем в custom_reports_section
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/section/div[1]/div/div/div[1]/ul/li[2]/ul/li[3]/a'))).click()

    #  Выбираем отчет reply_time
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/section/section/section/div/div/section/div/div/div/div/div[1]/div[2]/div[2]/button/span[1]/span'))).click()

    #  Генерим отчет (теперь отчет генерится сам)
    #wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/section/section/section/div/div/section/div/div/div/div[3]/div[1]/div[1]/button[3]'))).click()
    #time.sleep(1)
    #wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/section/section/section/div/div/section/div/div/div/div[3]/div[3]/button'))).click()

    #  Забираем данные
    total_reply_time_data = wait.until(EC.element_to_be_clickable((By.XPATH,
                                                              '//*[@id="main-content"]/section/div/div/section/div/div/div/div[4]/table/tbody/tr[3]/td[3]'))).text
    #logger.info(f'total_reply_time: {total_reply_time}')
    # Собираем все значения из таблицы

    reply_time_all_values = wait.until(EC.visibility_of_any_elements_located(
        (By.XPATH, '/html/body/section/section/section/div/div/section/div/div/div/div[4]/table/tbody/tr/td[3]')))
    reply_tymes_list = []
    # Кладем их в список
    for reply_times in reply_time_all_values:
        reply_tymes_list.append(reply_times.text)
    # Берем последнее значение из списка
    reply_time_average = str(reply_tymes_list[-1])
    #logger.info(f'reply_tymes_list: {reply_tymes_list}')
    #logger.info(f'reply_time_average: {reply_time_average}')

    stat['c2d']['reply_time_average'] = reply_time_average
    driver.quit()

    logger.debug(f'Данные из C2D добавлены: {stat}')


def get_answered_unanswered():
    s = requests.Session()
    answered_url = 'http://voip.bisys.ru/queue-stats/answered.php'
    headers = {
        "Content-Type": "application/x-www-form-urlencoded",
        "Authorization": "Basic c3VwcG9ydGhlbHA6cXdlcnR5",
        "User-Agent": "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:80.0) Gecko/20100101 Firefox/80.0"
    }
    yesterday = datetime.strftime((datetime.now() - timedelta(days=1)), "%Y-%m-%d")
    weekday = datetime.isoweekday((datetime.now() - timedelta(days=1)))
    logger.info(f'weekday: {weekday}')
    if weekday == 5:  # По пятнцицам смотреть звонки до 17 часов
        endtime = "16:59:59"
    else:  # В остальные дни смотреть до 18
        endtime = "17:59:59"

    payload = {
        "start": f"{yesterday} 09:00:00",
        "end": f"{yesterday} {endtime}",
        "List_Queue[]": ["'100'", "'200'"],
        "List_Agent[]": ["'125-Cvecih'", "'154-Yakovlev'", "'140-Voronina'", "'137-Zmeev'", "'148-Nazarov'", "'155-Kulakov'", "'147-Besonogov'", "'147-Besogonov'", "'156-Solcina'"]
    }
    logger.info(f'endtime: {payload["end"]}')
    ans_request = s.post(answered_url, data=payload, headers=headers)
    ans_soup = BeautifulSoup(ans_request.text, 'lxml')  # Передаем в суп полученную страницу
    # находим нужный тег, отсекая лишние символы
    answered = ans_soup.find('table').find_all('td')[11].text.replace(' выз.', '')

    unanswered_url = 'http://voip.bisys.ru/queue-stats/unanswered.php'
    unans_request = s.post(unanswered_url, data=payload, headers=headers)  # Идем во вкладку пропущенные
    unans_soup = BeautifulSoup(unans_request.text, 'lxml')  # Передаем в суп полученную страницу
    unanswered = unans_soup.find('table').find_all('td')[11].text.replace(' выз.', '')
    stat['calls']['answered'] = answered
    stat['calls']['unanswered'] = unanswered
    logger.debug(f'Данные по звонкам добавлены: {stat}')


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
    weekday = datetime.isoweekday((datetime.now() - timedelta(days=2)))
    cell = worksheet.find(yesterday)
    if weekday in [6, 7]:  # Если вчера была суббота, либо воскресение, то данные по звонкам не заполнять
        #  cell.row - строка cell.col - столбец. Заполняем данные
        #  Создано support
        try:
            worksheet.update_cell((cell.row + 1), cell.col, stat['yt'][0])
            #  Создано Customer Help Desk
            worksheet.update_cell((cell.row + 2), cell.col, stat['yt'][1])
            #  Создано Сustomer Help Desk Vezet
            worksheet.update_cell((cell.row + 3), cell.col, stat['yt'][2])
            #  Создано Support Kassir
            worksheet.update_cell((cell.row + 4), cell.col, stat['yt'][3])
            #  Закрыто SUPPORT
            worksheet.update_cell((cell.row + 6), cell.col, stat['yt'][4])
            #  Закрыто Сustomer Help Desk
            worksheet.update_cell((cell.row + 7), cell.col, stat['yt'][5])
            #  Закрыто Сustomer Help Desk Vezet
            worksheet.update_cell((cell.row + 8), cell.col, stat['yt'][6])
            #  Закрыто support Kassir
            worksheet.update_cell((cell.row + 9), cell.col, stat['yt'][7])
            #  Осталось в задачнике на конец суток
            worksheet.update_cell((cell.row + 12), cell.col, stat['yt'][8])
        except IndexError:
            logger.error("Can't record YT data")
        try:
            #  Обработано диалогов в открытых линиях
            worksheet.update_cell((cell.row + 13), cell.col, stat['c2d']['appreal_sum'])
            #  Средняя оценка в чатах
            worksheet.update_cell((cell.row + 14), cell.col, stat['c2d']['raiting_average'])
            #  Среднее время реакции
            worksheet.update_cell((cell.row + 19), cell.col, stat['c2d']['reply_time_average'])
        except KeyError:
            logger.error("Can't record c2d data")
        # Количество принятых звонков
        #worksheet.update_cell((cell.row + 15), cell.col, stat['calls']['answered'])
        # Количество пропущенных звонков
        #worksheet.update_cell((cell.row + 16), cell.col, stat['calls']['unanswered'])
        logger.info("It's day off. Ignore calls info.")
    else:
        try:
            worksheet.update_cell((cell.row + 1), cell.col, stat['yt'][0])
            worksheet.update_cell((cell.row + 2), cell.col, stat['yt'][1])
            worksheet.update_cell((cell.row + 3), cell.col, stat['yt'][2])
            worksheet.update_cell((cell.row + 4), cell.col, stat['yt'][3])
            worksheet.update_cell((cell.row + 6), cell.col, stat['yt'][4])
            worksheet.update_cell((cell.row + 7), cell.col, stat['yt'][5])
            worksheet.update_cell((cell.row + 8), cell.col, stat['yt'][6])
            worksheet.update_cell((cell.row + 9), cell.col, stat['yt'][7])
            worksheet.update_cell((cell.row + 12), cell.col, stat['yt'][8])
        except IndexError:
            logger.error("Can't record YT data")
        try:
            worksheet.update_cell((cell.row + 13), cell.col, stat['c2d']['appreal_sum'])
            worksheet.update_cell((cell.row + 14), cell.col, stat['c2d']['raiting_average'])
            worksheet.update_cell((cell.row + 19), cell.col, stat['c2d']['reply_time_average'])
        except KeyError:
            logger.error("Can't record c2d data")
        try:
            worksheet.update_cell((cell.row + 15), cell.col, stat['calls']['answered'])
            worksheet.update_cell((cell.row + 16), cell.col, stat['calls']['unanswered'])
        except KeyError:
            logger.error("Can't record calls data")


def tg_alarm(alarmtext):
    headers = {"Content-type": "application/json"}
    payload = {"text": f"{alarmtext}", "chat_id": f"{admin_id}"}
    requests.post(url=tg_webhook_url, data=json.dumps(payload), headers=headers)


@logger.catch()
def main():
    yt_stat()
    c2d_stat()
    get_answered_unanswered()
    record_data()


if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print('Program stopped')
    except Exception as e:
        alarmtext = f'Support_stat: (app.py): {str(e)}'
        tg_alarm(alarmtext)