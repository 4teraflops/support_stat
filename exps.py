import requests
import json
from lxml import html
from bs4 import BeautifulSoup
from datetime import datetime, timedelta


s = requests.Session()
cookies = {}


def get_answered_unanswered():
    answered_url = 'http://voip.bisys.ru/queue-stats/answered.php'
    headers = {
        "Content-Type": "application/x-www-form-urlencoded",
        "Authorization": "Basic c3VwcG9ydGhlbHA6cXdlcnR5",
        "User-Agent": "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:80.0) Gecko/20100101 Firefox/80.0"
    }
    yesterday = datetime.strftime((datetime.now() - timedelta(days=2)), "%Y-%m-%d")
    payload = {
        "start": f"{yesterday} 00:00:00",
        "end": f"{yesterday} 23:59:59",
        "List_Queue[]": ["'100'", "'200'"],
        "List_Agent[]": ["'125-Cvecih'", "'154-Yakovlev'", "'140-Voronina'", "'137-Zmeev'", "'148-Nazarov'", "'155-Kulakov'", "'147-Besonogov'", "'147-Besogonov'"]
    }
    ans_request = s.post(answered_url, data=payload, headers=headers)
    ans_soup = BeautifulSoup(ans_request.text, 'lxml')  # Передаем в суп полученную страницу
    # находим нужный тег, отсекая лишние символы
    answered = ans_soup.find('table').find_all('td')[11].text.replace(' выз.', '')

    unanswered_url = 'http://voip.bisys.ru/queue-stats/unanswered.php'
    unans_request = s.post(unanswered_url, data=payload, headers=headers)  # Идем во вкладку пропущенные
    unans_soup = BeautifulSoup(unans_request.text, 'lxml')  # Передаем в суп полученную страницу
    unanswered = unans_soup.find('table').find_all('td')[11].text.replace(' выз.', '')
    print(f'answered: {answered}\nunanswered: {unanswered}')


if __name__ == '__main__':
    try:
        get_answered_unanswered()
    except KeyboardInterrupt:
        print('Program stopped')


#html = open('index.html', 'r')
#soup = BeautifulSoup(html, 'lxml')  # Передаем в суп полученную страницу
## находим нужный тег, отсекая лишние символы
#data = soup.find('table').find_all('td')[11].text.replace(' выз.', '')
#print(data)
