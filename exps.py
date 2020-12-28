from datetime import datetime, timedelta, time


def fr():

    now = datetime.now()
    print(now.time())
    print(time(17, 0))
    print(now.time() < time(17, 0))
    if now.time() <= time(17, 0):
        print(now.time() <= time(17, 0))

day = '2020-12-28'

iso_day = datetime.strptime(day, '%Y-%m-%d')

format_day = datetime.strftime(iso_day, "%d.%m.%Y")
weekday = datetime.isoweekday(datetime.now())

print(format_day, weekday)