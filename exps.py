from datetime import datetime, timedelta, time


def fr():

    now = datetime.now()
    print(now.time())
    print(time(17, 0))
    print(now.time() < time(17, 0))
    if now.time() <= time(17, 0):
        print(now.time() <= time(17, 0))

fr()