from datetime import datetime, timedelta


yesterday = datetime.strftime((datetime.now() - timedelta(days=1)), "%Y-%m-%d")
print(datetime.isoweekday((datetime.now() - timedelta(days=2))))
