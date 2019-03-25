import datetime, calendar
from dateutil.relativedelta import *
import pprint


now = datetime.datetime.now()  # current date and time as datetime object

last_day_in_month = calendar.monthrange(now.year, now.month)[1]
print(last_day_in_month)
