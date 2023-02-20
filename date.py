from datetime import datetime
from dateutils.relativedelta import relativedelta

# 전달 첫일 ~ 말일
def get_prev_date_range()->list:
  today = datetime.today()
  prev_start = datetime(today.year, today.month, 1) + relativedelta(months=-1)
  prev_end = datetime(today.year, today.month, 1) + relativedelta(seconds=-1)
  return prev_start, prev_end
