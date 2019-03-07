import os, time, pprint, logging, sqlite3
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from config import Config   # this imports the config file where the private data sits
import pandas as pd
import datetime
import calendar
from dateutil.relativedelta import *

logging.basicConfig(level=logging.DEBUG, format=' %(asctime)s - %(levelname)s - %(message)s')  # turns on logging
# logging.disable(logging.CRITICAL)     # switches off logging when desired

cfg = Config()  # create an instance of the Config class, essentially brings private config data into play
os.chdir(cfg.cwd)  # change the current working directory to the one stipulated in config file

# logging.debug('Imported modules')
# logging.debug('Start of program')
# logging.debug(f'Current cwd = {os.getcwd()}')

p_number_to_search = 'P-46240'  # just used in the testing phase
date_example = "07/03/2019 00:00:00"  # not coded, just a visual reminder of date format requirement


def generate_dates():
    now = datetime.datetime.now()  # current date and time as datetime object
    start_date_string = now.strftime("%d/%m/%Y %H:%M:%S")  # current date and time as string
    # print(f'Start Date is {start_date_string}')
    this_time_next_month = now + relativedelta(months=+1)  # date time object for today's date + time in a month
    next_month = this_time_next_month.month  # isolating just the number of next month
    year_next_month = this_time_next_month.year  # isolating just the year it will be next month
    next_month_string = str(next_month)  # convert number of next month to a string
    if len(next_month_string) == 1:  # if next month is only 1 digit...
        next_month_string = "0" + next_month_string  # ...add a leading zero
    end_date_string = str(calendar.monthrange(year_next_month, next_month)[1]) + "/" + str(next_month_string) \
                      + "/" + str(year_next_month) + " 00:00:00"  # compile full end date string
    return start_date_string, end_date_string


start_date, end_date = generate_dates()  # generates start and end dates through the function

qf_msg = cfg.qf_msg
so_msg = cfg.so_msg
comp_msg = cfg.comp_msg

survey_name_col = 'Survey Name'
topic_col = 'Topic'
p_number_col = 'SE Project Number'
expected_loi_col = 'Expected LOI'
client_name_col = 'Client name'
sales_contact_col = 'Sales Contact'
edge_credits_col = 'Edge Credits'

external_survey_url = 'tbc'
prize_draw_entries = '1'
qf_outcome_reward_id = 'Disqualified Survey - Regular Prize Draw'
so_outcome_reward_id = 'Disqualified Survey - Regular Prize Draw'
comp_outcome_reward_id = 'Completed Survey - Regular Prize Draw'
comp_secondary_reward_type = 'Credits'

# TODO: create project dir - a directory in the appropriate client folder
# TODO: whilst on page, capture 3 x redirects and export them into an excel file in project dir


# TODO: import xlsm to sqlite

filename = "KP temp test copy"
conn = sqlite3.connect(filename + ".db")
c = conn.cursor()
# df = pd.read_excel(filename+'.xlsm', sheet_name='PPT')  # create dataframe from xlsm content
# df.to_sql('PPT', conn)  # populate database with dataframe content
table_name = "PPT"
conn.commit()


# TODO: using example P-number, look up all variables of interest

# 1) Contents of all columns for row that match a certain value in 1 column
c.execute('SELECT "{coi1}","{coi2}" FROM {tn} WHERE "{cn}"="{scn}"'.format(tn=table_name, cn=p_number_col, coi1=topic_col, coi2=sales_contact_col, scn=p_number_to_search))  # note I need to put speech marks around "{cn}" because the column name contains a space
all_rows = c.fetchall()
print(all_rows)





conn.close()

