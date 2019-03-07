import os, time, pprint, logging, sqlite3
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from config import Config   # this imports the config file where the private data sits
import pandas as pd

logging.basicConfig(level=logging.DEBUG, format=' %(asctime)s - %(levelname)s - %(message)s')  # turns on logging
# logging.disable(logging.CRITICAL)     # switches off logging when desired

cfg = Config()  # create an instance of the Config class, essentially brings private config data into play
os.chdir(cfg.cwd)  # change the current working directory to the one stipulated in config file

logging.debug('Imported modules')
logging.debug('Start of program')
logging.debug(f'Current cwd = {os.getcwd()}')

p_number_to_search = 'P-46240'
date_example = "07/03/2019 00:00:00"
# TODO: code to create Start Date (today's date) and End Date (last day of next month)

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

