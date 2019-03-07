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

p_number = 'P-46240'


# TODO: import xlsm to sqlite

filename = "KP temp test copy"
con = sqlite3.connect(filename+".db")
df = pd.read_excel(filename+'.xlsm', sheet_name='PPT')
df.to_sql('PPT', con)
con.commit()
con.close()


# TODO: lookup Title and other variables corresponding to project number of interest
