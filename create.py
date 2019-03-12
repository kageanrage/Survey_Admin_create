import os, time, pprint, logging, sqlite3, subprocess, pyautogui, shutil, send2trash, datetime, calendar
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from config import Config   # this imports the config file where the private data sits
import pandas as pd
from dateutil.relativedelta import *


# logging.basicConfig(level=logging.DEBUG, format=' %(asctime)s - %(levelname)s - %(message)s')  # turns on logging
# logging.disable(logging.CRITICAL)     # switches off logging when desired

cfg = Config()  # create an instance of the Config class, essentially brings private config data into play
os.chdir(cfg.cwd)  # change the current working directory to the one stipulated in config file


p_number_to_search = 'P-46251'


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

# define all variables here
qf_msg = cfg.qf_msg
so_msg = cfg.so_msg
comp_msg = cfg.comp_msg

# varied for each project
p_number_col = 'SE Project Number'

# columns of interest for SQL query
survey_name_col = 'Survey Name'
status = 'Active'
topic_col = 'Topic'
expected_loi_col = 'Expected LOI'
client_name_col = 'Client name'
sales_contact_col = 'Sales Contact'
edge_credits_col = 'Edge Credits'

# Fixed variables - same for all projects
external_survey_url = 'tbc'
prize_draw_entries = '1'
qf_outcome_reward_id = 'Disqualified Survey - Regular Prize Draw'
so_outcome_reward_id = 'Disqualified Survey - Regular Prize Draw'
comp_outcome_reward_id = 'Completed Survey - Regular Prize Draw'
comp_secondary_reward_type = 'Credits'
tc_filepath = cfg.tc_filepath

live_excel_name_path = cfg.live_excel_file_path + "\\" + cfg.live_excel_filename
live_excel_file_name_path_ext = live_excel_name_path + ".xlsm"


# TODO: import xlsm to sqlite

conn = sqlite3.connect(cfg.live_excel_filename + ".db")
c = conn.cursor()

df = pd.read_excel(live_excel_file_name_path_ext, sheet_name='PPT')  # create dataframe from xlsm content
df.to_sql('PPT', conn)  # populate database with dataframe content
conn.commit()

table_name = "PPT"


# TODO: using example P-number, look up all variables of interest in the SQL database

# 1) Contents of all columns for row that match a certain value in 1 column
c.execute('SELECT "{coi1}","{coi2}","{coi3}","{coi4}","{coi5}","{coi6}" FROM {tn} WHERE "{cn}"="{scn}"'.format(tn=table_name, cn=p_number_col, coi1=survey_name_col, coi2=topic_col, coi3=expected_loi_col, coi4=client_name_col, coi5=sales_contact_col, coi6=edge_credits_col, scn=p_number_to_search))  # note I need to put speech marks around "{cn}" because the column name contains a space
all_rows = c.fetchall()

survey_name = all_rows[0][0]  # assign outputs to variable names
topic = all_rows[0][1]
expected_loi = all_rows[0][2]
client_name = all_rows[0][3]
sales_contact = all_rows[0][4]
edge_credits = int(all_rows[0][5])


# TODO: open web browser, navigate to Create Survey page


def login():
    driver.get(cfg.assign_URL)  # use selenium webdriver to open web browser and desired URL from config file
    driver.execute_script("document.getElementById('UserName').value = '" + str(cfg.uname) + "';")  # insert username
    driver.execute_script("document.getElementById('Password').value = '" + str(cfg.pwd) + "';")  # insert password
    pass_elem = driver.find_element_by_id('Password')  # find the 'Password' text box using its element ID
    pass_elem.submit()  # submit password
    time.sleep(2)   # wait 2 seconds for the login process to take place (tested and this is necessary)


def grab_redirects():
    quota_full_url = driver.find_element_by_id('OutcomeFullUrl').get_attribute('value')  # find the right element and grab URL from box
    screened_url = driver.find_element_by_id('OutcomeScreenedUrl').get_attribute('value')  # find the right element and grab URL from box
    complete_url = driver.find_element_by_id('OutcomeCompleteUrl').get_attribute('value')  # find the right element and grab URL from box
    quota_full_url = quota_full_url[0:83]  # trim off the last 3 characters
    screened_url = screened_url[0:83]  # trim off the last 3 characters
    complete_url = complete_url[0:83]  # trim off the last 3 characters
    return quota_full_url, screened_url, complete_url


def establish_project_dir():
    qf, so, comp = grab_redirects()
    print(f'Quota Full: {qf}')
    print(f'Screened: {so}')
    print(f'Complete: {comp}')
    new_dir_path = cfg.projects_dir_path + "\\" + client_name + "\\" + p_number_to_search + " - " + survey_name
    logging.debug("Creating new directory:", new_dir_path)
    os.mkdir(new_dir_path)  # creates new directory
    subprocess.Popen(f'explorer "{new_dir_path}"')  # opens new dir in windows explorer

def enter_data():
    driver.execute_script("document.getElementById('Name').value = '" + str(survey_name) + "';")
    driver.execute_script("document.getElementById('Status').value = '" + str(status) + "';")
    driver.execute_script("document.getElementById('Title').value = '" + str(topic) + "';")
    driver.execute_script("document.getElementById('ProjectIONumber').value = '" + str(p_number_to_search) + "';")
    driver.execute_script("document.getElementById('ExpectedLength').value = '" + str(expected_loi) + "';")
    driver.execute_script("document.getElementById('ClientCompanyName').value = '" + str(client_name) + "';")
    driver.execute_script("document.getElementById('ExternalSurveyUrl').value = '" + str(external_survey_url) + "';")
    driver.execute_script("document.getElementById('StartDate').value = '" + str(start_date) + "';")
    driver.execute_script("document.getElementById('EndDate').value = '" + str(end_date) + "';")
    driver.execute_script("document.getElementById('OutcomeFull').value = '" + str(qf_msg) + "';")
    driver.execute_script("document.getElementById('OutcomeScreened').value = '" + str(so_msg) + "';")
    driver.execute_script("document.getElementById('OutcomeComplete').value = '" + str(comp_msg) + "';")
    driver.execute_script("document.getElementById('OutcomeFullRewardValue').value = '" + str(prize_draw_entries) + "';")
    driver.execute_script("document.getElementById('OutcomeScreenedRewardValue').value = '" + str(prize_draw_entries) + "';")
    driver.execute_script("document.getElementById('OutcomeCompleteRewardValue').value = '" + str(prize_draw_entries) + "';")
    driver.find_element_by_id('FullOutcomeRewardId').send_keys(qf_outcome_reward_id)  # this didn't work via JS execution method
    # driver.execute_script("document.getElementById('FullOutcomeRewardId').value = '" + str(qf_outcome_reward_id) + "';")
    driver.find_element_by_id('ScreenedOutcomeRewardId').send_keys(so_outcome_reward_id)  # this didn't work via JS execution method
    # driver.execute_script("document.getElementById('ScreenedOutcomeRewardId').value = '" + str(so_outcome_reward_id) + "';")
    driver.find_element_by_id('CompleteOutcomeRewardId').send_keys(comp_outcome_reward_id)  # this didn't work via JS execution method
    # driver.execute_script("document.getElementById('CompleteOutcomeRewardId').value = '" + str(comp_outcome_reward_id) + "';")
    driver.execute_script("document.getElementById('OutcomeCompleteSecondaryRewardValue').value = '" + str(edge_credits) + "';")
    driver.find_element_by_id('OutcomeCompleteSecondaryRewardType').send_keys(comp_secondary_reward_type)
    driver.find_element_by_id('TermsAndConditionsPdf').click()
    time.sleep(2)
    pyautogui.typewrite(tc_filepath)  # since popup window is outside web browser, need a diff package to control
    pyautogui.press('enter')
    driver.find_element_by_css_selector('#add-edit-survey > fieldset > dl > div.form_navigation > button').click()  # Submits / creates new project


chrome_path = cfg.chrome_path  # location of chromedriver.exe on local drive
driver = webdriver.Chrome(chrome_path)  # specify webdriver (selenium)

login()
establish_project_dir()
# grab_redirects()
# enter_data()



# TODO: capture redirect info


# TODO: create project dir - a directory in the appropriate client folder - NB will have to occur after DB is queried

# TODO: create excel file with redirects


conn.close()

