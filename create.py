import os, time, pprint, logging, sqlite3
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from config import Config   # this imports the config file where the private data sits
import pandas as pd
import datetime
import calendar
from dateutil.relativedelta import *
import subprocess


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


# TODO: import xlsm to sqlite

filename = "KP temp test copy"
conn = sqlite3.connect(filename + ".db")
c = conn.cursor()
# df = pd.read_excel(filename+'.xlsm', sheet_name='PPT')  # create dataframe from xlsm content
# df.to_sql('PPT', conn)  # populate database with dataframe content
table_name = "PPT"
conn.commit()


# TODO: using example P-number, look up all variables of interest in the SQL database

# 1) Contents of all columns for row that match a certain value in 1 column
c.execute('SELECT "{coi1}","{coi2}","{coi3}","{coi4}","{coi5}","{coi6}" FROM {tn} WHERE "{cn}"="{scn}"'.format(tn=table_name, cn=p_number_col, coi1=survey_name_col, coi2=topic_col, coi3=expected_loi_col, coi4=client_name_col, coi5=sales_contact_col, coi6=edge_credits_col, scn=p_number_to_search))  # note I need to put speech marks around "{cn}" because the column name contains a space
all_rows = c.fetchall()
# print(all_rows)

survey_name = all_rows[0][0]  # assign outputs to variable names
topic = all_rows[0][1]
expected_loi = all_rows[0][2]
client_name = all_rows[0][3]
sales_contact = all_rows[0][4]
edge_credits = all_rows[0][5]


# TODO: open web browser, navigate to Create Survey page


def login(driv):
    driv.get(cfg.assign_URL)  # use selenium webdriver to open web browser and desired URL from config file
    email_elem = driv.find_element_by_id('UserName')  # find the 'Username' text box on web page using its element ID
    email_elem.send_keys(cfg.uname)  # enter username from config file
    pass_elem = driv.find_element_by_id('Password')  # find the 'Password' text box using its element ID
    pass_elem.send_keys(cfg.pwd)  # enter password from config file
    pass_elem.submit()
    time.sleep(2)   # wait 2 seconds for the login process to take place (unsure if this is necessary)


def enter_data(driv, surveyname):
    try:  # structured as try / except statement in case something's gone wrong
        driv.find_element_by_id('Name').send_keys(surveyname)  # find the 'Survey name' text box on web page using its element ID and populate with survey name
        driv.find_element_by_id('Status').send_keys(status)  # find the 'Survey name' text box on web page using its element ID and populate with survey name
        driv.find_element_by_id('Title').send_keys(topic)  # find the 'Survey name' text box on web page using its element ID and populate with survey name
        # TODO: T&C Upload section
        driv.find_element_by_id('ProjectIONumber').send_keys(p_number_to_search)  # NB REPLACE WITH REAL P NUMBER
        driv.find_element_by_id('ExpectedLength').send_keys(expected_loi)  # find the 'Survey name' text box on web page using its element ID and populate with survey name
        driv.find_element_by_id('ClientCompanyName').send_keys(client_name)  # find the 'Survey name' text box on web page using its element ID and populate with survey name
        driv.find_element_by_id('OutcomeCompleteSecondaryRewardValue').send_keys(external_survey_url)
        driv.find_element_by_id('StartDate').send_keys(start_date)
        driv.find_element_by_id('EndDate').send_keys(end_date)
        driv.find_element_by_id('OutcomeFull').send_keys(qf_msg)
        driv.find_element_by_id('OutcomeScreened').send_keys(so_msg)
        driv.find_element_by_id('OutcomeComplete').send_keys(comp_msg)
        driv.find_element_by_id('OutcomeFullRewardValue').send_keys(prize_draw_entries)
        driv.find_element_by_id('OutcomeScreenedRewardValue').send_keys(edge_credits)
        driv.find_element_by_id('OutcomeCompleteRewardValue').send_keys(edge_credits)   


        driv.find_element_by_id('OutcomeCompleteSecondaryRewardValue').send_keys(edge_credits)  # find the 'Survey name' text box on web page using its element ID and populate with survey name





        driv.find_element_by_id('OutcomeCompleteSecondaryRewardValue').send_keys(prize_draw_entries)



        driv.find_element_by_id('OutcomeCompleteSecondaryRewardValue').send_keys(edge_credits)
        driv.find_element_by_id('OutcomeCompleteSecondaryRewardValue').send_keys(edge_credits)
        driv.find_element_by_id('OutcomeCompleteSecondaryRewardValue').send_keys(edge_credits)
        driv.find_element_by_id('OutcomeCompleteSecondaryRewardValue').send_keys(edge_credits)

        # reas_elem = driv.find_element_by_id('Reason')  # find the 'Reason' text box on web page using its element ID
        # reas_elem.clear()  # delete any text present in that field
        # reas_elem.send_keys(reas)  # enter Reason string
        # quan_elem = driv.find_element_by_id('Quantity')  # find the 'Quantity' text box on web page using its element ID
        # quan_elem.clear()  # delete any text present in that field
        # quan_elem.send_keys(quan)  # enter Quantity string
        # quan_elem.send_keys(Keys.ENTER)  # Press Enter key
    except:
        print(f"(in enter data function) - issue arose.")


def grab_redirects(driv):
    try:  # structured as try / except statement in case something's gone wrong
        quota_full_url = driv.find_element_by_id('OutcomeFullUrl').get_attribute('value')  # find the right element and grab URL from box
        screened_url = driv.find_element_by_id('OutcomeScreenedUrl').get_attribute('value')  # find the right element and grab URL from box
        complete_url = driv.find_element_by_id('OutcomeCompleteUrl').get_attribute('value')  # find the right element and grab URL from box
        quota_full_url = quota_full_url[0:83]  # trim off the last 3 characters
        screened_url = screened_url[0:83]  # trim off the last 3 characters
        complete_url = complete_url[0:83]  # trim off the last 3 characters
        return quota_full_url, screened_url, complete_url
    except:
        print(f"(in enter data function) - issue arose.")


chrome_path = cfg.chrome_path  # location of chromedriver.exe on local drive
driver = webdriver.Chrome(chrome_path)  # specify webdriver (selenium)
login(driver)

enter_data(driver, survey_name)
grab_redirects(driver)













# TODO: capture redirect info


# TODO: input required data





# TODO: create project dir - a directory in the appropriate client folder - NB will have to occur after DB is queried
# new_dir_path = cfg.projects_dir_path + "\\" + client_name + "\\" + p_number_to_search + " - " + survey_name  # clientname + projectname to be redefined by SQL queries
# print(new_dir_path)
# os.mkdir(new_dir_path)  # creates new directory
# subprocess.Popen(f'explorer "{new_dir_path}"')  # opens new dir in windows explorer
# TODO: create excel file with redirects


conn.close()

