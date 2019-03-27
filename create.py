import os, time, pprint, logging, sqlite3, subprocess, pyautogui, shutil, send2trash, datetime, calendar, sys, zcrmsdk
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from config import Config   # this imports the config file where the private data sits
import pandas as pd
from dateutil.relativedelta import *
import openpyxl
from openpyxl.styles import Font, Border, Side
import bs4
import re
from zcrmsdk import *
from pprint import pprint

# to avoid errors:
# Client directory must exist
# Survey name must be unique in xls

logging.basicConfig(level=logging.DEBUG, format=' %(asctime)s - %(levelname)s - %(message)s')  # turns on logging
# logging.disable(logging.CRITICAL)     # switches off logging when desired

cfg = Config()  # create an instance of the Config class, essentially brings private config data into play
os.chdir(cfg.cwd)  # change the current working directory to the one stipulated in config file


def pass_in_survey_name():    # reads in arg string from batch file
    if len(sys.argv) > 1:
        surv_name = str(sys.argv[1])  # takes the desired survey name from the command line arg, passed by the batch file
    else:
        surv_name = "No arguments passed"
    return surv_name


def generate_dates_sa():
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
    this_time_last_month = now + relativedelta(months=-1)  # date time object for today's date + time a month ago
    last_month = this_time_last_month.month  # isolating just the number of last month
    year_last_month = this_time_last_month.year  # isolating just the year it was last month
    last_month_string = str(last_month)  # convert number of last month to a string
    if len(last_month_string) == 1:  # if last month is only 1 digit...
        last_month_string = "0" + last_month_string  # ...add a leading zero
    proposal_date_string = "01/" + str(last_month_string) \
                      + "/" + str(year_last_month)  # compile full prop date string
    return start_date_string, end_date_string, proposal_date_string


def generate_closing_date():
    close_month_trimmed = close_month_raw[0:10]
    close_month = datetime.datetime.strptime(close_month_trimmed, '%Y-%m-%d')
    last_day_in_close_month = calendar.monthrange(close_month.year, close_month.month)[1]
    closing_date_string = str(last_day_in_close_month) + '/' + str(close_month.month) + '/' + str(close_month.year)
    return closing_date_string


def date_reshuffler(original_date):
    revised_date = original_date[6:10] + "-" + original_date[3:5] + "-" + original_date[0:2]
    logging.debug(f"revised date is {revised_date}")
    return revised_date


# p_number = 'P-46262'  # this is hardcoded for testing - won't be needed once Zoho is in the chain
# survey_name_to_search = 'Test KP'  # this is hardcoded for testing - won't be needed once argument passed in bat file

survey_name_to_search = pass_in_survey_name()  # grabs survey name from batch file as argument


start_date, end_date, proposal_date = generate_dates_sa()  # generates start and end dates through the function


# define all variables here
qf_msg = cfg.qf_msg
so_msg = cfg.so_msg
comp_msg = cfg.comp_msg

# varied for each project
p_number_col = 'SE Project Number'

# Project Tracking Sheet - columns of interest for SQL query
survey_name_col = 'Survey Name'
status = 'Active'
topic_col = 'Topic'
expected_loi_col = 'Expected LOI'
client_name_col = 'Client name'
sales_contact_col = 'Sales Contact'
edge_credits_col = 'Edge Credits'
close_month_col = 'Close month'

# Fixed variables - same for all projects
external_survey_url = 'tbc'
prize_draw_entries = '1'
qf_outcome_reward_id = 'Disqualified Survey - Regular Prize Draw'
so_outcome_reward_id = 'Disqualified Survey - Regular Prize Draw'
comp_outcome_reward_id = 'Completed Survey - Regular Prize Draw'
comp_secondary_reward_type = 'Credits'
tc_filepath = cfg.tc_filepath

excel_name_path = cfg.live_excel_file_path + "\\" + cfg.live_excel_filename  # LIVE MODE VERSION
excel_file_name_path_ext = excel_name_path + ".xlsm"  # LIVE MODE VERSION
excel_filename = cfg.live_excel_filename  # LIVE MODE VERSION

# excel_name_path = cfg.test_excel_file_path + "\\" + cfg.test_excel_filename  # TEST MODE VERSION
# excel_file_name_path_ext = excel_name_path + ".xlsm"  # TEST MODE VERSION
# excel_filename = cfg.test_excel_filename  # TEST MODE VERSION


conn = sqlite3.connect(excel_filename + ".db")
c = conn.cursor()

df = pd.read_excel(excel_file_name_path_ext, sheet_name='PPT')  # create dataframe from xlsm content
df.to_sql('PPT', conn)  # populate database with dataframe content
conn.commit()

table_name = "PPT"


# 1) Contents of columns of interest for row that matches survey name
c.execute('SELECT "{coi2}","{coi3}","{coi4}","{coi5}","{coi6}","{coi7}" FROM {tn} WHERE "{cn}"="{scn}"'.format(tn=table_name, cn=survey_name_col, coi2=topic_col, coi3=expected_loi_col, coi4=client_name_col, coi5=sales_contact_col, coi6=edge_credits_col, coi7=close_month_col, scn=survey_name_to_search))  # note I need to put speech marks around "{cn}" because the column name contains a space
all_rows = c.fetchall()
print(f"Searched on project name '{survey_name_to_search}'")
print('Project row looked up and found in excel db looks like this:')
print(all_rows)
conn.close()


# New SQL variable names to be used in Zoho and then Survey Admin
survey_name = survey_name_to_search  # assign outputs to variable names
topic = str(all_rows[0][0])
expected_loi = str(all_rows[0][1])
client_name = str(all_rows[0][2])
sales_contact = str(all_rows[0][3])
edge_credits = str(int(all_rows[0][4]))
close_month_raw = all_rows[0][5]


# define fixed variables for Zoho
zoho_url = cfg.zoho_create_potential_URL
industry = "Other"
account_type = "Research Panel"
stage = "Closed Won - Signed IO Received"
campaign_start_date = start_date[0:10]  # same as start date in survey admin, which is today's date, but trimmed
closing_date = generate_closing_date()
campaign_end_date = closing_date  # same as closing date



def login_sa():
    driver.get(cfg.create_survey_URL)  # use selenium webdriver to open web browser and desired URL from config file
    driver.execute_script("document.getElementById('UserName').value = '" + str(cfg.uname) + "';")  # insert username
    driver.execute_script("document.getElementById('Password').value = '" + str(cfg.pwd) + "';")  # insert password
    pass_elem = driver.find_element_by_id('Password')  # find the 'Password' text box using its element ID
    pass_elem.submit()  # submit password
    time.sleep(2)   # wait 2 seconds for the login process to take place (tested and this is necessary)


def login_zoho():
    driver.get(cfg.zoho_login_url)  # use selenium webdriver to open web browser and desired URL from config file
    driver.execute_script("document.getElementById('lid').value = '" + str(cfg.zoho_uname) + "';")  # insert username
    driver.execute_script("document.getElementById('pwd').value = '" + str(cfg.zoho_pw) + "';")  # insert password
    pass_elem = driver.find_element_by_id('pwd')  # find the 'Password' text box using its element ID
    pass_elem.submit()  # submit password
    time.sleep(2)   # wait 2 seconds for the login process to take place (tested and this is necessary)
    driver.get(cfg.zoho_create_potential_URL)


def enter_data_zoho():
    driver.find_element_by_id('Crm_Potentials_CONTACTID').send_keys(sales_contact)  # inconsistent with execute_script, same with send_keys. It enters, but can't make it save
    time.sleep(1)
    pyautogui.press('tab')
    # driver.execute_script("document.getElementById('Crm_Potentials_CONTACTID').value = '" + str(sales_contact) + "';")  # sometimes this doesn't input so starting with it
    time.sleep(1)
    driver.find_element_by_id('select2-Crm_Potentials_POTENTIALCF9-container').click()  # industry. Couldn't figure out how to make selenium do this so had to use pyautogui
    time.sleep(1)
    pyautogui.typewrite(industry)
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    driver.execute_script("document.getElementById('Crm_Potentials_ACCOUNTID').value = '" + str(client_name) + "';")
    driver.execute_script("document.getElementById('Crm_Potentials_POTENTIALNAME').value = '" + str(survey_name) + "';")
    driver.find_element_by_id('select2-Crm_Potentials_POTENTIALCF10-container').click()  # account_type. Couldn't figure out how to make selenium do this so had to use pyautogui
    pyautogui.typewrite(account_type)
    pyautogui.press('enter')
    driver.execute_script("document.getElementById('Crm_Potentials_POTENTIALCF86').value = '" + str(proposal_date) + "';")
    driver.execute_script("document.getElementById('Crm_Potentials_CLOSINGDATE').value = '" + str(closing_date) + "';")
    driver.find_element_by_id('select2-Crm_Potentials_STAGE-container').click()  # stage. Couldn't figure out how to make selenium do this so had to use pyautogui
    pyautogui.typewrite(stage)
    pyautogui.press('enter')
    driver.execute_script("document.getElementById('Crm_Potentials_POTENTIALCF84').value = '" + str(campaign_start_date) + "';")
    driver.execute_script("document.getElementById('Crm_Potentials_POTENTIALCF83').value = '" + str(campaign_end_date) + "';")
    time.sleep(5)
    driver.find_element_by_id('savePotentialsBtn').click()  # save potential
    time.sleep(5)

    # TODO: figure out how to put the Contact Name in and make it stick. Might need to be very manual

    # p_num_location = driver.find_element_by_id("subvalue_POTENTIALCF6")

    # ID 'subvalue_CONTACTID' - this one seems to work but no value is then returned (p_num is None)
    # print(f"p_num_location is {p_num_location}")

    # This is a way to take an element and find all attributes for it - e.g. to try to track down the value. Will come in handy later.
    # VARIABLE 1 - p_num_location
    # attempt to get all attributes of element (method 1):
    # attributes = driver.execute_script('var items = {}; for (index = 0; index < arguments[0].attributes.length; ++index) { items[arguments[0].attributes[index].name] = arguments[0].attributes[index].value }; return items;', p_num_location)
    # pprint.pprint(attributes)  # attempting to list all attributes for p_num_location to then see if it has a 'value'
    # # attempt to get all attributes of element (method 2):
    # p1_method2 = p_num_location.get_property('attributes')[0]
    # pprint.pprint(p1_method2)

    # p_num = p_num_location.get_attribute("value_Reference Number")
    # print(f"p_num is {p_num}")
    # driver.find_element_by_id('subvalue_CONTACTID').click()  # ID 'subvalue_CONTACTID' or 'value_CONTACTID' or 'labelTD_CONTACTID'
    #
    # pyautogui.typewrite(sales_contact)
    # driver.find_element_by_name('button__CONTACTID').click()
    # LAST ROW COMMENTED OUT TO AVOID ACTUAL POTENTIAL CREATION  ###########


def grab_p_number():
    # test_html_file = open(cfg.test_html_file)  # on for test mode
    # soup = bs4.BeautifulSoup(test_html_file, "html.parser")  # on for test mode
    content = driver.page_source  # on for live mode
    soup = bs4.BeautifulSoup(content, "html.parser")  # on for live mode
    soup_string = str(soup)
    # print(soup_string)
    text_segment = r"P-\d\d\d\d\d"  # hopefully I've used those escape characters correctly
    regex = re.compile(text_segment)
    mo = regex.findall(soup_string)
    p_num = mo[0]
    return p_num


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
    # print(f'Quota Full: {qf}')
    # print(f'Screened: {so}')
    # print(f'Complete: {comp}')
    logging.debug("Creating new directory:", new_project_dir_path)
    if not os.path.exists(new_project_dir_path):
        os.mkdir(new_project_dir_path)  # creates new directory
    create_redirects_xls(qf, so, comp)


def create_redirects_xls(q, s, c):
    wb = openpyxl.Workbook()

    sheet1 = wb.active
    sheet1.title = f'Redirects - {p_number}'
    sheet1['A1'] = f'Redirects for {p_number} - {survey_name}'
    sheet1['B2'] = 'Quota Full:'
    sheet1['B3'] = 'Screened:'
    sheet1['B4'] = 'Complete:'
    sheet1['C2'] = q
    sheet1['C3'] = s
    sheet1['C4'] = c
    sheet1.column_dimensions['B'].width = 15
    sheet1.column_dimensions['C'].width = 95

    emboldened = Font(bold=True)
    sheet1['A1'].font = emboldened
    sheet1['B2'].font = emboldened
    sheet1['B3'].font = emboldened
    sheet1['B4'].font = emboldened

    thin = Side(border_style='thin')
    surrounded = Border(top=thin, left=thin, right=thin, bottom=thin)
    sheet1['B2'].border = surrounded
    sheet1['B3'].border = surrounded
    sheet1['B4'].border = surrounded
    sheet1['C2'].border = surrounded
    sheet1['C3'].border = surrounded
    sheet1['C4'].border = surrounded

    wb.save(redirects_wb_path_name_ext)


def enter_data_sa():
    driver.find_element_by_id('Name').send_keys(survey_name)  # using send_keys instead of script command here due to potential inclusion of apostrophes etc which stuff up the js syntax
    driver.execute_script("document.getElementById('Status').value = '" + str(status) + "';")
    driver.execute_script("document.getElementById('Title').value = '" + str(topic) + "';")
    driver.execute_script("document.getElementById('ProjectIONumber').value = '" + str(p_number) + "';")
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
    # COMMENT OUT THE LAST ROW FOR TEST MODE, TO AVOID ACTUAL PROJECT CREATION  ###########


def clean_up():
    send2trash.send2trash(cfg.live_excel_filename + ".db")
    send2trash.send2trash(cfg.test_excel_filename + ".db")


# Variable Definition

chrome_path = cfg.chrome_path  # location of chromedriver.exe on local drive
chrome_options = Options()
chrome_options.add_argument("--disable-notifications")  # to disable notifications popup in Chrome (affects Zoho page)
driver = webdriver.Chrome(chrome_path, chrome_options=chrome_options)  # specify webdriver (chrome via selenium)

# PULLING LEVERS HERE #############
clean_up()
login_zoho()
enter_data_zoho()
p_number = grab_p_number()

new_project_dir_path = cfg.projects_dir_path + "\\" + client_name + "\\" + p_number + " - " + survey_name
redirects_wb_path_name_ext = new_project_dir_path + "\\" + p_number + " redirects.xlsx"

login_sa()
establish_project_dir()
enter_data_sa()

subprocess.Popen(f'explorer "{new_project_dir_path}"')  # opens new dir in windows explorer  # DISABLE FOR TESTING
subprocess.Popen(f'explorer "{redirects_wb_path_name_ext}"')  # opens file in windows  # DISABLE FOR TESTING
subprocess.Popen(f'explorer "{excel_file_name_path_ext}"')  # opens Survey Tracking file in windows, so I can add in project number manually  # DISABLE FOR TESTING


# TODO: update Zoho section so it uses API instead of GUI
