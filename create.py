import os, time, pprint, logging, sqlite3, subprocess, pyautogui, shutil, send2trash, datetime, calendar, sys, zcrmsdk, pyperclip
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


# FUNCTIONS ################################################


def clean_up():
    os.chdir(cfg.repo_dir)
    send2trash.send2trash(cfg.live_excel_filename + ".db")
    send2trash.send2trash(cfg.test_excel_filename + ".db")


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
    print(f"close_date_raw is {close_date_raw}")
    close_date_trimmed = close_date_raw[0:10]
    print(f"close_date_trimmed is {close_date_trimmed}")
    close_date = datetime.datetime.strptime(close_date_trimmed, '%Y-%m-%d')
    last_day_in_close_month = calendar.monthrange(close_date.year, close_date.month)[1]
    close_month_string = str(close_date.month)  # convert number of close month to a string
    print(f"close_month_string is {close_month_string}")
    if len(close_month_string) == 1:  # if close month is only 1 digit...
        close_month_string = "0" + close_month_string  # ...add a leading zero
    print(f"After adding a leading zero where needed, close_month_string is now {close_month_string}")
    closing_date_string = str(last_day_in_close_month) + '/' + close_month_string + '/' + str(close_date.year)
    print(f"closing_date_string is {closing_date_string}")
    return closing_date_string


def date_reshuffler(original_date):
    logging.debug(f"original date is {original_date}")
    revised_date = original_date[6:10] + "-" + original_date[3:5] + "-" + original_date[0:2]
    logging.debug(f"revised date is {revised_date}")
    return revised_date


def search_contact_by_name(keyword):
    logging.debug('Function: search_contact_by_name')
    try:
        module_ins = ZCRMModule.get_instance('Contacts')  # module API Name
        resp = module_ins.search_records(keyword)  # search key word
        # print(resp.status_code)
        resp_info = resp.info
        # print('resp_info looks like this:')
        # print(resp_info)
        record_ins_arr = resp.data
        # print('record_ins_array looks like this:')
        # print(record_ins_arr)
        logging.debug(f'Number of contacts found with this name (i.e. len of record_ins_arr) is {len(record_ins_arr)}')
        first_record = record_ins_arr[0]
        contact_id = first_record.entity_id

        record_ins_item_1_data = record_ins_arr[0].field_data
        # print('record looks like this:')
        # pprint(record_ins_item_1_data)
        full_name = record_ins_item_1_data['Full_Name']

        logging.debug(f"Returning Contact ID corresponding to {full_name}")
        logging.debug(f"ID found: {contact_id}")
        return contact_id

    except zcrmsdk.ZCRMException as ex:
        print(ex.status_code)
        print(ex.error_message)
        print(ex.error_code)
        print(ex.error_details)
        print(ex.error_content)


def create_potential():
    logging.debug('Function: create_potential')
    try:
        record = ZCRMRecord.get_instance('Potentials')  # module API Name

        record.set_field_value('Account_Name', client_name)
        record.set_field_value('Deal_Name', survey_name)
        record.set_field_value('Industry', industry)
        record.set_field_value('Account_Type', account_type)
        record.set_field_value('Campaign_Start_Date', campaign_start_date_api)  # note this format is different from the one I generated for the selenium input
        record.set_field_value('Campaign_End_Date', campaign_end_date_api)  # note this format is different from the one I generated for the selenium input
        record.set_field_value('Proposal_Date_Sent', proposal_date_api)  # note this format is different from the one I generated for the selenium input
        record.set_field_value('Closing_Date', closing_date_api)  # note this format is different from the one I generated for the selenium input
        record.set_field_value('Stage', stage)

        contact_id = search_contact_by_name(sales_contact)  # uses my fn to grab the contact's ID
        ac_dynamic_dict = {'name': sales_contact, 'id': contact_id}  # ... then pipes the ID and name into this dict for insertion into potential
        # print('ac_dynamic_dict looks like this:')
        # print(ac_dynamic_dict)
        # print('Attempting to set Contact_Name using dynamic dict')
        record.set_field_value('Contact_Name', ac_dynamic_dict)

        resp = record.create()
        new_potential_id = record.entity_id  # grabs the entity ID of the record, to then use this to look up the newly created potential
        print(resp.status_code)
        logging.debug(f"New potential's ID is {new_potential_id}")
        return new_potential_id

    except ZCRMException as ex:
        print(ex.status_code)
        print(ex.error_message)
        print(ex.error_code)
        print(ex.error_details)
        print(ex.error_content)


def get_potential_record_by_id(id):
    try:
        record = ZCRMRecord.get_instance('Potentials', id)
        resp = record.get()
        # print(resp.status_code)
        # print(f"entity ID for this record is {resp.data.entity_id}")
        # print(resp.data.created_by.id)
        # print(resp.data.modified_by.id)
        # print(resp.data.owner.id)
        # print(resp.data.created_by.name)
        # print(resp.data.created_time)
        # print(resp.data.modified_time)
        # print(resp.data.get_field_value('Email'))
        # print(resp.data.get_field_value('Last_Name'))
        # print('######### Attempting to display Full_Name - if this works I can return this value')
        deal_name = resp.data.get_field_value('Deal_Name')
        p_number = resp.data.get_field_value('Reference_Number')
        # print(full_name)
        print('Data for the potential looks like this:')
        pprint(resp.data.field_data)
        # if resp.data.line_items is not None:
        #     for line_item in resp.data.line_items:
        #         print("::::::LINE ITEM DETAILS::::::")
        #         print(line_item.id)
        #         print(line_item.product.get_field_value('Product_Code'))
        return p_number

    except ZCRMException as ex:
        print(ex.status_code)
        print(ex.error_message)
        print(ex.error_code)
        print(ex.error_details)
        print(ex.error_content)


def login_sa():
    driver.get(cfg.create_survey_URL)  # use selenium webdriver to open web browser and desired URL from config file
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
    driver.execute_script("document.getElementById('OutcomeCompleteRewardValue').value = '" + str(edge_credits) + "';")
    driver.find_element_by_id('FullOutcomeRewardId').send_keys(qf_outcome_reward_id)  # this didn't work via JS execution method
    driver.find_element_by_id('ScreenedOutcomeRewardId').send_keys(so_outcome_reward_id)  # this didn't work via JS execution method
    # driver.find_element_by_id('CompleteOutcomeRewardId').send_keys(str_saying_comp_surv_reg_p_d)  # 23-06 commented out so no longer changes from default
    driver.find_element_by_id('CompleteOutcomeSecondaryRewardId').send_keys(str_saying_comp_surv_reg_p_d)  # 23-06 added
    driver.execute_script("document.getElementById('OutcomeCompleteSecondaryRewardValue').value = '" + str(prize_draw_entries) + "';")
    driver.find_element_by_id('OutcomeCompleteSecondaryRewardType').send_keys('Reward')
    driver.find_element_by_id('OutcomeCompleteRewardType').send_keys(the_word_credits)  # added 23-06-19
    driver.find_element_by_id('TermsAndConditionsPdf').click()
    time.sleep(2)
    pyautogui.typewrite(tc_filepath)  # since popup window is outside web browser, need a diff package to control
    time.sleep(2)
    pyautogui.press('enter')
    time.sleep(2)
    driver.find_element_by_css_selector('#add-edit-survey > fieldset > dl > div.form_navigation > button').click()  # Submits / creates new project
    # COMMENT OUT THE LAST ROW FOR TEST MODE, TO AVOID ACTUAL PROJECT CREATION  ###########


# TEST / LIVE MODE DETERMINING VARIABLES
# p_number = 'P-46262'  # this is hardcoded for testing - won't be needed once Zoho is in the chain
# survey_name_to_search = 'KP test 28-03-19'  # TEST MODE - this is hardcoded for testing - won't be needed once argument passed in bat file
survey_name_to_search = pass_in_survey_name()  # LIVE MODE - grabs survey name from batch file as argument

excel_name_path = cfg.live_excel_file_path + "\\" + cfg.live_excel_filename  # LIVE MODE VERSION
excel_file_name_path_ext = excel_name_path + ".xlsm"  # LIVE MODE VERSION
excel_filename = cfg.live_excel_filename  # LIVE MODE VERSION

# excel_name_path = cfg.test_excel_file_path + "\\" + cfg.test_excel_filename  # TEST MODE VERSION
# excel_file_name_path_ext = excel_name_path + ".xlsm"  # TEST MODE VERSION
# excel_filename = cfg.test_excel_filename  # TEST MODE VERSION


# DATABASE / PROJECT TRACKING SHEET OPERATIONS
conn = sqlite3.connect(excel_filename + ".db")  # create database file
c = conn.cursor()  # define cursor
table_name = "PPT"  # define table name
df = pd.read_excel(excel_file_name_path_ext, sheet_name='PPT')  # create dataframe from xlsm content
df.to_sql('PPT', conn)  # populate database with dataframe content
conn.commit()  # commit to db aka save file


# VARIABLES ################################################
# DATABASE / PROJECT TRACKING SHEET
# column names of interest for SQL query
survey_name_col = 'Survey Name'
status = 'Active'
topic_col = 'Topic'
expected_loi_col = 'Expected LOI'
client_name_col = 'Client name'
sales_contact_col = 'Sales Contact'
edge_credits_col = 'Edge Credits'
close_month_col = 'Close month'
p_number_col = 'SE Project Number'


# DB query
c.execute('SELECT "{coi2}","{coi3}","{coi4}","{coi5}","{coi6}","{coi7}" FROM {tn} WHERE "{cn}"="{scn}"'.format(tn=table_name, cn=survey_name_col, coi2=topic_col, coi3=expected_loi_col, coi4=client_name_col, coi5=sales_contact_col, coi6=edge_credits_col, coi7=close_month_col, scn=survey_name_to_search))  # note I need to put speech marks around "{cn}" because the column name contains a space
all_rows = c.fetchall()
print(f"Searched on project name '{survey_name_to_search}'")
print('Project row looked up and found in excel db looks like this:')
print(all_rows)
conn.close()


# SQL variables for Zoho + Survey Admin
survey_name = survey_name_to_search  # assign outputs to variable names
new_project_name = survey_name  # redundant and can be adjusted when ready
topic = str(all_rows[0][0])
expected_loi = str(all_rows[0][1])
client_name = str(all_rows[0][2])
sales_contact = str(all_rows[0][3])
edge_credits = str(int(all_rows[0][4]))
close_date_raw = all_rows[0][5]

# ZOHO variables
# zoho_url = cfg.zoho_create_potential_URL
industry = "Other"
account_type = "Research Panel"
stage = "Closed Won - Signed IO Received"

# SURVEY ADMIN variables
# Fixed variables - same for all projects
qf_msg = cfg.qf_msg
so_msg = cfg.so_msg
comp_msg = cfg.comp_msg
external_survey_url = 'tbc'
prize_draw_entries = '1'
qf_outcome_reward_id = 'Disqualified Survey - Regular Prize Draw'
so_outcome_reward_id = 'Disqualified Survey - Regular Prize Draw'
str_saying_comp_surv_reg_p_d = 'Completed Survey - Regular Prize Draw'
the_word_credits = 'Credits'
tc_filepath = cfg.tc_filepath  # file path of T&Cs pdf file


# DATE VARIABLES
start_date, end_date, proposal_date = generate_dates_sa()  # generates start and end dates through the function
campaign_start_date = start_date[0:10]  # same as start date in survey admin, which is today's date, but trimmed
closing_date = generate_closing_date()
campaign_end_date = closing_date  # same as closing date
# format-adjusted date variables for use by API
proposal_date_api = date_reshuffler(proposal_date)
closing_date_api = date_reshuffler(closing_date)
campaign_start_date_api = date_reshuffler(campaign_start_date)
campaign_end_date_api = date_reshuffler(campaign_end_date)


# ZOHO API OPERATIONS ##########################
# 0 run this code every single time:
zcrmsdk.ZCRMRestClient.initialize()

# 2 - second chunk of code (run in isolation) - I ran this to attempt to generate 'access token through refresh token' i.e. add it to the token file
oauth_client = zcrmsdk.ZohoOAuth.get_client_instance()
refresh_token = cfg.refresh_token
user_identifier = cfg.zoho_uname
oauth_tokens = oauth_client.generate_access_token_from_refresh_token(refresh_token, user_identifier)

# 3 - if access token already refreshed in past hour, can proceed without any initialisation code apart from what's specified at '# 0'



# Zoho levers
new_job_id = create_potential()  # create the new potential and store its ID in this variable
p_number = get_potential_record_by_id(new_job_id)  # use that ID to look up the newly created potential and store its P-number in this variable
# print(f"p-number for new project is {p_number}")

# Survey Admin variables
chrome_path = cfg.chrome_path  # location of chromedriver.exe on local drive
chrome_options = Options()
chrome_options.add_argument("--disable-notifications")  # to disable notifications popup in Chrome (affects Zoho page)
driver = webdriver.Chrome(chrome_path, options=chrome_options)  # specify webdriver (chrome via selenium)

new_project_dir_path = cfg.projects_dir_path + "\\" + client_name + "\\" + p_number + " - " + survey_name
redirects_wb_path_name_ext = new_project_dir_path + "\\" + p_number + " redirects.xlsx"

# Survey Admin levers
login_sa()
establish_project_dir()
enter_data_sa()

subprocess.Popen(f'explorer "{new_project_dir_path}"')  # opens new dir in windows explorer  # DISABLE FOR TESTING
subprocess.Popen(f'explorer "{redirects_wb_path_name_ext}"')  # opens file in windows  # DISABLE FOR TESTING
subprocess.Popen(f'explorer "{excel_file_name_path_ext}"')  # opens Survey Tracking file in windows, so I can add in project number manually  # DISABLE FOR TESTING

clean_up()
pyperclip.copy(p_number)  # copy p_number to clipboard to then manually paste once script is done
