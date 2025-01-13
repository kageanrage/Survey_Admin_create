import os, time
from pyotp import TOTP
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from config import Config
import pyautogui, random, pyperclip, se_general
from bs4 import BeautifulSoup
import se_admin
import openpyxl
from config import Config
from pprint import pprint
from collections import OrderedDict


def read_in_template(cfg):
    wb = openpyxl.load_workbook(cfg.cmv_quota_template, data_only=True)
    sheet = wb.active
    headers_range = sheet['A1':'N1']
    headers_horizontal = headers_range[0]

    contents_range = sheet['A2':'N9']
    # pprint(contents_range[0])

    dicts_list = []

    for i in range(0, len(contents_range)):
        quota_dict = OrderedDict()  # may  want this as regular dict as ordered dict may have weird tuples in it
        for index, cell in enumerate(contents_range[i], 0):
            quota_dict.setdefault(headers_horizontal[index].value, cell.value)
        dicts_list.append(quota_dict)

    return dicts_list


def create_a_single_quota(driver, quota_dict, survey_id):
    # WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.LINK_TEXT,
    #                                                             'Add a new quota')))  # wait til button visible before attempting to click
    driver.get(f"https://data.studentedge.org/admin/survey/createquota?surveyId={survey_id}")
    time.sleep(2)
    # time.sleep(2)
    name_field = driver.find_element('id', 'Name')
    name_field.send_keys(quota_dict['Quota_name'])
    if quota_dict['Gender']:
        gender_dropdown = Select(driver.find_element("id", "Demographic_Gender")).select_by_visible_text(f"{quota_dict['Gender']}")

    if quota_dict['Region']:
        metro_regional_dropdown = Select(driver.find_element("id", "Area")).select_by_visible_text(f"{quota_dict['Region']}")

    target_field = driver.find_element('id', 'Target')
    target_field.clear()
    target_field.send_keys('100')

    min_age_field = driver.find_element('id', 'Demographic_MinimumAge')
    min_age_field.clear()
    min_age_field.send_keys(f"{quota_dict['Min age']}")
    max_age_field = driver.find_element('id', 'Demographic_MaximumAge')
    max_age_field.clear()
    max_age_field.send_keys(f"{quota_dict['Max age']}")

    state_names_dict = {"NSW": "New South Wales, Australia",
                       "ACT": "Australian Capital Territory, Australia",
                       "VIC": "Victoria, Australia",
                       "QLD": "Queensland, Australia",
                       "SA": "South Australia, Australia",
                       "WA": "Western Australia, Australia",
                       "TAS": "Tasmania, Australia",
                       "NT": "Northern Territory, Australia"
                       }

    state_options = ["NSW", "ACT", "VIC", "QLD", "SA", "WA", "TAS", "NT"]

    states_to_add = []
    for state in state_options:
        if quota_dict[state]:
            states_to_add.append(state_names_dict[state])

    state_selector = Select(driver.find_element("id", "AllLocationOptions"))
    for state in states_to_add:
        state_selector.select_by_visible_text(f'{state}')

    move_to_right_button = driver.find_element('link text', '>')
    move_to_right_button.click()

    # section save button currently disabled
    save_new_quota_button = driver.find_element('css selector', '.green')
    save_new_quota_button.click()


# this function ties it all together for external access
def generate_cmv_quotas(cfg, driver, survey_id):
    print('running generate_cmv_quotas()')
    quota_inputs = read_in_template(cfg)
    for i in range(0, len(quota_inputs)):
        create_a_single_quota(driver, quota_inputs[i], survey_id)


def main():
    cfg = Config()  # create an instance of the Config class, essentially brings private config data into play
    os.chdir(cfg.cwd)  # change the current working directory to the one stipulated in config file

    survey_id = "test_survey_id"  # basically serving as a placeholder for now

    quota_inputs = read_in_template(cfg)

    # configure webdriver
    chrome_path = cfg.chrome_path  # location of chromedriver.exe on local drive
    chrome_options = Options()
    chrome_options.add_argument(
        "--disable-notifications")  # to disable notifications popup in Chrome (affects Zoho page)
    chrome_options.add_experimental_option("detach", True)
    driver = webdriver.Chrome(chrome_path, options=chrome_options)  # specify webdriver (chrome via selenium)
    driver.implicitly_wait(10)

    se_admin.login_sa(driver, cfg.create_survey_URL)  # now using fn from module
    driver.get(cfg.kp_test_sa_project_url)
    # time.sleep(3)
    for i in range(0, len(quota_inputs)):
        create_a_single_quota(driver, quota_inputs[i], survey_id)


if __name__ == '__main__':
    main()

