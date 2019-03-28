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



# login_zoho()  # not needed in API mode
# enter_data_zoho()  # not needed in API mode
# p_number = grab_p_number()  # not needed in API mode
