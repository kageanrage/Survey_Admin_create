import os, time, pprint, logging, sqlite3, subprocess, pyautogui, send2trash, datetime, calendar, sys, zcrmsdk, pyperclip
from selenium.webdriver.common.action_chains import ActionChains
from config import Config   # this imports the config file where the private data sits
import pandas as pd
from dateutil.relativedelta import *
import openpyxl
from openpyxl.styles import Font, Border, Side, PatternFill
from zcrmsdk import *
from pprint import pprint
import se_general, se_admin, se_zoho


logging.basicConfig(level=logging.DEBUG, format=' %(asctime)s - %(levelname)s - %(message)s')  # turns on logging
# logging.disable(logging.CRITICAL)     # switches off logging when desired

cfg = Config()  # create an instance of the Config class, essentially brings private config data into play
os.chdir(cfg.cwd)  # change the current working directory to the one stipulated in config file

# proj_dict for newly created project


def determine_exclusion_survey_ids(survey_name):
    proj_dict = se_general.look_up_project(survey_name)
    if str(proj_dict['survey names to exclude']) != 'nan':  # if survey names listed for exclusion
        print('Survey names to exclude looks like this:')
        print(proj_dict['survey names to exclude'])
        str_of_survey_names_to_exclude = proj_dict['survey names to exclude']
        excl_surv_names = str_of_survey_names_to_exclude.split(',')

        survey_id_excl_list = []

        for s_name in excl_surv_names:
            # proj_dict for each of the survey ids I'm looking up
            proj_dict = se_general.look_up_project(s_name)
            survey_id_excl_list.append(proj_dict['Survey ID'])

        print('survey ids to exclude list looks like this:')
        print(survey_id_excl_list)
    elif str(proj_dict['survey ids to exclude']) != 'nan':  # if survey_ids are manually listed for exclusion
        survey_id_excl_list = proj_dict['survey ids to exclude']
        print('survey ids to exclude list looks like this:')
        print(survey_id_excl_list)
    else:
        print('No survey ids to exclude')
        survey_id_excl_list = ""

    return survey_id_excl_list


proj_dict = se_general.look_up_latest_project()
survey_name = proj_dict['Survey Name']
determine_exclusion_survey_ids(survey_name)