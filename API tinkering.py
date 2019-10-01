import os, time, pprint, logging, zcrmsdk
from config import Config   # this imports the config file where the private data sits

from zcrmsdk import *
from pprint import pprint

# to avoid errors:
# Survey name must be unique in xls
# For 'use last row in database', Sales Contact for that last one must be populated, and that column must be empty in excel from that last one onwards

logging.basicConfig(level=logging.DEBUG, format=' %(asctime)s - %(levelname)s - %(message)s')  # turns on logging
# logging.disable(logging.CRITICAL)     # switches off logging when desired

cfg = Config()  # create an instance of the Config class, essentially brings private config data into play
os.chdir(cfg.cwd)  # change the current working directory to the one stipulated in config file


# FUNCTIONS ################################################


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
        assert len(record_ins_arr) == 1, 'Number of contacts found was not equal to 1'
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
        logging.debug(f'deal_name = {deal_name}')
        p_number = resp.data.get_field_value('Reference_Number')
        logging.debug(f'p_number = {p_number}')
        account_name_in_potential = resp.data.get_field_value('Account_Name')['name']
        logging.debug(f'account_name = {account_name_in_potential}')
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
new_job_id = '1756226000007342001'  # hardcode an ID in here to look it up
p_number = get_potential_record_by_id(new_job_id)  # use that ID to look up the newly created potential and store its P-number in this variable