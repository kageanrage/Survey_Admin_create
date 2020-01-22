# TEST / LIVE MODE DETERMINING VARIABLES
survey_name_to_search = pass_in_survey_name()  # LIVE MODE - grabs survey name from batch file as argument

excel_name_path = cfg.live_excel_file_path + "\\" + cfg.live_excel_filename  # LIVE MODE VERSION
excel_file_name_path_ext = excel_name_path + ".xlsx"  # LIVE MODE VERSION
excel_filename = cfg.live_excel_filename  # LIVE MODE VERSION


clean_up()

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

c.execute('DELETE FROM {tn} WHERE "Sales Contact" IS NULL'.format(tn=table_name))  # deletes db rows below last valid project row so that last row can be accurately located

if survey_name_to_search == "last_row_in_table":
    c.execute('SELECT "{coi2}","{coi3}","{coi4}","{coi5}","{coi6}","{coi7}","{coi8}" FROM {tn} ORDER BY "index" DESC LIMIT 1'.format(tn=table_name, coi2=topic_col, coi3=expected_loi_col, coi4=client_name_col, coi5=sales_contact_col, coi6=edge_credits_col, coi7=close_month_col, coi8=survey_name_col))  # note I need to put speech marks around "{cn}" because the column name contains a space
else:
    c.execute('SELECT "{coi2}","{coi3}","{coi4}","{coi5}","{coi6}","{coi7}" FROM {tn} WHERE "{cn}"="{scn}"'.format(tn=table_name, cn=survey_name_col, coi2=topic_col, coi3=expected_loi_col, coi4=client_name_col, coi5=sales_contact_col, coi6=edge_credits_col, coi7=close_month_col, scn=survey_name_to_search))  # note I need to put speech marks around "{cn}" because the column name contains a space
all_rows = c.fetchall()
print(f"Searched on project name '{survey_name_to_search}'")
print('Project row looked up and found in excel db looks like this:')
print(all_rows)
conn.close()
