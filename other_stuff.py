# import win32com.client
import os

# excelApp = win32com.client.dynamic.Dispatch('Excel.Application')
# excelApp.Application.Quit()



# TODO: Replace above chunk with a function which does the following:
# TODO: make copy of live excel file
local_filename = 'local_file'
local_file_name_path = os.getcwd() + "\\" + local_filename
local_file_name_path_ext = local_file_name_path + ".xlsm"



# shutil.copyfile(live_excel_file_name_path_ext, local_file_name_path_ext)  # ERROR - must be closed to avoid permission error. Try via 'xlrd' module or through Reddit responses to my query



# TODO: delete the copy
# send2trash.send2trash(local_file_name_path_ext)





"""
# Old SQL query - no longer necessary as we are pulling all the same data but using Survey Name instead of P-number
# TODO: to prep data for Admin Survey Creation, using example P-number var, look up all variables of interest in the SQL database

# 1) Contents of columns of interest for row that matches P-number
c.execute('SELECT "{coi1}","{coi2}","{coi3}","{coi4}","{coi5}","{coi6}" FROM {tn} WHERE "{cn}"="{scn}"'.format(tn=table_name, cn=p_number_col, coi1=survey_name_col, coi2=topic_col, coi3=expected_loi_col, coi4=client_name_col, coi5=sales_contact_col, coi6=edge_credits_col, scn=p_number))  # note I need to put speech marks around "{cn}" because the column name contains a space
all_rows = c.fetchall()

conn.close()

# Old SQL variable names - no longer necessary as the query is different so doesn't match with the variable names
survey_name = str(all_rows[0][0])  # assign outputs to variable names
topic = str(all_rows[0][1])
expected_loi = str(all_rows[0][2])
client_name = str(all_rows[0][3])
sales_contact = str(all_rows[0][4])
edge_credits = str(int(all_rows[0][5]))
"""
