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
