# Import the necessary packages
import os
import time
import pandas as pd


# User Parameters/Constants to Set
AOI_PUBLIC_PATH = "//mcrtp-file-01.mcusa.local/public/000-AOI_Tool_Output/"
EXCEL_TEMPLATE_PATH = "C:/Users/ait.lab/.spyder-py3/Wafer_Map_Creator/Report_Template.xlsx"


def time_convert(sec):
    mins = sec // 60
    sec = sec % 60
    hours = mins // 60
    mins = mins % 60
    print("Time Lapsed = {0}h:{1}m:{2}s".format(int(hours), int(mins), round(sec) ) )



# Imports Excel report template to use
report_to_export = pd.read_excel(EXCEL_TEMPLATE_PATH, 
                                 "Report",
                                 header=None)


# Runs through each project file in the AOI output public sharedrive folder
for project_file_name in os.listdir(AOI_PUBLIC_PATH):
    
    # This script only runs for HBCOSA lots, so if not that lot, then skip
    if "Window" not in project_file_name:
        continue
    
    # Creates file name
    project_file_path = os.path.join(AOI_PUBLIC_PATH, project_file_name)
    
    # Runs through each slot file
    for slot_file_name in os.listdir(project_file_path):
        
        # If not Excel sheet folder, then skip
        if "Excel" not in slot_file_name:
            continue
        
        # Creates slot path
        slot_file_path = os.path.join(project_file_path, slot_file_name)
        
        # If report excel file already exist, skip to next project
        if any("Report" in s for s in os.listdir(slot_file_path)):
            break
        
        # Runs through each Excel sheet
        for excel_file_name in os.listdir(slot_file_path):
            
            # Skips report excel sheet
            if "Report" in excel_file_name:
                continue
            
            # Creates Excel file path
            excel_file_path = os.path.join(slot_file_path, excel_file_name)
            
            # imports excel sheet data
            report_to_import = pd.read_excel(excel_file_path, 
                                             sheet_name=excel_file_name.replace(".xlsx",""),
                                             header=None)
            
            