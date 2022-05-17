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



# Main
# =============================================================================
# Starting stopwatch to see how long process takes
overall_start_time = time.time()

# Clears some of the screen for asthetics
print("\n\n\n\n\n\n\n")


# Imports Excel report template to use
excel_report_template = pd.read_excel(EXCEL_TEMPLATE_PATH, 
                                 "Report",
                                 header=None)


# Runs through each project file in the AOI output public sharedrive folder
for project_file_name in os.listdir(AOI_PUBLIC_PATH):
    
    # This script only runs for HBCOSA lots, so if not that lot, then skip
    if "Window" not in project_file_name:
        continue
    
    # Creates file name
    project_file_path = os.path.join(AOI_PUBLIC_PATH, project_file_name)
    
    # Creates an excel sheet used from template to save/export as final report
    excel_final_report = excel_report_template.copy()
    
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
        for excel_file_name_index, excel_file_name in enumerate(os.listdir(slot_file_path) ):
            
            # Skips report excel sheet
            if "Report" in excel_file_name:
                continue
            
            # Creates Excel file path
            excel_file_path = os.path.join(slot_file_path, excel_file_name)
            
            # IF REPORT EXIST, SKIP. CAN DELETE AFTER TESTING
            if os.path.isfile(slot_file_path + "/" + project_file_name + "-Report.xlsx"):
                break
            
            # imports excel sheet data
            excel_import_data = pd.read_excel(excel_file_path, 
                                             sheet_name=excel_file_name.replace(".xlsx",""),
                                             header=None)
            
            # Copys wafer data
            wafer_data = excel_import_data.iloc[:16, 1:].copy()
            wafer_bin_count = excel_import_data.iloc[18:20, 11].copy()
            
            # Copying wafer data to final report
            excel_final_report.iloc[(0+17*excel_file_name_index):(16+17*excel_file_name_index), :16] = wafer_data
            # Excel slot name
            excel_final_report.iloc[(1+17*excel_file_name_index), 18] = excel_file_name.replace(".xlsx","")
            # Excel bin count
            excel_final_report.iloc[(4+17*excel_file_name_index):(6+17*excel_file_name_index), 18] = wafer_bin_count
        
        # Saves Excel final report
        if not os.path.isfile(slot_file_path + "/" + project_file_name + "-Report.xlsx"):
            print("Saving", slot_file_path + "/" + project_file_name + "-Report.xlsx")
            excel_final_report.to_excel(slot_file_path + "/" + project_file_name + "-Report_To_Copy.xlsx",
                                        header=None,
                                        index=None)



print("\nThis program is done!")

# Starting stopwatch to see how long process takes
print("Total Time: ")
overall_end_time = time.time()
time_lapsed = overall_end_time - overall_start_time
time_convert(time_lapsed)

