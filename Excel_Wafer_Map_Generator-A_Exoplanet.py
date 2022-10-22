# Import the necessary packages
import os
import sys
import shutil
import glob
import time
import numpy as np
import xlsxwriter
import re
import yaml
from math import sqrt


# User Parameters/Constants to Set
PREDICTED_DIR = "//mcrtp-file-01.mcusa.local/public/000-AOI_Tool_Output/"
STORED_WAFER_DATA = "C:/Users/ait.lab/.spyder-py3/Automated_AOI/Lot_Data/"
EXCEL_GENERATOR_TOGGLE = True
MAXIMUM_DEFECTS_TO_PASS_CLASS_1 = 0
MAXIMUM_DEFECTS_TO_PASS_CLASS_2 = 0
MAXIMUM_DEFECTS_TO_PASS_CLASS_3 = 1


def time_convert(sec):
    mins = sec // 60
    sec = sec % 60
    hours = mins // 60
    mins = mins % 60
    print("Time Lapsed = {0}h:{1}m:{2}s".format(int(hours), int(mins), round(sec) ) )



# MAIN():
# =============================================================================
# def main():
# Starting stopwatch to see how long process takes
start_time = time.time()

# Clears some of the screen for asthetics
print("\n\n\n")

max_defects_to_pass_class_list = []
max_defects_to_pass_class_list.append(MAXIMUM_DEFECTS_TO_PASS_CLASS_1)
max_defects_to_pass_class_list.append(MAXIMUM_DEFECTS_TO_PASS_CLASS_2)
max_defects_to_pass_class_list.append(MAXIMUM_DEFECTS_TO_PASS_CLASS_3)

# Cycles through each lot folder
for lotPathIndex, lotPath in enumerate(glob.glob(PREDICTED_DIR + "*") ):
    lotPathName = os.listdir(PREDICTED_DIR)[lotPathIndex]
    
    # Removes Thumbs.db in wafer map path if found
    if os.path.isfile(STORED_WAFER_DATA + "Thumbs.db"):
        os.remove(STORED_WAFER_DATA + "Thumbs.db")
        
    if "Exoplanet" not in lotPathName:
        continue
    
    # Checks to see if lot existing wafer map found in wafer map generator
    #  area for the current lotPath location
    shouldContinue = True
    for waferMapName in os.listdir(STORED_WAFER_DATA):
        if waferMapName in lotPathName:
            shouldContinue = False
            break
    if shouldContinue:
        continue
    
    # Imports correct dieNames and dieCoordinates data
    dieNames = np.load(STORED_WAFER_DATA + waferMapName + "/dieNames.npy")
    len_dieNames = len(dieNames)
    
    if not os.path.isfile(STORED_WAFER_DATA + waferMapName + "/config.yaml"):
        continue
    settings = yaml.safe_load( open(STORED_WAFER_DATA + waferMapName + "/config.yaml") )
    # Gets each appropriate settings
    classes_2 = settings['classes']
    good_class_index_2 = settings['good_class_index']
    excel_toggle_2 = settings['excel_toggle']
    num_excel_sheets_2 = settings['num_excel_sheets']
    
    if not excel_toggle_2:
        continue
    
    # Removes Thumbs.db in lot path if found
    if os.path.isfile(lotPath + "/Thumbs.db"):
        os.remove(lotPath + "/Thumbs.db")
    
    # Cycles through each slot folder within the lot folder
    for slot_name in os.listdir(lotPath):
        slot_path = os.path.join(lotPath, slot_name)
        
        # Checks if Excel sheet already exist, and only skips if selected not to
        if ( (os.path.isfile(slot_path + "/" + slot_name+ ".xlsx") 
        and os.path.isfile(lotPath + '/ZZZ-Excel_Sheets/' + slot_name + '.xlsx') )
        or slot_name == "ZZZ-Excel_Sheets"):
            continue
        
        print("Starting", slot_path)
        
        # Removes Thumbs.db in slot path if found
        if os.path.isfile(slot_path + "/Thumbs.db"):
            os.remove(slot_path + "/Thumbs.db")
        
        # Making list of bad die names
        badDieNames = []
        badDieBinNumbers = []
        bad_die_class_1_defect_count = []
        bad_die_class_2_defect_count = []
        bad_die_class_3_defect_count = []
        bad_die_classes_defect_count = []
        # For getting ro and col numbers and finding max
        row_list = []
        col_list = []
        bad_row_list = []
        bad_col_list = []
        bad_die_defect_num = []
        bad_defect_microns = []
        
        
        full_list = os.listdir(slot_path)
        for list_name_index, list_name in enumerate(full_list):
            if ".xlsx" in list_name or ".jpg" in list_name:
                del full_list[list_name_index]
        
        # Within each slot, cycle through each class
        for class_index, class_name in enumerate(full_list):
            class_path = os.path.join(slot_path, class_name)
            # Skips directory if first class (non-defect) folder or if it 
            # includes the wafer map with failing dies image (if this program 
            # already created one from a previous run)
            
            # WAS class_index == 0 NOW 1 FOR X-Display!! CHANGGGEE BACCKKKKK
            if (class_index == good_class_index_2 
            or "ZZ-" in class_name 
            or ".jpg" in class_name 
            or ".xlsx" in class_name
            or "Cropped" not in class_name
            ):
                continue
            # Removes Thumbs.db in class path if found
            if os.path.isfile(class_path + "/Thumbs.db"):
                os.remove(class_path + "/Thumbs.db")
            
            # Looks at die names in previously created wafer map and sees
            #  if this slot has the same die names within its defect folders.
            # If so, then create the new wafer map with red ovals in die 
            #  location within the wafer map image, and save this image.
            class_dies_list = os.listdir(class_path)
            previous_percent_index = 0
            for dieNameIndex, dieName in enumerate(dieNames):
                isBadDie = False
                
                if dieNameIndex == 0:
                    continue
                
                # Writes row and col numbers to find max
                row_list.append( int( re.findall(r'\d+', dieName)[0] ) )
                col_list.append( int( re.findall(r'\d+', dieName)[1] ) )
                
                # Shows progress in current slot
                if len_dieNames > 1000:
                    percent_index = round(dieNameIndex/len_dieNames*100)
                    
                    if percent_index != previous_percent_index:
                        sys.stdout.write('\033[2K\033[1G')
                        print("  ", class_name, "Progress:", 
                              str(round(dieNameIndex/len_dieNames*100) ) + "%",
                              end="\r"
                              )
                    
                    previous_percent_index = percent_index
                
                # # Checks if same die name already claimed as bad in previous class folder
                # if dieName in badDieNames:
                #     continue
                
                # Checks to see if current die name from general wafer 
                #  map die names is in any of the image names from current
                #  class folder
                called_bad_die = False
                for bad_die_name in class_dies_list:
                    if dieName in bad_die_name:
                        if called_bad_die:
                            if "1" in class_name:
                                bad_die_class_1_defect_count[-1] += 1
                            elif "2" in class_name:
                                bad_die_class_2_defect_count[-1] += 1
                            elif "3" in class_name:
                                bad_die_class_3_defect_count[-1] += 1
                        else:
                            badDieNames.append(dieName)
                            badDieBinNumbers.append(class_index)
                            
                            if "1" in class_name:
                                bad_die_class_1_defect_count.append(1)
                            elif "2" in class_name:
                                bad_die_class_2_defect_count.append(1)
                            elif "3" in class_name:
                                bad_die_class_3_defect_count.append(1)
                                
                            bad_row_list.append( int( re.findall(r'\d+', dieName)[0] ) )
                            bad_col_list.append( int( re.findall(r'\d+', dieName)[1] ) )
                            called_bad_die = True

                
                # if isBadDie:
                #     if len(dieNames) > 1000 and dieNameIndex % 1000 == 0:
                #         for list_index, image_name in enumerate(class_dies_list):
                #             if dieName in image_name:
                #                 del class_dies_list[:list_index]
                #                 break
                    
                #     continue
            
            # Makes new line so that next class progress status can show in terminal/shell
            print("")
                
        # Creates defect count for each class per each die
        bad_die_classes_defect_count.append(bad_die_class_1_defect_count)
        bad_die_classes_defect_count.append(bad_die_class_2_defect_count)
        bad_die_classes_defect_count.append(bad_die_class_3_defect_count)

        # XLS Section
        # -----------------------------------------------------------------------------
        if EXCEL_GENERATOR_TOGGLE:
            # XLSX PART
            max_row = max(row_list)
            max_col = max(col_list)
            
            # print("   Starting", slot_name, "Excel sheet results..")
            # Create a workbook and add a worksheet.
            workbook = xlsxwriter.Workbook(slot_path + '/' + slot_name + '.xlsx')
            
            
            worksheet_list = []
            # Checks to see how many worksheet list needed
            if num_excel_sheets_2 == 1:
                worksheet_list.append(workbook.add_worksheet(slot_name))
            else:
                for sheet_index in range(num_excel_sheets_2):
                    worksheet_list.append(workbook.add_worksheet(str(sheet_index)))
            
            # Chooses each font and background color for the Excel sheet
            font_color_list = ['black', 'black', 'black', 'white', 
                               'white', 'black', 'black', 'gray']
            bg_color_list = ['lime', 'magenta', 'magenta', 'red', 
                             'red', 'yellow', 'yellow', 'gray']
            
            
            
            # Chooses which font and background associated with each class
            bin_colors_list = []
            bin_bold_colors_list = []
            for class_index in range(len(classes_2)):
                bin_colors_list.append(workbook.add_format(
                    {'font_color': font_color_list[class_index],
                     'bg_color': bg_color_list[class_index],
                     'border': 4
                     }))
                bin_bold_colors_list.append(workbook.add_format(
                    {'bold': True,
                     'font_color': font_color_list[class_index],
                     'bg_color': bg_color_list[class_index]}))
                
            # For the "Not Tested Count" gray class
            bin_colors_list.append(workbook.add_format(
                {'font_color': font_color_list[-1],
                 'bg_color': bg_color_list[-1]}))
            bin_bold_colors_list.append(workbook.add_format(
                {'bold': True,
                 'font_color': font_color_list[-1],
                 'bg_color': bg_color_list[-1]}))
            
            # Finds how many rows and columns per Excel sheet
            row_per_sheet = int( max_row/( sqrt(num_excel_sheets_2) ) )
            col_per_sheet = int( max_col/( sqrt(num_excel_sheets_2) ) )
            
            # Iterates over each row_per_sheet x col_per_sheet dies 
            #  and defaults bin number to 8 - Untested
            for row in range(row_per_sheet*2):
                for col in range(col_per_sheet*2):
                    for worksheet in worksheet_list:
                        worksheet.write(row, col, "", bin_colors_list[-1])
            
            # Merges below wafer to say notch
            merge_format = workbook.add_format(
                {'align': 'center',
                 # 'bold': True,
                 'font_color': 'orange',
                 # 'border': 1,
                 'bg_color': bg_color_list[-1]
                 }
                )
            worksheet.merge_range(row_per_sheet*2, 0, row_per_sheet*2, 
                                  col_per_sheet*2-1, 'Notch', merge_format
                                  )
            
            
            # Combines all die names and bin numbers
            all_dieNames = badDieNames
            all_dieBinNumbers = badDieBinNumbers
            if len_dieNames > 1000:
                print("   Started making good bins..")
            
            full_list = glob.glob(slot_path + "/*")
            for list_name_index, list_name in enumerate(full_list):
                if ".xlsx" in list_name or ".jpg" in list_name:
                    del full_list[list_name_index]
                    
            list_items = os.listdir(full_list[good_class_index_2])
            
            # Checks to see which are good dies since previous scan in classes skipped good dies
            for dieNameIndex, dieName in enumerate(dieNames):
                should_skip = False
                if any(dieName in s for s in list_items):
                    row = int( re.findall(r'\d+', dieName)[0] )
                    col = int( re.findall(r'\d+', dieName)[1] )
                    
                    for bad_row_index, bad_row in enumerate(bad_row_list):
                        if row == bad_row and col == bad_col_list[bad_row_index]:
                            should_skip = True
                            break
                    if should_skip:
                        continue
                    all_dieNames.append(dieName)
                    all_dieBinNumbers.append(0) # USED TO BE good_class_index_2
            
                if len(dieNames) > 1000 and dieNameIndex % 1000 == 0:
                    for list_index, image_name in enumerate(class_dies_list):
                        if dieName in image_name:
                            del list_items[:list_index]
                            break
            
            if len_dieNames > 1000:
                print("   Started writing Excel sheet bin numbers..")
            
            # Writes all dies info in Excel
            for all_dieName_index, all_dieName in enumerate(all_dieNames):
                row = int( re.findall(r'\d+', all_dieName)[0] )
                col = int( re.findall(r'\d+', all_dieName)[1] )
                
                # Checks to see which background bin number to use
                bin_number = all_dieBinNumbers[all_dieName_index]
                background = bin_colors_list[bin_number]
                
                class_bin_number = bin_number
                
                # Figures out which class to be in as all_dieName_index 
                #  iterates through bad die names starting with classes in order
                if all_dieName_index < len(bad_die_classes_defect_count[0]):
                    class_index_to_use = all_dieName_index
                    bad_die_classes_defect_count_index = 0
                    
                elif (all_dieName_index >= len(bad_die_classes_defect_count[0])
                      and all_dieName_index < len(bad_die_classes_defect_count[1])
                      ):
                    class_index_to_use = all_dieName_index
                    class_index_to_use -= len(bad_die_classes_defect_count[0])
                    bad_die_classes_defect_count_index = 1
                    
                elif (all_dieName_index >= len(bad_die_classes_defect_count[1])
                      and all_dieName_index < len(bad_die_classes_defect_count[2])
                      ):
                    class_index_to_use = all_dieName_index
                    class_index_to_use -= len(bad_die_classes_defect_count[0])
                    class_index_to_use -= len(bad_die_classes_defect_count[1])
                    bad_die_classes_defect_count_index = 2
                    
                else:
                    class_index_to_use = all_dieName_index
                    bad_die_classes_defect_count_index = 99
                
                # Incase number of defects is below a desired threshold, 
                #  make it green - INEFFICIENT, PLEASE CHANGE IN FUTURE
                if (bin_number != good_class_index_2 
                    and bad_die_classes_defect_count_index <= 2
                    ):
                    if (bad_die_classes_defect_count[bad_die_classes_defect_count_index][class_index_to_use] 
                        <= max_defects_to_pass_class_list[bad_die_classes_defect_count_index]
                        ):
                        background = bin_colors_list[0]
                
                # If row or col is below 10 adds "0"s
                # --------------------------------------------------------------------
                
                # Row Section
                if row < 10:
                    row_string = "0" + str(row)
                else:
                    row_string = str(row)
                
                # Col Section
                if col < 10:
                    col_string = "0" + str(col)
                else:
                    col_string = str(col)
                
                # TEST SECTION
                # DELETE BELOW UNTIL ---- line
                os.listdir(slot_path + '/' + classes_2[class_bin_number])
                
                if os.path.isfile(slot_path + '/' 
                                  + classes_2[class_bin_number] + "Thumbs.db"):
                    os.remove(slot_path + '/' 
                              + classes_2[class_bin_number] + "Thumbs.db")
                    
                for image_name_jpg in os.listdir(slot_path + '/' + classes_2[class_bin_number]):
                    if 'Row_{}.Col_{}'.format(row_string, col_string) in image_name_jpg:
                        break
                        
                # --------------------------------------------------------------------
                if class_index == 1-1:
                    row_to_use = (row-1)*2+0
                    col_to_use = (col-1)*2+1
                elif class_index == 2-1:
                    row_to_use = (row-1)*2+0
                    col_to_use = (col-1)*2+0
                elif class_index == 3-1:
                    row_to_use = (row-1)*2+1
                    col_to_use = (col-1)*2+0
                else:
                    row_to_use = (row-1)*2+1
                    col_to_use = (col-1)*2+1
                
                blank_row = (row-1)*2+1
                blank_col = (col-1)*2+1
                
                if row <= row_per_sheet:
                    if col <= col_per_sheet:
                        # Hyperlink
                        if bin_number != good_class_index_2:
                            # Non Hyperlink - Just writes bins
                            if bad_die_classes_defect_count_index <= 2:
                                worksheet_list[0].write(row_to_use, col_to_use, 
                                                    bad_die_classes_defect_count[bad_die_classes_defect_count_index][class_index_to_use], 
                                                    background)
                        else:
                            # Non Hyperlink - Just writes bins
                            worksheet_list[0].write(row_to_use, col_to_use, 
                                                0, 
                                                background)
                        # Blank cell
                        worksheet_list[0].write(blank_row, blank_col, 
                                            "", 
                                            bin_colors_list[0])
                            
            
            
            # Selects appropriate "Not Tested Count" name
            not_tested_name = "8-Not_Tested-Count"
            
            # For each sheet, sets column width and zoom
            for worksheet_index, worksheet in enumerate(worksheet_list):
                
                # Sets the appropriate width for each column
                for row_index in range(12):
                    worksheet.set_row(row_index, height=23)
                worksheet.set_column(0, (col_per_sheet*2), width=round((3.3*max_row/max_col), 2) )
                
                # Sets zoom
                worksheet.set_zoom( 120 )
                
                
            
            workbook.close()
            
            os.makedirs(lotPath + '/ZZZ-Excel_Sheets/', exist_ok=True)
            shutil.copy( (slot_path + '/' + slot_name + '.xlsx'), (lotPath + '/ZZZ-Excel_Sheets/') )
        # -----------------------------------------------------------------------------

print("Done!")

# Stopping stopwatch to see how long process takes
end_time = time.time()
time_lapsed = end_time - start_time
time_convert(time_lapsed)


# if __name__ == "__main__":
#     main()


