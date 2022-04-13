# Import the necessary packages
import os
import shutil
import glob
import cv2
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


def time_convert(sec):
    mins = sec // 60
    sec = sec % 60
    hours = mins // 60
    mins = mins % 60
    print("Time Lapsed = {0}h:{1}m:{2}s".format(int(hours), int(mins), round(sec) ) )


def deleteDirContents(dir):
    # Deletes photos in path "dir"
    # # Used for deleting previous cropped photos from last run
    for f in os.listdir(dir):
        fullName = os.path.join(dir, f)
        shutil.rmtree(fullName)



# MAIN():
# =============================================================================
# Starting stopwatch to see how long process takes
start_time = time.time()

# Clears some of the screen for asthetics
print("\n\n\n\n\n\n\n\n\n\n\n\n\n")

# Cycles through each lot folder
for lotPathIndex, lotPath in enumerate(glob.glob(PREDICTED_DIR + "*") ):
    lotPathName = os.listdir(PREDICTED_DIR)[lotPathIndex]
    
    print("Starting", lotPathName)
    
    # Removes Thumbs.db in wafer map path if found
    if os.path.isfile(STORED_WAFER_DATA + "Thumbs.db"):
        os.remove(STORED_WAFER_DATA + "Thumbs.db")
    
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
    dieCoordinates = np.load(STORED_WAFER_DATA + waferMapName + "/Coordinates.npy")
    
    if not os.path.isfile(STORED_WAFER_DATA + waferMapName + "/config.yaml"):
        continue
    settings = yaml.safe_load( open(STORED_WAFER_DATA + waferMapName + "/config.yaml") )
    # Gets each appropriate settings
    model_path = settings['model_path']
    classes = settings['classes']
    good_class_index = settings['good_class_index']
    golden_image_toggle = settings['golden_image_toggle']
    split_match_CL = settings['split_match_CL']
    row_size = settings['row_size']
    col_size = settings['col_size']
    row_resize = settings['row_resize']
    col_resize = settings['col_resize']
    wafer_map_toggle = settings['wafer_map_toggle']
    excel_toggle = settings['excel_toggle']
    num_excel_sheets = settings['num_excel_sheets']
    grayscale_toggle = settings['grayscale_toggle']
    
    # Removes Thumbs.db in lot path if found
    if os.path.isfile(lotPath + "/Thumbs.db"):
        os.remove(lotPath + "/Thumbs.db")
    
    # Cycles through each slot folder within the lot folder
    for slot_name in os.listdir(lotPath):
        slot_path = os.path.join(lotPath, slot_name)
        # Creates wafer map
        waferMap = cv2.imread(STORED_WAFER_DATA + waferMapName +  "/Wafer_Map.jpg")
        
        # Checks if Excel sheet already exist, and only skips if selected not to
        if (os.path.isfile(slot_path + "/" + slot_name+ ".xlsx")
        or slot_name == "ZZZ-Excel_Sheets"):
            continue
        
        # Removes Thumbs.db in slot path if found
        if os.path.isfile(slot_path + "/Thumbs.db"):
            os.remove(slot_path + "/Thumbs.db")
        
        # Making list of bad die names
        badDieNames = []
        badDieBinNumbers = []
        # For getting ro and col numbers and finding max
        row_list = []
        col_list = []
        bad_row_list = []
        bad_col_list = []
        
        # Below is needed incase there is no bad dies
        isFirstImageRun = True
         
        # Within each slot, cycle through each class
        for class_index, class_name in enumerate(os.listdir(slot_path) ):
            class_path = os.path.join(slot_path, class_name)
            # Skips directory if first class (non-defect) folder or if it 
            # includes the wafer map with failing dies image (if this program 
            # already created one from a previous run)
            
            # WAS class_index == 0 NOW 1 FOR X-Display!! CHANGGGEE BACCKKKKK
            if (class_index == good_class_index 
            or "ZZ-" in class_name 
            or ".jpg" in class_name 
            or ".xlsx" in class_name):
                continue
            # Removes Thumbs.db in class path if found
            if os.path.isfile(class_path + "/Thumbs.db"):
                os.remove(class_path + "/Thumbs.db")
            
            # Looks at die names in previously created wafer map and sees
            #  if this slot has the same die names within its defect folders.
            # If so, then create the new wafer map with red ovals in die 
            #  location within the wafer map image, and save this image.
            list = os.listdir(class_path)
            (shown_progress_25, shown_progress_50, 
             shown_progress_75, shown_progress_100) = False, False, False, False
            for dieNameIndex, dieName in enumerate(dieNames):
                isBadDie = False
                
                if dieNameIndex == 0:
                    continue
                
                # Writes row and col numbers to find max
                # map(int, re.findall(r'\d+', string1))
                # import re ??? EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE
                row_list.append( int( re.findall(r'\d+', dieName)[0] ) )
                col_list.append( int( re.findall(r'\d+', dieName)[1] ) )
                # row_list.append( int(dieName[2:5]) )
                # col_list.append( int(dieName[-3:]) )
                
                # Shows progress in current slot
                if len_dieNames > 1000:
                    if (round(dieNameIndex/len_dieNames, 2) == 0.25
                    and shown_progress_25 == False):
                        print("  ", class_name, "Progress:", 
                              str(round(dieNameIndex/len_dieNames*100) ) + "%")
                        shown_progress_25 = True
                    if (round(dieNameIndex/len_dieNames, 2) == 0.50
                    and shown_progress_50 == False):
                        print("  ", class_name, "Progress:", 
                              str(round(dieNameIndex/len_dieNames*100) ) + "%")
                        shown_progress_50 = True
                    if (round(dieNameIndex/len_dieNames, 2) == 0.75
                    and shown_progress_75 == False):
                        print("  ", class_name, "Progress:", 
                              str(round(dieNameIndex/len_dieNames*100) ) + "%")
                        shown_progress_75 = True
                    if (round(dieNameIndex/len_dieNames, 2) == 1.00
                    and shown_progress_100 == False):
                        print("  ", class_name, "Progress:", 
                              str(round(dieNameIndex/len_dieNames*100) ) + "%")
                        shown_progress_100 = True
                
                # Checks if same die name already claimed as bad in previous class folder
                if dieName in badDieNames:
                    continue
                
                # Checks to see if current die name from general wafer 
                #  map die names is in any of the image names from current
                #  class folder
                if any(dieName in s for s in list):
                    isBadDie = True
                    badDieNames.append(dieName)
                    badDieBinNumbers.append(class_index)
                    bad_row_list.append( int( re.findall(r'\d+', dieName)[0] ) )
                    bad_col_list.append( int( re.findall(r'\d+', dieName)[1] ) )
                else:
                    isBadDie = False

                
                # Prevents green circles from being drawn
                isFirstImageRun = False
                
                if isBadDie:
                    if len(dieNames) > 1000 and dieNameIndex % 1000 == 0:
                        for list_index, image_name in enumerate(list):
                            if dieName in image_name:
                                del list[:list_index]
                                break
                    
                    continue
                


        # XLS Section
        # -----------------------------------------------------------------------------
        if EXCEL_GENERATOR_TOGGLE:
            # XLSX PART
            max_row = max(row_list)
            max_col = max(col_list)
            
            print("   Starting Excel sheet results..")
            # Create a workbook and add a worksheet.
            workbook = xlsxwriter.Workbook(slot_path + '/' + slot_name + '.xlsx')
            # Checks to see how many worksheet list needed
            if num_excel_sheets == 1:
                worksheet_TL = workbook.add_worksheet(slot_name)
                
                worksheet_list = [worksheet_TL]
            else:
                worksheet_TL = workbook.add_worksheet("TL")
                worksheet_TR = workbook.add_worksheet("TR")
                worksheet_BL = workbook.add_worksheet("BL")
                worksheet_BR = workbook.add_worksheet("BR")
            
                worksheet_list = [worksheet_TL, worksheet_TR, worksheet_BL, worksheet_BR]
            
            # Add a bold format to use to highlight cells.
            bold = workbook.add_format({'bold': True})
            # bin 0 - black
            bin0_background = workbook.add_format({'font_color': 'white',
                                                   'bg_color': 'black'})
            bin0_bold_background = workbook.add_format({'bold': True,
                                                        'font_color': 'white',
                                                        'bg_color': 'black'})
            # bin 1 - light green
            bin1_background = workbook.add_format({'bg_color': 'lime'})
            bin1_bold_background = workbook.add_format({'bold': True, 
                                                        'bg_color': 'lime'})
            # bin 2 - red
            bin2_background = workbook.add_format({'font_color': 'white',
                                                   'bg_color': 'red'})
            bin2_bold_background = workbook.add_format({'bold': True, 
                                                        'font_color': 'white',
                                                        'bg_color': 'red'})
            # bin 3 - dark green or green
            bin3_background = workbook.add_format({'font_color': 'white',
                                                   'bg_color': 'green'})
            bin3_bold_background = workbook.add_format({'bold': True, 
                                                        'font_color': 'white',
                                                        'bg_color': 'green'})
            # bin 4 - yellow
            bin4_background = workbook.add_format({'bg_color': 'yellow'})
            bin4_bold_background = workbook.add_format({'bold': True, 
                                                        'bg_color': 'yellow'})
            # bin 5 - blue
            bin5_background = workbook.add_format({'font_color': 'white',
                                                   'bg_color': 'blue'})
            bin5_bold_background = workbook.add_format({'bold': True, 
                                                        'font_color': 'white',
                                                        'bg_color': 'blue'})
            # bin 6 - pink
            bin6_background = workbook.add_format({'bg_color': 'magenta'})
            bin6_bold_background = workbook.add_format({'bold': True, 
                                                        'bg_color': 'magenta'})
            # bin 7 - cyan
            bin7_background = workbook.add_format({'bg_color': 'cyan'})
            bin7_bold_background = workbook.add_format({'bold': True, 
                                                        'bg_color': 'cyan'})
            # bin 8 - gray
            bin8_background = workbook.add_format({'font_color': 'white',
                                                   'bg_color': 'gray'})
            bin8_bold_background = workbook.add_format({'bold': True, 
                                                        'font_color': 'white',
                                                        'bg_color': 'gray'})
            
            # Finds how many rows and columns per Excel sheet
            row_per_sheet = int( max_row/( sqrt(num_excel_sheets) ) )
            col_per_sheet = int( max_col/( sqrt(num_excel_sheets) ) )
            
            # Iterates over each row_per_sheet x col_per_sheet dies 
            #  and defaults bin number to 8 - Untested
            if waferMapName == "HBCOSA":
                for row in range(row_per_sheet+1):
                    for col in range(col_per_sheet+1):
                        for worksheet in worksheet_list:
                            worksheet.write(row, col, 0, bin8_background)
            else:
                for row in range(row_per_sheet):
                    for col in range(col_per_sheet):
                        for worksheet in worksheet_list:
                            worksheet.write(row, col, 8, bin8_background)
            
            
            # Combines all die names and bin numbers
            all_dieNames = badDieNames
            all_dieBinNumbers = badDieBinNumbers
            print("   Started making good bins..")
            list = os.listdir(glob.glob(slot_path + "/*")[good_class_index])
            
            # Checks to see which are good dies since previous scan in classes skipped good dies
            for dieNameIndex, dieName in enumerate(dieNames):
                should_skip = False
                if any(dieName in s for s in list):
                    row = int( re.findall(r'\d+', dieName)[0] )
                    col = int( re.findall(r'\d+', dieName)[1] )
                    
                    for bad_row_index, bad_row in enumerate(bad_row_list):
                        if row == bad_row and col == bad_col_list[bad_row_index]:
                            should_skip = True
                            break
                    if should_skip:
                        continue
                    all_dieNames.append(dieName)
                    all_dieBinNumbers.append(good_class_index)
            
                if len(dieNames) > 1000 and dieNameIndex % 1000 == 0:
                    for list_index, image_name in enumerate(list):
                        if dieName in image_name:
                            del list[:list_index]
                            break
            
            
            print("   Started writing Excel sheet bin numbers..")
            # Writes all dies info in Excel
            for all_dieName_index, all_dieName in enumerate(all_dieNames):
                # row = int(all_dieName[2:5])
                # col = int(all_dieName[-3:])
                # Replace EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE
                row = int( re.findall(r'\d+', all_dieName)[0] )
                col = int( re.findall(r'\d+', all_dieName)[1] )
                
                # Checks to see which background bin number to use
                if sqrt(num_excel_sheets) == 2:
                    # Sees if it is X-Display LED dies
                    if all_dieBinNumbers[all_dieName_index] == 0:
                        background = bin0_background
                    elif all_dieBinNumbers[all_dieName_index] == 1:
                        background = bin1_background
                    elif all_dieBinNumbers[all_dieName_index] == 2:
                        background = bin2_background
                    elif all_dieBinNumbers[all_dieName_index] == 3:
                        background = bin3_background
                    elif all_dieBinNumbers[all_dieName_index] == 4:
                        background = bin4_background
                    elif all_dieBinNumbers[all_dieName_index] == 5:
                        background = bin5_background
                    elif all_dieBinNumbers[all_dieName_index] == 6:
                        background = bin6_background
                    elif all_dieBinNumbers[all_dieName_index] == 7:
                        background = bin7_background
                elif sqrt(num_excel_sheets) == 1:
                    # Sees if it is non-LED dies
                    if all_dieBinNumbers[all_dieName_index] == good_class_index:
                        background = bin1_background
                    elif all_dieBinNumbers[all_dieName_index] == 1:
                        background = bin2_background
                    elif all_dieBinNumbers[all_dieName_index] == 2:
                        background = bin6_background
                
                if waferMapName == "HBCOSA":
                    bin_number = all_dieBinNumbers[all_dieName_index] + 1
                else:
                    bin_number = all_dieBinNumbers[all_dieName_index]
                
                if row <= row_per_sheet:
                    if col <= col_per_sheet:
                        worksheet_TL.write(row-1, col-1, 
                                           bin_number, 
                                           background)
                        
                    else:
                        worksheet_TR.write(row-1, col-1-col_per_sheet, 
                                           bin_number,
                                           background)
                else:
                    if col <= col_per_sheet:
                        worksheet_BL.write(row-1-row_per_sheet, col-1, 
                                           bin_number,
                                           background)
                    else:
                        worksheet_BR.write(row-1-row_per_sheet, col-1-col_per_sheet, 
                                           bin_number,
                                           background)
            
            
            # Bin count for each quadrant
            tl_bin0 = 0
            tl_bin1 = 0
            tl_bin2 = 0
            tl_bin3 = 0
            tl_bin4 = 0
            tl_bin5 = 0
            tl_bin6 = 0
            tl_bin7 = 0
            tl_bin8 = 0
            
            tr_bin0 = 0
            tr_bin1 = 0
            tr_bin2 = 0
            tr_bin3 = 0
            tr_bin4 = 0
            tr_bin5 = 0
            tr_bin6 = 0
            tr_bin7 = 0
            tr_bin8 = 0
            
            bl_bin0 = 0
            bl_bin1 = 0
            bl_bin2 = 0
            bl_bin3 = 0
            bl_bin4 = 0
            bl_bin5 = 0
            bl_bin6 = 0
            bl_bin7 = 0
            bl_bin8 = 0
            
            br_bin0 = 0
            br_bin1 = 0
            br_bin2 = 0
            br_bin3 = 0
            br_bin4 = 0
            br_bin5 = 0
            br_bin6 = 0
            br_bin7 = 0
            br_bin8 = 0
            
            # Counts how many bins in each worksheet
            for row in range(row_per_sheet):
                for col in range(col_per_sheet):
                    if sqrt(num_excel_sheets) == 2:
                        tl_bin_number = worksheet_TL.table[row][col].number
                        if tl_bin_number == 0:
                            tl_bin0 += 1
                        elif tl_bin_number == 1:
                            tl_bin1 += 1
                        elif tl_bin_number == 2:
                            tl_bin2 += 1
                        elif tl_bin_number == 3:
                            tl_bin3 += 1
                        elif tl_bin_number == 4:
                            tl_bin4 += 1
                        elif tl_bin_number == 5:
                            tl_bin5 += 1
                        elif tl_bin_number == 6:
                            tl_bin6 += 1
                        elif tl_bin_number == 7:
                            tl_bin7 += 1
                        
                        tr_bin_number = worksheet_TR.table[row][col].number
                        if tr_bin_number == 0:
                            tr_bin0 += 1
                        elif tr_bin_number == 1:
                            tr_bin1 += 1
                        elif tr_bin_number == 2:
                            tr_bin2 += 1
                        elif tr_bin_number == 3:
                            tr_bin3 += 1
                        elif tr_bin_number == 4:
                            tr_bin4 += 1
                        elif tr_bin_number == 5:
                            tr_bin5 += 1
                        elif tr_bin_number == 6:
                            tr_bin6 += 1
                        elif tr_bin_number == 7:
                            tr_bin7 += 1
                        
                        bl_bin_number = worksheet_BL.table[row][col].number
                        if bl_bin_number == 0:
                            bl_bin0 += 1
                        elif bl_bin_number == 1:
                            bl_bin1 += 1
                        elif bl_bin_number == 2:
                            bl_bin2 += 1
                        elif bl_bin_number == 3:
                            bl_bin3 += 1
                        elif bl_bin_number == 4:
                            bl_bin4 += 1
                        elif bl_bin_number == 5:
                            bl_bin5 += 1
                        elif bl_bin_number == 6:
                            bl_bin6 += 1
                        elif bl_bin_number == 7:
                            bl_bin7 += 1
                        
                        br_bin_number = worksheet_BR.table[row][col].number
                        if br_bin_number == 0:
                            br_bin0 += 1
                        elif br_bin_number == 1:
                            br_bin1 += 1
                        elif br_bin_number == 2:
                            br_bin2 += 1
                        elif br_bin_number == 3:
                            br_bin3 += 1
                        elif br_bin_number == 4:
                            br_bin4 += 1
                        elif br_bin_number == 5:
                            br_bin5 += 1
                        elif br_bin_number == 6:
                            br_bin6 += 1
                        elif br_bin_number == 7:
                            br_bin7 += 1
                    
                    elif sqrt(num_excel_sheets) == 1:
                        tl_bin_number = worksheet_TL.table[row][col].number
                        if tl_bin_number == 0:
                            tl_bin0 += 1
                        elif tl_bin_number == 1:
                            tl_bin1 += 1
                        elif tl_bin_number == 2:
                            tl_bin2 += 1
                        elif tl_bin_number == 3:
                            tl_bin3 += 1
                        elif tl_bin_number == 4:
                            tl_bin4 += 1
                        elif tl_bin_number == 5:
                            tl_bin5 += 1
                        elif tl_bin_number == 6:
                            tl_bin6 += 1
                        elif tl_bin_number == 7:
                            tl_bin7 += 1
            
            
            # TL 
            # -----------------------------------------------------------
            # Dividing this to LED and non-LED dies
            if sqrt(num_excel_sheets) == 2:
                worksheet_list[0].write(202, 0, classes[0], 
                                     bin0_bold_background)
                worksheet_list[0].write(202, 1, "", bin0_background)
                worksheet_list[0].write(202, 2, "", bin0_background)
                worksheet_list[0].write(202, 3, "", bin0_background)
                worksheet_list[0].write(202, 4, "", bin0_background)
                worksheet_list[0].write(202, 5, "", bin0_background)
                worksheet_list[0].write(202, 6, "", bin0_background)
                worksheet_list[0].write(202, 7, "", bin0_background)
                worksheet_list[0].write(202, 8, "", bin0_background)
                worksheet_list[0].write(202, 9, "", bin0_background)
                worksheet_list[0].write(202, 10, "", bin0_background)
                worksheet_list[0].write(202, 11, tl_bin0,
                                     bin0_bold_background)
                
                worksheet_list[0].write(203, 0, classes[1], 
                                     bin1_bold_background)
                worksheet_list[0].write(203, 1, "", bin1_background)
                worksheet_list[0].write(203, 2, "", bin1_background)
                worksheet_list[0].write(203, 3, "", bin1_background)
                worksheet_list[0].write(203, 4, "", bin1_background)
                worksheet_list[0].write(203, 5, "", bin1_background)
                worksheet_list[0].write(203, 6, "", bin1_background)
                worksheet_list[0].write(203, 7, "", bin1_background)
                worksheet_list[0].write(203, 8, "", bin1_background)
                worksheet_list[0].write(203, 9, "", bin1_background)
                worksheet_list[0].write(203, 10, "", bin1_background)
                worksheet_list[0].write(203, 11, tl_bin1,
                                     bin1_bold_background)
                
                worksheet_list[0].write(204, 0, classes[2], 
                                     bin2_bold_background)
                worksheet_list[0].write(204, 1, "", bin2_background)
                worksheet_list[0].write(204, 2, "", bin2_background)
                worksheet_list[0].write(204, 3, "", bin2_background)
                worksheet_list[0].write(204, 4, "", bin2_background)
                worksheet_list[0].write(204, 5, "", bin2_background)
                worksheet_list[0].write(204, 6, "", bin2_background)
                worksheet_list[0].write(204, 7, "", bin2_background)
                worksheet_list[0].write(204, 8, "", bin2_background)
                worksheet_list[0].write(204, 9, "", bin2_background)
                worksheet_list[0].write(204, 10, "", bin2_background)
                worksheet_list[0].write(204, 11, tl_bin2,
                                     bin2_bold_background)
                
                worksheet_list[0].write(205, 0, classes[3], 
                                     bin3_bold_background)
                worksheet_list[0].write(205, 1, "", bin3_background)
                worksheet_list[0].write(205, 2, "", bin3_background)
                worksheet_list[0].write(205, 3, "", bin3_background)
                worksheet_list[0].write(205, 4, "", bin3_background)
                worksheet_list[0].write(205, 5, "", bin3_background)
                worksheet_list[0].write(205, 6, "", bin3_background)
                worksheet_list[0].write(205, 7, "", bin3_background)
                worksheet_list[0].write(205, 8, "", bin3_background)
                worksheet_list[0].write(205, 9, "", bin3_background)
                worksheet_list[0].write(205, 10, "", bin3_background)
                worksheet_list[0].write(205, 11, tl_bin3,
                                     bin3_bold_background)
                
                worksheet_list[0].write(206, 0, classes[4], 
                                     bin4_bold_background)
                worksheet_list[0].write(206, 1, "", bin4_background)
                worksheet_list[0].write(206, 2, "", bin4_background)
                worksheet_list[0].write(206, 3, "", bin4_background)
                worksheet_list[0].write(206, 4, "", bin4_background)
                worksheet_list[0].write(206, 5, "", bin4_background)
                worksheet_list[0].write(206, 6, "", bin4_background)
                worksheet_list[0].write(206, 7, "", bin4_background)
                worksheet_list[0].write(206, 8, "", bin4_background)
                worksheet_list[0].write(206, 9, "", bin4_background)
                worksheet_list[0].write(206, 10, "", bin4_background)
                worksheet_list[0].write(206, 11, tl_bin4,
                                     bin4_bold_background)
                
                worksheet_list[0].write(207, 0, classes[5], 
                                     bin5_bold_background)
                worksheet_list[0].write(207, 1, "", bin5_background)
                worksheet_list[0].write(207, 2, "", bin5_background)
                worksheet_list[0].write(207, 3, "", bin5_background)
                worksheet_list[0].write(207, 4, "", bin5_background)
                worksheet_list[0].write(207, 5, "", bin5_background)
                worksheet_list[0].write(207, 6, "", bin5_background)
                worksheet_list[0].write(207, 7, "", bin5_background)
                worksheet_list[0].write(207, 8, "", bin5_background)
                worksheet_list[0].write(207, 9, "", bin5_background)
                worksheet_list[0].write(207, 10, "", bin5_background)
                worksheet_list[0].write(207, 11, tl_bin5,
                                     bin5_bold_background)
                
                worksheet_list[0].write(208, 0, classes[6], 
                                     bin6_bold_background)
                worksheet_list[0].write(208, 1, "", bin6_background)
                worksheet_list[0].write(208, 2, "", bin6_background)
                worksheet_list[0].write(208, 3, "", bin6_background)
                worksheet_list[0].write(208, 4, "", bin6_background)
                worksheet_list[0].write(208, 5, "", bin6_background)
                worksheet_list[0].write(208, 6, "", bin6_background)
                worksheet_list[0].write(208, 7, "", bin6_background)
                worksheet_list[0].write(208, 8, "", bin6_background)
                worksheet_list[0].write(208, 9, "", bin6_background)
                worksheet_list[0].write(208, 10, "", bin6_background)
                worksheet_list[0].write(208, 11, tl_bin6,
                                     bin6_bold_background)
                
                worksheet_list[0].write(209, 0, classes[7], 
                                     bin7_bold_background)
                worksheet_list[0].write(209, 1, "", bin7_background)
                worksheet_list[0].write(209, 2, "", bin7_background)
                worksheet_list[0].write(209, 3, "", bin7_background)
                worksheet_list[0].write(209, 4, "", bin7_background)
                worksheet_list[0].write(209, 5, "", bin7_background)
                worksheet_list[0].write(209, 6, "", bin7_background)
                worksheet_list[0].write(209, 7, "", bin7_background)
                worksheet_list[0].write(209, 8, "", bin7_background)
                worksheet_list[0].write(209, 9, "", bin7_background)
                worksheet_list[0].write(209, 10, "", bin7_background)
                worksheet_list[0].write(209, 11, tl_bin7,
                                     bin7_bold_background)
                
                worksheet_list[0].write(210, 0, "8 - Not_Tested-Count", 
                                     bin8_bold_background)
                worksheet_list[0].write(210, 1, "", bin8_background)
                worksheet_list[0].write(210, 2, "", bin8_background)
                worksheet_list[0].write(210, 3, "", bin8_background)
                worksheet_list[0].write(210, 4, "", bin8_background)
                worksheet_list[0].write(210, 5, "", bin8_background)
                worksheet_list[0].write(210, 6, "", bin8_background)
                worksheet_list[0].write(210, 7, "", bin8_background)
                worksheet_list[0].write(210, 8, "", bin8_background)
                worksheet_list[0].write(210, 9, "", bin8_background)
                worksheet_list[0].write(210, 10, "", bin8_background)
                worksheet_list[0].write(210, 11, "="+str((len(dieNames)-1)/num_excel_sheets)+"-sum(L203:L210)",
                                     bin8_bold_background)
            elif sqrt(num_excel_sheets) == 1:
                if waferMapName == "HBCOSA":
                    tl_bin0 = tl_bin1
                    tl_bin1 = tl_bin2
                    tl_bin2 = tl_bin3
                
                worksheet_list[0].write(max_row + 2, 0, classes[0], 
                                     bin1_bold_background)
                worksheet_list[0].write(max_row + 2, 1, "", bin1_background)
                worksheet_list[0].write(max_row + 2, 2, "", bin1_background)
                worksheet_list[0].write(max_row + 2, 3, "", bin1_background)
                worksheet_list[0].write(max_row + 2, 4, "", bin1_background)
                worksheet_list[0].write(max_row + 2, 5, "", bin1_background)
                worksheet_list[0].write(max_row + 2, 6, "", bin1_background)
                worksheet_list[0].write(max_row + 2, 7, "", bin1_background)
                worksheet_list[0].write(max_row + 2, 8, "", bin1_background)
                worksheet_list[0].write(max_row + 2, 9, "", bin1_background)
                worksheet_list[0].write(max_row + 2, 10, "", bin1_background)
                worksheet_list[0].write(max_row + 2, 11, tl_bin0,
                                     bin1_bold_background)
                
                worksheet_list[0].write(max_row + 3, 0, classes[1], 
                                     bin2_bold_background)
                worksheet_list[0].write(max_row + 3, 1, "", bin2_background)
                worksheet_list[0].write(max_row + 3, 2, "", bin2_background)
                worksheet_list[0].write(max_row + 3, 3, "", bin2_background)
                worksheet_list[0].write(max_row + 3, 4, "", bin2_background)
                worksheet_list[0].write(max_row + 3, 5, "", bin2_background)
                worksheet_list[0].write(max_row + 3, 6, "", bin2_background)
                worksheet_list[0].write(max_row + 3, 7, "", bin2_background)
                worksheet_list[0].write(max_row + 3, 8, "", bin2_background)
                worksheet_list[0].write(max_row + 3, 9, "", bin2_background)
                worksheet_list[0].write(max_row + 3, 10, "", bin2_background)
                worksheet_list[0].write(max_row + 3, 11, tl_bin1,
                                     bin2_bold_background)
                
                # worksheet_name.set_column(0, 0, width=len(classes[7]))
                worksheet_list[0].set_column(0, (col_per_sheet), width=2)
                worksheet_list[0].set_column(11, 11, width=3)
            
                
                if len(classes) > 2:
                    worksheet_list[0].write(max_row + 4, 0, classes[2], 
                                         bin6_bold_background)
                    worksheet_list[0].write(max_row + 4, 1, "", bin6_background)
                    worksheet_list[0].write(max_row + 4, 2, "", bin6_background)
                    worksheet_list[0].write(max_row + 4, 3, "", bin6_background)
                    worksheet_list[0].write(max_row + 4, 4, "", bin6_background)
                    worksheet_list[0].write(max_row + 4, 5, "", bin6_background)
                    worksheet_list[0].write(max_row + 4, 6, "", bin6_background)
                    worksheet_list[0].write(max_row + 4, 7, "", bin6_background)
                    worksheet_list[0].write(max_row + 4, 8, "", bin6_background)
                    worksheet_list[0].write(max_row + 4, 9, "", bin6_background)
                    worksheet_list[0].write(max_row + 4, 10, "", bin6_background)
                    worksheet_list[0].write(max_row + 4, 11, tl_bin2,
                                         bin6_bold_background)
                
                if waferMapName == "HBCOSA":
                    worksheet_list[0].write(max_row + 5, 0, "0-Not_Tested-Count", 
                                         bin8_bold_background)
                else:
                    worksheet_list[0].write(max_row + 5, 0, "8-Not_Tested-Count", 
                                         bin8_bold_background)
                worksheet_list[0].write(max_row + 5, 1, "", bin8_background)
                worksheet_list[0].write(max_row + 5, 2, "", bin8_background)
                worksheet_list[0].write(max_row + 5, 3, "", bin8_background)
                worksheet_list[0].write(max_row + 5, 4, "", bin8_background)
                worksheet_list[0].write(max_row + 5, 5, "", bin8_background)
                worksheet_list[0].write(max_row + 5, 6, "", bin8_background)
                worksheet_list[0].write(max_row + 5, 7, "", bin8_background)
                worksheet_list[0].write(max_row + 5, 8, "", bin8_background)
                worksheet_list[0].write(max_row + 5, 9, "", bin8_background)
                worksheet_list[0].write(max_row + 5, 10, "", bin8_background)
                worksheet_list[0].write(max_row + 5, 11, 
                                    "="+str((len(dieNames)-1)/num_excel_sheets)
                                    +"-sum(L{}:L{})".format((max_row + 2 + 1),
                                                            (max_row + 4 + 1)),
                                     bin8_bold_background)
            
            # -----------------------------------------------------------
            
            
            if sqrt(num_excel_sheets) == 2:
                # TR
                # -----------------------------------------------------------
                worksheet_list[1].write(202, 0, classes[0], 
                                     bin0_bold_background)
                worksheet_list[1].write(202, 1, "", bin0_background)
                worksheet_list[1].write(202, 2, "", bin0_background)
                worksheet_list[1].write(202, 3, "", bin0_background)
                worksheet_list[1].write(202, 4, "", bin0_background)
                worksheet_list[1].write(202, 5, "", bin0_background)
                worksheet_list[1].write(202, 6, "", bin0_background)
                worksheet_list[1].write(202, 7, "", bin0_background)
                worksheet_list[1].write(202, 8, "", bin0_background)
                worksheet_list[1].write(202, 9, "", bin0_background)
                worksheet_list[1].write(202, 10, "", bin0_background)
                worksheet_list[1].write(202, 11, tr_bin0,
                                     bin0_bold_background)
                
                worksheet_list[1].write(203, 0, classes[1], 
                                     bin1_bold_background)
                worksheet_list[1].write(203, 1, "", bin1_background)
                worksheet_list[1].write(203, 2, "", bin1_background)
                worksheet_list[1].write(203, 3, "", bin1_background)
                worksheet_list[1].write(203, 4, "", bin1_background)
                worksheet_list[1].write(203, 5, "", bin1_background)
                worksheet_list[1].write(203, 6, "", bin1_background)
                worksheet_list[1].write(203, 7, "", bin1_background)
                worksheet_list[1].write(203, 8, "", bin1_background)
                worksheet_list[1].write(203, 9, "", bin1_background)
                worksheet_list[1].write(203, 10, "", bin1_background)
                worksheet_list[1].write(203, 11, tr_bin1,
                                     bin1_bold_background)
                
                worksheet_list[1].write(204, 0, classes[2], 
                                     bin2_bold_background)
                worksheet_list[1].write(204, 1, "", bin2_background)
                worksheet_list[1].write(204, 2, "", bin2_background)
                worksheet_list[1].write(204, 3, "", bin2_background)
                worksheet_list[1].write(204, 4, "", bin2_background)
                worksheet_list[1].write(204, 5, "", bin2_background)
                worksheet_list[1].write(204, 6, "", bin2_background)
                worksheet_list[1].write(204, 7, "", bin2_background)
                worksheet_list[1].write(204, 8, "", bin2_background)
                worksheet_list[1].write(204, 9, "", bin2_background)
                worksheet_list[1].write(204, 10, "", bin2_background)
                worksheet_list[1].write(204, 11, tr_bin2,
                                     bin2_bold_background)
                
                worksheet_list[1].write(205, 0, classes[3], 
                                     bin3_bold_background)
                worksheet_list[1].write(205, 1, "", bin3_background)
                worksheet_list[1].write(205, 2, "", bin3_background)
                worksheet_list[1].write(205, 3, "", bin3_background)
                worksheet_list[1].write(205, 4, "", bin3_background)
                worksheet_list[1].write(205, 5, "", bin3_background)
                worksheet_list[1].write(205, 6, "", bin3_background)
                worksheet_list[1].write(205, 7, "", bin3_background)
                worksheet_list[1].write(205, 8, "", bin3_background)
                worksheet_list[1].write(205, 9, "", bin3_background)
                worksheet_list[1].write(205, 10, "", bin3_background)
                worksheet_list[1].write(205, 11, tr_bin3,
                                     bin3_bold_background)
                
                worksheet_list[1].write(206, 0, classes[4], 
                                     bin4_bold_background)
                worksheet_list[1].write(206, 1, "", bin4_background)
                worksheet_list[1].write(206, 2, "", bin4_background)
                worksheet_list[1].write(206, 3, "", bin4_background)
                worksheet_list[1].write(206, 4, "", bin4_background)
                worksheet_list[1].write(206, 5, "", bin4_background)
                worksheet_list[1].write(206, 6, "", bin4_background)
                worksheet_list[1].write(206, 7, "", bin4_background)
                worksheet_list[1].write(206, 8, "", bin4_background)
                worksheet_list[1].write(206, 9, "", bin4_background)
                worksheet_list[1].write(206, 10, "", bin4_background)
                worksheet_list[1].write(206, 11, tr_bin4,
                                     bin4_bold_background)
                
                worksheet_list[1].write(207, 0, classes[5], 
                                     bin5_bold_background)
                worksheet_list[1].write(207, 1, "", bin5_background)
                worksheet_list[1].write(207, 2, "", bin5_background)
                worksheet_list[1].write(207, 3, "", bin5_background)
                worksheet_list[1].write(207, 4, "", bin5_background)
                worksheet_list[1].write(207, 5, "", bin5_background)
                worksheet_list[1].write(207, 6, "", bin5_background)
                worksheet_list[1].write(207, 7, "", bin5_background)
                worksheet_list[1].write(207, 8, "", bin5_background)
                worksheet_list[1].write(207, 9, "", bin5_background)
                worksheet_list[1].write(207, 10, "", bin5_background)
                worksheet_list[1].write(207, 11, tr_bin5,
                                     bin5_bold_background)
                
                worksheet_list[1].write(208, 0, classes[6], 
                                     bin6_bold_background)
                worksheet_list[1].write(208, 1, "", bin6_background)
                worksheet_list[1].write(208, 2, "", bin6_background)
                worksheet_list[1].write(208, 3, "", bin6_background)
                worksheet_list[1].write(208, 4, "", bin6_background)
                worksheet_list[1].write(208, 5, "", bin6_background)
                worksheet_list[1].write(208, 6, "", bin6_background)
                worksheet_list[1].write(208, 7, "", bin6_background)
                worksheet_list[1].write(208, 8, "", bin6_background)
                worksheet_list[1].write(208, 9, "", bin6_background)
                worksheet_list[1].write(208, 10, "", bin6_background)
                worksheet_list[1].write(208, 11, tr_bin6,
                                     bin6_bold_background)
                
                worksheet_list[1].write(209, 0, classes[7], 
                                     bin7_bold_background)
                worksheet_list[1].write(209, 1, "", bin7_background)
                worksheet_list[1].write(209, 2, "", bin7_background)
                worksheet_list[1].write(209, 3, "", bin7_background)
                worksheet_list[1].write(209, 4, "", bin7_background)
                worksheet_list[1].write(209, 5, "", bin7_background)
                worksheet_list[1].write(209, 6, "", bin7_background)
                worksheet_list[1].write(209, 7, "", bin7_background)
                worksheet_list[1].write(209, 8, "", bin7_background)
                worksheet_list[1].write(209, 9, "", bin7_background)
                worksheet_list[1].write(209, 10, "", bin7_background)
                worksheet_list[1].write(209, 11, tr_bin7,
                                     bin7_bold_background)
                
                worksheet_list[1].write(210, 0, "8 - Not_Tested-Count", 
                                     bin8_bold_background)
                worksheet_list[1].write(210, 1, "", bin8_background)
                worksheet_list[1].write(210, 2, "", bin8_background)
                worksheet_list[1].write(210, 3, "", bin8_background)
                worksheet_list[1].write(210, 4, "", bin8_background)
                worksheet_list[1].write(210, 5, "", bin8_background)
                worksheet_list[1].write(210, 6, "", bin8_background)
                worksheet_list[1].write(210, 7, "", bin8_background)
                worksheet_list[1].write(210, 8, "", bin8_background)
                worksheet_list[1].write(210, 9, "", bin8_background)
                worksheet_list[1].write(210, 10, "", bin8_background)
                worksheet_list[1].write(210, 11, "="+str((len(dieNames)-1)/num_excel_sheets)+"-sum(L203:L210)",
                                     bin8_bold_background)
                
                # BL
                # -----------------------------------------------------------
                worksheet_list[2].write(202, 0, classes[0], 
                                     bin0_bold_background)
                worksheet_list[2].write(202, 1, "", bin0_background)
                worksheet_list[2].write(202, 2, "", bin0_background)
                worksheet_list[2].write(202, 3, "", bin0_background)
                worksheet_list[2].write(202, 4, "", bin0_background)
                worksheet_list[2].write(202, 5, "", bin0_background)
                worksheet_list[2].write(202, 6, "", bin0_background)
                worksheet_list[2].write(202, 7, "", bin0_background)
                worksheet_list[2].write(202, 8, "", bin0_background)
                worksheet_list[2].write(202, 9, "", bin0_background)
                worksheet_list[2].write(202, 10, "", bin0_background)
                worksheet_list[2].write(202, 11, bl_bin0,
                                     bin0_bold_background)
                
                worksheet_list[2].write(203, 0, classes[1], 
                                     bin1_bold_background)
                worksheet_list[2].write(203, 1, "", bin1_background)
                worksheet_list[2].write(203, 2, "", bin1_background)
                worksheet_list[2].write(203, 3, "", bin1_background)
                worksheet_list[2].write(203, 4, "", bin1_background)
                worksheet_list[2].write(203, 5, "", bin1_background)
                worksheet_list[2].write(203, 6, "", bin1_background)
                worksheet_list[2].write(203, 7, "", bin1_background)
                worksheet_list[2].write(203, 8, "", bin1_background)
                worksheet_list[2].write(203, 9, "", bin1_background)
                worksheet_list[2].write(203, 10, "", bin1_background)
                worksheet_list[2].write(203, 11, bl_bin1,
                                     bin1_bold_background)
                
                worksheet_list[2].write(204, 0, classes[2], 
                                     bin2_bold_background)
                worksheet_list[2].write(204, 1, "", bin2_background)
                worksheet_list[2].write(204, 2, "", bin2_background)
                worksheet_list[2].write(204, 3, "", bin2_background)
                worksheet_list[2].write(204, 4, "", bin2_background)
                worksheet_list[2].write(204, 5, "", bin2_background)
                worksheet_list[2].write(204, 6, "", bin2_background)
                worksheet_list[2].write(204, 7, "", bin2_background)
                worksheet_list[2].write(204, 8, "", bin2_background)
                worksheet_list[2].write(204, 9, "", bin2_background)
                worksheet_list[2].write(204, 10, "", bin2_background)
                worksheet_list[2].write(204, 11, bl_bin2,
                                     bin2_bold_background)
                
                worksheet_list[2].write(205, 0, classes[3], 
                                     bin3_bold_background)
                worksheet_list[2].write(205, 1, "", bin3_background)
                worksheet_list[2].write(205, 2, "", bin3_background)
                worksheet_list[2].write(205, 3, "", bin3_background)
                worksheet_list[2].write(205, 4, "", bin3_background)
                worksheet_list[2].write(205, 5, "", bin3_background)
                worksheet_list[2].write(205, 6, "", bin3_background)
                worksheet_list[2].write(205, 7, "", bin3_background)
                worksheet_list[2].write(205, 8, "", bin3_background)
                worksheet_list[2].write(205, 9, "", bin3_background)
                worksheet_list[2].write(205, 10, "", bin3_background)
                worksheet_list[2].write(205, 11, bl_bin3,
                                     bin3_bold_background)
                
                worksheet_list[2].write(206, 0, classes[4], 
                                     bin4_bold_background)
                worksheet_list[2].write(206, 1, "", bin4_background)
                worksheet_list[2].write(206, 2, "", bin4_background)
                worksheet_list[2].write(206, 3, "", bin4_background)
                worksheet_list[2].write(206, 4, "", bin4_background)
                worksheet_list[2].write(206, 5, "", bin4_background)
                worksheet_list[2].write(206, 6, "", bin4_background)
                worksheet_list[2].write(206, 7, "", bin4_background)
                worksheet_list[2].write(206, 8, "", bin4_background)
                worksheet_list[2].write(206, 9, "", bin4_background)
                worksheet_list[2].write(206, 10, "", bin4_background)
                worksheet_list[2].write(206, 11, bl_bin4,
                                     bin4_bold_background)
                
                worksheet_list[2].write(207, 0, classes[5], 
                                     bin5_bold_background)
                worksheet_list[2].write(207, 1, "", bin5_background)
                worksheet_list[2].write(207, 2, "", bin5_background)
                worksheet_list[2].write(207, 3, "", bin5_background)
                worksheet_list[2].write(207, 4, "", bin5_background)
                worksheet_list[2].write(207, 5, "", bin5_background)
                worksheet_list[2].write(207, 6, "", bin5_background)
                worksheet_list[2].write(207, 7, "", bin5_background)
                worksheet_list[2].write(207, 8, "", bin5_background)
                worksheet_list[2].write(207, 9, "", bin5_background)
                worksheet_list[2].write(207, 10, "", bin5_background)
                worksheet_list[2].write(207, 11, bl_bin5,
                                     bin5_bold_background)
                
                worksheet_list[2].write(208, 0, classes[6], 
                                     bin6_bold_background)
                worksheet_list[2].write(208, 1, "", bin6_background)
                worksheet_list[2].write(208, 2, "", bin6_background)
                worksheet_list[2].write(208, 3, "", bin6_background)
                worksheet_list[2].write(208, 4, "", bin6_background)
                worksheet_list[2].write(208, 5, "", bin6_background)
                worksheet_list[2].write(208, 6, "", bin6_background)
                worksheet_list[2].write(208, 7, "", bin6_background)
                worksheet_list[2].write(208, 8, "", bin6_background)
                worksheet_list[2].write(208, 9, "", bin6_background)
                worksheet_list[2].write(208, 10, "", bin6_background)
                worksheet_list[2].write(208, 11, bl_bin6,
                                     bin6_bold_background)
                
                worksheet_list[2].write(209, 0, classes[7], 
                                     bin7_bold_background)
                worksheet_list[2].write(209, 1, "", bin7_background)
                worksheet_list[2].write(209, 2, "", bin7_background)
                worksheet_list[2].write(209, 3, "", bin7_background)
                worksheet_list[2].write(209, 4, "", bin7_background)
                worksheet_list[2].write(209, 5, "", bin7_background)
                worksheet_list[2].write(209, 6, "", bin7_background)
                worksheet_list[2].write(209, 7, "", bin7_background)
                worksheet_list[2].write(209, 8, "", bin7_background)
                worksheet_list[2].write(209, 9, "", bin7_background)
                worksheet_list[2].write(209, 10, "", bin7_background)
                worksheet_list[2].write(209, 11, bl_bin7,
                                     bin7_bold_background)
                
                worksheet_list[2].write(210, 0, "8 - Not_Tested-Count", 
                                     bin8_bold_background)
                worksheet_list[2].write(210, 1, "", bin8_background)
                worksheet_list[2].write(210, 2, "", bin8_background)
                worksheet_list[2].write(210, 3, "", bin8_background)
                worksheet_list[2].write(210, 4, "", bin8_background)
                worksheet_list[2].write(210, 5, "", bin8_background)
                worksheet_list[2].write(210, 6, "", bin8_background)
                worksheet_list[2].write(210, 7, "", bin8_background)
                worksheet_list[2].write(210, 8, "", bin8_background)
                worksheet_list[2].write(210, 9, "", bin8_background)
                worksheet_list[2].write(210, 10, "", bin8_background)
                worksheet_list[2].write(210, 11, "="+str((len(dieNames)-1)/num_excel_sheets)+"-sum(L203:L210)",
                                     bin8_bold_background)
                # -----------------------------------------------------------
                
                
                # BR
                # -----------------------------------------------------------
                worksheet_list[3].write(202, 0, classes[0], 
                                     bin0_bold_background)
                worksheet_list[3].write(202, 1, "", bin0_background)
                worksheet_list[3].write(202, 2, "", bin0_background)
                worksheet_list[3].write(202, 3, "", bin0_background)
                worksheet_list[3].write(202, 4, "", bin0_background)
                worksheet_list[3].write(202, 5, "", bin0_background)
                worksheet_list[3].write(202, 6, "", bin0_background)
                worksheet_list[3].write(202, 7, "", bin0_background)
                worksheet_list[3].write(202, 8, "", bin0_background)
                worksheet_list[3].write(202, 9, "", bin0_background)
                worksheet_list[3].write(202, 10, "", bin0_background)
                worksheet_list[3].write(202, 11, br_bin0,
                                     bin0_bold_background)
                
                worksheet_list[3].write(203, 0, classes[1], 
                                     bin1_bold_background)
                worksheet_list[3].write(203, 1, "", bin1_background)
                worksheet_list[3].write(203, 2, "", bin1_background)
                worksheet_list[3].write(203, 3, "", bin1_background)
                worksheet_list[3].write(203, 4, "", bin1_background)
                worksheet_list[3].write(203, 5, "", bin1_background)
                worksheet_list[3].write(203, 6, "", bin1_background)
                worksheet_list[3].write(203, 7, "", bin1_background)
                worksheet_list[3].write(203, 8, "", bin1_background)
                worksheet_list[3].write(203, 9, "", bin1_background)
                worksheet_list[3].write(203, 10, "", bin1_background)
                worksheet_list[3].write(203, 11, br_bin1,
                                     bin1_bold_background)
                
                worksheet_list[3].write(204, 0, classes[2], 
                                     bin2_bold_background)
                worksheet_list[3].write(204, 1, "", bin2_background)
                worksheet_list[3].write(204, 2, "", bin2_background)
                worksheet_list[3].write(204, 3, "", bin2_background)
                worksheet_list[3].write(204, 4, "", bin2_background)
                worksheet_list[3].write(204, 5, "", bin2_background)
                worksheet_list[3].write(204, 6, "", bin2_background)
                worksheet_list[3].write(204, 7, "", bin2_background)
                worksheet_list[3].write(204, 8, "", bin2_background)
                worksheet_list[3].write(204, 9, "", bin2_background)
                worksheet_list[3].write(204, 10, "", bin2_background)
                worksheet_list[3].write(204, 11, br_bin2,
                                     bin2_bold_background)
                
                worksheet_list[3].write(205, 0, classes[3], 
                                     bin3_bold_background)
                worksheet_list[3].write(205, 1, "", bin3_background)
                worksheet_list[3].write(205, 2, "", bin3_background)
                worksheet_list[3].write(205, 3, "", bin3_background)
                worksheet_list[3].write(205, 4, "", bin3_background)
                worksheet_list[3].write(205, 5, "", bin3_background)
                worksheet_list[3].write(205, 6, "", bin3_background)
                worksheet_list[3].write(205, 7, "", bin3_background)
                worksheet_list[3].write(205, 8, "", bin3_background)
                worksheet_list[3].write(205, 9, "", bin3_background)
                worksheet_list[3].write(205, 10, "", bin3_background)
                worksheet_list[3].write(205, 11, br_bin3,
                                     bin3_bold_background)
                
                worksheet_list[3].write(206, 0, classes[4], 
                                     bin4_bold_background)
                worksheet_list[3].write(206, 1, "", bin4_background)
                worksheet_list[3].write(206, 2, "", bin4_background)
                worksheet_list[3].write(206, 3, "", bin4_background)
                worksheet_list[3].write(206, 4, "", bin4_background)
                worksheet_list[3].write(206, 5, "", bin4_background)
                worksheet_list[3].write(206, 6, "", bin4_background)
                worksheet_list[3].write(206, 7, "", bin4_background)
                worksheet_list[3].write(206, 8, "", bin4_background)
                worksheet_list[3].write(206, 9, "", bin4_background)
                worksheet_list[3].write(206, 10, "", bin4_background)
                worksheet_list[3].write(206, 11, br_bin4,
                                     bin4_bold_background)
                
                worksheet_list[3].write(207, 0, classes[5], 
                                     bin5_bold_background)
                worksheet_list[3].write(207, 1, "", bin5_background)
                worksheet_list[3].write(207, 2, "", bin5_background)
                worksheet_list[3].write(207, 3, "", bin5_background)
                worksheet_list[3].write(207, 4, "", bin5_background)
                worksheet_list[3].write(207, 5, "", bin5_background)
                worksheet_list[3].write(207, 6, "", bin5_background)
                worksheet_list[3].write(207, 7, "", bin5_background)
                worksheet_list[3].write(207, 8, "", bin5_background)
                worksheet_list[3].write(207, 9, "", bin5_background)
                worksheet_list[3].write(207, 10, "", bin5_background)
                worksheet_list[3].write(207, 11, br_bin5,
                                     bin5_bold_background)
                
                worksheet_list[3].write(208, 0, classes[6], 
                                     bin6_bold_background)
                worksheet_list[3].write(208, 1, "", bin6_background)
                worksheet_list[3].write(208, 2, "", bin6_background)
                worksheet_list[3].write(208, 3, "", bin6_background)
                worksheet_list[3].write(208, 4, "", bin6_background)
                worksheet_list[3].write(208, 5, "", bin6_background)
                worksheet_list[3].write(208, 6, "", bin6_background)
                worksheet_list[3].write(208, 7, "", bin6_background)
                worksheet_list[3].write(208, 8, "", bin6_background)
                worksheet_list[3].write(208, 9, "", bin6_background)
                worksheet_list[3].write(208, 10, "", bin6_background)
                worksheet_list[3].write(208, 11, br_bin6,
                                     bin6_bold_background)
                
                worksheet_list[3].write(209, 0, classes[7], 
                                     bin7_bold_background)
                worksheet_list[3].write(209, 1, "", bin7_background)
                worksheet_list[3].write(209, 2, "", bin7_background)
                worksheet_list[3].write(209, 3, "", bin7_background)
                worksheet_list[3].write(209, 4, "", bin7_background)
                worksheet_list[3].write(209, 5, "", bin7_background)
                worksheet_list[3].write(209, 6, "", bin7_background)
                worksheet_list[3].write(209, 7, "", bin7_background)
                worksheet_list[3].write(209, 8, "", bin7_background)
                worksheet_list[3].write(209, 9, "", bin7_background)
                worksheet_list[3].write(209, 10, "", bin7_background)
                worksheet_list[3].write(209, 11, br_bin7,
                                     bin7_bold_background)
                
                worksheet_list[3].write(210, 0, "8 - Not_Tested-Count", 
                                     bin8_bold_background)
                worksheet_list[3].write(210, 1, "", bin8_background)
                worksheet_list[3].write(210, 2, "", bin8_background)
                worksheet_list[3].write(210, 3, "", bin8_background)
                worksheet_list[3].write(210, 4, "", bin8_background)
                worksheet_list[3].write(210, 5, "", bin8_background)
                worksheet_list[3].write(210, 6, "", bin8_background)
                worksheet_list[3].write(210, 7, "", bin8_background)
                worksheet_list[3].write(210, 8, "", bin8_background)
                worksheet_list[3].write(210, 9, "", bin8_background)
                worksheet_list[3].write(210, 10, "", bin8_background)
                worksheet_list[3].write(210, 11, "="+str((len(dieNames)-1)/num_excel_sheets)+"-sum(L203:L210)",
                                     bin8_bold_background)
                # -----------------------------------------------------------
            
            
            
            if sqrt(num_excel_sheets) == 2:
                # Write a count for each bin (all four quadrants) at the bottom
                for worksheet_name in worksheet_list:
                    worksheet_name.write(212, 0, "Total - " + classes[0], 
                                         bin0_bold_background)
                    worksheet_name.write(212, 1, "", bin0_background)
                    worksheet_name.write(212, 2, "", bin0_background)
                    worksheet_name.write(212, 3, "", bin0_background)
                    worksheet_name.write(212, 4, "", bin0_background)
                    worksheet_name.write(212, 5, "", bin0_background)
                    worksheet_name.write(212, 6, "", bin0_background)
                    worksheet_name.write(212, 7, "", bin0_background)
                    worksheet_name.write(212, 8, "", bin0_background)
                    worksheet_name.write(212, 9, "", bin0_background)
                    worksheet_name.write(212, 10, "", bin0_background)
                    worksheet_name.write(212, 11, 
                                         (tl_bin0 + tr_bin0 + bl_bin0 + br_bin0),
                                         bin0_bold_background)
                    
                    worksheet_name.write(213, 0, "Total - " + classes[1], 
                                         bin1_bold_background)
                    worksheet_name.write(213, 1, "", bin1_background)
                    worksheet_name.write(213, 2, "", bin1_background)
                    worksheet_name.write(213, 3, "", bin1_background)
                    worksheet_name.write(213, 4, "", bin1_background)
                    worksheet_name.write(213, 5, "", bin1_background)
                    worksheet_name.write(213, 6, "", bin1_background)
                    worksheet_name.write(213, 7, "", bin1_background)
                    worksheet_name.write(213, 8, "", bin1_background)
                    worksheet_name.write(213, 9, "", bin1_background)
                    worksheet_name.write(213, 10, "", bin1_background)
                    worksheet_name.write(213, 11, 
                                         (tl_bin1 + tr_bin1 + bl_bin1 + br_bin1),
                                         bin1_bold_background)
                    
                    worksheet_name.write(214, 0, "Total - " + classes[2], 
                                         bin2_bold_background)
                    worksheet_name.write(214, 1, "", bin2_background)
                    worksheet_name.write(214, 2, "", bin2_background)
                    worksheet_name.write(214, 3, "", bin2_background)
                    worksheet_name.write(214, 4, "", bin2_background)
                    worksheet_name.write(214, 5, "", bin2_background)
                    worksheet_name.write(214, 6, "", bin2_background)
                    worksheet_name.write(214, 7, "", bin2_background)
                    worksheet_name.write(214, 8, "", bin2_background)
                    worksheet_name.write(214, 9, "", bin2_background)
                    worksheet_name.write(214, 10, "", bin2_background)
                    worksheet_name.write(214, 11, 
                                         (tl_bin2 + tr_bin2 + bl_bin2 + br_bin2),
                                         bin2_bold_background)
                    
                    worksheet_name.write(215, 0, "Total - " + classes[3], 
                                         bin3_bold_background)
                    worksheet_name.write(215, 1, "", bin3_background)
                    worksheet_name.write(215, 2, "", bin3_background)
                    worksheet_name.write(215, 3, "", bin3_background)
                    worksheet_name.write(215, 4, "", bin3_background)
                    worksheet_name.write(215, 5, "", bin3_background)
                    worksheet_name.write(215, 6, "", bin3_background)
                    worksheet_name.write(215, 7, "", bin3_background)
                    worksheet_name.write(215, 8, "", bin3_background)
                    worksheet_name.write(215, 9, "", bin3_background)
                    worksheet_name.write(215, 10, "", bin3_background)
                    worksheet_name.write(215, 11, 
                                         (tl_bin3 + tr_bin3 + bl_bin3 + br_bin3),
                                         bin3_bold_background)
                    
                    worksheet_name.write(216, 0, "Total - " + classes[4], 
                                         bin4_bold_background)
                    worksheet_name.write(216, 1, "", bin4_background)
                    worksheet_name.write(216, 2, "", bin4_background)
                    worksheet_name.write(216, 3, "", bin4_background)
                    worksheet_name.write(216, 4, "", bin4_background)
                    worksheet_name.write(216, 5, "", bin4_background)
                    worksheet_name.write(216, 6, "", bin4_background)
                    worksheet_name.write(216, 7, "", bin4_background)
                    worksheet_name.write(216, 8, "", bin4_background)
                    worksheet_name.write(216, 9, "", bin4_background)
                    worksheet_name.write(216, 10, "", bin4_background)
                    worksheet_name.write(216, 11, 
                                         (tl_bin4 + tr_bin4 + bl_bin4 + br_bin4),
                                         bin4_bold_background)
                    
                    worksheet_name.write(217, 0, "Total - " + classes[5], 
                                         bin5_bold_background)
                    worksheet_name.write(217, 1, "", bin5_background)
                    worksheet_name.write(217, 2, "", bin5_background)
                    worksheet_name.write(217, 3, "", bin5_background)
                    worksheet_name.write(217, 4, "", bin5_background)
                    worksheet_name.write(217, 5, "", bin5_background)
                    worksheet_name.write(217, 6, "", bin5_background)
                    worksheet_name.write(217, 7, "", bin5_background)
                    worksheet_name.write(217, 8, "", bin5_background)
                    worksheet_name.write(217, 9, "", bin5_background)
                    worksheet_name.write(217, 10, "", bin5_background)
                    worksheet_name.write(217, 11, 
                                         (tl_bin5 + tr_bin5 + bl_bin5 + br_bin5),
                                         bin5_bold_background)
                    
                    worksheet_name.write(218, 0, "Total - " + classes[6], 
                                         bin6_bold_background)
                    worksheet_name.write(218, 1, "", bin6_background)
                    worksheet_name.write(218, 2, "", bin6_background)
                    worksheet_name.write(218, 3, "", bin6_background)
                    worksheet_name.write(218, 4, "", bin6_background)
                    worksheet_name.write(218, 5, "", bin6_background)
                    worksheet_name.write(218, 6, "", bin6_background)
                    worksheet_name.write(218, 7, "", bin6_background)
                    worksheet_name.write(218, 8, "", bin6_background)
                    worksheet_name.write(218, 9, "", bin6_background)
                    worksheet_name.write(218, 10, "", bin6_background)
                    worksheet_name.write(218, 11, 
                                         (tl_bin6 + tr_bin6 + bl_bin6 + br_bin6),
                                         bin6_bold_background)
                    
                    worksheet_name.write(219, 0, "Total - " + classes[7], 
                                         bin7_bold_background)
                    worksheet_name.write(219, 1, "", bin7_background)
                    worksheet_name.write(219, 2, "", bin7_background)
                    worksheet_name.write(219, 3, "", bin7_background)
                    worksheet_name.write(219, 4, "", bin7_background)
                    worksheet_name.write(219, 5, "", bin7_background)
                    worksheet_name.write(219, 6, "", bin7_background)
                    worksheet_name.write(219, 7, "", bin7_background)
                    worksheet_name.write(219, 8, "", bin7_background)
                    worksheet_name.write(219, 9, "", bin7_background)
                    worksheet_name.write(219, 10, "", bin7_background)
                    worksheet_name.write(219, 11, 
                                         (tl_bin7 + tr_bin7 + bl_bin7 + br_bin7),
                                         bin7_bold_background)
                    
                    worksheet_name.write(220, 0, "Total - 8 - Not_Tested-Count", 
                                         bin8_bold_background)
                    worksheet_name.write(220, 1, "", bin8_background)
                    worksheet_name.write(220, 2, "", bin8_background)
                    worksheet_name.write(220, 3, "", bin8_background)
                    worksheet_name.write(220, 4, "", bin8_background)
                    worksheet_name.write(220, 5, "", bin8_background)
                    worksheet_name.write(220, 6, "", bin8_background)
                    worksheet_name.write(220, 7, "", bin8_background)
                    worksheet_name.write(220, 8, "", bin8_background)
                    worksheet_name.write(220, 9, "", bin8_background)
                    worksheet_name.write(220, 10, "", bin8_background)
                    worksheet_name.write(220, 11, "="+str(len(dieNames)-1)+"-sum(L213:L220)",
                                         bin8_bold_background)
                
                
                    # worksheet_name.set_column(0, 0, width=len(classes[7]))
                    worksheet_name.set_column(0, (col_per_sheet), width=2)
                    worksheet_name.set_column(11, 11, width=6)
            
            workbook.close()
            
            os.makedirs(lotPath + '/ZZZ-Excel_Sheets/', exist_ok=True)
            shutil.copy( (slot_path + '/' + slot_name + '.xlsx'), (lotPath + '/ZZZ-Excel_Sheets/') )
        # -----------------------------------------------------------------------------

print("Done!")

# Stopping stopwatch to see how long process takes
end_time = time.time()
time_lapsed = end_time - start_time
time_convert(time_lapsed)