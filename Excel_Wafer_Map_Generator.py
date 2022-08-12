# Import the necessary packages
import os
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


def time_convert(sec):
    mins = sec // 60
    sec = sec % 60
    hours = mins // 60
    mins = mins % 60
    print("Time Lapsed = {0}h:{1}m:{2}s".format(int(hours), int(mins), round(sec) ) )



# MAIN():
# =============================================================================
def main():
    # Starting stopwatch to see how long process takes
    start_time = time.time()
    
    # Clears some of the screen for asthetics
    print("\n\n\n")
    
    # Cycles through each lot folder
    for lotPathIndex, lotPath in enumerate(glob.glob(PREDICTED_DIR + "*") ):
        lotPathName = os.listdir(PREDICTED_DIR)[lotPathIndex]
        
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
            # For getting ro and col numbers and finding max
            row_list = []
            col_list = []
            bad_row_list = []
            bad_col_list = []
            
            
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
                or ".xlsx" in class_name):
                    continue
                # Removes Thumbs.db in class path if found
                if os.path.isfile(class_path + "/Thumbs.db"):
                    os.remove(class_path + "/Thumbs.db")
                
                # Looks at die names in previously created wafer map and sees
                #  if this slot has the same die names within its defect folders.
                # If so, then create the new wafer map with red ovals in die 
                #  location within the wafer map image, and save this image.
                class_dies_list = os.listdir(class_path)
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
                    if any(dieName in s for s in class_dies_list):
                        isBadDie = True
                        badDieNames.append(dieName)
                        badDieBinNumbers.append(class_index)
                        bad_row_list.append( int( re.findall(r'\d+', dieName)[0] ) )
                        bad_col_list.append( int( re.findall(r'\d+', dieName)[1] ) )
                    else:
                        isBadDie = False
    
                    
                    if isBadDie:
                        if len(dieNames) > 1000 and dieNameIndex % 1000 == 0:
                            for list_index, image_name in enumerate(class_dies_list):
                                if dieName in image_name:
                                    del class_dies_list[:list_index]
                                    break
                        
                        continue
                    
    
    
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
                if waferMapName == "SMiPE4" or waferMapName == "TPv2":
                    font_color_list = ['white', 'black', 'white', 'white', 'black', 
                                       'white', 'white', 'black', 'white']
                    bg_color_list = ['black', 'lime', 'red', 'green', 'yellow', 
                                     'blue', 'magenta', 'cyan', 'gray']
                else:
                    font_color_list = ['black', 'white', 'white', 'white', 'white']
                    bg_color_list = ['lime', 'red', 'blue', 'magenta', 'gray']
                
                
                # Chooses which font and background associated with each class
                bin_colors_list = []
                bin_bold_colors_list = []
                for class_index in range(len(classes_2)):
                    bin_colors_list.append(workbook.add_format(
                        {'font_color': font_color_list[class_index],
                         'bg_color': bg_color_list[class_index]}))
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
                if waferMapName == "HBCOSA":
                    for row in range(row_per_sheet+1):
                        for col in range(col_per_sheet+1):
                            for worksheet in worksheet_list:
                                worksheet.write(row, col, 0, bin_colors_list[-1])
                else:
                    for row in range(row_per_sheet):
                        for col in range(col_per_sheet):
                            for worksheet in worksheet_list:
                                worksheet.write(row, col, 8, bin_colors_list[-1])
                
                
                # Combines all die names and bin numbers
                all_dieNames = badDieNames
                all_dieBinNumbers = badDieBinNumbers
                if len_dieNames > 1000:
                    print("   Started making good bins..")
                
                full_list = glob.glob(slot_path + "/*")
                for list_name_index, list_name in enumerate(full_list):
                    if ".xlsx" in list_name or ".jpg" in list_name:
                        del full_list[list_name_index]
                        
                list = os.listdir(full_list[good_class_index_2])
                
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
                        all_dieBinNumbers.append(good_class_index_2)
                
                    if len(dieNames) > 1000 and dieNameIndex % 1000 == 0:
                        for list_index, image_name in enumerate(list):
                            if dieName in image_name:
                                del list[:list_index]
                                break
                
                if len_dieNames > 1000:
                    print("   Started writing Excel sheet bin numbers..")
                # Writes all dies info in Excel
                for all_dieName_index, all_dieName in enumerate(all_dieNames):
                    row = int( re.findall(r'\d+', all_dieName)[0] )
                    col = int( re.findall(r'\d+', all_dieName)[1] )
                    
                    # Checks to see which background bin number to use
                    background = bin_colors_list[all_dieBinNumbers[all_dieName_index]]
                    
                    if waferMapName == "HBCOSA":
                        bin_number = all_dieBinNumbers[all_dieName_index] + 1
                    else:
                        bin_number = all_dieBinNumbers[all_dieName_index]
                    
                    if row <= row_per_sheet:
                        if col <= col_per_sheet:
                            # Hyperlink
                            if bin_number != good_class_index_2:
                                worksheet_list[0].write_url(row-1, col-1,
                                    slot_path + '/' + classes_2[bin_number] + '/Row_{}.Col_{}.jpg'.format(str(row), str(col))
                                                            )
                            # Non Hyperlink - Just writes bins
                            worksheet_list[0].write(row-1, col-1, 
                                                bin_number, 
                                                background)
                            
                        else:
                            # Hyperlink
                            if bin_number != good_class_index_2:
                                worksheet_list[1].write_url(row-1, col-1,
                                    slot_path + '/' + classes_2[bin_number] + '/Row_{}.Col_{}.jpg'.format(str(row), str(col))
                                                            )
                            # Non Hyperlink - Just writes bins
                            worksheet_list[1].write(row-1, col-1-col_per_sheet, 
                                               bin_number,
                                               background)
                    else:
                        if col <= col_per_sheet:
                            # Hyperlink
                            if bin_number != good_class_index_2:
                                worksheet_list[2].write_url(row-1, col-1,
                                    slot_path + '/' + classes_2[bin_number] + '/Row_{}.Col_{}.jpg'.format(str(row), str(col))
                                                            )
                            # Non Hyperlink - Just writes bins
                            worksheet_list[2].write(row-1-row_per_sheet, col-1, 
                                               bin_number,
                                               background)
                        else:
                            # Hyperlink
                            if bin_number != good_class_index_2:
                                worksheet_list[3].write_url(row-1, col-1,
                                    slot_path + '/' + classes_2[bin_number] + '/Row_{}.Col_{}.jpg'.format(str(row), str(col))
                                                            )
                            # Non Hyperlink - Just writes bins
                            worksheet_list[3].write(row-1-row_per_sheet, col-1-col_per_sheet, 
                                               bin_number,
                                               background)
                
                
                bin_count_dict = {}
                for worksheet_index in range(len(worksheet_list)):
                    bin_count_dict[worksheet_index] = {}
                    for bin_index in range(len(classes_2)):
                        bin_count_dict[worksheet_index]["bin{}".format(bin_index)] = 0
                
                # Counts how many bins in each worksheet
                for worksheet_index, worksheet in enumerate(worksheet_list):
                    for row in range(row_per_sheet):
                        for col in range(col_per_sheet):
                            bin_num = worksheet.table[row][col].number
                            if waferMapName == "HBCOSA":
                                if bin_num == 0:
                                    continue
                                bin_num -= 1
                            if bin_num == 8:
                                continue
                            bin_count_dict[worksheet_index]["bin{}".format(bin_num)] += 1
                
                
                # Selects appropriate "Not Tested Count" name
                if waferMapName == "SMiPE4":
                    not_tested_name = "8 - Not_Tested-Count"
                elif waferMapName == "HBCOSA":
                    not_tested_name = "0-Not_Tested-Count"
                else:
                    not_tested_name = "8-Not_Tested-Count"
                
                # For each sheet, writes bin class count and colors background
                for worksheet_index, worksheet in enumerate(worksheet_list):
                    for class_index, class_name in enumerate(classes_2):
                        # Writes in bold and makes color background for each sheet a count of class bins
                        worksheet.write(int(max_row/sqrt(len(worksheet_list) ) ) + 2 + class_index, 0, 
                            class_name, bin_bold_colors_list[class_index]
                            )
                        worksheet.write(int(max_row/sqrt(len(worksheet_list) ) ) + 2 + class_index, 11, 
                            bin_count_dict[worksheet_index]["bin{}".format(class_index)], 
                            bin_bold_colors_list[class_index]
                            )
                        
                        # Not Tested Count Section
                        worksheet.write(int(max_row/sqrt(len(worksheet_list) ) ) + 2 + len(classes_2), 0, 
                            not_tested_name, bin_bold_colors_list[-1]
                            )
                        worksheet.write(int(max_row/sqrt(len(worksheet_list) ) ) + 2 + len(classes_2), 11, 
                            "="+str((len(dieNames)-1)/num_excel_sheets_2)
                            +"-sum(L{}:L{})".format((int(max_row/sqrt(len(worksheet_list) ) ) + 3),
                                                    (int(max_row/sqrt(len(worksheet_list) ) ) + 2 + len(classes_2))), 
                            bin_bold_colors_list[-1]
                            )
                        
                        for index in range(10):
                            worksheet.write(int(max_row/sqrt(len(worksheet_list) ) ) + 2 + class_index, (index+1), 
                                "", bin_colors_list[class_index]
                                )
                            worksheet.write(int(max_row/sqrt(len(worksheet_list) ) ) + 2 + len(classes_2), (index+1), 
                                "", bin_colors_list[-1]
                                )
                        
                        # Writes an additional total count in case more than one sheet with total of sum of each sheet
                        if len(worksheet_list) > 1:
                            # Writes in bold and makes color background for each sheet a count of class bins
                            worksheet.write(int(max_row/sqrt(len(worksheet_list) ) ) + 4 + len(classes_2) + class_index, 0, 
                                "Total - " + class_name, 
                                bin_bold_colors_list[class_index]
                                )
                            tot_count = 0
                            for worksheet_index_v2 in range(len(worksheet_list)):
                                tot_count += bin_count_dict[worksheet_index_v2]["bin{}".format(class_index)]
                            worksheet.write(int(max_row/sqrt(len(worksheet_list) ) ) + 4 + len(classes_2) + class_index, 11, 
                                tot_count, 
                                bin_bold_colors_list[class_index]
                                )
                            
                            # Not Tested Count Section
                            worksheet.write(int(max_row/sqrt(len(worksheet_list) ) ) + 4 + len(classes_2) + len(classes_2), 0, 
                                "Total - " + not_tested_name, 
                                bin_bold_colors_list[-1]
                                )
                            worksheet.write(int(max_row/sqrt(len(worksheet_list) ) ) + 4 + len(classes_2) + len(classes_2), 11, 
                                "="+str(len(dieNames)-1)
                                +"-sum(L{}:L{})".format((int(max_row/sqrt(len(worksheet_list) ) ) + 5 + len(classes_2)),
                                                        (int(max_row/sqrt(len(worksheet_list) ) ) + 4 + len(classes_2) + len(classes_2))), 
                                bin_bold_colors_list[-1]
                                )
                            
                            for index in range(10):
                                worksheet.write(int(max_row/sqrt(len(worksheet_list) ) ) + 4 + len(classes_2) + class_index, (index+1), 
                                    "", bin_colors_list[class_index]
                                    )
                                worksheet.write(int(max_row/sqrt(len(worksheet_list) ) ) + 4 + len(classes_2) + len(classes_2), (index+1), 
                                    "", bin_colors_list[-1]
                                    )
                    
                    # Sets the appropriate width for each column
                    worksheet.set_column(0, (col_per_sheet), width=round((20*max_row/max_col)*.12, 2) )
                    
                    if len(worksheet_list) > 1: 
                        worksheet.set_column(11, 11, width=8)
                    elif len(all_dieNames) < 1000:
                        worksheet.set_column(11, 11, width=3.5)
                    else:
                        worksheet.set_column(11, 11, width=7)
                    
                    # Sets zoom
                    worksheet.set_zoom( max( int(2080.6*(max_row/sqrt(len(worksheet_list)))**-0.867 ) , 20 ) )
                    
                    
                
                workbook.close()
                
                os.makedirs(lotPath + '/ZZZ-Excel_Sheets/', exist_ok=True)
                shutil.copy( (slot_path + '/' + slot_name + '.xlsx'), (lotPath + '/ZZZ-Excel_Sheets/') )
            # -----------------------------------------------------------------------------
    
    print("Done!")
    
    # Stopping stopwatch to see how long process takes
    end_time = time.time()
    time_lapsed = end_time - start_time
    time_convert(time_lapsed)


if __name__ == "__main__":
    main()


