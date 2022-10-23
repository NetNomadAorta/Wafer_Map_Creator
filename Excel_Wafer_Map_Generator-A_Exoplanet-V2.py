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

# Cycles through each lot folder
for lot_name_index, lot_name in enumerate(os.listdir(PREDICTED_DIR)):
    lot_path = os.path.join(PREDICTED_DIR, lot_name)
    
    # Removes Thumbs.db in wafer map path if found
    if os.path.isfile(STORED_WAFER_DATA + "Thumbs.db"):
        os.remove(STORED_WAFER_DATA + "Thumbs.db")
        
    if "Exoplanet" not in lot_name:
        continue
    
    # Checks to see if lot existing wafer map found in wafer map generator
    #  area for the current lot_path location
    shouldContinue = True
    for waferMapName in os.listdir(STORED_WAFER_DATA):
        if waferMapName in lot_path:
            shouldContinue = False
            break
    if shouldContinue:
        continue
    
    # Imports correct die_names and die_coordinates data
    die_names = np.load(STORED_WAFER_DATA + waferMapName + "/die_names.npy")
    len_die_names = len(die_names)
    
    # Skips if config settings not found
    if not os.path.isfile(STORED_WAFER_DATA + waferMapName + "/config.yaml"):
        continue
    
    settings = yaml.safe_load( open(STORED_WAFER_DATA + waferMapName + "/config.yaml") )
    # Gets each appropriate settings
    classes = settings['classes']
    good_class_index = settings['good_class_index']
    excel_toggle = settings['excel_toggle']
    
    if not excel_toggle:
        continue
    
    # Removes Thumbs.db in lot path if found
    if os.path.isfile(lot_path + "/Thumbs.db"):
        os.remove(lot_path + "/Thumbs.db")
    
    # Cycles through each slot folder within the lot folder
    for slot_name in os.listdir(lot_path):
        slot_path = os.path.join(lot_path, slot_name)
        
        # Checks if Excel sheet already exist, and only skips if selected not to
        if ( (os.path.isfile(slot_path + "/" + slot_name+ ".xlsx") 
        and os.path.isfile(lot_path + '/ZZZ-Excel_Sheets/' + slot_name + '.xlsx') )
        or slot_name == "ZZZ-Excel_Sheets"):
            continue
        
        print("Starting", lot_name, "-", slot_name)
        
        # Removes Thumbs.db in slot path if found
        if os.path.isfile(slot_path + "/Thumbs.db"):
            os.remove(slot_path + "/Thumbs.db")
        
        # For getting ro and col numbers and finding max
        row_list = []
        col_list = []
        
        
        full_list = os.listdir(slot_path)
        # Removes Excel or jpg files from "full_list" variable
        for list_name_index, list_name in enumerate(full_list):
            if ".xlsx" in list_name or ".jpg" in list_name:
                del full_list[list_name_index]
        
        
        # Creates dictionary of die names with defect counts per class
        for die_name_index, die_name in enumerate(die_names):
            
            # Skips first die_names as it contains just string header
            if die_name_index == 0:
                continue
            # Creates dictionary
            elif die_name_index == 1:
                die_dictionary = {
                                  die_name: {
                                             "class_1_count": 0, 
                                             "class_2_count": 0, 
                                             "class_3_count": 0
                                             } 
                                  }
            # Appends dictionary already created
            else:
                die_dictionary[die_name] = {
                                            "class_1_count": 0, 
                                            "class_2_count": 0, 
                                            "class_3_count": 0
                                            } 
            
            # Writes row and col numbers to find max later on
            row_list.append( int( re.findall(r'\d+', die_name)[0] ) )
            col_list.append( int( re.findall(r'\d+', die_name)[1] ) )
        
        
        # Within each slot, cycle through each class
        for class_index, class_name in enumerate(full_list):
            class_path = os.path.join(slot_path, class_name)
            
            # Skips directory if first class (non-defect) folder or if it 
            # includes the wafer map with failing dies image (if this program 
            # already created one from a previous run)
            if (class_index == good_class_index 
            or "ZZ-" in class_name 
            or ".jpg" in class_name 
            or ".xlsx" in class_name
            or "Cropped" not in class_name
            ):
                continue
            
            # Removes Thumbs.db in class path if found
            if os.path.isfile(class_path + "/Thumbs.db"):
                os.remove(class_path + "/Thumbs.db")
            
            
            previous_percent_index = 0
            # Iterates through each class folder's die images
            for class_die_name in os.listdir(class_path):
                
                # Adds defect count of die image files in class folders for respected class count
                die_dictionary[class_die_name.split(".P")[0]]['class_{}_count'.format(class_index+1)] += 1
                
                
                # Shows progress in current slot
                if len_die_names > 1000:
                    percent_index = round(die_name_index/len_die_names*100)
                    
                    if percent_index != previous_percent_index:
                        sys.stdout.write('\033[2K\033[1G')
                        print("  ", class_name, "Progress:", 
                              str(round(die_name_index/len_die_names*100) ) + "%",
                              end="\r"
                              )
                    
                    previous_percent_index = percent_index
            
            
            # Makes new line so that next class progress status can show in terminal/shell
            print("")
        
        
        # XLS Section
        # -----------------------------------------------------------------------------
        if EXCEL_GENERATOR_TOGGLE:
            # XLSX PART
            max_row = max(row_list)
            max_col = max(col_list)
            
            # print("   Starting", slot_name, "Excel sheet results..")
            # Create a workbook and add a worksheet.
            workbook = xlsxwriter.Workbook(slot_path + '/' + slot_name + '.xlsx')
            
            worksheet = workbook.add_worksheet(slot_name)
            
            # Chooses each font and background color for the Excel sheet
            font_color_list = ['black', 'black', 'black', 'white', 
                               'white', 'black', 'black', 'gray']
            bg_color_list = ['lime', 'magenta', 'magenta', 'red', 
                             'red', 'yellow', 'yellow', 'gray']
            
            
            
            # Chooses which font and background associated with each class
            bin_colors_list = []
            bin_bold_colors_list = []
            for class_index in range(len(classes)):
                
                bin_colors_list.append(workbook.add_format(
                    {'font_color': font_color_list[class_index],
                     'bg_color': bg_color_list[class_index],
                     'border': 4
                     }
                    ))
                bin_bold_colors_list.append(workbook.add_format(
                    {'bold': True,
                     'font_color': font_color_list[class_index],
                     'bg_color': bg_color_list[class_index]
                     }
                    ))
                
            # For the "Not Tested Count" gray class
            bin_colors_list.append(workbook.add_format(
                {'font_color': font_color_list[-1],
                 'bg_color': bg_color_list[-1]
                 }
                ))
            bin_bold_colors_list.append(workbook.add_format(
                {'bold': True,
                 'font_color': font_color_list[-1],
                 'bg_color': bg_color_list[-1]
                 }
                ))
            
            # Iterates over each row_per_sheet x col_per_sheet dies 
            #  and defaults bin number to 8 - Untested
            for row in range(max_row*2):
                for col in range(max_col*2):
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
            worksheet.merge_range(max_row*2, 0, max_row*2, 
                                  max_col*2-1, 'Notch', merge_format
                                  )
            
            
            # Iterates through die dictionary names and info
            for die_name, class_count_names in die_dictionary.items():
                row = int( re.findall(r'\d+', die_name)[0] )
                col = int( re.findall(r'\d+', die_name)[1] )
                
                # Iterates through each class count defect info
                for class_count_name, defect_count in class_count_names.items():
                    class_number = int( re.findall(r'\d+', class_count_name)[0] )
                    
                    if class_number == 1:
                        row_to_use = (row-1)*2+0
                        col_to_use = (col-1)*2+1
                        
                        if defect_count > MAXIMUM_DEFECTS_TO_PASS_CLASS_1:
                            background = bin_colors_list[class_number]
                        else:
                            background = bin_colors_list[0]
                        
                    elif class_number == 2:
                        row_to_use = (row-1)*2+0
                        col_to_use = (col-1)*2+0
                        
                        if defect_count > MAXIMUM_DEFECTS_TO_PASS_CLASS_2:
                            background = bin_colors_list[class_number]
                        else:
                            background = bin_colors_list[0]
                        
                    elif class_number == 3:
                        row_to_use = (row-1)*2+1
                        col_to_use = (col-1)*2+0
                        
                        if defect_count > MAXIMUM_DEFECTS_TO_PASS_CLASS_3:
                            background = bin_colors_list[class_number]
                        else:
                            background = bin_colors_list[0]
                    
                    
                    # Writes in Excel sheet each cell appropriate info
                    worksheet.write(row_to_use, col_to_use, 
                                    defect_count, background)
                    
                    # Writes in blank Excel sheet's cell
                    worksheet.write((row-1)*2+1, (col-1)*2+1, 
                                    "", bin_colors_list[0])
            
            
            # Sets the appropriate width for each column
            for row_index in range(max_row*2):
                worksheet.set_row(row_index, height=23)
            worksheet.set_column(0, (max_col*2), width=round((3.3*max_row/max_col), 2) )
            
            # Sets zoom
            worksheet.set_zoom( 120 )
                
                
            
            workbook.close()
            
            os.makedirs(lot_path + '/ZZZ-Excel_Sheets/', exist_ok=True)
            shutil.copy( (slot_path + '/' + slot_name + '.xlsx'), (lot_path + '/ZZZ-Excel_Sheets/') )
        # -----------------------------------------------------------------------------

print("Done!")

# Stopping stopwatch to see how long process takes
end_time = time.time()
time_lapsed = end_time - start_time
time_convert(time_lapsed)


# if __name__ == "__main__":
#     main()


