"""
This python script uses the wafer map and coordinates created in 
 "Main_Wafer_Map_Creator.py" to create a wafer map with red ovals around each
 die location in the wafer map that are failing. 
It does this by looking at existing images, with row and column numbers 
 included in the images' name, that are found in the non-defect folder from the
 "Automated_AOI.py" output.
If an inlet and outlet folder exist and COMPARE_OVERLAY is true, then this will
 create an output wafer map with green/ovals around dies that are no longer 
 failing from the inlet folder
"""

# Import the necessary packages
import os
import shutil
import glob
import cv2
import time
import numpy as np
import xlsxwriter

# User Parameters/Constants to Set
PREDICTED_DIR = "C:/Users/troya/.spyder-py3/ML-Defect_Detection/Images/Prediction_Images/Predicted_Images/"
STORED_WAFER_DATA = "C:/Users/troya/.spyder-py3/Wafer_Map_Creator/Images/002-Wafer_Map/"
COMPARE_OVERLAY = False # Will compare "*-In" and "*-Out" wafer maps and output in "*-Out" folder
SHOULD_REPLACE_ALL_MAPS = False # Will remake each wafer map that already exist in AOI Output folder if set true
WAFER_MAP_SIZE_LIMIT = 500 # mb # If wafer map size above this value, reduce quality until size is under this value


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


    
# Imports correct dieNames and dieCoordinates data
dieNames = np.load("C:/Users/troya/.spyder-py3/Wafer_Map_Creator/Images/002-Wafer_Map/LED"+ "/dieNames.npy")
len_dieNames = len(dieNames)
dieCoordinates = np.load("C:/Users/troya/.spyder-py3/Wafer_Map_Creator/Images/002-Wafer_Map/LED" + "/Coordinates.npy")



# Removes Thumbs.db in lot path if found
if os.path.isfile(PREDICTED_DIR + "/Thumbs.db"):
    os.remove(PREDICTED_DIR + "/Thumbs.db")

# Creates wafer map
waferMap = cv2.imread("C:/Users/troya/.spyder-py3/Wafer_Map_Creator/Images/002-Wafer_Map/LED" +  "/Wafer_Map.jpg")

# Making list of bad die names
badDieNames = []
badDieCoordinates = []
badDieBinNumbers = []
# Start count for failing dies
numFailingDies = 0

# Below is needed incase there is no bad dies
isFirstImageRun = True

# Within each slot, cycle through each class
for classIndex, classPath in enumerate(glob.glob(PREDICTED_DIR + "*") ):
    # Skips directory if first class (non-defect) folder or if it 
    # includes the wafer map with failing dies image (if this program 
    # already created one from a previous run)
    if classIndex == 1 \
    or "ZZ-" in classPath \
    or ".jpg" in classPath:
        continue
    # Removes Thumbs.db in class path if found
    if os.path.isfile(classPath + "/Thumbs.db"):
        os.remove(classPath + "/Thumbs.db")
    
    # Looks at die names in previously created wafer map and sees
    #  if this slot has the same die names within its defect folders.
    # If so, then create the new wafer map with red ovals in die 
    #  location within the wafer map image, and save this image.
    list = os.listdir(classPath)
    
    (shown_progress_25, shown_progress_50, 
     shown_progress_75, shown_progress_100) = False, False, False, False
    
    for dieNameIndex, dieName in enumerate(dieNames):
        isBadDie = False
        
        if dieNameIndex == 0:
            continue
        
        # Shows progress in current slot
        if len_dieNames > 100:
            if (round(dieNameIndex/len_dieNames, 2) == 0.25
            and shown_progress_25 == False):
                print("  ", classPath[-25:], "Progress:", 
                      str(round(dieNameIndex/len_dieNames*100) ) + "%")
                shown_progress_25 = True
            if (round(dieNameIndex/len_dieNames, 2) == 0.50
            and shown_progress_50 == False):
                print("  ", classPath[-25:], "Progress:", 
                      str(round(dieNameIndex/len_dieNames*100) ) + "%")
                shown_progress_50 = True
            if (round(dieNameIndex/len_dieNames, 2) == 0.75
            and shown_progress_75 == False):
                print("  ", classPath[-25:], "Progress:", 
                      str(round(dieNameIndex/len_dieNames*100) ) + "%")
                shown_progress_75 = True
            if (round(dieNameIndex/len_dieNames, 2) == 1.00
            and shown_progress_100 == False):
                print("  ", classPath[-25:], "Progress:", 
                      str(round(dieNameIndex/len_dieNames*100) ) + "%")
                shown_progress_100 = True

        # Checks if same die name already claimed as bad in previous class folder
        if dieName in badDieNames:
            continue
            
        if any(dieName in s for s in list):
            isBadDie = True
            badDieNames.append(dieName)
            badDieCoordinates.append(dieCoordinates[dieNameIndex])
            badDieBinNumbers.append(classIndex)
        else:
            isBadDie = False
        
        
        x1 = dieCoordinates[dieNameIndex][0]
        y1 = dieCoordinates[dieNameIndex][1]
        x2 = dieCoordinates[dieNameIndex][2]
        y2 = dieCoordinates[dieNameIndex][3]
        
        midX    = round( (x1 + x2)/2)
        midY    = round( (y1 + y2)/2)
        
        lengthX = round( ( (x2 - x1) * 0.97 )/2)
        lengthY = round( ( (y2 - y1) * 0.97 )/2)
        
        lengthXInner = round( ( (x2 - x1) * 0.87 )/2)
        lengthYInner = round( ( (y2 - y1) * 0.87 )/2)
        
        # Places red ovals over wafer map using bad die's coordinate
        center      = (midX, midY)
        axes        = (lengthX, lengthY)
        angle       = 0
        startAngle  = 0
        endAngle    = 360
        if isBadDie:
            color       = (0, 0, 255)
        else:
            color       = (0, 255, 0)
        thickness   = round(lengthX * 0.03)
        
        cv2.ellipse(waferMap, center, 
                    axes, 
                    angle, 
                    startAngle, 
                    endAngle, 
                    color,
                    thickness
                    )
        
        # Prevents green circles from being drawn
        isFirstImageRun = False
        
        if isBadDie:
            
            if len(dieNames) > 1000 and dieNameIndex % 1000 == 0:
                for list_index, image_name in enumerate(list):
                    if dieName in image_name:
                        del list[:list_index]
                        break
            
            numFailingDies += 1
            continue
        
        if isFirstImageRun:
            x1 = dieCoordinates[dieNameIndex][0]
            y1 = dieCoordinates[dieNameIndex][1]
            x2 = dieCoordinates[dieNameIndex][2]
            y2 = dieCoordinates[dieNameIndex][3]
            
            midX    = round( (x1 + x2)/2)
            midY    = round( (y1 + y2)/2)
            
            lengthX = round( ( (x2 - x1) * 0.97 )/2)
            lengthY = round( ( (y2 - y1) * 0.97 )/2)
            
            lengthXInner = round( ( (x2 - x1) * 0.87 )/2)
            lengthYInner = round( ( (y2 - y1) * 0.87 )/2)
            
            # Places red ovals over wafer map using bad die's coordinate
            center      = (midX, midY)
            axes        = (lengthX, lengthY)
            angle       = 0
            startAngle  = 0
            endAngle    = 360
            color       = (0, 255, 0)
            thickness   = round(lengthX * 0.03)
            
            cv2.ellipse(waferMap, center, 
                        axes, 
                        angle, 
                        startAngle, 
                        endAngle, 
                        color,
                        thickness
                        )
    
    # Insert if classpath[i] in classpath[i+1] then numFailingDies -= 1
    
    classIndex += 1



image_size_limit = WAFER_MAP_SIZE_LIMIT # in mb
image_size_limit = image_size_limit * 1000000 # now im bytes

# Saves Wafer Map and deletes Temp Wafer Map if needed
percent_knockoff = 5
while True:
    cv2.imwrite(PREDICTED_DIR + "Wafer_Map_with_Failing_Dies.jpg", 
                waferMap, 
                [cv2.IMWRITE_JPEG_QUALITY, 100-percent_knockoff])
    size = os.path.getsize(PREDICTED_DIR + "Wafer_Map_with_Failing_Dies.jpg")
    if size > image_size_limit:
        os.remove(PREDICTED_DIR + "Wafer_Map_with_Failing_Dies.jpg")
        percent_knockoff += 10
    else:
        break

# XLS Section
# -----------------------------------------------------------------------------
print("   Starting Excel sheet results..")
# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook(PREDICTED_DIR + 'Results.xlsx')
worksheet_TL = workbook.add_worksheet("TL")
worksheet_TR = workbook.add_worksheet("TR")
worksheet_BL = workbook.add_worksheet("BL")
worksheet_BR = workbook.add_worksheet("BR")
worksheet_list = [worksheet_TL, worksheet_TR, worksheet_BL, worksheet_BR]

# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})

# Start from the third row cell. Rows and columns are zero indexed.
row = 0
col = 3

# Iterates over each 200x200 dies and defaults bin number to 8
for row in range(200):
    for col in range(200):
        worksheet_TL.write(row, col, 8)
        worksheet_TR.write(row, col, 8)
        worksheet_BL.write(row, col, 8)
        worksheet_BR.write(row, col, 8)


# Combines all die names and bin numbers
all_dieNames = badDieNames
all_dieBinNumbers = badDieBinNumbers
print("   Started making good bins..")
list = os.listdir(glob.glob(PREDICTED_DIR + "*")[1])
# Checks to see which are good dies since previous scan in classes skipped good dies
for dieNameIndex, dieName in enumerate(dieNames):
    if any(dieName in s for s in list):
        all_dieNames.append(dieName)
        all_dieBinNumbers.append(1)

    if len(dieNames) > 1000 and dieNameIndex % 1000 == 0:
        for list_index, image_name in enumerate(list):
            if dieName in image_name:
                del list[:list_index]
                break

print("   Started writing Excel sheet bin numbers..")
# Writes all dies info in Excel
for all_dieName_index, all_dieName in enumerate(all_dieNames):
    row = int(all_dieName[4:7])
    col = int(all_dieName[-3:])
    
    if row <= 200:
        if col <= 200:
            worksheet_TL.write(row, col, all_dieBinNumbers[all_dieName_index])
        else:
            worksheet_TR.write(row, col, all_dieBinNumbers[all_dieName_index])
    else:
        if col <= 200:
            worksheet_BL.write(row, col, all_dieBinNumbers[all_dieName_index])
        else:
            worksheet_BR.write(row, col, all_dieBinNumbers[all_dieName_index])

# Write a count for each bin at the bottom
for worksheet_name in worksheet_list:
    worksheet_name.write(202, 0, "0-Bad-Count", bold)
    worksheet_name.write(202, 4, all_dieBinNumbers.count(0))
    
    worksheet_name.write(203, 0, "1-All_Good-Count", bold)
    worksheet_name.write(203, 4, all_dieBinNumbers.count(1))
    
    worksheet_name.write(204, 0, "2-Red_Only_Present-Count", bold)
    worksheet_name.write(204, 4, all_dieBinNumbers.count(2))
    
    worksheet_name.write(205, 0, "3-Green_Only_Present", bold)
    worksheet_name.write(205, 4, all_dieBinNumbers.count(3))
    
    worksheet_name.write(206, 0, "4-Red_Green_Only_Present-Count", bold)
    worksheet_name.write(206, 4, all_dieBinNumbers.count(4))
    
    worksheet_name.write(207, 0, "5-Blue_Only_Present", bold)
    worksheet_name.write(207, 4, all_dieBinNumbers.count(5))
    
    worksheet_name.write(208, 0, "6-Red_Blue_Only_Present-Count", bold)
    worksheet_name.write(208, 4, all_dieBinNumbers.count(6))
    
    worksheet_name.write(209, 0, "7-Green_Blue_Only_Present-Count", bold)
    worksheet_name.write(209, 4, all_dieBinNumbers.count(7))
    
    worksheet_name.write(210, 0, "8-Not_Tested-Count", bold)
    worksheet_name.write(210, 4, "=160000-sum(E203:E210)")
    
    
    # worksheet_name.set_column(0, 0, width=len("7-Green_Blue_Only_Present-Count"))
    worksheet_name.set_column(4, 4, width=7)

workbook.close()
# -----------------------------------------------------------------------------


print("Done!")

# Stopping stopwatch to see how long process takes
end_time = time.time()
time_lapsed = end_time - start_time
time_convert(time_lapsed)