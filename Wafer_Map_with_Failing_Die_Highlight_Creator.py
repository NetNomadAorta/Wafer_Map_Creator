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
PREDICTED_DIR = "//mcrtp-file-01.mcusa.local/public/000-AOI_Tool_Output/"
STORED_WAFER_DATA = "//mcrtp-file-01.mcusa.local/public/000-AOI_Tool_Output/ZZZ-General_Wafer_Map_Data/"
COMPARE_OVERLAY = False # Will compare "*-In" and "*-Out" wafer maps and output in "*-Out" folder
SHOULD_REPLACE_ALL_MAPS = False # Will remake each wafer map that already exist in AOI Output folder if set true
WAFER_MAP_SIZE_LIMIT = 300 # mb # If wafer map size above this value, reduce quality until size is under this value
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
    
    # Sets parameter to enable comparing in and out files
    isInletLot = False
    isCompareMap = False
    if "-In" in lotPath and COMPARE_OVERLAY == True:
        isInletLot = True
    if "-Out" in lotPath and COMPARE_OVERLAY == True:
        isCompareMap = True
        inletLotPath = lotPath.replace("-Out", "-In")
    
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


    
    # Removes Thumbs.db in lot path if found
    if os.path.isfile(lotPath + "/Thumbs.db"):
        os.remove(lotPath + "/Thumbs.db")
    
    # Cycles through each slot folder within the lot folder
    for slotPath in glob.glob(lotPath + "/*"):
        # Creates wafer map
        waferMap = cv2.imread(STORED_WAFER_DATA + waferMapName +  "/Wafer_Map.jpg")
        
        # Checks if wafer map already exist, and only skips if selected not to
        if os.path.isfile(slotPath + "/Wafer_Map_with_Failing_Dies.jpg") \
        and SHOULD_REPLACE_ALL_MAPS == False:
            continue
        
        isUsingOriginalMap = False
        if isInletLot:
            tempWaferMap = waferMap.copy()
        elif isCompareMap:
            inletSlotPath = slotPath.replace("-Out", "-In")
            if os.path.isfile(inletSlotPath + "/Temp_Wafer_Map_to_Compare.jpg"):
                waferMap = cv2.imread(inletSlotPath + "/Temp_Wafer_Map_to_Compare.jpg")
                isUsingOriginalMap = False
            else:
                isUsingOriginalMap = True
        
        # Removes Thumbs.db in slot path if found
        if os.path.isfile(slotPath + "/Thumbs.db"):
            os.remove(slotPath + "/Thumbs.db")
        
        # Making list of bad die names
        badDieNames = []
        badDieCoordinates = []
        badDieBinNumbers = []
        # Start count for failing dies
        numFailingDies = 0
        
        # Below is needed incase there is no bad dies
        isFirstImageRun = True
         
        # Within each slot, cycle through each class
        for classIndex, classPath in enumerate(glob.glob(slotPath + "/*") ):
            # Skips directory if first class (non-defect) folder or if it 
            # includes the wafer map with failing dies image (if this program 
            # already created one from a previous run)
            
            # WAS classIndex == 0 NOW 1 FOR X-Display!! CHANGGGEE BACCKKKKK
            if classIndex == 1 \
            or "ZZ-" in os.listdir(slotPath)[classIndex] \
            or ".jpg" in os.listdir(slotPath)[classIndex]:
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
                if len_dieNames > 1000:
                    if (round(dieNameIndex/len_dieNames, 2) == 0.25
                    and shown_progress_25 == False):
                        print("  ", classPath[-30:], "Progress:", 
                              str(round(dieNameIndex/len_dieNames*100) ) + "%")
                        shown_progress_25 = True
                    if (round(dieNameIndex/len_dieNames, 2) == 0.50
                    and shown_progress_50 == False):
                        print("  ", classPath[-30:], "Progress:", 
                              str(round(dieNameIndex/len_dieNames*100) ) + "%")
                        shown_progress_50 = True
                    if (round(dieNameIndex/len_dieNames, 2) == 0.75
                    and shown_progress_75 == False):
                        print("  ", classPath[-30:], "Progress:", 
                              str(round(dieNameIndex/len_dieNames*100) ) + "%")
                        shown_progress_75 = True
                    if (round(dieNameIndex/len_dieNames, 2) == 1.00
                    and shown_progress_100 == False):
                        print("  ", classPath[-30:], "Progress:", 
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
                
                if isInletLot:
                    # Places green/orange ovals over wafer map using bad die's coordinate
                    # # This is used for "*-Out" file to overlay
                    axes  = (lengthXInner, lengthYInner)
                    if isBadDie:
                        color       = (0, 0, 255)
                    else:
                        color       = (0, 255, 0)
                    
                    cv2.ellipse(tempWaferMap, center, 
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
        
        slotName = slotPath[len(lotPath)+1:]
        
        # Writes slot name on top left
        font                   = cv2.FONT_HERSHEY_SIMPLEX
        bottomLeftCornerOfText = (round(waferMap.shape[1]*1/80), 
                                  round(waferMap.shape[0]*1/35))
        fontScale              = round(0.0007*waferMap.shape[1], 2)
        fontColor              = (255, 255, 255)
        thickness              = round(0.0015*waferMap.shape[1])
        lineType               = 2
        
        
        if not isCompareMap or isUsingOriginalMap:
            bottomLeftCornerOfText = (round(waferMap.shape[1]*1/80), 
                                      round(waferMap.shape[0]*1/35))
            cv2.putText(waferMap, 
                        "Lot Name: " + str(lotPathName), 
                        bottomLeftCornerOfText, 
                        font, 
                        fontScale,
                        fontColor,
                        thickness,
                        lineType
                        )
            
            bottomLeftCornerOfText = (round(waferMap.shape[1]*1/80), 
                                      round(waferMap.shape[0]*2/35))
            cv2.putText(waferMap, 
                        "Slot Name: " + str(slotName), 
                        bottomLeftCornerOfText, 
                        font, 
                        fontScale,
                        fontColor,
                        thickness,
                        lineType
                        )
        
        if isCompareMap and not isUsingOriginalMap:
            bottomLeftCornerOfText = (round(waferMap.shape[1]*1/80), 
                                      round(waferMap.shape[0]*1/35))
            cv2.putText(waferMap, 
                        "Lot Name: " + str(lotPathName), 
                        bottomLeftCornerOfText, 
                        font, 
                        fontScale,
                        fontColor,
                        thickness,
                        lineType
                        )
        
        # Also writes name in temporary wafer map if available
        if isInletLot:
            bottomLeftCornerOfText = (round(waferMap.shape[1]*1/80), 
                                      round(waferMap.shape[0]*2/35))
            cv2.putText(tempWaferMap, 
                        "Slot Name: " + str(slotName), 
                        bottomLeftCornerOfText, 
                        font, 
                        fontScale,
                        fontColor,
                        thickness,
                        lineType
                        )
        
        
        # Writes legend info on top right
        font                   = cv2.FONT_HERSHEY_SIMPLEX
        bottomLeftCornerOfText = (round(waferMap.shape[1]*23/30), 
                                  round(waferMap.shape[0]*1/50))
        fontScale              = round(0.0004*waferMap.shape[1], 2)
        fontColor              = (255, 255, 255)
        thickness              = round(0.0013*waferMap.shape[1])
        lineType               = 2
        
        if not isCompareMap or isUsingOriginalMap:
            bottomLeftCornerOfText = (round(waferMap.shape[1]*24/30), 
                                      round(waferMap.shape[0]*3/100))
            cv2.putText(waferMap, 
                        "Green: Passing; Red: Failing", 
                        bottomLeftCornerOfText, 
                        font, 
                        fontScale,
                        fontColor,
                        thickness,
                        lineType
                        )
            
        # Only writes extra legend info on comparing maps using temp map
        if isCompareMap and not isUsingOriginalMap:
            bottomLeftCornerOfText = (round(waferMap.shape[1]*24/30), 
                                      round(waferMap.shape[0]*5/100))
            cv2.putText(waferMap, 
                        "Inner Circle: Incoming Wafer", 
                        bottomLeftCornerOfText, 
                        font, 
                        fontScale,
                        fontColor,
                        thickness,
                        lineType
                        )
            
            bottomLeftCornerOfText = (round(waferMap.shape[1]*24/30), 
                                      round(waferMap.shape[0]*7/100))
            cv2.putText(waferMap, 
                        "Outer Circle: Final Wafer", 
                        bottomLeftCornerOfText, 
                        font, 
                        fontScale,
                        fontColor,
                        thickness,
                        lineType
                        )
            
            
        
        # Also writes name in temporary wafer map if available
        if isInletLot:
            bottomLeftCornerOfText = (round(waferMap.shape[1]*24/30), 
                                      round(waferMap.shape[0]*3/100))
            cv2.putText(tempWaferMap, 
                        "Green: Passing; Red: Failing", 
                        bottomLeftCornerOfText, 
                        font, 
                        fontScale,
                        fontColor,
                        thickness,
                        lineType
                        )
        
        
        # Writes how many failing defects on bottom right
        font                   = cv2.FONT_HERSHEY_SIMPLEX
        bottomLeftCornerOfText = (round(waferMap.shape[1]*23/30), 
                                  round(waferMap.shape[0]*79/80))
        fontScale              = round(0.00055*waferMap.shape[1], 2)
        fontColor              = (0, 0, 255)
        thickness              = round(0.0014*waferMap.shape[1])
        lineType               = 2
        
        if isInletLot:
            bottomLeftCornerOfText = (round(waferMap.shape[1]*22.5/30), 
                                      round(waferMap.shape[0]*77/80))
            
            cv2.putText(tempWaferMap, 
                        "Failing Dies of Inlet: " + str(numFailingDies), 
                        bottomLeftCornerOfText, 
                        font, 
                        fontScale,
                        fontColor,
                        thickness,
                        lineType
                        )
        if isCompareMap:
            bottomLeftCornerOfText = (round(waferMap.shape[1]*22.5/30), 
                                      round(waferMap.shape[0]*79/80))
            cv2.putText(waferMap, 
                        "Failing Dies of Outlet: " + str(numFailingDies), 
                        bottomLeftCornerOfText, 
                        font, 
                        fontScale,
                        fontColor,
                        thickness,
                        lineType
                        )
        else:
            bottomLeftCornerOfText = (round(waferMap.shape[1]*25/30), 
                                      round(waferMap.shape[0]*79/80))
            cv2.putText(waferMap, 
                        "Failing Dies: " + str(numFailingDies), 
                        bottomLeftCornerOfText, 
                        font, 
                        fontScale,
                        fontColor,
                        thickness,
                        lineType
                        )
        
        print("   Saving wafer map..")
        # Wafer map size limit set
        image_size_limit = WAFER_MAP_SIZE_LIMIT # in mb
        image_size_limit = image_size_limit * 1000000 # now in bytes
        # Saves Wafer Map and deletes Temp Wafer Map if needed
        percent_knockoff = 5
        while True:
            cv2.imwrite(slotPath + "/Wafer_Map_with_Failing_Dies.jpg", 
                        waferMap, 
                        [cv2.IMWRITE_JPEG_QUALITY, 100-percent_knockoff])
            size = os.path.getsize(slotPath + "/Wafer_Map_with_Failing_Dies.jpg")
            if size > image_size_limit:
                os.remove(slotPath + "/Wafer_Map_with_Failing_Dies.jpg")
                percent_knockoff += 10
            else:
                break
        if isInletLot:
            percent_knockoff = 5
            while True:
                cv2.imwrite(slotPath + "/Temp_Wafer_Map_to_Compare.jpg", 
                            tempWaferMap, 
                            [cv2.IMWRITE_JPEG_QUALITY, 100-percent_knockoff])
                size = os.path.getsize(slotPath + "/Temp_Wafer_Map_to_Compare.jpg")
                if size > image_size_limit:
                    os.remove(slotPath + "/Temp_Wafer_Map_to_Compare.jpg")
                    percent_knockoff += 10
                else:
                    break
        if isCompareMap:
            if os.path.isfile(inletSlotPath + "/Temp_Wafer_Map_to_Compare.jpg"):
                os.remove(inletSlotPath + "/Temp_Wafer_Map_to_Compare.jpg")

        # XLS Section
        # -----------------------------------------------------------------------------
        if EXCEL_GENERATOR_TOGGLE and os.path.isfile(slotPath + "/Wafer_Map_with_Failing_Dies.jpg"):
            print("   Starting Excel sheet results..")
            # Create a workbook and add a worksheet.
            workbook = xlsxwriter.Workbook(slotPath + '/Results.xlsx')
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
            bin2_background = workbook.add_format({'bg_color': 'red'})
            bin2_bold_background = workbook.add_format({'bold': True, 
                                                        'bg_color': 'red'})
            # bin 3 - dark green or green
            bin3_background = workbook.add_format({'bg_color': 'green'})
            bin3_bold_background = workbook.add_format({'bold': True, 
                                                        'bg_color': 'green'})
            # bin 4 - yellow
            bin4_background = workbook.add_format({'bg_color': 'yellow'})
            bin4_bold_background = workbook.add_format({'bold': True, 
                                                        'bg_color': 'yellow'})
            # bin 5 - blue
            bin5_background = workbook.add_format({'bg_color': 'blue'})
            bin5_bold_background = workbook.add_format({'bold': True, 
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
            bin8_background = workbook.add_format({'bg_color': 'gray'})
            bin8_bold_background = workbook.add_format({'bold': True, 
                                                        'bg_color': 'gray'})
            
            
            # Iterates over each 200x200 dies and defaults bin number to 8
            for row in range(200):
                for col in range(200):
                    worksheet_TL.write(row, col, 8, bin8_background)
                    worksheet_TR.write(row, col, 8, bin8_background)
                    worksheet_BL.write(row, col, 8, bin8_background)
                    worksheet_BR.write(row, col, 8, bin8_background)
            
            
            # Combines all die names and bin numbers
            all_dieNames = badDieNames
            all_dieBinNumbers = badDieBinNumbers
            print("   Started making good bins..")
            list = os.listdir(glob.glob(slotPath + "/*")[1])
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
                
                # Checks to see which background bin number to use
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
                
                if row <= 200:
                    if col <= 200:
                        worksheet_TL.write(row-1, col-1, 
                                           all_dieBinNumbers[all_dieName_index], 
                                           background)
                    else:
                        worksheet_TR.write(row-1, col-1-200, 
                                           all_dieBinNumbers[all_dieName_index],
                                           background)
                else:
                    if col <= 200:
                        worksheet_BL.write(row-1-200, col-1, 
                                           all_dieBinNumbers[all_dieName_index],
                                           background)
                    else:
                        worksheet_BR.write(row-1-200, col-1-200, 
                                           all_dieBinNumbers[all_dieName_index],
                                           background)
            
            # Write a count for each bin at the bottom
            for worksheet_name in worksheet_list:
                worksheet_name.write(202, 0, "0-Bad-Count", 
                                     bin0_bold_background)
                worksheet_name.write(202, 1, "", bin0_background)
                worksheet_name.write(202, 2, "", bin0_background)
                worksheet_name.write(202, 3, "", bin0_background)
                worksheet_name.write(202, 4, all_dieBinNumbers.count(0),
                                     bin0_bold_background)
                
                worksheet_name.write(203, 0, "1-All_Good-Count", 
                                     bin1_bold_background)
                worksheet_name.write(203, 1, "", bin1_background)
                worksheet_name.write(203, 2, "", bin1_background)
                worksheet_name.write(203, 3, "", bin1_background)
                worksheet_name.write(203, 4, all_dieBinNumbers.count(1),
                                     bin1_bold_background)
                
                worksheet_name.write(204, 0, "2-Red_Only_Present-Count", 
                                     bin2_bold_background)
                worksheet_name.write(204, 1, "", bin2_background)
                worksheet_name.write(204, 2, "", bin2_background)
                worksheet_name.write(204, 3, "", bin2_background)
                worksheet_name.write(204, 4, all_dieBinNumbers.count(2),
                                     bin2_bold_background)
                
                worksheet_name.write(205, 0, "3-Green_Only_Present", 
                                     bin3_bold_background)
                worksheet_name.write(205, 1, "", bin3_background)
                worksheet_name.write(205, 2, "", bin3_background)
                worksheet_name.write(205, 3, "", bin3_background)
                worksheet_name.write(205, 4, all_dieBinNumbers.count(3),
                                     bin3_bold_background)
                
                worksheet_name.write(206, 0, "4-Red_Green_Only_Present-Count", 
                                     bin4_bold_background)
                worksheet_name.write(206, 1, "", bin4_background)
                worksheet_name.write(206, 2, "", bin4_background)
                worksheet_name.write(206, 3, "", bin4_background)
                worksheet_name.write(206, 4, all_dieBinNumbers.count(4),
                                     bin4_bold_background)
                
                worksheet_name.write(207, 0, "5-Blue_Only_Present", 
                                     bin5_bold_background)
                worksheet_name.write(207, 1, "", bin5_background)
                worksheet_name.write(207, 2, "", bin5_background)
                worksheet_name.write(207, 3, "", bin5_background)
                worksheet_name.write(207, 4, all_dieBinNumbers.count(5),
                                     bin5_bold_background)
                
                worksheet_name.write(208, 0, "6-Red_Blue_Only_Present-Count", 
                                     bin6_bold_background)
                worksheet_name.write(208, 1, "", bin6_background)
                worksheet_name.write(208, 2, "", bin6_background)
                worksheet_name.write(208, 3, "", bin6_background)
                worksheet_name.write(208, 4, all_dieBinNumbers.count(6),
                                     bin6_bold_background)
                
                worksheet_name.write(209, 0, "7-Green_Blue_Only_Present-Count", 
                                     bin7_bold_background)
                worksheet_name.write(209, 1, "", bin7_background)
                worksheet_name.write(209, 2, "", bin7_background)
                worksheet_name.write(209, 3, "", bin7_background)
                worksheet_name.write(209, 4, all_dieBinNumbers.count(7),
                                     bin7_bold_background)
                
                worksheet_name.write(210, 0, "8-Not_Tested-Count", 
                                     bin8_bold_background)
                worksheet_name.write(210, 1, "", bin8_background)
                worksheet_name.write(210, 2, "", bin8_background)
                worksheet_name.write(210, 3, "", bin8_background)
                worksheet_name.write(210, 4, "=160000-sum(E203:E210)",
                                     bin8_bold_background)
                
                
                # worksheet_name.set_column(0, 0, width=len("7-Green_Blue_Only_Present-Count"))
                worksheet_name.set_column(4, 4, width=7)
            
            workbook.close()
        # -----------------------------------------------------------------------------

print("Done!")

# Stopping stopwatch to see how long process takes
end_time = time.time()
time_lapsed = end_time - start_time
time_convert(time_lapsed)