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
STORED_WAFER_DATA = "C:/Users/ait.lab/.spyder-py3/Automated_AOI/Lot_Data/"
COMPARE_OVERLAY = False # Will compare "*-In" and "*-Out" wafer maps and output in "*-Out" folder
SHOULD_REPLACE_ALL_MAPS = False # Will remake each wafer map that already exist in AOI Output folder if set true
WAFER_MAP_SIZE_LIMIT = 300 # mb # If wafer map size above this value, reduce quality until size is under this value


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
            if (classIndex == 0
            or "ZZ-" in os.listdir(slotPath)[classIndex]
            or ".jpg" in os.listdir(slotPath)[classIndex]
            or ".xlsx" in os.listdir(slotPath)[classIndex]):
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
                percent_knockoff += 15
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
                    percent_knockoff += 15
                else:
                    break
        if isCompareMap:
            if os.path.isfile(inletSlotPath + "/Temp_Wafer_Map_to_Compare.jpg"):
                os.remove(inletSlotPath + "/Temp_Wafer_Map_to_Compare.jpg")


print("Done!")

# Stopping stopwatch to see how long process takes
end_time = time.time()
time_lapsed = end_time - start_time
time_convert(time_lapsed)