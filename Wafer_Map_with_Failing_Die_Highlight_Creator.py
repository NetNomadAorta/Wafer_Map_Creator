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

# User Parameters/Constants to Set
PREDICTED_DIR = "//mcrtp-file-01.mcusa.local/public/000-AOI_Tool_Output/"
STORED_WAFER_DATA = "//mcrtp-file-01.mcusa.local/public/000-AOI_Tool_Output/ZZZ-General_Wafer_Map_Data/"
COMPARE_OVERLAY = True # Will compare "*-In" and "*-Out" wafer maps and output in "*-Out" folder


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
    
    # Sets parameter to enable comparing in and out files
    isInletLot = False
    isCompareMap = False
    if "-In" in lotPath:
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
        if waferMapName in os.listdir(PREDICTED_DIR):
            shouldContinue = False
            break
    if shouldContinue:
        continue
    
    # Imports correct dieNames and dieCoordinates data
    dieNames = np.load(STORED_WAFER_DATA + waferMapName + "/dieNames.npy")
    dieCoordinates = np.load(STORED_WAFER_DATA + waferMapName + "/Coordinates.npy")
    
    # Creates wafer map
    waferMap = cv2.imread(STORED_WAFER_DATA + waferMapName +  "/Wafer_Map.jpg")
    
    
    # Removes Thumbs.db in lot path if found
    if os.path.isfile(lotPath + "/Thumbs.db"):
        os.remove(lotPath + "/Thumbs.db")
    
    # Cycles through each slot folder within the lot folder
    for slotPath in glob.glob(lotPath + "/*"):
        
        if isInletLot:
            tempWaferMap = waferMap.copy()
        elif isCompareMap:
            inletSlotPath = slotPath.replace("-Out", "-In")
            if os.path.isfile(inletSlotPath + "/Temp_Wafer_Map_to_Compare.jpg"):
                waferMap = cv2.imread(inletSlotPath + "/Temp_Wafer_Map_to_Compare.jpg")
        
        # Removes Thumbs.db in slot path if found
        if os.path.isfile(slotPath + "/Thumbs.db"):
            os.remove(slotPath + "/Thumbs.db")
        
        # Start count for failing dies
        numFailingDies = 0
        
        # Within each slot, cycle through each class
        for classIndex, classPath in enumerate(glob.glob(slotPath + "/*") ):
            # Skips directory if first class (non-defect) folder or if it 
            # includes the wafer map with failing dies image (if this program 
            # already created one from a previous run)
            if classIndex == 0 \
            or os.listdir(slotPath)[classIndex] == "Wafer_Map_with_Failing_Dies.jpg"\
            or os.listdir(slotPath)[classIndex] == "Temp_Wafer_Map_to_Compare.jpg":
                # ABOVE LAST LINE AFTER OR STATEMENT MIGHT BE REDUNDANT. Should I remove?
                continue
            # Removes Thumbs.db in class path if found
            if os.path.isfile(classPath + "/Thumbs.db"):
                os.remove(classPath + "/Thumbs.db")
            
            # Looks at die names in previously created wafer map and sees
            #  if this slot has the same die names within its defect folders.
            # If so, then create the new wafer map with red ovals in die 
            #  location within the wafer map image, and save this image.
            for dieNameIndex, dieName in enumerate(dieNames):
                isBadDie = False
                for imageName in os.listdir(classPath):
                    if dieName in imageName:
                        isBadDie = True
                    else:
                        isBadDie = False
                        
                    x1 = dieCoordinates[dieNameIndex][0]
                    y1 = dieCoordinates[dieNameIndex][1]
                    x2 = dieCoordinates[dieNameIndex][2]
                    y2 = dieCoordinates[dieNameIndex][3]
                    
                    midX    = round( (x1 + x2)/2)
                    midY    = round( (y1 + y2)/2)
                    
                    lengthX = round( ( (x2 - x1) * 0.95 )/2)
                    lengthY = round( ( (y2 - y1) * 0.95 )/2)
                    
                    lengthXInner = round( ( (x2 - x1) * 0.85 )/2)
                    lengthYInner = round( ( (y2 - y1) * 0.85 )/2)
                    
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
                    thickness   = round(waferMap.shape[0] * 0.0009)
                    
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
                    
                    if isBadDie:
                        numFailingDies += 1
                        break
                    
                if isBadDie:
                    continue
            
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
        
        
        if not isCompareMap:
            cv2.putText(waferMap, 
                        "Slot Name: " + str(slotName), 
                        bottomLeftCornerOfText, 
                        font, 
                        fontScale,
                        fontColor,
                        thickness,
                        lineType
                        )
        
        # Also writes name in temporary wafer map if available
        if isInletLot:
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
                                  round(waferMap.shape[0]*1/35))
        fontScale              = round(0.0004*waferMap.shape[1], 2)
        fontColor              = (255, 255, 255)
        thickness              = round(0.0013*waferMap.shape[1])
        lineType               = 2
        
        if not isCompareMap:
            cv2.putText(waferMap, 
                        "Green: Passing; Red: Failing\n"\
                        + "Inner Circle: Incoming Wafer\n"\
                        + "Outer Circle: Final Wafer", 
                        bottomLeftCornerOfText, 
                        font, 
                        fontScale,
                        fontColor,
                        thickness,
                        lineType
                        )
        
        # Also writes name in temporary wafer map if available
        if isInletLot:
            cv2.putText(tempWaferMap, 
                        "Green: Passing; Red: Failing\n"\
                        + "Inner Circle: Incoming Wafer\n"\
                        + "Outer Circle: Final Wafer", 
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
            bottomLeftCornerOfText = (round(waferMap.shape[1]*23/30), 
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
            cv2.putText(waferMap, 
                        "Failing Dies: " + str(numFailingDies), 
                        bottomLeftCornerOfText, 
                        font, 
                        fontScale,
                        fontColor,
                        thickness,
                        lineType
                        )
        
        # Saves Wafer Map and deletes Temp Wafer Map if needed
        cv2.imwrite(slotPath + "/Wafer_Map_with_Failing_Dies.jpg", waferMap)
        if isInletLot:
            cv2.imwrite(slotPath + "/Temp_Wafer_Map_to_Compare.jpg", tempWaferMap)
        if isCompareMap:
            if os.path.isfile(inletSlotPath + "/Temp_Wafer_Map_to_Compare.jpg"):
                os.remove(inletSlotPath + "/Temp_Wafer_Map_to_Compare.jpg")



print("Done!")

# Stopping stopwatch to see how long process takes
end_time = time.time()
time_lapsed = end_time - start_time
time_convert(time_lapsed)