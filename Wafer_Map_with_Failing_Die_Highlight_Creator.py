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
import imutils
import cv2
import time
import numpy as np

# User Parameters/Constants to Set
PREDICTED_DIR = "//mcrtp-file-01.mcusa.local/public/000-AOI_Tool_Output/"
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
for lotPath in glob.glob(PREDICTED_DIR + "*"):
    
    # Sets parameter to enable comparing in and out files
    isInletLot = False
    compareMap = False
    if "-In" in lotPath:
        isInletLot = True
    if "-Out" in lotPath and COMPARE_OVERLAY == True:
        compareMap = True
        inletLotPath = lotPath.replace("-Out", "-In")
    
    # If wafer map is not found in lot folder, then skip creating wafer map
    # with failing dies image
    slotLists = os.listdir(lotPath)
    if "Coordinates.npy" not in slotLists \
    and "dieNames.npy" not in slotLists \
    and "Wafer_Map.jpg" not in slotLists:
        continue
    
    if compareMap:
        dieNames = np.load(inletLotPath + "/dieNames.npy")
        dieCoordinates = np.load(inletLotPath + "/Coordinates.npy")
    else:
        dieNames = np.load(lotPath + "/dieNames.npy")
        dieCoordinates = np.load(lotPath + "/Coordinates.npy")
    
    # Removes Thumbs.db in lot path if found
    if os.path.isfile(lotPath + "/Thumbs.db"):
        os.remove(lotPath + "/Thumbs.db")
    
    # Cycles through each slot folder within the lot folder
    for slotPath in glob.glob(lotPath + "/*"):
        
        if isInletLot:
            waferMap = cv2.imread(lotPath + "/Wafer_Map.jpg")
            tempWaferMap = waferMap.copy()
        elif compareMap:
            inletSlotPath = slotPath.replace("-Out", "-In")
            waferMap = cv2.imread(inletSlotPath + "/Temp_Wafer_Map_to_Compare.jpg")
        else:
            waferMap = cv2.imread(lotPath + "/Wafer_Map.jpg")
        
        # If current slotPath is looking at a non-slot folder, then skip
        slotName = slotPath[len(lotPath)+1:]
        if "Coordinates.npy" in slotName \
        or "dieNames.npy" in slotName \
        or "Wafer_Map.jpg" in slotName \
        or "Wafer_Map_with_Failing_Dies.jpg" in os.listdir(slotPath):
            continue
        
        # Removes Thumbs.db in slot path if found
        if os.path.isfile(slotPath + "/Thumbs.db"):
            os.remove(slotPath + "/Thumbs.db")
        
        # Start count for failing dies
        numFailingDies = 0
        
        # Within each slot, cycle through each class
        for classIndex, classPath in enumerate(glob.glob(slotPath + "/*")):
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
            # if this slot has the same die names within its defect folders.
            # If so, then create the new wafer map with red ovals in die 
            # location within the wafer map image, and save this image.
            for dieNameIndex, dieName in enumerate(dieNames):
                for imageName in os.listdir(classPath):
                    if dieName in imageName:
                        x1 = dieCoordinates[dieNameIndex][0]
                        y1 = dieCoordinates[dieNameIndex][1]
                        x2 = dieCoordinates[dieNameIndex][2]
                        y2 = dieCoordinates[dieNameIndex][3]
                        
                        midX    = round( (x1 + x2)/2)
                        midY    = round( (y1 + y2)/2)
                        
                        lengthX = round( (x2 - x1)/2)
                        lengthY = round( (y2 - y1)/2)
                        
                        # Places red ovals over wafer map using bad die's coordinate
                        center      = (midX, midY)
                        axes        = (lengthX, lengthY)
                        angle       = 0
                        startAngle  = 0
                        endAngle    = 360
                        color       = (0, 0, 255)
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
                            color = (0, 200, 150)
                            
                            cv2.ellipse(tempWaferMap, center, 
                                        axes, 
                                        angle, 
                                        startAngle, 
                                        endAngle, 
                                        color,
                                        thickness
                                        )
                        
                        numFailingDies += 1
                        
                        break
            
            classIndex += 1
        
        # Writes slot name on top left
        font                   = cv2.FONT_HERSHEY_SIMPLEX
        bottomLeftCornerOfText = (round(waferMap.shape[1]*1/80), 
                                  round(waferMap.shape[0]*1/35))
        fontScale              = round(0.0007*waferMap.shape[1], 2)
        fontColor              = (255, 255, 255)
        thickness              = round(0.0015*waferMap.shape[1])
        lineType               = 2
        
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
                          round(waferMap.shape[0]*75/80))
            
            cv2.putText(tempWaferMap, 
                        "Failing Dies of Inlet: " + str(numFailingDies), 
                        bottomLeftCornerOfText, 
                        font, 
                        fontScale,
                        fontColor,
                        thickness,
                        lineType
                        )
        elif compareMap:
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
        if compareMap:
            os.remove(inletSlotPath + "/Temp_Wafer_Map_to_Compare.jpg")















print("Done!")

# Stopping stopwatch to see how long process takes
end_time = time.time()
time_lapsed = end_time - start_time
time_convert(time_lapsed)