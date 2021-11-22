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
    # If wafer map is not found in lot folder, then skip creating wafer map
    # with failing dies image
    slotLists = os.listdir(lotPath)
    if "Coordinates.npy" not in slotLists \
    and "dieNames.npy" not in slotLists \
    and "Wafer_Map.jpg" not in slotLists:
        continue
    
    dieNames = np.load(lotPath + "/dieNames.npy")
    dieCoordinates = np.load(lotPath + "/Coordinates.npy")
    
    # Removes Thumbs.db in lot path if found
    if os.path.isfile(lotPath + "/Thumbs.db"):
        os.remove(lotPath + "/Thumbs.db")
    
    # Cycles through each slot folder within the lot folder
    for slotPath in glob.glob(lotPath + "/*"):
        
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
            or os.listdir(slotPath)[classIndex] == "Wafer_Map_with_Failing_Dies.jpg":
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
                        
                        midX = round( (x1 + x2)/2)
                        midY = round( (y1 + y2)/2)
                        
                        lengthX = round( (x2 - x1)/2)
                        lengthY = round( (y2 - y1)/2)
                        
                        # Places red ovals over wafer map using bad die's coordinate
                        cv2.ellipse(waferMap, (midX, midY), 
                                    (lengthX, lengthY), 
                                    0, 
                                    0, 
                                    360, 
                                    (0, 0, 255),
                                    round(waferMap.shape[0] * 0.0009))
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
                    lineType)
        
        
        # Writes how many failing defects on bottom right
        font                   = cv2.FONT_HERSHEY_SIMPLEX
        bottomLeftCornerOfText = (round(waferMap.shape[1]*22/30), 
                                  round(waferMap.shape[0]*79/80))
        fontScale              = round(0.00055*waferMap.shape[1], 2)
        fontColor              = (0, 0, 255)
        thickness              = round(0.0014*waferMap.shape[1])
        lineType               = 2
        
        cv2.putText(waferMap, 
                    "Number of Failing Dies: " + str(numFailingDies), 
                    bottomLeftCornerOfText, 
                    font, 
                    fontScale,
                    fontColor,
                    thickness,
                    lineType)
        
        cv2.imwrite(slotPath + "/Wafer_Map_with_Failing_Dies.jpg", waferMap)
















print("Done!")

# Stopping stopwatch to see how long process takes
end_time = time.time()
time_lapsed = end_time - start_time
time_convert(time_lapsed)