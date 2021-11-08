# Import the necessary packages
import os
import shutil
import glob
import imutils
import cv2
import time
import numpy as np

# User Parameters/Constants to Set
PREDICTED_DIR = "R:/000-AOI_Tool_Output/"


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

for lotPath in glob.glob(PREDICTED_DIR + "*"):
    
    slotLists = os.listdir(lotPath)
    if "Coordinates.npy" not in slotLists \
    and "dieNames.npy" not in slotLists \
    and "Wafer_Map.jpg" not in slotLists:
        continue
    dieNames = np.load(lotPath + "/dieNames.npy")
    dieCoordinates = np.load(lotPath + "/Coordinates.npy")
    waferMap = cv2.imread(lotPath + "/Wafer_Map.jpg")
    
    for slotPath in glob.glob(lotPath + "/*"):
        slotName = slotPath[len(lotPath)+1:]
        if "Coordinates.npy" in slotName \
        or "dieNames.npy" in slotName \
        or "Wafer_Map.jpg" in slotName:
            continue
        classIndex = 0
        for classPath in glob.glob(slotPath +"/*"):
            if classIndex == 0:
                classIndex += 1
                continue
            
            for dieNameIndex, dieName in enumerate(dieNames):
                for imageName in os.listdir(classPath):
                    if dieName in imageName:
                        print(dieName)
                        print(dieCoordinates[dieNameIndex])
                        
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
                                    3)
                        
                        break
            
            classIndex += 1
            
        cv2.imwrite(slotPath + "/Wafer_Map_with_Failing_Dies.jpg", waferMap)
















print("Done!")

# Stopping stopwatch to see how long process takes
end_time = time.time()
time_lapsed = end_time - start_time
time_convert(time_lapsed)