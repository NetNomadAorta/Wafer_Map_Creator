# Import the necessary packages
import os
import shutil
import glob
import imutils
import cv2
import time
import numpy as np
import math

# User Parameters/Constants to Set
MATCH_CL = 0.75 # Minimum confidence level (CL) required to match golden-image to scanned image
STICHED_IMAGES_DIRECTORY = "./Images/000-Stitched_Images/"
GOLDEN_IMAGES_DIRECTORY = "./Images/001-Golden_Images/"
WAFER_MAP_DIRECTORY = "./Images/002-Wafer_Map/"
SLEEP_TIME = 0.0 # Time to sleep in seconds between each window step
TOGGLE_DELETE_WAFER_MAP = False
TOGGLE_SHOW_WINDOW_IMAGE = False # Set equal to "True" and it will show a graphical image of where it's at
TOGGLE_STITCHED_OVERLAY = True # Will use original stitched image in final wafer map
DIE_SPACING_SCALE = 0.95

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


def decrease_brightness(img):
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
    h, s, v = cv2.split(hsv)
    
    brightness_scale = round((255-sum(cv2.mean(stitchImage))/3) * 0.00138, 2)
    
    v[v < 25] = 0
    v = np.uint8(v*brightness_scale)

    final_hsv = cv2.merge((h, s, v))
    img = cv2.cvtColor(final_hsv, cv2.COLOR_HSV2BGR)
    return img


def slidingWindow(stitchImage, stepSizeX, stepSizeY, windowSize):
    # Slides a window across the stitched-image
    for y in range(0, stitchImage.shape[0], stepSizeY):
        for x in range(0, stitchImage.shape[1], stepSizeX):
            # Yield the current window
            yield (x, y, stitchImage[y:y + windowSize[1], x:x + windowSize[0]])


# Comparison scan window-image to golden-image
def getMatch(window, goldenImage, x, y):
    h1, w1, c1 = window.shape
    h2, w2, c2 = goldenImage.shape
    
    if c1 == c2 and h2 <= h1 and w2 <= w1:
        method = eval('cv2.TM_CCOEFF_NORMED')
        res = cv2.matchTemplate(window, goldenImage, method)   
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
        
        if max_val > MATCH_CL: 
            print("\nFOUND MATCH - max_val =", round(max_val, 4) )
            print("Window Coordinates: x1:", x + max_loc[0], "y1:", y + max_loc[1], \
                  "x2:", x + max_loc[0] + w2, "y2:", y + max_loc[1] + h2)
            
            # Gets coordinates of cropped image
            return (max_loc[0], max_loc[1], max_loc[0] + w2, max_loc[1] + h2, max_val)
        
        else:
            return ("null", "null", "null", "null", "null")



# MAIN():
# =============================================================================
# Starting stopwatch to see how long process takes
start_time = time.time()

# Clears some of the screen for asthetics
print("\n\n\n\n\n\n\n\n\n\n\n\n\n")

# Deletes contents in cropped- and split-image folders
if TOGGLE_DELETE_WAFER_MAP == True:
    deleteDirContents(WAFER_MAP_DIRECTORY)

lenStitchDir = len(STICHED_IMAGES_DIRECTORY)


# Runs through each slot file within the main file within stitched-image folder
for stitchFolderPath in glob.glob(STICHED_IMAGES_DIRECTORY + "*"): 
    print("Working on", stitchFolderPath[len(STICHED_IMAGES_DIRECTORY):])
    
    # If wafermap already exist (and delete wafer map directory didn't happen)
    # then skip creating a wafer map for this stitch-image folder path
    doesWaferMapExist = os.path.isdir(WAFER_MAP_DIRECTORY + \
                                      stitchFolderPath[len(STICHED_IMAGES_DIRECTORY):])
    if doesWaferMapExist == True:
        print(" ", stitchFolderPath[len(STICHED_IMAGES_DIRECTORY):], "exists already. Skipping.")
        continue
    
    stitchImagePath = glob.glob(stitchFolderPath + "/*")[0]
    goldenFolderPath = GOLDEN_IMAGES_DIRECTORY + stitchFolderPath[lenStitchDir:]
    goldenImagePath = glob.glob(goldenFolderPath + "/*")[0]
    
    os.makedirs(WAFER_MAP_DIRECTORY + stitchFolderPath[lenStitchDir:], 
                exist_ok=True)
    
    # Load images
    stitchImage = cv2.imread(stitchImagePath)
    goldenImage = cv2.imread(goldenImagePath)
    
    # Parameter set
    winW = round(goldenImage.shape[1] * 1.5) # Scales window width with full image resolution
    winH = round(goldenImage.shape[0] * 1.5) # Scales window height with full image resolution
    windowSize = (winW, winH)
    stepSizeX = round(winW / 2.95)
    stepSizeY = round(winH / 2.95)
    
    # Predefine next for loop's parameters 
    prev_y1 = stepSizeY * 9 # Number that prevents y = 0 = prev_y1
    prev_x1 = stepSizeX * 9
    rowNum = 0
    colNum = 0
    prev_matchedCL = 0
    
    # Adding list and arrray entry
    dieNames = ["Row_#.Col_#"]
    dieCoordinates = np.zeros([1, 4], np.int32)
    
    # loop over the sliding window
    for (x, y, window) in slidingWindow(stitchImage, stepSizeX, stepSizeY, windowSize):
        # if the window does not meet our desired window size, ignore it
        if window.shape[0] != winH or window.shape[1] != winW:
            continue
        
        # Draw rectangle over sliding window for debugging and easier visual
        if TOGGLE_SHOW_WINDOW_IMAGE == True:
            displayImage = stitchImage.copy()
            cv2.rectangle(displayImage, 
                          (x, y), 
                          (x + winW, y + winH), 
                          (255, 0, 180), 
                          round(stitchImage.shape[0]*0.0027))
            displayImageResize = cv2.resize(displayImage, (1000, round(stitchImage.shape[0] / stitchImage.shape[1] * 1000)))
            cv2.imshow("Stitched Image", displayImageResize) # TOGGLE TO SHOW OR NOT
            cv2.waitKey(1)
            time.sleep(SLEEP_TIME) # sleep time in ms after each window step
        
        # Scans window for matched image
        # ==================================================================================
        # Scans window and grabs cropped image coordinates relative to window
        # Uses each golden image in the file if multiple part types are present
        
        # Gets coordinates relative to window of matched dies within a Stitched-Image
        win_x1, win_y1, win_x2, win_y2, matchedCL = getMatch(window, goldenImage, x, y)
        
        # Saves cropped image and names with coordinates
        if win_x1 != "null":
            # Turns cropped image coordinates relative to window to stitched-image coordinates
            x1 = x + win_x1
            y1 = y + win_y1
            x2 = x + win_x2
            y2 = y + win_y2
            
            # Makes sure same image does not get saved as different names
            if y1 >= (prev_y1 + round(stepSizeY / 2.95)) or y1 <= (prev_y1 - round(stepSizeY / 2.95)):
                rowNum += 1
                colNum = 1
                sameCol = False
            else:
                if x1 >= (prev_x1 + round(stepSizeX / 2.95)) or x1 <= (prev_x1 - round(stepSizeX / 2.95)):
                    colNum += 1
                    prev_matchedCL = 0
                    sameCol = False
                else: 
                    sameCol = True
            
            # NEEDS A CHECK TO SEE IF FIRST X IN PREVIOUS Y-ROW IS THE SAME
            #   IF IT ISN'T, THEN MAKE PREVIOUS FIRST X IN PREVIOUS ROW
            #   HAVE A COLUMN_NUMBER += 1 AND DELETE OLD SAVE AND RESAVE
            #   WITH NEW NAME
            
            # Puts 0 in front of single digit row nad column number
            if rowNum < 10:
                rZ = 0
            else: 
                rZ = ""
            if colNum < 10:
                cZ = 0
            else: 
                cZ = ""
            
            if sameCol == False: 
                dieNames.append("Row_{}{}.Col_{}{}".format(rZ, rowNum, cZ, colNum) )
                dieCoordinates = np.append(dieCoordinates, [[x1, y1, x2, y2]], axis=0)
            elif sameCol == True and matchedCL > prev_matchedCL:
                dieCoordinates[len(dieCoordinates)-1] = np.array([x1, y1, x2, y2], ndmin=2)
            
            prev_y1 = y1
            prev_x1 = x1
            prev_matchedCL = matchedCL
        # ==================================================================================
    rowNum += 1
    colNum = 0
    sameCol = False
    
    # Sets spacing between dies
    die_spacing_list = []
    for i in range(len(dieCoordinates)-2):
        die_spacing_temp = dieCoordinates[i+2, 0] - dieCoordinates[i+1, 2]
        if die_spacing_temp < (goldenImage.shape[1] * 0.6) and die_spacing_temp > 0:
            die_spacing_list.append(die_spacing_temp)

    die_spacing_max = max(die_spacing_list)
    die_spacing = 1 + round( (die_spacing_max/goldenImage.shape[1])*DIE_SPACING_SCALE, 3)
    
    # Grabbing max and min x and y coordinate values
    maxX = np.amax(dieCoordinates[:, 2] )
    maxY = np.amax(dieCoordinates[:, 3] )
    minX = np.amin(dieCoordinates[1:, 0] )
    minY = np.amin(dieCoordinates[1:, 1] )
    
    # Create blank image array
    waferMap = np.zeros([maxY+minY, maxX+minX, 3], np.uint8)
    
    # Overlays wafermap with darkened stitched image
    if TOGGLE_STITCHED_OVERLAY:
        dark_stitched_image = decrease_brightness(stitchImage)
        y_limit = dark_stitched_image[:maxY+minY, :maxX+minX].shape[0]
        x_limit = dark_stitched_image[:maxY+minY, :maxX+minX].shape[1]
        waferMap[:y_limit, :x_limit] = dark_stitched_image[:y_limit, :x_limit]
    
    # Creates wafer map and changes column names in dieNames
    for i in range(len(dieCoordinates)):
        # JUST ADDED TO SKIP FIRST ITERATION BECAUSE coord:0 and name:R#C#
        if i == 0:
            continue
        
        x1 = dieCoordinates[i, 0]
        y1 = dieCoordinates[i, 1]
        x2 = dieCoordinates[i, 2]
        y2 = dieCoordinates[i, 3]
        
        midX = round((x1 + x2)/2)
        midY = round((y1 + y2)/2)
        
        # Places green boxes over wafer map using each die's coordinate
        cv2.rectangle(waferMap, 
                      (x1, y1), 
                      (x2, y2), 
                      (255, 255, 255), 
                      round(goldenImage.size * 0.000002))
        
        # Replaces dieNames list column number with correct value
        colNumber = str(math.floor((x1-minX)/(goldenImage.shape[1]*die_spacing)+1) )
        if int(colNumber) < 10:
            colNumber = "0" + colNumber
        dieNames[i] = dieNames[i].replace("Col_" + dieNames[i][-2:], 
                                          "Col_" + str(colNumber))
        
        # Writes row and column number text in wafer map
        font                   = cv2.FONT_HERSHEY_SIMPLEX
        bottomLeftCornerOfText = (x1 + round(winW * 0.056), midY)
        fontScale              = round(winW*0.0023, 2)
        fontColor              = (255, 150, 150)
        thickness              = round(goldenImage.size * 0.0000025)
        lineType               = 2
        
        cv2.putText(waferMap, dieNames[i], 
                    bottomLeftCornerOfText, 
                    font, 
                    fontScale,
                    fontColor,
                    thickness,
                    lineType)
    
    cv2.imwrite(WAFER_MAP_DIRECTORY + stitchFolderPath[lenStitchDir:] \
                + "/Wafer_Map.jpg", waferMap)
    
    # Saving die names and coordinates as npy file
    np.save(WAFER_MAP_DIRECTORY + stitchFolderPath[lenStitchDir:] \
                + "/dieNames", dieNames)
    np.save(WAFER_MAP_DIRECTORY + stitchFolderPath[lenStitchDir:] \
                + "/Coordinates", dieCoordinates)
    # test1 = np.load(WAFER_MAP_DIRECTORY + stitchFolderPath[lenStitchDir:] \
    #             + "/dieNames.npy")
    # test2 = np.load(WAFER_MAP_DIRECTORY + stitchFolderPath[lenStitchDir:] \
    #             + "/Coordinates.npy")


print("Done!")

# Stopping stopwatch to see how long process takes
end_time = time.time()
time_lapsed = end_time - start_time
time_convert(time_lapsed)