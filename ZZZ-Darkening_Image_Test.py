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
MATCH_CL = 0.70 # Minimum confidence level (CL) required to match golden-image to scanned image
DIE_SPACING = 1.03 # Scale of die to die plus spacing between die
STICHED_IMAGES_DIRECTORY = "./Images/000-Stitched_Images/"
GOLDEN_IMAGES_DIRECTORY = "./Images/001-Golden_Images/"
WAFER_MAP_DIRECTORY = "./Images/002-Wafer_Map/"
SLEEP_TIME = 0.0 # Time to sleep in seconds between each window step
TOGGLE_DELETE_WAFER_MAP = False
TOGGLE_SHOW_WINDOW_IMAGE = False # Set equal to "True" and it will show a graphical image of where it's at


def time_convert(sec):
    mins = sec // 60
    sec = sec % 60
    hours = mins // 60
    mins = mins % 60
    print("Time Lapsed = {0}h:{1}m:{2}s".format(int(hours), int(mins), round(sec) ) )


def decrease_brightness(img, value):
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
    h, s, v = cv2.split(hsv)
    
    v[v < 25] = 0
    v = np.uint8(v*0.3)

    final_hsv = cv2.merge((h, s, v))
    img = cv2.cvtColor(final_hsv, cv2.COLOR_HSV2BGR)
    return img



# MAIN():
# =============================================================================
# Starting stopwatch to see how long process takes
start_time = time.time()

# Clears some of the screen for asthetics
print("\n\n\n\n\n\n\n\n\n\n\n\n\n")


lenStitchDir = len(STICHED_IMAGES_DIRECTORY)


# Runs through each slot file within the main file within stitched-image folder
for stitchFolderPath in glob.glob(STICHED_IMAGES_DIRECTORY + "*"): 
    print("Working on", stitchFolderPath[len(STICHED_IMAGES_DIRECTORY):])
    
    # If wafermap already exist (and delete wafer map directory didn't happen)
    # then skip creating a wafer map for this stitch-image folder path
    doesWaferMapExist = os.path.isdir(WAFER_MAP_DIRECTORY + \
                                      stitchFolderPath[len(STICHED_IMAGES_DIRECTORY):])
    
    doesWaferMapExist = False
    
    if doesWaferMapExist == True:
        print(" ", stitchFolderPath[len(STICHED_IMAGES_DIRECTORY):], "exists already. Skipping.")
        continue
    
    stitchImagePath = glob.glob(stitchFolderPath + "/*")[0]
    # goldenFolderPath = GOLDEN_IMAGES_DIRECTORY + stitchFolderPath[lenStitchDir:]
    # goldenImagePath = glob.glob(goldenFolderPath + "/*")[0]
    
    os.makedirs(WAFER_MAP_DIRECTORY + stitchFolderPath[lenStitchDir:], 
                exist_ok=True)
    
    # Load images
    stitchImage = cv2.imread(stitchImagePath)
    # goldenImage = cv2.imread(goldenImagePath)
    
    
    stitched_image_dark_mode = decrease_brightness(stitchImage, value=100)
    cv2.imwrite(stitchImagePath[:-4] + "-Darkened" + stitchImagePath[-4:], stitched_image_dark_mode)
    