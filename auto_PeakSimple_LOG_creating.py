# This script was created to automate the task of creating a single PeakSimple LOG file across multiple CHR files
# Instruments: GC-FID
# Software: works with Peak Simple
# Authors: Davi de Ferreyro Monticelli, iREACH group (University of British Columbia)
# Date: 2024-04-22
# Version: 1.0.0

# How it works: With PeakSimple opened in your computer, run this script after specifying the files you'll
#               work with in the lists below. !!! Within 10s after running, open the PeakSimple software window !!!
#               The script will assume control of your mouse and keyboard and open the files for you.
#               Once the files are opened, you can modify it (baseline correction, calibration etc.), and
#               once you are satisfied just press 'r' and the script will save the results to a PeakSimple
#               LOG file that can be later post-processed.

#######################################################################################################################
# NOTE: You'll need to have all you .CPT, .CHR, .CON, .CAL files in a single folder for this to work
# NOTE: Your CONTROL file should be already set with your preferences
#######################################################################################################################

import time
from pynput.mouse import Controller
from pynput.mouse import Button as Button_pyn
import pyautogui
import keyboard

mouse = Controller()

# Define the list of CHR files to open:
chr_files_1min = [
    "Terpenes_Sampling_1min_23_07_17.CHR",
    "Terpenes_Sampling_1min_23_10_18.CHR"
]

# Define the list of CPT files to open:
cpt_files_1min = [
    "Terpenes_Sampling_1min_23_07_17.CPT",
    "Terpenes_Sampling_1min_23_10_18.CPT"
]

# Define function to open CHR and CPT files
def open_files(chr_file, cpt_file):
    # Open the CHR file
    mouse.position = (38, 66)
    mouse.press(Button_pyn.left)
    time.sleep(0.2)
    mouse.release(Button_pyn.left)
    time.sleep(2)
    pyautogui.typewrite(chr_file)
    time.sleep(2)
    pyautogui.press('enter')
    time.sleep(2)

    # Open the CPT file
    mouse.position = (141, 66)
    mouse.press(Button_pyn.left)
    time.sleep(0.2)
    mouse.release(Button_pyn.left)
    time.sleep(2)
    mouse.position = (792, 425)
    mouse.press(Button_pyn.left)
    time.sleep(0.2)
    mouse.release(Button_pyn.left)
    time.sleep(2)
    mouse.position = (768, 705)
    mouse.press(Button_pyn.left)
    time.sleep(0.2)
    mouse.release(Button_pyn.left)
    time.sleep(2)
    pyautogui.typewrite(cpt_file)
    time.sleep(2)
    pyautogui.press('enter')
    time.sleep(2)

    # Close open tabs
    mouse.position = (960, 748)
    mouse.press(Button_pyn.left)
    time.sleep(0.2)
    mouse.release(Button_pyn.left)
    time.sleep(2)
    mouse.position = (897, 679)
    mouse.press(Button_pyn.left)
    time.sleep(0.2)
    mouse.release(Button_pyn.left)
    time.sleep(2)

    # Wait for user to press R to proceed
    print("Waiting for script to resume after checking chromatograph results and modifications.")
    keyboard.wait('r')
    print("Resuming.")

    # Results, add to log
    mouse.position = (97, 38)
    mouse.press(Button_pyn.left)
    time.sleep(0.2)
    mouse.release(Button_pyn.left)
    time.sleep(2)
    mouse.position = (129, 68)
    mouse.press(Button_pyn.left)
    time.sleep(0.2)
    mouse.release(Button_pyn.left)
    time.sleep(2)
    mouse.position = (719, 714)
    mouse.press(Button_pyn.left)
    time.sleep(0.2)
    mouse.release(Button_pyn.left)
    time.sleep(2)
    mouse.position = (1007, 158)
    mouse.press(Button_pyn.left)
    time.sleep(0.2)
    mouse.release(Button_pyn.left)
    time.sleep(2)


# Time to start
time.sleep(10)

# Loop to operate PeakSimple:
for chr_file, cpt_file in zip(chr_files_1min, cpt_files_1min):
    open_files(chr_file, cpt_file)

print("Script completed.")
