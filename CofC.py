# Script that will make a copy of a template to another directory for archive and Edit.
# Leaving template unchanged.

# import necessary libraries 
import os
import datetime
import openpyxl
import shutil
import pyautogui
import time

# This finds the absolute path of the directory
THIS_FOLDER = os.path.dirname(os.path.abspath(__file__))

#link to template excel file
exfile = os.path.join(THIS_FOLDER, 'C of C NEW 2018 (version 1).xlsm')

# This makes a copy of the excel file named year, month, day and minute 
def createNew() :
    global newFileName
    global today
   
    today = datetime.datetime.now()
    newFileName ='C of C New ' + today.strftime("%Y_%m%d%M") + ".xlsm"
    os.chdir(THIS_FOLDER + '\\Saved\\')
    shutil.copyfile(exfile, newFileName)

# starts excel and open new editable file 
def startProg(var) :
    os.startfile(var)

# loads the function to convert list of scans into formatted printable item   
def loadxcl() :
   
    wb = openpyxl.load_workbook(newFileName, read_only=False, keep_vba=True)
    sheet = wb['Scan']
    sheet2 = wb['CofC']

    #function to remove prefixes or suffixes and check for blank cells
    def splitStr(a, b, c) :
        if sheet[a].value is None :
            prefixStr = " "
        else:
            prefixStr = sheet[a].value.split()
            sheet2[b] = prefixStr[c]
        

    # adds model number based on first program
    splitStr('A1','D6',0)
    sheet2['J6']= today.strftime("%m/%d/%Y")
    
    # for loop to get data in list on scan page 
    for x in range(50) :
        y = 'A' + str(x+1)
           
        if x < 20 :
            d = 'B' + str(x + 10)
            splitStr(y,d,1)
        
            

        elif x >= 20 and x < 40 :
            d = 'F' + str(x - 10)
            splitStr(y,d,1)
        
            

        else: 
            d = 'J' + str(x - 30)
            splitStr(y,d,1)

    #saves the file        
    wb.save(newFileName)

#function to print the page
def printxcl() :
    
    #function for multi button press
    def mpress(x,y) :
        pyautogui.keyDown(x)
        pyautogui.press(y)
        pyautogui.keyUp(x)
        time.sleep(2)
        return
    #macro start to print the page by utilizing built in excel workbook macro
    time.sleep(5)
    mpress('ctrl','t')
    time.sleep(10)
    mpress('ctrl', 'p')
    pyautogui.press('enter')
    time.sleep(5)
    mpress('alt','f4')
    pyautogui.press('enter')

# process of functions
createNew()
startProg(newFileName)
pausesys = input("Press Enter after you finish scanning to continue")
loadxcl()
startProg(newFileName)
printxcl()


    