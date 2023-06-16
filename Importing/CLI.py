#This class provides the methods for cli elements
#Author: Caleb Harris - UCCS OIT Services Professional
#Date Created: 7/20/2022

import tkinter as tk
from tkinter import filedialog

def getFilePath():
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename()
    return file_path

def getSheetName():
    return input("Please enter the name of the sheet in excel to import from: ")

def cmdProgressBar(progress, total):

    '''
    This function displays a progress bar for the current operation in the command prompt
    Function written by NeuralNine on Youtube: https://www.youtube.com/watch?v=x1eaT88vJUA
    '''

    percent = 100 * (progress / float(total))
    bar = '█' * int(percent) + '-' * (100 - int(percent))
    print(f"\r|{bar}| {percent:.2f}", end="\r")

def getUserConfirm():
    '''
    Function get user confirmation for their selected option
    '''
    userVerfied = None
    while (userVerfied != True and userVerfied != False):
        
        confirmation = input("(y/n): ")
        
        if (confirmation == 'y' or confirmation == 'Y'):
            userVerfied = True
        elif (confirmation == 'n' or confirmation == 'N'):
            userVerfied = False
        else:
            print("Invalid selection.")

    return userVerfied

def welcome_user():
    #Welcome to user
    print(
    '''Welcome to the UCCS OIT Service Excel to MYSQL DB import tool!
-----------------------------------------------------------------
Author: Caleb Harris
Title: OIT Services Professional
Date Created: 5/11/2023
Last Date Modified: 5/25/2023
-----------------------------------------------------------------''')

