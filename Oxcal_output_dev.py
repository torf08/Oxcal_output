#!/usr/bin/env python3

"""
This script is designed to use json strings generated from the javascript files produced by the Oxcal 14C calibration program. See -- https://c14.arch.ox.ac.uk/oxcalhelp/readme.html#local for instructions on local or server installations. 

It is intended to be run using python3 with tcl-tk built-in. See -- https://stackoverflow.com/questions/36760839/why-my-python-installed-via-home-brew-not-include-tkinter for information on how to install python3 with tcl-tk on Mac and Linux machines.

Also needed is the xlsxwriter package for python3 for writing to excel spreadsheets. This package can be installed through 'pip install xlsxwriter'. See https://xlsxwriter.readthedocs.io/ for full documentation.

Feature currently commented out * (Lastly, it needs argparse to pass in and parse command line arguments chosen when initally running the script.)

This script can accommodate both Unmodeled Calibrated and Bayesian Modeled Calibrated .json files. When running from the command line include either 'Yes' or 'No' depending on whether or not the .json file includes Bayesian posterior information. 

It will clean and seperate names, RYCBP, sigma ranges, percentages, and median dates in one worksheet. It will also seperate probability density values and increment years from the start date of the range for making Probability Density Function (PDF) graphs. It will also capture any boundary information from the Bayesian model, but skips over sequences and phase sections in the indexes. This will likely change at a later date for ease of publication.

Planned improvements include writing a GUI window with tkinter and tk for all aspects of the file and Bayesian boolean selection process. 

Contact Matthew A. Fort (fort1@illinois.edu) with any questions, concerns or ideas for possible improvements 
""" 

#import tkinter for gui uses, JSON parser to load JSON string, and xlsxwriter to create and write to an excelsheet
from tkinter import Tk
from tkinter import filedialog
import json
from Define_Functions import Non_Bayesian_Workbook, Bayesian_Workbook 
#import argparse


#hide the main Tk window
root = Tk()
root.withdraw()

'''
Will have to come back to play with this more. 

#Command-line argument parser for asking whether .json file has Bayesian posterior info
parser = argparse.ArgumentParser(description = 'State what type of Oxcal .json file is')

parser.add_argument('Bayesian', metavar = 'Yes or No', type = str, nargs = '?', default = 'No', help ='Does the Oxcal .json file include Bayesian modeled ranges? Default response is No')

'''

#Simpler way to prompt user to determine which method to use from Define_Functions
Bayesian = int(input("Does this file contain Bayesian model data? 1 for Yes or 0 for No: "))

#What age scale should be used? 
Age_Scale = int(input("What age scale should be used? 1 for BC/AD, 2 for BCE/CE or 3 for Cal BP: "))

#ask for and set .JSON file to variable json_filename & ask for and save excel file name and location
json_filename=filedialog.askopenfile(initialdir ="~/", title = "Select .json file", filetypes = (("json files", "*.json"),("all files","*.*")))

excel_filename = filedialog.asksaveasfilename(initialdir ="~/", title = "Save to .xlsx file", filetypes = (("Excel Workbook", "*.xlsx"),("Excel Workbook", "*.xls"),("all files","*.*")))

#Check that both json and excel file are chosen properly
if json_filename == None:
    print ("No file selected for open file. Should be a .json")
    quit()
else:
    x = (json_filename.name)
    json_name_split = x.split(".")
    
    if 'json' in json_name_split:
        print ("Opening read file")
        Oxcal_Data=json.load(json_filename)
        json_filename.close()
        print (type(Oxcal_Data))
    else:
        print ("The selected file must be a .json file extension!")
        quit()
  
if excel_filename == None:
    print ("No file selected for save file. Should be a .xlsx or .xls")
    quit()
else:
    excel_name_split = (excel_filename.split("."))
    if 'xlsx' in excel_name_split or 'xls' in excel_name_split:
        print("Opening save file")
    else:
        print ("The save file must be a .xlsx or .xls file etension!")
        quit()
        
#Use user input to determine function to use 
if Bayesian == 1:
    Bayesian_Workbook(excel_filename, Oxcal_Data, Age_Scale)
    print("Opening Bayesian Output")
    
elif Bayesian == 0:
    Non_Bayesian_Workbook(excel_filename, Oxcal_Data, Age_Scale)
    print("Opening Non-Bayesian Output")
else:
    print("Only 1 or 0 are accepted responses!")


'''
#Check what response was parsed through from the command line and direct towards the appropriate function
args = parser.parse_args()

if args.Bayesian[0] == 'Yes':
    Bayesian_Workbook(excel_filename, Oxcal_Data)
    
    print("Bayesian Workbook should be set up")

elif args.Bayesian[0] == 'No':
    Non_Bayesian_Workbook(excel_filename, Oxcal_Data)
    print("Non-Bayesian Workbook should be set up")
    
else:
    print("Only Yes or No are accpeted responses, what don't you get about that")
'''
