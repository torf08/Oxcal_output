#!/usr/bin/env python3

"""
This script is designed to use json strings generated from the javascript files produced by the Oxcal 14C calibration plugin for Firefox. 

It is intended to be run using python3 with tcl-tk built-in. See -- https://stackoverflow.com/questions/36760839/why-my-python-installed-via-home-brew-not-include-tkinter
for information on how to install python3 with tcl-tk on Mac and Linux machines.

Also needed is the xlsxwriter package for python3 for writing to excel spreadsheets. This package can be installed through 'pip install xlsxwriter'. See https://xlsxwriter.readthedocs.io/ for full documentation.

The script will clean and seperate names, RYCBP, sigma ranges, percentages, and median dates in one worksheet. It will also seperate probability density values and increment years from the start date of the range for the ease of making PDF curve graphs.

Contact Matthew A. Fort (fort1@illinois.edu) with any questions, concerns or ideas for possible improvements 
""" 

#import tkinter for gui uses, JSON parser to load JSON string, and xlsxwriter to create and write to an excelsheet
from tkinter import Tk
from tkinter import filedialog
import json
import xlsxwriter

#hide the main Tk window
root = Tk()
root.withdraw()

#ask for and equal .JSON file to variable json_filename & ask for and save excel file name and location
json_filename=filedialog.askopenfile(initialdir ="~/", title = "Select .json file", filetypes = (("json files", "*.json"),("all files","*.*")))

excel_filename = filedialog.asksaveasfilename(initialdir ="~/", title = "Save to .xlsx file", filetypes = (("Excel Workbook", "*.xlsx"),("Excel Workbook", "*.xls"),("all files","*.*")))

#Check that both json and excel file chosen properly
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
    else:
        print ("The opened file must be a .json file extension!")
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
          
#Create Woorkbooks for Oxcal_Data using xlsxwriter
workbook = xlsxwriter.Workbook(excel_filename)
Date_ranges = workbook.add_worksheet()
worksheet2 = workbook.add_worksheet()
PDF_Graphing = workbook.add_worksheet()


#Name varibles for various cell formats
header = workbook.add_format({'bold': True, 'center_across': True})
format = workbook.add_format({'num_format': '0', 'center_across': True})
center = workbook.add_format({'center_across': True})
percent = workbook.add_format({'num_format': '0.0%', 'center_across': True})
italics = workbook.add_format({'italic': True})

#Set row, col and row_adj variables and row_count list; set header labels; adjust column withds; and advance rows
sheet1_row1 = 0 
sheet1_row2 = 0
sheet1_col1 = 0
sheet1_row_adj = 0
sheet2_row1 = 0
sheet2_col1 = 0
sheet2_row_adj = 0
row_count = []
number_samples = 0

Labels = ['Name', 'RYCBP', 'Plus or Minus', '1s.d. Cal', '%', '2s.d. Cal', '%', 'Median', 'Plus or Minus', '','Posterior', '1s.d. Cal', '%', '2s.d. Cal', '%', 'Median', 'Plus or Minus', 'Agreement', 'Convergence', 'probNorm' ]
for count, name in enumerate(Labels):
    Date_ranges.write(sheet1_row1, sheet1_col1+count, name, header)
    
Date_ranges.set_column('D:D', 13)
Date_ranges.set_column('F:F', 13)

#Advance and create row varibles necessary for age ranges and percentages; create global c varible
sheet1_row1+=1
sheet1_row2+=1
sheet1_row3 = 1
sheet1_row4 = 1

#Apply AD or BC labels to median dates
def Medians(x, y, a, b):
    if x > 0: 
        AD_Median = ("AD " + str(int(x)))
        Date_ranges.write(sheet1_row1, sheet1_col1+a, AD_Median, center)
        Date_ranges.write(sheet1_row1, sheet1_col1+b, y, format)
    elif x < 0:
        BC_Median = ("BC " + str(int(abs(x))))
        Date_ranges.write(sheet1_row1, sheet1_col1+a, BC_Median, center)
        Date_ranges.write(sheet1_row1, sheet1_col1+b, y, format)
        
    return
    
#Seperate and write to spreadsheet probabilities
def Ranges(range, indpos1, indpos2, col1, col2, row):
    
    global c
    c = row
    
    for prob in range[indpos1:indpos2]:
        step_in_1a = prob
        #print (step_in_1a)
        for sets in step_in_1a:
            step_in_2a = sets
            step_in_3a = step_in_2a[0]
            step_in_3b = step_in_2a[1]
            step_in_3c = (step_in_2a[2]/100)
            
            if step_in_3a >= 0:
                AD_Date = ('A.D. ' + str(int(step_in_3a)) + '- A.D. ' +    
                str(int(step_in_3b)))
                Date_ranges.write(c, sheet1_col1+col1, AD_Date, center)
                
            elif step_in_3a <= 0 and step_in_3b <= 0:
                step_in_3a = abs(step_in_3a)
                step_in_3b = abs(step_in_3b)
                
                BC_Date = ('B.C. ' + str(int(step_in_3a)) + '- B.C. ' + 
                str(int(step_in_3b)))
                Date_ranges.write(c, sheet1_col1+col1, BC_Date, center)
                
            else:
                step_in_3a = abs(step_in_3a)
                
                BC_AD_Date = ('B.C. ' + str(int(step_ind_3a)) + '- A.D. '+ 
                str(int(step_in_3b)))
                Date_ranges.write(c, sheet1_col1+col1, BC_AD_Date, center)
            
            Date_ranges.write(c, sheet1_col1+col2, step_in_3c, percent)
            
            
            c += 1
            
        return

#Determine 
def Row_Shift (x, y):
    
    global rowshift1
    global rowshift2
    
    rowshift1 = x
    rowshift2 = y
    
    Row_Adj = rowshift1 - rowshift2
    
    if Row_Adj < 0:
        rowshift1 = rowshift1 + abs(Row_Adj) + 1
        rowshift2 += 1
    if Row_Adj > 0:
        rowshift2 = rowshift2 + abs(Row_Adj) + 1
        rowshift1 += 1
    if Row_Adj == 0:
        rowshift1 += 1
        rowshift2 += 1
        
    return

#Seperate indices dictonary types from Oxcal_Data list
for dict in Oxcal_Data[0:]:
    IndvData = dict
    
    list_liklihood = IndvData['likelihood']
    list_comment = list_liklihood['comment']
    
    #very janky way to get around varibles that don't exist in the first few items in the list
    if list_comment[0] == "OxCal v4.3.2 Bronk Ramsey (2017); r:5":
        continue
    
    list_op = IndvData['op']
    
    if list_op == "Sequence" or list_op == "Boundary":
        continue

    #Sample Name, RCYBP date, and error on RCYBP measurment
    list_name = IndvData['name']
    list_date = IndvData['date']
    list_error = IndvData['error']
    
    #Unmodeled Range, Median, Sigma, Probability, Range Start, and Resolution
    unmodeled_range = list_liklihood['range']
    unmodeled_median = list_liklihood['median']
    unmodeled_sigma = list_liklihood['sigma']
    list_prob = list_liklihood['prob']
    list_start = list_liklihood['start']
    list_res = list_liklihood['resolution']
    
    #Modeled Range, Median, Sigma, Probability, 
    list_posterior = IndvData['posterior']
    modeled_range = list_posterior['range']
    modeled_median = list_posterior['median']
    modeled_sigma = list_posterior['sigma']
    modeled_agreement = list_posterior['agreement']
    modeled_probNorm = list_posterior['probNorm']
    modeled_convergence = list_posterior['convergence']
    
    Date_ranges.write(sheet1_row1, sheet1_col1, list_name)
    Date_ranges.write(sheet1_row1, sheet1_col1+1, list_date, format)
    Date_ranges.write(sheet1_row1, sheet1_col1+2, list_error, format) 
    Date_ranges.write(sheet1_row1, sheet1_col1+17, modeled_agreement, format) #need another digit
    Date_ranges.write(sheet1_row1, sheet1_col1+18, modeled_convergence, format) #need another digit
    Date_ranges.write(sheet1_row1, sheet1_col1+19, modeled_probNorm, format) #need 7 digits
    
    #Writing Unmodeled and Modeled Medians and Sigmas
    Medians(unmodeled_median, unmodeled_sigma, 7, 8)
    Medians(modeled_median, modeled_sigma, 15, 16)
    
    #Writing Unmodeled for 1 and 2 sigma
    Ranges(unmodeled_range, 1, 2, 3, 4, sheet1_row1)
    sheet1_row1 = c
    Ranges(unmodeled_range, 2, 3, 5, 6, sheet1_row2)
    sheet1_row2 = c
    Ranges(modeled_range, 1, 2, 11, 12, sheet1_row3)
    sheet1_row3 = c
    Ranges(modeled_range, 2, 3, 13, 14, sheet1_row4)
    sheet1_row4 = c
    
    #Adjust sheet_row_num values to keep a consistent 1 row space between samples
    
    Row_Shift(sheet1_row1, sheet1_row2)
    sheet1_row1 = rowshift1
    sheet1_row2 = rowshift2
    
    Row_Shift(sheet1_row3, sheet1_row4)
    sheet1_row3 = rowshift1
    sheet1_row4 = rowshift2
    
    Row_adj = sheet1_row2 - sheet1_row3
    if Row_adj < 0:
        sheet1_row2 = sheet1_row2 + abs(Row_adj)
        sheet1_row1 = sheet1_row2
    if Row_adj > 0: 
        sheet1_row3 = sheet1_row3 + abs(Row_adj)
        sheet1_row4 = sheet1_row3
    if Row_adj == 0:
        sheet1_row2 = sheet1_row3

#Add reference to calibration software used
#print(sheet1_row1)

Date_ranges.write(sheet1_row1, sheet1_col1,((Oxcal_Data[0]['likelihood']['comment'][0]) + (Oxcal_Data[0]['likelihood']['comment'][1])), italics) 
                    
workbook.close()


        
'''     This is something to play around with, perhaps if I make this all a function I could include agruments about which prob density scheme to use.  
        #write probability density to worksheet2 and count number of rows
        for prob in list_prob:
            mod_prob = (prob/2) + number_samples
            worksheet2.write(sheet2_row1, sheet2_col1, mod_prob)
            sheet2_row1 += 1
            sheet2_row_adj += 1
            row_count.append(sheet2_row_adj)
'''