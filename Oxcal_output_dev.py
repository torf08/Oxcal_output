#!/usr/bin/env python3

"""
This script is designed to use json strings generated from the javascripts produced by the Oxcal 14C calibration plugin for Firefox. 

It is intended to be run using python3 with tcl-tk built-in. See -- https://stackoverflow.com/questions/36760839/why-my-python-installed-via-home-brew-not-include-tkinter
for information on how to install python3 with tcl-tk on Mac and Linux machines.

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

#ask for and equal .JSON file to variable json_filename and ask for and save excel file name and location
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
    print ("No file selected for save file. Should be a .xlsx")
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
worksheet = workbook.add_worksheet()
worksheet2 = workbook.add_worksheet()

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

Labels = ['Name', 'RYCBP', '±', '1s.d. Cal', '%', '2s.d. Cal', '%', 'Median', '±']
for count, name in enumerate(Labels):
    worksheet.write(sheet1_row1, sheet1_col1+count, name, header)
    
worksheet.set_column('D:D', 13)
worksheet.set_column('F:F', 13)

sheet1_row1+=1
sheet1_row2+=1


#Seperate indices dictonary types from Oxcal_Data list
for dict in Oxcal_Data[1:]:
    IndvData = dict

    list_name = IndvData['name']
    list_date = IndvData['date']
    list_error = IndvData['error']
    list_liklihood = IndvData['likelihood']
    list_range = list_liklihood['range']
    list_median = list_liklihood['median']
    list_sigma = list_liklihood['sigma']
    list_prob = list_liklihood['prob']
    list_start = list_liklihood['start']
    list_res = list_liklihood['resolution']
    
    worksheet.write(sheet1_row1, sheet1_col1, list_name)
    worksheet.write(sheet1_row1, sheet1_col1+1, list_date, format)
    worksheet.write(sheet1_row1, sheet1_col1+2, list_error, format)
    
    #Apply AD or BC labels to median dates
    if list_median > 0: 
        AD_Median = ("AD " + str(int(list_median)))
        worksheet.write(sheet1_row1, sheet1_col1+7, AD_Median, center)
        worksheet.write(sheet1_row1, sheet1_col1+8, list_sigma, format)
    elif list_median < 0:
        BC_Median = ("BC " + str(int(abs(list_median))))
        worksheet.write(sheet1_row1, sheet1_col1+7, BC_Median, center)
        worksheet.write(sheet1_row1, sheet1_col1+8, list_sigma, format)
    
    #Seperate and write to spreadsheet 1 sigma probs
    for prob in list_range[1:2]:
        step_in_1a = prob
        #print (step_in_1a)
        for sets in step_in_1a:
            step_in_2a = sets
            step_in_3a = step_in_2a[0]
            step_in_3b = step_in_2a[1]
            step_in_3c = (step_in_2a[2]/100)
            
            if step_in_3a >= 0:
                AD_Date = ('AD ' + str(int(step_in_3a)) + '-' + str(int(step_in_3b)))
                worksheet.write(sheet1_row1, sheet1_col1+3, AD_Date, center)
                
            elif step_in_3a <= 0 and step_in_3b <= 0:
                step_in_3a = abs(step_in_3a)
                step_in_3b = abs(step_in_3b)
                
                BC_Date = ('BC ' + str(int(step_in_3a)) + '-' + str(int(step_in_3b)))
                worksheet.write(sheet1_row1, sheet1_col1+3, BC_Date, center)
                
            else:
                step_in_3a = abs(step_in_3a)
                
                BC_AD_Date = ('BC ' + str(int(step_ind_3a)) + '-AD ', str(int(step_in_3b)))
                worksheet.write(sheet1_row1, sheet1_col1+3, BC_AD_Date, center)
            
            worksheet.write(sheet1_row1, sheet1_col1+4, step_in_3c, percent)
           
            sheet1_row1 += 1
            
    #Seperate and write to spreadsheet 2 sigma probs        
    for prob in list_range[2:3]:
        step_in_1b = prob
        #print (step_in_1b)
        for sets in step_in_1b:
            step_in_2b = sets
            
            step_in_3d = step_in_2b[0]
            step_in_3e = step_in_2b[1]
            step_in_3f = (step_in_2b[2]/100)
            
            if step_in_3d >0:
                AD_Date = ('AD ' + str(int(step_in_3d)) + '-' + str(int(step_in_3e)))
                worksheet.write(sheet1_row2, sheet1_col1+5, AD_Date, center)
                
            elif step_in_3d <= 0 and step_in_3e <= 0:
                step_in_3d = abs(step_in_3d)
                step_in_3e = abs(step_in_3e)
                
                BC_Date = ('BC ' + str(int(step_in_3d)) + '-' + str(int(step_in_3e)))
                worksheet.write(sheet1_row2, sheet1_col1+5, BC_Date, center)
                
            else:
                step_in_3d = abs(step_in_3d)
                
                BC_AD_Date = ('BC ' + str(int(step_ind_3d)) + '-AD ', str(int(step_in_3e)))
                worksheet.write(sheet1_row2, sheet1_col1+5, BC_AD_Date, center)
            
           
            worksheet.write(sheet1_row2, sheet1_col1+6, step_in_3f, percent)
            
            sheet1_row2 += 1
        
    #Adjust sheet1_row1 or sheet1_row2 to keep a consistent 1 row space between samples
    sheet1_row_adj = sheet1_row1 - sheet1_row2
    if sheet1_row_adj < 0:
        sheet1_row1 = sheet1_row1 + abs(sheet1_row_adj)+1
        sheet1_row2 += 1
    elif sheet1_row_adj > 0:
        sheet1_row2 = sheet1_row2 + abs(sheet1_row_adj)+1
        sheet1_row1 += 1
    elif sheet1_row_adj == 0:
        sheet1_row1 += 1
        sheet1_row2 += 1
    
    #Set headers for worksheet2 and advance row1
    header_dates = (list_name + ' dates')
    header_prob = (list_name + ' prob')
    
    worksheet2.write(sheet2_row1, sheet2_col1, header_dates, header)
    sheet2_col1 += 1
    worksheet2.write(sheet2_row1, sheet2_col1, header_prob, header)
    sheet2_row1 += 1
    
    #write probability density to worksheet2 and count number of rows
    for prob in list_prob:
        worksheet2.write(sheet2_row1, sheet2_col1, prob)
        sheet2_row1 += 1
        sheet2_row_adj += 1
        row_count.append(sheet2_row_adj)
        
    #reset worksheet2 row1 back to 1 and create seperate variable for list of start dates
    sheet2_row1 = 1
    start_date = list_start
    
    #write dates of probability density to worksheet2 and increment up the start dates by resolution (5 years)
    for count in row_count:
        worksheet2.write(sheet2_row1, sheet2_col1-1, start_date)
        start_date += 5
        sheet2_row1 += 1
    
    #reset worksheet2 columns, rows, and row_count to necessary values for restart of loop     
    sheet2_col1 += 1
    sheet2_row1 = 0
    sheet2_row_adj = 0
    row_count = []

#Add reference to calibration software used
worksheet.write(sheet1_row1, sheet1_col1,((Oxcal_Data[0]['likelihood']['comment'][0]) + (Oxcal_Data[0]['likelihood']['comment'][1])), italics) 
                    
workbook.close()


