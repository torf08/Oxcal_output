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

#import tkinter for gui uses, JSON parser to load JSON string, and xlsxwriter to create and write to an excelsheet and os to open excel file after creation

import tkinter as tk
from tkinter import filedialog
from tkinter import *
import json
import xlsxwriter
import os

#Main Tk Window Frame
root = tk.Tk()
root.title("OXCAL JSON to XLSX")
root.geometry('+750+500')

Bayesian = tk.IntVar()
Age_Scale = tk.IntVar()
excel_filename = tk.StringVar()
json_filename = tk.StringVar()
        
class Oxcal_output(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.grid()
        self.create_widgets()
        
    def create_widgets(self):
        self.json = tk.Button(self, text = "Select JSON File", fg = 'blue', bg = 'grey')
        self.json["command"] = self.json_file
        self.json.grid(row =0, column=0)
        
        self.JSON_Loaded = tk.Label(self, text = "")
        self.JSON_Loaded.grid(row = 1, column = 0)
        
        self.xlsx = tk.Button(self, text = "Select XLSX File", fg = 'blue', bg = 'grey')
        self.xlsx["command"] = self.xlsx_file
        self.xlsx.grid(row=0, column=2)
        
        self.XLSX_Loaded = tk.Label(self, text = "")
        self.XLSX_Loaded.grid(row = 1, column = 2)
        
        self.Blank_1 = tk.Label(self, text = "")
        self.Blank_1.grid(row = 6, column = 0)
        
        self.Blank_2 = tk.Label(self, text = "")
        self.Blank_2.grid(row = 6, column = 2)
        
        self.files_loaded = tk.Label(self, text="")
        self.files_loaded.grid(row = 7, column = 1)
        
        self.clear_files = tk.Button(self, text = "Clear Selections", fg='blue', bg = 'grey')
        self.clear_files["command"] = self.clear_selections
        self.clear_files.grid(row = 0, column = 1)
        
        self.Bayesian_Label = tk.Label(self, text="Does the JSON file\n contain Bayesian\n modeling data?")
        self.Bayesian_1 = tk.Radiobutton(self, text = "Yes", variable = Bayesian, value = 1)
        self.Bayesian_2 = tk.Radiobutton(self, text = "No", variable = Bayesian, value = 2)
        
        self.AgeScale_Label = tk.Label(self, text = "What age scale\nshould be used?")
        self.AgeScale_0 = tk.Radiobutton(self, text = "BCE/CE", variable = Age_Scale, value = 2)
        self.AgeScale_1 = tk.Radiobutton(self, text = "BC/AD", variable = Age_Scale, value = 1)
        self.AgeScale_2 = tk.Radiobutton(self, text = "Cal BP", variable = Age_Scale, value = 3)
                
        self.cont = tk.Button(self, text="Continue", fg = 'green', command=self.continue_script)
        self.quit = tk.Button(self, text="Quit", fg = 'red', command=self.master.quit)
        
        self.cont.grid(row = 7, column = 0)
        self.quit.grid(row=7, column = 2)
        
        self.Bayesian_1.grid(row = 3, column = 0)
        self.Bayesian_2.grid(row = 3, column = 2)
        self.Bayesian_Label.grid(row = 2, column = 1)
        
        self.AgeScale_0.grid(row = 5, column = 0)
        self.AgeScale_1.grid(row = 5, column = 1)
        self.AgeScale_2.grid(row = 5, column = 2)
        self.AgeScale_Label.grid(row = 4, column = 1)
       
    def json_file(self):
        json_filename_v2 = filedialog.askopenfile(initialdir ="~/", title = "Select .JSON file", filetypes = (("JSON files", "*.json"),("all files","*.*")))
        if json_filename_v2 == None:
            print ("No file selected for open file. Should be a .json")
            self.JSON_Loaded.config(text='No file selected', fg = 'red')
            json_filename.set(json_filename_v2)
        else:
            x = (json_filename_v2.name)
            json_name_split = x.split(".")
            if 'json' in json_name_split:
                print ("Opening read file")
                json_filename.set(x)
                self.JSON_Loaded.config(text="JSON File Loaded", fg = "green", bg = 'white')
            else:
                print ("The selected file must be a .json file extension!")
                self.JSON_Loaded.config(text="Must be a JSON File", fg='red', bg = 'white')
                
    def xlsx_file(self):
        excel_filename_v2 = filedialog.asksaveasfilename(initialdir ="~/", title = "Save to .xlsx file", filetypes = (("Excel Workbook", "*.xlsx"),("Excel Workbook", "*.xls"),("all files","*.*")))
        if excel_filename_v2 == "":
            print ("No file selected for save file. Should be a .xlsx or .xls")
            self.XLSX_Loaded.config(text="No file selected", fg='red')
            excel_filename.set(excel_filename_v2)
        else:
            excel_name_split = (excel_filename_v2.split("."))
            if 'xlsx' in excel_name_split or 'xls' in excel_name_split:
                excel_filename.set(excel_filename_v2)
                print("Opening save file")
                self.XLSX_Loaded.config(text = "XLSX File Loaded!", fg = 'green')
            else:
                print ("The save file must be a .xlsx or .xls file etension!")
                self.XLSX_Loaded.config(text="Must be a XLSX file", fg='red')
                         
    def clear_selections(self):
        json_filename = ""
        self.JSON_Loaded.config(text="File Cleared!", fg='steel blue')
        self.json
        xlsx_filename = ""
        self.XLSX_Loaded.config(text="File Cleared!", fg='steel blue')
        Bayesian.set(0)
        Age_Scale.set(0)
        
    def Bayesian_Workbook(self, excel_filename, json_filename, Age_Scale): 
        
            #Create Workbooks for Oxcal_Data using xlsxwriter
            excel_open = excel_filename.get()
            
            workbook = xlsxwriter.Workbook(excel_open)
            Date_ranges = workbook.add_worksheet('14C Date Ranges')
            Date_PDF_Graphing = workbook.add_worksheet('Dates PDF_Graphing')
            CMD_PDF_Graphing = workbook.add_worksheet('Oxcal Cmds Graphing ')
        
            with open(json_filename.get()) as json_open:
                Oxcal_Data=json.load(json_open)
                json_open.close()
                print(type(Oxcal_Data))

            #Name varibles for various cell formats
            header = workbook.add_format({'bold': True, 'center_across': True})
            format = workbook.add_format({'num_format': '0', 'center_across': True})
            center = workbook.add_format({'center_across': True})
            percent = workbook.add_format({'num_format': '0.0%', 'center_across': True})
            italics = workbook.add_format({'italic': True})
            probNum = workbook.add_format({'num_format': '0.0000000', 'center_across': True})
            one_digit = workbook.add_format({'num_format': '0.0', 'center_across': True})

            #Set row, col and row_adj variables and row_count list; set header labels; adjust column withds; and advance rows
            sheet1_row1 = 0 
            sheet1_row2 = 0
            sheet1_col1 = 0
            sheet1_row_adj = 0
            sheet1_row3 = 1
            sheet1_row4 = 1

            sheet2_row1 = 0    
            sheet2_col1 = 0    
            sheet2_row_adj = 0
            row_count = []
            global record_count
            record_count = 0 

            sheet3_row1 = 0
            sheet3_col1 = 0
            sheet3_row_adj = 0
            row_count2 = []
            record_count2 = 0

            Labels = ['Name', 'RYCBP', 'Plus or Minus', '1s.d. Cal', '%', '2s.d. Cal', '%', 'Median','Plus or Minus','Posterior', '1s.d. Cal', '%', '2s.d. Cal', '%', 'Median', 'Plus or Minus', 'Agreement', 'Convergence', 'probNorm' ]
            for count, name in enumerate(Labels):
                Date_ranges.write(sheet1_row1, sheet1_col1+count, name, header)

            #Set Column widths
            Date_ranges.set_column('A:A', 20)
            Date_ranges.set_column('C:C', 11)    
            Date_ranges.set_column('D:D', 17)
            Date_ranges.set_column('F:F', 17)
            Date_ranges.set_column('I:I', 11)
            Date_ranges.set_column('K:K', 17)
            Date_ranges.set_column('M:M', 17)
            Date_ranges.set_column('I:I', 11)
            Date_ranges.set_column('P:S', 11)

            #Advance and create row varibles necessary for age ranges and percentages
            sheet1_row1+=1
            sheet1_row2+=1

            #Apply AD or BC labels to median dates
            def Medians(Median, Deviation, a, b, Age_Scale):

                #Had to set specific cell format for Plus or Minus column as "format" cell designation wasn't working inside function
                cell_format01 = workbook.add_format()
                cell_format01.set_num_format('0')
                cell_format01.set_align('center')

                if Age_Scale.get() == 1:
                    if Median > 0: 
                        AD_Median = ("AD " + str(int(Median)))
                        Date_ranges.write(sheet1_row1, sheet1_col1+a, AD_Median, center)
                        Date_ranges.write(sheet1_row1, sheet1_col1+b, Deviation, cell_format01)
                    elif Median < 0:
                        BC_Median = ("BC " + str(int(abs(Median))))
                        Date_ranges.write(sheet1_row1, sheet1_col1+a, BC_Median, center)
                        Date_ranges.write(sheet1_row1, sheet1_col1+b, Deviation, cell_format01)
                elif Age_Scale.get() == 2:
                    if Median > 0: 
                        AD_Median = ("CE " + str(int(Median)))
                        Date_ranges.write(sheet1_row1, sheet1_col1+a, AD_Median, center)
                        Date_ranges.write(sheet1_row1, sheet1_col1+b, Deviation, cell_format01)
                    elif Median < 0:
                        BC_Median = ("BCE " + str(int(abs(Median))))
                        Date_ranges.write(sheet1_row1, sheet1_col1+a, BC_Median, center)
                        Date_ranges.write(sheet1_row1, sheet1_col1+b, Deviation, cell_format01)
                else:
                   Cal_BP_Median = round(1949 - Median)
                   Date_ranges.write(sheet1_row1, sheet1_col1+a, Cal_BP_Median, center)
                   Date_ranges.write(sheet1_row1, sheet1_col1+b, Deviation, cell_format01)

                return

            #Seperate and write to spreadsheet probabilities
            def Ranges(range, indpos1, indpos2, col1, col2, row, Age_Scale):

                global c
                c = row

                for prob in range[indpos1:indpos2]:
                    step_in_1a = prob
    
                    if Age_Scale.get() == 1:
                        for sets in step_in_1a:
                            step_in_2a = sets
                            step_in_3a = step_in_2a[0]
                            step_in_3b = step_in_2a[1]
                            step_in_3c = (step_in_2a[2]/100)
    
                            if step_in_3a >= 0:
                                if isinstance(step_in_3b, float):
                                    AD_Date = ('AD ' + str(int(round(step_in_3a,0))) + '- AD ' + str(int(round(step_in_3b,0))))
                                    Date_ranges.write(c, sheet1_col1+col1, AD_Date, center)
                                
                                elif isinstance(step_in_3b, int):
                                    AD_Date = ('AD ' + str(int(round(step_in_3a,0))) + '- AD ' + str(step_in_3b))
                                    Date_ranges.write(c, sheet1_col1+col1, AD_Date, center)
                                
                                else:
                                    AD_Date = ('AD ' + str(int(round(step_in_3a,0))) + '- AD ' + step_in_3b)
                                    Date_ranges.write(c, sheet1_col1+col1, AD_Date, center)
        
                            elif step_in_3a <= 0 and step_in_3b <= 0:
                                step_in_3a = abs(step_in_3a)
                                step_in_3b = abs(step_in_3b)
        
                                BC_Date = ('BC ' + str(int(round(step_in_3a,0))) + '- BC ' + str(int(round(step_in_3b,0))))
                                Date_ranges.write(c, sheet1_col1+col1, BC_Date, center)
        
                            else:
                                step_in_3a = abs(step_in_3a)
                                BC_AD_Date = ('BC ' + str(int(round(step_ind_3a,0))) + '- AD '+ str(int(round(step_in_3b,0))))
                                Date_ranges.write(c, sheet1_col1+col1, BC_AD_Date, center)
                
                            Date_ranges.write(c, sheet1_col1+col2, step_in_3c, percent)
            
                            c += 1
            
                    elif Age_Scale.get() == 2:
                        for sets in step_in_1a:
                            step_in_2a = sets
                            step_in_3a = step_in_2a[0]
                            step_in_3b = step_in_2a[1]
                            step_in_3c = (step_in_2a[2]/100)

                            if step_in_3a >= 0:
                                if isinstance(step_in_3b, float):
                                    CE_Date = ('CE ' + str(int(round(step_in_3a,0))) + '- CE ' + str(int(round(step_in_3b,0))))
                                    Date_ranges.write(c, sheet1_col1+col1, CE_Date, center)
                            
                                elif isinstance(step_in_3b, int):
                                    CE_Date = ('CE ' + str(int(round(step_in_3a,0))) + '- CE ' + str(step_in_3b))
                                    Date_ranges.write(c, sheet1_col1+col1, step_in_3a, center)
                                
                                else:
                                    CE_Date = ('CE ' + str(int(round(step_in_3a,0))) + '- CE ' + step_in_3b)
                                    Date_ranges.write(c, sheet1_col1+col1, step_in_3a, center)
    
                            elif step_in_3a <= 0 and step_in_3b <= 0:
                                step_in_3a = abs(step_in_3a)
                                step_in_3b = abs(step_in_3b)
    
                                BCE_Date = ('BCE ' + str(int(round(step_in_3a,0))) + '- BCE ' + 
                                str(int(round(step_in_3b,0))))
                                Date_ranges.write(c, sheet1_col1+col1, BCE_Date, center)
                
                            else:
                                step_in_3a = abs(step_in_3a)
                
                                BCE_CE_Date = ('BCE ' + str(int(round(step_ind_3a,0))) + '- CE '+ 
                                str(int(round(step_in_3b,0))))
                                Date_ranges.write(c, sheet1_col1+col1, BCE_CE_Date, center)
                
                            Date_ranges.write(c, sheet1_col1+col2, step_in_3c, percent)
            
                            c += 1
            
                    else:
                        for sets in step_in_1a:
                            step_in_2a = sets
                            print(step_in_2a[0])
                            print(step_in_2a[1])
                            step_in_3a = str(round(1950 - step_in_2a[0]))
                            step_in_3b = str(round(1950 - step_in_2a[1]))
                            step_in_3c = (step_in_2a[2]/100)
            


                            Cal_BP_Date = (step_in_3a + '-' + step_in_3b + ' Cal Yr BP')
                            Date_ranges.write(c, sheet1_col1+col1, Cal_BP_Date, center)

                            Date_ranges.write(c, sheet1_col1+col2, step_in_3c, percent)

                            c += 1
            
                        return

            #Determine number of rows that each samples probability ranges used. Correct the rowshift so that a 1 row spacing is kept between each sample.
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

            #Add probabilites and dates to second worksheet to aid in Datagraph usage
            def Probabilities (name, prob, start, resolution, x, y , z, a, b, Age_Scale, work_sheet):
                if name == 'unmodeled':
                    header_dates = (list_name + ' _dates')
                    header_prob = (list_name + ' _prob')
                    header_prob_mod = (list_name + 'prob_mod')
                elif name == 'modeled':
                    header_dates = (list_name + ' B dates')
                    header_prob = (list_name + ' B prob')
                    header_prob_mod = (list_name + 'B prob_mod')

                global row1
                global col1
                global row_adj
                global rowcount

                row1 = x
                col1 = y
                row_adj = z
                rowcount = a
                recordcount = b

                work_sheet.write(row1, col1, header_dates,  header)
                col1 += 1
                col_prob_mod = col1 + 1

                work_sheet.write(row1, col1, header_prob, header)
                work_sheet.write(row1, col_prob_mod, header_prob_mod, header)
                row1 += 1

                #write probability density to worksheet2 and count number of rows    

                for values in prob:
                    prob_mod = (values/2) + recordcount
                    work_sheet.write(row1, col1, values)
                    work_sheet.write(row1, col_prob_mod, prob_mod)
                    row1 += 1
                    row_adj += 1
                    rowcount.append(row_adj)

                #reset PDF_Graphing row1 back to 1 and create seperate variable for list of dates beginning with start date
                row1 = 1
                if Age_Scale.get() == 1 or Age_Scale.get() == 2:
                    start_date = start
                else: 
                    start_date = 1949 - start

                #write dates of probability density to PDF_Graphing and increment up from the start date by resolution (set in Oxcal calibration)
                for count in rowcount:
                    if Age_Scale.get() == 1 or Age_Scale.get() == 2:
                        work_sheet.write(row1, col1-1, start_date)
                        start_date += resolution
                        row1 += 1
                    else:
                        work_sheet.write(row1, col1-1, start_date)
                        start_date -= resolution
                        row1 += 1
                #reset PDF_Graphing columns, rows, and row_count to necessary values for restart of loop     
                col1 += 2
                row1 = 0
                row_adj = 0
                rowcount = []

                return

            #Seperate indices dictonary types from Oxcal_Data list
            for IndvData in Oxcal_Data:

                list_liklihood = IndvData['likelihood']
                list_comment = list_liklihood['comment']

                #very janky way to get around varibles that don't exist in the first few items in the list
                if list_comment[0] == "OxCal v4.4.2 Bronk Ramsey (2020); r:5": 
                    continue
                
                elif list_comment[0] == "OxCal v4.3.2 Bronk Ramsey (2017); r:5": 
                    continue

                list_op = IndvData['op']

                if list_op == "Sequence" or list_op == "Phase":
                    continue

                #Check if Operator is Boundary and pull out posterior information for Boundary Ranges
                if list_op == "Boundary":

                    list_name = IndvData['name']
                    list_posterior = IndvData['posterior']
                    modeled_range = list_posterior['range']
                    modeled_median = list_posterior['median']
                    modeled_sigma = list_posterior['sigma']
                    modeled_probNorm = list_posterior['probNorm']
                    modeled_convergence = list_posterior['convergence']
                    modeled_prob = list_posterior['prob']
                    modeled_start = list_posterior['start']
                    modeled_res = list_posterior['resolution']

                    boundary_name = list_op + " " + list_name

                    Date_ranges.write(sheet1_row1, sheet1_col1, boundary_name )
                    Date_ranges.write(sheet1_row1, sheet1_col1+17, modeled_convergence, one_digit) 
                    Date_ranges.write(sheet1_row1, sheet1_col1+18, modeled_probNorm, probNum)

                    Medians(modeled_median, modeled_sigma, 14, 15, Age_Scale)

                    Ranges(modeled_range, 1, 2, 10, 11, sheet1_row3, Age_Scale)
                    sheet1_row3 = c

                    Ranges(modeled_range, 2, 3, 12, 13, sheet1_row4, Age_Scale)
                    sheet1_row4 = c

                    Row_Shift(sheet1_row3, sheet1_row4)
                    sheet1_row3 = rowshift1
                    sheet1_row4 = rowshift2
    
                    #Adjust rows between unmodeled and modeled dates to keep 1 row space between samples
                    Row_adj = sheet1_row2 - sheet1_row3
                    if Row_adj < 0:
                        sheet1_row2 = sheet1_row2 + abs(Row_adj)
                        sheet1_row1 = sheet1_row2
                    if Row_adj > 0: 
                        sheet1_row3 = sheet1_row3 + abs(Row_adj)
                        sheet1_row4 = sheet1_row3
                    if Row_adj == 0:
                        sheet1_row2 = sheet1_row3
    
                    #Add to PDF_Graphing sheet dates and probabilities
                    Probabilities('modeled', modeled_prob, modeled_start, modeled_res, sheet2_row1, sheet2_col1, sheet2_row_adj, row_count, record_count, Age_Scale, Date_PDF_Graphing)
                    sheet2_col1 = col1
                    sheet2_row1 = row1
                    sheet2_row_adj = row_adj
                    row_count = rowcount
    
                    record_count += 0.75
    
                #Check if Operator is R_Date and pull out both likelihood and posterior information
                elif list_op == "R_Date" or list_op == "R_Simulate":

                    #Sample Name, RCYBP date, and error on RCYBP measurment
                    list_name = IndvData['name']
                    list_date = IndvData['date']
                    list_error = IndvData['error']

                    #Unmodeled Range, Median, Sigma, Probability, Range Start, and Resolution
                    unmodeled_range = list_liklihood['range']
                    unmodeled_median = list_liklihood['median']
                    unmodeled_sigma = list_liklihood['sigma']
                    unmodeled_prob = list_liklihood['prob']
                    unmodeled_start = list_liklihood['start']
                    unmodeled_res = list_liklihood['resolution']

                    #Modeled Range, Median, Sigma, Probability, Model Agreement, ProbNorm, and Convergence
                    list_posterior = IndvData['posterior']
                    modeled_range = list_posterior['range']
                    modeled_median = list_posterior['median']
                    modeled_sigma = list_posterior['sigma']
                    modeled_agreement = list_posterior['agreement']
                    modeled_probNorm = list_posterior['probNorm']
                    modeled_convergence = list_posterior['convergence']
                    modeled_prob = list_posterior['prob']
                    modeled_start = list_posterior['start']
                    modeled_res = list_posterior['resolution']

                    Date_ranges.write(sheet1_row1, sheet1_col1, list_name)
                    Date_ranges.write(sheet1_row1, sheet1_col1+1, list_date, format)
                    Date_ranges.write(sheet1_row1, sheet1_col1+2, list_error, format) 
                    Date_ranges.write(sheet1_row1, sheet1_col1+16, modeled_agreement, one_digit) 
                    Date_ranges.write(sheet1_row1, sheet1_col1+17, modeled_convergence, one_digit) 
                    Date_ranges.write(sheet1_row1, sheet1_col1+18, modeled_probNorm, probNum) 

                    #Writing Unmodeled and Modeled Medians and Plus or Minuses to excelsheet
                    Medians(unmodeled_median, unmodeled_sigma, 7, 8, Age_Scale)
                    Medians(modeled_median, modeled_sigma, 14, 15, Age_Scale)

                    #Writing Unmodeled and Modeled ranges for 1 and 2 sigma to excelsheet
                    Ranges(unmodeled_range, 1, 2, 3, 4, sheet1_row1, Age_Scale)
                    sheet1_row1 = c
                    Ranges(unmodeled_range, 2, 3, 5, 6, sheet1_row2, Age_Scale)
                    sheet1_row2 = c
                    Ranges(modeled_range, 1, 2, 10, 11, sheet1_row3, Age_Scale)
                    sheet1_row3 = c
                    Ranges(modeled_range, 2, 3, 12, 13, sheet1_row4, Age_Scale)
                    sheet1_row4 = c

                    #Adjust sheet_row_num values to keep a consistent 1 row space between samples
                    Row_Shift(sheet1_row1, sheet1_row2)
                    sheet1_row1 = rowshift1
                    sheet1_row2 = rowshift2

                    Row_Shift(sheet1_row3, sheet1_row4)
                    sheet1_row3 = rowshift1
                    sheet1_row4 = rowshift2

                    #Adjust rows between unmodeled and modeled dates to keep 1 row space between samples
                    Row_adj = sheet1_row2 - sheet1_row3
                    if Row_adj < 0:
                        sheet1_row2 = sheet1_row2 + abs(Row_adj)
                        sheet1_row1 = sheet1_row2
                    if Row_adj > 0: 
                        sheet1_row3 = sheet1_row3 + abs(Row_adj)
                        sheet1_row4 = sheet1_row3
                    if Row_adj == 0:
                        sheet1_row2 = sheet1_row3

                    #Add to PDF_Graphing sheet dates and probabilities
                    Probabilities('unmodeled', unmodeled_prob, unmodeled_start, unmodeled_res, sheet2_row1, sheet2_col1, sheet2_row_adj, row_count, record_count, Age_Scale, Date_PDF_Graphing)
                    sheet2_row1 = row1
                    sheet2_col1 = col1
                    sheet2_row_adj = row_adj
                    row_count = rowcount
    
    
                    Probabilities('modeled', modeled_prob, modeled_start, modeled_res, sheet2_row1, sheet2_col1, sheet2_row_adj, row_count, record_count, Age_Scale, Date_PDF_Graphing)
                    sheet2_row1 = row1
                    sheet2_col1 = col1
                    sheet2_row_adj = row_adj
                    row_count = rowcount
    
                    record_count += 0.75
    
                elif list_op  == "Span" or list_op == "Interval" or list_op == "Start" or list_op == "End" or list_op == "Transition" or list_op == "After" or list_op == "Before" or list_op == "Outlier_Model":
    
                    list_name = IndvData['name']
                    list_posterior = IndvData['posterior']
                    modeled_range = list_posterior['range']
                    modeled_median = list_posterior['median']
                    modeled_sigma = list_posterior['sigma']
                    modeled_probNorm = list_posterior['probNorm']
                    modeled_convergence = list_posterior['convergence']
                    modeled_prob = list_posterior['prob']
                    modeled_start = list_posterior['start']
                    modeled_res = list_posterior['resolution']
    
                    #Add to CMD_PDF_Graphing sheet dates and probabilities
                    Probabilities('modeled', modeled_prob, modeled_start, modeled_res, sheet3_row1, sheet3_col1, sheet3_row_adj, row_count2, record_count2, Age_Scale, CMD_PDF_Graphing)
                    sheet3_col1 = col1
                    sheet3_row1 = row1
                    sheet3_row_adj = row_adj
                    row_count2 = rowcount
    
                    record_count2 += 0.75

            #Add reference to calibration software used
            Date_ranges.write(sheet1_row1, sheet1_col1,((Oxcal_Data[0]['likelihood']['comment'][0]) + (Oxcal_Data[0]['likelihood']['comment'][1])), italics) 
            
            workbook.close()
            self.files_loaded.config(text = "Opening Excel!", fg='green')
            print("Opening Excel workbook")
            command = "open -a '/Applications/Microsoft Excel.app' '" + excel_open + "'"
            os.system(command)
            
    def Non_Bayesian_Workbook(self, excel_filename, json_filename, Age_Scale): 
        
        #Create Workbooks for Oxcal_Data using xlsxwriter
        excel_open = excel_filename.get()
        
        workbook = xlsxwriter.Workbook(excel_open)
        Date_ranges = workbook.add_worksheet("14C Date Ranges")
        PDF_Graphing = workbook.add_worksheet("Dates PDF Graphing")
    
        with open(json_filename.get()) as json_open:
            Oxcal_Data=json.load(json_open)
            json_open.close()
            
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
        sheet1_row3 = 1
        sheet1_row4 = 1

        sheet2_row1 = 0    
        sheet2_col1 = 0    
        sheet2_row_adj = 0
        row_count = []
        record_count = 0

        Labels = ['Name', 'RYCBP', 'Plus or Minus', '1s.d. Cal', '%', '2s.d. Cal', '%', 'Median','Plus or Minus']
        for count, name in enumerate(Labels):
            Date_ranges.write(sheet1_row1, sheet1_col1+count, name, header)

        #Set Column widths
        Date_ranges.set_column('A:A', 20)
        Date_ranges.set_column('C:C', 11)    
        Date_ranges.set_column('D:D', 17)
        Date_ranges.set_column('F:F', 17)
        Date_ranges.set_column('I:I', 11)

        #Advance and create row varibles necessary for age ranges and percentages
        sheet1_row1+=1
        sheet1_row2+=1

        #Apply AD or BC labels to median dates
        def Medians(Median, Deviation, a, b, Age_Scale):

            #Had to set specific cell format for Plus or Minus column as "format" cell designation wasn't working inside function
            cell_format01 = workbook.add_format()
            cell_format01.set_num_format('0')
            cell_format01.set_align('center')

            if Age_Scale.get() == 1:
                if Median > 0: 
                    AD_Median = ("AD " + str(int(Median)))
                    Date_ranges.write(sheet1_row1, sheet1_col1+a, AD_Median, center)
                    Date_ranges.write(sheet1_row1, sheet1_col1+b, Deviation, cell_format01)
                elif Median < 0:
                    BC_Median = ("BC " + str(int(abs(Median))))
                    Date_ranges.write(sheet1_row1, sheet1_col1+a, BC_Median, center)
                    Date_ranges.write(sheet1_row1, sheet1_col1+b, Deviation, cell_format01)
            elif Age_Scale.get() == 2:
                if Median > 0: 
                    AD_Median = ("CE " + str(int(Median)))
                    Date_ranges.write(sheet1_row1, sheet1_col1+a, AD_Median, center)
                    Date_ranges.write(sheet1_row1, sheet1_col1+b, Deviation, cell_format01)
                elif Median < 0:
                    BC_Median = ("BCE " + str(int(abs(Median))))
                    Date_ranges.write(sheet1_row1, sheet1_col1+a, BC_Median, center)
                    Date_ranges.write(sheet1_row1, sheet1_col1+b, Deviation, cell_format01)
            else:
               Cal_BP_Median = round(1949 - Median)
               Date_ranges.write(sheet1_row1, sheet1_col1+a, Cal_BP_Median, center)
               Date_ranges.write(sheet1_row1, sheet1_col1+b, Deviation, cell_format01)

            return

        #Seperate and write to spreadsheet probabilities
        def Ranges(range, indpos1, indpos2, col1, col2, row, Age_Scale):

            global c
            c = row

            for prob in range[indpos1:indpos2]:
                step_in_1a = prob
    
                if Age_Scale.get() == 1:
                    for sets in step_in_1a:
                        step_in_2a = sets
                        step_in_3a = step_in_2a[0]
                        step_in_3b = step_in_2a[1]
                        step_in_3c = (step_in_2a[2]/100)
    
                        if step_in_3a >= 0:
                            if isinstance(step_in_3b, float):
                                AD_Date = ('AD ' + str(int(round(step_in_3a,0))) + '- AD ' + str(int(round(step_in_3b,0))))
                                Date_ranges.write(c, sheet1_col1+col1, AD_Date, center)
                                
                            elif isinstance(step_in_3b, int):
                                AD_Date = ('AD ' + str(int(round(step_in_3a,0))) + '- AD ' + str(step_in_3b))
                                Date_ranges.write(c, sheet1_col1+col1, AD_Date, center)
                                
                            else:
                                AD_Date = ('AD ' + str(int(round(step_in_3a,0))) + '- AD ' + step_in_3b)
                                Date_ranges.write(c, sheet1_col1+col1, AD_Date, center)
        
                        elif step_in_3a <= 0 and step_in_3b <= 0:
                            step_in_3a = abs(step_in_3a)
                            step_in_3b = abs(step_in_3b)
        
                            BC_Date = ('BC ' + str(int(round(step_in_3a,0))) + '- BC ' + str(int(round(step_in_3b,0))))
                            Date_ranges.write(c, sheet1_col1+col1, BC_Date, center)
        
                        else:
                            step_in_3a = abs(step_in_3a)
                            BC_AD_Date = ('BC ' + str(int(round(step_ind_3a,0))) + '- AD '+ str(int(round(step_in_3b,0))))
                            Date_ranges.write(c, sheet1_col1+col1, BC_AD_Date, center)
                
                        Date_ranges.write(c, sheet1_col1+col2, step_in_3c, percent)
            
                        c += 1
            
                elif Age_Scale.get() == 2:
                    for sets in step_in_1a:
                        step_in_2a = sets
                        step_in_3a = step_in_2a[0]
                        step_in_3b = step_in_2a[1]
                        step_in_3c = (step_in_2a[2]/100)

                        if step_in_3a >= 0:
                            if isinstance(step_in_3b, float):
                                CE_Date = ('CE ' + str(int(round(step_in_3a,0))) + '- CE ' + str(int(round(step_in_3b,0))))
                                Date_ranges.write(c, sheet1_col1+col1, CE_Date, center)
                            
                            elif isinstance(step_in_3b, int):
                                CE_Date = ('CE ' + str(int(round(step_in_3a,0))) + '- CE ' + str(step_in_3b))
                                Date_ranges.write(c, sheet1_col1+col1, step_in_3a, center)
                                
                            else:
                                CE_Date = ('CE ' + str(int(round(step_in_3a,0))) + '- CE ' + step_in_3b)
                                Date_ranges.write(c, sheet1_col1+col1, step_in_3a, center)
    
                        elif step_in_3a <= 0 and step_in_3b <= 0:
                            step_in_3a = abs(step_in_3a)
                            step_in_3b = abs(step_in_3b)
    
                            BCE_Date = ('BCE ' + str(int(round(step_in_3a,0))) + '- BCE ' + 
                            str(int(round(step_in_3b,0))))
                            Date_ranges.write(c, sheet1_col1+col1, BCE_Date, center)
                
                        else:
                            step_in_3a = abs(step_in_3a)
                
                            BCE_CE_Date = ('BCE ' + str(int(round(step_ind_3a,0))) + '- CE '+ 
                            str(int(round(step_in_3b,0))))
                            Date_ranges.write(c, sheet1_col1+col1, BCE_CE_Date, center)
                
                        Date_ranges.write(c, sheet1_col1+col2, step_in_3c, percent)
            
                        c += 1
            
                else:
                    for sets in step_in_1a:
                        step_in_2a = sets
                        print(step_in_2a[0])
                        print(step_in_2a[1])
                        step_in_3a = str(round(1950 - step_in_2a[0]))
                        step_in_3b = str(round(1950 - step_in_2a[1]))
                        step_in_3c = (step_in_2a[2]/100)
            


                        Cal_BP_Date = (step_in_3a + '-' + step_in_3b + ' Cal Yr BP')
                        Date_ranges.write(c, sheet1_col1+col1, Cal_BP_Date, center)

                        Date_ranges.write(c, sheet1_col1+col2, step_in_3c, percent)

                        c += 1
            
                    return

        #Determine number of rows that each samples probability ranges used. Correct the rowshift so that a 1 row spacing is kept between each sample.
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

        #Add probabilites and dates to second worksheet to aid in Datagraph usage
        def Probabilities (name, prob, start, resolution, x, y , z, a, b, Age_Scale):
            if name == 'unmodeled':
                header_dates = (list_name + ' _dates')
                header_prob = (list_name + ' _prob')
                header_prob_mod = (list_name + 'prob_mod')
            elif name == 'modeled':
                header_dates = (list_name + ' B dates')
                header_prob = (list_name + ' B prob')

            global row1
            global col1
            global row_adj
            global rowcount
            global record_count

            row1 = x
            col1 = y
            row_adj = z
            rowcount = a
            record_count = b

            PDF_Graphing.write(row1, col1, header_dates,  header)
            col1 += 1
            col_prob_mod = col1 + 1

            #print (col1, col_prob_mod)
            PDF_Graphing.write(row1, col1, header_prob, header)
            PDF_Graphing.write(row1, col_prob_mod, header_prob_mod, header)
            row1 += 1
            #write probability density to worksheet2 and count number of rows
            for dates in prob:
                prob_mod = (dates/2) + record_count
                PDF_Graphing.write(row1, col1, dates)
                PDF_Graphing.write(row1, col_prob_mod, prob_mod)
                #print(dates)
                #print(prob_mod)
                row1 += 1
                row_adj += 1
                rowcount.append(row_adj)

            #reset PDF_Graphing row1 back to 1 and create seperate variable for list of dates beginning with start date
            row1 = 1
            if Age_Scale.get() == 1 or Age_Scale.get() == 2:
                start_date = start
            else: 
                start_date = 1949 - start

            #write dates of probability density to PDF_Graphing and increment up from the start date by resolution (set in Oxcal calibration)
            for count in row_count:
                if Age_Scale.get() == 1 or Age_Scale.get() == 2:
                    PDF_Graphing.write(row1, col1-1, start_date)
                    start_date += resolution
                    row1 += 1
                else:
                    PDF_Graphing.write(row1, col1-1, start_date)
                    start_date -= resolution
                    row1 += 1
            #reset PDF_Graphing columns, rows, and row_count to necessary values for restart of loop     
            col1 += 2
            row1 = 0
            row_adj = 0
            rowcount = []
            
            return
            
        #Seperate indices dictonary types from Oxcal_Data list
        for IndvData in Oxcal_Data:

            list_liklihood = IndvData['likelihood']
            list_comment = list_liklihood['comment']
            
            #very janky way to get around varibles that don't exist in the first few items in the list
            if list_comment[0] == "OxCal v4.4.2 Bronk Ramsey (2020); r:5": 
                continue
                
            elif list_comment[0] == "OxCal v4.3.2 Bronk Ramsey (2017); r:5": 
                continue
                
            list_op = IndvData['op']
            
            if list_op == "Sequence" or list_op == "Phase":
                continue

            #Check if Operator is R_Date and pull out both likelihood and posterior information
            elif list_op == "R_Date" or list_op == 'R_Simulate':

                #Sample Name, RCYBP date, and error on RCYBP measurment
                list_name = IndvData['name']
                list_date = IndvData['date']
                list_error = IndvData['error']

                #Unmodeled Range, Median, Sigma, Probability, Range Start, and Resolution
                unmodeled_range = list_liklihood['range']
                unmodeled_median = list_liklihood['median']
                unmodeled_sigma = list_liklihood['sigma']
                unmodeled_prob = list_liklihood['prob']
                unmodeled_start = list_liklihood['start']
                unmodeled_res = list_liklihood['resolution']

                Date_ranges.write(sheet1_row1, sheet1_col1, list_name)
                Date_ranges.write(sheet1_row1, sheet1_col1+1, list_date, format)
                Date_ranges.write(sheet1_row1, sheet1_col1+2, list_error, format) 
    
                #Writing Unmodeled Medians and Plus or Minuses to excelsheet
                Medians(unmodeled_median, unmodeled_sigma, 7, 8, Age_Scale)

                #Writing Unmodeled and Modeled ranges for 1 and 2 sigma to excelsheet
                print (list_name)
                Ranges(unmodeled_range, 1, 2, 3, 4, sheet1_row1, Age_Scale)
                sheet1_row1 = c
                Ranges(unmodeled_range, 2, 3, 5, 6, sheet1_row2, Age_Scale)
                sheet1_row2 = c

                #Adjust sheet_row_num values to keep a consistent 1 row space between samples
                Row_Shift(sheet1_row1, sheet1_row2)
                sheet1_row1 = rowshift1
                sheet1_row2 = rowshift2

                #Add to PDF_Graphing sheet dates and probabilities
                Probabilities('unmodeled', unmodeled_prob, unmodeled_start, unmodeled_res, sheet2_row1, sheet2_col1, sheet2_row_adj, row_count, record_count, Age_Scale)
                sheet2_row1 = row1
                sheet2_col1 = col1
                sheet2_row_adj = row_adj
                row_count = rowcount
    
                record_count += 0.75
    
    
        #Add reference to calibration software used
        Date_ranges.write(sheet1_row1, sheet1_col1,((Oxcal_Data[0]['likelihood']['comment'][0]) + (Oxcal_Data[0]['likelihood']['comment'][1])), italics)
        
        
        workbook.close()
        self.files_loaded.config(text = "Opening Excel!", fg='green')
        print("Opening Excel workbook")
        command = "open -a '/Applications/Microsoft Excel.app' '" + excel_open + "'"
        os.system(command)
                
    def continue_script(self):
        json_var = json_filename.get()
        excel_var = excel_filename.get()
        
        if json_var  == 'None' or json_var == "":
            if excel_var == "":
                self.files_loaded.config(text = "No files loaded", fg = 'red')
            else: 
                self.files_loaded.config(text = "No JSON file loaded", fg = 'red')
        elif excel_var == "":
            self.files_loaded.config(text = "No XLSX file loaded", fg = 'red' )
        else:
            self.files_loaded.config(text = "Files Loaded!", fg='green')
            print("All files have been loaded")

            if Bayesian.get() == 1:
                if Age_Scale.get() != 0:
                    print("Opening Bayesian Output")
                    self.Bayesian_Workbook(excel_filename, json_filename, Age_Scale)
                else:
                    self.files_loaded.config(text = 'Choose Age Scale', fg = 'red')
            
            elif Bayesian.get() == 2:
                if Age_Scale.get() != 0:
                    print("Opening Non-Bayesian Output")
                    self.Non_Bayesian_Workbook(excel_filename, json_filename, Age_Scale)
                else:
                    self.files_loaded.config(text = 'Choose Age Scale', fg = 'red')
            else:
                print("Only 1 or 0 are accepted responses!")
                self.files_loaded.config(text = 'Choose Bayesian Option', fg = 'red')
                
app = Oxcal_output(master=root)
app.mainloop()

