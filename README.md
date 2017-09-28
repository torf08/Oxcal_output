This script is designed to use json strings generated from the javascript files produced by the Oxcal 14C calibration plugin for Firefox. 

It is intended to be run using python3 with tcl-tk built-in. See -- https://stackoverflow.com/questions/36760839/why-my-python-installed-via-home-brew-not-include-tkinter
for information on how to install python3 with tcl-tk on Mac and Linux machines.

Also needed is the xlsxwriter package for python3 for writing to excel spreadsheets. This package can be installed through 'pip install xlsxwriter'. See https://xlsxwriter.readthedocs.io/ for full documentation.

The script will clean and seperate names, RYCBP, sigma ranges, percentages, and median dates in one worksheet. It will also seperate probability density values and increment years from the start date of the range for the ease of making PDF curve graphs.

Contact Matthew A. Fort (fort1@illinois.edu) with any questions, concerns or ideas for possible improvements 
