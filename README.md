VBA Finance Tools
=================

What is this?
-------------
I used these VBA scripts to automate common excel tasks and make my work more efficient.  
This is not a "proper" Git repo, the files were simply copied from my PC.

Short description of the scripts
--------------------------------

**FCF_model**: automated sensitivity analysis (heat map and tornado chart) in an FCF model  

**LinkBreaker**: breaks the links in all xlsx files in a folder.  

**DataSplitter**: The purpose of this file is to split large data sets (master data) based on values in a specific column and save the new sets of data into separate excel files. For example, creating monthly transaction lists from the data for the whole year. Basically the following steps are automated: filter for X, copy filtered data into new file, save & close, filter for Y, copy filtered data,etc.  

**CombineBooks**: copy all worksheets from all excel files in a given folder into one single excel file  

**PivotAddValues**: if you have lots of fields and want to add all of them to a pivot talbe as sum values, it's easier to do it with VBA than manually adding them one by one (and changing the values to "sum" if excel doesn't do that by default).
