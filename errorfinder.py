# -*- coding: utf-8 -*-
"""
Created on Wed Apr 15 07:30:52 2020

@author: Bartek Dell

Script imports values from one excel file calculates them and 
compares with values taken from another excel file. 
It also generates .xls report at the end.
"""
# (days*919)-5%+site*208

import xlrd
import xlwt

book = xlrd.open_workbook("effort.xlsx") # Opening effort file
arkusz = book.sheet_by_name('Sheet1')    # Worksheet definition
book2 = xlrd.open_workbook("fees.xlsx")
arkusz2 = book2.sheet_by_name('Sheet1')
book3 = xlwt.Workbook(encoding="utf-8")

num_row = arkusz.nrows                   # Counting number of rows
num_row2 = arkusz2.nrows

#$$ My libraries:

A = {}
Office_lib = {}
Site_lib = {}
Os_lib = {}

#$$ Reading data from first spreadsheet and filling libraries

for i in range(1,num_row):    
    site = arkusz.cell(i,5)
    days = arkusz.cell(i,4)
    role = arkusz.cell(i,9)
    resource = arkusz.cell(i,6)
    if site.value == "Office":
        discount = (((days.value)*919)*5)/100
        kwota = (((days.value)*919)-discount)
        print("{0:<26} {1:<6} {2:>10}".format(resource.value, site.value, 
                                              (round(kwota, 2))))
        Office_lib[resource.value] = round(kwota, 2)
        Os_lib[resource.value] = round(kwota, 2) 
    elif site.value == "Site":
        if role.value == 'Observer':
            print("{0:<26} {1:<6} {2:>10}".format(resource.value, 
                                                  site.value, "Observer"))
            Site_lib[resource.value] = 0
            Os_lib[resource.value] = 0
        else:
            discount = (((days.value)*919)*5)/100
            kwota = (((days.value)*919)-discount)
            kwota2 = kwota+(days.value*208)
            print("{0:<26} {1:<6} {2:>10}".format(resource.value, site.value, 
                                                 (round(kwota2, 2))))
            Site_lib[resource.value] = round(kwota2, 2)
            Os_lib[resource.value] = Os_lib[resource.value] + kwota2

#$$ Reading data from second spreadsheet and printing values and 
#$$ libraries for checking purposes
  
for j in range(1,num_row2):
    res_name = arkusz2.cell(j,4)
    sales = arkusz2.cell(j,6)
    if res_name.value not in A:
        A[res_name.value] = sales.value
    else:
        A[res_name.value] = round((A[res_name.value] + sales.value), 2)

print(A)
print(Office_lib)
print(Site_lib)
print(Os_lib) 

#$$ Creating report "raport.xls".

rap_sheet1 = book3.add_sheet("Sheet1")
rap_sheet1.write(0, 0, "Name")
rap_sheet1.write(0, 2, "Total")
rap_sheet1.write(0, 4, "Office")
rap_sheet1.write(0, 5, "Site")
rap_sheet1.write(0, 3, "O + S")

line = 1
for k in A.keys():    
    rap_sheet1.write(line, 0, k)
    line = line + 1
    
line = 1
for m in A.values():    
    rap_sheet1.write(line, 2, m)
    line = line + 1

line = 1    
for n in A.keys():
    rap_sheet1.write(line, 4, (Office_lib[n]))
    line = line + 1
    
line = 1    
for p in A.keys():
    if p not in Site_lib:
        rap_sheet1.write(line, 5, (0))
        line = line + 1
    else:    
        rap_sheet1.write(line, 5, (Site_lib[p]))
        line = line + 1

line = 1    
for t in A.keys():
    rap_sheet1.write(line, 3, (Os_lib[t]))
    line = line + 1        


book3.save("raport.xls") 
input("Press enter to end program:")

          
        
        