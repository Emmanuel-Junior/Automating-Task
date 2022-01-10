'''
Author: Nwabufo.T.Emmanuel Jr
Obj: Updating Emails
Course: Python
'''

from openpyxl import load_workbook
import csv
 
#load excel file
workbook = load_workbook(filename="employees.xlsx")
 
#open workbook
sheet = workbook.active
 
#modify the desired cell
for i in range (2,15):
    index = "B"+ str(i)
    #print(sheet[index].value)
    sheet[index].value = str(sheet[index].value).replace(
       "@helpinghands.cm","@handsinhands.org"
   )
 
#save the file
workbook.save(filename="employee.xlsx")

'''reads the csv file only'''
s= open('employees.csv').read()

#replaces all data having @helpinghands.cm with @handsinhands.org 
s= s.replace("@helpinghands.cm","@handsinhands.org")

#Permits to write in the csv file
f = open('employees.csv', 'w')

#writes what was replaced into the csv filename
f.write(s)
f.close()