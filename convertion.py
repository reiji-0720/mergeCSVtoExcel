import pandas
import openpyxl

# Creating a new excel file
wb = openpyxl.Workbook()

# Getting a new sheet
sheet = book['Sheet']

# Changing file name 'Sheet' --> 'Sheet test' 
sheet.title = 'Sheet test'

# Saving a work book
wb.save('./data-list.elsx')
print('aaa')