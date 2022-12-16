import os
import csv
import openpyxl
from openpyxl.utils import column_index_from_string

# 作成したエクセルにCSVデータを反映させる
def merge(csv_filename,i,Path,wb):
    csv_filename_only = os.path.splitext(csv_filename)[0].strip('/') # CSVのファイル名だけ取り出す

    # Creating a new excel file
    # wb = openpyxl.Workbook()
    csv_filename = Path.rstrip('/') + csv_filename

    print('csv-filename:'+ csv_filename)
    print('csv-filename-only:'+ csv_filename_only)

    # open a csv file
    with open(csv_filename,newline='') as csvf:
        # reading a csv file
        data = csv.reader(csvf)
        
        # Getting a new sheet
        wb.create_sheet(index = i, title=csv_filename_only)
        sheet = wb[csv_filename_only]
    
       
        r = 1
        for line in data:
            c = 1 #colomn
            for v in line:
                if r == 2:
                    if c == 1:
                        sheet.cell(row=2,column=1).number_format = 'yyyy-mm-dd'
                    elif c == 2:
                        sheet.cell(row=2,column=2).number_format = 'hh:mm:ss.ssss'
                    # elif c == 3:
                    #     sheet.cell(row=2,column=3).number_format = 0
                sheet.cell(row = r,column = c).value = v
                c += 1
            r += 1 # row

        # for row in sheet:
        #     for cell in row:
        #         col = column_index_from_string(cell.column)

        #         if sheet.cell(row = 1, column = col).value == 'Z[m/s^2]':
        #             cell.number_format = '0.00'

        # # シート上でのセルのフォーマットの指定
    
        # sheet.cell(row=2,column=1).number_format = 'yyyy-mm-dd'
        # sheet.cell(row=2,column=2).number_format = 'hh:mm:ss.ssss'
        # sheet.cell(row=2,column=3).number_format = '0'
        # sheet.cell(row=2,column=4).number_format = '0.00'
        # sheet.cell(row=2,column=5).number_format = '0.00'
        # sheet.cell(row=2,column=6).number_format = '0.00'
