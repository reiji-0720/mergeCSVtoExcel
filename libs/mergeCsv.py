import os
import csv
import openpyxl

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
                sheet.cell(row = r,column = c).value = v
                c += 1
            r += 1 # row
