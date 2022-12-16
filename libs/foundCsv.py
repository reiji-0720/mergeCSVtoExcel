import os
import openpyxl
from .import mergeCsv

# ディレクトリ探索用関数

def simple_check(path,excel_filename):
    i = 0
    wb = openpyxl.Workbook()
    for pathname,dirname, filenames in os.walk(path):
        for filename in filenames:
            if filename.endswith('.csv'): # フィルタ処理 csvのみ処理
                filename = os.path.join(*pathname,filename)
                # print(os.path.join(pathname,filename))
                print(filename)
                mergeCsv.merge(filename,i,path,wb)
                i += 1

    excel_filename_only = os.path.basename(excel_filename) # 拡張子無しのファイル名取得
    wb.save(filename = excel_filename_only)
    wb.close()
