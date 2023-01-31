from libs import foundCsv

# 作成するエクセルのファイル名
excel_filename = './excel_data1.xlsx'

# csvデータがあるところまでのPATH
Path = './data1/'

foundCsv.simple_check(Path,excel_filename)

print('End')