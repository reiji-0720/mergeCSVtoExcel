from libs import foundCsv

# 作成するエクセルのファイル名
excel_filename = './test.xlsx'

# csvデータがあるところまでのPATH
Path = './sampledata/'

foundCsv.simple_check(Path,excel_filename)

print('End')