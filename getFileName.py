import glob
from natsort import natsorted
import os
import openpyxl

# ファイルを取得
dir_path = '/content/drive/AAA/**/*'
excel_path = '/content/drive/MyDrive/excel/Book2.xlsx'
# 入力列
column = 1
# dir_path = '/content/drive/AAA/**/*.sql'
files = natsorted(glob.glob(dir_path))
print(files)
# files = natsorted(glob.glob('dir/*.txt'))

file_name_contents = []

for i in files:
    # print('        パス名:' + i)
    # print('ディレクトリ名:' + os.path.split(i)[0])
    print('    ファイル名:' + os.path.split(i)[1])
    file_name_contents.append(os.path.split(i)[1])
    # print('-'*20)

print(file_name_contents)
# 配列の要素数を取得
num = len(file_name_contents)

# file名をエクセルの列に書き出す
actBook = openpyxl.load_workbook(excel_path)
actSheet = actBook.worksheets[0]

for i in range(0, len(file_name_contents) - 1, 1):
    actSheet.cell(column, i + 1).value = file_name_contents[i]
actSheet.save()