import glob
import os
import openpyxl
import pathlib


#excel_path = '/content/drive/MyDrive/テスト/Book1.xlsx'
excel_path = ''
contents = []

# ブックの取得
actBook = openpyxl.load_workbook(excel_path)

# シート数分ループ
for actSheetName in actBook.sheetnames:

    # アクティブシートを取得
    actSheet = actBook[actSheetName]
    # 該当シートの最大行と最大列を取得
    maxRow = actSheet.max_row
    maxCol = actSheet.max_column
    # for 行変数 in シート変数.iter_rows(開始行,終了行,開始列,終了列)
    for row in actSheet.iter_rows(min_row=2, max_row=maxRow, min_col=1, max_col=maxCol):
    # １行分のテキストの宣言と初期化
        text = ''
        # for セル変数 in 列変数
        for cell in row:
            # セルを取得
            cellData = cell.value
            colChart = cell.column
            rowChart = cell.row

            if colChart == 1:
                if rowChart == 2:
                    text += f'(\"{cellData}\", '
                else:
                    text += f',(\"{cellData}\", '
            elif colChart == maxCol:
                text += f'\"{cellData}\")  '
            else:
                text += f'\"{cellData}\") , '

        contents.append(text)
        print(text)

# ブック変数.save(Excelファイルのパス)
    actBook.close
print(contents)

dir_path = '/content/drive/MyDrive/SQL'
if not os.excel_path.exists(dir_path):
# ディレクトリが存在しない場合、ディレクトリを作成する
    os.makedirs(dir_path)
text_path = pathlib.Path(f'{dir_path}/text.txt')
# text_path = pathlib.Path(file_path)
for content in contents:
    with open(text_path, mode='w', encoding='cp932', newline='\n', errors='ignore') as f:
        for content in contents:
            f.write(content)
f.close()