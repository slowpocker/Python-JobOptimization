# シート名一括変更
import openpyxl
import glob
import os

before_wordlist = []
after_wordlist = []
# 対象ファイルのパス
path = '/content/drive/MyDrive/excel'
# 対象ファイル種別
fileType = '*.xlsx'
files = os.listdir(path)
fileList = [f for f in files if os.path.isfile(os.path.join(path, f))]

book_path = glob.glob(os.path.join(path,fileType ))
for book in book_path:
    # ブックの取得
    
    wb = openpyxl.load_workbook(file_path)

# すべてのシートの特定の文字列を置換する
for ws in wb:
    for i in range(len(before_wordlist)):
        ws.title = ws.title.replace(before_wordlist[i], after_wordlist[i])
wb.sheetnames  # シート名の確認
wb.save(path)

# # すべてのシートに特定の文字列を付与する
# for ws in wb:
#     ws.title = "特定の文字列" + ws.title

# # すべてのシートに連番を付ける
# for i, ws in enumerate(wb):
#     ws.title = str(i+1) + "_" + ws.title


