# excelのコピー
import glob
import os
import shutil
import openpyxl

ori_path = '/content/drive/MyDrive/original/excel'
path = '/content/drive/MyDrive/excel'
fileType = '*.xlsx'
original_file_contents = []
lists = []


for list in lists:
    # コピー元ファイルの情報をリストとして取得する
    files = os.listdir(ori_path)
    ori_fileList = [f for f in files if os.path.isfile(os.path.join(ori_path, f))]
    for ori_file in ori_fileList:
        # コピー元ファイルを別フォルダ内にコピーする
        src = f'{ori_path}/{ori_file}'
        dest = f'{path}/{ori_file}'
        shutil.copyfile(src, dest)

        after_file_name = ori_file.replace(original_file_contents[0], list[0])
        os.rename(dest, f'{path}/{after_file_name}')
        # ファイル名に合わせてパスを更新
        dest = f'{path}/{after_file_name}'

# 「①対象ファイルのパス」配下にある全てのExcelファイルのパスを出力
books_path = glob.glob(os.path.join(path, fileType))
print(books_path)
