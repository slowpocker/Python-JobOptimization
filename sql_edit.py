import os
import shutil

# コピー元ファイルの情報
ori_workname = ''
ori_formtitle = ''
ori_workcode = ''
ori_id = ''
ori_path = '/content/drive/MyDrive/***'

# 作成するファイル情報
lists = [
        ['***', '***', '***', '***'],
        ['***', '***', '***', '***'],
        ['***', '***', '***', '***'],
]

for list in lists:
    workname = list[0]
    formtitle = list[1]
    workcode = list[2]
    id = list[3]
    path = f'/content/drive/MyDrive/***/SQL/{formtitle}'

    if not os.path.exists(path):
        # ディレクトリが存在しない場合、ディレクトリを作成する
        os.makedirs(path)

    # コピー元ファイルの情報をリストとして取得する
    files = os.listdir(ori_path)
    ori_fileList = [f for f in files if os.path.isfile(
        os.path.join(ori_path, f))]

    for ori_file in ori_fileList:
        # コピー元ファイルを別フォルダ内にコピーする
        src = f'{ori_path}/{ori_file}'
        dest = f'{path}/{ori_file}'
        shutil.copyfile(src, dest)

        # コピーしたファイル名を取得する
        before_file = os.path.basename(dest)

        # コピーしたファイル名の文字置換
        after_file = before_file.replace(ori_formtitle, formtitle).replace(
            ori_workcode, workcode).replace(ori_id, id)
        os.rename(dest, f'{path}/{after_file}')

        # ファイル名に合わせてパスを更新
        dest = f'{path}/{after_file}'

        # 書き込み用リストを作成
        contents = []

        # ファイル情報を1行ごとに分割したリストとして取得して、作成するファイル情報に合わせて文字置換、contentsに代入していく
        with open(dest, mode='r', encoding='cp932', newline='\n', errors='ignore') as f:
            data_lines = f.readlines()
            for data_line in data_lines:
                data_line = data_line.replace(ori_formtitle, formtitle).replace(
                    ori_workcode, workcode).replace(ori_id, id).replace(ori_workname, workname)
                contents.append(data_line)
        f.close()

        # contentsに代入された1行ごとのテキストをファイルに書き込んでいく
        with open(dest, mode='w', encoding='cp932', newline='\n', errors='ignore') as f:
            for content in contents:
                f.write(content)
        f.close()
