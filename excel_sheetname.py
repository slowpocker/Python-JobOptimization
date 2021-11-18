# シート名一括変更
import openpyxl as px


# ブックの取得
# openpyxl.load_workbook('Excelファイルのパス')
wb = px.load_workbook('path')

# すべてのシートに特定の文字列を付与する
for ws in wb:
    ws.title = "特定の文字列" + ws.title

# すべてのシートに連番を付ける
for i, ws in enumerate(wb):
    ws.title = str(i+1) + "_" + ws.title


# すべてのシートの特定の文字列を置換する
for ws in wb:
    ws.title = ws.title.replace('特定の文字列', '置換後の文字列')


wb.sheetnames  # シート名の確認

wb.save('path')
