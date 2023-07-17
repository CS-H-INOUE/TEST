import os
from openpyxl import load_workbook

# ファイルパスの設定
directory = os.path.expanduser("~/Desktop/")
file_path = os.path.join(directory, "PythonScripts\Code\Inspection\p_inspection.xlsx")

# Excelファイルの読み込み
workbook = load_workbook(file_path)

# シート1の内容をコピーして新しいシートを作成
source_sheet = workbook["Sheet1"]

for i in range(2, 22):
    new_sheet_name = "Copy_Sheet{}".format(i)
    new_sheet = workbook.copy_worksheet(source_sheet)
    new_sheet.title = new_sheet_name

# ファイルを上書き保存
workbook.save(file_path)
