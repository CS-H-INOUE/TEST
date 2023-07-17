import openpyxl
import pandas as pd

# ワークブックを読み込む
xlsx = openpyxl.load_workbook('p_inspection.xlsx')

# シート名の取得
sheet_names = xlsx.sheetnames

# 出力データを格納するリスト
output_data = []

# シートごとに処理
for sheet_name in sheet_names:
    # シートを取得
    sheet = xlsx[sheet_name]

    # A列のデータを取得
    column_a = [cell.value for cell in sheet['A']]

    # B列のデータを取得
    column_b = [cell.value for cell in sheet['B']]

    # C列のデータを取得
    column_c = [cell.value for cell in sheet['C']]

    # D列のデータを取得
    column_d = [cell.value for cell in sheet['D']]

    # 最大の列の要素数を取得
    max_length = max(len(column_b), len(column_c), len(column_d))

    # データを抽出してoutput_dataに追加
    for i in range(max_length):
        b_value = column_b[min(i, len(column_b)-1)]
        c_value = column_c[min(i, len(column_c)-1)]
        d_value = column_d[min(i, len(column_d)-1)]
        output_data.append([b_value, c_value, d_value])

    # シート名のデータをA列に追加
    output_data.extend([[sheet_name]] * max_length)

# DataFrameに変換
output_df = pd.DataFrame(output_data, columns=['B列', 'C列', 'D列'])

# CSVファイルとして保存
output_df.to_csv('output.csv', index=False)
