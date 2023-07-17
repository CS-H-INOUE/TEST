import qrcode
import openpyxl
import os
from datetime import date


def generate_qr_code(qr_text, qr_directory):
    """
    テキストを使用してQRコードを生成して保存する関数

    Args:
        qr_text (str): QRコードのテキスト
        qr_directory (str): QRコードの保存先ディレクトリ名

    Returns:
        None
    """
    # QRコードの保存先ディレクトリのパスを作成
    qr_directory_path = os.path.join(os.getcwd(), qr_directory)

    # QRコードの保存先ディレクトリが存在しない場合は作成
    if not os.path.exists(qr_directory_path):
        os.makedirs(qr_directory_path)

    # QRコードを生成
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=10,
        border=4,
    )
    qr.add_data(qr_text)
    qr.make(fit=True)

    # QRコードをPIL Imageオブジェクトに変換
    qr_image = qr.make_image(fill_color="black", back_color="white")

    # QRコードを保存
    qr_filename = f"qr_{qr_text}.png"
    qr_filepath = os.path.join(qr_directory_path, qr_filename)
    qr_image.save(qr_filepath)


def generate_qr_codes(excel_file, sheet_name, qr_directory):
    """
    エクセルファイルからデータを読み取り、QRコードを生成して保存する関数

    Args:
        excel_file (str): エクセルファイルのパス
        sheet_name (str): シート名
        qr_directory (str): QRコードの保存先ディレクトリ名

    Returns:
        None
    """
    # エクセルファイルを読み込み
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook[sheet_name]

    # 今日の日付を取得
    today = date.today().strftime("%Y%m%d")

    # QRコードを生成して保存
    for row in sheet.iter_rows(values_only=True):
        text = str(row[0])
        qr_text = f"中部産商_管理デバイスPC_購入日付{today}_管理対象者{text}"

        generate_qr_code(qr_text, qr_directory)

    # エクセルファイルを閉じる
    workbook.close()


# メインの処理
if __name__ == "__main__":
    # エクセルファイルのパスとシート名
    excel_file = 'data.xlsx'
    sheet_name = 'Sheet1'

    # QRコードの保存先ディレクトリ
    qr_directory = 'QRディレクトリ'

    # QRコード生成の関数を呼び出す
