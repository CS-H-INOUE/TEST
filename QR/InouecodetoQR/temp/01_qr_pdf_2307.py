import os
import glob
from PIL import Image
from reportlab.lib.pagesizes import A6
from reportlab.pdfgen import canvas
import os.path
import qrcode
from datetime import date
import datetime
from PIL import Image, ImageDraw, ImageFont
import shutil
import sys

def check_daily_folder():
    # 実行ファイルの直下パスを取得
    base_path = os.path.dirname(os.path.abspath(__file__))

    # 当日の日付を取得
    today = datetime.date.today()
    folder_name = today.strftime("%Y%m%d")

    # 当日の日付フォルダを確認
    folder_path = os.path.join(base_path, folder_name)
    if not os.path.exists(folder_path):
        print(f"CheckFolderError!!!: {folder_path}")
        sys.exit(0)

    # 実行ファイルの直下パスを返す
    return folder_path

def print_images_as_a6(directory):
    """
    指定されたディレクトリ内の画像ファイルをA6サイズで印刷する関数
    :param directory: 画像ファイルが存在するディレクトリのパス
    """
    # ディレクトリの存在を確認
    if not os.path.isdir(directory):
        print(f"指定されたディレクトリ '{directory}' は存在しません。")
        return

    # ディレクトリ内の画像ファイルを取得
    image_files = glob.glob(os.path.join(directory, "*.jpg")) + glob.glob(os.path.join(directory, "*.png"))

    # 画像ファイルが存在しない場合は終了
    if not image_files:
        print("ディレクトリ内に画像ファイルが見つかりません。")
        return

    # PDFを作成し、画像ファイルを追加
    pdf = canvas.Canvas("printout.pdf", pagesize=A6)
    for image_file in image_files:
        try:
            image = Image.open(image_file)
            # 画像サイズをA6にリサイズ
            image = resize_image_to_a6(image)
            # PDFに画像を追加
            pdf.drawImage(image_file, 0, 0, width=A6[0], height=A6[1])
            pdf.showPage()
        except (IOError, OSError):
            print(f"画像ファイル '{image_file}' の処理中にエラーが発生しました。")
            continue

    # PDFを保存して印刷
    pdf.save()
    print(f"PDFファイル 'printout.pdf' が作成されました。印刷してください。")

def resize_image_to_a6(image):
    """
    画像をA6サイズにリサイズする関数
    :param image: PILのImageオブジェクト
    :return: リサイズされたImageオブジェクト
    """
    a6_size = (105, 148)  # A6サイズ（mm）
    image.thumbnail(a6_size)
    return image

# テスト用例
if __name__ == "__main__":
    directory_path = check_daily_folder()
    print_images_as_a6(directory_path)
