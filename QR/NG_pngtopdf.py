import os
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from PIL import Image

def create_pdf_from_png(png_directory, pdf_file):
    """
    PNGファイルを1つのPDFファイルに変換・結合する関数

    Args:
        png_directory (str): PNGファイルが保存されているディレクトリのパス
        pdf_file (str): 作成するPDFファイルのパス

    Returns:
        None
    """
    # PNGファイルのリストを取得
    png_files = [file for file in os.listdir(png_directory) if file.endswith('.png')]

    # PDFキャンバスを作成
    pdf_canvas = canvas.Canvas(pdf_file, pagesize=A4)

    # PNGファイルを1つずつPDFに変換・結合
    for png_file in png_files:
        png_path = os.path.join(png_directory, png_file)

        # PNGファイルを開く
        image = Image.open(png_path)

        # 画像サイズをA4用紙に合わせる
        image_width, image_height = image.size
        if image_width > A4[0] or image_height > A4[1]:
            image.thumbnail((A4[0] - 10, A4[1] - 10), Image.ANTIALIAS)

        # PNGファイルをPDFに描画
        pdf_canvas.drawImage(image, 0, 0, width=image.width, height=image.height)

        # ページを追加
        pdf_canvas.showPage()

    # PDFファイルを保存
    pdf_canvas.save()

    print(f"PDFファイル {pdf_file} が作成されました。")

if __name__ == "__main__":
    # PNGファイルが保存されているディレクトリのパス
    png_directory = 'QRディレクトリ'

    # 作成するPDFファイルのパス
    pdf_file = 'output.pdf'

    # PNGファイルを1つのPDFに変換・結合
    create_pdf_from_png(png_directory, pdf_file)

    # PNGファイルを削除
    png_files = [file for file in os.listdir(png_directory) if file.endswith('.png')]
    for png_file in png_files:
        png_path = os.path.join(png_directory, png_file)
        os.remove(png_path)

    print("PNGファイルが削除されました。")
