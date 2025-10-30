import sys
import os
import pandas as pd
import re
from pathlib import Path
import fitz  # PyMuPDF

class Tools():
    def __init__(self, pdf_path):
        self.original_pdf_path = pdf_path

    def natural_sort_key(s):
        return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', s)]

    def pdf2img(self, zoom=4.0):
        """
        PDF → PNG 変換
        zoom=4.0 は 72dpi × 4 ≈ 288dpi 相当（pdf2image の dpi=300 に近い）
        """

        # One-File exe / One-Folder / 通常スクリプト 実行時のパス解決
        if getattr(sys, 'frozen', False):
            base_dir = Path(sys.executable).parent
        else:
            base_dir = Path(__file__).resolve().parent

        base_path = Path(self.original_pdf_path)
        pdf_dir = base_path.parent / "pdf"
        pdf_dir.mkdir(exist_ok=True)

        # PDF を開く
        doc = fitz.open(self.original_pdf_path)

        output_paths = []
        for i, page in enumerate(doc):
            # DPI 相当の拡大率を設定
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat, alpha=False)  # ★ alpha=False に変更

            if len(doc) == 1:
                output_path = pdf_dir / f"{base_path.stem}.png"
            else:
                output_path = pdf_dir / f"{base_path.stem}_page_{i+1}.png"

            pix.save(str(output_path))
            output_paths.append(str(output_path))

        doc.close()
        return output_paths
    
    def separate_words(self, text_l):
        S = len("".join(text_l))
        if S > 15:
            c = 0
            count_l = []
            for t in text_l:
                c += len(t)  # 文字数を加算
                count_l.append(c)

            res = float("inf")
            ind = -1

            # 全体の半分に最も近い位置を探す
            for i, cnt in enumerate(count_l):
                if abs(S/2 - cnt) < res:
                    res = abs(S/2 - cnt)
                    ind = i

            return text_l[:ind+1], text_l[ind+1:]
        else:
            # 15文字以下ならそのまま返す
            return text_l, []



    def df2excel(self, df, name):
        with pd.ExcelWriter(name, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Sheet1', index=False)

            workbook  = writer.book
            worksheet = writer.sheets['Sheet1']

            # 和文フォント指定
            font_format = workbook.add_format({
                'font_name': 'ＭＳ 明朝',
                'font_size': 11
            })

            # 列幅とフォントを適用
            col_count = len(df.columns)
            for i in range(col_count):
                col_letter = chr(ord('A') + i)
                worksheet.set_column(f'{col_letter}:{col_letter}', 20, font_format)