import cv2
import os
import re
import shutil
import numpy as np
from numpy.lib.stride_tricks import sliding_window_view
import pandas as pd
from operator import itemgetter
from pathlib import Path
from pdf2image import convert_from_path
from natsort import natsorted
import unicodedata
import tkinter as tk
from tkinter import ttk
import unicodedata
import json
import requests
import urllib3
from normalization import *
from tools import *


class Convert_to_df():
    def __init__(self, original_img_path, api_base_url, append_log):
        self.original_img_path = original_img_path
        self.save_dir = None
        self.api_base_url = api_base_url.rstrip("/")  # URL末尾の/を除去
        self.append_log = append_log
        self.normalization = Normalization()
        self.tools = Tools(None)

        # 自己署名証明書の警告を抑制（開発用）
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


    def text_rec_gpu(self, dir_path):
        # dir_path: 送信する画像が入っているディレクトリ
        self.append_log("APIサーバに接続中...\n")

        # 送信ファイルの準備（自然順ソート）
        files = []
        for filename in natsorted(os.listdir(dir_path)):
            local_file = os.path.join(dir_path, filename)
            if os.path.isfile(local_file):
                files.append(("files", (filename, open(local_file, "rb"), "image/png")))
                self.append_log(f"送信: {filename}\n")

        if not files:
            self.append_log("[ERROR] 送信する画像が見つかりません\n")
            return None

        # APIにPOSTリクエスト
        self.append_log("リクエスト送信中...\n")
        try:
            res = requests.post(
                f"{self.api_base_url}/ocr",
                files=files,
                verify=False  # 自己署名証明書の場合
            )
            res.raise_for_status()
        except requests.RequestException as e:
            self.append_log(f"[ERROR] API通信失敗: {e}\n")
            return None

        # レスポンス処理
        try:
            result_json = res.json()
        except json.JSONDecodeError:
            self.append_log("[ERROR] APIからのレスポンスがJSONではありません\n")
            return None

        # 正規化（必要なら）
        normalized_text = json.dumps(result_json, ensure_ascii=False)
        normalized_text = unicodedata.normalize('NFKC', normalized_text)

        return json.loads(normalized_text)


    # cv2.imreadの日本語パス対応 + Path対応
    def imread(self,filepath: Path, flags=cv2.IMREAD_COLOR, dtype=np.uint8):
        try:
            n = np.fromfile(str(filepath), dtype)
            img = cv2.imdecode(n, flags)
            return img
        except Exception as e:
            print(f"読み込みエラー: {e}")
            return None

    # cv2.imwriteの日本語パス対応 + Path対応
    def imwrite(self,filepath: Path, img, params=None):
        try:
            ext = filepath.suffix
            result, n = cv2.imencode(ext, img, params)

            if result:
                with filepath.open('wb') as f:
                    n.tofile(f)
                return True
            else:
                return False
        except Exception as e:
            print(f"書き込みエラー: {e}")
            return False

    #横線を検出する関数
    def horizontal_line(self):
        bgr_img = self.imread(self.original_img_path)
        print(self.original_img_path)
        img = bgr_img[:, :, 0]
        height, width = img.shape
        minlength = width * 0.7
        gap = 0
        judge_img = cv2.bitwise_not(img)
        
        # 検出しやすくするために二値化
        # th, judge_img = cv2.threshold(judge_img, 128, 255, cv2.THRESH_BINARY)
        th, judge_img = cv2.threshold(judge_img, 64, 255, cv2.THRESH_BINARY)
        
        # 平行な横線のみ取り扱うため、ふるいにかける
        lines = []
        lines = cv2.HoughLinesP(judge_img, rho=1, theta=np.pi/360, threshold=100, minLineLength=minlength, maxLineGap=gap)
        
        line_list = []
        for line in lines:
            x1, y1, x2, y2 = line[0]
            # 傾きが threshold_slope px以内の線を横線と判断
            threshold_slope = 3
            if abs(y1 - y2) < threshold_slope:
                whiteline = 3
                lineadd_img = cv2.line(bgr_img, (line[0][0], line[0][1]), (line[0][2], line[0][3]), (0, 0, 255), whiteline)
                x1 = line[0][0]
                y1 = line[0][1]
                x2 = line[0][2]
                y2 = line[0][3]
                line = (x1, y1, x2, y2)
                line_list.append(line)
        
        # y座標をキーとして並び変え
        line_list.sort(key=itemgetter(1, 0, 2, 3))
        
        hoz_line = 0
        hoz_line_list = []
        y1 = 0
        for line in line_list:
            judge_y1 = line[1]
            # ほぼ同じ位置の横線は除外
            if abs(judge_y1 - y1) < 2 and hoz_line_list != []:
                y1 = judge_y1
            else: 
                y1 = judge_y1
                hoz_line = hoz_line + 1
                hoz_line_list.append(line)
        
        line_list = pd.DataFrame(hoz_line_list)
        return line_list[1].tolist() #横線のy座標のみ取得

    #縦線を検出する関数
    def vertical_line(self):
        bgr_img = self.imread(self.original_img_path)
        img = bgr_img[:,:,0]
        height, width = img.shape
        minlength = height * 0.7
        gap = 5
        judge_img = cv2.bitwise_not(img)
        
        # 検出しやすくするために二値化
        # th, judge_img = cv2.threshold(judge_img, 128, 255, cv2.THRESH_BINARY)
        th, judge_img = cv2.threshold(judge_img, 64, 255, cv2.THRESH_BINARY)
        
        lines = []
        lines = cv2.HoughLinesP(judge_img, rho=1, theta=np.pi/360, threshold=100, minLineLength=minlength, maxLineGap=gap)
        
        line_list = []
        for line in lines:
            x1, y1, x2, y2 = line[0]
            # 傾きが threshold_slope px以内の線を縦線と判断
            threshold_slope = 3
            if abs(x1 - x2) < threshold_slope:
                whiteline = 3
                lineadd_img = cv2.line(bgr_img, (line[0][0], line[0][1]), (line[0][2], line[0][3]), (0, 0, 255), whiteline)
                x1 = line[0][0]
                y1 = line[0][1]
                x2 = line[0][2]
                y2 = line[0][3]
                line = (x1, y1, x2, y2)
                line_list.append(line)
        
        # y座標をキーとして並び変え(今回は必要ない)
        line_list.sort(key=itemgetter(0, 1, 2, 3))
        
        ver_line = 0
        ver_line_list = []
        x1 = 0
        for line in line_list:
            judge_x1 = line[0]
            # ほぼ同じ位置の縦線は除外
            if abs(judge_x1 - x1) < 2 and ver_line_list != []:
                x1 = judge_x1
            else:
                x1 = judge_x1
                ver_line = ver_line + 1
                ver_line_list.append(line)
        
        line_list = pd.DataFrame(ver_line_list)
        return line_list[0].tolist()

    #検出した縦線と横線から格子点を割り出す関数（行ごとに格子点のy座標が等しい）
    def grid_points(self):
        horizontal_list = self.horizontal_line()
        vertical_list = self.vertical_line()
        grid_arr = []
        
        for i in horizontal_list:
            small = []
            for j in vertical_list:
                small.append((j,i))
            grid_arr.append(small)
            
        return grid_arr

    #行列から4要素ずつスライス
    def mk_grid_list(self,arr,size=2):
        h = len(arr)
        w = len(arr[0]) if h > 0 else 0
        windows = []
        for i in range(h - size + 1):
            for j in range(w - size + 1):
                window = [arr[i + di][j + dj] for di in range(size) for dj in range(size)]
                windows.append(window)
        
        grid_list = []
        
        for i in windows:
            k = (i[0],i[1],i[2],i[3])
            grid_list.append(k)

        return grid_list

    #画像をセルごとに分けて保存
    def split_img(self, g_list):
        # Pathオブジェクトにする
        img_path = Path(self.original_img_path)

        # 画像読み込み
        bgr_img = self.imread(img_path)
        if bgr_img is None:
            raise FileNotFoundError(f"画像が読み込めません: {img_path}")
        
        # 0チャンネルだけ取り出す
        img = bgr_img[:, :, 0]

        # 保存先ディレクトリを作成（親ディレクトリ + "img")
        if self.save_dir:
            shutil.rmtree(self.save_dir)
        parent_dir = img_path.parent.parent
        self.save_dir = parent_dir / "img"
        self.save_dir.mkdir(exist_ok=True)

        black_ratio_list = []  # 黒の割合を格納するリスト

        for k in range(len(g_list)):
            timg = img[g_list[k][0][1]+2:g_list[k][3][1], g_list[k][0][0]+2:g_list[k][3][0]]
            
            # 黒の割合を計算（0:黒, 255:白）
            # 2値化して黒画素数をカウント
            _, bin_img = cv2.threshold(timg, 128, 255, cv2.THRESH_BINARY)
            black_pixels = np.sum(bin_img == 0)
            total_pixels = bin_img.size
            black_ratio = black_pixels / total_pixels
            black_ratio_list.append(black_ratio)
            # パディングのサイズ（上, 下, 左, 右）
            top, bottom, left, right = 35, 35, 40, 40

            # 白色のBGR値（OpenCVはBGR形式）
            white = [255, 255, 255]

            # パディングを追加
            padded_image = cv2.copyMakeBorder(
                timg, top, bottom, left, right,
                borderType=cv2.BORDER_CONSTANT,
                value=white
            )
            # 保存先をsave_dir配下に修正
            self.save_path = self.save_dir / f"img{k+1}.png"
            result = self.imwrite(self.save_path, padded_image)
            if not result:
                print(f"書き込みに失敗しました: {self.save_path}")

        return black_ratio_list  # 黒の割合リストを返す
       
                
    #画像の縦横のセルの数をカウント($x\times y$のx,yを求める関数)
    def count_2D_cell(self,g_list):
        i = 1
        j = 0
        while g_list[j][0][1] == g_list[j+1][0][1]:
            i += 1
            j += 1
        return (len(g_list)//i,i)

        
    #画像から格子点の組み合わせ、セル内のテキストの二次元のリストを作成する
    def img2list(self):
        arr = self.grid_points()
        g_list = self.mk_grid_list(arr, size=2)
        black_ratio_list = self.split_img(g_list)
        
        black_density_flag = np.empty(len(g_list), dtype=bool)
        file_path_l = natsorted(os.listdir(self.save_dir))
        for i in range(len(file_path_l)):
            if black_ratio_list[i] > 0.1:
                black_density_flag[i] = True
            else:
                black_density_flag[i] = False

        # OCR処理(リモートGPUマシンで行う)
        text_arr = np.array(self.text_rec_gpu(self.save_dir), dtype=object)

        shape = self.count_2D_cell(g_list)
        text_arr = text_arr.reshape(shape)
        black_density_flag = black_density_flag.reshape(shape)

        return text_arr, black_density_flag
    
    def split_half_arrays(self,text_arr):
        mid = text_arr.shape[1] // 2

        text_arr_left = text_arr[:, :mid]
        text_arr_right = text_arr[:, mid:]

        return text_arr_left, text_arr_right


    def arr2df(self, text_arr_left, text_arr_right):
        # ---- 左側 ----
        text_df_left = pd.DataFrame(text_arr_left)

        # 1行目を空白除去＋正規化してヘッダに設定
        text_df_left.columns = text_df_left.iloc[0].apply(
            lambda x: unicodedata.normalize("NFKC", str(x)).replace(" ", "") if isinstance(x, str) else x
        )
        text_df_left = text_df_left[1:].reset_index(drop=True)

        if "備考" not in text_df_left.columns:
            text_df_left["備考"] = ""

        text_df_left.loc[text_df_left["手配"] != '', "備考"] = (
            text_df_left.loc[text_df_left["手配"] != '', "備考"].astype(str) + " 御支給品"
        )

        # ---- 右側 ----
        text_df_right = pd.DataFrame(text_arr_right)

        text_df_right.columns = text_df_right.iloc[0].apply(
            lambda x: unicodedata.normalize("NFKC", str(x)).replace(" ", "") if isinstance(x, str) else x
        )
        text_df_right = text_df_right[1:].reset_index(drop=True)

        if "備考" not in text_df_right.columns:
            text_df_right["備考"] = ""

        text_df_right.loc[text_df_right["手配"] != '', "備考"] = (
            text_df_right.loc[text_df_right["手配"] != '', "備考"].astype(str) + "  御支給品"
        )

        # ---- 空行を挿入する処理 ----
        empty_row = pd.DataFrame([[ '' for _ in text_df_left.columns ]], columns=text_df_left.columns)

        df_with_empty_left = pd.DataFrame(columns=text_df_left.columns)
        for _, row in text_df_left.iterrows():
            df_with_empty_left = pd.concat([df_with_empty_left, pd.DataFrame([row]), empty_row], ignore_index=True)

        df_with_empty_right = pd.DataFrame(columns=text_df_right.columns)
        for _, row in text_df_right.iterrows():
            df_with_empty_right = pd.concat([df_with_empty_right, pd.DataFrame([row]), empty_row], ignore_index=True)

        return text_df_left, text_df_right
    
    def out_df(self, df):
        out = pd.DataFrame()
        out = pd.concat([out, df[['購入先', '名称','仕様','数量']]])
        out.columns = ["メーカー","名称","型式・仕様","数量"]

        # 「仕様」と「備考」を結合して型式・仕様列にする
        out["型式・仕様"] = df["仕様"].str.cat(df["備考"], na_rep='')

        out.replace(r'^\s*$', pd.NA, regex=True, inplace=True)
        out = out.dropna(how="all")

        # 正規化と空白削除（型式・仕様以外）
        for col in out.columns:
            if col != "型式・仕様":
                out[col] = out[col].map(self.normalization.normalize_and_strip)

        def process_spec(val):
            if pd.isna(val):
                return [""]
            if isinstance(val, str):
                parts = val.split()
            elif isinstance(val, list):
                parts = val
            else:
                parts = [str(val)]

            left, right = self.tools.separate_words(parts)
            if left and right:
                return [" ".join(left), " ".join(right)]
            else:
                return [" ".join(parts)]

        processed_rows = []
        for _, row in out.iterrows():
            spec_vals = process_spec(row["型式・仕様"])

            # 1行目（必ず出力）
            first_row = row.copy()
            first_row["型式・仕様"] = spec_vals[0]
            processed_rows.append(first_row)

            # 2行目（空行を必ず追加）
            second_row = pd.Series({col: "" for col in out.columns})
            if len(spec_vals) > 1:
                # 型式・仕様だけ2段目を入れる
                second_row["型式・仕様"] = spec_vals[1]
            processed_rows.append(second_row)

        df_out = pd.DataFrame(processed_rows, columns=out.columns).fillna(" ")

        return df_out


    