import tkinter as tk
from tkinter import ttk, filedialog, messagebox,Toplevel,Text,Entry,Button,simpledialog
from PIL import Image, ImageTk
import json
import os
import sys
import re
import copy
import shutil
import threading
import time
import configparser
from pdf2df import *
from tools import *
from normalization import *

def dict_keys_to_str(d):
    return {str(k): v for k, v in d.items()}

class ShortcutMixin:
    def __init__(self,master):
        # ショートカットの有効無効変数
        self.shortcut_enabled = tk.BooleanVar(value=True)
        self.bind_shortcuts()

    def bind_shortcuts(self):
        self.bind_all("<Control-s>", self.on_save_shortcut)
        self.bind_all("<Up>", self.on_left_arrow)
        self.bind_all("<Down>", self.on_right_arrow)
        self.bind_all("<Return>",self.on_enter)

    def unbind_shortcuts(self):
        self.unbind_all("<Control-s>")
        self.unbind_all("<Up>")
        self.unbind_all("<Down>")
        self.unbind_all("<Return>")
        
    def toggle_shortcuts(self):
        if self.shortcut_enabled.get():
            self.bind_shortcuts()
        else:
            self.unbind_shortcuts()

    def on_save_shortcut(self, event=None):
        # 実際の保存処理は親クラスのメソッドを呼ぶ
        self.save_xlsx()
        return "break"  # これでイベント伝搬を止める

class PDF2xlsxApp(tk.Tk,ShortcutMixin):
    def __init__(self):
        super().__init__()
        ShortcutMixin.__init__(self,self)
        self.title("pdf2xlsx")
        self.geometry("900x600")
        self.project_data = None
        self.project_path = None
        self.current_dir = "./proj/" # 初期ディレクトリ

        self.create_menu()

        self.main_frame = ttk.Frame(self)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        self.tree_frame = ttk.Frame(self.main_frame)
        self.image_frame = ttk.Frame(self.main_frame)

        self.tree_frame.pack(side=tk.LEFT, fill=tk.Y)
        self.tree_frame.pack_propagate(False)
        self.tree_frame.config(width=275)
        self.image_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.tree_scrollbar = ttk.Scrollbar(self.tree_frame, orient="vertical")
        self.tree = ttk.Treeview(self.tree_frame, yscrollcommand=self.tree_scrollbar.set)
        self.tree.bind("<Up>", lambda e: "break") #ツリーのデフォルトショートカットキーを無効化し、自分でカスタマイズしたショートカットとの競合を避ける
        self.tree.bind("<Down>", lambda e: "break")
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.tree_scrollbar.config(command=self.tree.yview)
        self.tree_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 現在表示中の画像パス
        self.current_image_path = None  
        
        # 編集フラグ
        #self.is_modified = False 
        
        self.log_window = None
        
        #未保存時の確認ダイアログ(現在は使用しない)
        #self.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # 起動時はツリービューを表示しない（コメントアウト）
        # self.build_treeview()    
        
        #リモート先の情報
        config = configparser.ConfigParser()
        config.read("settings.ini",encoding="utf-8")
        self.api_base_url = config.get("settings","api_url")
                
        self.convert_to_df_class = Convert_to_df(
            None,
            self.api_base_url,
            self.append_log
        )
        self.tools = Tools(pdf_path=None)
        self.normal = Normalization()
        
    def reset_state(self):
        # Notebookの破棄
        if getattr(self, "notebook", None) is not None and self.notebook.winfo_exists():
            self.notebook.destroy()
        self.notebook = None

        # Treeviewのクリア
        if getattr(self, "tree", None) is not None and self.tree.winfo_exists():
            self.tree.delete(*self.tree.get_children())

        # データ構造の初期化
        self.project_data_all = {}
        self.page_tab_map = {}
        self.pdf_path = None


        
    def build_treeview(self, tree_widget, path):
        for item in tree_widget.get_children():
            tree_widget.delete(item)

        tree_widget.insert('', 'end', text="🔼 ../", values=[".."], open=False)

        try:
            entries = os.listdir(path)
            dirs = sorted([e for e in entries if os.path.isdir(os.path.join(path, e))], key=Tools.natural_sort_key)
            files = sorted([e for e in entries if os.path.isfile(os.path.join(path, e))], key=Tools.natural_sort_key)
            sorted_entries = dirs + files

            for entry in sorted_entries:
                full_path = os.path.join(path, entry)
                if os.path.isdir(full_path):
                    display_name = f"📁 {entry}"
                else:
                    display_name = f"📄 {entry}"
                tree_widget.insert('', 'end', text=display_name, values=[full_path])
        except PermissionError:
            messagebox.showerror("エラー", f"{path} にアクセスできません。")


        
    #ツリービューで画像が選択されたときに表示させる    
    def on_tree_select_for_page(self, event, tree_widget, page_data):
        selected_item = tree_widget.selection()
        if not selected_item:
            return
        selected = tree_widget.item(selected_item[0], 'values')[0]
        if os.path.isfile(selected) and selected.lower().endswith(('.png', '.jpg', '.jpeg')):
            self.load_image_for_page(selected, page_data, page_data["canvas"], page_data["text_entry"])


    
    #上下の矢印キーを押したときにツリービューで画像が選択されていれば画像を表示する
    def on_tree_selection_change(self, event):
        selected_item = self.tree.selection()
        if not selected_item:
            return

        selected_path = self.tree.item(selected_item[0], 'values')[0]

        # 画像ファイルだけ表示
        if os.path.isfile(selected_path) and selected_path.lower().endswith(('.png', '.jpg', '.jpeg')):
            self.load_image(selected_path)
        else:
            self.image_canvas.delete("all")  # 画像以外のときはキャンバスをクリア
            self.current_image_path = None
            
    def on_tab_changed(self, event):
        selected_tab = event.widget.select()
        page_data = self.page_tab_map[selected_tab]

        # 既存のTreeview(self.tree)の中身を切り替え
        self.build_treeview(self.tree, page_data["page_dir"])

        # Treeview選択イベントをこのページ用にバインド
        self.tree.bind(
            "<<TreeviewSelect>>",
            lambda e: self.on_tree_select_for_page(e, self.tree, page_data)
        )


    def on_enter(self, event):
        try:
            if self.text_entry.get().strip():
                self.register_text()
        except AttributeError:
            pass  # text_entry がまだ存在しない場合は何もしない


    def on_left_arrow(self, event):
        #上矢印キーで前の画像へ
        self.select_previous_image()

    def on_right_arrow(self, event):
        #下矢印キーで次の画像へ
        self.select_next_image()
        
    def select_previous_image(self):
        # 現在選択されているアイテムを取得
        current = self.tree.focus()
        prev = self.tree.prev(current)
        if prev:
            self.tree.selection_set(prev)
            self.tree.focus(prev)
            self.tree.see(prev)
            self.on_tree_select(None)

    def select_next_image(self):
        # 現在選択されているアイテムを取得
        current = self.tree.focus()
        next = self.tree.next(current)
        if next:
            self.tree.selection_set(next)
            self.tree.focus(next)
            self.tree.see(next)
            self.on_tree_select(None)
    
    def show_next_image(self, event=None):
        items = self.tree.get_children()
        if not items:
            return

        selected_item = self.tree.selection()
        if not selected_item:
            return

        try:
            current_index = items.index(selected_item[0])
        except ValueError:
            return

        next_index = current_index + 1

        while next_index < len(items):
            next_item = items[next_index]
            try:
                selected = self.tree.item(next_item, 'values')[0]
                if os.path.isfile(selected) and selected.lower().endswith(('.png', '.jpg', '.jpeg')):
                    self.tree.selection_set(next_item)
                    self.tree.see(next_item)  # 選択箇所にスクロール
                    self.load_image(selected)
                    return
            except PermissionError:
                messagebox.showerror("アクセス拒否", f"{selected} にアクセスできません。")
                return
            next_index += 1

    def show_prev_image(self, event=None):
        items = self.tree.get_children()
        if not items:
            return

        selected_item = self.tree.selection()
        if not selected_item:
            return

        try:
            current_index = items.index(selected_item[0])
        except ValueError:
            return

        prev_index = current_index - 1

        while prev_index >= 0:
            prev_item = items[prev_index]
            try:
                selected = self.tree.item(prev_item, 'values')[0]
                if os.path.isfile(selected) and selected.lower().endswith(('.png', '.jpg', '.jpeg')):
                    self.tree.selection_set(prev_item)
                    self.tree.see(prev_item)  # 選択箇所にスクロール
                    self.load_image(selected)
                    return
            except PermissionError:
                messagebox.showerror("アクセス拒否", f"{selected} にアクセスできません。")
                return
            prev_index -= 1

                    
    def process_directory(self, parent, path):
        try:
            for entry in os.listdir(path):
                full_path = os.path.join(path, entry)
                if os.path.isdir(full_path):
                    node = self.tree.insert(parent, 'end', text=entry, values=[full_path])
                    self.tree.insert(node, 'end')  # ダミーを追加して展開可能に
        except PermissionError as e:
            print(f"アクセスできません: {path} - {e}")

    def on_tree_double_click(self, event):
        try:
            item_id = self.tree.selection()[0]
            selected = self.tree.item(item_id, 'values')[0]

            if selected == "..":
                parent = os.path.dirname(self.current_dir)
                if parent != self.current_dir:
                    self.current_dir = parent
                    self.build_treeview()
            elif os.path.isdir(selected):
                self.current_dir = selected
                self.build_treeview()
            elif os.path.isfile(selected):
                self.load_image(selected)
        except PermissionError:
            messagebox.showerror("アクセス拒否", f"{selected} にアクセスできません。")
        except Exception as e:
            messagebox.showerror("エラー", f"エラーが発生しました: {e}")
                
    def go_back_directory(self):
        parent = os.path.dirname(self.current_dir)
        if parent != self.current_dir:  # ルートディレクトリでないなら
            self.current_dir = parent
            self.build_treeview()
        
    def populate_file_tree(self, parent_node, path):
        try:
            items = sorted(os.listdir(path))
            for item in items:
                full_path = os.path.join(path, item)
                if os.path.isdir(full_path):
                    node = self.tree.insert(parent_node, 'end', text=f"📁 {item}", values=[full_path])
                    self.populate_file_tree(node, full_path)
                else:
                    self.tree.insert(parent_node, 'end', text=f"📄 {item}", values=[full_path])
        except PermissionError:
            pass


    def create_menu(self):
        self.menubar = tk.Menu(self)
        self.filemenu = tk.Menu(self.menubar, tearoff=0)
        self.filemenu.add_command(label="pdfファイルを開く", command=self.open_pdf)
        #self.filemenu.add_command(label="作業ファイルを開く", command=self.load_json) #作業中のtmp.jsonがあれば開く
        #self.filemenu.add_command(label="作業ファイルを保存", command=self.save_json) #作業中のtmp.jsonがあれば保存
        self.filemenu.add_command(label="Excelファイルに出力", command=self.save_all_pages)
        self.editmenu = tk.Menu(self.menubar, tearoff=0)
        self.editmenu.add_checkbutton(label="ショートカットキーを有効にする", variable=self.shortcut_enabled, command=self.toggle_shortcuts)
        self.menubar.add_cascade(label="ファイル", menu=self.filemenu)
        self.menubar.add_cascade(label="編集",menu=self.editmenu)  # メニューの中身は未実装
        self.menubar.add_cascade(label="設定")
        self.config(menu=self.menubar)
        
    def append_log(self, text):
        if self.log_text:
            self.log_text.config(state="normal")
            self.log_text.after(0, lambda: self.log_text.insert(tk.END, text))
            self.log_text.after(0, lambda: self.log_text.see(tk.END))
            self.log_text.config(state="disabled")
            
    
    def all_process_pdf(self, pdf_path):
        #エラーにより実行後の状態が残っている場合にリフレッシュ
        if self.log_window:
            self.log_window.destroy()
        self.reset_state()
        
        #projディレクトリ内にディレクトリがあれば削除(不具合を避けるため)
        proj_dir = "./proj/"
        for item in os.listdir(proj_dir):
            item_path = os.path.join(proj_dir, item)
            # ディレクトリであれば削除
            if os.path.isdir(item_path):
                shutil.rmtree(item_path)

        # サブウィンドウでログ表示
        self.log_window = tk.Toplevel(self)
        self.log_window.title("Console_Window")
        self.log_text = tk.Text(self.log_window, wrap="word", height=20, width=80)
        self.log_text.pack(padx=10, pady=10)
        self.log_text.config(state="disabled")
        
        self.append_log("PDFを画像に変換中...\n")
        img_path_l = Tools(pdf_path).pdf2img()
        self.append_log("PDFを画像に変換完了。\n")

        # 全ページ分の結果を格納する辞書
        self.project_data_all = {}

        for page_index, img_path in enumerate(img_path_l):
            self.append_log(f"ページ {page_index+1} のOCR処理を開始します...\n")
            page_data = self.process_pdf(img_path, page_index)  # page_index を渡す
            if page_data:
                self.project_data_all[page_index] = page_data

        self.append_log("全ページのOCR処理が完了しました。\n")

        # 全ページ分のタブを作成
        self.create_all_tabs()


    def process_pdf(self, img_path, page_index):
        """
        1ページ分のOCR処理を行い、結果を辞書で返す。
        共通 ./proj/img に生成されたナンバリング画像をページ専用 img_pageN に移動し、
        failed_img_pageN を作成して保存する。
        """
        try:
            self.append_log(f"OCR処理開始: {os.path.basename(img_path)}\n")

            # OCR実行クラスを初期化（output_dirは渡さない）
            convert_class = Convert_to_df(
                img_path,
                self.api_base_url,
                self.append_log
            )

            # OCR結果と黒塗り判定を取得（この呼び出しで ./proj/img/img{i}.png が生成される前提）
            text_arr, black_density_flag = convert_class.img2list()

            # 画像番号配列を作成
            N = text_arr.shape[0] * text_arr.shape[1]
            img_arr = np.arange(1, N + 1).reshape(text_arr.shape[0], text_arr.shape[1])

            # 認識失敗セルの抽出
            failed_img = img_arr[black_density_flag].tolist()
            failed_text = text_arr[black_density_flag].tolist()

            # ページファイル名（拡張子なし）
            img_filename = os.path.splitext(os.path.basename(img_path))[0]

            # ページ専用 img ディレクトリ作成
            page_img_dir = f".\\proj\\img_page{page_index+1}"
            os.makedirs(page_img_dir, exist_ok=True)

            # 共通 ./proj/img からページ専用へ移動
            src_img_dir = ".\\proj\\img"
            if os.path.isdir(src_img_dir):
                for name in os.listdir(src_img_dir):
                    if name.lower().startswith("img") and name.lower().endswith(".png"):
                        shutil.move(os.path.join(src_img_dir, name),
                                    os.path.join(page_img_dir, name))
                # 空になった共通imgは削除
                shutil.rmtree(src_img_dir, ignore_errors=True)

            # ページ専用 failed_img ディレクトリ作成
            page_failed_dir = f".\\proj\\failed_img_page{page_index+1}"
            os.makedirs(page_failed_dir, exist_ok=True)

            # failed_img をコピー
            for i in failed_img:
                src_path = os.path.join(page_img_dir, f"img{i}.png")
                if os.path.exists(src_path):
                    shutil.copy(src_path, page_failed_dir)

            # このページの結果を辞書でまとめる
            page_data = {
                "img_filename": img_filename,
                "text_arr": text_arr.tolist(),
                "img_arr": img_arr.tolist(),
                "failed_img": failed_img,
                "failed_text": failed_text,
                "convert_class": convert_class,
                "page_dir": page_failed_dir
            }
    

            self.append_log(f"OCR処理完了: {img_filename}\n")
            return page_data

        except Exception as e:
            self.append_log(f"エラー発生: {str(e)}\n")
            return None


    def create_all_tabs(self):
        """
        self.project_data_all に格納された全ページ分の結果をNotebookに展開する
        """
        self.notebook = ttk.Notebook(self.image_frame)
        self.notebook.pack(fill="both", expand=True)

        # タブIDとpage_dataの対応表
        self.page_tab_map = {}

        for page_index, page_data in self.project_data_all.items():
            frame = self.create_page_tab(page_index, page_data)
            tab_id = self.notebook.tabs()[-1]
            self.page_tab_map[tab_id] = page_data

        # タブ切り替えイベント登録
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)

        # 最初のタブのTreeviewを初期表示
        if self.notebook.tabs():
            first_tab = self.notebook.tabs()[0]
            first_page_data = self.page_tab_map[first_tab]
            self.build_treeview(self.tree, first_page_data["page_dir"])
            self.tree.bind(
                "<<TreeviewSelect>>",
                lambda e: self.on_tree_select_for_page(e, self.tree, first_page_data)
            )


    def create_page_tab(self, page_index, page_data):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text=f"Page {page_index+1}")

        canvas = tk.Canvas(frame, bg='white')
        canvas.pack(fill="both", expand=True)

        text_entry = tk.Entry(frame)
        text_entry.pack(fill="x", padx=5, pady=5)

        # UI要素をpage_dataに保持
        page_data["canvas"] = canvas
        page_data["text_entry"] = text_entry
        # 修正登録ボタン
        def register_text_for_page():
            if "current_image_path" not in page_data:
                messagebox.showwarning("警告", "画像が選択されていません。")
                return
            raw = text_entry.get().strip()
            if not raw:
                messagebox.showwarning("警告", "テキストが入力されていません。")
                return
            text_arr = page_data["text_arr"]
            v = len(text_arr)
            h = len(text_arr[0])
            img_num = int(os.path.splitext(os.path.basename(page_data["current_image_path"]))[0][3:])
            # 修正: img_num は1始まりなので -1 を使って行列インデックスを計算
            row = (img_num - 1) // h
            col = (img_num - 1) % h

            # 元セルの構造に合わせて代入（list なら split、str ならそのまま）
            try:
                orig = text_arr[row][col]
            except Exception as e:
                messagebox.showerror("エラー", f"内部インデックスエラー: {e}")
                return

            if isinstance(orig, list):
                new_val = raw.split()
            else:
                new_val = raw

            text_arr[row][col] = new_val
            self.append_log(f"Page {page_index+1} 修正登録: {new_val}\n")

        ttk.Button(frame, text="修正登録", command=register_text_for_page).pack(pady=5)

        return frame

        
    #pdfファイルを開く
    def open_pdf(self):
        pdf_path = filedialog.askopenfilename(
            title="pdfファイルを開く",
            filetypes=[("pdf file", "*.pdf")]
        )
        if not pdf_path.lower().endswith('.pdf'):
            messagebox.showwarning("警告", "PDFファイルを選択してください。")
            return

        # 並列処理開始
        threading.Thread(target=lambda: self.all_process_pdf(pdf_path), daemon=True).start()



    def is_special_spec_only(self, df):
        """
        DataFrame内に「仕様」列だけ非空で他列が空欄の行が存在すればTrue
        """
        if "仕様" not in df.columns:
            return False

        for _, row in df.iterrows():
            spec_val = str(row["仕様"]).strip()
            other_vals = [str(v).strip() for col, v in row.items() if col != "仕様"]
            if spec_val and all(v == "" for v in other_vals):
                return True
        return False


    def save_all_pages(self):
        os.makedirs("./out", exist_ok=True)

        # PDFファイル名（拡張子なし）を取得
        if hasattr(self, "pdf_path") and self.pdf_path:
            base_name = os.path.splitext(os.path.basename(self.pdf_path))[0]
        else:
            # 念のため、project_data_allの最初のページのimg_filenameを利用
            first_page = next(iter(self.project_data_all.values()))
            base_name = first_page["img_filename"]

        output_path = f"./out/{base_name}.xlsx"

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            for page_index, page_data in self.project_data_all.items():
                convert_class = page_data["convert_class"]
                text_arr = np.array(page_data["text_arr"])

                # 左右に分割してDataFrame化
                text_arr_left, text_arr_right = convert_class.split_half_arrays(text_arr)
                text_df_left, text_df_right = convert_class.arr2df(text_arr_left, text_arr_right)

                # 左右それぞれ「仕様列だけ埋まっている行があるか」を判定
                left_flag = self.is_special_spec_only(text_df_left)
                right_flag = self.is_special_spec_only(text_df_right)

                # 左側の処理
                if left_flag:
                    sep_name = f"./out/{base_name}_Page{page_index+1}_Left.xlsx"
                    with pd.ExcelWriter(sep_name, engine="openpyxl") as sep_writer:
                        convert_class.out_df(text_df_left).to_excel(sep_writer, sheet_name="Left", index=False)
                    self.append_log(f"Page {page_index+1} 左側を別ファイルに保存しました: {sep_name}\n")
                else:
                    sheet_left = f"Page{page_index+1}_Left"
                    convert_class.out_df(text_df_left).to_excel(writer, sheet_name=sheet_left, index=False)

                # 右側の処理
                if right_flag:
                    sep_name = f"./out/{base_name}_Page{page_index+1}_Right.xlsx"
                    with pd.ExcelWriter(sep_name, engine="openpyxl") as sep_writer:
                        convert_class.out_df(text_df_right).to_excel(sep_writer, sheet_name="Right", index=False)
                    self.append_log(f"Page {page_index+1} 右側を別ファイルに保存しました: {sep_name}\n")
                else:
                    sheet_right = f"Page{page_index+1}_Right"
                    convert_class.out_df(text_df_right).to_excel(writer, sheet_name=sheet_right, index=False)

        # ./proj 内のディレクトリを全削除
        proj_dir = "./proj"
        if os.path.isdir(proj_dir):
            for name in os.listdir(proj_dir):
                dir_path = os.path.join(proj_dir, name)
                if os.path.isdir(dir_path):
                    shutil.rmtree(dir_path, ignore_errors=True)

        # Notebookとタブ関連をリセット
        if hasattr(self, "notebook") and self.notebook.winfo_exists():
            self.notebook.destroy()
            self.notebook = None
        self.page_tab_map = {}
        if hasattr(self, "tree") and self.tree.winfo_exists():
            self.tree.delete(*self.tree.get_children())

        messagebox.showinfo("保存", f"全ページのExcel出力が完了しました。\n {output_path}")

        self.log_window.destroy()
        self.reset_state()


            
    def load_image_for_page(self, image_path, page_data, canvas, text_entry):
        try:
            frame_width = canvas.winfo_width()
            frame_height = canvas.winfo_height()

            if frame_width < 10 or frame_height < 10:
                self.after(100, lambda: self.load_image_for_page(image_path, page_data, canvas, text_entry))
                return

            img = Image.open(image_path)
            try:
                resample_filter = Image.Resampling.LANCZOS
            except AttributeError:
                resample_filter = Image.ANTIALIAS

            img = img.resize((frame_width, frame_height), resample_filter)
            tk_img = ImageTk.PhotoImage(img)
            canvas.delete("all")
            canvas.create_image(0, 0, anchor='nw', image=tk_img)
            canvas.image = tk_img

            # 現在の画像パスを保存
            page_data["current_image_path"] = image_path
            
            # エントリー更新
            text_entry.delete(0, tk.END)
            img_num = int(os.path.splitext(os.path.basename(image_path))[0][3:])
            text_entry.insert(
                tk.END,
                " ".join(np.array(page_data["text_arr"]).reshape(-1).tolist()[img_num - 1])
            )

        except Exception as e:
            messagebox.showerror("画像エラー", f"画像もしくは文字列を読み込めませんでした: {e}")
            canvas.delete("all")


         
    #プロジェクトを閉じる機能(現在はどこからも呼び出されない)
    def close_project(self):
        if not self.project_data:
            return  # 開いているプロジェクトがなければ何もしない

        if self.is_modified:
            result = messagebox.askyesnocancel(
                "保存の確認",
                "変更が保存されていません。保存しますか？"
            )
            if result is None:
                return  # キャンセルしたので終了しない
            elif result:
                self.save_json()
                if self.is_modified:  # 保存失敗
                    messagebox.showerror("エラー", "プロジェクトが正しく保存できませんでした")
                    return

        # ツリービューの中身をクリアして初期状態に戻す
        for item in self.tree.get_children():
            self.tree.delete(item)

        # スクロールバー設定を再設定（通常は初期のままでOK）
        self.tree.config(yscrollcommand=self.tree_scrollbar.set)
        self.tree_scrollbar.config(command=self.tree.yview)

        # 選択状態クリア
        self.tree.selection_remove(self.tree.selection())

        # 画像キャンバスをクリア
        self.image_canvas.delete("all")

        # テキストエントリーもクリア
        self.text_entry.delete(0, tk.END)
        
        #image_frame内にウィジットが存在すれば全削除
        for widget in self.image_frame.winfo_children():
            widget.destroy()

        # 変数を全て初期化
        self.__init__()

    
    #閉じる前に保存するか確認(プロジェクトファイル(tmp.json)に保存する形式を撤廃したため現在はどこからも呼び出されない)
    def on_close(self):
        if self.is_modified:
            result = messagebox.askyesnocancel("保存の確認", "変更内容が保存されていません。保存しますか？")
            if result is None:
                return  # キャンセル
            elif result:
                #tmp.jsonが存在する場合
                if self.project_path:
                    self.save_json()
                #tmp.jsonが存在しない場合
                else:
                    messagebox.showerror("エラー", "tmp.jsonが存在しません。PDFファイルを開いて生成してください。")
        self.destroy()

if __name__ == '__main__':
    app = PDF2xlsxApp()
    app.mainloop()