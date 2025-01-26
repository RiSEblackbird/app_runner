# -*- coding: utf-8 -*-
# アプリ名: Pythonファイル実行ツール

import os
import sys
import tkinter as tk
import traceback
from datetime import datetime, timedelta
from tkinter import filedialog
from tkinter import messagebox
from tkinter import font as tkfont
import subprocess
import threading
import yaml
import win32gui
import win32con
import win32api
import time
import socket
import csv
from tkinter import ttk

# 定数定義
TARGET_WINDOW_TITLE = "GUIツールランナー.exe"
DESTINATION_WINDOW_TITLE = "Pythonファイル実行ツール"
INNER_PADDING = 50

# 定数としてウィンドウ位置情報を保存するファイル名を設定
POSITION_FILE = 'window_position_app_runner.csv'

# 定数としてホスト名を取得
HOSTNAME = socket.gethostname()

# ウインドウを指定のウインドウの内側に移動する関数
def move_window_inside_relative(target_title, destination_title, padding):
    print("処理を開始します。")
    # 移動先のウインドウの位置とサイズを取得
    hwnd_dest = win32gui.FindWindow(None, destination_title)
    if hwnd_dest:
        rect = win32gui.GetWindowRect(hwnd_dest)
        x_dest, y_dest = rect[0], rect[1]

        print(f"'{destination_title}'のウインドウの位置: ({x_dest}, {y_dest})")

        # 初期位置を設定（エラー回避のため）
        new_x, new_y = x_dest, y_dest

        # モニターの全体的なサイズを取得
        monitors_info = win32api.EnumDisplayMonitors()
        for monitor in monitors_info:
            monitor_info = win32api.GetMonitorInfo(monitor[0])
            monitor_rect = monitor_info['Monitor']
            # ウインドウがこのモニター内にあるか判断
            if monitor_rect[0] <= x_dest < monitor_rect[2] and monitor_rect[1] <= y_dest < monitor_rect[3]:
                mx_center = (monitor_rect[0] + monitor_rect[2]) / 2
                my_center = (monitor_rect[1] + monitor_rect[3]) / 2
                # ウインドウの新しい位置を計算
                new_x = x_dest - padding if x_dest > mx_center else x_dest + padding
                new_y = y_dest - padding if y_dest > my_center else y_dest + padding
                break  # 適切なモニターが見つかったらループを抜ける

        print(f"移動後の '{target_title}' のウインドウの新しい位置: ({new_x}, {new_y})")
        win32gui.SetWindowPos(win32gui.FindWindow(None, target_title), None, new_x, new_y, 0, 0, win32con.SWP_NOZORDER | win32con.SWP_NOSIZE)


def save_position(root):
    """
    ウィンドウの位置とサイズをCSVファイルに保存する。
    """
    print("ウィンドウ位置を保存中...")
    position_data = [HOSTNAME, root.geometry()]
    print(f"保存データ: {position_data}")
    with open(POSITION_FILE, 'w', newline='', encoding="utf_8_sig") as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(position_data)
    print("保存完了")

def restore_position(root):
    """
    CSVファイルからウィンドウの位置とサイズを復元する。
    """
    print("ウィンドウ位置を復元中...")
    try:
        with open(POSITION_FILE, newline='', encoding="utf_8_sig") as csvfile:
            reader = csv.reader(csvfile)
            for row in reader:
                if row[0] == HOSTNAME:
                    print(f"復元データ: {row[1]}")
                    root.geometry(row[1])
                    break
    except FileNotFoundError:
        print("位置情報ファイルが見つかりません。")

# アプリケーションの終了時の処理をカスタマイズする
def on_close():
    save_position(app)  # ウィンドウの位置を保存
    app.destroy()  # ウィンドウを破壊する

# YAMLファイルを読み込むための関数
def load_yaml_settings(file_path):
    """
    指定されたパスのYAMLファイルを読み込んで、設定を辞書として返す。
    """
    with open(file_path, 'r', encoding="utf-8") as file:
        # yamlモジュールを使用して設定ファイルを読み込む
        settings = yaml.safe_load(file)
    return settings


# パス処理関数の定義
def path_treatment(arg):
    """
    パスが "/" で繋がれていたり、先頭末尾に '"' があっても読み取れる文字列に加工する
    """
    return r'{}'.format(arg).strip().strip('"')


# 例外処理関数の定義
def except_processing():
    """
    例外のトレースバックを取得し、メッセージボックスで表示する
    """
    t, v, tb = sys.exc_info()
    trace = traceback.format_exception(t, v, tb)
    messagebox.showerror("エラーが発生しました", ''.join(trace))


# スクリプトを非同期で実行し、必要に応じて再起動する関数
def execute_script(script_path, attempt=1):
    """
    新しいスレッドでPythonスクリプトを実行し、エラーが発生した場合は再起動する。
    """
    max_attempts = 3  # 最大再起動回数
    try:
        result = subprocess.run(["python", script_path], check=True, text=True)
        print(f"'{script_path}' executed successfully.")  # 実行成功をログに記録
    except subprocess.CalledProcessError as e:
        if attempt <= max_attempts:
            print(f"Error executing '{script_path}': {e}. Attempt {attempt}/{max_attempts}. Retrying...")
            execute_script(script_path, attempt + 1)  # 再試行
        else:
            messagebox.showerror("エラー", f"'{script_path}' failed after {max_attempts} attempts.")
    except Exception as e:
        messagebox.showerror("エラー", f"Unexpected error executing '{script_path}': {e}")


# アプリケーションクラスの定義
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(DESTINATION_WINDOW_TITLE)  # ウィンドウタイトルの設定
        self.create_widgets()  # ウィジェットの生成と配置
        self.folder_path_entry.insert(0, DEFAULT_FOLDER_PATH)  # デフォルトパスの設定
        self.update_file_list(DEFAULT_FOLDER_PATH)  # デフォルトパスに基づいてファイルリストを更新
        self.after(100, self.adjust_window_position)  # ウィンドウの位置調整をスケジュール
        
    def adjust_window_position(self):
        # アプリのウインドウが作成された後にウインドウ位置を調整
        move_window_inside_relative(TARGET_WINDOW_TITLE, DESTINATION_WINDOW_TITLE, INNER_PADDING)
        # DESTINATION_WINDOW_TITLE のウィンドウを最前面に表示
        hwnd = win32gui.FindWindow(None, DESTINATION_WINDOW_TITLE)
        if hwnd:
            win32gui.SetForegroundWindow(hwnd)

    def create_widgets(self):
        """
        ウィジェットを作成し、ウィンドウに配置する
        """
        # 使用するフォントの設定
        default_font = tkfont.Font(size=FONT_SIZE)

        # Treeviewのスタイル設定
        style = ttk.Style()
        style.configure("Treeview", font=('', FONT_SIZE))  # フォントサイズを設定
        style.configure("Treeview.Heading", font=('', FONT_SIZE))  # ヘッダーのフォントサイズを設定

        # フォルダ操作エリア（1行目）
        folder_frame = tk.Frame(self)
        folder_frame.pack(padx=10, pady=5)

        self.folder_path_entry = tk.Entry(folder_frame, width=20, font=default_font)  # 入力フォームの作成
        self.folder_path_entry.grid(row=0, column=0, sticky="ew", padx=(0, 10))

        self.browse_button = tk.Button(folder_frame, text="選択", command=self.browse_folder, font=default_font)  # 選択ボタンの作成
        self.browse_button.grid(row=0, column=1)

        self.open_folder_button = tk.Button(folder_frame, text="フォルダを開く", command=self.open_folder, font=default_font)  # フォルダを開くボタンの作成
        self.open_folder_button.grid(row=0, column=2)

        self.open_vscode_button = tk.Button(folder_frame, text="VSCodeで開く", command=self.open_vscode, font=default_font)  # VSCodeで開くボタンの作成
        self.open_vscode_button.grid(row=0, column=3)

        # 実行と終了ボタンエリア（2行目）
        run_exit_frame = tk.Frame(self)
        run_exit_frame.pack(padx=10, pady=5)

        self.run_button = tk.Button(run_exit_frame, text="RUN", command=self.run_python_files, font=default_font)
        self.run_button.pack(side=tk.LEFT, padx=5)

        self.run_main_button = tk.Button(run_exit_frame, text="RUN MAIN", command=self.run_main_files, font=default_font)
        self.run_main_button.pack(side=tk.LEFT, padx=5)

        self.exit_button = tk.Button(run_exit_frame, text="アプリ終了", command=self.destroy, font=default_font)
        self.exit_button.pack(side=tk.LEFT, padx=5)

        # ファイル一覧エリア（3行目以降）
        tree_frame = tk.Frame(self)
        tree_frame.pack(padx=10, pady=10, fill='both', expand=True)
        
        # Gridの設定
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        self.file_tree = ttk.Treeview(tree_frame, columns=('priority', 'app_name', 'filename', 'path'), show='headings', height=10, selectmode='extended')
        
        # マウスドラッグによる選択を有効にする
        self.file_tree.bind('<B1-Motion>', self.on_drag)
        self.file_tree.bind('<Button-1>', self.on_click)
        
        # 列の設定
        self.file_tree.heading('priority', text='優先度')
        self.file_tree.heading('app_name', text='アプリ名')
        self.file_tree.heading('filename', text='pyファイル名')
        self.file_tree.heading('path', text='絶対パス')
        
        # 列の幅設定
        self.file_tree.column('priority', width=50, anchor='center')
        self.file_tree.column('app_name', width=200)
        self.file_tree.column('filename', width=150)
        self.file_tree.column('path', width=300)
        
        # スクロールバーの追加
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.file_tree.yview)
        self.file_tree.configure(yscrollcommand=scrollbar.set)
        
        # TreeviewとスクロールバーをGridで配置
        self.file_tree.grid(row=0, column=0, sticky='nsew')
        scrollbar.grid(row=0, column=1, sticky='ns')

    def open_folder(self):
        folder_path = self.folder_path_entry.get()  # フォルダパスを取得
        if sys.platform == "win32":
            os.startfile(folder_path)
        else:
            subprocess.Popen(["open", folder_path])


    def open_vscode(self):
        folder_path = self.folder_path_entry.get()
        # パスを正規化
        folder_path = os.path.normpath(folder_path)
        print([CODE_EXE_PATH, folder_path])
        subprocess.Popen([CODE_EXE_PATH, folder_path])


    def browse_folder(self):
        """
        フォルダ選択ダイアログを開き、選択したフォルダのPythonファイルを一覧表示する
        """
        folder = filedialog.askdirectory()  # フォルダ選択ダイアログを表示
        if folder:  # フォルダが選択された場合
            self.folder_path_entry.delete(0, tk.END)  # 入力フォームの内容をクリア
            self.folder_path_entry.insert(0, folder)  # 入力フォームに選択したフォルダのパスを挿入
            self.update_file_list(folder)  # ファイルリストの更新

    
    def update_file_list(self, folder):
        """
        指定されたフォルダ内のPythonファイルをテーブル形式で表示する
        """
        # 既存の項目をクリア
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)

        # フォルダ内のPythonファイルとその情報をリストに追加
        files_info = []
        for filename in os.listdir(folder):
            if filename.endswith(PYTHON_FILE_EXTENSION) and filename.startswith(APP_TITLE_PARTS):
                full_path = os.path.join(folder, filename)
                tmp_name = self.extract_raw_app_name(full_path)  # 優先度を含むアプリ名を取得
                priority = self.extract_priority(tmp_name)  # 優先度を抽出
                app_name = self.extract_app_name(tmp_name)  # アプリ名から優先度部分を除去
                files_info.append((priority, app_name, filename, full_path))

        # 優先度でソート
        sorted_files = sorted(files_info, key=lambda x: x[0])

        # ソートされたファイルをTreeviewに追加
        for priority, app_name, filename, full_path in sorted_files:
            self.file_tree.insert('', 'end', values=(priority, app_name, filename, full_path))

    def extract_raw_app_name(self, file_path):
        """
        指定されたPythonファイルから優先度を含むアプリ名を抽出する。
        """
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                lines = [next(file) for _ in range(5)]  # ファイルから最初の5行を読み込む
            for line in lines:
                if line.startswith("# アプリ名:"):  # アプリ名の行を探す
                    return line.strip().split(": ", 1)[1]
            return "名前未設定"  # アプリ名が見つからなかった場合
        except Exception as e:
            return "読み込み失敗"

    def extract_app_name(self, raw_name):
        """
        優先度を含むアプリ名から優先度部分を除去する。
        """
        if ". " in raw_name:
            return raw_name.split(". ")[1]
        return raw_name

    def extract_priority(self, raw_name):
        """
        アプリ名から優先度を抽出する。
        例: "0. ザ・アプリ" -> 0
        """
        try:
            if raw_name and raw_name[0].isdigit() and ". " in raw_name:
                priority_str = raw_name.split(". ")[0]
                return int(priority_str)
            return 999  # 優先度が見つからない場合
        except (ValueError, IndexError):
            return 999  # 優先度の解析に失敗した場合
    
    
    # 各スクリプトを別々のスレッドで実行する
    def run_python_files(self):
        """
        選択されたPythonファイルを新しいスレッドで一括実行する
        """
        selected_items = self.file_tree.selection()
        if selected_items:
            for item in selected_items:
                # 選択された項目から絶対パスを取得
                values = self.file_tree.item(item)['values']
                file_path = values[3]  # 絶対パスは4番目の列
                try:
                    threading.Thread(target=execute_script, args=(file_path,), daemon=True).start()
                    time.sleep(0.1)
                except Exception:
                    except_processing()

    def run_main_files(self):
        """
        優先度0のPythonファイルを一括実行する
        """
        main_items = []
        for item in self.file_tree.get_children():
            values = self.file_tree.item(item)['values']
            if values[0] == 0:  # 優先度が0のアイテムを選択
                main_items.append(item)
        
        if main_items:
            for item in main_items:
                values = self.file_tree.item(item)['values']
                file_path = values[3]  # 絶対パスは4番目の列
                try:
                    threading.Thread(target=execute_script, args=(file_path,), daemon=True).start()
                    time.sleep(0.1)
                except Exception:
                    except_processing()

    def on_click(self, event):
        """クリック時の処理"""
        self.start_item = self.file_tree.identify_row(event.y)
        if not self.start_item:
            return
        
        # Ctrlキーが押されていない場合は既存の選択をクリア
        if not event.state & 0x4:  # Ctrlキーの状態をチェック
            self.file_tree.selection_set(self.start_item)

    def on_drag(self, event):
        """ドラッグ時の処理"""
        drag_item = self.file_tree.identify_row(event.y)
        if not drag_item:
            return
        
        # 全アイテムを取得
        all_items = self.file_tree.get_children()
        start_idx = all_items.index(self.start_item)
        current_idx = all_items.index(drag_item)
        
        # 選択範囲を決定
        if start_idx <= current_idx:
            items_to_select = all_items[start_idx:current_idx + 1]
        else:
            items_to_select = all_items[current_idx:start_idx + 1]
        
        # 選択を更新
        self.file_tree.selection_set(items_to_select)


# 定数定義
ENCODING_FOR_CSV = "utf_8_sig"
PYTHON_FILE_EXTENSION = ".py"  # Pythonファイルの拡張子
FONT_SIZE = 12  # 基本フォントサイズ
APP_TITLE_PARTS = "app_"

# YAML設定ファイルパス
yaml_settings_path = 'desktop_gui_settings.yaml'
settings = load_yaml_settings(yaml_settings_path)
DEFAULT_FOLDER_PATH = settings["app_runner"]["DEFAULT_FOLDER_PATH"]
CODE_EXE_PATH = settings["app_runner"]["CODE_EXE_PATH"]

# メイン処理
try:
    app = App()  # アプリケーションのインスタンスを作成
    restore_position(app)
    app.protocol("WM_DELETE_WINDOW", on_close)  # 終了時処理の設定

    app.mainloop()  # アプリケーションを実行

except Exception:
    save_position(app)
    except_processing()
