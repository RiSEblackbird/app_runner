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
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                              QHBoxLayout, QLineEdit, QPushButton, QTreeWidget, 
                              QTreeWidgetItem, QHeaderView)
from PySide6.QtCore import Qt, QThread, QTimer

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
    # QtのgeometryをTkinter形式の文字列に変換
    geom = root.geometry()
    geometry_str = f"{geom.width()}x{geom.height()}+{geom.x()}+{geom.y()}"
    position_data = [HOSTNAME, geometry_str]
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
                    # tkinterのgeometry文字列（例: "100x100+200+200"）をQtの形式に変換
                    geom = row[1].replace('x', '+').split('+')
                    if len(geom) == 4:
                        width, height, x, y = map(int, geom)
                        root.setGeometry(x, y, width, height)
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
        # シンプルに実行するだけ（コンソール制御なし）
        result = subprocess.run(
            ["python", script_path], 
            check=True, 
            text=True
        )
        print(f"'{script_path}' executed successfully.")
    except subprocess.CalledProcessError as e:
        if attempt <= max_attempts:
            print(f"Error executing '{script_path}': {e}. Attempt {attempt}/{max_attempts}. Retrying...")
            execute_script(script_path, attempt + 1)
        else:
            messagebox.showerror("エラー", f"'{script_path}' failed after {max_attempts} attempts.")
    except Exception as e:
        messagebox.showerror("エラー", f"Unexpected error executing '{script_path}': {e}")


# アプリケーションクラスの定義
class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(DESTINATION_WINDOW_TITLE)
        
        # メインウィジェットとレイアウトの設定
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        
        self.create_widgets()
        self.folder_path_entry.setText(DEFAULT_FOLDER_PATH)
        self.update_file_list(DEFAULT_FOLDER_PATH)
        
        # ウィンドウ位置の調整を遅延実行
        QTimer.singleShot(100, self.adjust_window_position)

    def adjust_window_position(self):
        """
        アプリケーションウィンドウとコンソールウィンドウの位置を調整
        """
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
        # フォルダ操作エリア（1行目）
        folder_frame = QWidget()
        folder_layout = QHBoxLayout(folder_frame)
        
        self.folder_path_entry = QLineEdit()
        self.browse_button = QPushButton("選択")
        self.open_folder_button = QPushButton("フォルダを開く")
        self.open_vscode_button = QPushButton("VSCodeで開く")
        
        folder_layout.addWidget(self.folder_path_entry)
        folder_layout.addWidget(self.browse_button)
        folder_layout.addWidget(self.open_folder_button)
        folder_layout.addWidget(self.open_vscode_button)
        
        # 実行と終了ボタンエリア（2行目）
        run_exit_frame = QWidget()
        run_exit_layout = QHBoxLayout(run_exit_frame)
        
        self.run_button = QPushButton("RUN")
        self.run_main_button = QPushButton("RUN MAIN")
        self.exit_button = QPushButton("アプリ終了")
        
        run_exit_layout.addWidget(self.run_button)
        run_exit_layout.addWidget(self.run_main_button)
        run_exit_layout.addWidget(self.exit_button)
        
        # ファイル一覧エリア（3行目以降）
        self.file_tree = QTreeWidget()
        self.file_tree.setHeaderLabels(['優先度', 'アプリ名', 'pyファイル名', '絶対パス'])
        self.file_tree.setSelectionMode(QTreeWidget.ExtendedSelection)
        
        # 列幅の設定
        header = self.file_tree.header()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)  # 優先度
        header.setSectionResizeMode(1, QHeaderView.Interactive)       # アプリ名
        header.setSectionResizeMode(2, QHeaderView.Interactive)       # pyファイル名
        header.setSectionResizeMode(3, QHeaderView.Stretch)          # 絶対パス
        
        # pyファイル名の列の初期幅を設定（例：200ピクセル）
        self.file_tree.setColumnWidth(2, 200)
        
        # シグナル/スロットの接続
        self.browse_button.clicked.connect(self.browse_folder)
        self.open_folder_button.clicked.connect(self.open_folder)
        self.open_vscode_button.clicked.connect(self.open_vscode)
        self.run_button.clicked.connect(self.run_python_files)
        self.run_main_button.clicked.connect(self.run_main_files)
        self.exit_button.clicked.connect(self.close)
        
        # レイアウトに追加
        layout = QVBoxLayout()
        layout.addWidget(folder_frame)
        layout.addWidget(run_exit_frame)
        layout.addWidget(self.file_tree)
        
        # メインウィジェットの設定
        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

    def open_folder(self):
        folder_path = self.folder_path_entry.text()  # フォルダパスを取得
        if sys.platform == "win32":
            os.startfile(folder_path)
        else:
            subprocess.Popen(["open", folder_path])


    def open_vscode(self):
        folder_path = self.folder_path_entry.text()
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
            self.folder_path_entry.setText(folder)  # 入力フォームに選択したフォルダのパスを挿入
            self.update_file_list(folder)  # ファイルリストの更新

    
    def update_file_list(self, folder):
        """
        指定されたフォルダ内のPythonファイルをツリー形式で表示する
        """
        self.file_tree.clear()
        
        files_info = []
        for filename in os.listdir(folder):
            if filename.endswith(PYTHON_FILE_EXTENSION) and filename.startswith(APP_TITLE_PARTS):
                full_path = os.path.join(folder, filename)
                tmp_name = self.extract_raw_app_name(full_path)
                priority = self.extract_priority(tmp_name)
                app_name = self.extract_app_name(tmp_name)
                files_info.append((priority, app_name, filename, full_path))
        
        sorted_files = sorted(files_info, key=lambda x: x[0])
        
        for priority, app_name, filename, full_path in sorted_files:
            item = QTreeWidgetItem([str(priority), app_name, filename, full_path])
            self.file_tree.addTopLevelItem(item)

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
        selected_items = self.file_tree.selectedItems()
        if selected_items:
            for item in selected_items:
                file_path = item.text(3)  # 絶対パスは4列目
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
        # すべてのトップレベルアイテムを取得
        for i in range(self.file_tree.topLevelItemCount()):
            item = self.file_tree.topLevelItem(i)
            if item.text(0) == "0":  # 優先度が0のアイテムを選択
                main_items.append(item)
        
        if main_items:
            for item in main_items:
                file_path = item.text(3)  # 絶対パスは4列目
                try:
                    threading.Thread(target=execute_script, args=(file_path,), daemon=True).start()
                    time.sleep(0.1)
                except Exception:
                    except_processing()


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
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = None  # windowをNoneで初期化
    try:
        window = App()
        restore_position(window)
        window.show()
        app.exec()
    except Exception:
        if window is not None:  # windowが作成されている場合のみ保存
            save_position(window)
        except_processing()
