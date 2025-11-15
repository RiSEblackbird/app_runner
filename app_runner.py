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
                              QTreeWidgetItem, QHeaderView, QTableWidget,
                              QTableWidgetItem, QMessageBox)
from PySide6.QtCore import Qt, QThread, QTimer
import ctypes

# 定数定義
TARGET_WINDOW_TITLE = "GUIツールランナー.exe"
DESTINATION_WINDOW_TITLE = "Pythonファイル実行ツール"
INNER_PADDING = 50

# 定数としてウィンドウ位置情報を保存するファイル名を設定
POSITION_FILE = 'window_position_app_runner.csv'

# アプリ情報を保存するCSVファイルパスを定数として定義
APPS_CSV_FILE = os.path.join(os.path.dirname(__file__), "app_runner_apps.csv")

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
        # 対象ウィンドウのハンドル取得（タイトル一致が無ければコンソールにフォールバック）
        hwnd_target = win32gui.FindWindow(None, target_title)
        if not hwnd_target:
            hwnd_target = ctypes.windll.kernel32.GetConsoleWindow()

        if hwnd_target:
            try:
                win32gui.SetWindowPos(hwnd_target, None, new_x, new_y, 0, 0, win32con.SWP_NOZORDER | win32con.SWP_NOSIZE)
            except Exception as e:
                print(f"SetWindowPos 失敗: {e}")
        else:
            print(f"対象ウィンドウが見つかりませんでした: '{target_title}'。処理をスキップします。")


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


def find_app_python_files(base_folder):
    """
    DEFAULT_FOLDER_PATH直下と、その直下のサブフォルダ（1階層目）のPythonアプリファイルを探索し、
    (priority, app_name, filename, full_path) のリストを返す。

    「自動追記」機能で未登録アプリを検出するために使用する。
    """
    files_info = []
    try:
        for entry in os.listdir(base_folder):
            path = os.path.join(base_folder, entry)

            # 直下のファイル
            if os.path.isfile(path):
                if entry.endswith(PYTHON_FILE_EXTENSION) and entry.startswith(APP_TITLE_PARTS):
                    tmp_name = App.extract_raw_app_name_static(path)
                    priority = App.extract_priority_static(tmp_name)
                    app_name = App.extract_app_name_static(tmp_name)
                    files_info.append((priority, app_name, entry, path))

            # 直下のサブフォルダ（1階層目）
            elif os.path.isdir(path):
                try:
                    for filename in os.listdir(path):
                        child = os.path.join(path, filename)
                        if os.path.isfile(child) and filename.endswith(PYTHON_FILE_EXTENSION) and filename.startswith(APP_TITLE_PARTS):
                            tmp_name = App.extract_raw_app_name_static(child)
                            priority = App.extract_priority_static(tmp_name)
                            app_name = App.extract_app_name_static(tmp_name)
                            files_info.append((priority, app_name, filename, child))
                except Exception:
                    # 個々のサブフォルダのエラーは無視し、全体の処理は継続する
                    continue
    except Exception:
        except_processing()

    return files_info


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
        QVBoxLayout(main_widget)

        # サブウィンドウ（CSV編集ウィンドウ）への参照を保持する
        self.csv_editor_window = None

        self.create_widgets()
        self.folder_path_entry.setText(DEFAULT_FOLDER_PATH)
        self.update_file_list()
        
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
        self.edit_csv_button = QPushButton("アプリCSV編集")
        
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
        run_exit_layout.addWidget(self.edit_csv_button)
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
        self.edit_csv_button.clicked.connect(self.open_csv_editor)
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

    def open_csv_editor(self):
        """
        アプリ情報CSVを編集する専用ウィンドウを開く。
        既に開いている場合は、そのウィンドウを前面に出す。
        """
        if self.csv_editor_window is None:
            self.csv_editor_window = CsvEditorWindow(self)
        self.csv_editor_window.show()
        self.csv_editor_window.raise_()
        self.csv_editor_window.activateWindow()

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
            # 現在はフォルダパスは「フォルダを開く」「VSCodeで開く」で利用し、
            # アプリ一覧自体はCSVから読み込む。
            self.folder_path_entry.setText(folder)  # 入力フォームに選択したフォルダのパスを挿入

    
    def update_file_list(self, folder=None):
        """
        アプリ情報CSVから一覧を読み込み、ツリー形式で表示する。
        画面でのアプリ一覧は常にCSVの内容を反映する。
        """
        self.file_tree.clear()

        try:
            with open(APPS_CSV_FILE, newline='', encoding=ENCODING_FOR_CSV) as csvfile:
                reader = csv.reader(csvfile)
                # 先頭行はヘッダーとして扱う
                _ = next(reader, None)
                for row in reader:
                    if len(row) < 4:
                        # 想定外の行はスキップする
                        continue
                    priority, app_name, filename, full_path = row[:4]
                    item = QTreeWidgetItem([str(priority), app_name, filename, full_path])
                    self.file_tree.addTopLevelItem(item)
        except FileNotFoundError:
            messagebox.showerror("エラー", f"アプリCSVファイルが見つかりません: {APPS_CSV_FILE}")
        except Exception:
            except_processing()

    @staticmethod
    def extract_raw_app_name_static(file_path):
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
        except Exception:
            return "読み込み失敗"

    @staticmethod
    def extract_app_name_static(raw_name):
        """
        優先度を含むアプリ名から優先度部分を除去する。
        """
        if ". " in raw_name:
            return raw_name.split(". ")[1]
        return raw_name

    @staticmethod
    def extract_priority_static(raw_name):
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


class CsvEditorWindow(QMainWindow):
    """
    アプリ情報CSV（app_runner_apps.csv）を編集するためのウィンドウ。
    任意編集に加え、「自動追記」で未登録アプリを探索し追記できる。
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("アプリCSV編集")
        self.resize(900, 500)

        # 親のAppインスタンス（更新通知用）
        self._app = parent

        central_widget = QWidget()
        layout = QVBoxLayout()

        # CSVの内容を表示・編集するテーブル
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(['優先度', 'アプリ名', 'pyファイル名', '絶対パス'])

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Interactive)
        header.setSectionResizeMode(2, QHeaderView.Interactive)
        header.setSectionResizeMode(3, QHeaderView.Stretch)

        # 操作ボタン群
        button_row = QWidget()
        button_layout = QHBoxLayout(button_row)
        self.reload_button = QPushButton("再読み込み")
        self.auto_append_button = QPushButton("自動追記")
        self.save_button = QPushButton("保存")
        self.close_button = QPushButton("閉じる")

        button_layout.addWidget(self.reload_button)
        button_layout.addWidget(self.auto_append_button)
        button_layout.addWidget(self.save_button)
        button_layout.addWidget(self.close_button)

        layout.addWidget(self.table)
        layout.addWidget(button_row)
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        # シグナル接続
        self.reload_button.clicked.connect(self.load_csv)
        self.auto_append_button.clicked.connect(self.auto_append_missing_apps)
        self.save_button.clicked.connect(self.save_csv)
        self.close_button.clicked.connect(self.close)

        # 初期表示
        self.load_csv()

    def load_csv(self):
        """
        CSVファイルからデータを読み込み、テーブルに表示する。
        """
        self.table.setRowCount(0)
        try:
            with open(APPS_CSV_FILE, newline='', encoding=ENCODING_FOR_CSV) as csvfile:
                reader = csv.reader(csvfile)
                # 先頭行はヘッダー行として読み飛ばす
                _ = next(reader, None)

                for row in reader:
                    # 空行や列数不足の行はスキップ
                    if not row or len(row) < 4:
                        continue
                    row_index = self.table.rowCount()
                    self.table.insertRow(row_index)
                    for col_index in range(4):
                        value = row[col_index] if col_index < len(row) else ""
                        item = QTableWidgetItem(value)
                        self.table.setItem(row_index, col_index, item)
        except FileNotFoundError:
            # ファイルが無ければ空の状態から編集開始できるようにする
            pass
        except Exception:
            except_processing()

    def save_csv(self):
        """
        テーブルの内容をCSVファイルに保存する。
        空行はスキップし、ヘッダー行を先頭に出力する。
        """
        try:
            with open(APPS_CSV_FILE, 'w', newline='', encoding=ENCODING_FOR_CSV) as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(['優先度', 'アプリ名', 'pyファイル名', '絶対パス'])

                for row in range(self.table.rowCount()):
                    row_values = []
                    for col in range(4):
                        item = self.table.item(row, col)
                        row_values.append(item.text().strip() if item else "")

                    # 全列が空の場合はスキップ
                    if not any(row_values):
                        continue

                    writer.writerow(row_values)

            QMessageBox.information(self, "保存完了", "アプリCSVを保存しました。")

            # メイン画面の一覧も更新する
            if self._app is not None:
                self._app.update_file_list()
        except Exception:
            except_processing()

    def auto_append_missing_apps(self):
        """
        DEFAULT_FOLDER_PATH 配下からアプリ用Pythonファイルを探索し、
        まだCSVに登録されていないものをテーブルの末尾に追記する。
        """
        # 既存の絶対パス一覧を取得
        existing_paths = set()
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 3)
            if item:
                existing_paths.add(item.text().strip())

        app_files = find_app_python_files(DEFAULT_FOLDER_PATH)

        added_count = 0
        for priority, app_name, filename, full_path in app_files:
            if full_path in existing_paths:
                continue

            row_index = self.table.rowCount()
            self.table.insertRow(row_index)
            self.table.setItem(row_index, 0, QTableWidgetItem(str(priority)))
            self.table.setItem(row_index, 1, QTableWidgetItem(app_name))
            self.table.setItem(row_index, 2, QTableWidgetItem(filename))
            self.table.setItem(row_index, 3, QTableWidgetItem(full_path))
            added_count += 1

        if added_count == 0:
            QMessageBox.information(self, "自動追記", "新しく追記するアプリはありませんでした。")
        else:
            QMessageBox.information(self, "自動追記", f"{added_count}件のアプリを追記しました。")
    
    
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
