# Pythonファイル実行ツール

![image](https://github.com/user-attachments/assets/3f013b77-eef0-4865-a904-f136d9f82c3f)

## 概要
このアプリケーションは、指定されたフォルダ内のPythonファイルを一覧表示し、選択したファイルを一括実行するためのツールです。ファイルの管理や実行を効率的に行うことができます。

## 主な機能
- フォルダ内のPythonファイル一覧表示（優先度順）
- ファイルの複数選択と一括実行（ドラッグ選択可能）
- 優先度0のファイルの一括実行（RUN MAIN機能）
- フォルダのブラウズ機能
- フォルダを直接開く機能
- VSCodeでフォルダを開く機能
- ウィンドウ位置の記憶と復元

## 技術仕様
- 開発言語: Python
- GUIフレームワーク: PySide6 (Qt)
- 設定ファイル: YAML
- ウィンドウ位置保存: CSV

## 使用方法
1. アプリケーションの起動
   - 以下のような内容でショートカットファイルを作成します：
     ```
     %windir%\System32\cmd.exe "/C" C:\ProgramData\Anaconda3\Scripts\activate.bat base && D: && cd D:\Tools\PythonRunner && python D:\Tools\PythonRunner\app_runner.py
     ```
   - このショートカットファイルをダブルクリックして起動します
   - ショートカットの内容説明：
     - Anaconda環境をアクティベート
     - 作業ドライブに移動
     - 作業ディレクトリに移動
     - Pythonスクリプトを実行

2. フォルダの選択
   - 「選択」ボタンでフォルダを選択する
   - パスを直接入力することも可能
   - 「フォルダを開く」でエクスプローラーで開く
   - 「VSCodeで開く」でエディタで開く

3. ファイルの選択と実行
   - リストから実行したいファイルを選択する
   - 複数のファイルを選択可能
     - Ctrl+クリックで個別選択
     - Shift+クリックで範囲選択
     - マウスドラッグで範囲選択
   - 「RUN」ボタンで選択したファイルを実行する
   - 「RUN MAIN」ボタンで優先度0のファイルを一括実行する

4. アプリケーションの終了
   - 「アプリ終了」ボタンで終了
   - ウィンドウ位置は自動的に保存される

## 機能詳細
### ファイル管理機能
- ファイル一覧表示
  - フォルダ内のPythonファイルを表形式で表示
  - 優先度、アプリ名、ファイル名、パスを表示
  - 優先度順にソートして表示
  - スクロールバーによる快適な閲覧

- フォルダ操作
  - フォルダパスの直接入力
  - フォルダ選択ダイアログ
  - エクスプローラーでの表示
  - VSCodeでの表示

### 実行機能
- 非同期実行
  - 選択したファイルを別スレッドで実行
  - エラー時の自動再試行（最大3回）
  - 実行状態のログ出力

- 一括実行
  - 複数ファイルの同時実行
  - 優先度0のファイルの一括実行（RUN MAIN）
  - 実行間隔の自動調整
  - エラーハンドリング

### ウィンドウ管理機能
- 位置記憶
  - 終了時の位置をCSVに保存
  - 起動時に前回位置を復元
  - ホスト名ごとの位置管理

- 相対位置調整
  - 他のウィンドウとの位置関係を調整
  - パディング設定による配置制御
  - マルチモニター対応

## 注意点
- Pythonの実行環境が必要です（Anaconda環境推奨）
- 適切なショートカット設定が必要です
- ショートカットのパスは実際の環境に合わせて変更してください
- フォルダの書き込み権限が必要です
- VSCode連携にはVSCodeのインストールが必要です
- ウィンドウ位置の保存にはCSVファイルの書き込み権限が必要です
- 実行時のエラーは自動的にメッセージボックスで表示されます
