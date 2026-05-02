# Kindle AutoCapture

PC版 Kindle を自動でページ送りしながら、各ページのスクリーンショットを連番保存するためのデスクトップアプリです。
個人利用（購入済み書籍の研究・学習用途）を前提にしています。

**操作・設定の詳細は [MANUAL.md](MANUAL.md) を参照してください。**

## 主な機能

- Kindle for PC のウィンドウ自動検出
- 指定範囲の高速スクリーンショット（mss / DXGI Desktop Duplication）
- 自動ページ送り（キーボード/マウスクリック方式）
- pHashによる重複ページ検知 → 最終ページ自動停止
- PNG / JPEG / WebP での連番保存
- 開始 / 一時停止 / 停止 / 再開
- グローバルホットキー（F9: 開始/停止, F10: 一時停止, Esc: 緊急停止）
- リアルタイムログ・進捗表示
- ROI（取得範囲）をドラッグで選択するオーバーレイ

## 動作環境

- Windows 10 / 11
- Python 3.10 以上（推奨: 3.11）
- Kindle for PC

## セットアップ

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python main.py
```

## EXE化

```powershell
pip install pyinstaller
pyinstaller build.spec --clean --noconfirm
```

## ディレクトリ構成

```
kindle_autocapture/
├── main.py                # エントリーポイント
├── app/                   # オーケストレーター
├── capture/               # ウィンドウ検出・SS取得・重複検知
├── navigation/            # ページ送り
├── storage/               # 画像保存
├── gui/                   # PyQt6 GUI
├── utils/                 # ロガー・パス管理
└── config/                # 設定読み書き
```

## 設定ファイル

実行時には `%APPDATA%\KindleAutoCapture\config.json` が生成されます。
GUIから変更した設定はここに永続化されます。

## 注意事項

本ツールは購入済みコンテンツの**個人的な複製の範囲**で使用してください。
取得画像の二次配布・公衆送信は著作権侵害となる可能性があります。
