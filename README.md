# 複数Excelファイル自動集計ツール

## 概要
複数の売上Excelファイルを自動で読み込み、
店舗別・担当者別に集計して1つのExcelへ出力するPythonツールです。

## 主な機能
- 複数Excelファイルの自動読み込み
- 必要列チェック
- 読み込み失敗ファイルのスキップ
- 店舗別売上集計
- 担当者×商品別集計
- 複数シートでExcel出力
- 列幅・見出し・数値表示の自動整形

## 使用技術
- Python
- pandas
- openpyxl

## 実行方法
pip install -r requirements.txt
python main.py