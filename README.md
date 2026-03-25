# Excel Automation Tool

複数の売上Excelファイルを集計し、テンプレートExcelに出力する自動化ツールです。  
CLI実行とGUI（Tkinter）実行の両方に対応しています。

---

## 機能

- 複数Excelファイルの読み込み
- 売上データの統合
- ピボット集計
- テンプレートExcelへの書き込み
- グラフ更新
- GUI / CLI 両対応
- loggingによるログ出力
- config.ini による設定管理

---

## 実行方法

### CLI

```bash
python run.py
```

### GUI

```bash
python run.py --gui
```

---
## フォルダ構成

```text
excel_automation/
├─ core/
│   ├─ main_logic.py
│   ├─ config_loader.py
│   └─ logger_setup.py
│
├─ gui/
│   └─ gui_app.py
│
├─ run.py
├─ config.ini
└─ requirements.txt
```

## 使用技術

- Python
- pandas
- openpyxl
- tkinter
- logging
- configparser

## 今後の改善予定

- GUIスレッドの安全性改善
- config検証の共通化
- テンプレートExcelの検証強化
- loggingの改善
- エラーハンドリング強化

## 補足
config.iniに実行パスを設定して使用します。