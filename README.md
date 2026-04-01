# Excel Automation Tool

複数の売上Excelファイルを集計し、テンプレートExcelに出力する自動化ツールです。  
CLIとGUI（Tkinter）の両方で実行できます。

---

## 特徴

- 複数Excelファイルの自動読み込み
- pandasによるデータ統合・集計
- テンプレートExcelへの書き込み（openpyxl）
- グラフの自動更新
- CLI / GUI 両対応
- config.ini による設定管理
- loggingによるログ出力（ファイル / GUI / コンソール）
- 入力・テンプレートの検証機能あり

---

## セットアップ

```bash
pip install -r requirements.txt
```

---
## 設定ファイル
`config.ini`を作成してください。  
(`config.example.ini`をコピーして使用)

```ini
[PATH]
input_folder = C:\path\to\input_folder
output_path = C:\path\to\output.xlsx
template_path = C:\path\to\template.xlsx

[LOG]
log_file = logs/app.log
```

---
## 実行方法

### CLI

```bash
python run.py
```

### GUI

```bash
python run.py ---gui
```

---

## フォルダ構成

```text
excel_automation/
├─ core/
│   ├─ main_logic.py
│   ├─ config_loader.py
│   ├─ config_validator.py
│   └─ logger_setup.py
│
├─ gui/
│   └─ gui_app.py
│
├─ run.py
├─ config.ini
├─ config.example.ini
└─ requirements.txt
```

---

## 設計ポイント

- 責務分離
  - core：ロジック
  - gui：表示
- 設定の外部化（config.ini）
- 例外処理とバリデーションの分離
- loggingの統合管理
  - ファイル / GUI / コンソールに出力

---

## 使用技術
- Python
- pandas
- openpyxl
- tkinter
- logging
- configparser

---

## 今後の改善予定
- データクリア範囲の動的化
- グラフ更新ロジックの改善
- loggingのさらなる改善
- エラーハンドリングの強化