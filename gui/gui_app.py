# ====================
# Tkinter GUI (GUI専用)
# ====================

import tkinter as tk
from tkinter.scrolledtext import ScrolledText
from tkinter import filedialog
from tkinter import messagebox
import threading
import queue
import sys
import os

from core.main_logic import main
from core.config_loader import load_config
from core.logger_setup import setup_logger
from core.config_validator import validate_config, validate_config_file_exists

log_queue = queue.Queue()

# --- GUI用：ログをGUIに流すためのクラス ---
class QueueLogger:
    def write(self, msg):
        log_queue.put(msg)
    def flush(self):
        pass   

# GUIログ出力の有効化
sys.stdout = QueueLogger()
sys.stderr = QueueLogger()

def safe_load_config():
    try:
        config_path = "config.ini"
        validate_config_file_exists(config_path)
        config = load_config(config_path)
        validate_config(config)
        return config
    
    except Exception as e:
        messagebox.showerror("設定エラー", str(e))
        return None

# --- GUI本体 ---
def gui_main():
    config = safe_load_config()
    if config is None:
        return

    setup_logger(config["LOG"]["log_file"])

    root = tk.Tk()
    root.title("売上レポート自動生成ツール")

    #　入力フォルダ
    tk.Label(root, text="入力フォルダ").grid(row=0, column=0)
    input_var = tk.StringVar()
    try :
        input_var.set(config["PATH"]["input_folder"].replace("\\*.xlsx", ""))
    except Exception:
        input_var.set("")
    input_entry = tk.Entry(root, textvariable=input_var, width=50)
    input_entry.grid(row=0, column=1)
    tk.Button(root, text="選択", command=lambda: (input_var.set(filedialog.askdirectory()), validate_paths())).grid(row=0, column=2)

    # テンプレート
    tk.Label(root, text="テンプレート").grid(row=1, column=0)
    template_var = tk.StringVar()
    try :
        template_var.set(config["PATH"]["template_path"])
    except Exception:
        template_var.set("")
    template_entry = tk.Entry(root, textvariable=template_var, width=50)
    template_entry.grid(row=1, column=1)
    tk.Button(root, text="選択", command=lambda: (template_var.set(filedialog.askopenfilename()),validate_paths())).grid(row=1, column=2)

    # 出力ファイル
    tk.Label(root, text="出力ファイル").grid(row=2, column=0)
    output_var = tk.StringVar()
    try :
        output_var.set(config["PATH"]["output_path"])
    except Exception:
        output_var.set("")
    output_entry = tk.Entry(root, textvariable=output_var, width=50)
    output_entry.grid(row=2, column=1)
    tk.Button(root, text="選択", command=lambda: (output_var.set(filedialog.asksaveasfilename()), validate_paths())).grid(row=2, column=2)

    # ログ表示
    log_box = ScrolledText(root, width=80, height=20)
    log_box.grid(row=4, column=0, columnspan=3)

    def update_log():
        while not log_queue.empty():
            msg = log_queue.get()
            log_box.insert(tk.END, msg)
            log_box.see(tk.END)
        root.after(100, update_log)

    def worker():
        try:
            main(config)
        finally:
            root.after(0, lambda: run_button.config(state="normal"))
    
    def run_main():
        run_button.config(state="disabled")
        config["PATH"]["input_folder"] = os.path.join(input_var.get(), "*.xlsx")
        config["PATH"]["template_path"] = template_var.get()
        config["PATH"]["output_path"] = output_var.get()
        threading.Thread(target=worker, daemon=True).start()
    
    def validate_paths():
        ok = True

        # 入力フォルダ
        if os.path.isdir(input_var.get()):
            input_entry.config(bg="white")
        else:
            input_entry.config(bg="#ffcccc")
            ok = False

        # テンプレート
        if os.path.isfile(template_var.get()):
            template_entry.config(bg="white")
        else:
            template_entry.config(bg="#ffcccc")
            ok = False

        # 出力ファイル（フォルダが存在するかチェック）
        out_dir = os.path.dirname(output_var.get())
        if out_dir and os.path.isdir(out_dir):
            output_entry.config(bg="white")
        else:
            output_entry.config(bg="#ffcccc")
            ok = False

        # ボタン制御
        if ok:
            run_button.config(state="normal")
        else:
            run_button.config(state="disabled")

    # Entry の変更を監視
    input_var.trace_add("write", lambda *args: validate_paths())
    template_var.trace_add("write", lambda *args: validate_paths())
    output_var.trace_add("write", lambda *args: validate_paths())

    # 実行ボタン（初期は無効）
    run_button = tk.Button(root, text="実行", command=run_main, state="disabled")
    run_button.grid(row=3, column=1)



    update_log()
    root.mainloop()

if __name__ == "__main__":
    gui_main()

