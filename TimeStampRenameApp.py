import os
import re
import threading
import shutil
import subprocess
import time
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog
from PIL import Image

# --- ドラッグ&ドロップ用のライブラリ ---
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    TkinterDnD = None 

def get_file_info(filepath):
    filename = os.path.basename(filepath)
    date_str = ""
    sort_key = ""
    try:
        with Image.open(filepath) as img:
            exif = img.getexif()
            if exif:
                raw_dt = exif.get(36867) or exif.get(36868) or exif.get(306)
                if not raw_dt:
                    try:
                        exif_sub = exif.get_ifd(0x8769)
                        raw_dt = exif_sub.get(36867) or exif_sub.get(36868)
                    except:
                        pass
                if raw_dt:
                    raw_dt = str(raw_dt).strip()
                    clean_dt = re.sub(r'\D', '', raw_dt)
                    if len(clean_dt) >= 8:
                        date_str = clean_dt[:8]
                        sort_key = raw_dt
    except Exception:
        pass 

    if not date_str:
        date_8 = re.search(r"(\d{8})", filename)
        if date_8:
            d8 = date_8.group(1)
            try:
                datetime.strptime(d8, '%Y%m%d') 
                date_str = d8
                sort_key = d8
            except ValueError:
                pass
        if not date_str:
            date_6 = re.search(r"(\d{6})", filename)
            if date_6:
                d6 = date_6.group(1)
                try:
                    month_part = int(d6[4:])
                    if 1 <= month_part <= 12:
                        date_str = d6
                        sort_key = d6
                except ValueError:
                    pass
    if not sort_key:
        sort_key = str(os.path.getmtime(filepath))
    return sort_key, date_str

def sanitize_filename(name):
    return re.sub(r'[\\/:*?"<>|]', '', name)

class RenameApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Time Stamp Rename App")
        self.root.geometry("600x800") 
        
        self.bg_main = "#ececec" 
        self.root.configure(bg=self.bg_main)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        if TkinterDnD:
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind('<<Drop>>', self.handle_drop)

        self.style = ttk.Style()
        self.style.theme_use('default') 
        self.style.configure("TLabelframe", background="#ffffff", relief="flat")
        self.style.configure("TLabelframe.Label", background="#ffffff", foreground="#333333", font=("", 10, "bold"))
        
        # --- ボタンのスタイル設定（他のUIパーツに馴染ませた薄い枠線） ---
        standard_border = "#d0d0d0"
        
        self.style.configure("Select.TButton", 
                             background="#dddddd", 
                             foreground="#000000", 
                             font=("", 9, "bold"), 
                             padding=(0, 10), 
                             relief="solid", 
                             borderwidth=1,
                             bordercolor=standard_border)
        self.style.map("Select.TButton", background=[('active', '#eeeeee')])
        
        self.style.configure("Clear.TButton", 
                             background="#ffffff",          # #f0faff から #ffffff (白) に変更
                             foreground="#000000", 
                             font=("", 9, "bold"), 
                             padding=(0, 10), 
                             relief="solid", 
                             borderwidth=1,
                             bordercolor=standard_border)
        self.style.map("Clear.TButton", background=[('active', '#f0faff')]) # #ffffff から #f0faff (薄い青) に変更
        
        self.style.configure("Run.TButton", background="#000000", foreground="#ffffff", font=("", 11, "bold"), padding=(0, 12), relief="flat")
        self.style.map("Run.TButton", background=[('active', '#ff9800')], foreground=[('active', '#ffffff')]) 

        self.raw_data = [] 
        self.target_folder = ""
        self.is_running = False
        self.create_widgets()

    def create_widgets(self):
        # 1. リネーム設定
        self.frame_setting = ttk.LabelFrame(self.root, text=" 1. リネーム設定 ", padding=12)
        self.frame_setting.pack(padx=15, pady=(15, 5), fill="x")
        
        input_frame = tk.Frame(self.frame_setting, bg="#ffffff")
        input_frame.pack(fill="x", pady=2)
        tk.Label(input_frame, text="接頭辞:", bg="#ffffff", fg="#000000").grid(row=0, column=0, sticky="w")
        self.prefix_var = tk.StringVar(value="A")
        self.prefix_var.trace_add("write", lambda *args: self.update_example())
        tk.Entry(input_frame, textvariable=self.prefix_var, width=15, relief="solid", borderwidth=1).grid(row=0, column=1, padx=5)
        tk.Label(input_frame, text="開始番号:", bg="#ffffff", fg="#000000").grid(row=0, column=2, sticky="w", padx=(15,0))
        self.num_var = tk.StringVar(value="0001")
        self.num_var.trace_add("write", lambda *args: self.update_example())
        tk.Entry(input_frame, width=8, textvariable=self.num_var, relief="solid", borderwidth=1).grid(row=0, column=3, padx=5)
        
        self.var_include_date = tk.BooleanVar(value=True)
        self.check_date = tk.Checkbutton(self.frame_setting, text=" 撮影日またはファイル名の数値を付ける ", variable=self.var_include_date, command=self.update_example, bg="#f5f5f5", selectcolor="#cce5ff", activebackground="#ffffff", font=("", 10), relief="groove", pady=6, padx=10)
        self.check_date.pack(anchor="w", pady=8)
        
        sort_frame = tk.Frame(self.frame_setting, bg="#ffffff")
        sort_frame.pack(fill="x", pady=2)
        tk.Label(sort_frame, text="並び替え:", bg="#ffffff", fg="#333333", font=("", 9)).pack(side="left")
        self.sort_mode = tk.StringVar(value="date")
        
        tk.Radiobutton(sort_frame, text="日付順", variable=self.sort_mode, value="date", command=self.resort_and_preview, bg="#ffffff", font=("", 9)).pack(side="left", padx=10)
        tk.Radiobutton(sort_frame, text="名前順", variable=self.sort_mode, value="name", command=self.resort_and_preview, bg="#ffffff", font=("", 9)).pack(side="left")

        # 2. 完成イメージ
        self.example_frame = tk.Frame(self.root, bg="#f0f8ff", padx=10, pady=10, highlightbackground="#4682b4", highlightthickness=1)
        self.example_frame.pack(padx=15, pady=8, fill="x")
        self.example_label = tk.Label(self.example_frame, text="", bg="#f0f8ff", fg="#000000", font=("", 10, "bold"))
        self.example_label.pack()
        self.update_example()

        # 3. 選択・クリアボタン
        tk.Label(self.root, text=" 2. ファイルの追加 ", bg=self.bg_main, fg="#333333", font=("", 9, "bold")).pack(anchor="w", padx=20, pady=(2,0))
        self.frame_action = tk.Frame(self.root, bg=self.bg_main)
        self.frame_action.pack(fill="x", padx=15, pady=5)
        
        self.btn_select_dir = ttk.Button(self.frame_action, text="フォルダ選択", style="Select.TButton", command=self.select_folder)
        self.btn_select_dir.pack(side="left", fill="x", expand=True, padx=(0,2))
        
        self.btn_select_files = ttk.Button(self.frame_action, text="ファイル選択", style="Select.TButton", command=self.select_files)
        self.btn_select_files.pack(side="left", fill="x", expand=True, padx=2)
        
        self.btn_clear = ttk.Button(self.frame_action, text="クリア", style="Clear.TButton", command=self.clear_data)
        self.btn_clear.pack(side="left", fill="x", expand=True, padx=(2,0))

        # 4. プレビューエリア
        tk.Label(self.root, text=" 3. プレビュー (ドロップ可能) ", bg=self.bg_main, fg="#333333", font=("", 9, "bold")).pack(anchor="w", padx=20, pady=(2,0))
        preview_container = tk.Frame(self.root, bg="#ffffff", padx=2, pady=2, highlightbackground="#bbbbbb", highlightthickness=1)
        preview_container.pack(fill="both", expand=True, padx=15, pady=5)
        sc = tk.Scrollbar(preview_container)
        sc.pack(side="right", fill="y")
        self.listbox = tk.Listbox(preview_container, yscrollcommand=sc.set, font=("Consolas", 10), bg="#ffffff", fg="#000000", relief="flat", highlightthickness=0)
        self.listbox.pack(side="left", fill="both", expand=True)
        sc.config(command=self.listbox.yview)

        # ステータス・実行
        self.status_label = tk.Label(self.root, text="項目を選択またはドロップしてください", bg=self.bg_main, fg="#333333", font=("", 10, "bold"))
        self.status_label.pack(pady=(2, 0))
        self.percent_label = tk.Label(self.root, text="", bg=self.bg_main, fg="#000000", font=("Consolas", 18, "bold"))
        self.percent_label.pack()

        self.btn_run = ttk.Button(self.root, text="実行開始", state="disabled", style="Run.TButton", command=self.start_execution)
        self.btn_run.pack(fill="x", padx=30, pady=(5, 15))

    # --- ロジック部分 ---
    def clear_data(self):
        self.raw_data = []
        self.target_folder = ""
        self.listbox.delete(0, tk.END)
        self.status_label.config(text="項目を選択またはドロップしてください", fg="#333333")
        self.percent_label.config(text="")
        self.btn_run.state(["disabled"])
        self.btn_run.config(text="実行開始")

    def handle_drop(self, event):
        data = event.data
        paths = re.findall(r'\{.*?\}|\S+', data)
        paths = [p.strip('{}') for p in paths]
        self.start_analysis(paths)

    def select_folder(self):
        path = filedialog.askdirectory()
        if path: self.start_analysis([path])

    def select_files(self):
        paths = filedialog.askopenfilenames()
        if paths: self.start_analysis(list(paths))

    def start_analysis(self, paths):
        if paths:
            first_path = paths[0]
            self.target_folder = first_path if os.path.isdir(first_path) else os.path.dirname(first_path)
        self.btn_select_dir.state(["disabled"])
        self.btn_select_files.state(["disabled"])
        self.btn_clear.state(["disabled"])
        self.percent_label.config(text="")
        self.status_label.config(text="解析中...", fg="#000000")
        threading.Thread(target=self.analyze_files_task, args=(paths,), daemon=True).start()

    def analyze_files_task(self, paths):
        try:
            self.raw_data = []
            exclude = ["changed", "Thumbs.db", ".DS_Store", "desktop.ini"]
            all_files = []
            for p in paths:
                if os.path.isdir(p):
                    files = [os.path.join(p, f) for f in os.listdir(p) if os.path.isfile(os.path.join(p, f)) and f not in exclude]
                    all_files.extend(files)
                elif os.path.isfile(p):
                    if os.path.basename(p) not in exclude:
                        all_files.append(p)
            total = len(all_files)
            for i, filepath in enumerate(all_files):
                sort_val, date_str = get_file_info(filepath)
                self.raw_data.append({'path': filepath, 'old_name': os.path.basename(filepath), 'date_str': date_str, 'sort_val': str(sort_val)})
                if i % 20 == 0: 
                    self.root.after(0, lambda i=i, t=total: self.status_label.config(text=f"解析中... ({i+1}/{total})"))
            self.root.after(0, self.resort_and_preview)
            self.root.after(0, lambda: self.btn_run.state(["!disabled"]))
            self.root.after(0, lambda: [self.btn_select_dir.state(["!disabled"]), self.btn_select_files.state(["!disabled"]), self.btn_clear.state(["!disabled"])])
        except Exception as e:
            self.root.after(0, lambda e=e: self.status_label.config(text=f"エラー: {e}"))

    def resort_and_preview(self):
        if not self.raw_data: return
        self.raw_data.sort(key=lambda x: x['sort_val'] if self.sort_mode.get() == "date" else x['old_name'])
        self.update_preview()

    def update_preview(self, *args):
        if not self.target_folder or not self.raw_data: return
        prefix = sanitize_filename(self.prefix_var.get().strip())
        num_str = self.num_var.get().strip()
        include_date = self.var_include_date.get()
        try: start_num, digits = int(num_str), len(num_str)
        except: return
        self.listbox.delete(0, tk.END)
        for i, item in enumerate(self.raw_data, start_num):
            ext = os.path.splitext(item['old_name'])[1]
            parts = []
            if include_date and item['date_str']:
                parts.append(item['date_str'])
            if prefix: parts.append(prefix)
            parts.append(str(i).zfill(digits))
            item['new_name'] = "_".join(parts) + ext
            self.listbox.insert(tk.END, f"{item['old_name']} → {item['new_name']}")
        self.status_label.config(text=f"準備完了: {len(self.raw_data)}件", fg="#000000")

    def update_example(self, *args):
        prefix = sanitize_filename(self.prefix_var.get().strip())
        num_str = self.num_var.get().strip()
        include_date = self.var_include_date.get()
        example_date = datetime.now().strftime("%Y%m%d")
        parts = [example_date] if include_date else []
        if prefix: parts.append(prefix)
        parts.append(num_str if num_str else "0001")
        self.example_label.config(text=f"完成イメージ： {'_'.join(parts)}.jpg")
        self.update_preview()

    def update_progress_ui(self, current, total):
        if total == 0: return
        percent = int(current / total * 100)
        self.root.after(0, lambda: self.percent_label.config(text=f"{percent}%"))

    def start_execution(self):
        if not self.raw_data: return
        self.is_running = True
        self.percent_label.config(text="0%") 
        self.btn_run.state(["disabled"])
        self.btn_select_dir.state(["disabled"])
        self.btn_select_files.state(["disabled"])
        self.btn_clear.state(["disabled"])
        threading.Thread(target=self.execute_copy, daemon=True).start()

    def execute_copy(self):
        out = os.path.join(self.target_folder, "changed")
        if not os.path.exists(out): os.makedirs(out)
        total = len(self.raw_data)
        for i, item in enumerate(self.raw_data):
            if not self.is_running: return 
            dest_path = os.path.join(out, item['new_name'])
            shutil.copy2(item['path'], dest_path)
            self.update_progress_ui(i + 1, total)
            self.root.after(0, lambda i=i: self.status_label.config(text=f"コピー中... ({i+1}/{total})"))
            time.sleep(0.01)
        self.is_running = False
        self.root.after(0, lambda: self.finish(out))

    def finish(self, out_path):
        self.status_label.config(text="完了しました。", fg="#000000")
        self.btn_select_dir.state(["!disabled"])
        self.btn_select_files.state(["!disabled"])
        self.btn_clear.state(["!disabled"])
        self.btn_run.config(text="完了")
        if os.name == 'nt': os.startfile(out_path)
        else: subprocess.run(["open", out_path])

    def on_closing(self):
        self.is_running = False
        self.root.destroy()

if __name__ == "__main__":
    if TkinterDnD:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    app = RenameApp(root)
    root.mainloop()