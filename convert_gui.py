import contextlib
import os
import queue
import shutil
import tempfile
import threading
import traceback
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import convert_pptx  # 同一資料夾內的 convert_pptx.py（不需修改）

DONE = object()
q = queue.Queue()


class QueueWriter:
    def write(self, s):
        if s:
            q.put(s)
        return len(s or "")

    def flush(self):
        pass


def run_convert(files, out_dir):
    tmp = Path(tempfile.mkdtemp(prefix="pptx_"))
    try:
        used = set()
        for f in files:
            name, stem, suffix, i = f.name, f.stem, f.suffix, 2
            while name.lower() in used:
                name = f"{stem}_{i}{suffix}"
                i += 1
            used.add(name.lower())
            shutil.copy2(f, tmp / name)
        out_dir.mkdir(parents=True, exist_ok=True)
        convert_pptx.INPUT_DIR = tmp
        convert_pptx.OUTPUT_DIR = out_dir
        w = QueueWriter()
        with contextlib.redirect_stdout(w), contextlib.redirect_stderr(w):
            convert_pptx.main()
    except SystemExit as e:
        q.put(f"{e}\n")
    except Exception:
        q.put(traceback.format_exc())
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
        q.put(DONE)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PPTX 文字轉換工具 (PPTX → Markdown / JSON)")
        self.geometry("760x600")
        self.files = []
        self.out_var = tk.StringVar()
        self.build()
        self.after(100, self.poll)

    def build(self):
        pad = {"padx": 10, "pady": 4}
        ttk.Label(self, text="1. 加入 PPTX 檔案").pack(anchor="w", **pad)
        box = ttk.Frame(self)
        box.pack(fill="both", expand=True, padx=10)
        self.lb = tk.Listbox(box, height=8, selectmode="extended")
        sb = ttk.Scrollbar(box, command=self.lb.yview)
        self.lb.config(yscrollcommand=sb.set)
        self.lb.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        bar = ttk.Frame(self)
        bar.pack(fill="x", **pad)
        ttk.Button(bar, text="加入檔案…", command=self.add_files).pack(side="left")
        ttk.Button(bar, text="加入資料夾…", command=self.add_folder).pack(side="left", padx=6)
        ttk.Button(bar, text="移除選取", command=self.remove_sel).pack(side="left")
        ttk.Button(bar, text="清空", command=self.clear).pack(side="left", padx=6)

        ttk.Label(self, text="2. 選擇輸出資料夾").pack(anchor="w", **pad)
        row = ttk.Frame(self)
        row.pack(fill="x", padx=10)
        ttk.Entry(row, textvariable=self.out_var).pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="瀏覽…", command=self.pick_out).pack(side="left", padx=6)

        self.run_btn = ttk.Button(self, text="3. 開始轉換", command=self.start)
        self.run_btn.pack(anchor="w", **pad)
        self.prog = ttk.Progressbar(self, mode="indeterminate")
        self.prog.pack(fill="x", padx=10)

        ttk.Label(self, text="執行記錄").pack(anchor="w", **pad)
        self.log = tk.Text(self, height=9, state="disabled")
        self.log.pack(fill="both", expand=True, padx=10, pady=(0, 6))
        self.open_btn = ttk.Button(self, text="開啟輸出資料夾", command=self.open_out, state="disabled")
        self.open_btn.pack(anchor="e", padx=10, pady=(0, 10))

    def add_paths(self, paths):
        for p in paths:
            p = Path(p)
            if p.suffix.lower() == ".pptx" and not p.name.startswith("~$") and p not in self.files:
                self.files.append(p)
                self.lb.insert("end", str(p))
        if self.files and not self.out_var.get():
            self.out_var.set(str(self.files[0].parent / "output"))

    def add_files(self):
        self.add_paths(filedialog.askopenfilenames(title="選擇 PPTX 檔案", filetypes=[("PowerPoint", "*.pptx")]))

    def add_folder(self):
        d = filedialog.askdirectory(title="選擇含有 PPTX 的資料夾")
        if d:
            self.add_paths(sorted(Path(d).glob("*.pptx")))

    def remove_sel(self):
        for i in reversed(self.lb.curselection()):
            self.lb.delete(i)
            del self.files[i]

    def clear(self):
        self.lb.delete(0, "end")
        self.files.clear()

    def pick_out(self):
        d = filedialog.askdirectory(title="選擇輸出資料夾")
        if d:
            self.out_var.set(d)

    def write_log(self, s):
        self.log.config(state="normal")
        self.log.insert("end", s)
        self.log.see("end")
        self.log.config(state="disabled")

    def start(self):
        if not self.files:
            messagebox.showwarning("提示", "請先加入至少一個 .pptx 檔案。")
            return
        if not self.out_var.get().strip():
            messagebox.showwarning("提示", "請選擇輸出資料夾。")
            return
        self.run_btn.config(state="disabled")
        self.open_btn.config(state="disabled")
        self.log.config(state="normal")
        self.log.delete("1.0", "end")
        self.log.config(state="disabled")
        self.prog.start(12)
        threading.Thread(target=run_convert, args=(list(self.files), Path(self.out_var.get())), daemon=True).start()

    def poll(self):
        try:
            while True:
                item = q.get_nowait()
                if item is DONE:
                    self.prog.stop()
                    self.run_btn.config(state="normal")
                    self.open_btn.config(state="normal")
                    self.write_log("\n完成。\n")
                else:
                    self.write_log(item)
        except queue.Empty:
            pass
        self.after(100, self.poll)

    def open_out(self):
        p = self.out_var.get()
        if os.path.isdir(p):
            os.startfile(p)


if __name__ == "__main__":
    App().mainloop()
