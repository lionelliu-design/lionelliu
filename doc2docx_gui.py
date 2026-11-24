#!/usr/bin/env python3
"""
DOC 轉 DOCX 批次轉換工具 (圖形介面版)

功能: 拖曳或瀏覽選擇檔案/資料夾 -> 選輸出位置 -> 按「開始轉換」
需求: Python 3.8+ 以及 LibreOffice (或 Windows + Microsoft Word + pywin32)
選用: pip install tkinterdnd2   (啟用拖曳功能)
"""
import os
import platform
import queue
import shutil
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except Exception:  # noqa: BLE001
    HAS_DND = False

EXT_MAP = {".doc": ".docx", ".dot": ".dotx", ".rtf": ".docx"}
IS_WIN = platform.system() == "Windows"
NO_WINDOW = subprocess.CREATE_NO_WINDOW if IS_WIN else 0


# ============================================================ 核心轉換邏輯
def app_dir():
    """exe 所在資料夾 (打包後) 或腳本所在資料夾。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def find_soffice():
    # 1) 優先使用隨程式一起打包的 LibreOffice (免安裝)
    base = app_dir()
    for rel in ("LibreOffice/program/soffice.exe", "LibreOffice/program/soffice"):
        c = base / rel
        if c.exists():
            return str(c)
    # 2) 退而求其次:使用系統已安裝的 LibreOffice
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    for c in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
              "/Applications/LibreOffice.app/Contents/MacOS/soffice",
              "/usr/bin/soffice", "/opt/libreoffice/program/soffice"):
        if os.path.exists(c):
            return c
    return None


def has_word():
    if not IS_WIN:
        return False
    try:
        import winreg
        import win32com.client  # noqa: F401
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Word.Application"))
        return True
    except Exception:  # noqa: BLE001
        return False


def dest_for(src, base, out_root):
    name = src.stem + EXT_MAP[src.suffix.lower()]
    if out_root is None:
        return src.with_name(name)
    return out_root / src.relative_to(base).parent / name


def lo_convert(soffice, srcs, outdir, timeout):
    outdir.mkdir(parents=True, exist_ok=True)
    profile = tempfile.mkdtemp(prefix="lo_profile_")
    err = ""
    try:
        by_fmt = {}
        for s in srcs:
            by_fmt.setdefault(EXT_MAP[s.suffix.lower()].lstrip("."), []).append(s)
        for fmt, group in by_fmt.items():
            cmd = [soffice, f"-env:UserInstallation={Path(profile).as_uri()}",
                   "--headless", "--norestore", "--convert-to", fmt,
                   "--outdir", str(outdir)] + [str(s) for s in group]
            try:
                r = subprocess.run(cmd, stdin=subprocess.DEVNULL,
                                   capture_output=True, text=True,
                                   timeout=timeout * len(group),
                                   creationflags=NO_WINDOW)
                err += r.stderr
            except subprocess.TimeoutExpired:
                err += "逾時"
    finally:
        shutil.rmtree(profile, ignore_errors=True)
    return err.strip()


def run_conversion(jobs, engine, workers, overwrite, cancel, emit, timeout=120):
    """
    jobs: [(src, dst)]。emit(kind, *args) 回報進度:
      ("status", src, 文字) / ("progress", done, total)
    回傳 (成功, 失敗, 略過, 取消)
    """
    total = len(jobs)
    counts = {"ok": 0, "fail": 0, "skip": 0, "cancel": 0, "done": 0}
    lock = threading.Lock()

    def finish(src, kind, text):
        with lock:
            counts[kind] += 1
            counts["done"] += 1
            done = counts["done"]
        emit("status", src, text)
        emit("progress", done, total)

    todo = []
    for src, dst in jobs:
        if dst.exists() and not overwrite:
            finish(src, "skip", "略過 (已存在)")
        else:
            todo.append((src, dst))

    if engine == "word":
        try:
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
        except Exception as e:  # noqa: BLE001
            for src, _ in todo:
                finish(src, "fail", f"無法啟動 Word: {e}")
            return tuple(counts[k] for k in ("ok", "fail", "skip", "cancel"))
        fmt = {".docx": 12, ".dotx": 14}
        try:
            for src, dst in todo:
                if cancel.is_set():
                    finish(src, "cancel", "已取消")
                    continue
                emit("status", src, "轉換中…")
                try:
                    dst.parent.mkdir(parents=True, exist_ok=True)
                    doc = word.Documents.Open(str(src), ReadOnly=True,
                                              AddToRecentFiles=False,
                                              ConfirmConversions=False)
                    doc.SaveAs2(str(dst), FileFormat=fmt[dst.suffix.lower()])
                    doc.Close(False)
                    finish(src, "ok", "完成")
                except Exception as e:  # noqa: BLE001
                    finish(src, "fail", f"失敗: {e}")
        finally:
            word.Quit()
    else:
        soffice = find_soffice()
        if not soffice:
            for src, _ in todo:
                finish(src, "fail", "找不到 LibreOffice")
            return tuple(counts[k] for k in ("ok", "fail", "skip", "cancel"))

        groups = {}
        for src, dst in todo:
            groups.setdefault(dst.parent, []).append((src, dst))
        tasks = []
        for outdir, items in groups.items():
            for i in range(0, len(items), 5):          # 小批次 => 進度較即時
                tasks.append((outdir, items[i:i + 5]))

        def task(outdir, items):
            if cancel.is_set():
                for src, _ in items:
                    finish(src, "cancel", "已取消")
                return
            for src, _ in items:
                emit("status", src, "轉換中…")
            err = lo_convert(soffice, [s for s, _ in items], outdir, timeout)
            for src, dst in items:
                if dst.exists() and dst.stat().st_size > 0:
                    finish(src, "ok", "完成")
                else:
                    finish(src, "fail", "失敗" + (f": {err[:80]}" if err else ""))

        with ThreadPoolExecutor(max_workers=max(1, workers)) as ex:
            futs = [ex.submit(task, o, i) for o, i in tasks]
            for f in futs:
                f.result()

    return tuple(counts[k] for k in ("ok", "fail", "skip", "cancel"))


# ============================================================ 圖形介面
class App:
    def __init__(self, root):
        self.root = root
        root.title("DOC 轉 DOCX 批次轉換工具")
        root.geometry("820x640")
        root.minsize(720, 560)

        self.items = {}            # iid(str) -> (src Path, base Path)
        self.q = queue.Queue()
        self.cancel = threading.Event()
        self.running = False

        self.recursive = tk.BooleanVar(value=True)
        self.out_mode = tk.StringVar(value="same")
        self.out_dir = tk.StringVar()
        self.overwrite = tk.BooleanVar(value=False)
        self.engine = tk.StringVar(value="auto")
        self.workers = tk.IntVar(value=min(4, os.cpu_count() or 2))

        self._build()
        self.root.after(100, self._poll)

    # ---------------------------------------------------------- 版面
    def _build(self):
        pad = {"padx": 10, "pady": 5}

        # --- 檔案清單
        top = ttk.LabelFrame(self.root, text="1. 選擇要轉換的檔案")
        top.pack(fill="both", expand=True, **pad)

        bar = ttk.Frame(top)
        bar.pack(fill="x", padx=8, pady=6)
        ttk.Button(bar, text="新增檔案…", command=self.add_files).pack(side="left")
        ttk.Button(bar, text="新增資料夾…", command=self.add_folder).pack(side="left", padx=6)
        ttk.Checkbutton(bar, text="包含子資料夾", variable=self.recursive).pack(side="left", padx=6)
        ttk.Button(bar, text="清空", command=self.clear).pack(side="right")
        ttk.Button(bar, text="移除選取", command=self.remove_selected).pack(side="right", padx=6)

        cols = ("name", "folder", "status")
        wrap = ttk.Frame(top)
        wrap.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.tree = ttk.Treeview(wrap, columns=cols, show="headings", selectmode="extended")
        for c, t, w in (("name", "檔名", 220), ("folder", "所在資料夾", 360), ("status", "狀態", 120)):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor="w")
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.bind("<Delete>", lambda e: self.remove_selected())

        hint = "可將檔案或資料夾直接拖曳到此視窗" if HAS_DND else \
               "提示: 安裝 tkinterdnd2 (pip install tkinterdnd2) 可啟用拖曳功能"
        self.count_lbl = ttk.Label(top, text=f"尚未加入檔案    {hint}", foreground="#666")
        self.count_lbl.pack(anchor="w", padx=10, pady=(0, 6))

        if HAS_DND:
            for w in (self.root, self.tree):
                w.drop_target_register(DND_FILES)
                w.dnd_bind("<<Drop>>", self._on_drop)

        # --- 輸出位置
        out = ttk.LabelFrame(self.root, text="2. 轉換後的儲存位置")
        out.pack(fill="x", **pad)
        ttk.Radiobutton(out, text="與原檔案相同的資料夾 (預設)", variable=self.out_mode,
                        value="same", command=self._sync_out).grid(row=0, column=0, columnspan=3,
                                                                    sticky="w", padx=8, pady=(6, 0))
        ttk.Radiobutton(out, text="自訂資料夾:", variable=self.out_mode, value="custom",
                        command=self._sync_out).grid(row=1, column=0, sticky="w", padx=8, pady=6)
        self.out_entry = ttk.Entry(out, textvariable=self.out_dir, state="disabled")
        self.out_entry.grid(row=1, column=1, sticky="ew", padx=4)
        self.out_btn = ttk.Button(out, text="瀏覽…", command=self.choose_out, state="disabled")
        self.out_btn.grid(row=1, column=2, padx=8)
        out.columnconfigure(1, weight=1)

        # --- 選項
        opt = ttk.LabelFrame(self.root, text="3. 選項")
        opt.pack(fill="x", **pad)
        ttk.Checkbutton(opt, text="覆寫已存在的 .docx", variable=self.overwrite).pack(side="left", padx=8, pady=6)
        ttk.Label(opt, text="引擎:").pack(side="left", padx=(16, 4))
        engines = ["auto", "libreoffice"] + (["word"] if IS_WIN else [])
        ttk.Combobox(opt, textvariable=self.engine, values=engines, width=12,
                     state="readonly").pack(side="left")
        ttk.Label(opt, text="平行數:").pack(side="left", padx=(16, 4))
        ttk.Spinbox(opt, from_=1, to=16, textvariable=self.workers, width=4).pack(side="left")

        # --- 執行
        run = ttk.Frame(self.root)
        run.pack(fill="x", padx=10, pady=(4, 10))
        self.start_btn = ttk.Button(run, text="開始轉換", command=self.start)
        self.start_btn.pack(side="left")
        self.cancel_btn = ttk.Button(run, text="取消", command=self.cancel.set, state="disabled")
        self.cancel_btn.pack(side="left", padx=6)
        self.open_btn = ttk.Button(run, text="開啟輸出資料夾", command=self.open_output, state="disabled")
        self.open_btn.pack(side="left", padx=6)
        self.progress = ttk.Progressbar(run, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=10)
        self.prog_lbl = ttk.Label(run, text="0 / 0", width=10)
        self.prog_lbl.pack(side="left")

    # ---------------------------------------------------------- 加入檔案
    def _on_drop(self, event):
        self.add_paths(self.root.tk.splitlist(event.data))

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="選擇檔案",
            filetypes=[("Word 舊格式", "*.doc *.dot *.rtf"), ("所有檔案", "*.*")])
        self.add_paths(paths)

    def add_folder(self):
        d = filedialog.askdirectory(title="選擇資料夾")
        if d:
            self.add_paths([d])

    def add_paths(self, paths):
        added = 0
        for raw in paths:
            p = Path(raw).resolve()
            found = []
            if p.is_file():
                found = [(p, p.parent)]
            elif p.is_dir():
                it = p.rglob("*") if self.recursive.get() else p.glob("*")
                found = [(f, p) for f in sorted(it) if f.is_file()]
            for f, base in found:
                if f.suffix.lower() not in EXT_MAP or f.name.startswith("~$"):
                    continue
                iid = str(f)
                if iid in self.items:
                    continue
                self.items[iid] = (f, base)
                self.tree.insert("", "end", iid=iid, values=(f.name, str(f.parent), "等待中"))
                added += 1
        self._update_count()
        if not added and paths:
            messagebox.showinfo("提示", "沒有找到新的 .doc / .dot / .rtf 檔案。")

    def remove_selected(self):
        if self.running:
            return
        for iid in self.tree.selection():
            self.tree.delete(iid)
            self.items.pop(iid, None)
        self._update_count()

    def clear(self):
        if self.running:
            return
        self.tree.delete(*self.tree.get_children())
        self.items.clear()
        self._update_count()

    def _update_count(self):
        n = len(self.items)
        self.count_lbl.config(text=f"已加入 {n} 個檔案" if n else "尚未加入檔案")

    # ---------------------------------------------------------- 輸出位置
    def _sync_out(self):
        state = "normal" if self.out_mode.get() == "custom" else "disabled"
        self.out_entry.config(state=state)
        self.out_btn.config(state=state)

    def choose_out(self):
        d = filedialog.askdirectory(title="選擇輸出資料夾")
        if d:
            self.out_dir.set(d)

    def open_output(self):
        target = self.out_dir.get() if self.out_mode.get() == "custom" else None
        if target is None and self.items:
            target = str(next(iter(self.items.values()))[0].parent)
        if not target or not Path(target).exists():
            return
        if IS_WIN:
            os.startfile(target)  # type: ignore[attr-defined]
        elif platform.system() == "Darwin":
            subprocess.run(["open", target])
        else:
            subprocess.run(["xdg-open", target])

    # ---------------------------------------------------------- 執行
    def start(self):
        if not self.items:
            messagebox.showwarning("提示", "請先加入要轉換的檔案。")
            return
        out_root = None
        if self.out_mode.get() == "custom":
            if not self.out_dir.get().strip():
                messagebox.showwarning("提示", "請選擇輸出資料夾。")
                return
            out_root = Path(self.out_dir.get()).resolve()

        engine = self.engine.get()
        if engine == "auto":
            engine = "word" if has_word() else "libreoffice"
        if engine == "libreoffice" and not find_soffice():
            messagebox.showerror("找不到 LibreOffice",
                                 "找不到轉換引擎。請確認 LibreOffice 資料夾與程式放在一起,\n"
                                 "或安裝 LibreOffice: https://www.libreoffice.org/download/")
            return
        if engine == "word" and not has_word():
            messagebox.showerror("無法使用 Word", "需要 Windows、Microsoft Word 與 pywin32。")
            return

        jobs = [(s, dest_for(s, b, out_root)) for s, b in self.items.values()]
        for src, _ in jobs:
            self.tree.set(str(src), "status", "等待中")
        self.progress.config(maximum=len(jobs), value=0)
        self.prog_lbl.config(text=f"0 / {len(jobs)}")
        self.cancel.clear()
        self._set_running(True)

        def work():
            try:
                res = run_conversion(jobs, engine, self.workers.get(), self.overwrite.get(),
                                     self.cancel, lambda *a: self.q.put(a))
            except Exception as e:  # noqa: BLE001
                self.q.put(("error", str(e)))
                return
            self.q.put(("finished", res))

        threading.Thread(target=work, daemon=True).start()

    def _set_running(self, running):
        self.running = running
        self.start_btn.config(state="disabled" if running else "normal")
        self.cancel_btn.config(state="normal" if running else "disabled")
        if running:
            self.open_btn.config(state="disabled")

    def _poll(self):
        try:
            while True:
                msg = self.q.get_nowait()
                kind = msg[0]
                if kind == "status":
                    if self.tree.exists(str(msg[1])):
                        self.tree.set(str(msg[1]), "status", msg[2])
                elif kind == "progress":
                    self.progress.config(value=msg[1])
                    self.prog_lbl.config(text=f"{msg[1]} / {msg[2]}")
                elif kind == "finished":
                    ok, fail, skip, cancel = msg[1]
                    self._set_running(False)
                    self.open_btn.config(state="normal")
                    text = f"成功 {ok} 個,失敗 {fail} 個,略過 {skip} 個"
                    if cancel:
                        text += f",取消 {cancel} 個"
                    messagebox.showinfo("轉換完成", text)
                elif kind == "error":
                    self._set_running(False)
                    messagebox.showerror("發生錯誤", msg[1])
        except queue.Empty:
            pass
        self.root.after(100, self._poll)


def main():
    root = TkinterDnD.Tk() if HAS_DND else tk.Tk()
    try:
        ttk.Style().theme_use("vista" if IS_WIN else "clam")
    except tk.TclError:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
