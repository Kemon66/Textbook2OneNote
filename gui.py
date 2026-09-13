#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Textbook2OneNote - 图形界面入口
将教材文件（PDF / Word）拖入拖放区，先「预览目录」确认自动识别的章节结构，
再点击「开始导入」生成 OneNote 笔记本。
"""

import os
import queue
import sys
import threading
import traceback
from datetime import datetime

from tkinterdnd2 import DND_FILES, TkinterDnD

# 保证打包为 exe 后也能 import importer
if getattr(sys, "frozen", False):
    _base = os.path.dirname(sys.executable)
else:
    _base = os.path.dirname(os.path.abspath(__file__))
if _base not in sys.path:
    sys.path.insert(0, _base)

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import importer

# 主题色
BG = "#141b24"
PANEL = "#1f2a38"
PANEL_2 = "#263444"
ACCENT = "#4f8ef7"
ACCENT_HOVER = "#6ea6ff"
ACCENT_DARK = "#3b6ec4"
TEXT = "#e8eef5"
MUTED = "#8fa3b8"
BORDER = "#35465a"
DROP_BORDER = "#5d7ea6"
DROP_ACTIVE_BG = "#2b3b4f"
SUCCESS = "#63d68b"
WARN = "#f0b45a"
DANGER = "#ef6a6a"
LOG_BG = "#101720"
LOG_FG = "#9fc3e8"


class App:
    def __init__(self, root):
        self.root = root
        root.title("Textbook2OneNote - 教材导入 OneNote")
        root.geometry("900x660")
        root.minsize(760, 560)
        root.configure(bg=BG)

        self.file_path = None
        self.file_name = ""
        self.file_size = 0
        self._analysis = None
        self._busy = False
        self._cancel = threading.Event()
        self._queue = queue.Queue()
        self._success_outdir = None
        self._drop_widgets = []

        self._init_style()
        self._build_ui()

        root.bind("<Control-o>", lambda e: self.choose())
        root.bind("<Control-Return>", lambda e: self.start())
        root.after(80, self._poll_queue)

    # ---------- 样式 ----------
    def _init_style(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Treeview", background=PANEL, fieldbackground=PANEL,
                        foreground=TEXT, borderwidth=0, rowheight=24,
                        font=("Microsoft YaHei UI", 9))
        style.configure("Treeview.Heading", background=PANEL_2, foreground=MUTED,
                        relief="flat", font=("Microsoft YaHei UI", 9, "bold"))
        style.map("Treeview", background=[("selected", ACCENT_DARK)],
                  foreground=[("selected", "#ffffff")])
        style.configure("TEntry", fieldbackground=PANEL, foreground=TEXT,
                        insertcolor=TEXT, bordercolor=BORDER)
        style.map("TEntry", fieldbackground=[("disabled", PANEL)])
        style.configure("TProgressbar", troughcolor=PANEL, background=ACCENT,
                        bordercolor=PANEL, lightcolor=ACCENT, darkcolor=ACCENT)

    # ---------- 界面 ----------
    def _build_ui(self):
        self.root.grid_columnconfigure(0, minsize=330)
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(1, weight=1)

        # 顶部标题
        header = tk.Frame(self.root, bg=BG)
        header.grid(row=0, column=0, columnspan=2, sticky="ew", padx=24, pady=(18, 10))
        tk.Label(header, text="Textbook2OneNote", bg=BG, fg=TEXT,
                 font=("Segoe UI Semibold", 18)).pack(side="left")
        tk.Label(header, text="识别教材目录 · 自动生成 OneNote 笔记本",
                 bg=BG, fg=MUTED, font=("Microsoft YaHei UI", 9)
                 ).pack(side="left", padx=(12, 0), pady=(7, 0))
        self.status_chip = tk.Label(header, text="就绪", bg=PANEL, fg=SUCCESS,
                                    font=("Microsoft YaHei UI", 9), padx=10, pady=3)
        self.status_chip.pack(side="right")

        # 左列：文件选择 / 输出目录 / 操作按钮 / 进度
        left = tk.Frame(self.root, bg=BG)
        left.grid(row=1, column=0, sticky="nsew", padx=(24, 12), pady=(0, 12))
        left.grid_columnconfigure(0, weight=1)

        self.drop = tk.Frame(left, bg=PANEL, highlightbackground=DROP_BORDER,
                             highlightthickness=2, bd=0, cursor="hand2")
        self.drop.grid(row=0, column=0, sticky="ew")
        self.drop.grid_columnconfigure(0, weight=1)

        # 底部文件信息先 pack，内容区再占满剩余空间，避免固定高度在高 DPI 下互相遮挡
        self.file_meta_var = tk.StringVar(value="")
        self._drop_label(self.drop, textvariable=self.file_meta_var, bg=PANEL,
                         fg=SUCCESS, font=("Consolas", 9), anchor="w"
                         ).pack(side="bottom", fill="x", padx=14, pady=(0, 10))

        body = tk.Frame(self.drop, bg=PANEL)
        body.bind("<Button-1>", self.choose)
        self._drop_widgets.append(body)
        body.pack(fill="both", expand=True, padx=12, pady=(12, 2))

        self._drop_label(body, text="⬇", bg=PANEL, fg=ACCENT,
                         font=("Segoe UI Symbol", 30)).pack(pady=(10, 0))
        self.drop_main_text = self._drop_label(
            body, text="拖拽教材到此处", bg=PANEL, fg=TEXT,
            font=("Microsoft YaHei UI", 13, "bold"))
        self.drop_main_text.pack(pady=(4, 2))
        self.drop_sub_text = self._drop_label(
            body, text="支持 PDF / Word（.docx）", bg=PANEL, fg=MUTED,
            font=("Microsoft YaHei UI", 9))
        self.drop_sub_text.pack(pady=(0, 10))

        self.drop.drop_target_register(DND_FILES)
        self.drop.dnd_bind("<<Drop>>", self.on_drop)
        self.drop.dnd_bind("<<DragEnter>>", lambda e: self.set_drop_style_active(True))
        self.drop.dnd_bind("<<DragLeave>>", lambda e: self.set_drop_style_active(False))
        self.drop.bind("<Button-1>", self.choose)

        # 输出目录
        out_block = tk.Frame(left, bg=BG)
        out_block.grid(row=1, column=0, sticky="ew", pady=(16, 0))
        out_block.grid_columnconfigure(0, weight=1)
        tk.Label(out_block, text="输出目录", bg=BG, fg=TEXT,
                 font=("Microsoft YaHei UI", 10)).grid(row=0, column=0, columnspan=2,
                                                       sticky="w", pady=(0, 6))
        self.out_var = tk.StringVar(value=importer.find_output_dir())
        self.out_entry = ttk.Entry(out_block, textvariable=self.out_var)
        self.out_entry.grid(row=1, column=0, sticky="ew", ipady=3)
        self._btn(out_block, "浏览…", self.browse_out, kind="plain"
                  ).grid(row=1, column=1, padx=(8, 0))

        # 笔记本名称（默认取书名/文件名，可编辑）
        name_block = tk.Frame(left, bg=BG)
        name_block.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        name_block.grid_columnconfigure(0, weight=1)
        tk.Label(name_block, text="笔记本名称", bg=BG, fg=TEXT,
                 font=("Microsoft YaHei UI", 10)).grid(row=0, column=0,
                                                       sticky="w", pady=(0, 6))
        self.name_var = tk.StringVar(value="")
        self.name_entry = ttk.Entry(name_block, textvariable=self.name_var)
        self.name_entry.grid(row=1, column=0, sticky="ew", ipady=3)

        # 操作按钮
        btn_row = tk.Frame(left, bg=BG)
        btn_row.grid(row=3, column=0, sticky="ew", pady=(16, 0))
        btn_row.grid_columnconfigure(0, weight=1)
        btn_row.grid_columnconfigure(1, weight=1)
        self.preview_btn = self._btn(btn_row, "预览目录", self.preview, kind="plain")
        self.preview_btn.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.start_btn = self._btn(btn_row, "开始导入", self.start, kind="accent")
        self.start_btn.grid(row=0, column=1, sticky="ew", padx=(6, 0))
        self.cancel_btn = self._btn(left, "取消", self.cancel, kind="danger")
        self.cancel_btn.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        self.cancel_btn.grid_remove()

        # 进度
        prog = tk.Frame(left, bg=BG)
        prog.grid(row=5, column=0, sticky="ew", pady=(16, 0))
        prog.grid_columnconfigure(0, weight=1)
        self.stage_var = tk.StringVar(value="")
        tk.Label(prog, textvariable=self.stage_var, anchor="w", bg=BG, fg=MUTED,
                 font=("Microsoft YaHei UI", 9)).grid(row=0, column=0,
                                                      sticky="ew", pady=(0, 4))
        self.progress = ttk.Progressbar(prog, mode="determinate", maximum=100)
        self.progress.grid(row=1, column=0, sticky="ew")

        # 右列：预览树 + 日志
        right = tk.Frame(self.root, bg=BG)
        right.grid(row=1, column=1, sticky="nsew", padx=(12, 24), pady=(0, 12))
        right.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(1, weight=1)

        pv_header = tk.Frame(right, bg=BG)
        pv_header.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        tk.Label(pv_header, text="目录结构预览", bg=BG, fg=TEXT,
                 font=("Microsoft YaHei UI", 12, "bold")).pack(side="left")
        self.summary_var = tk.StringVar(value="选择文件后点击「预览目录」")
        tk.Label(pv_header, textvariable=self.summary_var, bg=BG, fg=MUTED,
                 font=("Microsoft YaHei UI", 9), anchor="e").pack(side="right")

        tree_frame = tk.Frame(right, bg=PANEL, highlightbackground=BORDER,
                              highlightthickness=1)
        tree_frame.grid(row=1, column=0, sticky="nsew")
        tree_frame.grid_columnconfigure(0, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(tree_frame, columns=("num", "loc"),
                                 show="tree headings", selectmode="browse")
        self.tree.heading("#0", text="章节结构")
        self.tree.column("#0", width=300, minwidth=180, stretch=True)
        self.tree.heading("num", text="编号")
        self.tree.column("num", width=70, minwidth=50, stretch=False, anchor="w")
        self.tree.heading("loc", text="位置")
        self.tree.column("loc", width=90, minwidth=60, stretch=False, anchor="w")
        self.tree.tag_configure("part", foreground=ACCENT)
        self.tree.tag_configure("chapter", foreground="#dbe8ff")
        self.tree.tag_configure("main", foreground=TEXT)
        self.tree.tag_configure("end", foreground=MUTED)
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        self.tree_hint = tk.Label(
            tree_frame, text="尚未识别目录\n点击「预览目录」查看识别结果",
            bg=PANEL, fg=MUTED, font=("Microsoft YaHei UI", 10), justify="center")
        self.tree_hint.place(relx=0.5, rely=0.5, anchor="center")

        log_head = tk.Frame(right, bg=BG)
        log_head.grid(row=2, column=0, sticky="ew", pady=(12, 4))
        tk.Label(log_head, text="运行日志", bg=BG, fg=TEXT,
                 font=("Microsoft YaHei UI", 10)).pack(side="left")
        self._btn(log_head, "清空", self.clear_log, kind="plain").pack(side="right")
        self.log_box = tk.Text(right, height=6, bg=LOG_BG, fg=LOG_FG,
                               insertbackground=MUTED, relief="flat",
                               font=("Consolas", 9), state="disabled",
                               padx=8, pady=6, highlightbackground=BORDER,
                               highlightthickness=1)
        self.log_box.grid(row=3, column=0, sticky="ew")

        # 底部状态栏
        status = tk.Frame(self.root, bg=PANEL_2)
        status.grid(row=2, column=0, columnspan=2, sticky="ew")
        self.status_label = tk.Label(status, text="拖入教材文件即可开始", anchor="w",
                                     bg=PANEL_2, fg=MUTED,
                                     font=("Microsoft YaHei UI", 9), padx=16, pady=7)
        self.status_label.pack(side="left", fill="x", expand=True)
        self.open_btn = tk.Button(status, text="打开输出文件夹",
                                  command=self.open_folder, bg=PANEL_2, fg=ACCENT,
                                  activebackground=PANEL_2, activeforeground=ACCENT_HOVER,
                                  bd=0, cursor="hand2",
                                  font=("Microsoft YaHei UI", 9), padx=12, pady=5)
        self.open_btn.pack(side="right", padx=(0, 10), pady=4)
        self.open_btn.pack_forget()

    def _btn(self, parent, text, command, kind="plain"):
        if kind == "accent":
            bg, fg, abg = ACCENT, "#ffffff", ACCENT_HOVER
        elif kind == "danger":
            bg, fg, abg = "#5b3036", "#f5c6cb", "#6f3a41"
        else:
            bg, fg, abg = PANEL_2, TEXT, "#33445a"
        return tk.Button(parent, text=text, command=command, bg=bg, fg=fg,
                         activebackground=abg, activeforeground=fg, bd=0,
                         relief="flat", padx=12, pady=7, cursor="hand2",
                         font=("Microsoft YaHei UI", 10), disabledforeground=MUTED)

    def _drop_label(self, parent, **kw):
        kw.setdefault("bg", PANEL)
        lbl = tk.Label(parent, **kw)
        lbl.bind("<Button-1>", self.choose)
        self._drop_widgets.append(lbl)
        return lbl

    # ---------- 文件选择 ----------
    def set_drop_style_active(self, active):
        self.drop.configure(highlightbackground=ACCENT if active else DROP_BORDER,
                            highlightcolor=ACCENT if active else DROP_BORDER,
                            bg=DROP_ACTIVE_BG if active else PANEL)
        for w in self._drop_widgets:
            w.configure(bg=DROP_ACTIVE_BG if active else PANEL)

    def on_drop(self, event):
        self.set_drop_style_active(False)
        path = self._clean_drop_path(event.data)
        if path:
            self.set_file(path)
        return event.action

    @staticmethod
    def _clean_drop_path(data):
        """解析拖放数据（可能带花括号 / 多文件 / 空格路径），取第一个存在的文件。"""
        import re
        candidates = []
        for m in re.finditer(r"\{([^{}]+)\}", data):
            candidates.append(m.group(1).strip())
        rest = re.sub(r"\{[^{}]*\}", "", data)
        for seg in rest.split():
            seg = seg.strip()
            if seg and seg not in candidates:
                candidates.append(seg)
        for c in candidates:
            if os.path.exists(c):
                return c
        return None

    def set_file(self, path):
        if self._busy or not path or not os.path.isfile(path):
            return
        self.file_path = path
        self.file_name = os.path.basename(path)
        self.file_size = os.path.getsize(path)
        self._analysis = None
        self._success_outdir = None
        self.open_btn.pack_forget()

        ext = os.path.splitext(path)[1].lower()
        kind = "PDF" if ext == ".pdf" else ("Word" if ext == ".docx" else "未知类型")
        self.drop_main_text.configure(text="已选择文件")
        self.drop_sub_text.configure(text=self._shorten(self.file_name, 34))
        self.file_meta_var.set("%s  ·  %.1f MB" % (kind, self.file_size / 1048576.0))
        self.name_var.set(os.path.splitext(self.file_name)[0])
        self._clear_tree()
        self.summary_var.set("点击「预览目录」查看自动识别的章节结构")
        self.status_label.configure(text="已选择：" + self.file_name)
        self._set_state("ready")

    def choose(self, *_):
        if self._busy:
            return
        p = filedialog.askopenfilename(
            title="选择教材文件（PDF / Word）",
            filetypes=[("教材文件", "*.pdf *.docx"), ("PDF", "*.pdf"),
                       ("Word", "*.docx"), ("所有文件", "*.*")],
        )
        if p:
            self.set_file(p)

    def browse_out(self, *_):
        if self._busy:
            return
        cur = self.out_var.get().strip()
        d = filedialog.askdirectory(initialdir=cur or None, title="选择输出目录")
        if d:
            self.out_var.set(os.path.normpath(d))

    @staticmethod
    def _shorten(text, n=42):
        return text if len(text) <= n else text[: n - 1] + "…"

    # ---------- 日志 ----------
    def _append_log(self, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_box.configure(state="normal")
        self.log_box.insert("end", "[%s] %s\n" % (ts, msg))
        if int(self.log_box.index("end-1c").split(".")[0]) > 1500:
            self.log_box.delete("1.0", "1000.0")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def clear_log(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

    # ---------- 执行 ----------
    def _check_file(self):
        if not self.file_path or not os.path.isfile(self.file_path):
            messagebox.showwarning("提示", "请先选择或拖入教材文件")
            return False
        return True

    def preview(self):
        if self._busy:
            return
        if not self._check_file():
            return
        self._analysis = None
        self._clear_tree()
        self.summary_var.set("正在识别目录结构…")
        self._start_worker(self._preview_worker)

    def start(self, *_):
        if self._busy:
            return
        if not self._check_file():
            return
        notegen_exe = self._find_notegen()
        if not os.path.isfile(notegen_exe):
            messagebox.showerror("缺少组件",
                                 "未找到 OneNote 生成引擎 notegen.exe。\n"
                                 "请确认程序文件完整。")
            return
        out_root = self.out_var.get().strip() or importer.find_output_dir()
        if not os.path.isdir(out_root):
            messagebox.showerror("输出目录无效", "输出目录不存在：\n%s" % out_root)
            return
        self._start_worker(self._import_worker, args=(notegen_exe, out_root))

    def cancel(self):
        if not self._busy:
            return
        self._cancel.set()
        self.cancel_btn.configure(state="disabled")
        self.stage_var.set("正在取消…")
        self._append_log("正在取消…")

    def _start_worker(self, fn, args=()):
        self._cancel.clear()
        self._set_state("busy")
        self._append_log("开始处理 %s" % self.file_name)
        t = threading.Thread(target=self._wrap_worker, args=(fn, args), daemon=True)
        t.start()

    def _wrap_worker(self, fn, args):
        try:
            fn(*args)
        except importer.ImportCancelled:
            self._queue.put({"k": "cancelled"})
        except Exception as e:
            self._queue.put({"k": "crash", "err": str(e),
                             "tb": traceback.format_exc()})
        finally:
            self._queue.put({"k": "finished"})

    def _queue_log(self, msg):
        self._queue.put({"k": "log", "v": msg})

    def _queue_progress(self, stage, done, total):
        self._queue.put({"k": "progress", "stage": stage,
                         "done": done, "total": total})

    def _preview_worker(self):
        a = importer.analyze_file(self.file_path, cancel=self._cancel,
                                  progress=self._queue_progress,
                                  log_cb=self._queue_log)
        self._queue.put({"k": "preview_done", "data": a})

    def _import_worker(self, notegen_exe, out_root):
        a = self._analysis
        if a is None or a.get("path") != os.path.abspath(self.file_path):
            a = importer.analyze_file(self.file_path, cancel=self._cancel,
                                      progress=self._queue_progress,
                                      log_cb=self._queue_log)
            self._queue.put({"k": "analysis", "data": a})
        if not a.get("ok"):
            self._queue.put({"k": "reason", "data": a})
            return

        raw_name = self.name_var.get().strip()
        if raw_name:
            a = dict(a)
            a["name"] = importer.sanitize_book_name(raw_name)

        ext = os.path.splitext(self.file_path)[1].lower()
        if ext == ".pdf":
            r = importer.process_pdf(self.file_path, out_root, notegen_exe,
                                     cancel=self._cancel,
                                     progress=self._queue_progress,
                                     log_cb=self._queue_log, analysis=a)
        else:
            r = importer.process_word(self.file_path, out_root, notegen_exe,
                                      cancel=self._cancel,
                                      progress=self._queue_progress,
                                      log_cb=self._queue_log, analysis=a)
        if r is None:
            self._queue.put({"k": "reason", "data": a})
            return
        self._queue.put({"k": "import_done", "result": r[0], "outdir": r[1]})

    # ---------- 主线程消息泵 ----------
    def _poll_queue(self):
        try:
            while True:
                self._handle_msg(self._queue.get_nowait())
        except queue.Empty:
            pass
        except Exception:
            traceback.print_exc()
        self.root.after(80, self._poll_queue)

    def _handle_msg(self, m):
        k = m["k"]
        if k == "log":
            self._append_log(m["v"])
        elif k == "progress":
            self._set_progress(m.get("stage"), m.get("done"), m.get("total"))
        elif k == "preview_done":
            self._on_preview_done(m["data"])
        elif k == "analysis":
            self._store_analysis(m["data"])
        elif k == "reason":
            self._store_analysis(m["data"])
            self._show_reason_dialog(m["data"])
        elif k == "import_done":
            self._on_import_done(m["result"], m["outdir"])
        elif k == "cancelled":
            self._append_log("已取消")
            self.status_label.configure(text="已取消")
            self.status_chip.configure(text="已取消", fg=WARN)
        elif k == "crash":
            self._append_log("发生错误: %s" % m["err"])
            self._append_log(m["tb"])
            self.status_chip.configure(text="错误", fg=DANGER)
            messagebox.showerror("错误", m["err"])
        elif k == "finished":
            self._set_state("ready")

    def _set_state(self, mode):
        if mode == "busy":
            self._busy = True
            self.preview_btn.configure(state="disabled")
            self.start_btn.configure(state="disabled", text="处理中…")
            self.cancel_btn.configure(state="normal")
            self.cancel_btn.grid()
            self.status_chip.configure(text="处理中", fg=WARN)
            self.progress.configure(mode="indeterminate")
            self.progress.start(14)
            self.stage_var.set("准备中…")
        else:
            self._busy = False
            self.preview_btn.configure(state="normal")
            can_start = self._analysis is None or self._analysis.get("ok")
            self.start_btn.configure(state="normal" if can_start else "disabled",
                                     text="开始导入")
            self.cancel_btn.grid_remove()
            self.progress.stop()
            self.progress.configure(mode="determinate", value=0)
            self.stage_var.set("")
            if can_start and self.status_chip.cget("text") in ("处理中", "已取消"):
                self.status_chip.configure(text="就绪", fg=SUCCESS)

    def _set_progress(self, stage, done, total):
        if stage:
            self.stage_var.set(stage)
        if total and total > 0:
            self.progress.stop()
            self.progress.configure(mode="determinate", maximum=max(1, total),
                                    value=min(done, total))
        else:
            self.progress.configure(mode="indeterminate")
            self.progress.start(14)

    # ---------- 结果展示 ----------
    def _clear_tree(self):
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        self.tree_hint.place(relx=0.5, rely=0.5, anchor="center")

    def _populate_tree(self, a):
        self._clear_tree()
        self.tree_hint.place_forget()
        kind = a.get("kind")

        def add(node, parent=""):
            name = node.get("name") or ("—" if node.get("level") == "chapter" else "")
            values = (node.get("num") or "", self._loc_text(node, kind))
            iid = self.tree.insert(parent, "end", text=name, values=values,
                                   tags=(node.get("level") or "main",))
            if node.get("level") in ("part", "chapter"):
                self.tree.item(iid, open=True)
            for child in node.get("children", []):
                add(child, iid)

        for node in a.get("tree", []):
            add(node)

    @staticmethod
    def _loc_text(node, kind):
        if kind == "word":
            if node.get("level") == "main" and "count" in node:
                return "%d 段" % node["count"]
            return ""
        if node.get("level") == "part":
            return ""
        p = node.get("print")
        if p is None:
            return "—"
        return "p.%s" % p

    def _store_analysis(self, a):
        self._analysis = a
        if a.get("ok"):
            self._populate_tree(a)
            counts = a.get("counts", {})
            if a.get("kind") == "pdf":
                leftovers_pages = sum(len(g.get("pages", []))
                                      for g in a.get("leftovers", []))
                self.summary_var.set("%d 部分 · %d 章 · %d 小节 · 共 %d 页" %
                                     (counts.get("parts", 0), counts.get("chapters", 0),
                                      counts.get("mains", 0), a.get("render_pages", 0)))
                if leftovers_pages:
                    self._append_log("删减内容整理: %d 页自动归入「其他内容」分区组" %
                                     leftovers_pages)
            else:
                self.summary_var.set("%d 部分 · %d 章 · %d 小节（Word 标题样式）" %
                                     (counts.get("parts", 0), counts.get("chapters", 0),
                                      counts.get("mains", 0)))
            self.name_var.set(a.get("name") or self.name_var.get())
            self.status_label.configure(text="目录识别完成，可开始导入")
        else:
            self._clear_tree()
            self.summary_var.set(self._reason_summary(a))
            self.status_chip.configure(text="无法识别", fg=DANGER)
            self.status_label.configure(text=self._reason_summary(a))

    @staticmethod
    def _reason_summary(a):
        m = {
            "scanned": "扫描版 PDF：无文字层，请先 OCR",
            "no_toc": "未找到目录页（Contents / 目录）",
            "empty_toc": "目录页存在，但未解析出章节条目",
            "no_headings": "未识别到 Heading 1/2/3 标题样式",
            "doc_legacy": "旧版 .doc 不支持，请另存为 .docx",
            "parse_error": "Word 文档解析失败",
            "unsupported": "不支持的文件类型",
        }
        return m.get(a.get("reason"), "无法识别目录结构")

    def _on_preview_done(self, a):
        self._store_analysis(a)
        if not a.get("ok"):
            self._show_reason_dialog(a)

    def _show_reason_dialog(self, a):
        reason = a.get("reason")
        stats = a.get("stats") or {}
        if reason == "scanned":
            messagebox.showwarning(
                "无法识别",
                "该 PDF 为扫描版/图片型（无文字层），无法识别目录。\n"
                "请先对教材进行 OCR，或更换带文字层的版本后重试。\n\n"
                "文本层统计：采样 %d 页，有实质文本 %d 页（%.1f%%），"
                "总字符 %d，单页文本中位数 %d" %
                (stats.get("sampled", 0), stats.get("text_pages", 0),
                 stats.get("ratio", 0) * 100, stats.get("text_chars", 0),
                 stats.get("len_median", 0)))
        elif reason == "no_toc":
            messagebox.showwarning(
                "无法识别",
                "PDF 含文字层，但前 40 页未找到目录页（Contents / 目录）。\n"
                "请确认教材包含目录页后重试，或更换含标准目录的版本。")
        elif reason == "empty_toc":
            messagebox.showwarning(
                "无法识别",
                "已找到目录页，但未能解析出任何章节条目。\n"
                "目录可能是图片形式，或排版过于特殊。")
        elif reason == "no_headings":
            messagebox.showwarning(
                "无法识别",
                "未识别到标题层级。\n请确认 Word 文档使用了"
                "「标题 1 / 标题 2 / 标题 3」样式。")
        elif reason == "doc_legacy":
            messagebox.showwarning(
                "格式不支持",
                "暂不支持旧版 .doc 格式。\n请用 Word 打开后「另存为 .docx」再导入。")
        elif reason == "parse_error":
            messagebox.showerror("解析失败",
                                 "Word 文档解析失败：%s" % a.get("error"))
        else:
            messagebox.showwarning("不支持",
                                   "不支持的文件类型，请选择 PDF 或 .docx 文件。")

    def _on_import_done(self, result, outdir):
        if result.get("ok"):
            self._append_log("生成成功！")
            self._append_log("笔记本位置: %s" % outdir)
            self._append_log("结构: %d 分区组 / %d 分区 / %d 页" %
                             (result.get("groups", 0), result.get("sections", 0),
                              result.get("pages", 0)))
            self._append_log("排序校验: %s" %
                             ("通过" if result.get("sortedOk") else "失败"))
            self._success_outdir = outdir
            self.status_label.configure(text="导入完成：" + outdir)
            self.status_chip.configure(text="完成", fg=SUCCESS)
            self.open_btn.pack(side="right", padx=(0, 10), pady=4)
            messagebox.showinfo("完成", "生成成功！\n位置：%s" % outdir)
        else:
            self._append_log("生成失败: %s" % result.get("error"))
            self.status_chip.configure(text="失败", fg=DANGER)
            messagebox.showerror("失败", result.get("error"))

    def open_folder(self):
        d = self._success_outdir
        if not d or not os.path.isdir(d):
            return
        try:
            os.startfile(d)
        except Exception as e:
            messagebox.showerror("无法打开", str(e))

    def _find_notegen(self):
        """定位 notegen.exe：打包后从内部解压目录(_MEIPASS)取，否则取同目录。"""
        if getattr(sys, "frozen", False):
            mp = getattr(sys, "_MEIPASS", "")
            cand = os.path.join(mp, "notegen.exe")
            if cand and os.path.exists(cand):
                return cand
        return os.path.join(_base, "notegen.exe")


def main():
    try:
        root = TkinterDnD.Tk()
        App(root)
        root.mainloop()
    except Exception:
        traceback.print_exc()
        try:
            messagebox.showerror("错误", traceback.format_exc())
        except Exception:
            pass


if __name__ == "__main__":
    # 带参数时走 importer 命令行（便于脚本/复现问题），无参数时打开图形界面
    if len(sys.argv) > 1:
        sys.exit(importer.main())
    main()
