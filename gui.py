#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Textbook2OneNote - 图形界面入口（拖放版）
将教材文件（PDF / Word）拖入大拖放区，点击「开始导入」生成 OneNote 笔记本。
"""

import os
import sys
import threading
import traceback

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
BG = "#1b2430"
PANEL = "#25303f"
ACCENT = "#4f8ef7"
ACCENT_HOVER = "#6ea6ff"
TEXT = "#e8eef5"
MUTED = "#8fa3b8"
DROP_BORDER = "#5d7ea6"


class App:
    def __init__(self, root):
        self.root = root
        root.title("Textbook2OneNote - 教材导入 OneNote")
        root.geometry("720x560")
        root.minsize(600, 480)
        root.configure(bg=BG)

        # 主容器
        main = tk.Frame(root, bg=BG)
        main.pack(fill="both", expand=True, padx=36, pady=28)

        # ---------- 视觉重心：选择教材拖放区 ----------
        title = tk.Label(main, text="选择教材文件", bg=BG, fg=TEXT,
                         font=("Microsoft YaHei UI", 20, "bold"))
        title.pack(pady=(0, 6))
        sub = tk.Label(main, text="将 PDF / Word 教材拖入下方区域", bg=BG, fg=MUTED,
                       font=("Microsoft YaHei UI", 11))
        sub.pack(pady=(0, 18))

        self.drop = tk.Frame(main, bg=PANEL, highlightbackground=DROP_BORDER,
                             highlightthickness=2, bd=0, cursor="hand2")
        self.drop.pack(fill="both", expand=True)

        # 拖放区中央提示（图标 + 文案）
        center = tk.Frame(self.drop, bg=PANEL)
        center.place(relx=0.5, rely=0.5, anchor="center")
        icon = tk.Label(center, text="⬇", bg=PANEL, fg=ACCENT,
                        font=("Segoe UI Symbol", 52))
        icon.pack()
        self.drop_main_text = tk.Label(center, text="拖拽文件到此处", bg=PANEL, fg=TEXT,
                                       font=("Microsoft YaHei UI", 16, "bold"))
        self.drop_main_text.pack(pady=(6, 2))
        self.drop_sub_text = tk.Label(center, text="或点击此处选择文件", bg=PANEL, fg=MUTED,
                                      font=("Microsoft YaHei UI", 10))
        self.drop_sub_text.pack()

        # 已选文件路径（拖放/选择后显示）
        self.file_var = tk.StringVar()
        self.file_label = tk.Label(self.drop, textvariable=self.file_var, bg=PANEL,
                                   fg="#9fd3a5", font=("Consolas", 10), wraplength=560)
        self.file_label.pack(side="bottom", pady=(0, 14))

        # 拖放事件
        self.drop.drop_target_register(DND_FILES)
        self.drop.dnd_bind("<<Drop>>", self.on_drop)
        self.drop.dnd_bind("<<DragEnter>>", lambda e: self.set_drop_style_active(True))
        self.drop.dnd_bind("<<DragLeave>>", lambda e: self.set_drop_style_active(False))
        self.drop.bind("<Button-1>", self.choose)

        # ---------- 次要区域：输出目录 + 开始按钮 + 日志 ----------
        sec = tk.Frame(main, bg=BG)
        sec.pack(fill="x", pady=(20, 0))

        out_row = tk.Frame(sec, bg=BG)
        out_row.pack(fill="x")
        tk.Label(out_row, text="输出目录", bg=BG, fg=TEXT,
                 font=("Microsoft YaHei UI", 10)).pack(side="left", padx=(0, 10))
        self.out_var = tk.StringVar(value=importer.find_output_dir())
        self.out_entry = ttk.Entry(out_row, textvariable=self.out_var)
        self.out_entry.pack(side="left", fill="x", expand=True, ipady=3)

        self.start_btn = tk.Button(sec, text="开始导入", bg=ACCENT, fg="#ffffff",
                                   activebackground=ACCENT_HOVER, activeforeground="#ffffff",
                                   font=("Microsoft YaHei UI", 12, "bold"),
                                   bd=0, padx=22, pady=8, cursor="hand2",
                                   command=self.start)
        self.start_btn.pack(fill="x", pady=(14, 0))

        self.log_box = tk.Text(sec, height=7, bg="#11171f", fg="#9fc3e8",
                               insertbackground=MUTED, relief="flat",
                               font=("Consolas", 9), state="disabled")
        self.log_box.pack(fill="x", pady=(14, 0))

    # ---------- 拖放 ----------
    def set_drop_style_active(self, active):
        self.drop.configure(highlightbackground=ACCENT if active else DROP_BORDER)
        self.drop.configure(bg="#" if active else PANEL)

    def on_drop(self, event):
        self.set_drop_style_active(False)
        path = self._clean_drop_path(event.data)
        if path:
            self.file_var.set(path)
            self.drop_main_text.configure(text="已选择文件")
            self.drop_sub_text.configure(text=os.path.basename(path))
        return event.action

    @staticmethod
    def _clean_drop_path(data):
        """解析拖放数据（可能带花括号 / 多文件 / 空格路径），取第一个存在的文件。"""
        import re
        candidates = []
        # 花括号包裹的路径
        for m in re.finditer(r"\{([^{}]+)\}", data):
            candidates.append(m.group(1).strip())
        # 剩余未包裹片段
        rest = re.sub(r"\{[^{}]*\}", "", data)
        for seg in rest.split():
            seg = seg.strip()
            if seg and seg not in candidates:
                candidates.append(seg)
        for c in candidates:
            if os.path.exists(c):
                return c
        return None

    def choose(self, *_):
        p = filedialog.askopenfilename(
            title="选择教材文件（PDF / Word）",
            filetypes=[("教材文件", "*.pdf *.docx *.doc"), ("PDF", "*.pdf"),
                       ("Word", "*.docx *.doc"), ("所有文件", "*.*")],
        )
        if p:
            self.file_var.set(p)
            self.drop_main_text.configure(text="已选择文件")
            self.drop_sub_text.configure(text=os.path.basename(p))

    # ---------- 日志 ----------
    def log(self, msg):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", "[T2N] %s\n" % msg)
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        self.root.update_idletasks()

    # ---------- 执行 ----------
    def start(self):
        path = self.file_var.get().strip()
        out_root = self.out_var.get().strip() or importer.find_output_dir()
        if not path or not os.path.exists(path):
            messagebox.showwarning("提示", "请先选择或拖入教材文件")
            return
        self.start_btn.configure(state="disabled", text="导入中...")
        t = threading.Thread(target=self.worker, args=(path, out_root), daemon=True)
        t.start()

    def _find_notegen(self):
        """定位 notegen.exe：打包后从内部解压目录(_MEIPASS)取，否则取同目录。"""
        if getattr(sys, "frozen", False):
            mp = getattr(sys, "_MEIPASS", "")
            cand = os.path.join(mp, "notegen.exe")
            if cand and os.path.exists(cand):
                return cand
        return os.path.join(_base, "notegen.exe")

    def worker(self, path, out_root):
        notegen_exe = self._find_notegen()
        self.log("notegen: %s" % notegen_exe)
        ext = os.path.splitext(path)[1].lower()
        try:
            if ext == ".pdf":
                r = importer.process_pdf(path, out_root, notegen_exe)
            elif ext in (".docx", ".doc"):
                r = importer.process_word(path, out_root, notegen_exe)
            else:
                self.log("不支持的文件类型: %s" % ext)
                self._restore_btn()
                return
        except Exception as e:
            self.log("处理失败: %s" % e)
            self._restore_btn()
            messagebox.showerror("失败", str(e))
            return
        self._restore_btn()
        if r is None:
            reason, stats = importer.diagnose_pdf(path)
            if reason == "scanned":
                self.log("检测到扫描版/图片型 PDF（无文字层），无法识别目录。")
                self.log("文本层统计: 采样%d页 / 有实质文本%d页(%.1f%%) / 总字符%d / 单页文本中位数%d" %
                         (stats["sampled"], stats["text_pages"], stats["ratio"] * 100,
                          stats["text_chars"], stats["len_median"]))
                messagebox.showwarning(
                    "无法识别",
                    "该 PDF 为扫描版/图片型（无文字层），无法识别其目录。\n"
                    "请先对教材进行 OCR 处理，或更换带文字层的版本后重试。")
            else:
                self.log("检测到文字层，但前 40 页未定位到目录页（Contents/目录），无法自动识别章节层级。")
                messagebox.showwarning(
                    "无法识别",
                    "PDF 含文字层但未找到目录页（Contents / 目录）。\n"
                    "请确认教材包含目录页后重试，或更换含标准目录的版本。")
            return
        result, outdir = r
        if result.get("ok"):
            self.log("生成成功！")
            self.log("笔记本位置: %s" % outdir)
            self.log("结构: %d 分区组 / %d 分区 / %d 页" %
                     (result.get("groups", 0), result.get("sections", 0), result.get("pages", 0)))
            self.log("排序校验: %s" % ("通过" if result.get("sortedOk") else "失败"))
            messagebox.showinfo("完成", "生成成功！\n位置：%s" % outdir)
        else:
            self.log("生成失败: %s" % result.get("error"))
            messagebox.showerror("失败", result.get("error"))

    def _restore_btn(self):
        self.start_btn.configure(state="normal", text="开始导入")


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
    main()
