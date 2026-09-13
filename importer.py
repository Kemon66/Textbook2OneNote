#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Textbook2OneNote - 教材自动导入 OneNote 工具
支持 PDF / Word，自动识别目录层级，生成 .one 笔记本（OfficeIMO 引擎）。
用法:
    python importer.py [文件路径]
    无参数时弹出文件选择框。
"""

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import unicodedata

try:
    import pymupdf as fitz  # PyMuPDF >= 1.24 推荐的新导入名
except ImportError:
    import fitz


# ---------------------------------------------------------------------------
# 工具函数
# ---------------------------------------------------------------------------

def log(msg):
    try:
        print("[T2N]", msg, flush=True)
    except Exception:
        pass  # 打包为无控制台 exe 时 stdout 可能为 None


class ImportCancelled(Exception):
    """用户取消了当前操作。"""


def _emit(log_cb, msg):
    """输出一条日志：有 GUI 回调时走回调，否则打印到控制台。"""
    if log_cb is not None:
        log_cb(msg)
    else:
        log(msg)


class _TextCache:
    """按页缓存 PDF 文本与规范化 key，目录定位/删减整理时避免反复解析整页。"""

    def __init__(self, doc):
        self.doc = doc
        self._text = {}
        self._flat = {}

    def text(self, page):
        if page not in self._text:
            self._text[page] = self.doc[page].get_text() or ""
        return self._text[page]

    def flat(self, page):
        if page not in self._flat:
            self._flat[page] = _title_key(self.text(page))
        return self._flat[page]


def pick_file():
    """弹文件选择框，返回路径；用户取消则返回 None。"""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        path = filedialog.askopenfilename(
            title="选择教材文件（PDF / Word）",
            filetypes=[("教材文件", "*.pdf *.docx *.doc"), ("PDF", "*.pdf"),
                       ("Word", "*.docx *.doc"), ("所有文件", "*.*")],
        )
        root.destroy()
        return path or None
    except Exception as e:
        log("无法弹出文件框: %s" % e)
        return None


def find_output_dir():
    """默认输出目录：优先 OneDrive 桌面（实现移动端无感同步）。"""
    home = os.path.expanduser("~")
    for d in [os.path.join(home, "OneDrive", "Desktop"),
              os.path.join(home, "OneDrive", "桌面"),
              os.path.join(home, "Desktop"),
              os.path.join(home, "桌面")]:
        if os.path.isdir(d):
            return d
    return home


_WIN_ILLEGAL_NAME = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def sanitize_book_name(name):
    """清理书名：NFC 规范化、替换 Windows 非法字符、压缩空白并限制长度。"""
    s = unicodedata.normalize("NFC", str(name or ""))
    s = _WIN_ILLEGAL_NAME.sub(" ", s)
    s = re.sub(r"\s+", " ", s).strip(" .")
    return s[:120].strip() or "未命名笔记本"


def derive_book_name(path, doc=None):
    """确定笔记本（输出文件夹）名：优先文档元数据中的书名，回退到文件名。

    PDF 用内嵌元数据 title，Word 用文档属性 title；元数据缺失或是
    "Untitled / Microsoft Word - / PowerPoint" 之类占位值时回退文件名。
    """
    stem = sanitize_book_name(os.path.splitext(os.path.basename(path))[0])
    ext = os.path.splitext(path)[1].lower()
    title = ""
    if ext == ".pdf" and doc is not None:
        try:
            title = sanitize_book_name((doc.metadata or {}).get("title"))
        except Exception:
            title = ""
    elif ext == ".docx":
        try:
            import docx as _docx
            d = _docx.Document(path)
            title = sanitize_book_name(d.core_properties.title)
        except Exception:
            title = ""
    low = title.lower()
    if low.startswith("microsoft word - "):
        title = sanitize_book_name(title[len("microsoft word - "):])
    elif not title or low in ("untitled", "title", "pdf", "document",
                              "无标题", "新建"):
        title = stem
    return title or stem or "未命名笔记本"


def this_dir():
    """返回本脚本/可执行文件所在目录（用于定位 notegen.exe）。"""
    if getattr(sys, "frozen", False):
        mp = getattr(sys, "_MEIPASS", "")
        if mp and os.path.exists(os.path.join(mp, "notegen.exe")):
            return mp
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


# ---------------------------------------------------------------------------
# PDF 目录识别
# ---------------------------------------------------------------------------

def detect_toc_pages_pdf(doc):
    """返回 (start, end)：目录页范围（0-based，闭区间）。"""
    brief = None
    contents = []
    limit = min(40, doc.page_count)
    for i in range(limit):
        t = doc[i].get_text()
        # brief 目录：页脚常写作 "Brief Table of Contents"（空格可能是不换行空格\xa0）
        if re.search(r"Brief[\s\xa0]+(Table[\s\xa0]+of[\s\xa0]+)?Contents", t, re.I):
            brief = i
        # 正文目录页：英文 Contents / 中文「目 录」
        is_contents = (re.search(r"\bContents\b", t, re.I)
                       or re.search(r"目[\s\xa0]*录", t))
        if is_contents and not re.search(r"Brief[\s\xa0]+(Table[\s\xa0]+of[\s\xa0]+)?Contents", t, re.I):
            contents.append(i)
    if brief is not None:
        start = brief + 1
    elif contents:
        start = contents[0]
    else:
        return None, None
    end = contents[-1] if contents else start
    if end < start:
        end = start
    return start, min(end, doc.page_count - 1)



# 章末要素关键词白名单（正文内嵌框 FYI/CASE STUDY/IN THE NEWS 等忽略）
END_KEYWORDS = [
    "chapter in a nutshell", "key concepts", "questions for review",
    "problems and applications", "quick quiz", "summary", "glossary",
    "key terms", "review questions", "exercises", "critical thinking",
]

# 全大写标题的前缀黑名单（正文信息框，不是主节）
_NON_MAIN_CAPS_PREFIX = (
    "frontline focus", "ethical dilemma", "real world applications", "life skills",
    "thinking critically", "learning outcomes", "key terms", "for review", "summary",
    "glossary", "reference", "internet", "team exercises", "review questions",
    "review exercises", "chapter in a nutshell", "chapter review", "last word",
    "company profile", "in the news", "case study",
)

_HEADING_LINE = re.compile(r"^(PART\s|CHAPTER\s|UNIT\s|MODULE\s+([IVX\d]+)|SECTION\s|\d{1,3}\s*[>.:])", re.I)


def _looks_like_new_entry(text):
    """判断一行是否像新的目录条目起点（PART/CHAPTER/数字章节/主节/全大写主节）。
    用于大字号跨行合并循环的停止条件，避免把后续独立条目吞进上一个标题。"""
    if not text:
        return False
    if _HEADING_LINE.match(text):
        return True
    # 数字 + 空格/tab + 大写开头标题（"1 Business Now..." 式章节）
    if re.match(r"^\d{1,3}\s{1,}[A-Z0-9]", text):
        return True
    # 主节 1-1 / 1.1
    if re.match(r"^\d{1,2}[-.]\d{1,2}(?![a-zA-Z])", text):
        return True
    if _is_caps_label(text):
        return True
    return False


def _clustered_columns(page_lines):
    """按 x0 聚类分栏，返回栏内 (y,x0) 重排后的行序列副本。"""
    if not page_lines:
        return []
    xs = sorted(set(r["x0"] for r in page_lines))
    cols = []
    cur = [xs[0]]
    for x in xs[1:]:
        if x - cur[-1] > 40:
            cols.append(cur)
            cur = [x]
        else:
            cur.append(x)
    if cur:
        cols.append(cur)
    col_of = {}
    for ci, c in enumerate(cols):
        for x in c:
            col_of[x] = ci
    order = [dict(r, col=col_of[r["x0"]]) for r in page_lines]
    order.sort(key=lambda r: (r["col"], r["y"], r["x0"]))
    return order


def _layout_lines(doc, start, end):
    """版面感知目录行：(p, col, y, x0, fs, text)，跨行标题可按列内相邻 y 合并。"""
    raw = []
    for p in range(start, end + 1):
        d = doc[p].get_text("dict")
        for b in d.get("blocks", []):
            for line in b.get("lines", []):
                spans = [s for s in line.get("spans", []) if s["text"].strip()]
                if not spans:
                    continue
                txt = "".join(s["text"] for s in spans).strip()
                fs = max(s["size"] for s in spans)
                x0 = min(s["bbox"][0] for s in spans)
                y = min(s["bbox"][1] for s in spans)
                if re.fullmatch(r"[ivxlcdmIVXLCDM\s.,•]+", txt):
                    continue  # 页眉/页码
                if txt.startswith(("Table of Contents", "Brief Table", "©", "BRIEF")):
                    continue
                if re.match(r"^[vixlcIVXLC\d]+\s*•?\s*$", txt):
                    continue
                # 页眉/页脚/图片来源等噪声行
                low = txt.lower()
                if low == "contents" or low.startswith("copyright") or "editorial review has deemed" in low:
                    continue
                if low == "table of contents":
                    continue  # 目录页页眉
                if re.search(r"(shutterstock|istock|gettyimages|alamy|dreamstime)\b", low):
                    continue
                raw.append({"p": p, "fs": round(fs, 1), "x0": round(x0),
                            "y": round(y), "text": txt})
    out = []
    for p in range(start, end + 1):
        out.extend(_clustered_columns([r for r in raw if r["p"] == p]))
    return out


def _is_caps_label(text):
    """全大写且编号内无小写 -> 可能是主节标题（BECSR 风格全大写无编号）。"""
    letters = re.sub(r"[^A-Za-z]", "", text)
    return bool(letters) and letters.isupper() and len(letters) >= 5 and not re.search(r"[a-z]", text)


def _page_num_at(tail):
    m = re.search(r"(\d{1,3})\s*$", tail)
    return int(m.group(1)) if m else None


# 单页「有实质文本」的最小字符数：低于此视为无正文的空白/纯图页
_TEXT_PAGE_MIN = 3


def _sample_text_stats(doc, total, max_points=60):
    """均匀采样透视全文，返回 (有实质文本的单页长度列表, 采样总字符数, 全部采样长度列表)。
    采样点按间隔 max(1, total//max_points) 覆盖全文，最多 max_points 个，避免只看开篇几页。"""
    if total <= 0:
        return [], 0, []
    if total <= max_points:
        idxs = list(range(total))
    else:
        step = max(1, total // max_points)
        idxs = list(range(0, total, step))[:max_points]
    lens = []
    for i in idxs:
        lens.append(len(doc[i].get_text().strip()))
    text_lens = [L for L in lens if L >= _TEXT_PAGE_MIN]
    return text_lens, sum(lens), lens


def analyze_text_layer(doc):
    """分析 PDF 文本层（扫描版/图片型判定），返回统计 dict。

    判定信号：
      1) 内嵌 outline（get_toc）非空 → 文档带可检索大纲，必然有文本层；
      2) 否则均匀采样透视全文，统计「有实质文本页」占比与单页文本长度分布：
         扫描版（每页一张图、仅残留极少量 OCR 碎片）的典型特征是
         文本页占比极低（< 0.3），或占比中等（< 0.6）但文本页中位数极短（< 20 字符）。
         —— 避免「前几页任一页文本≥3字符就误判为有文本层」的粗判。
    """
    toc = doc.get_toc()
    toc_count = len(toc) if toc else 0
    total = doc.page_count
    text_lens, text_chars, all_lens = _sample_text_stats(doc, total)
    sampled = len(all_lens)
    text_pages = len(text_lens)
    ratio = round(text_pages / sampled, 3) if sampled else 0.0
    len_median = sorted(text_lens)[len(text_lens) // 2] if text_lens else 0
    len_max = max(text_lens) if text_lens else 0

    is_scanned = True
    if toc_count > 0:
        is_scanned = False          # 带大纲 → 必然有文本层
    elif text_pages == 0:
        is_scanned = True           # 采样页中没有一个实质文本页
    elif ratio >= 0.6:
        is_scanned = False          # 绝大多数采样页都有实质文本 → 正常文本教材
    elif ratio < 0.3:
        is_scanned = True           # 文本页占比过低 → 每页一图、零星残片
    else:
        is_scanned = len_median < 20  # 占比中等但单页文本极短 → OCR 残片型扫描件

    return {
        "toc_count": toc_count,
        "total_pages": total,
        "sampled": sampled,
        "text_pages": text_pages,
        "ratio": ratio,
        "text_chars": text_chars,
        "len_median": len_median,
        "len_max": len_max,
        "is_scanned": is_scanned,
    }


def has_text_layer(path):
    """稳健判定 PDF 是否有可提取的文本层（扫描版/图片型 PDF 为 False）。
    以 get_toc 大纲 + 全文多页采样统计为准，取代旧的「前 N 页任一页≥3字符」粗判。"""
    try:
        doc = fitz.open(path)
        try:
            return not analyze_text_layer(doc)["is_scanned"]
        finally:
            doc.close()
    except Exception:
        return False


def diagnose_pdf(path):
    """目录识别失败时给出原因分类（供 GUI 弹窗使用）。
    返回 (reason, stats)：reason ∈ {'scanned', 'no_toc'}。"""
    doc = fitz.open(path)
    try:
        stats = analyze_text_layer(doc)
    finally:
        doc.close()
    reason = "scanned" if stats["is_scanned"] else "no_toc"
    return reason, stats


def parse_toc_pdf(doc, start, end):
    """版面感知通用目录解析：支持分栏排版、`N > Title` 章节、全大写无编号主节、
    PART/CHAPTER 跨行、章末要素。返回 entries。"""
    rows = _layout_lines(doc, start, end)
    if not rows:
        return _parse_toc_text(doc, start, end)

    entries = []
    pending = None  # 无页码跨行标题暂存（文本）
    i, n = 0, len(rows)
    while i < n:
        r = rows[i]
        text, fs = r["text"], r["fs"]

        # ---- 大字号行：PART/CHAPTER/UNIT 等章节级标题，允许跨行合并 ----
        # 主节行（N-M 编号 / 全大写无编号）即使字体较大也不按章节级合并，直接走普通行逻辑
        is_mainrow = bool(re.match(r"^\d{1,2}[-.]\d{1,2}(?![a-zA-Z])", text)) or _is_caps_label(text)
        if fs >= 10.0 and not is_mainrow:
            buf, j = text, i + 1
            tpage = None
            # 合并续行：同为大体字且不是新条目（行尾带页码也算标题续行，页码并入该条目）
            while (j < n and rows[j]["fs"] >= 10.0
                   and not _looks_like_new_entry(rows[j]["text"])):
                seg = rows[j]["text"]
                sp = _page_num_at(seg)
                if sp is not None:
                    tpage = sp
                    seg = seg[:seg.rfind(str(sp))].strip()
                buf += " " + seg
                j += 1
            t = re.sub(r"\s+", " ", buf).strip()
            # 行尾 1-3 位数字若落在页数范围内，视为页码；识别标题时先剥离，
            # 未知行则保留原始文本（含页码）交给普通行逻辑处理。
            sp0 = _page_num_at(t)
            t_clean = t
            if sp0 is not None and 1 <= sp0 <= doc.page_count:
                t_clean = t[: t.rfind(str(sp0))].strip()
            mm = re.match(r"^PART\s+([IVXLCivxlcd]+|\d+)\s*(.*)$", t_clean, re.I)
            if mm:
                entries.append({"level": "part", "num": mm.group(1).upper(),
                                "name": mm.group(2).strip(), "print": 0})
            else:
                cm = re.match(r"^CHAPTER\s+(\d+)\s*(.*)$", t_clean, re.I)
                if cm:
                    entries.append({"level": "chapter", "num": cm.group(1),
                                    "name": cm.group(2).strip(),
                                    "print": tpage or sp0 or 0})
                else:
                    cm2 = re.match(r"^(\d{1,3})\s*[>.:]\s*(.+)$", t_clean)
                    if cm2 and int(cm2.group(1)) <= 200:
                        entries.append({"level": "chapter", "num": cm2.group(1),
                                        "name": cm2.group(2).strip(),
                                        "print": tpage or sp0 or 0})
                    else:
                        mm2 = re.match(r"^(\d{1,3})\s{1,}([A-Z0-9][^0-9].*)$", t_clean)
                        if mm2 and int(mm2.group(1)) <= 200:
                            entries.append({"level": "chapter", "num": mm2.group(1),
                                            "name": mm2.group(2).strip(),
                                            "print": tpage or sp0 or 0})
                        else:  # 未知大字号行，按普通行继续处理（保留 pending，不重置）
                            fs = 8.0
                            text = t
            if fs >= 10.0:
                pending = None
                i = j
                continue

        # ---- 普通行：尝试剥离行尾页码 ----
        if text is None:
            text = r["text"]
        pg = _page_num_at(text)
        page = pg if (pg is not None and 1 <= pg <= doc.page_count) else None
        body = text[:text.rfind(str(pg))].strip() if page is not None else text

        # 若 pending 仍挂着上一行完整条目，而本行又是新的条目起点（主节/PART/CHAPTER/全大写），
        # 先把 pending 落盘（作为无页码条目），再让本行独立处理，避免把两条拼成一条。
        if pending is not None and _looks_like_new_entry(body):
            cl_old = _classify(pending, allow_caps=True)
            if cl_old is not None and cl_old[0] == "main":
                entries.append({"level": "main", "num": cl_old[1],
                                "name": cl_old[2], "print": None})
            elif cl_old is not None and cl_old[0] in ("part", "chapter"):
                entries.append({"level": cl_old[0], "num": cl_old[1],
                                "name": cl_old[2], "print": 0})
            pending = None

        # 优先级1：跨行标题合并（pending + 本行续行）
        if pending is not None and not _looks_like_new_entry(body):
            cand = (pending + " " + body).strip()
            if page is not None:
                cnum = re.match(r"^(\d{1,2})[-.](\d{1,2})(?![a-zA-Z])\s*(.*)$", cand)
                cl = _classify(cand, allow_caps=True)
                if cnum:
                    entries.append({"level": "main", "num": cnum.group(1) + "-" + cnum.group(2),
                                    "name": cnum.group(3).strip(), "print": page})
                elif cl is not None and cl[0] == "main":
                    entries.append({"level": "main", "num": cl[1], "name": cl[2], "print": page})
                elif any(kw in cand.lower() for kw in END_KEYWORDS):
                    entries.append({"level": "end", "num": "", "name": cand, "print": page})
                # 有页码但合并结果不像标题 -> 整行丢弃（子节/正文行）
                pending = None
            else:
                # 本行仍无页码：可继续累积跨行标题
                pending = cand
            i += 1
            continue

        # 优先级2：无 pending，行内直接判定
        if page is not None:
            cnum = re.match(r"^(\d{1,2})[-.](\d{1,2})(?![a-zA-Z])\s*(.*)$", body)
            if cnum:
                entries.append({"level": "main", "num": cnum.group(1) + "-" + cnum.group(2),
                                "name": cnum.group(3).strip(), "print": page})
                i += 1
                continue
            cl = _classify(body, allow_caps=True)
            if cl is not None and cl[0] in ("main", "part", "chapter"):
                if cl[0] == "main" or (cl[0] == "chapter" and body.lower().startswith(("chapter ", "part ", "unit "))):
                    pass  # part/chapter 无页码时也记录（build_plan 可处理 print=0）
                entries.append({"level": cl[0], "num": cl[1], "name": cl[2], "print": page})
                i += 1
                continue
            if any(kw in body.lower() for kw in END_KEYWORDS):
                entries.append({"level": "end", "num": "", "name": body, "print": page})
                i += 1
                continue
        else:
            # 无页码：优先识别章节级标题（数字+标题 / 数字>标题 / PART/CHAPTER / 全大写主节）
            lc = body.lower()
            is_ch_like = bool(re.match(r"^\d{1,3}[\s>]", body)) or lc.startswith(("part ", "chapter ", "unit "))
            is_num_main = bool(re.match(r"^\d{1,2}[-.]\d{1,2}(?![a-zA-Z])", body))
            if is_ch_like or _is_caps_label(body) or is_num_main:
                # 新条目行（带编号/全大写）：把上一个 pending 的跨行条目先落盘
                if pending is not None:
                    cl_old = _classify(pending, allow_caps=True)
                    if cl_old is not None and cl_old[0] == "main":
                        entries.append({"level": "main", "num": cl_old[1],
                                        "name": cl_old[2], "print": None})
                    elif cl_old is not None and cl_old[0] in ("part", "chapter"):
                        entries.append({"level": cl_old[0], "num": cl_old[1],
                                        "name": cl_old[2], "print": 0})
                    pending = None
                cand = body
                cl = _classify(cand, allow_caps=True)
                if cl is not None and cl[0] in ("part", "chapter"):
                    entries.append({"level": cl[0], "num": cl[1], "name": cl[2], "print": 0})
                elif cl is not None and cl[0] == "main":
                    # 无页码主节：先暂存为 pending，等续行页码；若下一行又是新条目则在此落盘
                    pending = cand
                elif cl is None and is_ch_like:
                    dm = re.match(r"^(\d{1,3})\s+([A-Z0-9][A-Za-z0-9 :'%&-]+)$", cand)
                    if dm and int(dm.group(1)) <= 200:
                        entries.append({"level": "chapter", "num": dm.group(1),
                                        "name": dm.group(2).strip(), "print": 0})
                    else:
                        pending = cand
                else:
                    pending = (pending + " " + cand).strip() if pending else cand
            else:
                # 非新条目行（跨行续行/正文行）：若存在 pending 则拼接续行
                if pending is not None:
                    pending = pending + " " + body
        i += 1

    # 循环结束：把挂着的最后一个 pending 落盘
    if pending is not None:
        cl = _classify(pending, allow_caps=True)
        if cl is not None and cl[0] == "main":
            entries.append({"level": "main", "num": cl[1], "name": cl[2], "print": None})
        elif cl is not None and cl[0] in ("part", "chapter"):
            entries.append({"level": cl[0], "num": cl[1], "name": cl[2], "print": 0})

    return entries


def _classify(body, allow_caps=False):
    """识别条目类型，返回 (level, num, name) 或 None。
    allow_caps=True 时识别全大写无编号主节（BECSR 风格）。"""
    mm = re.match(r"^PART\s+([IVXLCivxlcd]+|\d+)\s*(.*)$", body, re.I)
    if mm:
        return "part", mm.group(1).upper(), mm.group(2).strip()
    mm = re.match(r"^CHAPTER\s+(\d+)\s*(.*)$", body, re.I)
    if mm:
        return "chapter", mm.group(1), mm.group(2).strip()
    # 中文章节: 第1章 / 第 1 章 / 第1节 / 第一〇章
    mm = re.match(r"^第\s*(\d+)\s*[章节]\s*(.*)$", body)
    if mm:
        return "chapter", mm.group(1), mm.group(2).strip()
    # 主节: 1-1 / 1.1 (无字母后缀)；子节 1-1a 忽略
    mm = re.match(r"^(\d{1,2})[-.](\d{1,2})(?![a-zA-Z])\s*(.*)$", body)
    if mm:
        return "main", mm.group(1) + "-" + mm.group(2), mm.group(3).strip()
    # 数字 > 标题（BECSR 商业教材章节风格）
    mm = re.match(r"^(\d{1,3})\s*>\s*(.+)$", body)
    if mm and int(mm.group(1)) <= 200:
        return "chapter", mm.group(1), mm.group(2).strip()
    # 全大写无编号主节（BECSR 主节风格）
    if allow_caps and _is_caps_label(body) and not body.lower().startswith(_NON_MAIN_CAPS_PREFIX):
        return "main", "", body
    return None


def _parse_toc_text(doc, start, end):
    """旧版逐行文本解析（表格/复杂版面布局解析失败时回退）。"""
    raw = []
    for i in range(start, end + 1):
        raw.extend(doc[i].get_text().split("\n"))
    entries = []
    pending = None
    for ln in raw:
        s = ln.strip()
        if not s:
            continue
        if re.fullmatch(r"[ivxlcdm]+", s, re.I):
            continue
        if re.match(r"^\d{1,2}[-.]\d{1,2}[a-z]", s):
            continue
        cm = re.match(r"^CHAPTER\s+(\d+)\s*$", s, re.I)
        if cm:
            pending = ("chapter", cm.group(1), "")
            continue
        m = re.search(r"(\d+)\s*$", s)
        page = int(m.group(1)) if m else None
        body = s[:m.start()].strip() if m else s
        c = _classify(body, allow_caps=False)
        if c is not None:
            level, num, name = c
            if page is not None and (level != "part" or name):
                entries.append({"level": level, "num": num, "name": name, "print": page})
                pending = None
            else:
                pending = (level, num, name)
        elif pending is not None:
            pl, pn, pname = pending
            full = (pname + " " + body).strip() if pname else body
            if page is not None:
                entries.append({"level": pl, "num": pn, "name": full, "print": page})
                pending = None
            else:
                pending = (pl, pn, full)
        else:
            if page is not None and body:
                for kw in END_KEYWORDS:
                    if kw in body.lower():
                        entries.append({"level": "end", "num": "", "name": body, "print": page})
                        break
    return entries


def calc_offset_pdf(doc, entries, toc_end, cancel=None):
    """锚定法 + 众数：用主节/章节标题在正文中定位，排除离群值求最稳 offset。"""
    scored = [(e, e["print"]) for e in entries
              if e["level"] in ("main", "chapter") and e["print"] is not None
              and e["name"] and len(e["name"].split()) >= 4]
    cache = _TextCache(doc)
    offsets = []
    for e, pr in scored[:60]:
        if cancel is not None and cancel.is_set():
            raise ImportCancelled()
        lo = max(toc_end + 1, pr - 1)
        # offset 只接受 0..160，因此只需在 pr-1..pr+160 窗口内搜索，避免扫到书尾
        hi = min(doc.page_count, pr + 161)
        abs_p = _locate_title_page(doc, e["name"], lo, hi, cache=cache, cancel=cancel)
        if abs_p is None:
            continue
        off = abs_p - pr
        if 0 <= off <= 160:
            offsets.append(off)
    if not offsets:
        return None
    from collections import Counter
    cnt = Counter(offsets)
    top, topn = cnt.most_common(1)[0]
    if topn >= max(2, len(offsets) * 0.35):
        return top
    offsets.sort()
    return offsets[len(offsets) // 2]


_STOP = set("a an the of and or for in on to with from by at as it is are was we you your its his her our their this that these those".split())


def _title_key(title):
    """标题规范化（忽略大小写/空白，保留字母数字）。"""
    return re.sub(r"[^a-z0-9%$]+", "", title.lower())


def _locate_title_page(doc, title, lo, hi, cache=None, cancel=None):
    """在正文 [lo,hi) 页范围内定位完整标题首次出现的绝对页（忽略空白/大小写）。
    标题可能跨行，故用压缩空白后的全文匹配。找不到返回 None。"""
    key = _title_key(title)
    if len(key) < 12:
        return None
    cache = cache or _TextCache(doc)
    for i in range(max(0, lo), min(hi, doc.page_count)):
        if cancel is not None and cancel.is_set():
            raise ImportCancelled()
        if key in cache.flat(i):
            return i
    return None


def fill_missing_pages(doc, entries, toc_end, offset, cancel=None):
    """对 print 为 None 的 main，用正文标题搜索回填其印刷页码。
    找不到的按上一主节+1 近似，仍覆盖不到则用 chapter 起始。"""
    cache = _TextCache(doc)
    last_abs = None          # 上一主节已定位的正文绝对页，用于推进搜索下界
    for e in entries:
        if cancel is not None and cancel.is_set():
            raise ImportCancelled()
        if e["level"] == "chapter":
            # 章节有页码时顺带锚定 last_abs，保证主节搜索从章节正文页起步
            if e["print"] is not None:
                lo = max(toc_end + 1, e["print"] - 1 if last_abs is None else last_abs - 1)
                ap = _locate_title_page(doc, e["name"], lo, doc.page_count,
                                        cache=cache, cancel=cancel)
                if ap is not None:
                    last_abs = ap
            continue
        if e["level"] != "main":
            continue
        if e["print"] is not None and last_abs is None:
            lo = max(toc_end + 1, e["print"] - 1)
            ap = _locate_title_page(doc, e["name"], lo, doc.page_count,
                                    cache=cache, cancel=cancel)
            if ap is not None:
                last_abs = ap
            continue
        if e["print"] is not None:
            continue
        lo = max(toc_end + 1, (last_abs + 1) if last_abs is not None else toc_end + 1)
        ap = _locate_title_page(doc, e["name"], lo, doc.page_count,
                                cache=cache, cancel=cancel)
        if ap is not None:
            e["print"] = ap - offset if offset is not None else ap
            last_abs = ap
        elif last_abs is not None and e["print"] is None:
            e["print"] = last_abs + 1 - offset if offset is not None else last_abs + 1
            last_abs += 1
    # 依然为 None 的 main：按章节内前序+1 兜底
    prev = None
    for e in entries:
        if e["level"] == "chapter":
            prev = e["print"] if e["print"] is not None else prev
        elif e["level"] == "main" and e["print"] is None:
            e["print"] = (prev + 1) if prev is not None else (toc_end + 1)
            prev = e["print"]
        elif e["level"] == "main":
            prev = e["print"]
    return entries


# ---------------------------------------------------------------------------
# Word 目录识别
# ---------------------------------------------------------------------------

def parse_docx(path):
    """用 Heading 1/2/3 识别层级，返回 (entries, texts_by_heading)。"""
    import docx
    d = docx.Document(path)
    entries = []
    texts = {}  # 每个 main 的正文段落列表
    cur_part = cur_chapter = cur_main = None
    order = []  # 记录 main 出现的顺序
    h1 = h2 = h3 = 0
    for p in d.paragraphs:
        style = (p.style.name or "").lower() if p.style else ""
        txt = p.text.strip()
        if not txt:
            continue
        if "heading 1" in style or style == "title":
            h1 += 1
            cur_part = "PART %d" % h1
            cur_chapter = cur_main = None
            entries.append({"level": "part", "num": str(h1), "name": txt, "print": 0})
        elif "heading 2" in style:
            h2 += 1
            cur_chapter = "Chapter %d" % h2
            cur_main = None
            entries.append({"level": "chapter", "num": str(h2), "name": txt, "print": 0})
        elif "heading 3" in style:
            h3 += 1
            cur_main = "%d-%d" % (h2, h3)
            entries.append({"level": "main", "num": cur_main, "name": txt, "print": 0})
            texts[cur_main] = []
            order.append(cur_main)
        else:
            if cur_main is not None and len(txt) > 1:
                texts[cur_main].append(txt)
    return entries, texts


# ---------------------------------------------------------------------------
# plan 构建
# ---------------------------------------------------------------------------

_ROMAN = {'I': 1, 'II': 2, 'III': 3, 'IV': 4, 'V': 5, 'VI': 6, 'VII': 7, 'VIII': 8,
          'IX': 9, 'X': 10, 'XI': 11, 'XII': 12, 'XIII': 13, 'XIV': 14, 'XV': 15,
          'XVI': 16, 'XVII': 17, 'XVIII': 18, 'XIX': 19, 'XX': 20}


def _pad_num(num, width=2):
    """把 '1-1' / '1' / 'I' 规范化为零填充、字典序正确的编号。"""
    s = (num or "").strip()
    if not s:
        return s
    u = s.upper()
    if u in _ROMAN:
        return "%0*d" % (width, _ROMAN[u])
    parts = re.split(r"[-.]", s)
    try:
        return "-".join("%0*d" % (width, int(p)) for p in parts if p != "")
    except ValueError:
        return s


def build_plan(entries, offset):
    """把 entries 组装成 plan（parts/chapters/mains/abs_pages/ends）。"""
    parts = []
    cur_part = None
    cur_ch = None
    part_seq = 0
    for e in entries:
        if e["level"] == "part":
            part_seq += 1
            cur_part = {"num": "%02d" % part_seq, "name": e["name"], "chapters": []}
            parts.append(cur_part)
            cur_ch = None
        elif e["level"] == "chapter":
            cur_ch = {"num": _pad_num(e["num"]), "name": e["name"], "mains": [],
                      "ends_print_pages": [], "ends_abs_pages": []}
            if cur_part is None:
                cur_part = {"num": "0", "name": "Contents", "chapters": []}
                parts.append(cur_part)
            cur_part["chapters"].append(cur_ch)
        elif e["level"] == "main":
            if cur_ch is None:
                cur_ch = {"num": "0", "name": "", "mains": [],
                          "ends_print_pages": [], "ends_abs_pages": []}
                if cur_part is None:
                    cur_part = {"num": "0", "name": "Contents", "chapters": []}
                    parts.append(cur_part)
                cur_part["chapters"].append(cur_ch)
            cur_ch["mains"].append({"num": _pad_num(e["num"]), "name": e["name"], "print_start": e["print"]})
        elif e["level"] == "end":
            if cur_ch is not None:
                cur_ch["ends_print_pages"].append(e["print"])

    # 计算 abs 页范围
    for part in parts:
        for ch in part["chapters"]:
            mains = ch["mains"]
            for i, m in enumerate(mains):
                sp = m["print_start"]
                if i + 1 < len(mains):
                    ep = max(sp, mains[i + 1]["print_start"] - 1)  # 同页相邻主节也至少占自身一页
                elif ch["ends_print_pages"]:
                    ep = max(sp, min(ch["ends_print_pages"]) - 1)
                else:
                    ep = sp
                m["abs_pages"] = list(range(sp + offset, max(sp + offset, ep + offset + 1)))
            ch["ends_abs_pages"] = [p + offset for p in ch["ends_print_pages"]]
            ch.pop("ends_print_pages", None)
            for m in ch["mains"]:
                m.pop("print_start", None)
    return {"parts": parts}


def build_plan_word(entries, texts):
    """Word 路径：无页码，用 texts 建文本页。"""
    parts = []
    cur_part = cur_ch = None
    for e in entries:
        if e["level"] == "part":
            cur_part = {"num": e["num"], "name": e["name"], "chapters": []}
            parts.append(cur_part)
            cur_ch = None
        elif e["level"] == "chapter":
            cur_ch = {"num": e["num"], "name": e["name"], "mains": []}
            if cur_part is None:
                cur_part = {"num": "0", "name": "Contents", "chapters": []}
                parts.append(cur_part)
            cur_part["chapters"].append(cur_ch)
        elif e["level"] == "main":
            if cur_ch is None:
                cur_ch = {"num": "0", "name": "", "mains": []}
                if cur_part is None:
                    cur_part = {"num": "0", "name": "Contents", "chapters": []}
                    parts.append(cur_part)
                cur_part["chapters"].append(cur_ch)
            cur_ch["mains"].append({"num": _pad_num(e["num"]), "name": e["name"],
                                    "texts": texts.get(e["num"], [])[:200]})
    return {"parts": parts}


# ---------------------------------------------------------------------------
# 删减内容整理（未被目录覆盖的 PDF 页自动归类建分区）
# ---------------------------------------------------------------------------

LO_PART_NAME = "其他内容"          # 删减内容所在分区组名
LO_CAT_NAMES = {
    "preface": "前言 Preface",
    "toc": "目录 Contents",
    "appendix": "附录 Appendix",
    "endnote": "尾注 Endnotes",
    "index": "索引 Index",
    "glossary": "术语表 Glossary",
    "back": "后记/结语 Afterword",
    "other": "其他内容 Other",
}

_LO_KW = {
    # 每类若干关键词，命中一次 +1 分（不区分大小写）
    "preface": ("preface", "foreword", "introduction", "acknowledg",
                "前言", "序言", "序文", "致谢", "序 "),
    "appendix": ("appendix", "附 录", "附录"),
    "back": ("afterword", "epilogue", "conclusion", "后记", "结语",
             "总 结", "结论", "尾声"),
}
# 书尾里程碑标题：单独成行/行首出现即视为一个独立部分的起始页
_LO_MILESTONES = [
    ("endnote",  ("endnote", "尾注", "参考文献", "references", "bibliography", "works cited")),
    ("glossary", ("glossary", "术语表", "词汇表")),
    ("index",    ("index", "索 引", "索引")),
    ("appendix", ("append", "附录", "附 录", "apéndice")),
    ("back",     ("afterword", "epilogue", "后记", "结语", "结 语", "尾声")),
]

# 尾注/术语/索引等段的里程碑词：命中"正文正文 Index"之类句子时要求独立成行才认可
_M_LINE_WORDS = {"index": "index", "glossary": "glossary", "endnote": "endnote",
                 "references": "references", "bibliography": "bibliography",
                 "appendix": "appendix"}


def _consecutive(pages):
    """把已排序页号切成连续区间列表。"""
    pages = sorted(set(pages))
    out, cur = [], []
    for p in pages:
        if cur and p == cur[-1] + 1:
            cur.append(p)
        else:
            if cur:
                out.append(cur)
            cur = [p]
    if cur:
        out.append(cur)
    return out


def _page_lines(doc, p, n=6, cache=None):
    """取页面文本的前 n 个非空行使于里程碑检测。"""
    txt = (cache.text(p) if cache is not None else (doc[p].get_text() or "")).strip()
    lines = [l.strip() for l in txt.splitlines() if l.strip()]
    return lines[:n]


def _milestone_page_kind(doc, p, cache=None):
    """书尾独立部分起始页探测：前几行命中独立标题词 -> (kind)或 None。"""
    for line in _page_lines(doc, p, 6, cache=cache):
        low = line.lower()
        for kind, kws in _LO_MILESTONES:
            for kw in kws:
                if low.rstrip(" \t:.").endswith(kw) or low.startswith(kw):
                    # 要求该词独立成行或行首短前缀，排除正文句子误命中
                    for word in _M_LINE_WORDS:
                        if kw.startswith(word):
                            if re.search(r"\b%s\b" % re.escape(word), low) and \
                               not re.match(r"^.*?\b%s\b.*?\b%s\b" % (word, word), low):
                                return kind
                    return kind
    return None


def _region_kind(doc, pages, toc_start, toc_end, cache=None):
    """对一段删减页整体归类：逐页关键词加权后按众数取 kind。
    避免单页中"Brief Contents 列出附录条目"之类的噪声词主导整段。
    """
    toc_cnt = sum(1 for p in pages if toc_start <= p <= toc_end)
    if toc_cnt >= max(1, len(pages) * 0.4):
        return "toc"
    votes = {}
    for p in pages:
        text = (cache.text(p) if cache is not None else (doc[p].get_text() or "")).lower()
        scores = {}
        for kind, kws in _LO_KW.items():
            s = sum(text.count(k) for k in kws)
            scores[kind] = scores.get(kind, 0) + s
        # 目录页强归 toc
        if toc_start <= p <= toc_end:
            scores["toc"] = max(scores.get("toc", 0), 1)
        if not any(scores.values()):
            best = "other"
        else:
            best = max(scores, key=scores.get)
        votes[best] = votes.get(best, 0) + 1
    return max(votes, key=votes.get) if votes else "other"


def classify_leftovers(doc, toc_start, toc_end, covered, cache=None, cancel=None):
    """识别未被目录覆盖的 PDF 页并按语义归类为删减分区。

    只处理两类明确区域，避免正文页因页码偏移被误当删减内容：
      · 书首区：绝对页 < 首个被覆盖页（封面/版权/前言/目录等）
      · 书尾区：绝对页 > 最后被覆盖页（附录/尾注/术语表/索引/后记等）
    正文中间散布的未覆盖页属"目录覆盖缺口"，不纳入删减分区（单独计数报告）。
    返回 [{"kind": str, "pages": [abs...]}]，每 kind 只会出现一次（跨区间合并）。
    """
    cache = cache or _TextCache(doc)
    total = doc.page_count
    uncov = sorted(set(range(total)) - set(covered))
    first_cov = min(covered) if covered else total
    last_cov = max(covered) if covered else total
    head_pages = [i for i in uncov if i < first_cov]
    tail_pages = [i for i in uncov if i > last_cov]
    # 正文中间漏覆盖页（不计入删减分区）
    n_mid = len(uncov) - len(head_pages) - len(tail_pages)

    groups = []                       # [(kind, pages)]
    # ---- 书首区：目录区间优先，其余按连续片断逐段归类 ----
    toc_set = set()
    for rg in _consecutive(head_pages):
        if cancel is not None and cancel.is_set():
            raise ImportCancelled()
        sub_toc = [p for p in rg if toc_start <= p <= toc_end]
        if sub_toc:
            groups.append(("toc", sub_toc))
        sub_rest = [p for p in rg if p not in sub_toc]
        for seg in _consecutive(sub_rest):
            groups.append((_region_kind(doc, seg, toc_start, toc_end, cache=cache), seg))
    # ---- 书尾区：里程碑切段，无里程碑段按关键词归并 ----
    for rg in _consecutive(tail_pages):
        seg_kind, seg_pages, pending = None, [], []
        for p in rg:
            if cancel is not None and cancel.is_set():
                raise ImportCancelled()
            mk = _milestone_page_kind(doc, p, cache=cache)
            if mk is not None:
                if pending:                       # 里程碑前积累的无标签页
                    groups.append((_region_kind(doc, pending, toc_start, toc_end, cache=cache), pending))
                    pending = []
                if seg_kind is not None and seg_kind != mk:
                    groups.append((seg_kind, seg_pages))
                    seg_pages = []
                seg_kind = mk
                seg_pages.append(p)
            elif seg_kind is not None:
                seg_pages.append(p)
            else:
                pending.append(p)
        if pending:
            groups.append((_region_kind(doc, pending, toc_start, toc_end, cache=cache), pending))
        if seg_kind is not None:
            groups.append((seg_kind, seg_pages))
    # ---- 同 kind 合并：仅页号仍连续才合并，避免书首/书尾不相干片段被硬拼 ----
    merged = []
    for kind, ps in groups:
        ps = sorted(set(ps))
        if merged and merged[-1]["kind"] == kind:
            cand = sorted(set(merged[-1]["pages"] + ps))
            if cand[-1] - cand[0] + 1 == len(cand):
                merged[-1]["pages"] = cand
                continue
        merged.append({"kind": kind, "pages": ps})
    return merged, len(merged), n_mid


def add_leftover_part(plan, leftovers):
    """把删减内容分区追加到 plan（保持 notegen 认识的 parts/chapters/mains 结构）。

    同 kind 的多个不相连片段合并为一个分区（分区内多个 main 片段）；
    返回 (chapters, all_pages)，all_pages 为此前未被渲染的删减绝对页集合。
    """
    if not leftovers:
        return [], []
    by_kind = {}
    for g in leftovers:
        by_kind.setdefault(g["kind"], []).append(g["pages"])
    chapters, all_pages = [], []
    for idx, (kind, page_groups) in enumerate(by_kind.items(), 1):
        name = LO_CAT_NAMES.get(kind, LO_CAT_NAMES["other"])
        mains = []
        for j, ps in enumerate(page_groups, 1):
            mname = name if len(page_groups) == 1 else "%s (%d)" % (name, j)
            mains.append({"num": "1-%d" % j, "name": mname, "abs_pages": ps})
            all_pages.extend(ps)
        chapters.append({"num": "%02d" % idx, "name": name, "mains": mains})
    part = {"num": "%02d" % (len(plan["parts"]) + 1),
            "name": LO_PART_NAME,
            "chapters": chapters}
    plan["parts"].append(part)
    return chapters, all_pages


# ---------------------------------------------------------------------------
# 渲染 & 生成
# ---------------------------------------------------------------------------

def render_pdf(doc, abs_pages, outdir, cancel=None, progress=None):
    os.makedirs(outdir, exist_ok=True)
    pages = sorted(abs_pages)
    total = len(pages)
    if progress:
        progress("渲染页面", 0, total)
    n = 0
    for idx, ap in enumerate(pages, 1):
        if cancel is not None and cancel.is_set():
            raise ImportCancelled()
        png = os.path.join(outdir, "page_%04d.png" % ap)
        if os.path.exists(png):
            continue
        doc[ap].get_pixmap(dpi=150).save(png)
        n += 1
        if progress:
            progress("渲染页面", idx, total)
    return n


def run_notegen(exe, plan_path, rendered, outdir, name):
    if not os.path.isfile(exe):
        raise FileNotFoundError("未找到 OneNote 生成引擎 notegen.exe：%s" % exe)
    cmd = [exe, plan_path, rendered, outdir, name]
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                       errors="replace", timeout=1800)
    for line in r.stdout.splitlines():
        line = line.strip()
        if line.startswith("RESULT "):
            try:
                return json.loads(line[len("RESULT "):])
            except Exception:
                return {"ok": False, "error": "parse result failed: " + line}
    return {"ok": False, "error": "no RESULT from notegen; stdout=" + r.stdout[-500:] + " stderr=" + r.stderr[-500:]}


# ---------------------------------------------------------------------------
# 分析 / 预览（不渲染、不生成 OneNote 文件）
# ---------------------------------------------------------------------------

def _collect_abs_pages(plan):
    """汇总 plan 中全部待渲染的绝对页号。"""
    all_abs = set()
    for part in plan.get("parts", []):
        for ch in part.get("chapters", []):
            for m in ch.get("mains", []):
                all_abs.update(m.get("abs_pages") or [])
            all_abs.update(ch.get("ends_abs_pages") or [])
    return all_abs


def _build_preview_tree(entries, counts=None):
    """把 entries 转成 GUI 预览用的嵌套树；counts 为 Word 主节段落数映射。"""
    counts = counts or {}
    tree = []
    cur_part = cur_ch = None
    for e in entries:
        lv = e["level"]
        if lv == "part":
            cur_part = {"level": "part", "num": e.get("num", ""), "name": e.get("name", ""),
                        "print": e.get("print"), "children": []}
            tree.append(cur_part)
            cur_ch = None
        elif lv == "chapter":
            if cur_part is None:
                cur_part = {"level": "part", "num": "0", "name": "Contents",
                            "print": None, "children": []}
                tree.append(cur_part)
            cur_ch = {"level": "chapter", "num": e.get("num", ""), "name": e.get("name", ""),
                      "print": e.get("print"), "children": []}
            cur_part["children"].append(cur_ch)
        elif lv in ("main", "end"):
            if cur_ch is None:
                if cur_part is None:
                    cur_part = {"level": "part", "num": "0", "name": "Contents",
                                "print": None, "children": []}
                    tree.append(cur_part)
                cur_ch = {"level": "chapter", "num": "0", "name": "",
                          "print": None, "children": []}
                cur_part["children"].append(cur_ch)
            node = {"level": lv, "num": e.get("num", ""), "name": e.get("name", ""),
                    "print": e.get("print")}
            if lv == "main" and counts:
                node["count"] = min(len(counts.get(e.get("num", ""), [])), 200)
            cur_ch["children"].append(node)
    return tree


def analyze_file(path, cancel=None, progress=None, log_cb=None):
    """预览/预分析入口：识别目录层级并返回结构树与生成计划，不渲染不写 OneNote。

    成功返回 ok=True 并附带 plan；失败返回 ok=False 且带 reason：
    scanned / no_toc / empty_toc（PDF），no_headings / doc_legacy（Word），
    unsupported（其他扩展名）。
    """
    path = os.path.abspath(path)
    name = sanitize_book_name(os.path.splitext(os.path.basename(path))[0])
    size = os.path.getsize(path) if os.path.exists(path) else 0
    ext = os.path.splitext(path)[1].lower()

    def _tick(stage, done=0, total=0):
        if progress:
            progress(stage, done, total)

    if ext == ".pdf":
        doc = fitz.open(path)
        try:
            name = derive_book_name(path, doc=doc)
            _tick("检测目录页")
            start, end = detect_toc_pages_pdf(doc)
            stats = analyze_text_layer(doc)
            if start is None:
                reason = "scanned" if stats["is_scanned"] else "no_toc"
                return {"ok": False, "kind": "pdf", "reason": reason, "name": name,
                        "path": path, "size": size, "pages": stats["total_pages"],
                        "stats": stats, "toc_pages": None}

            _tick("解析目录条目")
            entries = parse_toc_pdf(doc, start, end)
            n_part = sum(1 for e in entries if e["level"] == "part")
            n_ch = sum(1 for e in entries if e["level"] == "chapter")
            n_main = sum(1 for e in entries if e["level"] == "main")
            n_end = sum(1 for e in entries if e["level"] == "end")
            if n_ch == 0 and n_main == 0:
                return {"ok": False, "kind": "pdf", "reason": "empty_toc", "name": name,
                        "path": path, "size": size, "pages": stats["total_pages"],
                        "stats": stats, "toc_pages": [start, end]}

            _tick("计算页码偏移")
            offset = calc_offset_pdf(doc, entries, end, cancel=cancel)
            if offset is None:
                _emit(log_cb, "偏移计算失败，回退 0")
                offset = 0
            _emit(log_cb, "页码偏移 offset=%d (abs = print + offset)" % offset)

            n_missing = sum(1 for e in entries if e["level"] == "main" and e["print"] is None)
            if n_missing:
                _emit(log_cb, "正文定位回填 %d 个无页码主节..." % n_missing)
                fill_missing_pages(doc, entries, end, offset, cancel=cancel)
                still = sum(1 for e in entries if e["level"] == "main" and e["print"] is None)
                _emit(log_cb, "回填后仍缺 %d" % still)

            _tick("生成目录结构")
            plan = build_plan(entries, offset)
            all_abs = _collect_abs_pages(plan)

            _tick("整理删减内容")
            leftovers, n_lo, n_mid = classify_leftovers(doc, start, end, all_abs,
                                                        cancel=cancel)
            lo_chapters = []
            if n_lo:
                lo_chapters, _lo_pages = add_leftover_part(plan, leftovers)
                plan["leftover"] = {"name": LO_PART_NAME, "chapters": lo_chapters}

            leftovers_info = [{"kind": g["kind"],
                               "name": LO_CAT_NAMES.get(g["kind"], LO_CAT_NAMES["other"]),
                               "pages": g["pages"]} for g in leftovers]
            render_pages = len(_collect_abs_pages(plan))
            return {
                "ok": True, "kind": "pdf", "name": name, "path": path, "size": size,
                "pages": stats["total_pages"], "toc_pages": [start, end],
                "stats": stats, "offset": offset, "plan": plan,
                "counts": {"parts": n_part, "chapters": n_ch, "mains": n_main,
                           "ends": n_end},
                "tree": _build_preview_tree(entries),
                "covered_pages": len(all_abs), "render_pages": render_pages,
                "leftovers": leftovers_info, "gap_pages": n_mid,
            }
        finally:
            doc.close()

    if ext in (".docx", ".doc"):
        if ext == ".doc":
            return {"ok": False, "kind": "word", "reason": "doc_legacy",
                    "name": name, "path": path, "size": size}
        try:
            name = derive_book_name(path)
            entries, texts = parse_docx(path)
        except Exception as e:
            return {"ok": False, "kind": "word", "reason": "parse_error",
                    "name": name, "path": path, "size": size, "error": str(e)}
        plan = build_plan_word(entries, texts)
        counts = {
            "parts": sum(1 for e in entries if e["level"] == "part"),
            "chapters": sum(1 for e in entries if e["level"] == "chapter"),
            "mains": sum(1 for e in entries if e["level"] == "main"),
        }
        if counts["chapters"] == 0 and counts["mains"] == 0:
            return {"ok": False, "kind": "word", "reason": "no_headings",
                    "name": name, "path": path, "size": size}
        return {
            "ok": True, "kind": "word", "name": name, "path": path, "size": size,
            "counts": counts, "plan": plan,
            "tree": _build_preview_tree(entries, counts=texts),
        }

    return {"ok": False, "kind": None, "reason": "unsupported",
            "name": name, "path": path, "size": size}


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

def process_pdf(path, out_root, notegen_exe, cancel=None, progress=None,
                log_cb=None, analysis=None):
    """PDF 主流程：识别目录 → 渲染页面 → 生成 .one 笔记本。

    analysis 为 analyze_file 的返回（预览后直接导入时复用，避免重复解析）；
    为 None 时内部先分析。返回 (result, outdir)，无法识别时返回 None。
    """
    if analysis is None:
        analysis = analyze_file(path, cancel=cancel, progress=progress, log_cb=log_cb)

    if not analysis.get("ok"):
        reason = analysis.get("reason")
        stats = analysis.get("stats") or {}
        if reason == "scanned":
            _emit(log_cb, "检测到扫描版/图片型 PDF（无文字层），无法识别目录层级。")
            _emit(log_cb, "文本层统计: 采样%d页 / 有实质文本%d页(%.1f%%) / 总字符%d / 单页文本中位数%d" %
                  (stats.get("sampled", 0), stats.get("text_pages", 0),
                   stats.get("ratio", 0) * 100, stats.get("text_chars", 0),
                   stats.get("len_median", 0)))
            _emit(log_cb, "请先对教材进行 OCR 处理，或更换带文字层的版本后重试。")
        elif reason == "empty_toc":
            _emit(log_cb, "目录页已找到，但未能解析出任何章节条目，无法继续。")
        else:
            _emit(log_cb, "检测到文字层，但前 40 页未定位到目录页（Contents/目录），无法自动识别章节层级。")
            _emit(log_cb, "请确认教材包含目录页后重试，或更换含标准目录的版本。")
        return None

    counts = analysis.get("counts", {})
    toc_pages = analysis.get("toc_pages")
    if toc_pages:
        _emit(log_cb, "目录识别: %d PART / %d CHAPTER / %d 主节 (目录页 %d-%d)" %
              (counts.get("parts", 0), counts.get("chapters", 0),
               counts.get("mains", 0), toc_pages[0], toc_pages[1]))
    leftovers = analysis.get("leftovers") or []
    if leftovers:
        detail = ", ".join("%s(%d页)" % (g["name"], len(g["pages"])) for g in leftovers)
        _emit(log_cb, "删减内容整理: 新增分区组「%s」, 共 %d 个分区 [%s]" %
              (LO_PART_NAME, len(leftovers), detail))
    else:
        _emit(log_cb, "删减内容整理: 未发现被删减的页")
    if analysis.get("gap_pages"):
        _emit(log_cb, "删减内容整理: 正文中另有 %d 页未被目录覆盖（目录覆盖缺口，不并入删减分区）" %
              analysis["gap_pages"])

    plan = analysis["plan"]
    name = sanitize_book_name(analysis.get("name"))
    outdir = os.path.join(out_root, name)
    work = tempfile.mkdtemp(prefix="t2n_")
    doc = fitz.open(path)
    try:
        rendered = os.path.join(work, "rendered")
        plan_path = os.path.join(work, "plan.json")
        with open(plan_path, "w", encoding="utf-8") as f:
            json.dump(plan, f, ensure_ascii=False)

        all_abs = _collect_abs_pages(plan)
        _emit(log_cb, "渲染 %d 页 ..." % len(all_abs))
        render_pdf(doc, all_abs, rendered, cancel=cancel, progress=progress)
        if cancel is not None and cancel.is_set():
            raise ImportCancelled()

        _emit(log_cb, "生成 .one 笔记本 ...")
        result = run_notegen(notegen_exe, plan_path, rendered, outdir, name)
    finally:
        doc.close()
        shutil.rmtree(work, ignore_errors=True)
    return result, outdir


def process_word(path, out_root, notegen_exe, cancel=None, progress=None,
                 log_cb=None, analysis=None):
    """Word 主流程：识别标题层级 → 生成 .one 笔记本。

    analysis 为 analyze_file 的返回；为 None 时内部先分析。返回 (result, outdir)，
    无法识别时返回 None。
    """
    if analysis is None:
        analysis = analyze_file(path, cancel=cancel, progress=progress, log_cb=log_cb)

    if not analysis.get("ok"):
        reason = analysis.get("reason")
        if reason == "doc_legacy":
            _emit(log_cb, "暂不支持旧版 .doc 格式，请用 Word 另存为 .docx 后重试。")
        elif reason == "parse_error":
            _emit(log_cb, "Word 文档解析失败: %s" % analysis.get("error"))
        else:
            _emit(log_cb, "未识别到标题层级，请确认文档使用了 Heading 1/2/3 样式")
        return None

    counts = analysis["counts"]
    _emit(log_cb, "Word 目录识别: %d PART / %d CHAPTER / %d 主节" %
          (counts["parts"], counts["chapters"], counts["mains"]))

    name = sanitize_book_name(analysis.get("name"))
    outdir = os.path.join(out_root, name)
    work = tempfile.mkdtemp(prefix="t2n_")
    try:
        if cancel is not None and cancel.is_set():
            raise ImportCancelled()
        plan_path = os.path.join(work, "plan.json")
        with open(plan_path, "w", encoding="utf-8") as f:
            json.dump(analysis["plan"], f, ensure_ascii=False)
        result = run_notegen(notegen_exe, plan_path, work, outdir, name)
    finally:
        shutil.rmtree(work, ignore_errors=True)
    return result, outdir


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("file", nargs="?")
    ap.add_argument("--out", default=None, help="输出目录（默认 OneDrive 桌面）")
    ap.add_argument("--notegen", default=None, help="notegen.exe 路径")
    args = ap.parse_args()

    notegen_exe = args.notegen or os.path.join(this_dir(), "notegen.exe")
    if not os.path.exists(notegen_exe):
        log("找不到 notegen.exe: %s" % notegen_exe)
        return 2

    path = args.file or pick_file()
    if not path or not os.path.exists(path):
        log("未选择有效文件")
        return 1

    out_root = args.out or find_output_dir()
    log("输出目录: %s" % out_root)

    ext = os.path.splitext(path)[1].lower()
    try:
        if ext == ".pdf":
            r = process_pdf(path, out_root, notegen_exe)
        elif ext in (".docx", ".doc"):
            r = process_word(path, out_root, notegen_exe)
        else:
            log("不支持的文件类型: %s" % ext)
            return 1
    except ImportCancelled:
        log("已取消")
        return 130
    except FileNotFoundError as e:
        log("✗ %s" % e)
        return 2
    except Exception as e:
        log("✗ 处理失败: %s" % e)
        return 1

    if r is None:
        return 1

    result, outdir = r
    if result.get("ok"):
        log("✓ 生成成功")
        log("笔记本位置: %s" % outdir)
        log("结构: %d 分区组 / %d 分区 / %d 页" %
            (result.get("groups", 0), result.get("sections", 0), result.get("pages", 0)))
        log("排序校验: %s" % ("通过" if result.get("sortedOk") else "失败"))
        try:
            print("SUCCESS " + json.dumps({"out": outdir, **result}, ensure_ascii=False))
        except Exception:
            pass
        return 0
    else:
        log("✗ 生成失败: %s" % result.get("error"))
        return 1


if __name__ == "__main__":
    sys.exit(main())
