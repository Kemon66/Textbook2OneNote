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
import subprocess
import sys
import tempfile

import fitz  # PyMuPDF


# ---------------------------------------------------------------------------
# 工具函数
# ---------------------------------------------------------------------------

def log(msg):
    print("[T2N]", msg, flush=True)


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


def this_dir():
    """返回本脚本/可执行文件所在目录（用于定位 notegen.exe）。"""
    if getattr(sys, "frozen", False):
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
                       or re.search("目[\s\xa0]*录", t))
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


def has_text_layer(path):
    """粗略判断 PDF 是否有可提取的文本层（扫描版/图片型 PDF 为 False）。"""
    try:
        import fitz
        doc = fitz.open(path)
        try:
            for i in range(min(doc.page_count, 20)):
                if len(doc[i].get_text().strip()) >= 3:
                    return True
        finally:
            doc.close()
    except Exception:
        return False
    return False


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
            mm = re.match(r"^PART\s+([IVXLCivxlcd]+|\d+)\s*(.*)$", t, re.I)
            if mm:
                entries.append({"level": "part", "num": mm.group(1).upper(),
                                "name": mm.group(2).strip(), "print": 0})
            else:
                cm = re.match(r"^CHAPTER\s+(\d+)\s*(.*)$", t, re.I)
                if cm:
                    entries.append({"level": "chapter", "num": cm.group(1),
                                    "name": cm.group(2).strip(), "print": tpage or 0})
                else:
                    cm2 = re.match(r"^(\d{1,3})\s*[>.:]\s*(.+)$", t)
                    if cm2 and int(cm2.group(1)) <= 200:
                        entries.append({"level": "chapter", "num": cm2.group(1),
                                        "name": cm2.group(2).strip(), "print": tpage or 0})
                    else:
                        mm2 = re.match(r"^(\d{1,3})\s{1,}([A-Z0-9][^0-9].*)$", t)
                        if mm2 and int(mm2.group(1)) <= 200:
                            entries.append({"level": "chapter", "num": mm2.group(1),
                                            "name": mm2.group(2).strip(), "print": tpage or 0})
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


def calc_offset_pdf(doc, entries, toc_end):
    """锚定法 + 众数：用主节/章节标题在正文中定位，排除离群值求最稳 offset。"""
    scored = [(e, e["print"]) for e in entries
              if e["level"] in ("main", "chapter") and e["print"] is not None
              and e["name"] and len(e["name"].split()) >= 4]
    offsets = []
    for e, pr in scored[:60]:
        lo = max(toc_end + 1, pr - 1)
        abs_p = _locate_title_page(doc, e["name"], lo, doc.page_count)
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


def _locate_title_page(doc, title, lo, hi):
    """在正文 [lo,hi) 页范围内定位完整标题首次出现的绝对页（忽略空白/大小写）。
    标题可能跨行，故用压缩空白后的全文匹配。找不到返回 None。"""
    key = _title_key(title)
    if len(key) < 12:
        return None
    for i in range(lo, min(hi, doc.page_count)):
        text = doc[i].get_text()
        if not text:
            continue
        flat = _title_key(text)
        if key in flat:
            return i
    return None


def fill_missing_pages(doc, entries, toc_end, offset):
    """对 print 为 None 的 main，用正文标题搜索回填其印刷页码。
    找不到的按上一主节+1 近似，仍覆盖不到则用 chapter 起始。"""
    last_abs = None          # 上一主节已定位的正文绝对页，用于推进搜索下界
    for e in entries:
        if e["level"] == "chapter":
            # 章节有页码时顺带锚定 last_abs，保证主节搜索从章节正文页起步
            if e["print"] is not None:
                lo = max(toc_end + 1, e["print"] - 1 if last_abs is None else last_abs - 1)
                ap = _locate_title_page(doc, e["name"], lo, doc.page_count)
                if ap is not None:
                    last_abs = ap
            continue
        if e["level"] != "main":
            continue
        if e["print"] is not None and last_abs is None:
            lo = max(toc_end + 1, e["print"] - 1)
            ap = _locate_title_page(doc, e["name"], lo, doc.page_count)
            if ap is not None:
                last_abs = ap
            continue
        if e["print"] is not None:
            continue
        lo = max(toc_end + 1, (last_abs + 1) if last_abs is not None else toc_end + 1)
        ap = _locate_title_page(doc, e["name"], lo, doc.page_count)
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


def _page_lines(doc, p, n=6):
    """取页面文本的前 n 个非空行使于里程碑检测。"""
    txt = (doc[p].get_text() or "").strip()
    lines = [l.strip() for l in txt.splitlines() if l.strip()]
    return lines[:n]


def _milestone_page_kind(doc, p):
    """书尾独立部分起始页探测：前几行命中独立标题词 -> (kind)或 None。"""
    for line in _page_lines(doc, p, 6):
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


def _region_kind(doc, pages, toc_start, toc_end):
    """对一段删减页整体归类：逐页关键词加权后按众数取 kind。
    避免单页中"Brief Contents 列出附录条目"之类的噪声词主导整段。
    """
    toc_cnt = sum(1 for p in pages if toc_start <= p <= toc_end)
    if toc_cnt >= max(1, len(pages) * 0.4):
        return "toc"
    votes = {}
    for p in pages:
        text = (doc[p].get_text() or "").lower()
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


def classify_leftovers(doc, toc_start, toc_end, covered):
    """识别未被目录覆盖的 PDF 页并按语义归类为删减分区。

    只处理两类明确区域，避免正文页因页码偏移被误当删减内容：
      · 书首区：绝对页 < 首个被覆盖页（封面/版权/前言/目录等）
      · 书尾区：绝对页 > 最后被覆盖页（附录/尾注/术语表/索引/后记等）
    正文中间散布的未覆盖页属"目录覆盖缺口"，不纳入删减分区（单独计数报告）。
    返回 [{"kind": str, "pages": [abs...]}]，每 kind 只会出现一次（跨区间合并）。
    """
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
        sub_toc = [p for p in rg if toc_start <= p <= toc_end]
        if sub_toc:
            groups.append(("toc", sub_toc))
        sub_rest = [p for p in rg if p not in sub_toc]
        for seg in _consecutive(sub_rest):
            groups.append((_region_kind(doc, seg, toc_start, toc_end), seg))
    # ---- 书尾区：里程碑切段，无里程碑段按关键词归并 ----
    for rg in _consecutive(tail_pages):
        seg_kind, seg_pages, pending = None, [], []
        for p in rg:
            mk = _milestone_page_kind(doc, p)
            if mk is not None:
                if pending:                       # 里程碑前积累的无标签页
                    groups.append((_region_kind(doc, pending, toc_start, toc_end), pending))
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
            groups.append((_region_kind(doc, pending, toc_start, toc_end), pending))
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

def render_pdf(doc, abs_pages, outdir):
    os.makedirs(outdir, exist_ok=True)
    n = 0
    for ap in sorted(abs_pages):
        png = os.path.join(outdir, "page_%04d.png" % ap)
        if os.path.exists(png):
            continue
        doc[ap].get_pixmap(dpi=150).save(png)
        n += 1
    return n


def run_notegen(exe, plan_path, rendered, outdir, name):
    cmd = [exe, plan_path, rendered, outdir, name]
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    for line in r.stdout.splitlines():
        line = line.strip()
        if line.startswith("RESULT "):
            try:
                return json.loads(line[len("RESULT "):])
            except Exception:
                return {"ok": False, "error": "parse result failed: " + line}
    return {"ok": False, "error": "no RESULT from notegen; stdout=" + r.stdout[-500:] + " stderr=" + r.stderr[-500:]}


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

def process_pdf(path, out_root, notegen_exe):
    doc = fitz.open(path)
    start, end = detect_toc_pages_pdf(doc)
    if start is None:
        log("未能定位目录页")
        return None
    entries = parse_toc_pdf(doc, start, end)
    n_part = sum(1 for e in entries if e["level"] == "part")
    n_ch = sum(1 for e in entries if e["level"] == "chapter")
    n_main = sum(1 for e in entries if e["level"] == "main")
    log("目录识别: %d PART / %d CHAPTER / %d 主节 (目录页 %d-%d)" % (n_part, n_ch, n_main, start, end))
    if n_ch == 0 and n_main == 0:
        log("目录解析结果为空，无法继续")
        return None

    offset = calc_offset_pdf(doc, entries, end)
    if offset is None:
        log("偏移计算失败，回退 0")
        offset = 0
    log("页码偏移 offset=%d (abs = print + offset)" % offset)

    # 回填无页码主节（纯正文标题搜索定位）
    n_missing = sum(1 for e in entries if e["level"] == "main" and e["print"] is None)
    if n_missing:
        log("正文定位回填 %d 个无页码主节..." % n_missing)
        fill_missing_pages(doc, entries, end, offset)
        still = sum(1 for e in entries if e["level"] == "main" and e["print"] is None)
        log("回填后仍缺 %d" % still)

    plan = build_plan(entries, offset)
    all_abs = set()
    for part in plan["parts"]:
        for ch in part["chapters"]:
            for m in ch["mains"]:
                all_abs.update(m.get("abs_pages", []))
            all_abs.update(ch.get("ends_abs_pages", []))

    # --- 删减内容整理：未被目录覆盖的页自动归类为 前言/目录/附录/索引/后记 等分区 ---
    leftovers, n_lo, n_mid = classify_leftovers(doc, start, end, all_abs)
    if n_mid:
        log("删减内容整理: 正文中另有 %d 页未被目录覆盖（目录覆盖缺口，不并入删减分区）" % n_mid)
    if n_lo:
        lo_chapters, lo_pages = add_leftover_part(plan, leftovers)
        all_abs.update(lo_pages)
        detail = ", ".join("%s(%d页)" % (LO_CAT_NAMES.get(g["kind"], "其他内容"), len(g["pages"]))
                           for g in leftovers)
        log("删减内容整理: 新增分区组「%s」, 共 %d 个分区 [%s]" % (LO_PART_NAME, n_lo, detail))
        plan["leftover"] = {"name": LO_PART_NAME, "chapters": lo_chapters}
    else:
        log("删减内容整理: 未发现被删减的页")

    name = os.path.splitext(os.path.basename(path))[0]
    outdir = os.path.join(out_root, name)
    work = tempfile.mkdtemp(prefix="t2n_")
    rendered = os.path.join(work, "rendered")
    plan_path = os.path.join(work, "plan.json")
    with open(plan_path, "w", encoding="utf-8") as f:
        json.dump(plan, f, ensure_ascii=False)

    log("渲染 %d 页 ..." % len(all_abs))
    render_pdf(doc, all_abs, rendered)
    doc.close()

    log("生成 .one 笔记本 ...")
    result = run_notegen(notegen_exe, plan_path, rendered, outdir, name)
    return result, outdir


def process_word(path, out_root, notegen_exe):
    entries, texts = parse_docx(path)
    plan = build_plan_word(entries, texts)
    n_ch = sum(1 for e in entries if e["level"] == "chapter")
    n_main = sum(1 for e in entries if e["level"] == "main")
    log("Word 目录识别: %d CHAPTER / %d 主节" % (n_ch, n_main))
    if n_ch == 0 and n_main == 0:
        log("未识别到标题层级，请确认文档使用了 Heading 样式")
        return None

    name = os.path.splitext(os.path.basename(path))[0]
    outdir = os.path.join(out_root, name)
    work = tempfile.mkdtemp(prefix="t2n_")
    plan_path = os.path.join(work, "plan.json")
    with open(plan_path, "w", encoding="utf-8") as f:
        json.dump(plan, f, ensure_ascii=False)

    result = run_notegen(notegen_exe, plan_path, work, outdir, name)
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
    if ext == ".pdf":
        r = process_pdf(path, out_root, notegen_exe)
    elif ext in (".docx", ".doc"):
        r = process_word(path, out_root, notegen_exe)
    else:
        log("不支持的文件类型: %s" % ext)
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
        print("SUCCESS " + json.dumps({"out": outdir, **result}, ensure_ascii=False))
        return 0
    else:
        log("✗ 生成失败: %s" % result.get("error"))
        return 1


if __name__ == "__main__":
    sys.exit(main())
