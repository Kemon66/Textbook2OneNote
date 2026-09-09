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
        if re.search(r"Brief\s+Contents", t, re.I):
            brief = i
        if re.search(r"\bContents\b", t, re.I) and not re.search(r"Brief\s+Contents", t, re.I):
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


def _classify(body):
    """按行首识别条目类型，返回 (level, num, name) 或 None。"""
    if re.match(r"^PART\s+[IVXLCivxlcd\d]", body, re.I):
        mm = re.match(r"^PART\s+([IVXLCivxlcd]+|\d+)\s*(.*)$", body, re.I)
        if mm:
            return "part", mm.group(1).upper(), mm.group(2).strip()
    if re.match(r"^CHAPTER\s+\d+", body, re.I):
        mm = re.match(r"^CHAPTER\s+(\d+)\s*(.*)$", body, re.I)
        if mm:
            return "chapter", mm.group(1), mm.group(2).strip()
    # 主节: 1-1 / 1.1 (无字母后缀)；子节 1-1a 忽略
    mm = re.match(r"^(\d{1,2})[-.](\d{1,2})(?![a-zA-Z])\s*(.*)$", body)
    if mm:
        return "main", mm.group(1) + "-" + mm.group(2), mm.group(3).strip()
    return None


# 章末要素关键词白名单（正文内嵌框 FYI/CASE STUDY/IN THE NEWS 等忽略）
END_KEYWORDS = [
    "chapter in a nutshell", "key concepts", "questions for review",
    "problems and applications", "quick quiz", "summary", "glossary",
    "key terms", "review questions", "exercises", "critical thinking",
]


def parse_toc_pdf(doc, start, end):
    """逐行解析目录，返回 entries 列表。"""
    raw = []
    for i in range(start, end + 1):
        raw.extend(doc[i].get_text().split("\n"))

    entries = []
    pending = None  # (level, num, title_part) 跨行标题暂存
    for ln in raw:
        s = ln.strip()
        if not s:
            continue
        if re.fullmatch(r"[ivxlcdm]+", s, re.I):  # 页眉罗马页码
            continue
        if re.match(r"^\d{1,2}[-.]\d{1,2}[a-z]", s):  # 子节(1-1a)，忽略
            continue
        # 纯 "CHAPTER N" 章节号行：数字是章节号不是页码
        cm = re.match(r"^CHAPTER\s+(\d+)\s*$", s, re.I)
        if cm:
            pending = ("chapter", cm.group(1), "")
            continue
        m = re.search(r"(\d+)\s*$", s)
        page = int(m.group(1)) if m else None
        body = s[:m.start()].strip() if m else s

        c = _classify(body)
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
            # 章末要素等：无编号、有页码，仅保留白名单关键词
            if page is not None and body:
                for kw in END_KEYWORDS:
                    if kw in body.lower():
                        entries.append({"level": "end", "num": "", "name": body, "print": page})
                        break
    return entries


def calc_offset_pdf(doc, entries, toc_end):
    """锚定法：用主节标题在正文中定位，求 print->abs 偏移。"""
    mains = [e for e in entries if e["level"] == "main"]
    offsets = []
    for m in mains[:20]:
        words = m["name"].split()
        if len(words) < 3:
            continue
        key = " ".join(words[:4]).lower()
        for i in range(toc_end + 1, doc.page_count):
            if key in doc[i].get_text().lower():
                offsets.append(i - m["print"])
                break
    if not offsets:
        return None
    offsets.sort()
    return offsets[len(offsets) // 2]


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
                    ep = mains[i + 1]["print_start"] - 1
                elif ch["ends_print_pages"]:
                    ep = min(ch["ends_print_pages"]) - 1
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

    plan = build_plan(entries, offset)
    all_abs = set()
    for part in plan["parts"]:
        for ch in part["chapters"]:
            for m in ch["mains"]:
                all_abs.update(m.get("abs_pages", []))
            all_abs.update(ch.get("ends_abs_pages", []))

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
