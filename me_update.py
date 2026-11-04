# -*- coding: utf-8 -*-
"""
M&E 台账更新：用 M&E Documents Status Tracker 更新
  1) Transmittals List-Code G（只更新指定负责人、指定文件类型的行，默认 陈书翰 / SDR）
  2) Smart City Related Document Status（Material Submittals / Shopdrawing / MS / TQ / ITP）
改动的单元格标黄，并在文件末尾生成「更新记录」工作表。
"""
import collections
import datetime as dt
import os
import re

import openpyxl
from openpyxl.styles import Font, PatternFill

YELLOW = "FFFFFF00"
YEL = PatternFill("solid", fgColor=YELLOW)
LOG_SHEET = "更新记录"

DEFAULT_ME_CONFIG = {
    "codeg_people": ["陈书翰"],
    "codeg_types": ["SDR"],
    "done_codes": ["A", "B", "C", "D", "O"],
    "codeg_name_pattern": "*Code G*.xlsx",
    "smartcity_name_pattern": "*Smart City*.xlsx",
    "material_groups": {
        "MR/MTS": ["MR", "MTS"], "TBE": ["TBE"], "AR": ["AR"], "PO": ["PO"],
        "NMR 601": ["NMR 601"], "NMR 602": ["NMR 602"], "NMR 603": ["NMR 603"],
    },
}


def _norm(s):
    return re.sub(r"[^a-z0-9]", "", str(s).lower()) if s is not None else ""


# ======================================================================
# M&E 跟踪表索引
# ======================================================================

def tr_num(t):
    if t is None:
        return None
    m = re.search(r"(\d{3,5})\s*$", str(t).strip())
    return int(m.group(1)) if m else None


def doc_key(d):
    """文件编号去掉版本后缀，作为同一文件的识别键，例如 ('PCL','SDR',346,'')。"""
    if d is None:
        return None
    s = re.sub(r"\s", "", str(d).upper()).replace(".", "-").replace("_", "-").replace("(N)", "N")
    s = re.sub(r"-+", "-", s)
    s = re.sub(r"/\d+$", "", s)
    m = re.search(r"(PCL|QUA|INT|HSE|DMP)-?([A-Z0-9]+)-(\d+)(N?)$", s)
    if not m:
        return s
    return (m.group(1), m.group(2), int(m.group(3)), m.group(4))


def rev_num(r):
    if r is None:
        return -1
    m = re.search(r"(\d+)", str(r).upper().replace("O", "0"))
    return int(m.group(1)) if m else -1


def code_from(k, l):
    """Status Latest（k）优先；空白时（文件已升版）从 JCDC Comments（l）识别。"""
    if k is not None and str(k).strip():
        s = str(k).strip()
        m = re.match(r"(?i)code[- ]?([A-Z])$", s)
        if m:
            return "Code-" + m.group(1).upper()
        return s
    if not l:
        return None
    s = str(l).strip()
    u = s.upper()
    if "INFORMATION" in u or "(O)" in u:
        return "Code-O"
    if "RECALL" in u:
        return "Recall"
    m = re.search(r"[-\s(]\(?([A-G])\)?\.?\s*$", s)
    if m:
        return "Code-" + m.group(1)
    if "DONOTRESUBMIT" in u.replace(" ", ""):
        return "Code-B"
    if "RESUBMIT" in u and "APPROVED" in u:
        return "Code-G"
    if "REVISE" in u:
        return "Code-E"
    if "REJECT" in u:
        return "Code-F"
    if "NO OBJECTION" in u:
        return "Code-C"
    if u.startswith("APPROVED"):
        return "Code-A"
    return None


def to_dt(v):
    if isinstance(v, dt.datetime):
        return v
    if isinstance(v, dt.date):
        return dt.datetime(v.year, v.month, v.day)
    if v is None:
        return None
    s = str(v).strip()
    for f in ("%d-%B-%y", "%d-%B-%Y", "%d-%b-%y", "%d-%b-%Y", "%m/%d/%Y", "%d-%m-%Y",
              "%Y-%m-%d", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S"):
        try:
            return dt.datetime.strptime(s, f)
        except ValueError:
            pass
    return None


_ME_FIELDS = {
    "tr": lambda h: h.startswith("transmittalno"),
    "rev": lambda h: h == "rev",
    "doc": lambda h: h == "documentno",
    "date": lambda h: h == "datesubmitted",
    "desc": lambda h: h.startswith("description"),
    "resp": lambda h: h == "responsereceiveddate",
    "k": lambda h: h == "statuslatest",
    "l": lambda h: h == "jcdccomments",
}
_ME_FALLBACK = {"tr": 1, "rev": 2, "doc": 3, "date": 4, "desc": 5, "resp": 9, "k": 10, "l": 11}


def load_me(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    recs = []
    for ws in wb.worksheets:
        hdr_row, cols = None, None
        for r in range(1, 7):
            vals = [_norm(c.value) for c in ws[r]]
            if "datesubmitted" in vals and "documentno" in vals:
                hdr_row = r
                cols = {}
                for key, test in _ME_FIELDS.items():
                    cols[key] = next((i for i, h in enumerate(vals) if test(h)), _ME_FALLBACK[key])
                break
        if not hdr_row:
            continue
        for ridx, row in enumerate(ws.iter_rows(min_row=hdr_row + 1, values_only=True), start=hdr_row + 1):
            g = lambda k: row[cols[k]] if cols[k] < len(row) else None
            tr, doc = g("tr"), g("doc")
            if not tr and not doc:
                continue
            rec = dict(sheet=ws.title, row=ridx, tr=str(tr).strip() if tr else None, trn=tr_num(tr),
                       rev=g("rev"), revn=rev_num(g("rev")), doc=doc, key=doc_key(doc),
                       date=to_dt(g("date")), date_raw=g("date"), desc=g("desc"),
                       resp=to_dt(g("resp")), resp_raw=g("resp"), k=g("k"), l=g("l"))
            rec["code"] = code_from(rec["k"], rec["l"])
            recs.append(rec)
    wb.close()
    latest = {}
    groups = collections.defaultdict(list)
    for x in recs:
        if x["key"]:
            groups[x["key"]].append(x)
    for k, v in groups.items():
        latest[k] = max(v, key=lambda x: (x["revn"], x["trn"] or 0, x["date"] or dt.datetime(1900, 1, 1)))
    bytr = collections.defaultdict(list)
    for x in recs:
        if x["trn"]:
            bytr[x["trn"]].append(x)
    return recs, latest, bytr


# ======================================================================
# 公共：写单元格、更新记录、清除上次标黄
# ======================================================================

def ncode(c):
    if c is None:
        return None
    u = str(c).upper()
    if u == "OPEN":
        return "Open"
    if u.startswith("RECALL"):
        return "Recall"
    return c


def nst(s):
    s = re.sub(r"[^A-Z]", "", str(s).upper()).replace("CODE", "")
    return {"SUBMITTED": "OPEN", "PENDING": "OPEN", "RECALLED": "RECALL"}.get(s, s)


def _fill_repr(fill):
    if fill is None or fill.fill_type in (None, "none"):
        return "none"
    c = fill.fgColor
    if c.type == "rgb" and isinstance(c.rgb, str):
        return f"{fill.fill_type}|rgb|{c.rgb}"
    if c.type == "theme":
        return f"{fill.fill_type}|theme|{c.theme}|{c.tint}"
    if c.type == "indexed":
        return f"{fill.fill_type}|indexed|{c.indexed}"
    return "none"


def _fill_from(rep):
    from openpyxl.styles.colors import Color
    if not rep or rep == "none":
        return PatternFill(fill_type=None)
    p = rep.split("|")
    if p[1] == "rgb":
        col = Color(rgb=p[2])
    elif p[1] == "theme":
        col = Color(theme=int(p[2]), tint=float(p[3]))
    else:
        col = Color(indexed=int(p[2]))
    return PatternFill(fill_type=p[0], fgColor=col, bgColor=col)


def _fmt(v):
    if v is None:
        return ""
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d")
    return str(v)


class Writer:
    def __init__(self, wb):
        self.wb = wb
        self.log = []
        self.done = set()

    def set(self, ws, cell, val, ctx):
        c = ws[cell] if isinstance(cell, str) else cell
        old = c.value
        if _fmt(old) == _fmt(val):
            return False
        key = (ws.title, c.coordinate)
        orig_fill = None if key in self.done else _fill_repr(c.fill)
        if isinstance(val, dt.datetime) and c.number_format == "General":
            c.number_format = "yyyy-mm-dd"
        c.value = val
        c.fill = YEL
        self.done.add(key)
        self.log.append((ws.title, c.coordinate, ctx, _fmt(old), _fmt(val), orig_fill or ""))
        return True


def clear_previous_marks(wb, notes):
    """按上次的「更新记录」把标黄的单元格恢复成原来的底色。"""
    if LOG_SHEET not in wb.sheetnames:
        return 0
    ws = wb[LOG_SHEET]
    hdr = [c.value for c in ws[1]]
    if "原底色" not in hdr:
        notes.append("上次的「更新记录」是旧格式（没有记录原底色），上次的黄色标记无法自动清除，已保留。")
        return 0
    ci = hdr.index("原底色")
    n = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        sheet, coord, rep = row[0], row[1], row[ci]
        if sheet in wb.sheetnames and coord and rep:
            cell = wb[sheet][coord]
            if _fill_repr(cell.fill) == f"solid|rgb|{YELLOW}":
                cell.fill = _fill_from(rep)
                n += 1
    return n


def write_log(wb, log, title_note):
    if LOG_SHEET in wb.sheetnames:
        del wb[LOG_SHEET]
    ws = wb.create_sheet(LOG_SHEET)
    ws.append(["工作表", "单元格", "文件/说明", "原值", "新值", "原底色"])
    ws["G1"] = f"{title_note}    生成时间：{dt.datetime.now():%Y-%m-%d %H:%M}"
    for x in log:
        ws.append(list(x))
    for c in ws[1]:
        c.font = Font(bold=True, name="Arial")
    for c in ws[1][:6]:
        c.fill = YEL
    for w, col in zip([24, 10, 70, 32, 32, 18, 50], "ABCDEFG"):
        ws.column_dimensions[col].width = w
    ws.column_dimensions["F"].hidden = True
    ws.freeze_panes = "A2"


# ======================================================================
# Code-G
# ======================================================================

def _hnorm(s):
    """表头比较用：去掉空白和标点，保留中文。"""
    return re.sub(r"[\s.\-_/#:：()（）]", "", str(s).lower()) if s is not None else ""


def _find_cols(ws, header_row, names):
    hdr = [(_hnorm(c.value), c.column) for c in ws[header_row]]
    out = {}
    for key, (target, which) in names.items():
        hits = [col for h, col in hdr if h and h == _hnorm(target)]
        if hits:
            out[key] = hits[0] if which == "first" else hits[-1]
    return out


def update_codeg(path, me, out_path, cfg, log=print, clear_old=True):
    recs, latest, bytr = me
    wb = openpyxl.load_workbook(path)
    ws = next((w for w in wb.worksheets if "code g" in w.title.lower()), wb.worksheets[0])
    hdr_row = next((r for r in range(1, 8)
                    if any(_norm(c.value) == "documentno" for c in ws[r])), 3)
    cols = _find_cols(ws, hdr_row, {
        "tr": ("Transmittal No. CHEC", "first"), "rev": ("Rev.", "first"),
        "doc": ("Document No", "first"), "person": ("Responsible person", "first"),
        "upd": ("更新情况", "first"), "state": ("状态", "first"),
        "status": ("Status", "first"), "trnew": ("Transmittal No. CHEC", "last"),
    })
    need = ["tr", "doc", "person", "upd", "state", "status", "trnew"]
    miss = [k for k in need if k not in cols]
    if miss or cols["tr"] == cols["trnew"]:
        raise RuntimeError(f"Code-G 表第 {hdr_row} 行找不到这些列：{miss or ['最后一列 Transmittal No. CHEC']}")

    notes = []
    w = Writer(wb)
    if clear_old:
        n = clear_previous_marks(wb, notes)
        if n:
            notes.append(f"已清除上次更新留下的 {n} 处黄色标记。")
    people = {p.strip() for p in cfg.get("codeg_people", []) if p.strip()}
    types = {t.strip().upper() for t in cfg.get("codeg_types", []) if t.strip()}
    done_codes = {"Code-" + c.upper() for c in cfg.get("done_codes", ["A", "B", "C", "D", "O"])}

    checked = nomatch = 0
    suspicious = []
    last = max((r for r in range(hdr_row + 1, ws.max_row + 1) if ws.cell(r, cols["tr"]).value), default=hdr_row)
    for r in range(hdr_row + 1, last + 1):
        if not ws.cell(r, cols["tr"]).value:
            continue
        person = str(ws.cell(r, cols["person"]).value or "").strip()
        doc = ws.cell(r, cols["doc"]).value
        k = doc_key(doc)
        if people and person not in people:
            continue
        if types and not (isinstance(k, tuple) and k[1] in types):
            continue
        checked += 1
        if k not in latest:
            nomatch += 1
            notes.append(f"第 {r} 行 {doc}：M&E 里找不到这个文件，未更新。")
            continue
        L = latest[k]
        rowtrn = tr_num(ws.cell(r, cols["tr"]).value)
        Rtrn = tr_num(ws.cell(r, cols["trnew"]).value)
        Ltrn = L["trn"] or 0
        code = ncode(L["code"])
        if code is None:
            continue
        # R 列的 Transmittal 号在 M&E 里对应的是别的文件 → 可能填错，提示
        if Rtrn and Rtrn != rowtrn:
            hits = bytr.get(Rtrn, [])
            if hits and not any(h["key"] == k for h in hits):
                suspicious.append(f"第 {r} 行 {doc}：R 列 {ws.cell(r, cols['trnew']).value} 在 M&E 里对应 "
                                  + "、".join(str(h["doc"]) for h in hits) + "，可能填错（未改动）。")
        done = "已完成" if code in done_codes else "已上报"
        ctx = f"{doc} → M&E 最新 {L['tr']} {L['rev']} {code}"
        cur_status = ws.cell(r, cols["status"]).value
        if Ltrn > max(rowtrn or 0, Rtrn or 0):
            w.set(ws, ws.cell(r, cols["upd"]), "已更新", ctx)
            w.set(ws, ws.cell(r, cols["state"]), done, ctx)
            if nst(cur_status) != nst(code):
                w.set(ws, ws.cell(r, cols["status"]), code, ctx)
            w.set(ws, ws.cell(r, cols["trnew"]), L["tr"], ctx)
        elif Ltrn == (Rtrn or rowtrn) and nst(cur_status) != nst(code):
            w.set(ws, ws.cell(r, cols["status"]), code, ctx)
            w.set(ws, ws.cell(r, cols["state"]), done, ctx)

    rows_changed = len({re.sub(r"[A-Z]+", "", x[1]) for x in w.log})
    scope = f"负责人={'、'.join(sorted(people)) or '全部'}；类型={'、'.join(sorted(types)) or '全部'}"
    write_log(wb, w.log, f"Code-G 更新（{scope}）")
    wb.save(out_path)
    log(f"Code-G：核对 {checked} 行（{scope}），改动 {rows_changed} 行 / {len(w.log)} 个单元格。")
    for s in suspicious + notes:
        log("  · " + s)
    return dict(out=out_path, checked=checked, rows=rows_changed, cells=len(w.log),
                notes=suspicious + notes, log=w.log)


# ======================================================================
# Smart City
# ======================================================================

def _style_status(code, cur):
    if code == "Open":
        base = "OPEN" if cur and str(cur).isupper() else ("Submitted" if cur and "SUBMIT" in str(cur).upper() else "Open")
    elif code == "Recall":
        base = "Recalled"
    else:
        base = code
    if cur and str(cur).startswith("CODE"):
        base = base.upper()
    return base


_DOC_SHEET_COLS = {
    "tr": ["transmittalno", "transmittalnochec"],
    "rev": ["rev"],
    "doc": ["documentno"],
    "sub": ["submissiondate", "datesubmitted"],
    "appr": ["approvaldate", "responsereceiveddate"],
    "status": ["currentstatus", "status"],
}


def update_smartcity(path, me, out_path, cfg, log=print, clear_old=True):
    recs, latest, bytr = me
    wb = openpyxl.load_workbook(path)
    notes = []
    w = Writer(wb)
    if clear_old:
        n = clear_previous_marks(wb, notes)
        if n:
            notes.append(f"已清除上次更新留下的 {n} 处黄色标记。")

    per_sheet = collections.Counter()
    for ws in wb.worksheets:
        if ws.title == LOG_SHEET:
            continue
        # ---------- Material Submittals（分组表头）----------
        row1 = [_norm(c.value) for c in ws[1]]
        groups = []
        for name, sheets in cfg.get("material_groups", DEFAULT_ME_CONFIG["material_groups"]).items():
            if _norm(name) in row1:
                groups.append((name, row1.index(_norm(name)) + 1, set(sheets)))
        if len(groups) >= 2:
            for r in range(3, ws.max_row + 1):
                for name, c, sheets in groups:
                    tr = ws.cell(r, c).value
                    if not tr or str(tr).strip().upper() in ("N/A", "NA", "-") or "\n" in str(tr):
                        continue
                    cand = [x for x in bytr.get(tr_num(tr), []) if x["sheet"] in sheets]
                    if not cand:
                        continue
                    L = latest[cand[0]["key"]]
                    code = ncode(L["code"])
                    Ltrn = L["trn"] or 0
                    ctx = f"{ws.cell(r, 2).value} {name}：{cand[0]['doc']} → M&E {L['tr']} {L['rev']} {code}"
                    cur = ws.cell(r, c + 4).value
                    before = len(w.log)
                    if Ltrn > (tr_num(tr) or 0):
                        w.set(ws, ws.cell(r, c), L["tr"], ctx)
                        w.set(ws, ws.cell(r, c + 1), L["rev"], ctx)
                        w.set(ws, ws.cell(r, c + 2), L["date"] or L["date_raw"], ctx)
                        w.set(ws, ws.cell(r, c + 3), L["resp"] or ("Pending" if code == "Open" else None), ctx)
                        w.set(ws, ws.cell(r, c + 4), _style_status(code, cur), ctx)
                    elif code and nst(code) != nst(cur):
                        w.set(ws, ws.cell(r, c + 4), _style_status(code, cur), ctx)
                        if L["resp"]:
                            w.set(ws, ws.cell(r, c + 3), L["resp"], ctx)
                    if len(w.log) > before:
                        per_sheet[ws.title] += 1
            continue

        # ---------- 按文件编号的表（Shopdrawing / MS / TQ / ITP 等）----------
        hdr = [_norm(c.value) for c in ws[1]]
        cols = {}
        for key, names in _DOC_SHEET_COLS.items():
            for nm in names:
                if nm in hdr:
                    cols[key] = hdr.index(nm) + 1
                    break
        if not all(k in cols for k in ("tr", "doc", "status")):
            continue
        for r in range(2, ws.max_row + 1):
            tr = ws.cell(r, cols["tr"]).value
            doc = ws.cell(r, cols["doc"]).value
            if not tr or not doc or "\n" in str(tr):
                continue
            L = latest.get(doc_key(doc))
            if not L:
                continue
            code = ncode(L["code"])
            rowtrn = tr_num(tr) or 0
            Ltrn = L["trn"] or 0
            ctx = f"{doc} → M&E {L['sheet']}表 {L['tr']} {L['rev']} {code}"
            cur = ws.cell(r, cols["status"]).value
            before = len(w.log)
            if Ltrn > rowtrn:
                w.set(ws, ws.cell(r, cols["tr"]), L["tr"], ctx)
                if "rev" in cols:
                    w.set(ws, ws.cell(r, cols["rev"]), L["rev"], ctx)
                if re.search(r"/\d+\s*$", str(doc)):
                    w.set(ws, ws.cell(r, cols["doc"]), L["doc"], ctx)
                if "sub" in cols:
                    w.set(ws, ws.cell(r, cols["sub"]), L["date"] or L["date_raw"], ctx)
                if "appr" in cols:
                    w.set(ws, ws.cell(r, cols["appr"]), L["resp"] or L["resp_raw"], ctx)
                if code:
                    w.set(ws, ws.cell(r, cols["status"]), _style_status(code, cur), ctx)
            elif Ltrn == rowtrn and code and nst(code) != nst(cur):
                w.set(ws, ws.cell(r, cols["status"]), _style_status(code, cur), ctx)
                if L["resp"] and "appr" in cols:
                    w.set(ws, ws.cell(r, cols["appr"]), L["resp"], ctx)
            if len(w.log) > before:
                per_sheet[ws.title] += 1

    write_log(wb, w.log, "Smart City 更新")
    wb.save(out_path)
    detail = "，".join(f"{k} {v} 处" for k, v in per_sheet.items()) or "无改动"
    log(f"Smart City：改动 {sum(per_sheet.values())} 处 / {len(w.log)} 个单元格（{detail}）。")
    for s in notes:
        log("  · " + s)
    return dict(out=out_path, rows=sum(per_sheet.values()), cells=len(w.log), notes=notes, log=w.log)


# ======================================================================
# 输出文件名
# ======================================================================

def dated_name(src_path, out_dir, date=None):
    """文件名里有 8 位日期就换成今天；没有就在末尾加 -YYYYMMDD。"""
    date = date or dt.date.today()
    stem, ext = os.path.splitext(os.path.basename(src_path))
    stem = stem.strip()
    tag = date.strftime("%Y%m%d")
    if re.search(r"\d{8}", stem):
        stem = re.sub(r"\d{8}", tag, stem, count=1)
    else:
        stem = f"{stem}-{tag}"
    p = os.path.join(out_dir, stem + ext)
    if os.path.abspath(p) == os.path.abspath(src_path):
        p = os.path.join(out_dir, stem + "_新" + ext)
    return p
