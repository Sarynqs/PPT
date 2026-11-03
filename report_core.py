# -*- coding: utf-8 -*-
"""
周报更新核心逻辑：
从 M&E Documents Status Tracker (xlsx) 中提取指定日期范围内的提交 / 批复数据，
填入上周的周例会 PPT（第 1.1–1.4 小节表格、每页统计、标题日期、封面日期）。
"""
import copy
import datetime as dt
import json
import math
import os
import re

from openpyxl import load_workbook
from pptx import Presentation
from pptx.util import Emu

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

DEFAULT_CONFIG = {
    "pass_codes": ["A", "B", "C", "O"],
    "open_status": ["Open"],
    "sections": {
        "1.1": {"type_column": {"MTS": "MTS", "MR": "MR", "TBE": "TBE",
                                "NMR 601": "NMR 601", "NMR 602": "NMR 602"}},
        "1.2": {"sheets": ["Shop Dwg"]},
        "1.3": {"sheets": ["MS"]},
        "1.4": {"type_column": {"TQ": "TQ", "Proposal": "Proposal"}},
    },
    "prefix_patterns": [
        r"^Material\s+Technical\s+Submit+als?\s+(for|of)\s+",
        r"^Material\s+Requisitions?\s+(for|of)\s+",
        r"^Tech\w*\s+Bid\s+Evaluations?\s+(for|of)\s+",
        r"^Non[-\s]*Material\s+Requirements?\s*(\(\s*NMR[-\s]*\d*\s*\)|NMR[-\s]*\d+|\d+)?\s*(for|of)?\s*",
        r"^Shop\s+Drawings?\s+(for|of)\s+",
        r"^Method\s+(of\s+)?Statements?\s+(for|of)\s+",
        r"^Proposal\s+(for|of)\s+",
        r"^Technical\s+Query\s+(for|of|on|regarding)\s+",
        r"^TQ\s*[-–:]\s*",
    ],
    "tracker_pattern": "*Status*Tracker*.xlsx",
    "output_name": "{date} 技术部周例会汇报材料.pptx",
    "default_output_dir": "E:\\EdgeDOWNLODAD",
    "min_font_pt": 8,
    "slide_bottom_limit_emu": 4870000,
}


def load_config(path=None):
    cfg = copy.deepcopy(DEFAULT_CONFIG)
    if path and os.path.isfile(path):
        with open(path, encoding="utf-8") as f:
            user = json.load(f)
        for k, v in user.items():
            if not k.startswith("_"):
                cfg[k] = v
    return cfg


# ======================================================================
# Excel 读取
# ======================================================================

def _norm(s):
    return re.sub(r"[^a-z0-9]", "", str(s).lower()) if s is not None else ""


FIELDS = {
    "sno": ["sno"],
    "rev": ["rev"],
    "docno": ["documentno"],
    "submitted": ["datesubmitted"],
    "desc": ["descriptionsubject", "description"],
    "resp": ["responsereceiveddate"],
    "status": ["statuslatest"],
    "comments": ["jcdccomments"],
}


def _find_header(ws):
    for r in range(1, 7):
        vals = [_norm(c.value) for c in ws[r]]
        if "datesubmitted" in vals and any(v.startswith("description") for v in vals):
            cols = {}
            for key, names in FIELDS.items():
                for i, v in enumerate(vals):
                    if any(v == n or (n == "description" and v.startswith(n)) for n in names):
                        cols[key] = i
                        break
            return r, cols
    return None, None


def to_date(v):
    if v is None or v == "":
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, (int, float)) and 20000 < v < 80000:  # Excel 序列号
        return (dt.datetime(1899, 12, 30) + dt.timedelta(days=float(v))).date()
    s = str(v).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y",
                "%d-%b-%Y", "%d-%b-%y", "%d %b %Y", "%Y-%m-%d %H:%M:%S"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


_CODE_TEXT = [
    (r"do\s*not\s*re-?submit", "B"),
    (r"approved\s*w(ith)?\s*/?\s*comments?\s*,?\s*re-?submit", "G"),
    (r"revised?\s*and\s*re-?submit", "E"),
    (r"reject", "F"),
    (r"reviewed\s*with\s*comments", "D"),
    (r"no\s*objection", "C"),
    (r"information|record", "O"),
]


def parse_code(text):
    """从状态 / 评论文字中识别批复代码（A–I、O），识别不了返回 None。"""
    if text is None:
        return None
    t = str(text).strip()
    if not t:
        return None
    m = re.search(r"code\s*[-_ ]?\s*([A-IO])\b", t, re.I)
    if m:
        return m.group(1).upper()
    m = re.search(r"[-–(\s]\s*([A-IO])\s*\)?\s*\.?\s*$", t)
    if m:
        return m.group(1)
    for pat, code in _CODE_TEXT:
        if re.search(pat, t, re.I):
            return code
    return None


class Record:
    __slots__ = ("sheet", "sno", "rev", "desc", "submitted", "resp",
                 "status", "comments", "code", "row")

    def __repr__(self):
        return f"<{self.sheet} #{self.sno} {self.rev} {self.desc[:30]}>"


def read_tracker(xlsx_path):
    """返回 {sheet名: [Record,...]}，只包含能识别出表头的 sheet。"""
    wb = load_workbook(xlsx_path, data_only=True, read_only=True)
    data = {}
    for ws in wb.worksheets:
        hdr_row, cols = _find_header(ws)
        if not hdr_row or "desc" not in cols:
            continue
        recs = []
        for ridx, row in enumerate(ws.iter_rows(min_row=hdr_row + 1, values_only=True),
                                   start=hdr_row + 1):
            def g(k):
                i = cols.get(k)
                return row[i] if i is not None and i < len(row) else None
            desc = g("desc")
            if desc is None or str(desc).strip() == "":
                continue
            r = Record()
            r.sheet = ws.title
            r.row = ridx
            r.sno = g("sno")
            r.rev = (str(g("rev")).strip() if g("rev") is not None else "")
            r.desc = re.sub(r"\s+", " ", str(desc)).strip()
            r.submitted = to_date(g("submitted"))
            r.resp = to_date(g("resp"))
            st = g("status")
            r.status = str(st).strip() if st is not None else ""
            cm = g("comments")
            r.comments = str(cm).strip() if cm is not None else ""
            if r.status.lower() in ("open", "recall", "recalled"):
                r.code = None
            else:
                r.code = parse_code(r.status) or parse_code(r.comments)
            recs.append(r)
        data[ws.title] = recs
    wb.close()
    return data


def strip_prefix(desc, patterns):
    s = desc
    for p in patterns:
        new = re.sub(p, "", s, count=1, flags=re.I)
        if new != s:
            s = new
            break
    s = s.strip(" -–:")
    return s if s else desc


# ======================================================================
# 统计
# ======================================================================

class SheetStats:
    def __init__(self, sheet, recs, start, end, cfg):
        self.sheet = sheet
        pass_codes = {c.upper() for c in cfg["pass_codes"]}
        open_status = {s.lower() for s in cfg["open_status"]}
        self.submitted = [r for r in recs if r.submitted and start <= r.submitted <= end]
        resp_in = [r for r in recs if r.resp and start <= r.resp <= end]
        self.responded = [r for r in resp_in if r.code]
        self.unparsed = [r for r in resp_in if not r.code and r.status.lower() not in open_status]
        self.passed = [r for r in self.responded if r.code in pass_codes]
        self.rejected = [r for r in self.responded if r.code not in pass_codes]
        self.open = [r for r in recs if r.status.lower() in open_status]


def _get_stats(cache, data, sheet, start, end, cfg, warnings):
    if sheet not in cache:
        if sheet not in data:
            warnings.append(f"xlsx 中找不到 sheet “{sheet}”（或该 sheet 没有标准表头），按 0 处理。")
            cache[sheet] = SheetStats(sheet, [], start, end, cfg)
        else:
            cache[sheet] = SheetStats(sheet, data[sheet], start, end, cfg)
    return cache[sheet]


# ======================================================================
# PPT 工具函数
# ======================================================================

def _is_cjk(s):
    return any("\u2e80" <= c <= "\u9fff" or "\uff00" <= c <= "\uffef" for c in s)


def set_cell_lines(cell, lines, size_hundredths=None):
    """把单元格内容替换为多段文字，保留原有段落 / 字体格式。"""
    txb = cell.text_frame._txBody
    ps = txb.findall(A + "p")
    tpl = next((p for p in ps if p.find(A + "r") is not None), ps[0] if ps else None)
    tpl = copy.deepcopy(tpl)
    for p in ps:
        txb.remove(p)
    for ln in lines:
        p = copy.deepcopy(tpl)
        runs = p.findall(A + "r")
        if runs:
            for r in runs[1:]:
                p.remove(r)
            r = runs[0]
        else:  # 空单元格：按 endParaRPr 新建一个 run
            r = p.makeelement(A + "r", {})
            end = p.find(A + "endParaRPr")
            rpr = copy.deepcopy(end) if end is not None else p.makeelement(A + "rPr", {})
            rpr.tag = A + "rPr"
            r.append(rpr)
            t = r.makeelement(A + "t", {})
            r.append(t)
            if end is not None:
                end.addprevious(r)
            else:
                p.append(r)
        for br in p.findall(A + "br"):
            p.remove(br)
        r.find(A + "t").text = ln
        cjk = _is_cjk(ln)
        for e in (r.find(A + "rPr"), p.find(A + "endParaRPr")):
            if e is not None:
                e.set("lang", "zh-CN" if cjk else "en-US")
                e.set("altLang", "en-US" if cjk else "zh-CN")
                if size_hundredths:
                    e.set("sz", str(size_hundredths))
        txb.append(p)


def set_cell_size(cell, size_hundredths):
    for e in cell.text_frame._txBody.iter(A + "rPr", A + "endParaRPr"):
        e.set("sz", str(size_hundredths))


def cell_font_size(cell, default=1000):
    for e in cell.text_frame._txBody.iter(A + "rPr"):
        if e.get("sz"):
            return int(e.get("sz"))
    return default


def numbered(items):
    return [f"{i}. {t}" for i, t in enumerate(items, 1)] if items else ["无"]


def replace_in_paragraph(para, pattern, repl):
    """在段落内（可能跨多个 run）做一次正则替换，保留第一个 run 的格式。"""
    runs = list(para.runs)
    if not runs:
        return False
    texts = [r.text for r in runs]
    full = "".join(texts)
    m = re.search(pattern, full)
    if not m:
        return False
    start, end = m.span()
    new = repl(m) if callable(repl) else repl
    offs, o = [], 0
    for t in texts:
        offs.append(o)
        o += len(t)
    i0 = max(i for i in range(len(runs)) if offs[i] <= start)
    i1 = max(i for i in range(len(runs)) if offs[i] < end) if end > start else i0
    runs[i0].text = texts[i0][: start - offs[i0]] + new + texts[i1][end - offs[i1]:]
    for k in range(i0 + 1, i1 + 1):
        runs[k]._r.getparent().remove(runs[k]._r)
    return True


def _is_red(run):
    rpr = run._r.find(A + "rPr")
    return rpr is not None and rpr.find(A + "solidFill") is not None


def fill_slots(para, values):
    """把段落里红色数字位置依次替换成 values。返回实际找到的位置数。"""
    runs = list(para.runs)
    groups, cur = [], []
    for r in runs:
        if _is_red(r):
            cur.append(r)
        else:
            if cur:
                groups.append(cur)
            cur = []
    if cur:
        groups.append(cur)
    slots = []
    for g in groups:
        d = next((r for r in g if re.search(r"\d", r.text)), None)
        if d is None:
            d = next((r for r in g if r.text.strip() == ""), None)
        if d is not None:
            slots.append(d)
    for r, v in zip(slots, values):
        if re.search(r"\d", r.text):
            r.text = re.sub(r"\d+", str(v), r.text, count=1)
        else:
            r.text = f" {v} "
    return len(slots)


# ---------- 排版估算 ----------

def _text_width_pt(s, size_pt):
    w = 0.0
    for c in s:
        if _is_cjk(c):
            w += 1.0
        elif c == " ":
            w += 0.25
        elif c in "il.,;:|!'()[]":
            w += 0.3
        elif c.isupper() or c.isdigit():
            w += 0.6
        else:
            w += 0.48
    return w * size_pt


def _cell_margins(cell):
    tcPr = cell._tc.find(A + "tcPr")
    g = (lambda k, d: int(tcPr.get(k, d))) if tcPr is not None else (lambda k, d: d)
    return g("marL", 91440), g("marR", 91440), g("marT", 45720), g("marB", 45720)


def _cell_needed_height(cell, width_emu, size_pt):
    mL, mR, mT, mB = _cell_margins(cell)
    avail_pt = max(10, (width_emu - mL - mR) / 12700.0)
    lines = 0
    for p in cell.text_frame.paragraphs:
        t = "".join(r.text for r in p.runs)
        lines += max(1, math.ceil(_text_width_pt(t, size_pt) / avail_pt))
    line_h = size_pt * 1.45 * 12700
    return int(lines * line_h + mT + mB + 20000)


# ======================================================================
# 主流程
# ======================================================================

SECTION_RE = re.compile(r"^\s*(\d+\.\d+)\s*[、.．，,]")


def _slide_sections(slide):
    heads = []
    for sh in slide.shapes:
        if sh.has_text_frame:
            m = SECTION_RE.match(sh.text_frame.text)
            if m:
                heads.append((sh.top, m.group(1)))
    return sorted(heads)


def _section_for(top, heads):
    cands = [code for t, code in heads if t <= top + 60000]
    return cands[-1] if cands else None


def _match_type(text, mapping):
    k = _norm(text)
    for t, sheet in mapping.items():
        if _norm(t) == k:
            return sheet
    return None


def update_report(xlsx_path, pptx_path, start, end, out_path,
                  report_date=None, cfg=None, log=print):
    cfg = cfg or load_config()
    report_date = report_date or end
    warnings, info = [], []

    log(f"读取表格：{os.path.basename(xlsx_path)}")
    data = read_tracker(xlsx_path)
    log(f"  识别到 {len(data)} 个文件类型 sheet")

    prs = Presentation(pptx_path)
    cache = {}
    slide_results = []
    used_sheets_all = set()

    for sidx, slide in enumerate(prs.slides, start=1):
        heads = _slide_sections(slide)
        tables = [sh for sh in slide.shapes if getattr(sh, "has_table", False)]
        filled = []          # (shape, data_row_indices)
        slide_sheets = []    # 按出现顺序
        for gf in tables:
            sec = _section_for(gf.top, heads)
            if sec not in cfg["sections"]:
                continue
            scfg = cfg["sections"][sec]
            tbl = gf.table
            hdr = [_norm(c.text_frame.text) for c in tbl.rows[0].cells]
            type_col = next((i for i, h in enumerate(hdr) if h == "type"), None)
            sub_col = next((i for i, h in enumerate(hdr) if "submit" in h), None)
            app_col = next((i for i, h in enumerate(hdr) if "approv" in h), None)
            if sub_col is None and app_col is None:
                continue
            rows_done = []
            for ri in range(1, len(tbl.rows)):
                if type_col is not None and "type_column" in scfg:
                    ttext = tbl.cell(ri, type_col).text_frame.text.strip()
                    if not ttext:
                        continue
                    sheet = _match_type(ttext, scfg["type_column"])
                    if not sheet:
                        warnings.append(f"第{sidx}页 {sec} 表格 Type “{ttext}” 在配置里没有对应的 sheet，未更新该行。")
                        continue
                    sheets = [sheet]
                else:
                    sheets = scfg.get("sheets", [])
                stats = [_get_stats(cache, data, s, start, end, cfg, warnings) for s in sheets]
                pats = cfg["prefix_patterns"]
                if sub_col is not None:
                    items = [strip_prefix(r.desc, pats) for st in stats for r in st.submitted]
                    set_cell_lines(tbl.cell(ri, sub_col), numbered(items))
                if app_col is not None:
                    items = [strip_prefix(r.desc, pats) for st in stats for r in st.passed]
                    set_cell_lines(tbl.cell(ri, app_col), numbered(items))
                for s in sheets:
                    if s not in slide_sheets:
                        slide_sheets.append(s)
                rows_done.append(ri)
            if rows_done:
                filled.append((gf, rows_done))
            # 小节里配置了、但表格里没出现的 sheet 也计入本页统计
            for s in scfg.get("sheets", []) + list(scfg.get("type_column", {}).values()):
                if s not in slide_sheets:
                    slide_sheets.append(s)

        if not filled:
            continue
        used_sheets_all.update(slide_sheets)

        # ---- 本页统计 ----
        sts = [_get_stats(cache, data, s, start, end, cfg, warnings) for s in slide_sheets]
        tot = {
            "sub": sum(len(s.submitted) for s in sts),
            "resp": sum(len(s.responded) for s in sts),
            "pass": sum(len(s.passed) for s in sts),
            "rej": sum(len(s.rejected) for s in sts),
            "open": sum(len(s.open) for s in sts),
        }
        filled_summary = False
        for sh in slide.shapes:
            if not sh.has_text_frame or getattr(sh, "has_table", False):
                continue
            for para in sh.text_frame.paragraphs:
                t = "".join(r.text for r in para.runs)
                if "还未" in t:
                    vals = [tot["open"]]
                elif "通过" in t and "批复" in t:
                    vals = [tot["resp"], tot["pass"], tot["rej"]]
                elif "收到" in t and "批复" in t:
                    vals = [tot["resp"]]
                elif "提交" in t and "批复" not in t:
                    vals = [tot["sub"]]
                else:
                    continue
                n = fill_slots(para, vals)
                if n < len(vals):
                    warnings.append(f"第{sidx}页 “{t.strip()}” 里只找到 {n} 个红色数字位置（需要 {len(vals)} 个），请手动核对。")
                filled_summary = True
        if not filled_summary:
            warnings.append(f"第{sidx}页没有找到“提交 / 批复”统计文字。")

        # ---- 标题日期 ----
        for sh in slide.shapes:
            if sh.has_text_frame and "前几周工作完成情况" in sh.text_frame.text:
                for para in sh.text_frame.paragraphs:
                    replace_in_paragraph(
                        para,
                        r"[（(]\s*\d{1,2}\s*月\s*\d{1,2}\s*日\s*[–\-~—～至到]+\s*\d{1,2}\s*月\s*\d{1,2}\s*日\s*[）)]",
                        f"（{start.month}月{start.day:02d}日 – {end.month}月{end.day:02d}日）")
                if sh.width < 7000000:   # 加宽标题框，避免折行
                    sh.width = Emu(7200000)

        # ---- 排版：行高 / 推开下方元素 / 必要时缩小字号 ----
        _layout_slide(sidx, slide, filled, cfg, warnings, prs.slide_height)

        slide_results.append((sidx, slide_sheets, tot))

    # ---- 封面日期 ----
    if len(prs.slides):
        for sh in prs.slides[0].shapes:
            if sh.has_text_frame:
                for para in sh.text_frame.paragraphs:
                    replace_in_paragraph(
                        para, r"\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日",
                        f"{report_date.year}年{report_date.month:02d}月{report_date.day:02d}日")

    if not slide_results:
        raise RuntimeError("PPT 里没有找到 1.1–1.4 小节的表格，请确认选的是周例会汇报 PPT。")

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    prs.save(out_path)
    log(f"已生成：{out_path}")

    # ---- 未纳入 PPT 的 sheet 动态（供参考）----
    for sheet, recs in data.items():
        if sheet in used_sheets_all:
            continue
        st = SheetStats(sheet, recs, start, end, cfg)
        if st.submitted or st.responded:
            info.append((sheet, st))

    report = _write_checklist(out_path, start, end, cfg, cache, slide_results, warnings, info)
    log(f"核对清单：{report}")
    for sidx, sheets, tot in slide_results:
        log(f"  第{sidx}页（{' + '.join(sheets)}）：提交 {tot['sub']}，批复 {tot['resp']}，"
            f"通过 {tot['pass']}，被拒 {tot['rej']}，Open {tot['open']}")
    if warnings:
        log(f"⚠ 有 {len(warnings)} 条提醒，详见核对清单。")
    return out_path, report, warnings


def _layout_slide(sidx, slide, filled, cfg, warnings, slide_h):
    limit = cfg.get("slide_bottom_limit_emu", 4870000)
    min_sz = int(cfg.get("min_font_pt", 8) * 100)
    base_sz = cell_font_size(filled[0][0].table.cell(filled[0][1][0], 0))
    for gf, rows in filled:
        for ri in rows:
            for c in gf.table.rows[ri].cells:
                base_sz = max(base_sz, cell_font_size(c, base_sz))
                break

    # 参与排版的元素：表格和文本框（排除图片、背景、标题区以上的元素）
    first_top = min(gf.top for gf, _ in filled)
    movable = [sh for sh in slide.shapes
               if (getattr(sh, "has_table", False) or sh.has_text_frame)
               and sh.top >= first_top - 10000
               and (getattr(sh, "has_table", False) or sh.height < slide_h * 0.6)]
    orig = {sh.shape_id: (sh.left, sh.top, sh.width, sh.height) for sh in movable}
    filled_ids = {gf.shape_id: (gf, rows) for gf, rows in filled}
    orig_rows = {gf.shape_id: [r.height for r in gf.table.rows] for gf, _ in filled}

    def plan(sz):
        size_pt = sz / 100.0
        heights = {}
        for sid, (gf, rows) in filled_ids.items():
            tbl = gf.table
            cols = [c.width for c in tbl.columns]
            hs = list(orig_rows[sid])
            need = {}
            for ri in rows:
                need[ri] = max(260000, max(_cell_needed_height(tbl.cell(ri, ci), cols[ci], size_pt)
                                           for ci in range(len(cols))))
            other = sum(h for i, h in enumerate(hs) if i not in need)
            need_total = other + sum(need.values())
            orig_total = orig[sid][3]
            extra = max(0, orig_total - need_total)
            tot_need = sum(need.values()) or 1
            for ri, h in need.items():
                hs[ri] = int(h + extra * h / tot_need)
            heights[sid] = hs
        # 级联下推：保持元素原有的上下关系，被上方变高的表格挤到时往下移
        GAP, TOL = 30000, 40000
        pos, placed = {}, []
        bottom_ok = lambda sid, t, h: t + h <= max(limit, orig[sid][1] + orig[sid][3]) + TOL

        def collides(l, t, w, h, skip):
            return any(l < pl + pw and pl < l + w and t < pt + ph and pt < t + h
                       for (sid2, pl, pt, pw, ph, _ot, _ob) in placed if sid2 != skip)

        for sh in sorted(movable, key=lambda s: (orig[s.shape_id][1], orig[s.shape_id][0])):
            sid = sh.shape_id
            l, t, w, h = orig[sid]
            if sid in heights:
                h = sum(heights[sid])
            for (_s, pl, pt, pw, ph, ot, ob) in placed:
                if l < pl + pw and pl < l + w and orig[sid][1] >= ot:
                    gap = min(GAP, orig[sid][1] - ob)   # 原来就贴着/略重叠的保持原样
                    if t < pt + ph + gap:
                        t = pt + ph + gap
            # 文本框被挤出页面：试着挪到其它较矮的表格下方
            if not bottom_ok(sid, t, h) and sid not in heights:
                for gsid in sorted(heights, key=lambda k: pos[k][1] + pos[k][3] if k in pos else 1e12):
                    if gsid not in pos:
                        continue
                    gl, gt, gw, gh = pos[gsid]
                    nl, nt, nw = gl, gt + gh + GAP, min(w, gw)
                    if bottom_ok(sid, nt, h) and nt + h <= limit + TOL and not collides(nl, nt, nw, h, sid):
                        l, t, w = nl, nt, nw
                        break
            pos[sid] = (l, t, w, h)
            placed.append((sid, l, t, w, h, orig[sid][1], orig[sid][1] + orig[sid][3]))
        ok = all(bottom_ok(s.shape_id, pos[s.shape_id][1], pos[s.shape_id][3]) for s in movable)
        return heights, pos, ok

    sz = base_sz
    heights, pos, ok = plan(sz)
    while not ok and sz - 50 >= min_sz:
        sz -= 50
        heights, pos, ok = plan(sz)
    if sz != base_sz:
        for gf, rows in filled:
            for ri in rows:
                for c in gf.table.rows[ri].cells:
                    set_cell_size(c, sz)
        warnings.append(f"第{sidx}页内容较多，表格正文字号已从 {base_sz/100:g} 缩小到 {sz/100:g}。")
    if not ok:
        warnings.append(f"第{sidx}页内容太多，缩到最小字号仍可能超出页面，请打开 PPT 手动调整。")

    for sh in movable:
        l, t, w, h = pos[sh.shape_id]
        if sh.shape_id in heights:
            for r, hh in zip(sh.table.rows, heights[sh.shape_id]):
                r.height = Emu(hh)
            sh.height = Emu(h)
        if l != sh.left or w != sh.width:
            sh.left, sh.width = Emu(l), Emu(w)
        if t != sh.top:
            sh.top = Emu(t)


def _fmt_rec(r, with_code=False):
    d = r.resp if with_code else r.submitted
    s = f"    #{r.sno}  {r.rev:<4} {d.isoformat() if d else '':<10}  {r.desc}"
    if with_code:
        s += f"   →  Code-{r.code}"
    return s


def _write_checklist(out_path, start, end, cfg, cache, slide_results, warnings, info):
    path = os.path.splitext(out_path)[0] + "_核对清单.txt"
    L = []
    L.append(f"周报数据核对清单    统计范围：{start.isoformat()} ~ {end.isoformat()}")
    L.append(f"生成时间：{dt.datetime.now():%Y-%m-%d %H:%M}")
    L.append(f"通过 = Code-{'/'.join(cfg['pass_codes'])}；其它代码算被拒。")
    L.append("=" * 70)
    if warnings:
        L.append("【需要注意】")
        for w in warnings:
            L.append("  ⚠ " + w)
        L.append("=" * 70)
    for sidx, sheets, tot in slide_results:
        L.append(f"第{sidx}页  合计：提交 {tot['sub']}，批复 {tot['resp']}，通过 {tot['pass']}，"
                 f"被拒 {tot['rej']}，未批复(Open) {tot['open']}")
        for s in sheets:
            st = cache[s]
            L.append(f"  ▶ {s}：提交 {len(st.submitted)}，批复 {len(st.responded)}"
                     f"（通过 {len(st.passed)} / 被拒 {len(st.rejected)}），Open {len(st.open)}")
            if st.submitted:
                L.append("    [提交]")
                L += [_fmt_rec(r) for r in st.submitted]
            if st.responded:
                L.append("    [批复]")
                L += [_fmt_rec(r, True) for r in st.responded]
            if st.unparsed:
                L.append("    [有回复日期但识别不出代码，未计入]")
                L += [_fmt_rec(r) + f"   状态：{r.status or '(空)'} / {r.comments}" for r in st.unparsed]
        L.append("-" * 70)
    if info:
        L.append("【供参考】以下 sheet 不在 PPT 里，但统计范围内有提交或批复：")
        for sheet, st in info:
            L.append(f"  ▶ {sheet}：提交 {len(st.submitted)}，批复 {len(st.responded)}")
            L += [_fmt_rec(r) + (f"   ({r.status})" if r.status else "") for r in st.submitted]
            L += [_fmt_rec(r, True) for r in st.responded]
    with open(path, "w", encoding="utf-8-sig") as f:
        f.write("\n".join(L) + "\n")
    return path
