# -*- coding: utf-8 -*-
"""
周更 —— 每周固定更新工作的小工具

首页选择功能：
  1. 更新周报：从 M&E Documents Status Tracker 提取日期范围数据，更新技术部周例会 PPT
  2. 更新 M&E 台账：用 M&E Documents Status Tracker 更新 Code-G 表和 Smart City 表

命令行（可选）：
  周更.exe report --xlsx 台账.xlsx --ppt 上周.pptx --start 2026-10-14 --end 2026-11-03
  周更.exe me --xlsx 台账.xlsx [--codeg Code-G.xlsx] [--smartcity SmartCity.xlsx]
"""
import argparse
import datetime as dt
import fnmatch
import json
import os
import queue
import re
import sys
import threading
import traceback

import me_update as mu
import report_core as rc

APP_NAME = "周更"
VERSION = "2.0"


# ======================================================================
# 公共工具
# ======================================================================

def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def config_path():
    p = os.path.join(app_dir(), "config.json")
    src = os.path.join(getattr(sys, "_MEIPASS", app_dir()), "config.json")
    try:
        if not os.path.isfile(p):
            if os.path.isfile(src) and os.path.abspath(src) != os.path.abspath(p):
                with open(src, encoding="utf-8") as f, open(p, "w", encoding="utf-8") as g:
                    g.write(f.read())
        elif os.path.isfile(src) and os.path.abspath(src) != os.path.abspath(p):
            # 旧版 config.json 没有 me_update 段：自动补上，保留用户已改的内容
            with open(p, encoding="utf-8") as f:
                user = json.load(f)
            if "me_update" not in user:
                with open(src, encoding="utf-8") as f:
                    default = json.load(f)
                user["me_update"] = default.get("me_update", {})
                with open(p, "w", encoding="utf-8") as g:
                    json.dump(user, g, ensure_ascii=False, indent=2)
    except (OSError, ValueError):
        pass
    return p


def load_full_config():
    cfg = rc.load_config(config_path())
    me = dict(mu.DEFAULT_ME_CONFIG)
    try:
        with open(config_path(), encoding="utf-8") as f:
            raw = json.load(f)
        for k, v in raw.get("me_update", {}).items():
            if not k.startswith("_"):
                me[k] = v
    except (OSError, ValueError):
        pass
    return cfg, me


def settings_path():
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    d = os.path.join(base, "WeeklyReportUpdater")
    os.makedirs(d, exist_ok=True)
    return os.path.join(d, "settings.json")


def load_settings():
    try:
        with open(settings_path(), encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def save_settings(s):
    try:
        with open(settings_path(), "w", encoding="utf-8") as f:
            json.dump(s, f, ensure_ascii=False, indent=2)
    except OSError:
        pass


def parse_date(text, default_year=None):
    """支持 2026-10-14 / 2026.10.14 / 20261014 / 10-14 / 10/14 / 10月14日"""
    t = str(text).strip()
    if not t:
        raise ValueError("日期为空")
    default_year = default_year or dt.date.today().year
    m = re.fullmatch(r"(\d{4})[-/.年]?(\d{1,2})[-/.月]?(\d{1,2})日?", t)
    if m:
        return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    m = re.fullmatch(r"(\d{1,2})[-/.月](\d{1,2})日?", t)
    if m:
        return dt.date(default_year, int(m.group(1)), int(m.group(2)))
    raise ValueError(f"看不懂的日期：{t}（请写成 2026-10-14 或 10-14）")


def newest_file(folder, pattern):
    """文件夹里符合名字规则、修改时间最新的 xlsx（忽略 ~$ 临时文件）。"""
    if not folder or not os.path.isdir(folder):
        return None
    hits = [os.path.join(folder, f) for f in os.listdir(folder)
            if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$")
            and fnmatch.fnmatch(f.lower(), pattern.lower())]
    return max(hits, key=os.path.getmtime) if hits else None


def fmt_mtime(path):
    return dt.datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d %H:%M")


def resolve_out_dir(out_dir, cfg, fallback_dir):
    d = out_dir or cfg.get("default_output_dir") or fallback_dir
    if not os.path.isdir(d):
        try:
            os.makedirs(d, exist_ok=True)
        except OSError:
            d = fallback_dir
    return d


def report_output_path(cfg, out_dir, report_date, fallback_dir):
    d = resolve_out_dir(out_dir, cfg, fallback_dir)
    name = cfg.get("output_name", "{date} 技术部周例会汇报材料.pptx").format(date=report_date.isoformat())
    return os.path.join(d, name)


def split_names(text):
    return [x.strip() for x in re.split(r"[,，;；、\s]+", text or "") if x.strip()]


def run_me_update(tracker, codeg, smartcity, out_dir, me_cfg, clear_old, log):
    log(f"读取 M&E 台账：{os.path.basename(tracker)}（最后修改 {fmt_mtime(tracker)}）")
    me = mu.load_me(tracker)
    log(f"  共 {len(me[0])} 条文件记录")
    outs, notes = [], 0
    if codeg:
        log(f"更新 Code-G：{os.path.basename(codeg)}")
        out = mu.dated_name(codeg, out_dir)
        r = mu.update_codeg(codeg, me, out, me_cfg, log=log, clear_old=clear_old)
        outs.append(out)
        notes += len(r["notes"])
    if smartcity:
        log(f"更新 Smart City：{os.path.basename(smartcity)}")
        out = mu.dated_name(smartcity, out_dir)
        r = mu.update_smartcity(smartcity, me, out, me_cfg, log=log, clear_old=clear_old)
        outs.append(out)
    for o in outs:
        log(f"已生成：{o}")
    log("改动的单元格已标黄；每份表最后的「更新记录」工作表逐格列出原值和新值。")
    return outs


# ======================================================================
# 命令行
# ======================================================================

def run_cli(argv):
    ap = argparse.ArgumentParser(description=APP_NAME)
    sub = ap.add_subparsers(dest="cmd")
    a1 = sub.add_parser("report", help="更新周报 PPT")
    a1.add_argument("--xlsx", required=True)
    a1.add_argument("--ppt", required=True)
    a1.add_argument("--start", required=True)
    a1.add_argument("--end", required=True)
    a1.add_argument("--date")
    a1.add_argument("--out")
    a2 = sub.add_parser("me", help="更新 M&E 台账（Code-G / Smart City）")
    a2.add_argument("--xlsx", required=True, help="M&E Documents Status Tracker")
    a2.add_argument("--codeg")
    a2.add_argument("--smartcity")
    a2.add_argument("--people", help="Code-G 负责人，逗号分隔，默认用 config.json")
    a2.add_argument("--types", help="Code-G 文件类型，逗号分隔，默认用 config.json")
    a2.add_argument("--out")
    a2.add_argument("--keep-old-marks", action="store_true", help="保留上次的黄色标记")
    # 兼容旧版：直接写 --xlsx ... 视为 report
    if argv and argv[0].startswith("--"):
        argv = ["report"] + argv
    a = ap.parse_args(argv)
    cfg, me_cfg = load_full_config()
    if a.cmd == "report":
        start, end = parse_date(a.start), parse_date(a.end)
        rdate = parse_date(a.date) if a.date else end
        out = report_output_path(cfg, a.out, rdate, os.path.dirname(os.path.abspath(a.ppt)))
        print(f"使用台账：{a.xlsx}（最后修改 {fmt_mtime(a.xlsx)}）")
        rc.update_report(a.xlsx, a.ppt, start, end, out, report_date=rdate, cfg=cfg)
    elif a.cmd == "me":
        if a.people is not None:
            me_cfg["codeg_people"] = split_names(a.people)
        if a.types is not None:
            me_cfg["codeg_types"] = split_names(a.types)
        src = a.codeg or a.smartcity or a.xlsx
        out_dir = resolve_out_dir(a.out, cfg, os.path.dirname(os.path.abspath(src)))
        run_me_update(a.xlsx, a.codeg, a.smartcity, out_dir, me_cfg, not a.keep_old_marks, print)
    else:
        ap.print_help()


# ======================================================================
# 图形界面
# ======================================================================

def run_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    st = load_settings()
    cfg0, me0 = load_full_config()

    root = tk.Tk()
    root.title(f"{APP_NAME}  v{VERSION}")
    root.geometry("800x600")
    root.minsize(680, 520)

    style = ttk.Style()
    try:
        style.theme_use("vista" if sys.platform == "win32" else "clam")
    except tk.TclError:
        pass
    base_font = ("Microsoft YaHei UI", 10)
    style.configure(".", font=base_font)
    style.configure("Title.TLabel", font=("Microsoft YaHei UI", 22, "bold"))
    style.configure("Sub.TLabel", font=("Microsoft YaHei UI", 10), foreground="#666")
    style.configure("Card.TButton", font=("Microsoft YaHei UI", 13, "bold"), padding=(20, 18))
    style.configure("Back.TButton", padding=(6, 2))
    style.configure("Head.TLabel", font=("Microsoft YaHei UI", 14, "bold"))

    container = ttk.Frame(root)
    container.pack(fill="both", expand=True)
    pages = {}

    def show(name):
        for p in pages.values():
            p.pack_forget()
        pages[name].pack(fill="both", expand=True)
        root.title(f"{APP_NAME}  v{VERSION}" + ("" if name == "home" else f"  —  {PAGE_TITLES[name]}"))

    PAGE_TITLES = {"report": "更新周报", "me": "更新 M&E 台账"}

    # ---------------- 公共控件 ----------------
    def pick_file(var, title, types):
        init = os.path.dirname(var.get()) if var.get() else ""
        p = filedialog.askopenfilename(title=title, filetypes=types, initialdir=init or None)
        if p:
            var.set(os.path.normpath(p))

    def pick_dir(var):
        p = filedialog.askdirectory(title="选择输出文件夹", initialdir=var.get() or None)
        if p:
            var.set(os.path.normpath(p))

    def file_row(parent, r, label, var, title, types, check_var=None):
        if check_var is not None:
            ttk.Checkbutton(parent, text=label, variable=check_var).grid(row=r, column=0, sticky="w", pady=4)
        else:
            ttk.Label(parent, text=label).grid(row=r, column=0, sticky="w", pady=4)
        ttk.Entry(parent, textvariable=var).grid(row=r, column=1, sticky="ew", pady=4)
        ttk.Button(parent, text="选择…", command=lambda: pick_file(var, title, types)).grid(row=r, column=2, padx=(6, 0))

    XLSX = [("Excel", "*.xlsx *.xlsm"), ("全部", "*.*")]

    def open_path(p):
        if p and hasattr(os, "startfile"):
            os.startfile(p)

    def make_log(parent, r):
        box = tk.Text(parent, height=14, wrap="word", font=("Microsoft YaHei UI", 9))
        box.grid(row=r, column=0, columnspan=3, sticky="nsew", pady=(6, 0))
        parent.rowconfigure(r, weight=1)
        sb = ttk.Scrollbar(parent, command=box.yview)
        sb.grid(row=r, column=3, sticky="ns", pady=(6, 0))
        box.configure(yscrollcommand=sb.set)
        return box

    def header(parent, title, desc):
        top = ttk.Frame(parent)
        top.grid(row=0, column=0, columnspan=4, sticky="ew", pady=(0, 10))
        ttk.Button(top, text="← 返回首页", style="Back.TButton", command=lambda: show("home")).pack(side="left")
        ttk.Label(top, text=title, style="Head.TLabel").pack(side="left", padx=12)
        ttk.Label(top, text=desc, style="Sub.TLabel").pack(side="left")

    q = queue.Queue()

    def pump():
        try:
            while True:
                fn = q.get_nowait()
                fn()
        except queue.Empty:
            pass
        root.after(100, pump)

    def ui(fn):
        q.put(fn)

    def tracker_default():
        saved = st.get("xlsx", "")
        auto = newest_file(os.path.dirname(saved), cfg0.get("tracker_pattern", "*Status*Tracker*.xlsx")) if saved else None
        return auto or saved, auto

    def fresh_check(path, key):
        last = st.get(key)
        if last and os.path.isfile(path) and os.path.getmtime(path) <= last:
            return messagebox.askyesno(APP_NAME,
                f"这个 M&E 台账自上次运行以来没有修改过：\n{os.path.basename(path)}\n"
                f"（最后修改：{fmt_mtime(path)}）\n\n"
                "如果这周已经更新过台账，可能是：\n· 在 Excel 里改了但还没保存；\n· 另存成了别的文件名或放在别的文件夹。\n\n"
                "仍然用这个文件继续吗？")
        return True

    # ================= 首页 =================
    home = ttk.Frame(container, padding=30)
    pages["home"] = home
    ttk.Label(home, text="周更", style="Title.TLabel").pack(pady=(30, 4))
    ttk.Label(home, text="每周固定更新工作，选择要做的事：", style="Sub.TLabel").pack(pady=(0, 28))
    cards = ttk.Frame(home)
    cards.pack()
    ttk.Button(cards, text="更新周报\n\n技术部周例会 PPT", style="Card.TButton", width=20,
               command=lambda: show("report")).grid(row=0, column=0, padx=14, ipady=10)
    ttk.Button(cards, text="更新 M&E 台账\n\nCode-G 表 / Smart City 表", style="Card.TButton", width=20,
               command=lambda: show("me")).grid(row=0, column=1, padx=14, ipady=10)
    ttk.Label(home, text="两个功能共用同一份 M&E Documents Status Tracker，程序会自动选中文件夹里最新的一份。",
              style="Sub.TLabel").pack(pady=(30, 0))
    ttk.Button(home, text="编辑规则 (config.json)", command=lambda: open_path(config_path())).pack(pady=(12, 0))

    # ================= 更新周报 =================
    rp = ttk.Frame(container, padding=12)
    pages["report"] = rp
    rp.columnconfigure(1, weight=1)
    header(rp, "更新周报", "从 M&E 台账提取日期范围数据，更新周例会 PPT")

    today = dt.date.today()
    try:
        def_start = dt.date.fromisoformat(st.get("last_end", "")) + dt.timedelta(days=1)
    except ValueError:
        def_start = today - dt.timedelta(days=6)
    tr_path, tr_auto = tracker_default()
    r_xlsx = tk.StringVar(value=tr_path)
    r_ppt = tk.StringVar(value=st.get("last_output") or st.get("ppt", ""))
    r_start = tk.StringVar(value=def_start.isoformat())
    r_end = tk.StringVar(value=today.isoformat())
    r_date = tk.StringVar(value="")
    r_out = tk.StringVar(value=st.get("out_dir") or cfg0.get("default_output_dir", ""))

    file_row(rp, 1, "M&E 台账 (xlsx)：", r_xlsx, "选择 M&E Documents Status Tracker", XLSX)
    file_row(rp, 2, "上周的 PPT：", r_ppt, "选择上周的周例会 PPT", [("PowerPoint", "*.pptx"), ("全部", "*.*")])
    dates = ttk.Frame(rp)
    dates.grid(row=3, column=0, columnspan=3, sticky="w", pady=8)
    ttk.Label(dates, text="统计范围：").pack(side="left")
    ttk.Entry(dates, textvariable=r_start, width=12).pack(side="left")
    ttk.Label(dates, text=" 至 ").pack(side="left")
    ttk.Entry(dates, textvariable=r_end, width=12).pack(side="left")
    ttk.Label(dates, text="    封面日期：").pack(side="left")
    ttk.Entry(dates, textvariable=r_date, width=12).pack(side="left")
    ttk.Label(dates, text="（留空 = 结束日期）", style="Sub.TLabel").pack(side="left")
    ttk.Label(rp, text="输出文件夹：").grid(row=4, column=0, sticky="w", pady=4)
    ttk.Entry(rp, textvariable=r_out).grid(row=4, column=1, sticky="ew", pady=4)
    ttk.Button(rp, text="选择…", command=lambda: pick_dir(r_out)).grid(row=4, column=2, padx=(6, 0))
    rb = ttk.Frame(rp)
    rb.grid(row=5, column=0, columnspan=3, sticky="w", pady=(10, 6))
    r_run = ttk.Button(rb, text="生成新 PPT")
    r_run.pack(side="left")
    r_state = {"out": None}
    ttk.Button(rb, text="打开输出文件夹",
               command=lambda: open_path(os.path.dirname(r_state["out"]) if r_state["out"] else r_out.get())
               ).pack(side="left", padx=8)
    r_log = make_log(rp, 6)

    def rlog(msg):
        ui(lambda: (r_log.insert("end", str(msg) + "\n"), r_log.see("end")))

    def on_report():
        try:
            xlsx, ppt = r_xlsx.get().strip().strip('"'), r_ppt.get().strip().strip('"')
            if not os.path.isfile(xlsx):
                raise ValueError("请选择 M&E 台账 xlsx")
            if not os.path.isfile(ppt):
                raise ValueError("请选择上周的 PPT")
            start, end = parse_date(r_start.get()), parse_date(r_end.get())
            if start > end:
                raise ValueError("开始日期晚于结束日期")
            rdate = parse_date(r_date.get()) if r_date.get().strip() else end
            cfg, _ = load_full_config()
            out = report_output_path(cfg, r_out.get().strip(), rdate, os.path.dirname(ppt))
            if os.path.abspath(out) == os.path.abspath(ppt):
                base, ext = os.path.splitext(out)
                out = base + "_新" + ext
        except Exception as e:  # noqa
            messagebox.showerror(APP_NAME, str(e))
            return
        if not fresh_check(xlsx, "last_run_time"):
            return
        r_log.delete("1.0", "end")
        r_run.configure(state="disabled")
        rlog(f"使用台账：{xlsx}\n  最后修改时间：{fmt_mtime(xlsx)}")

        def work():
            try:
                res = rc.update_report(xlsx, ppt, start, end, out, report_date=rdate, cfg=cfg, log=rlog)
                st.update({"xlsx": xlsx, "ppt": ppt, "last_output": res[0], "last_end": end.isoformat(),
                           "out_dir": r_out.get().strip(), "last_run_time": dt.datetime.now().timestamp()})
                save_settings(st)

                def done():
                    r_run.configure(state="normal")
                    r_state["out"] = res[0]
                    msg = f"已生成：\n{res[0]}\n\n核对清单：\n{res[1]}"
                    if res[2]:
                        msg += f"\n\n有 {len(res[2])} 条提醒，请看核对清单。"
                    if messagebox.askyesno(APP_NAME, msg + "\n\n现在打开 PPT 吗？"):
                        open_path(res[0])
                ui(done)
            except PermissionError:
                ui(lambda: (r_run.configure(state="normal"),
                            messagebox.showerror(APP_NAME, f"无法写入：{out}\n文件可能正在 PowerPoint 里打开，请先关闭再试。")))
            except Exception as e:  # noqa
                rlog(traceback.format_exc())
                err = str(e)
                ui(lambda: (r_run.configure(state="normal"), messagebox.showerror(APP_NAME, f"出错了：{err}")))

        threading.Thread(target=work, daemon=True).start()

    r_run.configure(command=on_report)
    r_log.insert("end", "选择 M&E 台账和上周的 PPT，确认日期，点“生成新 PPT”。\n"
                        "会更新 1.1–1.4 的表格、每页提交/批复统计、标题和封面日期；第 5、6 页不改。\n")
    if tr_auto and os.path.normcase(tr_auto) != os.path.normcase(st.get("xlsx", "")):
        r_log.insert("end", f"已自动选中最新的台账：{os.path.basename(tr_auto)}（{fmt_mtime(tr_auto)}）\n")

    # ================= 更新 M&E 台账 =================
    mp = ttk.Frame(container, padding=12)
    pages["me"] = mp
    mp.columnconfigure(1, weight=1)
    header(mp, "更新 M&E 台账", "用 M&E 台账更新 Code-G 表和 Smart City 表，改动处标黄")

    def auto_pick(key, pattern):
        saved = st.get(key, "")
        auto = newest_file(os.path.dirname(saved), pattern) if saved else None
        return auto or saved

    m_xlsx = tk.StringVar(value=tr_path)
    m_codeg = tk.StringVar(value=auto_pick("me_codeg", me0.get("codeg_name_pattern", "*Code G*.xlsx")))
    m_sc = tk.StringVar(value=auto_pick("me_smartcity", me0.get("smartcity_name_pattern", "*Smart City*.xlsx")))
    m_do_codeg = tk.BooleanVar(value=st.get("me_do_codeg", True))
    m_do_sc = tk.BooleanVar(value=st.get("me_do_sc", True))
    m_people = tk.StringVar(value=st.get("me_people", "、".join(me0.get("codeg_people", []))))
    m_types = tk.StringVar(value=st.get("me_types", "、".join(me0.get("codeg_types", []))))
    m_clear = tk.BooleanVar(value=st.get("me_clear", True))
    m_out = tk.StringVar(value=st.get("me_out_dir") or st.get("out_dir") or cfg0.get("default_output_dir", ""))

    file_row(mp, 1, "M&E 台账 (xlsx)：", m_xlsx, "选择 M&E Documents Status Tracker", XLSX)
    file_row(mp, 2, "Code-G 表：", m_codeg, "选择 Transmittals List-Code G", XLSX, m_do_codeg)
    flt = ttk.Frame(mp)
    flt.grid(row=3, column=1, columnspan=2, sticky="w", pady=(0, 6))
    ttk.Label(flt, text="只更新  负责人：").pack(side="left")
    ttk.Entry(flt, textvariable=m_people, width=16).pack(side="left")
    ttk.Label(flt, text="   文件类型：").pack(side="left")
    ttk.Entry(flt, textvariable=m_types, width=12).pack(side="left")
    ttk.Label(flt, text="（多个用顿号隔开，空 = 全部）", style="Sub.TLabel").pack(side="left")
    file_row(mp, 4, "Smart City 表：", m_sc, "选择 Smart City Related Document Status", XLSX, m_do_sc)
    ttk.Checkbutton(mp, text="清除上次更新留下的黄色标记（恢复原底色，只标本次改动）",
                    variable=m_clear).grid(row=5, column=1, columnspan=2, sticky="w", pady=4)
    ttk.Label(mp, text="输出文件夹：").grid(row=6, column=0, sticky="w", pady=4)
    ttk.Entry(mp, textvariable=m_out).grid(row=6, column=1, sticky="ew", pady=4)
    ttk.Button(mp, text="选择…", command=lambda: pick_dir(m_out)).grid(row=6, column=2, padx=(6, 0))
    mb = ttk.Frame(mp)
    mb.grid(row=7, column=0, columnspan=3, sticky="w", pady=(10, 6))
    m_run = ttk.Button(mb, text="开始更新")
    m_run.pack(side="left")
    m_state = {"outs": []}
    ttk.Button(mb, text="打开输出文件夹",
               command=lambda: open_path(os.path.dirname(m_state["outs"][0]) if m_state["outs"] else m_out.get())
               ).pack(side="left", padx=8)
    m_log = make_log(mp, 8)

    def mlog(msg):
        ui(lambda: (m_log.insert("end", str(msg) + "\n"), m_log.see("end")))

    def on_me():
        try:
            xlsx = m_xlsx.get().strip().strip('"')
            codeg = m_codeg.get().strip().strip('"') if m_do_codeg.get() else ""
            sc = m_sc.get().strip().strip('"') if m_do_sc.get() else ""
            if not os.path.isfile(xlsx):
                raise ValueError("请选择 M&E 台账 xlsx")
            if not codeg and not sc:
                raise ValueError("请至少勾选 Code-G 表或 Smart City 表")
            for p, n in ((codeg, "Code-G 表"), (sc, "Smart City 表")):
                if p and not os.path.isfile(p):
                    raise ValueError(f"找不到{n}：{p}")
            _, me_cfg = load_full_config()
            me_cfg["codeg_people"] = split_names(m_people.get())
            me_cfg["codeg_types"] = [t.upper() for t in split_names(m_types.get())]
            out_dir = resolve_out_dir(m_out.get().strip(), cfg0, os.path.dirname(codeg or sc))
        except Exception as e:  # noqa
            messagebox.showerror(APP_NAME, str(e))
            return
        if not fresh_check(xlsx, "me_last_run_time"):
            return
        m_log.delete("1.0", "end")
        m_run.configure(state="disabled")
        clear_old = m_clear.get()

        def work():
            try:
                outs = run_me_update(xlsx, codeg, sc, out_dir, me_cfg, clear_old, mlog)
                st.update({"xlsx": xlsx, "me_codeg": outs[0] if codeg else st.get("me_codeg", ""),
                           "me_smartcity": outs[-1] if sc else st.get("me_smartcity", ""),
                           "me_do_codeg": bool(codeg), "me_do_sc": bool(sc),
                           "me_people": m_people.get(), "me_types": m_types.get(),
                           "me_clear": clear_old, "me_out_dir": m_out.get().strip(),
                           "me_last_run_time": dt.datetime.now().timestamp()})
                save_settings(st)

                def done():
                    m_run.configure(state="normal")
                    m_state["outs"] = outs
                    if messagebox.askyesno(APP_NAME, "已生成：\n" + "\n".join(outs) + "\n\n现在打开输出文件夹吗？"):
                        open_path(os.path.dirname(outs[0]))
                ui(done)
            except PermissionError as e:
                fn = getattr(e, "filename", "") or ""
                ui(lambda: (m_run.configure(state="normal"),
                            messagebox.showerror(APP_NAME, f"无法写入 {fn}\n文件可能正在 Excel 里打开，请先关闭再试。")))
            except Exception as e:  # noqa
                mlog(traceback.format_exc())
                err = str(e)
                ui(lambda: (m_run.configure(state="normal"), messagebox.showerror(APP_NAME, f"出错了：{err}")))

        threading.Thread(target=work, daemon=True).start()

    m_run.configure(command=on_me)
    m_log.insert("end",
                 "选择 M&E 台账，以及要更新的 Code-G 表 / Smart City 表，点“开始更新”。\n"
                 "· 按文件编号找 M&E 里同一文件的最新版本；有新 Transmittal 或状态变了才改。\n"
                 "· Code-G：只改上面填的负责人和文件类型的行；“状态”列 A/B/C/D/O 填“已完成”，其余填“已上报”。\n"
                 "· 新文件保存到输出文件夹，文件名带今天日期，原文件不动。\n"
                 "· 程序会记住这次的输出，下次自动作为输入。\n")

    show("home")
    root.after(100, pump)
    root.mainloop()


if __name__ == "__main__":
    if len(sys.argv) > 1:
        run_cli(sys.argv[1:])
    else:
        run_gui()
