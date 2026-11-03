# -*- coding: utf-8 -*-
"""
周报更新工具 —— 图形界面 / 命令行入口

双击运行：打开窗口，选择 Status Tracker 表格和上周的 PPT，填日期，点“生成”。
命令行：  WeeklyReportUpdater.exe --xlsx 表格.xlsx --ppt 上周.pptx --start 2026-10-14 --end 2026-11-03
"""
import argparse
import datetime as dt
import json
import os
import queue
import re
import sys
import threading
import traceback

import report_core as rc

APP_NAME = "周报更新工具"


def app_dir():
    """exe 所在目录（打包后）或脚本所在目录。config.json 放这里。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def config_path():
    p = os.path.join(app_dir(), "config.json")
    if not os.path.isfile(p):
        # 第一次运行：把默认配置写出来，方便以后修改
        try:
            src = os.path.join(getattr(sys, "_MEIPASS", app_dir()), "config.json")
            if os.path.isfile(src) and os.path.abspath(src) != os.path.abspath(p):
                with open(src, encoding="utf-8") as f, open(p, "w", encoding="utf-8") as g:
                    g.write(f.read())
        except OSError:
            pass
    return p


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


def newest_tracker(folder, pattern="*Status*Tracker*.xlsx"):
    """在文件夹里找修改时间最新的台账 xlsx（忽略 Excel 临时文件 ~$）。"""
    import fnmatch
    if not folder or not os.path.isdir(folder):
        return None
    files = [os.path.join(folder, f) for f in os.listdir(folder)
             if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$")]
    hits = [f for f in files if fnmatch.fnmatch(os.path.basename(f).lower(), pattern.lower())]
    cands = hits or []
    return max(cands, key=os.path.getmtime) if cands else None


def fmt_mtime(path):
    return dt.datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d %H:%M")


def output_path(cfg, out_dir, report_date, fallback_dir):
    d = out_dir or cfg.get("default_output_dir") or fallback_dir
    if not os.path.isdir(d):
        try:
            os.makedirs(d, exist_ok=True)
        except OSError:
            d = fallback_dir
    name = cfg.get("output_name", "{date} 技术部周例会汇报材料.pptx").format(date=report_date.isoformat())
    return os.path.join(d, name)


# ----------------------------------------------------------------------
# 命令行
# ----------------------------------------------------------------------

def run_cli(argv):
    ap = argparse.ArgumentParser(description=APP_NAME)
    ap.add_argument("--xlsx", required=True, help="M&E Documents Status Tracker 表格")
    ap.add_argument("--ppt", required=True, help="上周的周例会 PPT（作为模板）")
    ap.add_argument("--start", required=True, help="统计开始日期，如 2026-10-14")
    ap.add_argument("--end", required=True, help="统计结束日期，如 2026-11-03")
    ap.add_argument("--date", help="封面汇报日期，默认等于结束日期")
    ap.add_argument("--out", help="输出文件夹，默认用 config.json 里的 default_output_dir")
    a = ap.parse_args(argv)
    cfg = rc.load_config(config_path())
    start, end = parse_date(a.start), parse_date(a.end)
    rdate = parse_date(a.date) if a.date else end
    out = output_path(cfg, a.out, rdate, os.path.dirname(os.path.abspath(a.ppt)))
    print(f"使用台账：{a.xlsx}（最后修改 {fmt_mtime(a.xlsx)}）")
    rc.update_report(a.xlsx, a.ppt, start, end, out, report_date=rdate, cfg=cfg)


# ----------------------------------------------------------------------
# 图形界面
# ----------------------------------------------------------------------

def run_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    st = load_settings()
    today = dt.date.today()
    try:
        last_end = dt.date.fromisoformat(st.get("last_end", ""))
        def_start = last_end + dt.timedelta(days=1)
    except ValueError:
        def_start = today - dt.timedelta(days=6)

    root = tk.Tk()
    root.title(APP_NAME)
    root.geometry("760x560")
    root.minsize(640, 480)
    try:
        root.tk.call("tk", "scaling", 1.3)
    except tk.TclError:
        pass

    saved_xlsx = st.get("xlsx", "")
    pattern = rc.load_config(config_path()).get("tracker_pattern", "*Status*Tracker*.xlsx")
    auto_xlsx = newest_tracker(os.path.dirname(saved_xlsx), pattern) if saved_xlsx else None
    v_xlsx = tk.StringVar(value=auto_xlsx or saved_xlsx)
    v_ppt = tk.StringVar(value=st.get("last_output") or st.get("ppt", ""))
    v_start = tk.StringVar(value=def_start.isoformat())
    v_end = tk.StringVar(value=today.isoformat())
    v_date = tk.StringVar(value="")
    cfg0 = rc.load_config(config_path())
    v_out = tk.StringVar(value=st.get("out_dir") or cfg0.get("default_output_dir", ""))

    frm = ttk.Frame(root, padding=12)
    frm.pack(fill="both", expand=True)
    frm.columnconfigure(1, weight=1)

    def pick_file(var, title, types):
        init = os.path.dirname(var.get()) if var.get() else ""
        p = filedialog.askopenfilename(title=title, filetypes=types, initialdir=init or None)
        if p:
            var.set(os.path.normpath(p))

    def pick_dir(var):
        p = filedialog.askdirectory(title="选择输出文件夹", initialdir=var.get() or None)
        if p:
            var.set(os.path.normpath(p))

    r = 0
    ttk.Label(frm, text="文件台账 (xlsx)：").grid(row=r, column=0, sticky="w", pady=4)
    ttk.Entry(frm, textvariable=v_xlsx).grid(row=r, column=1, sticky="ew", pady=4)
    ttk.Button(frm, text="选择…", command=lambda: pick_file(
        v_xlsx, "选择 M&E Documents Status Tracker", [("Excel", "*.xlsx *.xlsm"), ("全部", "*.*")])
    ).grid(row=r, column=2, padx=(6, 0))
    r += 1
    ttk.Label(frm, text="上周的 PPT：").grid(row=r, column=0, sticky="w", pady=4)
    ttk.Entry(frm, textvariable=v_ppt).grid(row=r, column=1, sticky="ew", pady=4)
    ttk.Button(frm, text="选择…", command=lambda: pick_file(
        v_ppt, "选择上周的周例会 PPT", [("PowerPoint", "*.pptx"), ("全部", "*.*")])
    ).grid(row=r, column=2, padx=(6, 0))
    r += 1

    dates = ttk.Frame(frm)
    dates.grid(row=r, column=0, columnspan=3, sticky="w", pady=8)
    ttk.Label(dates, text="统计范围：").pack(side="left")
    ttk.Entry(dates, textvariable=v_start, width=12).pack(side="left")
    ttk.Label(dates, text=" 至 ").pack(side="left")
    ttk.Entry(dates, textvariable=v_end, width=12).pack(side="left")
    ttk.Label(dates, text="    封面日期：").pack(side="left")
    ttk.Entry(dates, textvariable=v_date, width=12).pack(side="left")
    ttk.Label(dates, text="（留空 = 结束日期）", foreground="#888").pack(side="left")
    r += 1

    ttk.Label(frm, text="输出文件夹：").grid(row=r, column=0, sticky="w", pady=4)
    ttk.Entry(frm, textvariable=v_out).grid(row=r, column=1, sticky="ew", pady=4)
    ttk.Button(frm, text="选择…", command=lambda: pick_dir(v_out)).grid(row=r, column=2, padx=(6, 0))
    r += 1

    btns = ttk.Frame(frm)
    btns.grid(row=r, column=0, columnspan=3, sticky="w", pady=(10, 6))
    b_run = ttk.Button(btns, text="生成新 PPT")
    b_run.pack(side="left")
    b_open = ttk.Button(btns, text="打开输出文件夹", state="disabled")
    b_open.pack(side="left", padx=8)
    b_cfg = ttk.Button(btns, text="编辑规则 (config.json)",
                       command=lambda: os.startfile(config_path()) if hasattr(os, "startfile") else None)
    b_cfg.pack(side="left")
    r += 1

    logbox = tk.Text(frm, height=16, wrap="word", font=("Microsoft YaHei UI", 9))
    logbox.grid(row=r, column=0, columnspan=3, sticky="nsew", pady=(6, 0))
    frm.rowconfigure(r, weight=1)
    sb = ttk.Scrollbar(frm, command=logbox.yview)
    sb.grid(row=r, column=3, sticky="ns", pady=(6, 0))
    logbox.configure(yscrollcommand=sb.set)

    q = queue.Queue()
    state = {"out": None, "report": None}

    def log(msg):
        q.put(("log", str(msg)))

    def pump():
        try:
            while True:
                kind, payload = q.get_nowait()
                if kind == "log":
                    logbox.insert("end", payload + "\n")
                    logbox.see("end")
                elif kind == "done":
                    b_run.configure(state="normal")
                    out, report, warns = payload
                    state["out"], state["report"] = out, report
                    b_open.configure(state="normal")
                    msg = f"已生成：\n{out}\n\n核对清单：\n{report}"
                    if warns:
                        msg += f"\n\n有 {len(warns)} 条提醒，请看核对清单。"
                    if messagebox.askyesno(APP_NAME, msg + "\n\n现在打开 PPT 吗？"):
                        _open(out)
                elif kind == "error":
                    b_run.configure(state="normal")
                    messagebox.showerror(APP_NAME, payload)
        except queue.Empty:
            pass
        root.after(100, pump)

    def _open(p):
        if hasattr(os, "startfile"):
            os.startfile(p)

    def on_run():
        try:
            xlsx, ppt = v_xlsx.get().strip().strip('"'), v_ppt.get().strip().strip('"')
            if not os.path.isfile(xlsx):
                raise ValueError("请选择文件台账 xlsx")
            last_run = st.get("last_run_time")
            if last_run and os.path.getmtime(xlsx) <= last_run:
                if not messagebox.askyesno(APP_NAME,
                        f"这个台账自上次生成 PPT 以来没有修改过：\n{os.path.basename(xlsx)}\n"
                        f"（最后修改：{fmt_mtime(xlsx)}）\n\n"
                        "如果你这周已经更新过台账，可能是：\n"
                        "· 在 Excel 里改了但还没保存；\n"
                        "· 另存成了别的文件名或放在别的文件夹。\n\n"
                        "仍然用这个文件继续吗？"):
                    return
            if not os.path.isfile(ppt):
                raise ValueError("请选择上周的 PPT")
            start, end = parse_date(v_start.get()), parse_date(v_end.get())
            if start > end:
                raise ValueError("开始日期晚于结束日期")
            rdate = parse_date(v_date.get()) if v_date.get().strip() else end
            cfg = rc.load_config(config_path())
            out = output_path(cfg, v_out.get().strip(), rdate, os.path.dirname(ppt))
            if os.path.abspath(out) == os.path.abspath(ppt):
                base, ext = os.path.splitext(out)
                out = base + "_新" + ext
        except Exception as e:  # noqa
            messagebox.showerror(APP_NAME, str(e))
            return
        logbox.delete("1.0", "end")
        b_run.configure(state="disabled")
        log(f"使用台账：{xlsx}")
        log(f"  最后修改时间：{fmt_mtime(xlsx)}")

        def work():
            try:
                res = rc.update_report(xlsx, ppt, start, end, out, report_date=rdate, cfg=cfg, log=log)
                new_st = {"xlsx": xlsx, "ppt": ppt, "last_output": res[0],
                          "last_end": end.isoformat(), "out_dir": v_out.get().strip(),
                          "last_run_time": dt.datetime.now().timestamp()}
                st.clear(); st.update(new_st)
                save_settings(new_st)
                q.put(("done", res))
            except PermissionError:
                q.put(("error", f"无法写入：{out}\n文件可能正在 PowerPoint 里打开，请先关闭再试。"))
            except Exception as e:  # noqa
                log(traceback.format_exc())
                q.put(("error", f"出错了：{e}"))

        threading.Thread(target=work, daemon=True).start()

    b_run.configure(command=on_run)
    b_open.configure(command=lambda: _open(os.path.dirname(state["out"])) if state["out"] else None)

    log("使用方法：选择本周的文件台账 xlsx 和上周的 PPT，确认日期，点“生成新 PPT”。")
    if auto_xlsx and os.path.normcase(auto_xlsx) != os.path.normcase(saved_xlsx):
        log(f"已自动选中文件夹里最新的台账：{os.path.basename(auto_xlsx)}（{fmt_mtime(auto_xlsx)}）")
    elif v_xlsx.get() and os.path.isfile(v_xlsx.get()):
        log(f"台账：{os.path.basename(v_xlsx.get())}（最后修改 {fmt_mtime(v_xlsx.get())}）")
    log("提示：Excel 里改完台账请先保存，程序读取的是已保存的内容。")
    log("程序会更新 1.1–1.4 的表格、每页提交/批复统计、标题和封面日期；第 5、6 页不改。")
    root.after(100, pump)
    root.mainloop()


if __name__ == "__main__":
    if len(sys.argv) > 1:
        run_cli(sys.argv[1:])
    else:
        run_gui()
