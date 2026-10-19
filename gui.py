#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
龙洲校区「家作一周一上传」催交助手（图形界面）

功能：
  · 选择要检查的文件夹（学年目录，或直接选某一周）
  · 勾选要检查的周，或一键全选
  · 总览矩阵：每个班每周 √ / ×
  · 一键复制钉钉催交文案（群发一条 / 逐个发送）
  · 导出 Excel（总览矩阵 / 单周明细）

运行：双击 打开界面.bat，或 python gui.py
"""

from __future__ import annotations

import math
import os
import queue
import sys
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import checker

APP_TITLE = "龙洲校区 家作催交助手"
FONT_UI = ("Microsoft YaHei UI", 9)
FONT_TITLE = ("Microsoft YaHei UI", 10, "bold")


def enable_dpi_awareness() -> None:
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # type: ignore[attr-defined]
    except Exception:
        pass


class ScrollFrame(ttk.Frame):
    """带竖向滚动条的容器。"""

    def __init__(self, master, height: int = 160):
        super().__init__(master)
        self.canvas = tk.Canvas(self, height=height, highlightthickness=0, bd=0)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.inner = ttk.Frame(self.canvas)
        self._win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

        self.inner.bind("<Configure>", self._on_inner)
        self.canvas.bind("<Configure>", self._on_canvas)
        self.canvas.bind("<Enter>", lambda _e: self.canvas.bind_all("<MouseWheel>", self._on_wheel))
        self.canvas.bind("<Leave>", lambda _e: self.canvas.unbind_all("<MouseWheel>"))

    def _on_inner(self, _e=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas(self, event):
        self.canvas.itemconfigure(self._win, width=event.width)

    def _on_wheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.q: queue.Queue = queue.Queue()
        self.stop_flag = threading.Event()
        self.busy = False

        self.week_vars: list[tuple[int, Path, tk.BooleanVar]] = []
        self.week_sig: tuple[str, ...] = ()
        self.week_q: queue.Queue = queue.Queue()
        self._probe_busy = False
        self.results: list[checker.WeekResult] = []
        self.roster: dict[str, dict] = {}
        self.all_rows: list[dict] = []

        self.year_dir = tk.StringVar(value=str(checker.DEFAULT_YEAR_DIR))
        self.roster_path = tk.StringVar(value=str(checker.DEFAULT_ROSTER_XLS))
        self.status = tk.StringVar(value="就绪")
        self.summary = tk.StringVar(value="")

        # 矩阵筛选
        self.filter_kw = tk.StringVar()
        self.filter_grade = tk.StringVar(value="全部年级")
        self.filter_status = tk.StringVar(value="全部状态")
        self.filter_week = tk.StringVar(value="不限周次")
        self.auto_refresh = tk.BooleanVar(value=True)

        self._build_style()
        self._build_ui()
        self.root.after(120, self.refresh_weeks)
        self._start_auto_refresh()

    # ---------------- 外观 ----------------

    def _build_style(self):
        st = ttk.Style()
        for theme in ("vista", "winnative", "clam"):
            try:
                st.theme_use(theme)
                break
            except tk.TclError:
                continue
        st.configure(".", font=FONT_UI)
        st.configure("Treeview", rowheight=26, font=FONT_UI)
        st.configure("Treeview.Heading", font=FONT_TITLE)
        st.configure("Title.TLabel", font=("Microsoft YaHei UI", 13, "bold"), foreground="#1F4E79")
        st.configure("Hint.TLabel", foreground="#666666")
        st.configure("Primary.TButton", font=("Microsoft YaHei UI", 10, "bold"))

    def _build_ui(self):
        self.root.title(APP_TITLE)
        self.root.geometry("1120x780")
        self.root.minsize(900, 620)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(3, weight=1)

        # ---- 顶部：路径 ----
        top = ttk.LabelFrame(self.root, text=" 文件夹与配班表 ")
        top.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 6))
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text="检查目录").grid(row=0, column=0, padx=(10, 6), pady=6, sticky="e")
        ttk.Entry(top, textvariable=self.year_dir).grid(row=0, column=1, sticky="ew", pady=6)
        ttk.Button(top, text="选择…", width=8, command=self.pick_year_dir).grid(
            row=0, column=2, padx=6, pady=6)
        ttk.Button(top, text="刷新周列表", width=11, command=self.refresh_weeks).grid(
            row=0, column=3, padx=(0, 10), pady=6)

        ttk.Label(top, text="配班表").grid(row=1, column=0, padx=(10, 6), pady=6, sticky="e")
        ttk.Entry(top, textvariable=self.roster_path).grid(row=1, column=1, sticky="ew", pady=6)
        ttk.Button(top, text="选择…", width=8, command=self.pick_roster).grid(
            row=1, column=2, padx=6, pady=6)
        ttk.Button(top, text="重新读取", width=11, command=self.refresh_roster).grid(
            row=1, column=3, padx=(0, 10), pady=6)

        # ---- 中部：周选择 ----
        mid = ttk.LabelFrame(self.root, text=" 要检查的周 ")
        mid.grid(row=1, column=0, sticky="ew", padx=12, pady=6)
        mid.columnconfigure(0, weight=1)

        bar = ttk.Frame(mid)
        bar.grid(row=0, column=0, sticky="ew", padx=10, pady=(6, 0))
        ttk.Button(bar, text="全选（全部周）", command=self.select_all).pack(side="left")
        ttk.Button(bar, text="全不选", command=self.select_none).pack(side="left", padx=6)
        ttk.Button(bar, text="反选", command=self.select_invert).pack(side="left")
        ttk.Button(bar, text="只看最新一周", command=self.select_latest).pack(side="left", padx=6)
        ttk.Button(bar, text="最近 4 周", command=lambda: self.select_recent(4)).pack(side="left")
        ttk.Checkbutton(bar, text="自动刷新周列表", variable=self.auto_refresh).pack(
            side="left", padx=(12, 6))
        self.week_count_lbl = ttk.Label(bar, text="正在读取目录…", style="Hint.TLabel")
        self.week_count_lbl.pack(side="left", padx=4)

        self.week_scroll = ScrollFrame(mid, height=132)
        self.week_scroll.grid(row=1, column=0, sticky="ew", padx=10, pady=(4, 10))

        # ---- 操作条 ----
        self._build_actions()

        # ---- 结果 Notebook ----
        self.nb = ttk.Notebook(self.root)
        self.nb.grid(row=3, column=0, sticky="nsew", padx=12, pady=(0, 6))

        self.tab_matrix = ttk.Frame(self.nb)
        self.tab_msg = ttk.Frame(self.nb)
        self.tab_log = ttk.Frame(self.nb)
        self.nb.add(self.tab_matrix, text="  总览矩阵  ")
        self.nb.add(self.tab_msg, text="  催交文案  ")
        self.nb.add(self.tab_log, text="  运行日志  ")

        self._build_matrix_tab()
        self._build_msg_tab()
        self._build_log_tab()

        # ---- 状态栏 ----
        sb = ttk.Frame(self.root)
        sb.grid(row=4, column=0, sticky="ew", padx=12, pady=(0, 10))
        sb.columnconfigure(0, weight=1)
        ttk.Label(sb, textvariable=self.status, style="Hint.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(sb, textvariable=self.summary, style="Hint.TLabel").grid(row=0, column=1, sticky="e")

    def _build_actions(self):
        bar = ttk.Frame(self.root)
        bar.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 6))

        self.btn_start = ttk.Button(bar, text="开始检查", style="Primary.TButton",
                                    width=12, command=self.start_check)
        self.btn_start.pack(side="left")

        self.btn_stop = ttk.Button(bar, text="停止", width=7, state="disabled",
                                   command=self.stop_check)
        self.btn_stop.pack(side="left", padx=8)

        self.prog = ttk.Progressbar(bar, mode="determinate", length=220)
        self.prog.pack(side="left", padx=(4, 8))

        self.prog_lbl = ttk.Label(bar, text="", style="Hint.TLabel", width=30)
        self.prog_lbl.pack(side="left")

        ttk.Button(bar, text="导出总览 Excel", command=self.export_matrix).pack(side="right")
        ttk.Button(bar, text="只导出筛选结果", command=self.export_filtered).pack(
            side="right", padx=6)
        ttk.Button(bar, text="导出本周 Excel", command=self.export_week).pack(
            side="right", padx=6)
        ttk.Button(bar, text="打开输出目录", command=self.open_out_dir).pack(
            side="right", padx=6)

    # ---------------- 三个页签 ----------------

    def _build_matrix_tab(self):
        f = self.tab_matrix
        f.columnconfigure(0, weight=1)
        f.rowconfigure(1, weight=1)

        fb = ttk.Frame(f)
        fb.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 2))

        ttk.Label(fb, text="筛选", font=FONT_TITLE).pack(side="left", padx=(0, 8))

        ttk.Label(fb, text="搜索").pack(side="left")
        ent = ttk.Entry(fb, textvariable=self.filter_kw, width=12)
        ent.pack(side="left", padx=(2, 10))
        ent.bind("<KeyRelease>", lambda _e: self._apply_filters())

        ttk.Label(fb, text="年级").pack(side="left")
        grades = ["全部年级"] + [checker.GRADE_NAMES[k] for k in sorted(checker.GRADE_NAMES)]
        cb_grade = ttk.Combobox(fb, textvariable=self.filter_grade, state="readonly",
                                width=8, values=grades)
        cb_grade.pack(side="left", padx=(2, 10))
        cb_grade.bind("<<ComboboxSelected>>", lambda _e: self._apply_filters())

        ttk.Label(fb, text="状态").pack(side="left")
        cb_status = ttk.Combobox(fb, textvariable=self.filter_status, state="readonly",
                                 width=11, values=["全部状态", "只看有未交", "只看全交", "只看全缺"])
        cb_status.pack(side="left", padx=(2, 10))
        cb_status.bind("<<ComboboxSelected>>", lambda _e: self._apply_filters())

        ttk.Label(fb, text="指定周").pack(side="left")
        self.cb_week = ttk.Combobox(fb, textvariable=self.filter_week, state="readonly",
                                    width=17, values=["不限周次"])
        self.cb_week.pack(side="left", padx=(2, 10))
        self.cb_week.bind("<<ComboboxSelected>>", lambda _e: self._apply_filters())

        ttk.Button(fb, text="重置", width=6, command=self._reset_filters).pack(side="left")

        self.filter_count = ttk.Label(fb, text="", style="Hint.TLabel")
        self.filter_count.pack(side="right")
        ttk.Button(fb, text="复制表格", width=9, command=self.copy_visible_table).pack(
            side="right", padx=(0, 8))

        wrap = ttk.Frame(f)
        wrap.grid(row=1, column=0, sticky="nsew", padx=8, pady=(4, 8))
        wrap.columnconfigure(0, weight=1)
        wrap.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(wrap, show="headings", selectmode="browse")
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        hsb = ttk.Scrollbar(wrap, orient="horizontal", command=self.tree.xview)
        hsb.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.tag_configure("allmiss", foreground="#C00000")
        self.tree.tag_configure("allok", foreground="#2E7D32")
        self.tree.tag_configure("odd", background="#FAFAFA")

        # 表格内容复制：双击单元格 / 右键菜单 / Ctrl+C
        self.tree.bind("<Double-1>", self._copy_clicked_cell)
        self.tree.bind("<Button-3>", self._show_tree_menu)
        self.tree.bind("<Control-c>", self._on_ctrl_c)

        self.empty_hint = ttk.Label(
            wrap, text="勾选上面的周，然后点「开始检查」",
            style="Hint.TLabel", font=("Microsoft YaHei UI", 11))
        self.empty_hint.place(relx=0.5, rely=0.45, anchor="center")

    def _build_msg_tab(self):
        f = self.tab_msg
        f.columnconfigure(0, weight=1)
        f.rowconfigure(1, weight=1)

        bar = ttk.Frame(f)
        bar.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 4))

        ttk.Label(bar, text="周次").pack(side="left", padx=(0, 4))
        self.week_combo = ttk.Combobox(bar, state="readonly", width=22, values=[])
        self.week_combo.pack(side="left")
        self.week_combo.bind("<<ComboboxSelected>>", lambda _e: self.render_message())

        self.msg_mode = tk.StringVar(value="group")
        ttk.Radiobutton(bar, text="群发一条", value="group", variable=self.msg_mode,
                        command=self.render_message).pack(side="left", padx=(14, 4))
        ttk.Radiobutton(bar, text="逐个发送", value="single", variable=self.msg_mode,
                        command=self.render_message).pack(side="left")

        ttk.Button(bar, text="复制文案", command=self.copy_message).pack(side="left", padx=14)

        wrap = ttk.Frame(f)
        wrap.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))
        wrap.columnconfigure(0, weight=1)
        wrap.rowconfigure(0, weight=1)

        self.msg_text = tk.Text(wrap, wrap="word", font=("Microsoft YaHei UI", 10),
                                relief="solid", bd=1, padx=10, pady=8)
        self.msg_text.grid(row=0, column=0, sticky="nsew")
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=self.msg_text.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        self.msg_text.configure(yscrollcommand=vsb.set)
        self.msg_text.insert("1.0", "检查完成后，这里会生成可直接粘贴到钉钉的催交文案。")
        self.msg_text.configure(state="disabled")

    def _build_log_tab(self):
        f = self.tab_log
        f.columnconfigure(0, weight=1)
        f.rowconfigure(0, weight=1)
        self.log_text = tk.Text(f, wrap="none", font=("Consolas", 9),
                                relief="solid", bd=1, padx=8, pady=6)
        self.log_text.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        vsb = ttk.Scrollbar(f, orient="vertical", command=self.log_text.yview)
        vsb.grid(row=0, column=1, sticky="ns", pady=8)
        self.log_text.configure(yscrollcommand=vsb.set)

    # ---------------- 路径选择 ----------------

    def pick_year_dir(self):
        d = filedialog.askdirectory(title="选择学年目录（里面有 第N周 文件夹），或直接选某一周",
                                    initialdir=self.year_dir.get() or str(Path.home()))
        if d:
            self.year_dir.set(os.path.normpath(d))
            self.refresh_weeks()

    def pick_roster(self):
        f = filedialog.askopenfilename(
            title="选择教师配班安排表",
            initialdir=str(Path(self.roster_path.get()).parent),
            filetypes=[("Excel 表格", "*.xls *.xlsx"), ("所有文件", "*.*")])
        if f:
            self.roster_path.set(os.path.normpath(f))
            self.log("已选择配班表：" + f)

    def refresh_roster(self):
        self.log("将强制重新读取配班表（下次检查时生效）")
        self.status.set("已标记：下次检查重新读取配班表")
        self._force_roster = True

    def open_out_dir(self):
        checker.OUT_DIR.mkdir(parents=True, exist_ok=True)
        try:
            os.startfile(str(checker.OUT_DIR))  # type: ignore[attr-defined]
        except Exception as e:
            messagebox.showinfo("输出目录", str(checker.OUT_DIR))

    # ---------------- 周列表 ----------------

    def refresh_weeks(self):
        try:
            weeks = checker.list_weeks(Path(self.year_dir.get()))
        except Exception as e:
            self.log(f"读取目录失败：{e}")
            self.status.set("读取目录失败，请检查路径或网络")
            self.week_count_lbl.configure(text="读取失败")
            return

        sig = tuple(p.name for _, p in weeks)
        if sig == self.week_sig and self.week_vars:
            return  # 目录没变化，保留现有勾选，不重建

        old_state = {no: v.get() for no, _, v in self.week_vars}
        prev_selected = {no for no, ok in old_state.items() if ok}
        all_selected = bool(old_state) and len(prev_selected) == len(old_state)
        first_load = not self.week_vars

        for w in self.week_scroll.inner.winfo_children():
            w.destroy()
        self.week_vars.clear()
        self.week_sig = sig

        if not weeks:
            self.log("该目录下没有「第N周」文件夹")
            self.status.set("没有找到周文件夹")
            self.week_count_lbl.configure(text="未发现周目录")
            return

        latest = weeks[-1][0]
        cols = 4 if len(weeks) > 12 else 3
        for idx, (no, path) in enumerate(weeks):
            if first_load:
                checked = (no == latest)
            elif all_selected:
                checked = True  # 之前是全选，新的一周自动加入
            else:
                checked = no in prev_selected

            var = tk.BooleanVar(value=checked)
            label = path.name if path.name.startswith("第") else f"第{no}周"
            ttk.Checkbutton(self.week_scroll.inner, text=label, variable=var).grid(
                row=idx // cols, column=idx % cols, sticky="w", padx=8, pady=3)
            self.week_vars.append((no, path, var))

        rows_needed = math.ceil(len(weeks) / cols)
        self.week_scroll.canvas.configure(height=max(36, min(210, rows_needed * 32 + 8)))

        total = len(weeks)
        span = (f"第{weeks[0][0]}–{weeks[-1][0]}周" if total > 1 else f"第{weeks[0][0]}周")
        self.week_count_lbl.configure(text=f"共发现 {total} 周（{span}）")
        self.log(f"已载入 {total} 个周目录：{'、'.join(f'第{n}周' for n, _ in weeks)}")

        new_nos = sorted({no for no, _ in weeks} - set(old_state))
        if new_nos and not first_load:
            tip = "、".join(f"第{x}周" for x in new_nos)
            self.log(f"发现新的周目录：{tip}（默认未勾选，需要查请手动勾上）")
            self.status.set(f"已载入 {total} 个周，新增 {tip}")
        else:
            self.status.set(f"已载入 {total} 个周")

    def select_all(self):
        for _, _, v in self.week_vars:
            v.set(True)

    def select_none(self):
        for _, _, v in self.week_vars:
            v.set(False)

    def select_invert(self):
        for _, _, v in self.week_vars:
            v.set(not v.get())

    def select_latest(self):
        for i, (_, _, v) in enumerate(self.week_vars):
            v.set(i == len(self.week_vars) - 1)

    def select_recent(self, n: int):
        for i, (_, _, v) in enumerate(self.week_vars):
            v.set(i >= len(self.week_vars) - n)

    # ---------------- 自动发现新周 ----------------

    def _start_auto_refresh(self):
        self.root.after(1000, self._drain_week_q)
        self.root.after(30000, self._auto_tick)

    def _auto_tick(self):
        if self.auto_refresh.get() and not self._probe_busy:
            self._probe_busy = True
            threading.Thread(target=self._auto_probe, daemon=True).start()
        self.root.after(30000, self._auto_tick)

    def _auto_probe(self):
        try:
            weeks = checker.list_weeks(Path(self.year_dir.get()))
            self.week_q.put((tuple(p.name for _, p in weeks), len(weeks)))
        except Exception:
            pass
        finally:
            self._probe_busy = False

    def _drain_week_q(self):
        try:
            while True:
                sig, count = self.week_q.get_nowait()
                if sig != self.week_sig:
                    self.log(f"检测到目录变化（现有 {count} 个周），自动刷新周列表")
                    self.refresh_weeks()
                    self.status.set("已自动发现新的周目录")
        except queue.Empty:
            pass
        self.root.after(800, self._drain_week_q)

    # ---------------- 检查流程 ----------------

    def start_check(self):
        if self.busy:
            return
        weeks = [(no, p) for no, p, v in self.week_vars if v.get()]
        if not weeks:
            messagebox.showwarning("还没选周", "请先勾选至少一个要检查的周。")
            return

        self.busy = True
        self.stop_flag.clear()
        self.btn_start.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        self.prog.configure(maximum=len(weeks), value=0)
        self.prog_lbl.configure(text="准备中…")
        self.status.set("正在检查…")
        self.empty_hint.place_forget()

        force = getattr(self, "_force_roster", False)
        self._force_roster = False
        threading.Thread(target=self._worker, args=(weeks, force), daemon=True).start()
        self.root.after(120, self._poll)

    def _worker(self, weeks, force):
        try:
            self.q.put(("log", "读取配班表…"))
            roster = checker.load_roster(
                Path(self.roster_path.get()), force=force,
                log=lambda m: self.q.put(("log", m)))
            self.q.put(("log", f"配班表：{len(roster)} 个班级"))

            results = checker.analyze_weeks(
                roster, weeks,
                progress=lambda i, n, name: self.q.put(("progress", i, n, name)),
                should_stop=self.stop_flag.is_set)
            self.q.put(("done", roster, results))
        except Exception as e:
            self.q.put(("error", str(e), traceback.format_exc()))

    def _poll(self):
        try:
            while True:
                item = self.q.get_nowait()
                kind = item[0]
                if kind == "log":
                    self.log(item[1])
                elif kind == "progress":
                    _, i, n, name = item
                    self.prog.configure(value=i - 1)
                    self.prog_lbl.configure(text=f"({i}/{n}) {name}")
                    self.status.set(f"正在扫描 {name} …")
                elif kind == "done":
                    _, roster, results = item
                    self._finish(roster, results)
                    return
                elif kind == "error":
                    _, msg, tb = item
                    self.log("检查失败：" + msg)
                    self.log(tb)
                    messagebox.showerror("检查失败", msg)
                    self._reset_busy()
                    return
        except queue.Empty:
            pass
        if self.busy:
            self.root.after(120, self._poll)

    def _reset_busy(self):
        self.busy = False
        self.btn_start.configure(state="normal")
        self.btn_stop.configure(state="disabled")
        self.prog_lbl.configure(text="")

    def stop_check(self):
        self.stop_flag.set()
        self.status.set("正在停止…")

    def _finish(self, roster, results):
        self.roster = roster
        self.results = results
        self._reset_busy()
        self.prog.configure(value=self.prog.cget("maximum"))

        if not results:
            self.status.set("已停止，未得到结果")
            return

        self._fill_matrix()
        self._fill_combo()
        self.render_message()

        total_missing = sum(r.missing_count for r in results)
        self.status.set(f"检查完成：{len(results)} 周，未交 {total_missing} 条　（双击单元格可复制）")
        self.prog_lbl.configure(text="完成")
        self.summary.set(
            f"{len(results)} 周｜{len(roster)} 个班｜未交 {total_missing} 条")
        if getattr(self, "_demo_filter", False):
            self.filter_status.set("只看有未交")
            self._apply_filters()
        self.nb.select(getattr(self, "_demo_tab", 0))

    # ---------------- 结果展示 ----------------

    def _fill_matrix(self):
        self.empty_hint.place_forget()
        self.tree.delete(*self.tree.get_children())

        results = self.results
        cols = ["code", "teacher", "partner"] + [f"w{i}" for i in range(len(results))] + ["miss"]
        self.tree.configure(columns=cols)

        self.tree.heading("code", text="班级")
        self.tree.heading("teacher", text="班主任")
        self.tree.heading("partner", text="配班老师")
        for i, res in enumerate(results):
            # 列头只要“第7周”，日期范围在勾选框和「指定周」里能看，避免列太宽
            short = f"第{res.week_no}周" if res.week_no is not None else res.label
            self.tree.heading(f"w{i}", text=short)
        self.tree.heading("miss", text="未交")

        self.tree.column("code", width=56, anchor="center", stretch=False)
        self.tree.column("teacher", width=76, anchor="center", stretch=False)
        self.tree.column("partner", width=76, anchor="center", stretch=False)
        for i in range(len(results)):
            self.tree.column(f"w{i}", width=64, anchor="center", stretch=False)
        self.tree.column("miss", width=64, anchor="center", stretch=False)

        self.all_rows = []
        order = sorted(self.roster, key=lambda c: (
            checker.GRADE_ORDER.get(self.roster[c]["grade"], 99), c))
        for n, code in enumerate(order):
            info = self.roster[code]
            vals = [code, info["teacher"], info["partner"]]
            miss = 0
            for res in results:
                ok = bool(res.submitted.get(code))
                vals.append("√" if ok else "×")
                if not ok:
                    miss += 1
            vals.append(miss)
            tag = ""
            if miss == len(results):
                tag = "allmiss"
            elif miss == 0:
                tag = "allok"
            elif n % 2 == 1:
                tag = "odd"
            self.all_rows.append({
                "code": code,
                "grade": info["grade"],
                "teacher": info["teacher"],
                "partner": info["partner"],
                "values": vals,
                "miss": miss,
                "weeks": len(results),
                "tag": tag,
            })

        labels = ["不限周次"] + [r.label for r in results]
        self.cb_week.configure(values=labels)
        if self.filter_week.get() not in labels:
            self.filter_week.set("不限周次")
        self._apply_filters()

    def _row_matches(self, row: dict) -> bool:
        kw = self.filter_kw.get().strip().lower()
        if kw:
            hay = f"{row['code']}{row['teacher']}{row['partner']}".lower()
            if kw not in hay:
                return False
        grade = self.filter_grade.get()
        if grade != "全部年级" and row["grade"] != grade:
            return False
        status = self.filter_status.get()
        if status == "只看有未交" and row["miss"] == 0:
            return False
        if status == "只看全交" and row["miss"] != 0:
            return False
        if status == "只看全缺" and row["miss"] != row["weeks"]:
            return False
        picked = self.filter_week.get()
        if picked and picked != "不限周次":
            try:
                week_idx = [r.label for r in self.results].index(picked)
            except ValueError:
                week_idx = None
            if week_idx is not None and row["values"][3 + week_idx] != "×":
                return False
        return True

    def filtered_rows(self) -> list[dict]:
        return [r for r in self.all_rows if self._row_matches(r)]

    def _filter_note(self) -> str:
        parts = []
        kw = self.filter_kw.get().strip()
        if kw:
            parts.append(f"搜索「{kw}」")
        if self.filter_grade.get() != "全部年级":
            parts.append(self.filter_grade.get())
        if self.filter_status.get() != "全部状态":
            parts.append(self.filter_status.get())
        picked = self.filter_week.get()
        if picked and picked != "不限周次":
            parts.append(f"只看{picked}未交")
        return "筛选条件：" + ("、".join(parts) if parts else "无（全部班级）")

    def _apply_filters(self):
        if not self.all_rows:
            return
        self.tree.delete(*self.tree.get_children())
        shown = 0
        for row in self.filtered_rows():
            self.tree.insert("", "end", values=row["values"],
                             tags=(row["tag"],) if row["tag"] else ())
            shown += 1
        self.filter_count.configure(text=f"显示 {shown} / {len(self.all_rows)} 个班")

    def _reset_filters(self):
        self.filter_kw.set("")
        self.filter_grade.set("全部年级")
        self.filter_status.set("全部状态")
        self.filter_week.set("不限周次")
        self._apply_filters()

    # ---------------- 表格文字复制 ----------------

    def _copy(self, text: str, msg: str) -> None:
        if not text:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.root.update_idletasks()
        self.status.set(msg)

    def _row_text(self, iid: str) -> str:
        return "\t".join(str(v) for v in self.tree.item(iid, "values"))

    def _col_titles(self) -> list[str]:
        return [self.tree.heading(c, "text") for c in self.tree.cget("columns")]

    def copy_tree_selection(self) -> None:
        """Ctrl+C / 右键：复制选中行（制表符分隔，可直接粘到 Excel）。"""
        sel = self.tree.selection()
        if not sel:
            self.status.set("先在表格里点一行，再复制")
            return
        if len(sel) == 1:
            self._copy(self._row_text(sel[0]), "已复制该行（制表符分隔）")
        else:
            self._copy("\n".join(self._row_text(i) for i in sel),
                       f"已复制 {len(sel)} 行")

    def _copy_clicked_cell(self, event) -> None:
        """双击某个单元格：只复制这一格。"""
        iid = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not iid or not col:
            return
        idx = int(col.lstrip("#")) - 1
        vals = self.tree.item(iid, "values")
        if idx < 0 or idx >= len(vals):
            return
        self.tree.selection_set(iid)
        self.tree.focus(iid)
        self._copy(str(vals[idx]), f"已复制：{vals[idx]}")

    def copy_visible_table(self) -> None:
        """复制当前筛选出来的整张表（含表头），粘到 Excel 直接成型。"""
        rows = self.tree.get_children()
        if not rows:
            self.status.set("当前没有可复制的行")
            return
        lines = ["\t".join(self._col_titles())]
        lines += [self._row_text(i) for i in rows]
        self._copy("\n".join(lines),
                   f"已复制 {len(rows)} 行（含表头，可直接粘到 Excel）")

    def copy_teacher_list(self) -> None:
        """复制“班级 班主任”清单，适合直接发群。"""
        rows = self.tree.get_children()
        if not rows:
            self.status.set("当前没有可复制的行")
            return
        parts = []
        for iid in rows:
            v = self.tree.item(iid, "values")
            parts.append(f"{v[0]} {v[1]}")
        self._copy("、".join(parts), f"已复制 {len(parts)} 个「班级 班主任」")

    def copy_teacher_only(self) -> None:
        sel = self.tree.selection()
        if not sel:
            self.status.set("先点一行，再复制班主任")
            return
        v = self.tree.item(sel[0], "values")
        self._copy(str(v[1]), f"已复制班主任：{v[1]}")

    def _on_ctrl_c(self, _event):
        self.copy_tree_selection()
        return "break"

    def _show_tree_menu(self, event):
        iid = self.tree.identify_row(event.y)
        if iid:
            self.tree.selection_set(iid)
            self.tree.focus(iid)
        m = tk.Menu(self.root, tearoff=0, font=FONT_UI)
        m.add_command(label="复制这一行", command=self.copy_tree_selection)
        m.add_command(label="只复制班主任姓名", command=self.copy_teacher_only)
        m.add_separator()
        m.add_command(label="复制「班级 班主任」清单", command=self.copy_teacher_list)
        m.add_command(label="复制整张表（含表头）", command=self.copy_visible_table)
        try:
            m.tk_popup(event.x_root, event.y_root)
        finally:
            m.grab_release()

    def _fill_combo(self):
        vals = [r.label for r in self.results]
        self.week_combo.configure(values=vals)
        if vals:
            self.week_combo.current(len(vals) - 1)

    def current_result(self) -> checker.WeekResult | None:
        idx = self.week_combo.current()
        if idx < 0 or idx >= len(self.results):
            return None
        return self.results[idx]

    def render_message(self):
        res = self.current_result()
        if res is None:
            return
        year_dir = Path(self.year_dir.get())
        text = checker.messages_for(res, self.msg_mode.get(), year_dir)
        self.msg_text.configure(state="normal")
        self.msg_text.delete("1.0", "end")
        self.msg_text.insert("1.0", text)
        self.msg_text.configure(state="disabled")

    def copy_message(self):
        text = self.msg_text.get("1.0", "end-1c").strip()
        if not text:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.root.update_idletasks()
        self.status.set("文案已复制到剪贴板")

    # ---------------- 导出 ----------------

    def export_filtered(self):
        """只导出当前筛选出来的班级（总览矩阵 + 未交明细）。"""
        if not self.results:
            messagebox.showinfo("暂无结果", "请先点「开始检查」。")
            return
        rows = self.filtered_rows()
        if not rows:
            messagebox.showinfo("筛选结果为空", "当前筛选条件下没有班级，请先调整筛选。")
            return

        checker.OUT_DIR.mkdir(parents=True, exist_ok=True)
        default = f"筛选结果_{len(rows)}个班.xlsx"
        path = filedialog.asksaveasfilename(
            title="导出筛选结果", initialdir=str(checker.OUT_DIR),
            initialfile=default, defaultextension=".xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx")])
        if not path:
            return
        try:
            checker.write_matrix_xlsx(
                Path(path), self.results, self.roster, Path(self.year_dir.get()),
                codes=[r["code"] for r in rows], note=self._filter_note())
            self.status.set(f"已导出筛选结果：{len(rows)} 个班")
            self.log(f"已导出筛选结果（{len(rows)} 个班）：{path}")
            if messagebox.askyesno("导出成功", f"已保存 {len(rows)} 个班：\n{path}\n\n现在打开吗？"):
                os.startfile(path)  # type: ignore[attr-defined]
        except Exception as e:
            messagebox.showerror("导出失败", str(e))
            self.log("导出失败：" + str(e))

    def export_matrix(self):
        if not self.results:
            messagebox.showinfo("暂无结果", "请先点「开始检查」。")
            return
        checker.OUT_DIR.mkdir(parents=True, exist_ok=True)
        default = checker.OUT_DIR / f"上交总览_{len(self.results)}周.xlsx"
        path = filedialog.asksaveasfilename(
            title="保存总览矩阵", initialdir=str(checker.OUT_DIR),
            initialfile=default.name, defaultextension=".xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx")])
        if not path:
            return
        try:
            checker.write_matrix_xlsx(Path(path), self.results, self.roster,
                                      Path(self.year_dir.get()))
            self.status.set("已导出：" + Path(path).name)
            self.log("已导出总览矩阵：" + path)
            if messagebox.askyesno("导出成功", f"已保存：\n{path}\n\n现在打开吗？"):
                os.startfile(path)  # type: ignore[attr-defined]
        except Exception as e:
            messagebox.showerror("导出失败", str(e))
            self.log("导出失败：" + str(e))

    def export_week(self):
        res = self.current_result()
        if res is None:
            messagebox.showinfo("暂无结果", "请先点「开始检查」。")
            return
        checker.OUT_DIR.mkdir(parents=True, exist_ok=True)
        default = f"催交名单_{res.label}.xlsx"
        path = filedialog.asksaveasfilename(
            title="保存本周明细", initialdir=str(checker.OUT_DIR),
            initialfile=default, defaultextension=".xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx")])
        if not path:
            return
        try:
            checker.write_xlsx(Path(path), res, Path(self.year_dir.get()))
            self.status.set("已导出：" + Path(path).name)
            self.log("已导出单周明细：" + path)
            if messagebox.askyesno("导出成功", f"已保存：\n{path}\n\n现在打开吗？"):
                os.startfile(path)  # type: ignore[attr-defined]
        except Exception as e:
            messagebox.showerror("导出失败", str(e))
            self.log("导出失败：" + str(e))

    # ---------------- 日志 ----------------

    def log(self, msg: str):
        from datetime import datetime
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{ts}] {msg}\n")
        self.log_text.see("end")


def _emit(lines: list[str]) -> None:
    """自检输出：有 --smoke-out 就写文件（打包成窗口程序后没有控制台）。"""
    out = None
    if "--smoke-out" in sys.argv:
        i = sys.argv.index("--smoke-out")
        if i + 1 < len(sys.argv):
            out = Path(sys.argv[i + 1])
    if out:
        try:
            out.parent.mkdir(parents=True, exist_ok=True)
            out.write_text("\n".join(lines), encoding="utf-8")
        except Exception:
            pass
    print("\n".join(lines))


def main() -> int:
    # 打包成 --windowed 后 sys.stdout / sys.stderr 是 None，print 会抛异常
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w", encoding="utf-8")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w", encoding="utf-8")

    enable_dpi_awareness()
    root = tk.Tk()
    app = App(root)

    if "--smoke" in sys.argv:
        # 无窗口自检：载入周、跑完一次检查、校验矩阵行数
        root.withdraw()
        root.update()
        app.refresh_weeks()
        root.update()
        if not app.week_vars:
            _emit(["SMOKE FAIL: 没有周目录"])
            root.destroy()
            return 1
        if "--smoke-all" in sys.argv:
            app.select_all()
        else:
            app.select_latest()
        weeks = [(no, p) for no, p, v in app.week_vars if v.get()]
        roster = checker.load_roster(Path(app.roster_path.get()))
        results = checker.analyze_weeks(roster, weeks)
        app._finish(roster, results)
        root.update()
        rows = len(app.tree.get_children())
        cols = app.tree.cget("columns")
        text = app.msg_text.get("1.0", "end-1c")
        _emit([
            f"SMOKE OK weeks={len(results)} roster={len(roster)} "
            f"rows={rows} cols={len(cols)}",
            f"msg_chars={len(text)}",
            f"out_dir={checker.OUT_DIR}",
            f"roster_csv={checker.ROSTER_CSV}",
        ])
        root.destroy()
        return 0

    if "--demo" in sys.argv:
        # 演示模式：自动全选并跑一次，方便截图/试用
        app._demo_tab = 1 if "--demo-msg" in sys.argv else 0
        app._demo_filter = "--demo-filter" in sys.argv

        def _demo():
            app.select_all()
            app.start_check()
        root.after(400, _demo)

    root.mainloop()
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        err = traceback.format_exc()
        if "--smoke" in sys.argv:
            try:
                _emit(["SMOKE FAIL", err])
            except Exception:
                pass
            sys.exit(1)
        try:
            from tkinter import messagebox
            messagebox.showerror(APP_TITLE, err)
        except Exception:
            print(err, file=sys.stderr)
        sys.exit(1)
