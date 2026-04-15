"""
Monte Carlo Projection Tool v2.4
==================================
Kanban throughput-based forecasting using Jira issue exports.
Runs entirely locally — no data leaves your machine.

REQUIREMENTS:
    pip install pandas numpy matplotlib
    (tkinter is built into Python — no install needed)

USAGE:
    python monte_carlo_jira.py

CHANGELOG:
    v2.4 — Capacity adjustment (% availability), HTML & PDF report export
    v2.3 — Cross-platform button fix (macOS compatibility)
    v2.2 — Forecast finish dates, scrollable left panel
    v2.1 — Kanban throughput mode, weekly throughput tab
    v2.0 — GUI with tkinter
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
import sys
import base64
import io
from datetime import date, datetime, timedelta

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib.ticker as mticker

# ── Version ────────────────────────────────────────────
VERSION = "2.4"

# ── Colours ────────────────────────────────────────────
BG        = "#1e1e2e"
BG2       = "#2a2a3e"
ACCENT    = "#7c6af7"
ACCENT2   = "#26c6da"
TEXT      = "#e0e0f0"
TEXT_DIM  = "#9090b0"
SUCCESS   = "#66bb6a"
WARNING   = "#ffa726"
ERROR     = "#ef5350"
CONF_COLS = {50: "#66bb6a", 70: "#26c6da", 85: "#ffa726", 95: "#ef5350"}
CONF_HEX  = {50: "#66bb6a", 70: "#26c6da", 85: "#ffa726", 95: "#ef5350"}

FONT       = ("Segoe UI", 10)
FONT_BOLD  = ("Segoe UI", 10, "bold")
FONT_TITLE = ("Segoe UI", 13, "bold")
FONT_SMALL = ("Segoe UI", 9)
FONT_MONO  = ("Consolas", 9)

IS_MAC = sys.platform == "darwin"


# ── Cross-platform button ──────────────────────────────
def mk_button(parent, text, command, font=FONT, bg=ACCENT, fg="white",
              padx=12, pady=6, cursor="hand2", **kwargs):
    frame = tk.Frame(parent, bg=bg, cursor=cursor)
    label = tk.Label(frame, text=text, font=font, bg=bg, fg=fg,
                     padx=padx, pady=pady, cursor=cursor)
    label.pack(fill="both", expand=True)

    def on_press(e):  label.config(relief="sunken")
    def on_release(e):
        label.config(relief="flat")
        command()

    label.bind("<Button-1>", on_press)
    label.bind("<ButtonRelease-1>", on_release)
    frame.bind("<Button-1>", on_press)
    frame.bind("<ButtonRelease-1>", on_release)

    def config(**kw):
        if "state" in kw:
            state = kw.pop("state")
            label.config(fg="#888888" if state == "disabled" else fg)
        if "text" in kw:
            label.config(text=kw.pop("text"))
        if kw:
            label.config(**kw)

    frame.config_btn = config
    frame.config = lambda **kw: config(**kw)
    return frame


# ══════════════════════════════════════════════════════
# DATA HELPERS
# ══════════════════════════════════════════════════════

def load_kanban_throughput(path):
    df = pd.read_csv(path, low_memory=False)
    df.columns = df.columns.str.strip()
    date_candidates = [c for c in df.columns if any(k in c.lower() for k in
                       ["resolved", "resolutiondate", "done date", "completed", "closed"])]
    if not date_candidates:
        raise ValueError("Cannot find a resolved/completed date column.\n"
                         "Columns found: " + str(list(df.columns)[:20]))
    date_col = date_candidates[0]
    dates = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True).dropna()
    if len(dates) == 0:
        raise ValueError("No valid dates found in column '%s'." % date_col)
    week_series = dates.dt.to_period("W")
    weekly_counts = week_series.value_counts().sort_index()
    if len(weekly_counts) > 1:
        full_range = pd.period_range(weekly_counts.index.min(),
                                     weekly_counts.index.max(), freq="W")
        weekly_counts = weekly_counts.reindex(full_range, fill_value=0)
    throughput = weekly_counts.values.astype(float)
    return throughput, date_col, len(dates), weekly_counts


def load_cycle_time(path):
    if not path or not os.path.exists(path):
        return None, None
    df = pd.read_csv(path)
    df.columns = df.columns.str.strip()
    candidates = [c for c in df.columns if "cycle" in c.lower() or "lead" in c.lower()
                  or ("days" in c.lower() and "time" in c.lower())]
    if not candidates:
        candidates = [c for c in df.columns if "days" in c.lower()]
    if not candidates:
        return None, None
    col = candidates[0]
    values = pd.to_numeric(df[col], errors="coerce").dropna()
    values = values[values > 0]
    return values.to_numpy(), col


# ── Monte Carlo engines ────────────────────────────────

def mc_weeks_to_done(samples, backlog, n):
    results = np.zeros(n, dtype=int)
    for i in range(n):
        remaining = backlog
        count = 0
        while remaining > 0:
            remaining -= np.random.choice(samples)
            count += 1
            if count > 2000:
                break
        results[i] = count
    return results


def mc_throughput_in_periods(samples, periods, n):
    return np.array([np.random.choice(samples, size=periods).sum() for _ in range(n)])


# ══════════════════════════════════════════════════════
# APPLICATION
# ══════════════════════════════════════════════════════

class MonteCarloApp(tk.Tk):

    def __init__(self):
        print("Initialising...")
        super().__init__()
        self.title("Monte Carlo Projection Tool  v%s" % VERSION)
        self.configure(bg=BG)
        self.resizable(True, True)
        self.minsize(940, 680)
        self._fig = None
        self._last_results = None   # stores last run data for report export
        print("Building UI...")
        self._build_ui()
        print("Centring window...")
        self._centre_window(1020, 780)
        print("Ready.")

    def _centre_window(self, w, h):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry("%dx%d+%d+%d" % (w, h, max(0, (sw-w)//2), max(0, (sh-h)//2)))

    # ── Layout ─────────────────────────────────────────

    def _build_ui(self):
        hdr = tk.Frame(self, bg=ACCENT, pady=10)
        hdr.pack(fill="x")
        tk.Label(hdr, text="Monte Carlo Projection Tool  v%s" % VERSION,
                 font=FONT_TITLE, bg=ACCENT, fg="white").pack()
        tk.Label(hdr, text="Kanban throughput-based forecasting  |  "
                            "All processing is local — no data leaves your machine",
                 font=FONT_SMALL, bg=ACCENT, fg="#ddd").pack()

        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=14, pady=12)

        # Scrollable left panel
        left_outer = tk.Frame(body, bg=BG, width=320)
        left_outer.pack(side="left", fill="y", padx=(0, 12))
        left_outer.pack_propagate(False)

        left_canvas = tk.Canvas(left_outer, bg=BG, highlightthickness=0)
        left_scroll = tk.Scrollbar(left_outer, orient="vertical", command=left_canvas.yview)
        left_canvas.configure(yscrollcommand=left_scroll.set)
        left_scroll.pack(side="right", fill="y")
        left_canvas.pack(side="left", fill="both", expand=True)

        left = tk.Frame(left_canvas, bg=BG)
        left_win = left_canvas.create_window((0, 0), window=left, anchor="nw")

        left.bind("<Configure>", lambda e: left_canvas.configure(
            scrollregion=left_canvas.bbox("all")))
        left_canvas.bind("<Configure>", lambda e: left_canvas.itemconfig(
            left_win, width=e.width))
        left_canvas.bind_all("<MouseWheel>", lambda e: left_canvas.yview_scroll(
            int(-1 * (e.delta / 120)), "units"))

        right = tk.Frame(body, bg=BG)
        right.pack(side="left", fill="both", expand=True)

        self._build_inputs(left)
        self._build_results(right)

    def _section(self, parent, title):
        tk.Label(parent, text=title, font=FONT_BOLD,
                 bg=BG, fg=ACCENT).pack(anchor="w", pady=(10, 4))
        tk.Frame(parent, bg=ACCENT, height=1).pack(fill="x", pady=(0, 8))

    def _labelled_entry(self, parent, label, default, width=10):
        row = tk.Frame(parent, bg=BG)
        row.pack(fill="x", pady=3)
        tk.Label(row, text=label, font=FONT, bg=BG, fg=TEXT,
                 width=24, anchor="w").pack(side="left")
        var = tk.StringVar(value=str(default))
        tk.Entry(row, textvariable=var, font=FONT_MONO, width=width,
                 bg=BG2, fg=TEXT, insertbackground=TEXT,
                 relief="flat", bd=4).pack(side="left")
        return var

    def _labelled_entry_dyn(self, parent, label_var, default, width=10):
        row = tk.Frame(parent, bg=BG)
        row.pack(fill="x", pady=3)
        tk.Label(row, textvariable=label_var, font=FONT, bg=BG, fg=TEXT,
                 width=24, anchor="w").pack(side="left")
        var = tk.StringVar(value=str(default))
        tk.Entry(row, textvariable=var, font=FONT_MONO, width=width,
                 bg=BG2, fg=TEXT, insertbackground=TEXT,
                 relief="flat", bd=4).pack(side="left")
        return var

    def _file_row(self, parent, label, optional=False):
        lbl = label + (" (optional)" if optional else "")
        tk.Label(parent, text=lbl, font=FONT,
                 bg=BG, fg=TEXT_DIM if optional else TEXT).pack(anchor="w", pady=(4, 1))
        row = tk.Frame(parent, bg=BG)
        row.pack(fill="x", pady=(0, 4))
        var = tk.StringVar()
        tk.Entry(row, textvariable=var, font=FONT_SMALL, bg=BG2, fg=TEXT,
                 insertbackground=TEXT, relief="flat", bd=4).pack(
                     side="left", fill="x", expand=True)

        def browse():
            path = filedialog.askopenfilename(
                title="Select " + label,
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
            if path:
                var.set(path)

        mk_button(row, text="Browse", command=browse, font=FONT_SMALL,
                  bg=ACCENT, fg="white", padx=8).pack(side="left", padx=(6, 0))
        return var

    def _build_inputs(self, parent):
        # CSV files
        self._section(parent, "CSV Files")
        tk.Label(parent, text="Jira Issues > Your filter > Export > CSV (all fields)",
                 font=FONT_SMALL, bg=BG, fg=TEXT_DIM,
                 wraplength=280, justify="left").pack(anchor="w", pady=(0, 4))

        tk.Label(parent, text="Issue Export CSV", font=FONT,
                 bg=BG, fg=TEXT).pack(anchor="w", pady=(4, 1))
        row = tk.Frame(parent, bg=BG)
        row.pack(fill="x", pady=(0, 4))
        self.primary_path = tk.StringVar()
        tk.Entry(row, textvariable=self.primary_path, font=FONT_SMALL,
                 bg=BG2, fg=TEXT, insertbackground=TEXT,
                 relief="flat", bd=4).pack(side="left", fill="x", expand=True)

        def browse_primary():
            path = filedialog.askopenfilename(
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
            if path:
                self.primary_path.set(path)

        mk_button(row, text="Browse", command=browse_primary,
                  font=FONT_SMALL, bg=ACCENT, fg="white",
                  padx=8).pack(side="left", padx=(6, 0))

        self.cycletime_path = self._file_row(parent, "Control Chart CSV", optional=True)

        # Simulation settings
        self._section(parent, "Simulation Settings")
        self._backlog_lbl = tk.StringVar(value="Backlog (no. of items)")
        self._period_lbl  = tk.StringVar(value="Weeks of history to use")
        self.backlog_var     = self._labelled_entry_dyn(parent, self._backlog_lbl, 50)
        self.period_var      = self._labelled_entry_dyn(parent, self._period_lbl, 16)
        self.simulations_var = self._labelled_entry(parent, "Simulations", 10000)

        today_str = date.today().strftime("%d/%m/%Y")
        self.start_date_var = self._labelled_entry(
            parent, "Start date (DD/MM/YYYY)", today_str, width=12)
        tk.Label(parent, text="  Leave blank to skip date forecast",
                 font=FONT_SMALL, bg=BG, fg=TEXT_DIM).pack(anchor="w")

        # ── Capacity adjustment ─────────────────────────
        self._section(parent, "Capacity Adjustment")
        tk.Label(parent,
                 text="Reduce effective throughput to account for\n"
                      "meetings, holidays, and overhead.",
                 font=FONT_SMALL, bg=BG, fg=TEXT_DIM,
                 wraplength=280, justify="left").pack(anchor="w", pady=(0, 6))

        cap_row = tk.Frame(parent, bg=BG)
        cap_row.pack(fill="x", pady=3)
        tk.Label(cap_row, text="Team availability (%)", font=FONT, bg=BG, fg=TEXT,
                 width=24, anchor="w").pack(side="left")
        self.capacity_var = tk.StringVar(value="80")
        tk.Entry(cap_row, textvariable=self.capacity_var, font=FONT_MONO, width=6,
                 bg=BG2, fg=TEXT, insertbackground=TEXT,
                 relief="flat", bd=4).pack(side="left")

        tk.Label(parent,
                 text="  80% = standard agile recommendation\n"
                      "  100% = no adjustment",
                 font=FONT_SMALL, bg=BG, fg=TEXT_DIM).pack(anchor="w")

        # Confidence levels
        self._section(parent, "Confidence Levels")
        self.conf_vars = {}
        cf = tk.Frame(parent, bg=BG)
        cf.pack(anchor="w", pady=4)
        for level in [50, 70, 85, 95]:
            v = tk.BooleanVar(value=True)
            self.conf_vars[level] = v
            tk.Checkbutton(cf, text="%d%%" % level, variable=v,
                           font=FONT, bg=BG, fg=CONF_COLS[level],
                           selectcolor=BG2, activebackground=BG,
                           relief="flat").pack(side="left", padx=4)

        tk.Frame(parent, bg=BG, height=10).pack()

        self.run_btn = mk_button(
            parent, text="Run Simulation",
            command=self._start_simulation,
            font=("Segoe UI", 11, "bold"),
            bg=ACCENT, fg="white", padx=12, pady=8)
        self.run_btn.pack(fill="x", pady=(4, 0))

        self.status_var = tk.StringVar(value="Ready.")
        tk.Label(parent, textvariable=self.status_var,
                 font=FONT_SMALL, bg=BG, fg=TEXT_DIM,
                 wraplength=280, justify="left").pack(anchor="w", pady=(6, 0))

        # Export buttons (shown after first run)
        tk.Frame(parent, bg=BG, height=6).pack()
        self._section(parent, "Export")
        self.save_chart_btn = mk_button(
            parent, text="Save Chart as PNG",
            command=self._save_chart,
            font=FONT_SMALL, bg=BG2, fg=TEXT, padx=8, pady=5)

        self.export_html_btn = mk_button(
            parent, text="Export Report as HTML",
            command=self._export_html,
            font=FONT_SMALL, bg=ACCENT2, fg=BG, padx=8, pady=5)

        self.export_pdf_btn = mk_button(
            parent, text="Export Report as PDF",
            command=self._export_pdf,
            font=FONT_SMALL, bg=ACCENT, fg="white", padx=8, pady=5)

        # Hide export buttons until a simulation has run
        self._export_buttons_shown = False

    def _show_export_buttons(self):
        if not self._export_buttons_shown:
            self.save_chart_btn.pack(fill="x", pady=(2, 0))
            self.export_html_btn.pack(fill="x", pady=(4, 0))
            self.export_pdf_btn.pack(fill="x", pady=(4, 0))
            self._export_buttons_shown = True

    def _build_results(self, parent):
        nb = ttk.Notebook(parent)
        nb.pack(fill="both", expand=True)

        style = ttk.Style()
        style.theme_use("default")
        style.configure("TNotebook", background=BG, borderwidth=0)
        style.configure("TNotebook.Tab", background=BG2, foreground=TEXT,
                        padding=[12, 5], font=FONT)
        style.map("TNotebook.Tab",
                  background=[("selected", ACCENT)],
                  foreground=[("selected", "white")])

        self.chart_frame = tk.Frame(nb, bg=BG)
        nb.add(self.chart_frame, text="  Chart  ")
        tk.Label(self.chart_frame,
                 text="Run a simulation to see the chart here.",
                 font=FONT, bg=BG, fg=TEXT_DIM).place(relx=0.5, rely=0.5, anchor="center")

        summary_frame = tk.Frame(nb, bg=BG)
        nb.add(summary_frame, text="  Summary  ")
        self.summary_text = scrolledtext.ScrolledText(
            summary_frame, font=FONT_MONO, bg=BG2, fg=TEXT,
            insertbackground=TEXT, relief="flat", bd=8,
            state="disabled", wrap="word")
        self.summary_text.pack(fill="both", expand=True, padx=4, pady=4)

        throughput_frame = tk.Frame(nb, bg=BG)
        nb.add(throughput_frame, text="  Weekly Throughput  ")
        self.throughput_text = scrolledtext.ScrolledText(
            throughput_frame, font=FONT_MONO, bg=BG2, fg=TEXT,
            insertbackground=TEXT, relief="flat", bd=8,
            state="disabled", wrap="word")
        self.throughput_text.pack(fill="both", expand=True, padx=4, pady=4)

        self._configure_tags()

    def _configure_tags(self):
        for t in [self.summary_text]:
            t.tag_config("heading", foreground=ACCENT,  font=("Consolas", 10, "bold"))
            t.tag_config("subhead", foreground=ACCENT2, font=("Consolas", 9, "bold"))
            t.tag_config("value",   foreground=TEXT)
            t.tag_config("good",    foreground=SUCCESS)
            t.tag_config("warn",    foreground=WARNING)
            t.tag_config("bad",     foreground=ERROR)
            t.tag_config("dim",     foreground=TEXT_DIM)
            for level, col in CONF_COLS.items():
                t.tag_config("conf%d" % level, foreground=col)

    # ── Simulation ─────────────────────────────────────

    def _start_simulation(self):
        self.run_btn.config(state="disabled", text="Running...")
        self.status_var.set("Running simulation...")
        threading.Thread(target=self._run, daemon=True).start()

    def _run(self):
        try:
            self._run_kanban()
        except Exception as ex:
            import traceback
            traceback.print_exc()
            self.after(0, lambda e=ex: self._show_error(str(e)))
        finally:
            self.after(0, lambda: self.run_btn.config(
                state="normal", text="Run Simulation"))

    def _run_kanban(self):
        path = self.primary_path.get().strip()
        if not path:
            raise ValueError("Please select an Issue Export CSV file.")
        if not os.path.exists(path):
            raise ValueError("File not found:\n%s" % path)

        backlog, weeks_history, n_sims, conf_levels, start_date, capacity = \
            self._parse_inputs()

        self.after(0, lambda: self.status_var.set("Calculating weekly throughput..."))
        tp_samples, date_col, n_issues, weekly_counts = load_kanban_throughput(path)

        if len(tp_samples) < 3:
            raise ValueError("Only %d weeks of data found.\n"
                             "Need at least 3 weeks of history." % len(tp_samples))

        if weeks_history > 0 and len(tp_samples) > weeks_history:
            tp_samples_used = tp_samples[-weeks_history:]
        else:
            tp_samples_used = tp_samples

        # Apply capacity adjustment
        adjusted_samples = tp_samples_used * (capacity / 100.0)

        ct_samples, ct_col = load_cycle_time(self.cycletime_path.get().strip())

        self.after(0, lambda: self.status_var.set(
            "Running %s simulations (%.0f%% capacity)..." % (
                "{:,}".format(n_sims), capacity)))

        week_r       = mc_weeks_to_done(adjusted_samples, backlog, n_sims)
        median_weeks = int(np.median(week_r))
        through_r    = mc_throughput_in_periods(adjusted_samples, median_weeks, n_sims)

        self.after(0, lambda: self._render_kanban(
            week_r, through_r, adjusted_samples, tp_samples_used,
            weekly_counts, ct_samples, ct_col, date_col,
            backlog, n_issues, n_sims, conf_levels,
            median_weeks, start_date, capacity
        ))

    def _parse_inputs(self):
        try:
            backlog   = int(self.backlog_var.get())
            period    = int(self.period_var.get())
            n_sims    = int(self.simulations_var.get())
            capacity  = float(self.capacity_var.get())
        except ValueError:
            raise ValueError("Backlog, period, simulations, and capacity must be numbers.")
        if backlog < 1:
            raise ValueError("Backlog must be at least 1.")
        if n_sims < 100:
            raise ValueError("Please use at least 100 simulations.")
        if not 1 <= capacity <= 100:
            raise ValueError("Capacity must be between 1 and 100.")
        conf_levels = [l for l, v in self.conf_vars.items() if v.get()]
        if not conf_levels:
            raise ValueError("Select at least one confidence level.")
        start_date = self._parse_start_date()
        return backlog, period, n_sims, conf_levels, start_date, capacity

    def _parse_start_date(self):
        raw = self.start_date_var.get().strip()
        if not raw:
            return None
        for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(raw, fmt)
            except ValueError:
                continue
        raise ValueError("Cannot parse start date '%s'.\nUse DD/MM/YYYY." % raw)

    def _finish_dates(self, percentiles, start_date, days_per_unit):
        if start_date is None:
            return {}
        return {label: (start_date + timedelta(days=int(periods) * days_per_unit)).strftime("%d %b %Y")
                for label, periods in percentiles.items()}

    # ── Render ─────────────────────────────────────────

    def _render_kanban(self, week_r, through_r, adj_samples, raw_samples,
                       weekly_counts, ct_samples, ct_col, date_col,
                       backlog, n_issues, n_sims, conf_levels,
                       median_weeks, start_date, capacity):

        # Store for report export
        self._last_results = dict(
            week_r=week_r, through_r=through_r,
            adj_samples=adj_samples, raw_samples=raw_samples,
            weekly_counts=weekly_counts, ct_samples=ct_samples,
            ct_col=ct_col, date_col=date_col, backlog=backlog,
            n_issues=n_issues, n_sims=n_sims, conf_levels=conf_levels,
            median_weeks=median_weeks, start_date=start_date, capacity=capacity
        )

        self._draw_chart(week_r, through_r, backlog, conf_levels,
                         median_weeks, n_sims, capacity)
        self._write_summary(week_r, through_r, adj_samples, ct_samples,
                            ct_col, date_col, backlog, n_issues, n_sims,
                            conf_levels, median_weeks, start_date, capacity)
        self._write_throughput_tab(weekly_counts, raw_samples, adj_samples, capacity)

        self.status_var.set("Done. %s simulations | %d weeks history | %.0f%% capacity." % (
            "{:,}".format(n_sims), len(raw_samples), capacity))
        self._show_export_buttons()

    # ── Chart ──────────────────────────────────────────

    def _draw_chart(self, week_r, through_r, backlog, conf_levels,
                    median_weeks, n_sims, capacity, fig_size=(9, 4.2)):
        for w in self.chart_frame.winfo_children():
            w.destroy()
        fig = self._make_chart_fig(week_r, through_r, backlog, conf_levels,
                                   median_weeks, n_sims, capacity, fig_size)
        canvas = FigureCanvasTkAgg(fig, master=self.chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        self._fig = fig

    def _make_chart_fig(self, week_r, through_r, backlog, conf_levels,
                        median_weeks, n_sims, capacity, fig_size=(9, 4.2)):
        fig, axes = plt.subplots(1, 2, figsize=fig_size)
        fig.patch.set_facecolor("#1e1e2e")
        plt.subplots_adjust(wspace=0.35, left=0.08, right=0.97, top=0.85, bottom=0.12)
        cap_note = "  |  Capacity: %.0f%%" % capacity if capacity != 100 else ""
        fig.suptitle(
            "Monte Carlo Projection  —  %s simulations  |  Backlog: %d items%s" % (
                "{:,}".format(n_sims), backlog, cap_note),
            color=TEXT, fontsize=11, fontweight="bold")

        for ax in axes:
            ax.set_facecolor("#2a2a3e")
            ax.tick_params(colors=TEXT_DIM, labelsize=8)
            for spine in ax.spines.values():
                spine.set_color("#444466")
            ax.grid(axis="y", color="#333355", linewidth=0.5)
            ax.yaxis.label.set_color(TEXT_DIM)
            ax.xaxis.label.set_color(TEXT_DIM)
            ax.title.set_color(TEXT)

        ax1, ax2 = axes
        bins = list(range(int(week_r.min()), int(week_r.max()) + 2))
        ax1.hist(week_r, bins=bins, color=ACCENT, alpha=0.75,
                 edgecolor="#1e1e2e", linewidth=0.4)
        for cl in conf_levels:
            pct = np.percentile(week_r, cl)
            ax1.axvline(pct, color=CONF_COLS[cl], linestyle="--", linewidth=1.5,
                        label="%d%% <= %.0f wks" % (cl, pct))
        ax1.set_xlabel("Weeks to clear backlog", fontsize=9)
        ax1.set_ylabel("Simulations", fontsize=9)
        ax1.set_title("How many weeks?", fontsize=10)
        ax1.legend(fontsize=8, facecolor="#2a2a3e", edgecolor="#444466", labelcolor=TEXT)
        ax1.xaxis.set_major_locator(mticker.MaxNLocator(integer=True))

        ax2.hist(through_r, bins=40, color=ACCENT2, alpha=0.75,
                 edgecolor="#1e1e2e", linewidth=0.4)
        for cl in conf_levels:
            pct = np.percentile(through_r, cl)
            ax2.axvline(pct, color=CONF_COLS[cl], linestyle="--", linewidth=1.5,
                        label="%d%% >= %.0f items" % (cl, pct))
        ax2.axvline(backlog, color="white", linestyle="-", linewidth=1.5,
                    label="Backlog (%d)" % backlog)
        ax2.set_xlabel("Items completed", fontsize=9)
        ax2.set_ylabel("Simulations", fontsize=9)
        ax2.set_title("Throughput in %d weeks" % median_weeks, fontsize=10)
        ax2.legend(fontsize=8, facecolor="#2a2a3e", edgecolor="#444466", labelcolor=TEXT)
        return fig

    # ── Summary tab ────────────────────────────────────

    def _write_summary(self, week_r, through_r, adj_samples, ct_samples,
                       ct_col, date_col, backlog, n_issues, n_sims,
                       conf_levels, median_weeks, start_date, capacity):
        st = self.summary_text
        st.config(state="normal")
        st.delete("1.0", "end")

        def line(text="", tag="value"):
            st.insert("end", text + "\n", tag)

        line("=" * 54, "dim")
        line(" KANBAN — MONTE CARLO RESULTS  v%s" % VERSION, "heading")
        line("=" * 54, "dim")
        line()
        line("  Generated        : %s" % datetime.now().strftime("%d %b %Y  %H:%M"), "dim")
        line("  Resolved col     : %s" % date_col, "dim")
        line("  Issues analysed  : %d" % n_issues, "dim")
        line("  Weeks of history : %d" % len(adj_samples), "dim")
        line("  Throughput range : %.0f – %.0f items/week  (mean %.1f)" % (
            adj_samples.min(), adj_samples.max(), adj_samples.mean()), "dim")
        line("  Backlog          : %d items" % backlog, "dim")
        line("  Capacity         : %.0f%%" % capacity, "dim")
        line("  Simulations      : %s" % "{:,}".format(n_sims), "dim")
        line()
        line("-" * 54, "dim")
        line("  WEEKS TO CLEAR %d-ITEM BACKLOG" % backlog, "subhead")
        line("-" * 54, "dim")
        line("  Mean   : %.1f weeks" % week_r.mean())
        line("  Median : %.0f weeks" % np.median(week_r))
        line()
        for cl in conf_levels:
            pct = np.percentile(week_r, cl)
            tag = "good" if cl <= 70 else ("warn" if cl <= 85 else "bad")
            line("  %d%% confidence  <=  %.0f weeks" % (cl, pct), tag)

        if start_date is not None:
            line()
            line("-" * 54, "dim")
            line("  FORECAST FINISH DATES  (from %s)" %
                 start_date.strftime("%d %b %Y"), "subhead")
            line("-" * 54, "dim")
            for cl in conf_levels:
                pct = np.percentile(week_r, cl)
                d = self._finish_dates({cl: pct}, start_date, 7)
                tag = "good" if cl <= 70 else ("warn" if cl <= 85 else "bad")
                line("  %d%% confidence by  :  %s" % (cl, d[cl]), tag)

        line()
        line("-" * 54, "dim")
        line("  THROUGHPUT IN %d WEEKS (MEDIAN)" % median_weeks, "subhead")
        line("-" * 54, "dim")
        line("  Mean   : %.0f items" % through_r.mean())
        line("  Median : %.0f items" % np.median(through_r))
        line()
        for cl in conf_levels:
            pct = np.percentile(through_r, cl)
            line("  %d%%ile  >=  %.0f items  (%.0f%% of backlog)" % (
                cl, pct, (pct / backlog) * 100), "conf%d" % cl)

        if ct_samples is not None:
            line()
            line("-" * 54, "dim")
            line("  CYCLE TIME  ('%s')" % ct_col, "subhead")
            line("-" * 54, "dim")
            line("  Median : %.1f days" % np.median(ct_samples))
            line("  85th % : %.1f days" % np.percentile(ct_samples, 85))
            line("  95th % : %.1f days" % np.percentile(ct_samples, 95))

        line()
        line("=" * 54, "dim")
        st.config(state="disabled")

    # ── Throughput tab ─────────────────────────────────

    def _write_throughput_tab(self, weekly_counts, raw_samples, adj_samples, capacity):
        tt = self.throughput_text
        tt.config(state="normal")
        tt.delete("1.0", "end")
        tt.insert("end", "  WEEKLY THROUGHPUT BREAKDOWN\n")
        tt.insert("end", "  (items completed per week — raw data)\n\n")
        tt.insert("end", "  %-22s  %s\n" % ("Week", "Items"))
        tt.insert("end", "  " + "-" * 32 + "\n")
        for period, count in weekly_counts.items():
            bar = "#" * int(count)
            tt.insert("end", "  %-22s  %3d  %s\n" % (str(period), count, bar))
        tt.insert("end", "\n  " + "-" * 32 + "\n")
        tt.insert("end", "  RAW THROUGHPUT\n")
        tt.insert("end", "  Mean   : %.1f items/week\n" % raw_samples.mean())
        tt.insert("end", "  Median : %.1f items/week\n" % np.median(raw_samples))
        tt.insert("end", "  Min : %.0f  |  Max : %.0f\n" % (
            raw_samples.min(), raw_samples.max()))
        if capacity != 100:
            tt.insert("end", "\n  ADJUSTED THROUGHPUT (%.0f%% capacity)\n" % capacity)
            tt.insert("end", "  Mean   : %.1f items/week\n" % adj_samples.mean())
            tt.insert("end", "  Median : %.1f items/week\n" % np.median(adj_samples))
        tt.config(state="disabled")

    # ── Save chart PNG ─────────────────────────────────

    def _save_chart(self):
        if self._fig is None:
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG image", "*.png")],
            initialfile="monte_carlo_chart.png",
            title="Save chart")
        if path:
            self._fig.savefig(path, dpi=150, bbox_inches="tight",
                              facecolor=self._fig.get_facecolor())
            messagebox.showinfo("Saved", "Chart saved to:\n%s" % path)

    # ── Report helpers ─────────────────────────────────

    def _build_summary_lines(self, r):
        """Build a list of (text, colour_hex) tuples for report output."""
        lines = []
        week_r       = r["week_r"]
        through_r    = r["through_r"]
        adj_samples  = r["adj_samples"]
        raw_samples  = r["raw_samples"]
        ct_samples   = r["ct_samples"]
        ct_col       = r["ct_col"]
        date_col     = r["date_col"]
        backlog      = r["backlog"]
        n_issues     = r["n_issues"]
        n_sims       = r["n_sims"]
        conf_levels  = r["conf_levels"]
        median_weeks = r["median_weeks"]
        start_date   = r["start_date"]
        capacity     = r["capacity"]

        dim   = "#9090b0"
        white = "#e0e0f0"

        def add(text, col=white):
            lines.append((text, col))

        add("Generated        : %s" % datetime.now().strftime("%d %b %Y  %H:%M"), dim)
        add("Resolved col     : %s" % date_col, dim)
        add("Issues analysed  : %d" % n_issues, dim)
        add("Weeks of history : %d" % len(adj_samples), dim)
        add("Throughput range : %.0f – %.0f items/week  (mean %.1f)" % (
            adj_samples.min(), adj_samples.max(), adj_samples.mean()), dim)
        add("Backlog          : %d items" % backlog, dim)
        add("Capacity         : %.0f%%" % capacity, dim)
        add("Simulations      : %s" % "{:,}".format(n_sims), dim)
        add("")
        add("WEEKS TO CLEAR %d-ITEM BACKLOG" % backlog, "#26c6da")
        add("Mean   : %.1f weeks" % week_r.mean())
        add("Median : %.0f weeks" % np.median(week_r))
        add("")
        for cl in conf_levels:
            pct = np.percentile(week_r, cl)
            col = CONF_HEX[cl]
            add("%d%% confidence  <=  %.0f weeks" % (cl, pct), col)

        if start_date is not None:
            add("")
            add("FORECAST FINISH DATES  (from %s)" %
                start_date.strftime("%d %b %Y"), "#26c6da")
            for cl in conf_levels:
                pct = np.percentile(week_r, cl)
                d = self._finish_dates({cl: pct}, start_date, 7)
                add("%d%% confidence by  :  %s" % (cl, d[cl]), CONF_HEX[cl])

        add("")
        add("THROUGHPUT IN %d WEEKS (MEDIAN)" % median_weeks, "#26c6da")
        add("Mean   : %.0f items" % through_r.mean())
        add("Median : %.0f items" % np.median(through_r))
        add("")
        for cl in conf_levels:
            pct = np.percentile(through_r, cl)
            add("%d%%ile  >=  %.0f items  (%.0f%% of backlog)" % (
                cl, pct, (pct / backlog) * 100), CONF_HEX[cl])

        if ct_samples is not None:
            add("")
            add("CYCLE TIME  ('%s')" % ct_col, "#26c6da")
            add("Median : %.1f days" % np.median(ct_samples))
            add("85th % : %.1f days" % np.percentile(ct_samples, 85))
            add("95th % : %.1f days" % np.percentile(ct_samples, 95))

        return lines

    def _build_throughput_lines(self, r):
        lines = []
        weekly_counts = r["weekly_counts"]
        raw_samples   = r["raw_samples"]
        adj_samples   = r["adj_samples"]
        capacity      = r["capacity"]
        dim = "#9090b0"
        white = "#e0e0f0"

        for period, count in weekly_counts.items():
            bar = "█" * int(count)
            lines.append(("%-22s  %3d  %s" % (str(period), count, bar), white))

        lines.append(("", white))
        lines.append(("Raw throughput — Mean: %.1f  Median: %.1f  Min: %.0f  Max: %.0f" % (
            raw_samples.mean(), np.median(raw_samples),
            raw_samples.min(), raw_samples.max()), dim))
        if capacity != 100:
            lines.append(("Adjusted (%.0f%%) — Mean: %.1f  Median: %.1f" % (
                capacity, adj_samples.mean(), np.median(adj_samples)), dim))
        return lines

    # ── HTML export ─────────────────────────────────────

    def _export_html(self):
        if self._last_results is None:
            messagebox.showwarning("No Data", "Run a simulation first.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".html",
            filetypes=[("HTML file", "*.html")],
            initialfile="monte_carlo_report.html",
            title="Export HTML Report")
        if not path:
            return

        r = self._last_results

        # Render chart to base64 PNG for embedding
        fig = self._make_chart_fig(
            r["week_r"], r["through_r"], r["backlog"], r["conf_levels"],
            r["median_weeks"], r["n_sims"], r["capacity"], fig_size=(10, 4.5))
        buf = io.BytesIO()
        fig.savefig(buf, format="png", dpi=120, bbox_inches="tight",
                    facecolor=fig.get_facecolor())
        plt.close(fig)
        chart_b64 = base64.b64encode(buf.getvalue()).decode("utf-8")

        summary_lines   = self._build_summary_lines(r)
        throughput_lines = self._build_throughput_lines(r)

        def html_rows(lines):
            rows = []
            for text, col in lines:
                if not text:
                    rows.append('<tr><td colspan="1" style="padding:4px 0;">&nbsp;</td></tr>')
                else:
                    rows.append('<tr><td style="font-family:Consolas,monospace;font-size:13px;'
                                'color:%s;padding:2px 8px;white-space:pre;">%s</td></tr>' % (
                                    col, text.replace("&", "&amp;").replace("<", "&lt;")))
            return "\n".join(rows)

        html = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Monte Carlo Projection Report</title>
<style>
  body {{ background:#1e1e2e; color:#e0e0f0; font-family:'Segoe UI',sans-serif; margin:0; padding:24px; }}
  h1   {{ color:#7c6af7; font-size:22px; margin-bottom:4px; }}
  h2   {{ color:#26c6da; font-size:15px; margin-top:28px; margin-bottom:8px;
          border-bottom:1px solid #444466; padding-bottom:6px; }}
  .sub {{ color:#9090b0; font-size:12px; margin-bottom:20px; }}
  .chart {{ max-width:100%%; border-radius:8px; margin:16px 0; }}
  table {{ border-collapse:collapse; width:100%%; }}
  .footer {{ color:#555577; font-size:11px; margin-top:32px; border-top:1px solid #333355;
             padding-top:12px; }}
</style>
</head>
<body>
<h1>Monte Carlo Projection Report</h1>
<div class="sub">Generated {generated} &nbsp;|&nbsp; Monte Carlo Projection Tool v{version}</div>

<h2>Simulation Chart</h2>
<img class="chart" src="data:image/png;base64,{chart}" alt="Monte Carlo Chart">

<h2>Summary</h2>
<table>{summary_rows}</table>

<h2>Weekly Throughput Breakdown</h2>
<table>{throughput_rows}</table>

<div class="footer">
  Monte Carlo Projection Tool v{version} &nbsp;|&nbsp;
  All data processed locally — no data was transmitted externally.
</div>
</body>
</html>""".format(
            generated=datetime.now().strftime("%d %b %Y  %H:%M"),
            version=VERSION,
            chart=chart_b64,
            summary_rows=html_rows(summary_lines),
            throughput_rows=html_rows(throughput_lines)
        )

        with open(path, "w", encoding="utf-8") as f:
            f.write(html)
        messagebox.showinfo("Exported", "HTML report saved to:\n%s" % path)

    # ── PDF export ──────────────────────────────────────

    def _export_pdf(self):
        if self._last_results is None:
            messagebox.showwarning("No Data", "Run a simulation first.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF file", "*.pdf")],
            initialfile="monte_carlo_report.pdf",
            title="Export PDF Report")
        if not path:
            return

        r = self._last_results

        with PdfPages(path) as pdf:
            # Page 1 — Chart
            fig1 = self._make_chart_fig(
                r["week_r"], r["through_r"], r["backlog"], r["conf_levels"],
                r["median_weeks"], r["n_sims"], r["capacity"], fig_size=(11, 5))
            fig1.text(0.5, 0.97,
                      "Monte Carlo Projection Report  —  v%s  —  %s" % (
                          VERSION, datetime.now().strftime("%d %b %Y %H:%M")),
                      ha="center", va="top", color=TEXT_DIM, fontsize=9,
                      transform=fig1.transFigure)
            pdf.savefig(fig1, bbox_inches="tight", facecolor=fig1.get_facecolor())
            plt.close(fig1)

            # Page 2 — Summary
            summary_lines = self._build_summary_lines(r)
            fig2 = self._make_text_page(
                "Summary", summary_lines,
                "Monte Carlo Projection Report  —  v%s  —  %s" % (
                    VERSION, datetime.now().strftime("%d %b %Y %H:%M")))
            pdf.savefig(fig2, bbox_inches="tight", facecolor=fig2.get_facecolor())
            plt.close(fig2)

            # Page 3 — Weekly Throughput
            tp_lines = self._build_throughput_lines(r)
            fig3 = self._make_text_page(
                "Weekly Throughput Breakdown", tp_lines,
                "Monte Carlo Projection Report  —  v%s  —  %s" % (
                    VERSION, datetime.now().strftime("%d %b %Y %H:%M")))
            pdf.savefig(fig3, bbox_inches="tight", facecolor=fig3.get_facecolor())
            plt.close(fig3)

        messagebox.showinfo("Exported", "PDF report saved to:\n%s" % path)

    def _make_text_page(self, title, lines, footer):
        fig = plt.figure(figsize=(11, 8.5))
        fig.patch.set_facecolor("#1e1e2e")
        fig.text(0.06, 0.96, title, fontsize=14, fontweight="bold",
                 color="#7c6af7", va="top", transform=fig.transFigure)
        fig.text(0.5, 0.97, footer, ha="center", va="top",
                 color=TEXT_DIM, fontsize=8, transform=fig.transFigure)

        y = 0.90
        line_h = 0.028
        for text, col in lines:
            if y < 0.04:
                break
            fig.text(0.06, y, text if text else "",
                     fontsize=9, color=col, va="top",
                     fontfamily="monospace", transform=fig.transFigure)
            y -= line_h

        fig.text(0.5, 0.02,
                 "Monte Carlo Projection Tool v%s  |  All data processed locally" % VERSION,
                 ha="center", va="bottom", color="#555577", fontsize=8,
                 transform=fig.transFigure)
        return fig

    # ── Error ──────────────────────────────────────────

    def _show_error(self, msg):
        self.status_var.set("Error — see message.")
        messagebox.showerror("Simulation Error", msg)


# ══════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════

if __name__ == "__main__":
    print("Starting Monte Carlo Projection Tool v%s..." % VERSION)
    try:
        app = MonteCarloApp()
        print("Entering main loop...")
        app.mainloop()
    except Exception as e:
        import traceback
        traceback.print_exc()
        input("Press Enter to exit...")
