"""
Monte Carlo Sprint/Flow Projection Tool
========================================
Supports both Scrum (velocity-based) and Kanban (throughput-based) teams.
Runs entirely locally — no data is sent anywhere.

REQUIREMENTS:
    pip install pandas numpy matplotlib
    (tkinter is built into Python — no install needed)

USAGE:
    python monte_carlo_jira.py
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.ticker as mticker

# ── Colours ───────────────────────────────────────────
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

FONT       = ("Segoe UI", 10)
FONT_BOLD  = ("Segoe UI", 10, "bold")
FONT_TITLE = ("Segoe UI", 13, "bold")
FONT_SMALL = ("Segoe UI", 9)
FONT_MONO  = ("Consolas", 9)


# ══════════════════════════════════════════════════════
# DATA HELPERS
# ══════════════════════════════════════════════════════

def load_velocity(path):
    """Load completed story points per sprint (Scrum mode)."""
    df = pd.read_csv(path)
    df.columns = df.columns.str.strip()
    candidates = [c for c in df.columns
                  if "complet" in c.lower() or "done" in c.lower()]
    if not candidates:
        numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
        candidates = numeric_cols[1:2] if len(numeric_cols) >= 2 else numeric_cols[:1]
    if not candidates:
        raise ValueError(
            "Cannot find a 'completed' column in the velocity CSV.\n"
            "Columns found: " + str(list(df.columns))
        )
    col = candidates[0]
    values = pd.to_numeric(df[col], errors="coerce").dropna()
    values = values[values > 0]
    return values.to_numpy(), col


def load_kanban_throughput(path):
    """
    Load a Jira issue export (filter export) and calculate
    weekly throughput (items completed per week).
    Looks for a resolved/completed date column.
    """
    df = pd.read_csv(path, low_memory=False)
    df.columns = df.columns.str.strip()

    # Find resolved date column
    date_candidates = [c for c in df.columns
                       if any(k in c.lower() for k in
                              ["resolved", "resolutiondate", "done date",
                               "completed", "closed"])]
    if not date_candidates:
        raise ValueError(
            "Cannot find a resolved/completed date column.\n"
            "Columns found: " + str(list(df.columns)[:20])
        )
    date_col = date_candidates[0]

    dates = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True)
    dates = dates.dropna()

    if len(dates) == 0:
        raise ValueError(
            "No valid dates found in column '%s'.\n"
            "Check the CSV contains resolved dates." % date_col
        )

    # Group by ISO week and count items
    week_series = dates.dt.to_period("W")
    weekly_counts = week_series.value_counts().sort_index()

    # Fill gaps with 0 for weeks with no completions
    if len(weekly_counts) > 1:
        full_range = pd.period_range(
            weekly_counts.index.min(),
            weekly_counts.index.max(),
            freq="W"
        )
        weekly_counts = weekly_counts.reindex(full_range, fill_value=0)

    throughput = weekly_counts.values.astype(float)
    throughput = throughput[throughput >= 0]  # keep zeros (low weeks are real)

    return throughput, date_col, len(dates), weekly_counts


def load_cycle_time(path):
    """Load cycle time data from control chart CSV (optional)."""
    if not path or not os.path.exists(path):
        return None, None
    df = pd.read_csv(path)
    df.columns = df.columns.str.strip()
    candidates = [c for c in df.columns
                  if "cycle" in c.lower() or "lead" in c.lower()
                  or ("days" in c.lower() and "time" in c.lower())]
    if not candidates:
        # fallback — any column with 'days'
        candidates = [c for c in df.columns if "days" in c.lower()]
    if not candidates:
        return None, None
    col = candidates[0]
    values = pd.to_numeric(df[col], errors="coerce").dropna()
    values = values[values > 0]
    return values.to_numpy(), col


# ── Monte Carlo engines ────────────────────────────────

def mc_sprints_to_done(samples, backlog, n):
    """Scrum: how many sprints to clear backlog of story points."""
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
    """How much will be completed in a fixed number of periods."""
    return np.array([
        np.random.choice(samples, size=periods).sum()
        for _ in range(n)
    ])


def mc_weeks_to_done(samples, backlog_items, n):
    """Kanban: how many weeks to clear backlog of items."""
    results = np.zeros(n, dtype=int)
    for i in range(n):
        remaining = backlog_items
        count = 0
        while remaining > 0:
            remaining -= np.random.choice(samples)
            count += 1
            if count > 2000:
                break
        results[i] = count
    return results


# ══════════════════════════════════════════════════════
# APPLICATION
# ══════════════════════════════════════════════════════

class MonteCarloApp(tk.Tk):

    def __init__(self):
        print("Initialising...")
        super().__init__()
        self.title("Monte Carlo Projection Tool")
        self.configure(bg=BG)
        self.resizable(True, True)
        self.minsize(920, 660)
        self._fig = None
        self._mode = tk.StringVar(value="scrum")
        print("Building UI...")
        self._build_ui()
        print("Centring window...")
        self._centre_window(1000, 760)
        print("Ready.")

    def _centre_window(self, w, h):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = max(0, (sw - w) // 2)
        y = max(0, (sh - h) // 2)
        self.geometry("%dx%d+%d+%d" % (w, h, x, y))

    # ── Layout ─────────────────────────────────────────

    def _build_ui(self):
        # Header
        hdr = tk.Frame(self, bg=ACCENT, pady=10)
        hdr.pack(fill="x")
        tk.Label(hdr, text="Monte Carlo Projection Tool",
                 font=FONT_TITLE, bg=ACCENT, fg="white").pack()
        tk.Label(hdr, text="Supports Scrum (velocity) and Kanban (throughput)  |  "
                            "All processing is local — no data leaves your machine",
                 font=FONT_SMALL, bg=ACCENT, fg="#ddd").pack()

        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=14, pady=12)

        left_outer = tk.Frame(body, bg=BG, width=320)
        left_outer.pack(side="left", fill="y", padx=(0, 12))
        left_outer.pack_propagate(False)

        # Scrollable canvas for left panel
        left_canvas = tk.Canvas(left_outer, bg=BG, highlightthickness=0)
        left_scroll = tk.Scrollbar(left_outer, orient="vertical",
                                   command=left_canvas.yview)
        left_canvas.configure(yscrollcommand=left_scroll.set)
        left_scroll.pack(side="right", fill="y")
        left_canvas.pack(side="left", fill="both", expand=True)

        left = tk.Frame(left_canvas, bg=BG)
        left_window = left_canvas.create_window((0, 0), window=left, anchor="nw")

        def on_frame_configure(event):
            left_canvas.configure(scrollregion=left_canvas.bbox("all"))

        def on_canvas_configure(event):
            left_canvas.itemconfig(left_window, width=event.width)

        left.bind("<Configure>", on_frame_configure)
        left_canvas.bind("<Configure>", on_canvas_configure)

        # Mouse wheel scrolling
        def on_mousewheel(event):
            left_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        left_canvas.bind_all("<MouseWheel>", on_mousewheel)

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
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
            )
            if path:
                var.set(path)

        tk.Button(row, text="Browse", command=browse, font=FONT_SMALL,
                  bg=ACCENT, fg="white", relief="flat",
                  padx=8, cursor="hand2").pack(side="left", padx=(6, 0))
        return var

    def _build_inputs(self, parent):
        # Mode selector
        self._section(parent, "Team Type")
        mode_frame = tk.Frame(parent, bg=BG)
        mode_frame.pack(fill="x", pady=(0, 4))

        def on_mode_change():
            self._update_mode_ui()

        tk.Radiobutton(
            mode_frame, text="Scrum  (velocity CSV)",
            variable=self._mode, value="scrum",
            command=on_mode_change,
            font=FONT, bg=BG, fg=TEXT, selectcolor=BG2,
            activebackground=BG
        ).pack(anchor="w")
        tk.Radiobutton(
            mode_frame, text="Kanban  (issue export CSV)",
            variable=self._mode, value="kanban",
            command=on_mode_change,
            font=FONT, bg=BG, fg=TEXT, selectcolor=BG2,
            activebackground=BG
        ).pack(anchor="w")

        # CSV files
        self._section(parent, "CSV Files")
        self._csv_hint_var = tk.StringVar(
            value="Jira Board > Reports > Velocity Chart > Export")
        self._csv_hint_lbl = tk.Label(
            parent, textvariable=self._csv_hint_var,
            font=FONT_SMALL, bg=BG, fg=TEXT_DIM, wraplength=280, justify="left")
        self._csv_hint_lbl.pack(anchor="w", pady=(0, 4))

        self._primary_label_var = tk.StringVar(value="Velocity CSV")
        tk.Label(parent, textvariable=self._primary_label_var,
                 font=FONT, bg=BG, fg=TEXT).pack(anchor="w", pady=(4, 1))
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

        tk.Button(row, text="Browse", command=browse_primary, font=FONT_SMALL,
                  bg=ACCENT, fg="white", relief="flat",
                  padx=8, cursor="hand2").pack(side="left", padx=(6, 0))

        self.cycletime_path = self._file_row(parent, "Control Chart CSV", optional=True)

        # Settings
        self._section(parent, "Simulation Settings")
        self._backlog_label_var = tk.StringVar(value="Backlog (story points)")
        self._period_label_var  = tk.StringVar(value="Sprint length (days)")

        self.backlog_var     = self._labelled_entry_dynamic(
            parent, self._backlog_label_var, 150)
        self.period_var      = self._labelled_entry_dynamic(
            parent, self._period_label_var, 14)
        self.simulations_var = self._labelled_entry(parent, "Simulations", 10000)

        # Start date
        from datetime import date
        today_str = date.today().strftime("%d/%m/%Y")
        self.start_date_var = self._labelled_entry(
            parent, "Start date (DD/MM/YYYY)", today_str, width=12)
        tk.Label(parent, text="  Leave blank to skip date forecast",
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
        self.run_btn = tk.Button(
            parent, text="Run Simulation",
            command=self._start_simulation,
            font=("Segoe UI", 11, "bold"),
            bg=ACCENT, fg="white", relief="flat",
            padx=12, pady=8, cursor="hand2"
        )
        self.run_btn.pack(fill="x", pady=(4, 0))

        self.status_var = tk.StringVar(value="Ready.")
        tk.Label(parent, textvariable=self.status_var,
                 font=FONT_SMALL, bg=BG, fg=TEXT_DIM,
                 wraplength=280, justify="left").pack(anchor="w", pady=(6, 0))

        self.save_btn = tk.Button(
            parent, text="Save Chart as PNG",
            command=self._save_chart,
            font=FONT_SMALL, bg=BG2, fg=TEXT,
            relief="flat", padx=8, pady=5, cursor="hand2"
        )

    def _parse_start_date(self):
        """Parse start date field. Returns datetime or None."""
        from datetime import datetime
        raw = self.start_date_var.get().strip()
        if not raw:
            return None
        for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(raw, fmt)
            except ValueError:
                continue
        raise ValueError(
            "Cannot parse start date '%s'.\n"
            "Please use DD/MM/YYYY format, e.g. 09/04/2026" % raw)

    def _finish_dates(self, percentiles, start_date, days_per_unit):
        """
        Given a dict of {label: periods}, a start date, and days per period,
        return a dict of {label: date_string}.
        """
        from datetime import timedelta
        if start_date is None:
            return {}
        results = {}
        for label, periods in percentiles.items():
            finish = start_date + timedelta(days=int(periods) * days_per_unit)
            results[label] = finish.strftime("%d %b %Y")
        return results

    def _labelled_entry_dynamic(self, parent, label_var, default, width=10):
        row = tk.Frame(parent, bg=BG)
        row.pack(fill="x", pady=3)
        tk.Label(row, textvariable=label_var, font=FONT, bg=BG, fg=TEXT,
                 width=24, anchor="w").pack(side="left")
        var = tk.StringVar(value=str(default))
        tk.Entry(row, textvariable=var, font=FONT_MONO, width=width,
                 bg=BG2, fg=TEXT, insertbackground=TEXT,
                 relief="flat", bd=4).pack(side="left")
        return var

    def _update_mode_ui(self):
        if self._mode.get() == "scrum":
            self._primary_label_var.set("Velocity CSV")
            self._csv_hint_var.set(
                "Jira Board > Reports > Velocity Chart > Export")
            self._backlog_label_var.set("Backlog (story points)")
            self._period_label_var.set("Sprint length (days)")
            self.period_var.set("14")
        else:
            self._primary_label_var.set("Issue Export CSV")
            self._csv_hint_var.set(
                "Jira Issues > Your filter > Export > CSV (all fields)")
            self._backlog_label_var.set("Backlog (no. of items)")
            self._period_label_var.set("Weeks of history to use")
            self.period_var.set("16")

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
                 font=FONT, bg=BG, fg=TEXT_DIM).place(
                     relx=0.5, rely=0.5, anchor="center")

        summary_frame = tk.Frame(nb, bg=BG)
        nb.add(summary_frame, text="  Summary  ")
        self.summary_text = scrolledtext.ScrolledText(
            summary_frame, font=FONT_MONO,
            bg=BG2, fg=TEXT, insertbackground=TEXT,
            relief="flat", bd=8, state="disabled", wrap="word"
        )
        self.summary_text.pack(fill="both", expand=True, padx=4, pady=4)
        self._configure_tags()

        throughput_frame = tk.Frame(nb, bg=BG)
        nb.add(throughput_frame, text="  Weekly Throughput  ")
        self.throughput_text = scrolledtext.ScrolledText(
            throughput_frame, font=FONT_MONO,
            bg=BG2, fg=TEXT, insertbackground=TEXT,
            relief="flat", bd=8, state="disabled", wrap="word"
        )
        self.throughput_text.pack(fill="both", expand=True, padx=4, pady=4)

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
            if self._mode.get() == "scrum":
                self._run_scrum()
            else:
                self._run_kanban()
        except Exception as ex:
            import traceback
            traceback.print_exc()
            self.after(0, lambda e=ex: self._show_error(str(e)))
        finally:
            self.after(0, lambda: self.run_btn.config(
                state="normal", text="Run Simulation"))

    # ── Scrum mode ─────────────────────────────────────

    def _run_scrum(self):
        path = self.primary_path.get().strip()
        if not path:
            raise ValueError("Please select a Velocity CSV file.")
        if not os.path.exists(path):
            raise ValueError("File not found:\n%s" % path)

        backlog, sprint_d, n_sims, conf_levels, start_date = self._parse_common_inputs()

        self.after(0, lambda: self.status_var.set("Loading velocity data..."))
        vel_samples, vel_col = load_velocity(path)

        if len(vel_samples) < 3:
            raise ValueError(
                "Only %d valid sprint(s) found.\n"
                "Need at least 3 completed sprints." % len(vel_samples))

        ct_samples, ct_col = load_cycle_time(self.cycletime_path.get().strip())

        self.after(0, lambda: self.status_var.set(
            "Running %s simulations..." % "{:,}".format(n_sims)))

        sprint_r     = mc_sprints_to_done(vel_samples, backlog, n_sims)
        median_count = int(np.median(sprint_r))
        through_r    = mc_throughput_in_periods(vel_samples, median_count, n_sims)

        self.after(0, lambda: self._render_scrum(
            sprint_r, through_r, vel_samples,
            ct_samples, ct_col, vel_col,
            backlog, sprint_d, n_sims, conf_levels, median_count, start_date
        ))

    # ── Kanban mode ────────────────────────────────────

    def _run_kanban(self):
        path = self.primary_path.get().strip()
        if not path:
            raise ValueError("Please select an Issue Export CSV file.")
        if not os.path.exists(path):
            raise ValueError("File not found:\n%s" % path)

        backlog, weeks_history, n_sims, conf_levels, start_date = self._parse_common_inputs()

        self.after(0, lambda: self.status_var.set("Calculating weekly throughput..."))
        tp_samples, date_col, n_issues, weekly_counts = load_kanban_throughput(path)

        if len(tp_samples) < 3:
            raise ValueError(
                "Only %d weeks of data found.\n"
                "Need at least 3 weeks of history." % len(tp_samples))

        # Optionally limit to recent N weeks
        if weeks_history > 0 and len(tp_samples) > weeks_history:
            tp_samples_used = tp_samples[-weeks_history:]
        else:
            tp_samples_used = tp_samples

        ct_samples, ct_col = load_cycle_time(self.cycletime_path.get().strip())

        self.after(0, lambda: self.status_var.set(
            "Running %s simulations..." % "{:,}".format(n_sims)))

        week_r       = mc_weeks_to_done(tp_samples_used, backlog, n_sims)
        median_weeks = int(np.median(week_r))
        through_r    = mc_throughput_in_periods(tp_samples_used, median_weeks, n_sims)

        self.after(0, lambda: self._render_kanban(
            week_r, through_r, tp_samples_used, weekly_counts,
            ct_samples, ct_col, date_col,
            backlog, n_issues, n_sims, conf_levels, median_weeks, start_date
        ))

    def _parse_common_inputs(self):
        try:
            backlog  = int(self.backlog_var.get())
            period   = int(self.period_var.get())
            n_sims   = int(self.simulations_var.get())
        except ValueError:
            raise ValueError("Backlog, period, and simulations must be whole numbers.")
        if backlog < 1:
            raise ValueError("Backlog must be at least 1.")
        if n_sims < 100:
            raise ValueError("Please use at least 100 simulations.")
        conf_levels = [l for l, v in self.conf_vars.items() if v.get()]
        if not conf_levels:
            raise ValueError("Select at least one confidence level.")
        start_date = self._parse_start_date()
        return backlog, period, n_sims, conf_levels, start_date

    # ── Render: Scrum ──────────────────────────────────

    def _render_scrum(self, sprint_r, through_r, vel_samples,
                      ct_samples, ct_col, vel_col,
                      backlog, sprint_d, n_sims, conf_levels, median_sprints, start_date):
        self._draw_chart_scrum(
            sprint_r, through_r, backlog, conf_levels, median_sprints, n_sims)
        self._write_summary_scrum(
            sprint_r, through_r, vel_samples,
            ct_samples, ct_col, vel_col,
            backlog, sprint_d, n_sims, conf_levels, median_sprints, start_date)
        self._clear_throughput_tab("Scrum mode — no weekly throughput breakdown.")
        self.status_var.set("Done. %s simulations on %d sprints of history." % (
            "{:,}".format(n_sims), len(vel_samples)))
        self.save_btn.pack(fill="x", pady=(6, 0))

    def _draw_chart_scrum(self, sprint_r, through_r, backlog,
                          conf_levels, median_sprints, n_sims):
        for w in self.chart_frame.winfo_children():
            w.destroy()
        fig, axes = plt.subplots(1, 2, figsize=(9, 4.2))
        self._style_fig(fig, axes, n_sims, backlog, "story points")

        ax1, ax2 = axes
        bins = list(range(int(sprint_r.min()), int(sprint_r.max()) + 2))
        ax1.hist(sprint_r, bins=bins, color=ACCENT, alpha=0.75,
                 edgecolor="#1e1e2e", linewidth=0.4)
        for cl in conf_levels:
            pct = np.percentile(sprint_r, cl)
            ax1.axvline(pct, color=CONF_COLS[cl], linestyle="--", linewidth=1.5,
                        label="%d%% <= %.0f sprints" % (cl, pct))
        ax1.set_xlabel("Sprints to clear backlog", fontsize=9)
        ax1.set_ylabel("Simulations", fontsize=9)
        ax1.set_title("How many sprints?", fontsize=10)
        ax1.legend(fontsize=8, facecolor="#2a2a3e", edgecolor="#444466", labelcolor=TEXT)
        ax1.xaxis.set_major_locator(mticker.MaxNLocator(integer=True))

        ax2.hist(through_r, bins=40, color=ACCENT2, alpha=0.75,
                 edgecolor="#1e1e2e", linewidth=0.4)
        for cl in conf_levels:
            pct = np.percentile(through_r, cl)
            ax2.axvline(pct, color=CONF_COLS[cl], linestyle="--", linewidth=1.5,
                        label="%d%% >= %.0f pts" % (cl, pct))
        ax2.axvline(backlog, color="white", linestyle="-", linewidth=1.5,
                    label="Backlog (%d)" % backlog)
        ax2.set_xlabel("Points completed", fontsize=9)
        ax2.set_ylabel("Simulations", fontsize=9)
        ax2.set_title("Throughput in %d sprints" % median_sprints, fontsize=10)
        ax2.legend(fontsize=8, facecolor="#2a2a3e", edgecolor="#444466", labelcolor=TEXT)

        self._embed_fig(fig)

    def _write_summary_scrum(self, sprint_r, through_r, vel_samples,
                             ct_samples, ct_col, vel_col,
                             backlog, sprint_d, n_sims, conf_levels, median_sprints,
                             start_date):
        st = self.summary_text
        st.config(state="normal")
        st.delete("1.0", "end")

        def line(text="", tag="value"):
            st.insert("end", text + "\n", tag)

        line("=" * 54, "dim")
        line(" SCRUM — MONTE CARLO RESULTS", "heading")
        line("=" * 54, "dim")
        line()
        line("  Mode           : Scrum (velocity-based)", "dim")
        line("  Velocity col   : %s" % vel_col, "dim")
        line("  Sprints history: %d" % len(vel_samples), "dim")
        line("  Velocity range : %.0f – %.0f pts  (mean %.1f)" % (
            vel_samples.min(), vel_samples.max(), vel_samples.mean()), "dim")
        line("  Backlog        : %d story points" % backlog, "dim")
        line("  Sprint length  : %d days" % sprint_d, "dim")
        line("  Simulations    : %s" % "{:,}".format(n_sims), "dim")
        line()
        line("-" * 54, "dim")
        line("  SPRINTS TO CLEAR %d-POINT BACKLOG" % backlog, "subhead")
        line("-" * 54, "dim")
        line("  Mean   : %.1f sprints  (~%.0f days)" % (
            sprint_r.mean(), sprint_r.mean() * sprint_d))
        line("  Median : %.0f sprints  (~%.0f days)" % (
            np.median(sprint_r), np.median(sprint_r) * sprint_d))
        line()
        for cl in conf_levels:
            pct = np.percentile(sprint_r, cl)
            tag = "good" if cl <= 70 else ("warn" if cl <= 85 else "bad")
            line("  %d%% confidence  <=  %.0f sprints  (~%.0f days)" % (
                cl, pct, pct * sprint_d), tag)

        # Date forecast
        if start_date is not None:
            line()
            line("-" * 54, "dim")
            line("  FORECAST FINISH DATES  (from %s)" %
                 start_date.strftime("%d %b %Y"), "subhead")
            line("-" * 54, "dim")
            for cl in conf_levels:
                pct = np.percentile(sprint_r, cl)
                dates = self._finish_dates({cl: pct}, start_date, sprint_d)
                tag = "good" if cl <= 70 else ("warn" if cl <= 85 else "bad")
                line("  %d%% confidence by  :  %s" % (cl, dates[cl]), tag)
        line()
        line("-" * 54, "dim")
        line("  THROUGHPUT IN %d SPRINTS (MEDIAN)" % median_sprints, "subhead")
        line("-" * 54, "dim")
        line("  Mean   : %.0f pts" % through_r.mean())
        line("  Median : %.0f pts" % np.median(through_r))
        line()
        for cl in conf_levels:
            pct = np.percentile(through_r, cl)
            line("  %d%%ile  >=  %.0f pts  (%.0f%% of backlog)" % (
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

    # ── Render: Kanban ─────────────────────────────────

    def _render_kanban(self, week_r, through_r, tp_samples, weekly_counts,
                       ct_samples, ct_col, date_col,
                       backlog, n_issues, n_sims, conf_levels, median_weeks, start_date):
        self._draw_chart_kanban(
            week_r, through_r, tp_samples, backlog,
            conf_levels, median_weeks, n_sims)
        self._write_summary_kanban(
            week_r, through_r, tp_samples,
            ct_samples, ct_col, date_col,
            backlog, n_issues, n_sims, conf_levels, median_weeks, start_date)
        self._write_throughput_tab(weekly_counts, tp_samples)
        self.status_var.set("Done. %s simulations on %d weeks of history (%d issues)." % (
            "{:,}".format(n_sims), len(tp_samples), n_issues))
        self.save_btn.pack(fill="x", pady=(6, 0))

    def _draw_chart_kanban(self, week_r, through_r, tp_samples, backlog,
                           conf_levels, median_weeks, n_sims):
        for w in self.chart_frame.winfo_children():
            w.destroy()
        fig, axes = plt.subplots(1, 2, figsize=(9, 4.2))
        self._style_fig(fig, axes, n_sims, backlog, "items")

        ax1, ax2 = axes

        # Left: weeks to completion
        bins = list(range(int(week_r.min()), int(week_r.max()) + 2))
        ax1.hist(week_r, bins=bins, color=ACCENT, alpha=0.75,
                 edgecolor="#1e1e2e", linewidth=0.4)
        for cl in conf_levels:
            pct = np.percentile(week_r, cl)
            ax1.axvline(pct, color=CONF_COLS[cl], linestyle="--", linewidth=1.5,
                        label="%d%% <= %.0f weeks" % (cl, pct))
        ax1.set_xlabel("Weeks to clear backlog", fontsize=9)
        ax1.set_ylabel("Simulations", fontsize=9)
        ax1.set_title("How many weeks?", fontsize=10)
        ax1.legend(fontsize=8, facecolor="#2a2a3e", edgecolor="#444466", labelcolor=TEXT)
        ax1.xaxis.set_major_locator(mticker.MaxNLocator(integer=True))

        # Right: throughput histogram
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

        self._embed_fig(fig)

    def _write_summary_kanban(self, week_r, through_r, tp_samples,
                              ct_samples, ct_col, date_col,
                              backlog, n_issues, n_sims, conf_levels, median_weeks,
                              start_date):
        st = self.summary_text
        st.config(state="normal")
        st.delete("1.0", "end")

        def line(text="", tag="value"):
            st.insert("end", text + "\n", tag)

        line("=" * 54, "dim")
        line(" KANBAN — MONTE CARLO RESULTS", "heading")
        line("=" * 54, "dim")
        line()
        line("  Mode              : Kanban (throughput-based)", "dim")
        line("  Resolved col      : %s" % date_col, "dim")
        line("  Issues analysed   : %d" % n_issues, "dim")
        line("  Weeks of history  : %d" % len(tp_samples), "dim")
        line("  Throughput range  : %.0f – %.0f items/week  (mean %.1f)" % (
            tp_samples.min(), tp_samples.max(), tp_samples.mean()), "dim")
        line("  Backlog           : %d items" % backlog, "dim")
        line("  Simulations       : %s" % "{:,}".format(n_sims), "dim")
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

        # Date forecast
        if start_date is not None:
            line()
            line("-" * 54, "dim")
            line("  FORECAST FINISH DATES  (from %s)" %
                 start_date.strftime("%d %b %Y"), "subhead")
            line("-" * 54, "dim")
            for cl in conf_levels:
                pct = np.percentile(week_r, cl)
                dates = self._finish_dates({cl: pct}, start_date, 7)
                tag = "good" if cl <= 70 else ("warn" if cl <= 85 else "bad")
                line("  %d%% confidence by  :  %s" % (cl, dates[cl]), tag)
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

    def _write_throughput_tab(self, weekly_counts, tp_samples):
        tt = self.throughput_text
        tt.config(state="normal")
        tt.delete("1.0", "end")
        tt.insert("end", "  WEEKLY THROUGHPUT BREAKDOWN\n", )
        tt.insert("end", "  (items completed per week)\n\n")
        tt.insert("end", "  %-20s  %s\n" % ("Week", "Items"))
        tt.insert("end", "  " + "-" * 30 + "\n")
        for period, count in weekly_counts.items():
            bar = "#" * int(count)
            tt.insert("end", "  %-20s  %3d  %s\n" % (str(period), count, bar))
        tt.insert("end", "\n  " + "-" * 30 + "\n")
        tt.insert("end", "  Mean   : %.1f items/week\n" % tp_samples.mean())
        tt.insert("end", "  Median : %.1f items/week\n" % np.median(tp_samples))
        tt.insert("end", "  Min    : %.0f  |  Max : %.0f\n" % (
            tp_samples.min(), tp_samples.max()))
        tt.config(state="disabled")

    def _clear_throughput_tab(self, msg):
        tt = self.throughput_text
        tt.config(state="normal")
        tt.delete("1.0", "end")
        tt.insert("end", "  " + msg + "\n")
        tt.config(state="disabled")

    # ── Shared chart helpers ───────────────────────────

    def _style_fig(self, fig, axes, n_sims, backlog, unit):
        fig.patch.set_facecolor("#1e1e2e")
        plt.subplots_adjust(wspace=0.35, left=0.08, right=0.97, top=0.88, bottom=0.12)
        fig.suptitle(
            "Monte Carlo Projection  —  %s simulations  |  Backlog: %d %s" % (
                "{:,}".format(n_sims), backlog, unit),
            color=TEXT, fontsize=11, fontweight="bold"
        )
        for ax in axes:
            ax.set_facecolor("#2a2a3e")
            ax.tick_params(colors=TEXT_DIM, labelsize=8)
            for spine in ax.spines.values():
                spine.set_color("#444466")
            ax.grid(axis="y", color="#333355", linewidth=0.5)
            ax.yaxis.label.set_color(TEXT_DIM)
            ax.xaxis.label.set_color(TEXT_DIM)
            ax.title.set_color(TEXT)

    def _embed_fig(self, fig):
        canvas = FigureCanvasTkAgg(fig, master=self.chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        self._fig = fig

    # ── Save / Error ───────────────────────────────────

    def _save_chart(self):
        if self._fig is None:
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG image", "*.png")],
            initialfile="monte_carlo_results.png",
            title="Save chart"
        )
        if path:
            self._fig.savefig(path, dpi=150, bbox_inches="tight",
                              facecolor=self._fig.get_facecolor())
            messagebox.showinfo("Saved", "Chart saved to:\n%s" % path)

    def _show_error(self, msg):
        self.status_var.set("Error — see message.")
        messagebox.showerror("Simulation Error", msg)


# ══════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════

if __name__ == "__main__":
    print("Starting Monte Carlo Projection Tool...")
    try:
        app = MonteCarloApp()
        print("Entering main loop...")
        app.mainloop()
    except Exception as e:
        import traceback
        traceback.print_exc()
        input("Press Enter to exit...")
