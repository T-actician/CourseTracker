"""
Course Tracker - your 90-day course, ticked off and noted down.

Import the whole course document (.docx, .txt or .md, or paste it), and the app
splits it into days, turns every checklist line into a real tick-box, and gives
each day five notepads (Learn, Code, Build, Test, Document) with code colouring.

Run:   python course_tracker.py
Build: double-click build.bat  (makes dist\\CourseTracker.exe)
"""
import datetime as dt
import json
import os
import re
import shutil
import tkinter as tk
import tkinter.font as tkfont
from functools import partial
from tkinter import filedialog, messagebox

import customtkinter as ctk

APP = "Course Tracker"

# key, tab label, minutes
BLOCKS = [
    ("learn", "Learn", 45),
    ("code", "Code", 60),
    ("build", "Build", 90),
    ("test", "Test", 30),
    ("doc", "Document", 15),
]
MARKERS = {
    "Learn": "learn",
    "Code without AI": "code",
    "Build": "build",
    "Test and debug": "test",
    "Document": "doc",
}
SCORE_FIELDS = [
    ("Explain concepts without notes", 20),
    ("Solve coding tasks without AI", 25),
    ("Build the week's deliverable", 25),
    ("Debug and test correctly", 20),
    ("Document work and commit to Git", 10),
]

PALETTES = {
    "Gold": dict(
        BG="#0e0b07", PANEL="#15110a", FIELD="#1d1810", FIELD2="#2a2110", BORDER="#4a3a15",
        ACC="#E3B23C", ACC_HI="#F5CB5C", ACC_DIM="#7a5d1e", TEXT="#f2e9d0", MUTED="#a89c7a",
        OK="#2FA866", EDBG="#0b0905", SEL="#3a2e14", ON_ACC="#000000",
        kw="#E3B23C", st="#9ccf7a", com="#7d7358", num="#e58b5a", fn="#7fc8d6"),
    "Matrix Green": dict(
        BG="#080b08", PANEL="#0d130d", FIELD="#121a12", FIELD2="#1b2a1b", BORDER="#24472a",
        ACC="#3ddc84", ACC_HI="#7dffb0", ACC_DIM="#1d6b3c", TEXT="#d7f2dc", MUTED="#7fa089",
        OK="#3ddc84", EDBG="#060906", SEL="#1d3a24", ON_ACC="#000000",
        kw="#61d6d6", st="#f9f1a5", com="#6b7f6f", num="#e0a85a", fn="#7db4ff"),
}
KEYWORDS = set("""
def class return if elif else for while in not and or is import from as with try except finally
raise pass break continue lambda None True False self yield async await global nonlocal assert del
function const let var new this null undefined typeof export default switch case throw catch
SELECT FROM WHERE INSERT INTO VALUES UPDATE SET DELETE JOIN LEFT INNER ON GROUP BY ORDER HAVING
LIMIT CREATE TABLE PRIMARY KEY FOREIGN REFERENCES NOT NULL UNIQUE CHECK INDEX DROP ALTER AS AND OR
select from where insert into values update set delete join left inner on group by order having
limit create table primary key foreign references unique check index drop alter
""".split())

UI_FONT = "Palatino Linotype"
PAIRS = {"(": ")", "[": "]", "{": "}", '"': '"', "'": "'", "`": "`"}
CLOSERS = set(")]}\"'`")


def set_palette(name):
    globals().update(PALETTES.get(name, PALETTES["Gold"]))


set_palette("Gold")


def read_course_file(path):
    """Return the text of a .docx / .txt / .md file."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        try:
            from docx import Document
        except ImportError:
            raise RuntimeError("python-docx is not installed. Run:  pip install python-docx")
        doc = Document(path)
        return "\n".join(p.text for p in doc.paragraphs)
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        return f.read()


# ----------------------------------------------------------------------------
# Storage
# ----------------------------------------------------------------------------
def data_dir():
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    path = os.path.join(base, "CourseTracker")
    os.makedirs(path, exist_ok=True)
    return path


DATA_FILE = os.path.join(data_dir(), "data.json")


def default_state():
    return {"v": 1, "days": {}, "activity": [], "theme": "vscode",
            "font": 11, "last_day": 1}


def load_state():
    if not os.path.exists(DATA_FILE):
        return default_state()
    try:
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            st = json.load(f)
        base = default_state()
        base.update(st)
        return base
    except Exception:
        try:
            shutil.copy(DATA_FILE, DATA_FILE + ".corrupt.bak")
        except Exception:
            pass
        return default_state()


def save_state(state):
    tmp = DATA_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=1)
    os.replace(tmp, DATA_FILE)


# ----------------------------------------------------------------------------
# Course parser
# ----------------------------------------------------------------------------
DAY_RE = re.compile(r"^\s*(?:#{1,6}\s*)?Day\s+(\d{1,2})\s*[:\-\u2013\u2014.]\s*(\S.*)$", re.I)
STOP_RE = re.compile(
    r"^\s*(?:#{1,6}\s*)?(WEEK\s+\d+\s*[:\-\u2013\u2014]|PHASE\s+\d+\s*[:\-\u2013\u2014]"
    r"|FINAL SCORECARD|-{3,}\s*$|\|)", re.I)
CHECK_HEAD_RE = re.compile(r"^\s*(final\s+)?check\s?list\s*:?\s*$", re.I)
BLOCK_RE = re.compile(
    r"(Learn|Code without AI|Build|Test and debug|Document)\s*\(\s*\d+\s*min\s*\)\s*:")
SKIP_ITEM_RE = re.compile(r"^(Day\s+\d+\s+note|Week\s+\d+\s+assessment|Report template)", re.I)
ITEM_RE = re.compile(r"^[-*\u2022]?\s*(\[[ xX]?\]|\u2610|\u2611|\u2612|\u25a1|\u25a0)?\s*(.*)$")
DONE_MARKS = {"[x]", "[X]", "\u2611", "\u2612"}


def clean_line(raw):
    s = raw.replace("\t", " ").replace("`", "")
    s = re.sub(r"\*\*(.+?)\*\*", r"\1", s)
    s = re.sub(r"^\s*>\s?", "", s)
    s = re.sub(r"^\s*#{1,6}\s+", "", s)
    s = s.strip()
    m = re.match(r"^\*(.+)\*$", s)
    if m:
        s = m.group(1).strip()
    return s


def parse_course(text):
    """Return {day_number: {title, lesson{key:text}, checks[{t,d}]}}."""
    raw_days = {}
    cur = None
    stopped = False
    for raw in text.replace("\r", "").split("\n"):
        line = clean_line(raw)
        m = DAY_RE.match(line)
        if m:
            n = int(m.group(1))
            if 1 <= n <= 90:
                cur = raw_days[n] = {"title": m.group(2).strip(), "lines": []}
                stopped = False
                continue
        if cur is None:
            continue
        if STOP_RE.match(line):
            stopped = True
        if not stopped:
            cur["lines"].append(line)

    days = {}
    for n, d in raw_days.items():
        lines = d["lines"]
        idx = None
        for i, l in enumerate(lines):
            if CHECK_HEAD_RE.match(l):
                idx = i
                break
        body_lines = lines if idx is None else lines[:idx]
        check_lines = [] if idx is None else lines[idx + 1:]
        body = "\n".join(body_lines)

        lesson = {k: "" for k, _, _ in BLOCKS}
        ms = list(BLOCK_RE.finditer(body))
        for i, m in enumerate(ms):
            end = ms[i + 1].start() if i + 1 < len(ms) else len(body)
            seg = re.sub(r"\s+", " ", body[m.end():end]).strip()
            seg = re.sub(r"\s+(Week\s+\d+\s+assessment)", r"\n\n\1", seg)
            lesson[MARKERS[m.group(1)]] = seg

        checks = []
        for l in check_lines:
            s = l.strip()
            if not s or SKIP_ITEM_RE.match(s):
                continue
            mm = ITEM_RE.match(s)
            mark = mm.group(1) or ""
            txt = mm.group(2).strip()
            if txt:
                checks.append({"t": txt, "d": mark in DONE_MARKS})
        days[n] = {"title": d["title"], "lesson": lesson, "checks": checks}
    return days


def merge_days(state, new_days):
    """Add parsed days, keeping any notes/progress already stored."""
    for n, nd in new_days.items():
        key = str(n)
        old = state["days"].get(key)
        notes = {k: "" for k, _, _ in BLOCKS}
        bdone = {k: False for k, _, _ in BLOCKS}
        score = {}
        if old:
            notes.update(old.get("notes", {}))
            bdone.update(old.get("bdone", {}))
            score = old.get("score", {})
            done_map = {c["t"]: c["d"] for c in old.get("checks", [])}
            for c in nd["checks"]:
                if done_map.get(c["t"]):
                    c["d"] = True
        state["days"][key] = {"title": nd["title"], "lesson": nd["lesson"],
                              "checks": nd["checks"], "notes": notes,
                              "bdone": bdone, "score": score}


# ----------------------------------------------------------------------------
# Widget helpers. Rounded CTk widgets are used sparingly (they are slow to
# redraw); lists, chips and text use light native widgets in the same colours.
# ----------------------------------------------------------------------------
def gold_button(parent, text, command, **kw):
    return ctk.CTkButton(parent, text=text, command=command, fg_color=ACC, hover_color=ACC_HI,
                         text_color=ON_ACC, corner_radius=10, **kw)


def dark_button(parent, text, command, **kw):
    return ctk.CTkButton(parent, text=text, command=command, fg_color=FIELD, hover_color=FIELD2,
                         text_color=TEXT, border_width=1, border_color=BORDER, corner_radius=10, **kw)


def card(parent, **kw):
    return ctk.CTkFrame(parent, fg_color=PANEL, border_width=1, border_color=BORDER,
                        corner_radius=14, **kw)


def tl(parent, text="", size=12, color=None, bold=False, bg=None, **kw):
    """Light native label."""
    return tk.Label(parent, text=text, fg=color or TEXT, bg=bg or BG, anchor="w", bd=0,
                    font=(UI_FONT, size, "bold" if bold else "normal"), **kw)


def track(parent, bg, height=6):
    """Native progress bar: returns (frame, setter)."""
    fr = tk.Frame(parent, bg="#2a2417", height=height)
    fill = tk.Frame(fr, bg=ACC)
    fill.place(x=0, y=0, relheight=1, relwidth=0.0)

    def setv(v):
        fill.place_configure(relwidth=max(0.0, min(1.0, v)))
    return fr, setv


def chip(parent, text, bg, fg, cmd, width=5, bgp=None):
    b = tk.Label(parent, text=text, bg=bg, fg=fg, width=width, pady=5, cursor="hand2",
                 font=(UI_FONT, 10, "bold"), bd=0)
    b.bind("<Button-1>", lambda e: cmd())
    return b


class FastScroll(tk.Frame):
    """Smooth scroll area: one canvas, one inner frame, one scrollbar."""

    def __init__(self, parent, bg):
        super().__init__(parent, bg=bg)
        self.cv = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0, yscrollincrement=28)
        self.sb = ctk.CTkScrollbar(self, command=self.cv.yview, button_color=BORDER,
                                   button_hover_color=ACC_DIM, width=12)
        self.cv.configure(yscrollcommand=self.sb.set)
        self.sb.pack(side="right", fill="y")
        self.cv.pack(side="left", fill="both", expand=True)
        self.inner = tk.Frame(self.cv, bg=bg)
        self.win = self.cv.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", self._inner)
        self.cv.bind("<Configure>", lambda e: self.cv.itemconfigure(self.win, width=e.width))

    def _inner(self, _):
        self.cv.configure(scrollregion=self.cv.bbox("all"))

    def wheel(self, e):
        if self.inner.winfo_reqheight() > self.cv.winfo_height():
            self.cv.yview_scroll(int(-e.delta / 120) * 2, "units")

    def clear(self):
        for w in self.inner.winfo_children():
            w.destroy()
        self.cv.yview_moveto(0)


NAV = [("Dashboard", "\u2302"), ("Today", "\u25b6"), ("Search", "\u2315"),
       ("Timer", "\u25f7"), ("Import", "\u2193"), ("Settings", "\u2699")]


# ----------------------------------------------------------------------------
# Main app
# ----------------------------------------------------------------------------
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.state_ = load_state()
        pal = self.state_.get("palette", "Gold")
        set_palette(pal if pal in PALETTES else "Gold")
        try:
            ctk.ThemeManager.theme["CTkFont"]["family"] = UI_FONT
            ctk.ThemeManager.theme["CTkFont"]["size"] = 14
        except Exception:
            pass
        ctk.set_appearance_mode("dark")
        self.title(APP)
        self.configure(fg_color=BG)
        self.center_window(1240, 820)

        fams = set(tkfont.families())
        self.mono = next((f for f in ("Cascadia Mono", "Consolas", "Courier New") if f in fams), "Courier New")
        self.ef = tkfont.Font(family=self.mono, size=int(self.state_.get("font", 12)))
        self.cur = None
        self.cur_block = "learn"
        self.loading = False
        self.save_job = None
        self.font_job = None
        self.search_job = None
        self.hl_jobs = {}
        self.page = None
        self.editors = {}
        self.loaded = set()
        self.shown = None
        self.dash_dirty = True
        self.panel_week = None
        self.panel_rows = {}
        self.week_w = {}
        self.tab_text = {}
        self.score_vars = []
        self.timer = dict(step=0, left=BLOCKS[0][2] * 60, running=False, job=None)

        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.bind("<Control-s>", lambda e: self.flush())
        self.bind("<Control-f>", lambda e: self.show_page("Search"))
        for seq in ("<Control-plus>", "<Control-equal>"):
            self.bind(seq, lambda e: self.zoom(1))
        self.bind("<Control-minus>", lambda e: self.zoom(-1))
        self.bind_all("<MouseWheel>", self.on_wheel)

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.build_sidebar()
        self.main = tk.Frame(self, bg=BG)
        self.main.grid(row=0, column=1, sticky="nsew", padx=20, pady=14)
        self.main.grid_columnconfigure(0, weight=1)
        self.main.grid_rowconfigure(0, weight=1)
        self.build_week_panel()
        self.pages = {}
        self.build_dashboard()
        self.build_day()
        self.build_search()
        self.build_timer()
        self.build_import()
        self.build_settings()
        for p in self.pages.values():
            p.grid(row=0, column=0, sticky="nsew")

        self.build_week_cards()
        self.refresh_overall()
        if self.state_["days"]:
            last = self.state_.get("last_day")
            self.open_day(last if str(last) in self.state_["days"] else self.today_target(), go=False)
        self.show_page("Dashboard" if self.state_["days"] else "Import")

    # ---------- window ----------
    def center_window(self, w, h):
        try:
            sc = ctk.ScalingTracker.get_window_scaling(self)
        except Exception:
            sc = self.winfo_fpixels("1i") / 96
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        pw = min(int(w * sc), sw - int(60 * sc))
        ph = min(int(h * sc), sh - int(130 * sc))
        self.geometry(f"{int(pw / sc)}x{int(ph / sc)}")
        self.minsize(min(1000, int(pw / sc)), min(620, int(ph / sc)))
        x, y = (sw - pw) // 2, max(0, (sh - ph) // 2 - int(20 * sc))
        tk.Tk.geometry(self, f"{pw}x{ph}+{x}+{y}")

    def on_wheel(self, e):
        w = self.winfo_containing(e.x_root, e.y_root)
        while w is not None:
            if isinstance(w, FastScroll):
                w.wheel(e)
                return
            w = w.master

    # ---------- sidebar ----------
    def build_sidebar(self):
        sb = ctk.CTkFrame(self, width=210, fg_color=PANEL, corner_radius=0)
        sb.grid(row=0, column=0, sticky="ns")
        sb.pack_propagate(False)
        ctk.CTkLabel(sb, text="\u25c8 Course Tracker", text_color=ACC, anchor="w",
                     font=ctk.CTkFont(size=20, weight="bold")).pack(fill="x", padx=18, pady=(24, 22))
        self.nav = {}
        self.nav_on = None
        for name, icon in NAV:
            b = ctk.CTkButton(sb, text=f"  {icon}   {name}", anchor="w", height=44, corner_radius=12,
                              fg_color="transparent", hover_color=FIELD2, text_color=TEXT,
                              font=ctk.CTkFont(size=14), command=partial(self.show_page, name))
            b.pack(fill="x", padx=12, pady=3)
            self.nav[name] = b
        foot = tk.Frame(sb, bg=PANEL)
        foot.pack(side="bottom", fill="x", padx=16, pady=18)
        self.streak_lbl = tl(foot, "", 11, ACC_HI, True, PANEL)
        self.streak_lbl.pack(fill="x", pady=(0, 6))
        self.side_pct = tl(foot, "", 10, MUTED, False, PANEL)
        self.side_pct.pack(fill="x")
        fr, self.side_set = track(foot, PANEL, 8)
        fr.pack(fill="x", pady=(4, 0))

    def style_nav(self, sel):
        if self.nav_on == sel:
            return
        if self.nav_on:
            self.nav[self.nav_on].configure(fg_color="transparent", text_color=TEXT)
        self.nav[sel].configure(fg_color=FIELD2, text_color=ACC_HI)
        self.nav_on = sel

    def show_page(self, name):
        if name == "Today":
            if not self.state_["days"]:
                name = "Import"
            else:
                self.open_day(self.today_target(), go=False)
                name = "Day"
        if self.big_open:
            self.close_big()
        if name == self.page:
            return
        self.pages[name].tkraise()
        self.page = name
        self.style_nav("Today" if name == "Day" else name)
        if name == "Dashboard" and self.dash_dirty:
            self.refresh_dashboard()
        if name == "Day":
            self.update_day_widgets()

    # ---------- right panel (this week) ----------
    def build_week_panel(self):
        q = tk.Frame(self, width=230, bg=PANEL)
        q.grid(row=0, column=2, sticky="ns")
        q.grid_propagate(False)
        q.pack_propagate(False)
        self.week_title = tl(q, "This Week", 15, ACC, True, PANEL)
        self.week_title.pack(fill="x", padx=18, pady=(24, 12))
        self.week_rows = tk.Frame(q, bg=PANEL)
        self.week_rows.pack(fill="x")

    def status_color(self, d):
        p = self.day_progress(d)
        return OK if p >= 1 else (ACC_DIM if p > 0 else FIELD2)

    def refresh_week_panel(self):
        n = self.cur
        if n is None:
            return
        wk = min((n - 1) // 7 + 1, 13)
        if wk != self.panel_week:
            for w in self.week_rows.winfo_children():
                w.destroy()
            self.panel_rows = {}
            self.panel_week = wk
            self.week_title.configure(text=f"Week {wk}")
            for d in range((wk - 1) * 7 + 1, min(wk * 7, 90) + 1):
                if str(d) not in self.state_["days"]:
                    continue
                f = tk.Frame(self.week_rows, bg="#17130b", highlightthickness=1,
                             highlightbackground=BORDER, height=44, cursor="hand2")
                f.pack(fill="x", padx=14, pady=4)
                f.pack_propagate(False)
                ic = tk.Label(f, text=str(d), width=4, bg=FIELD2, fg="white", bd=0,
                              font=(UI_FONT, 10, "bold"))
                ic.pack(side="left", padx=(10, 8), pady=9)
                t = self.state_["days"][str(d)]["title"]
                tx = tk.Label(f, text=(t if len(t) < 19 else t[:17] + "\u2026"), fg=TEXT, bg="#17130b",
                              anchor="w", bd=0, font=(UI_FONT, 10))
                tx.pack(side="left", fill="x", expand=True)
                for wdg in (f, ic, tx):
                    wdg.bind("<Button-1>", lambda e, k=d: self.open_day(k))
                self.panel_rows[d] = (f, ic, tx)
        for d, (f, ic, tx) in self.panel_rows.items():
            sel = d == n
            bg = FIELD2 if sel else "#17130b"
            f.configure(bg=bg, highlightbackground=ACC if sel else BORDER)
            tx.configure(bg=bg)
            ic.configure(bg=self.status_color(d))

    # ---------- dashboard ----------
    def build_dashboard(self):
        f = self.pages["Dashboard"] = tk.Frame(self.main, bg=BG)
        tl(f, "Dashboard", 26, TEXT, True).pack(fill="x")
        self.dash_sub = tl(f, "", 11, MUTED)
        self.dash_sub.pack(fill="x", pady=(0, 12))
        row = tk.Frame(f, bg=BG)
        row.pack(fill="x", pady=(0, 12))
        row.grid_columnconfigure((0, 1, 2), weight=1, uniform="s")
        self.stat = []
        for i, name in enumerate(("Overall progress", "Days completed", "Study streak")):
            c = card(row)
            c.grid(row=0, column=i, sticky="ew", padx=(0 if i == 0 else 8, 0))
            tl(c, name, 10, MUTED, False, PANEL).pack(fill="x", padx=16, pady=(12, 0))
            v = tl(c, "-", 22, ACC_HI, True, PANEL)
            v.pack(fill="x", padx=16, pady=(0, 12))
            self.stat.append(v)
        self.cont = card(f)
        self.cont.pack(fill="x", pady=(0, 12))
        self.cont_title = tl(self.cont, "", 14, TEXT, True, PANEL)
        self.cont_title.pack(fill="x", padx=18, pady=(14, 0))
        self.cont_sub = tl(self.cont, "", 10, MUTED, False, PANEL)
        self.cont_sub.pack(fill="x", padx=18)
        fr, self.cont_set = track(self.cont, PANEL, 8)
        fr.pack(fill="x", padx=18, pady=(8, 8))
        gold_button(self.cont, "Continue  \u25b6", lambda: self.show_page("Today"), height=36, width=160
                    ).pack(anchor="w", padx=18, pady=(0, 14))
        tl(f, "All weeks", 13, ACC, True).pack(fill="x", pady=(2, 6))
        self.weeks = FastScroll(f, BG)
        self.weeks.pack(fill="both", expand=True)

    def build_week_cards(self):
        """Built once (and after an import); afterwards only recoloured."""
        self.weeks.clear()
        self.week_w = {}
        days = self.state_["days"]
        for wk in range(1, 14):
            ds = [d for d in range((wk - 1) * 7 + 1, min(wk * 7, 90) + 1) if str(d) in days]
            if not ds:
                continue
            c = tk.Frame(self.weeks.inner, bg=PANEL, highlightthickness=1, highlightbackground=BORDER)
            c.pack(fill="x", pady=4, padx=(0, 6))
            top = tk.Frame(c, bg=PANEL)
            top.pack(fill="x", padx=16, pady=(10, 4))
            tl(top, f"Week {wk}", 12, TEXT, True, PANEL).pack(side="left")
            pct = tl(top, "", 11, ACC_HI, True, PANEL)
            pct.pack(side="right")
            tr, setter = track(c, PANEL, 6)
            tr.pack(fill="x", padx=16)
            chips = tk.Frame(c, bg=PANEL)
            chips.pack(fill="x", padx=12, pady=(8, 12))
            cd = {}
            for d in ds:
                lb = chip(chips, str(d), FIELD2, TEXT, partial(self.open_day, d), width=5)
                lb.pack(side="left", padx=3)
                cd[d] = lb
            self.week_w[wk] = (pct, setter, cd)
        self.dash_dirty = True

    def refresh_dashboard(self):
        self.dash_dirty = False
        days = self.state_["days"]
        nums = sorted(int(k) for k in days)
        total = len(nums)
        done = sum(1 for n in nums if self.day_progress(n) >= 1)
        avg = (sum(self.day_progress(n) for n in nums) / total) if total else 0
        self.dash_sub.configure(text=f"{total} days loaded  \u00b7  {done} complete" if total
                                else "Nothing loaded yet. Open Import and add your course.")
        self.stat[0].configure(text=f"{avg * 100:.0f}%")
        self.stat[1].configure(text=f"{done} / {total or 90}")
        s = self.streak()
        self.stat[2].configure(text=f"{s} day{'s' if s != 1 else ''}")
        if total:
            n = self.today_target()
            self.cont_title.configure(text=f"Next up: Day {n} \u00b7 {days[str(n)]['title']}")
            self.cont_sub.configure(text=f"{self.day_progress(n) * 100:.0f}% of this day complete")
            self.cont_set(self.day_progress(n))
        else:
            self.cont_title.configure(text="Import your course to begin")
            self.cont_sub.configure(text="")
            self.cont_set(0)
        for wk, (pct, setter, cd) in self.week_w.items():
            ps = []
            for d, lb in cd.items():
                p = self.day_progress(d)
                ps.append(p)
                lb.configure(bg=OK if p >= 1 else (ACC_DIM if p > 0 else FIELD2),
                             fg="white" if p > 0 else TEXT)
            avgw = sum(ps) / len(ps)
            pct.configure(text=f"{avgw * 100:.0f}%")
            setter(avgw)

    # ---------- day workspace ----------
    def build_day(self):
        f = self.pages["Day"] = tk.Frame(self.main, bg=BG)
        self.d_title = tl(f, "", 22, TEXT, True)
        self.d_title.pack(fill="x")
        head = tk.Frame(f, bg=BG)
        head.pack(fill="x", pady=(0, 8))
        self.d_sub = tl(head, "", 11, MUTED)
        self.d_sub.pack(side="left")
        dark_button(head, "Next \u25b6", lambda: self.step_day(1), width=80, height=28).pack(side="right")
        dark_button(head, "\u25c0 Prev", lambda: self.step_day(-1), width=80, height=28).pack(side="right", padx=6)

        # checklist (one native Text: no per-item widgets)
        self.chk_card = card(f)
        self.chk_card.pack(fill="x", pady=(0, 8))
        ch = tk.Frame(self.chk_card, bg=PANEL)
        ch.pack(fill="x", padx=16, pady=(8, 2))
        tl(ch, "Checklist  (click a line to tick it)", 12, ACC, True, PANEL).pack(side="left")
        self.chk_count = tl(ch, "", 10, MUTED, False, PANEL)
        self.chk_count.pack(side="right")
        cw = tk.Frame(self.chk_card, bg=PANEL)
        cw.pack(fill="x", padx=10, pady=(0, 8))
        self.chk = tk.Text(cw, bg=PANEL, fg=TEXT, relief="flat", wrap="word", height=5, bd=0,
                           highlightthickness=0, font=(UI_FONT, 11), cursor="arrow", padx=6, pady=2,
                           spacing1=3, spacing3=3, state="disabled", selectbackground=PANEL,
                           insertwidth=0)
        csb = ctk.CTkScrollbar(cw, command=self.chk.yview, button_color=BORDER,
                               button_hover_color=ACC_DIM, width=12)
        self.chk.configure(yscrollcommand=csb.set)
        csb.pack(side="right", fill="y")
        self.chk.pack(fill="x")
        self.chk.tag_configure("done", foreground=MUTED, overstrike=True)
        self.score_card = tk.Frame(self.chk_card, bg=FIELD)

        # block tabs
        tabs = tk.Frame(f, bg=BG)
        tabs.pack(fill="x", pady=(0, 6))
        self.tab_btns = {}
        for key, lab, mins in BLOCKS:
            b = ctk.CTkButton(tabs, text=f"{lab}  {mins}m", height=34, corner_radius=10, fg_color=FIELD,
                              hover_color=FIELD2, text_color=TEXT, border_width=1, border_color=BORDER,
                              command=partial(self.select_block, key))
            b.pack(side="left", padx=(0, 6))
            self.tab_btns[key] = b
            self.tab_text[key] = b.cget("text")
        self.tab_on = None

        self.less_card = card(f)
        self.less_card.pack(fill="x", pady=(0, 8))
        tl(self.less_card, "WHAT TO DO", 8, MUTED, True, PANEL).pack(fill="x", padx=16, pady=(8, 0))
        self.lesson = tk.Text(self.less_card, height=4, wrap="word", bg=PANEL, fg=TEXT, relief="flat",
                              font=(UI_FONT, 11), padx=10, pady=4, bd=0, highlightthickness=0,
                              selectbackground=SEL, state="disabled", insertwidth=0)
        self.lesson.pack(fill="x", padx=8, pady=(0, 8))

        self.ed_card = card(f)
        self.ed_card.pack(fill="both", expand=True)
        eh = tk.Frame(self.ed_card, bg=PANEL)
        eh.pack(fill="x", padx=16, pady=(6, 0))
        tl(eh, "MY NOTES / CODE", 8, MUTED, True, PANEL).pack(side="left")
        self.blk_var = tk.BooleanVar(value=False)
        self.blk_check = ctk.CTkCheckBox(eh, text="Block complete", variable=self.blk_var,
                                         command=self.toggle_block, fg_color=ACC, hover_color=ACC_HI,
                                         border_color=ACC_DIM, checkmark_color="#000000", text_color=TEXT,
                                         corner_radius=5, border_width=2, checkbox_width=20, checkbox_height=20)
        self.blk_check.pack(side="right")
        dark_button(eh, "Copy", self.copy_notes, width=64, height=26).pack(side="right", padx=10)
        dark_button(eh, "\u2922 Expand", self.open_big, width=84, height=26).pack(side="right")
        host = tk.Frame(self.ed_card, bg=PANEL)
        host.pack(fill="both", expand=True, padx=8, pady=(2, 8))
        for key, _, _ in BLOCKS:
            fr = tk.Frame(host, bg=EDBG)
            ed = tk.Text(fr, bg=EDBG, fg=TEXT, insertbackground=ACC, relief="flat", font=self.ef,
                         undo=True, wrap="none", padx=12, pady=10, selectbackground=SEL,
                         highlightthickness=0, bd=0)
            ys = ctk.CTkScrollbar(fr, command=ed.yview, button_color=BORDER, button_hover_color=ACC_DIM, width=12)
            xs = ctk.CTkScrollbar(fr, orientation="horizontal", command=ed.xview, button_color=BORDER,
                                  button_hover_color=ACC_DIM, height=12)
            ed.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
            ys.pack(side="right", fill="y")
            xs.pack(side="bottom", fill="x")
            ed.pack(fill="both", expand=True)
            self.setup_editor(ed, lambda k=key: k)
            self.editors[key] = (fr, ed)
        self.build_big_editor()

    def setup_editor(self, ed, keyfn):
        """Colours + key behaviour shared by the normal and the expanded editor."""
        for tag, col in (("kw", kw), ("num", num), ("fn", fn), ("str", st), ("com", com)):
            ed.tag_configure(tag, foreground=col)
        ed.bind("<<Modified>>", lambda e: self.on_edit(e, keyfn()))
        ed.bind("<Tab>", self.on_tab)
        ed.bind("<Return>", self.on_return)
        ed.bind("<KeyPress>", self.on_key)
        ed.bind("<BackSpace>", self.on_backspace)
        ed.bind("<Control-MouseWheel>", lambda e: (self.zoom(1 if e.delta > 0 else -1), "break")[1])

    # ---------- expanded (floating) notes card ----------
    BIG_MARGIN = 28

    def build_big_editor(self):
        wrap = tk.Frame(self.main, bg=BG)          # plain frame: CTk widgets cannot be placed with width/height
        ov = ctk.CTkFrame(wrap, fg_color=PANEL, border_width=2, border_color=ACC_DIM, corner_radius=16)
        ov.pack(fill="both", expand=True)
        self.big = wrap
        self.big_open = False
        head = tk.Frame(ov, bg=PANEL)
        head.pack(fill="x", padx=18, pady=(14, 6))
        self.big_title = tl(head, "", 13, ACC, True, PANEL)
        self.big_title.pack(side="left")
        dark_button(head, "Close  \u2715", self.close_big, width=92, height=30).pack(side="right")
        ctk.CTkCheckBox(head, text="Block complete", variable=self.blk_var, command=self.toggle_block,
                        fg_color=ACC, hover_color=ACC_HI, border_color=ACC_DIM, checkmark_color="#000000",
                        text_color=TEXT, corner_radius=5, border_width=2, checkbox_width=20,
                        checkbox_height=20).pack(side="right", padx=14)
        dark_button(head, "Copy", self.copy_big, width=64, height=30).pack(side="right")
        host = tk.Frame(ov, bg=EDBG)
        host.pack(fill="both", expand=True, padx=14, pady=(2, 14))
        ed = tk.Text(host, bg=EDBG, fg=TEXT, insertbackground=ACC, relief="flat", font=self.ef, undo=True,
                     wrap="none", padx=16, pady=14, selectbackground=SEL, highlightthickness=0, bd=0)
        ys = ctk.CTkScrollbar(host, command=ed.yview, button_color=BORDER, button_hover_color=ACC_DIM, width=12)
        xs = ctk.CTkScrollbar(host, orientation="horizontal", command=ed.xview, button_color=BORDER,
                              button_hover_color=ACC_DIM, height=12)
        ed.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        ys.pack(side="right", fill="y")
        xs.pack(side="bottom", fill="x")
        ed.pack(fill="both", expand=True)
        self.setup_editor(ed, lambda: self.cur_block)
        self.big_ed = ed
        self.bind("<Escape>", lambda e: self.close_big())

    def open_big(self):
        if self.cur is None or self.big_open:
            return
        key = self.cur_block
        label_ = {k: lab for k, lab, _ in BLOCKS}[key]
        self.big_title.configure(text=f"Day {self.cur}  \u00b7  {label_} notes")
        ed = self.big_ed
        self.loading = True
        ed.delete("1.0", "end")
        ed.insert("1.0", self.state_["days"][str(self.cur)]["notes"].get(key, ""))
        ed.edit_reset()
        ed.edit_modified(False)
        self.highlight(ed)
        self.loading = False
        m = self.BIG_MARGIN
        self.big.place(x=m, y=m, relwidth=1, relheight=1, width=-2 * m, height=-2 * m)
        self.big.lift()
        self.big_open = True
        ed.mark_set("insert", "end-1c")
        ed.see("insert")
        ed.focus_set()

    def close_big(self):
        if not self.big_open:
            return
        self.big.place_forget()
        self.big_open = False
        if self.cur is not None:
            key = self.cur_block
            ed = self.editors[key][1]
            self.loading = True
            ed.delete("1.0", "end")
            ed.insert("1.0", self.state_["days"][str(self.cur)]["notes"].get(key, ""))
            ed.edit_reset()
            ed.edit_modified(False)
            self.highlight(ed)
            self.loading = False
            ed.focus_set()

    def copy_big(self):
        self.clipboard_clear()
        self.clipboard_append(self.big_ed.get("1.0", "end-1c"))
        self.set_status("Copied to clipboard")

    def select_block(self, key):
        self.cur_block = key
        if self.shown and self.shown != key:
            self.editors[self.shown][0].pack_forget()
        fr, ed = self.editors[key]
        if self.shown != key:
            fr.pack(fill="both", expand=True)
        self.shown = key
        if self.tab_on != key:
            if self.tab_on:
                self.tab_btns[self.tab_on].configure(fg_color=FIELD, text_color=TEXT, border_color=BORDER)
            self.tab_btns[key].configure(fg_color=ACC, text_color=ON_ACC, border_color=ACC)
            self.tab_on = key
        if self.cur is None:
            return
        d = self.state_["days"][str(self.cur)]
        if key not in self.loaded:           # load notes lazily, only for the visible tab
            self.loaded.add(key)
            self.loading = True
            ed.delete("1.0", "end")
            ed.insert("1.0", d["notes"].get(key, ""))
            ed.edit_reset()
            ed.edit_modified(False)
            self.highlight(ed)
            self.loading = False
        self.lesson.configure(state="normal")
        self.lesson.delete("1.0", "end")
        self.lesson.insert("1.0", d["lesson"].get(key) or "(No instructions found for this block.)")
        self.lesson.configure(state="disabled")
        self.blk_var.set(bool(d["bdone"].get(key)))
        ed.focus_set()

    def on_tab(self, ev):
        ev.widget.insert("insert", "    ")
        return "break"

    def on_return(self, ev):
        w = ev.widget
        line = w.get("insert linestart", "insert")
        indent = re.match(r"[ \t]*", line).group(0)
        prv, nxt = w.get("insert-1c", "insert"), w.get("insert", "insert+1c")
        if prv in "([{" and prv and PAIRS.get(prv) == nxt:
            # cursor sits between a bracket pair: open a block like VS Code does
            w.insert("insert", "\n" + indent + "    " + "\n" + indent)
            w.mark_set("insert", "insert-1l lineend")
        else:
            if line.rstrip().endswith((":", "{", "(", "[")):
                indent += "    "
            w.insert("insert", "\n" + indent)
        w.see("insert")
        return "break"

    def on_key(self, ev):
        """Auto-close brackets and quotes like VS Code."""
        w, ch = ev.widget, ev.char
        if not ch or len(ch) != 1 or (ch not in PAIRS and ch not in CLOSERS):
            return None
        sel = bool(w.tag_ranges("sel"))
        prv, nxt = w.get("insert-1c", "insert"), w.get("insert", "insert+1c")
        if ch in CLOSERS and nxt == ch and not sel:      # typing a closer over an existing one
            w.mark_set("insert", "insert+1c")
            return "break"
        if ch not in PAIRS:
            return None
        if sel:                                          # wrap the selection
            a, b = w.index("sel.first"), w.index("sel.last")
            text = w.get(a, b)
            w.delete(a, b)
            w.insert(a, ch + text + PAIRS[ch])
            return "break"
        if ch in "\"'`":
            before = w.get("insert linestart", "insert")
            # triple quotes: """ -> """|"""
            if ch != "`" and before.endswith(ch * 2) and not before.endswith(ch * 3) and not nxt.isalnum():
                w.insert("insert", ch * 4)
                w.mark_set("insert", "insert-3c")
                w.see("insert")
                return "break"
            # f"..." r"..." b"..." u"..." prefixes still auto-close
            prefix = re.search(r"(?<![A-Za-z0-9_])[rRbBuUfF]{1,2}$", before)
            if (prv.isalnum() or prv == "_") and not prefix:
                return None                              # e.g. don't -> no pairing
            if nxt.isalnum() or nxt == "_":
                return None
        elif nxt and not nxt.isspace() and nxt not in ")]}\"'`,;:.":
            return None
        w.insert("insert", ch + PAIRS[ch])
        w.mark_set("insert", "insert-1c")
        w.see("insert")
        return "break"

    def on_backspace(self, ev):
        w = ev.widget
        if w.tag_ranges("sel"):
            return None
        prv, nxt = w.get("insert-1c", "insert"), w.get("insert", "insert+1c")
        if prv in PAIRS and PAIRS[prv] == nxt:           # delete the empty pair together
            w.delete("insert-1c", "insert+1c")
            return "break"
        return None

    def on_edit(self, ev, key):
        w = ev.widget
        if not w.edit_modified():
            return
        w.edit_modified(False)
        if self.loading or self.cur is None:
            return
        self.state_["days"][str(self.cur)]["notes"][key] = w.get("1.0", "end-1c")
        self.mark_dirty()
        if key in self.hl_jobs:
            self.after_cancel(self.hl_jobs[key])
        self.hl_jobs[key] = self.after(350, lambda: self.highlight(w))

    def highlight(self, w):
        for tag in ("kw", "num", "fn", "str", "com"):
            w.tag_remove(tag, "1.0", "end")
        text = w.get("1.0", "end-1c")
        if not text or len(text) > 40000:
            return

        def tag(name, pattern, group=0):
            for m in re.finditer(pattern, text):
                w.tag_add(name, f"1.0+{m.start(group)}c", f"1.0+{m.end(group)}c")

        tag("num", r"\b\d+(?:\.\d+)?\b")
        tag("fn", r"\b([A-Za-z_]\w*)(?=\()", 1)
        for m in re.finditer(r"\b[A-Za-z_]\w*\b", text):
            if m.group(0) in KEYWORDS:
                w.tag_add("kw", f"1.0+{m.start()}c", f"1.0+{m.end()}c")
        tag("str", r"\"[^\"\n]*\"|'[^'\n]*'")
        tag("com", r"(?:#|//|--)[^\n]*")

    def zoom(self, d):
        size = max(8, min(30, int(self.ef.cget("size")) + d))
        self.apply_font(size)

    def apply_font(self, size):
        self.ef.configure(size=size)
        self.state_["font"] = size
        if hasattr(self, "font_lbl"):
            self.font_lbl.configure(text=f"{size} pt")
        self.mark_dirty(activity=False)

    # ---------- progress ----------
    def day_progress(self, n):
        d = self.state_["days"].get(str(n))
        if not d:
            return 0.0
        total = len(d["checks"]) + len(BLOCKS)
        done = sum(1 for c in d["checks"] if c["d"]) + sum(1 for v in d["bdone"].values() if v)
        return done / total

    def today_target(self):
        for n in range(1, 91):
            if str(n) in self.state_["days"] and self.day_progress(n) < 1:
                return n
        return 90 if "90" in self.state_["days"] else 1

    def streak(self):
        days = set(self.state_["activity"])
        d = dt.date.today()
        if d.isoformat() not in days:
            d -= dt.timedelta(days=1)
        n = 0
        while d.isoformat() in days:
            n += 1
            d -= dt.timedelta(days=1)
        return n

    def refresh_overall(self):
        nums = [int(k) for k in self.state_["days"]]
        avg = (sum(self.day_progress(n) for n in nums) / max(len(nums), 1)) if nums else 0
        self.side_set(avg)
        self.side_pct.configure(text=f"Overall  {avg * 100:.0f}%")
        s = self.streak()
        self.streak_lbl.configure(text=f"\U0001f525 {s}-day streak" if s else "")

    # ---------- day view ----------
    def step_day(self, delta):
        nums = sorted(int(k) for k in self.state_["days"])
        if not nums:
            return
        i = nums.index(self.cur) if self.cur in nums else 0
        self.open_day(nums[max(0, min(len(nums) - 1, i + delta))])

    def open_day(self, n, tab=None, go=True):
        if str(n) not in self.state_["days"]:
            return
        if self.big_open:
            self.close_big()
        if n != self.cur:
            self.cur = n
            self.loaded = set()
            self.state_["last_day"] = n
            d = self.state_["days"][str(n)]
            wk = min((n - 1) // 7 + 1, 13)
            self.d_title.configure(text=f"Day {n}  \u00b7  {d['title']}")
            self.d_sub.configure(text=f"Week {wk}")
            self.render_checks()
            self.render_score()
        want = self.cur_block
        for k, lab, _ in BLOCKS:
            if tab and (k == tab or lab.lower() == str(tab).lower()):
                want = k
        self.loaded.discard(want)
        self.select_block(want)
        self.update_day_widgets()
        if go:
            self.show_page("Day")

    def render_checks(self):
        d = self.state_["days"][str(self.cur)]
        w = self.chk
        w.configure(state="normal")
        w.delete("1.0", "end")
        for tg in w.tag_names():
            if tg.startswith("c"):
                w.tag_delete(tg)
        if not d["checks"]:
            w.insert("end", "No checklist found for this day. Use 'Block complete' on each tab.")
        for i, c in enumerate(d["checks"]):
            tg = f"c{i}"
            box = "\u2611" if c["d"] else "\u2610"
            txt = c["t"].replace("\n", " ")
            w.insert("end", f"{box}  {txt}" + ("\n" if i < len(d["checks"]) - 1 else ""),
                     (tg, "done") if c["d"] else (tg,))
            w.tag_configure(tg, lmargin1=2, lmargin2=26)
            w.tag_bind(tg, "<Button-1>", lambda e, k=i: self.toggle_check(k))
            w.tag_bind(tg, "<Enter>", lambda e: w.config(cursor="hand2"))
            w.tag_bind(tg, "<Leave>", lambda e: w.config(cursor="arrow"))
        w.tag_raise("done")
        review = self.cur % 7 == 0 or self.cur == 90
        w.configure(height=max(2, min(len(d["checks"]), 4 if review else 7)), state="disabled")
        w.yview_moveto(0)

    def toggle_check(self, i):
        d = self.state_["days"][str(self.cur)]
        c = d["checks"][i]
        c["d"] = not c["d"]
        w = self.chk
        ln = i + 1
        w.configure(state="normal")
        w.delete(f"{ln}.0", f"{ln}.1")
        w.insert(f"{ln}.0", "\u2611" if c["d"] else "\u2610", (f"c{i}",))
        if c["d"]:
            w.tag_add("done", f"{ln}.0", f"{ln}.end")
        else:
            w.tag_remove("done", f"{ln}.0", f"{ln}.end")
        w.configure(state="disabled")
        self.mark_dirty()
        self.update_day_widgets()

    def toggle_block(self):
        d = self.state_["days"][str(self.cur)]
        d["bdone"][self.cur_block] = bool(self.blk_var.get())
        self.mark_dirty()
        self.update_day_widgets()

    def copy_notes(self):
        self.clipboard_clear()
        self.clipboard_append(self.editors[self.cur_block][1].get("1.0", "end-1c"))
        self.set_status("Copied to clipboard")

    def update_day_widgets(self):
        n = self.cur
        if n is None:
            return
        d = self.state_["days"][str(n)]
        p = self.day_progress(n)
        done = sum(1 for c in d["checks"] if c["d"])
        self.chk_count.configure(text=f"{done}/{len(d['checks'])} ticked  \u00b7  day {p * 100:.0f}% complete")
        for key, lab, mins in BLOCKS:
            tick = "\u2713 " if d["bdone"].get(key) else ""
            txt = f"{tick}{lab}  {mins}m"
            if self.tab_text.get(key) != txt:
                self.tab_text[key] = txt
                self.tab_btns[key].configure(text=txt)
        self.refresh_week_panel()
        self.refresh_overall()
        self.dash_dirty = True

    # ---------- scorecard ----------
    def render_score(self):
        for w in self.score_card.winfo_children():
            w.destroy()
        self.score_card.pack_forget()
        self.score_vars = []
        n = self.cur
        if not (n % 7 == 0 or n == 90):
            return
        self.score_card.pack(fill="x", padx=12, pady=(0, 10))
        sc = self.state_["days"][str(n)].setdefault("score", {})
        tl(self.score_card, "WEEKLY SCORECARD", 8, MUTED, True, FIELD).grid(row=0, column=0, columnspan=4,
                                                                          sticky="w", padx=12, pady=(8, 2))
        for i, (lab, mx) in enumerate(SCORE_FIELDS):
            v = tk.StringVar(value=str(sc.get(str(i), "")))
            r, c = 1 + i // 2, (i % 2) * 2
            tl(self.score_card, f"{lab} (/{mx})", 10, TEXT, False, FIELD).grid(row=r, column=c, sticky="w",
                                                                               padx=(12, 4), pady=3)
            tk.Entry(self.score_card, textvariable=v, width=5, bg=EDBG, fg=TEXT, insertbackground=ACC,
                     relief="flat", highlightthickness=1, highlightbackground=BORDER,
                     highlightcolor=ACC).grid(row=r, column=c + 1, padx=(0, 10))
            self.score_vars.append(v)
        for i, v in enumerate(self.score_vars):
            v.trace_add("write", partial(self.on_score, i, v))
        self.score_total = tl(self.score_card, "", 11, OK, True, FIELD)
        self.score_total.grid(row=3, column=2, columnspan=2, sticky="w", padx=12, pady=(2, 8))
        self.update_total()

    def on_score(self, k, var, *_):
        if self.cur is None:
            return
        self.state_["days"][str(self.cur)].setdefault("score", {})[str(k)] = var.get()
        self.update_total()
        self.mark_dirty()

    def update_total(self):
        total = 0
        for i, v in enumerate(self.score_vars):
            try:
                total += max(0, min(SCORE_FIELDS[i][1], int(v.get())))
            except ValueError:
                pass
        ok = total >= 80
        self.score_total.configure(
            text=f"Total {total}/100   " + ("\u2713 you can move on" if ok else "below 80: revisit weak topics"),
            fg=OK if ok else ACC_HI)

    # ---------- search ----------
    def build_search(self):
        f = self.pages["Search"] = tk.Frame(self.main, bg=BG)
        tl(f, "Search", 26, TEXT, True).pack(fill="x")
        tl(f, "Find anything in the lessons, checklists and your own notes.", 11, MUTED).pack(fill="x", pady=(0, 12))
        self.q_var = tk.StringVar()
        self.q_entry = ctk.CTkEntry(f, textvariable=self.q_var, height=40, corner_radius=10, fg_color=FIELD,
                                    border_color=BORDER, text_color=TEXT, placeholder_text="Type at least 2 letters...",
                                    placeholder_text_color=MUTED)
        self.q_entry.pack(fill="x", pady=(0, 10))
        self.q_var.trace_add("write", self.search_later)
        self.res = FastScroll(f, BG)
        self.res.pack(fill="both", expand=True)

    def search_later(self, *_):
        if self.search_job:
            self.after_cancel(self.search_job)
        self.search_job = self.after(220, self.run_search)

    def run_search(self):
        self.res.clear()
        q = self.q_var.get().strip().lower()
        if len(q) < 2:
            return
        count = 0
        for n in range(1, 91):
            d = self.state_["days"].get(str(n))
            if not d:
                continue
            src = [("title", "learn", d["title"])]
            src += [("lesson " + k, k, v) for k, v in d["lesson"].items()]
            src += [("note " + k, k, v) for k, v in d["notes"].items()]
            src += [("checklist", "learn", c["t"]) for c in d["checks"]]
            for where, blk, text in src:
                i = text.lower().find(q)
                if i < 0:
                    continue
                snip = text[max(0, i - 30): i + 70].replace("\n", " ")
                row = tk.Label(self.res.inner, text=f"Day {n} \u00b7 {where}   \u2026{snip}\u2026", anchor="w",
                               bg=PANEL, fg=TEXT, font=(UI_FONT, 10), padx=12, pady=8, cursor="hand2",
                               highlightthickness=1, highlightbackground=BORDER)
                row.pack(fill="x", pady=2, padx=(0, 6))
                row.bind("<Button-1>", lambda e, a=n, b=blk: self.open_day(a, b))
                row.bind("<Enter>", lambda e, r=row: r.configure(bg=FIELD2))
                row.bind("<Leave>", lambda e, r=row: r.configure(bg=PANEL))
                count += 1
                if count >= 40:
                    return

    # ---------- timer ----------
    def build_timer(self):
        f = self.pages["Timer"] = tk.Frame(self.main, bg=BG)
        tl(f, "Study Timer", 26, TEXT, True).pack(fill="x")
        tl(f, "Your 4-hour day: 45 learn \u00b7 60 code \u00b7 90 build \u00b7 30 test \u00b7 15 document.", 11, MUTED
           ).pack(fill="x", pady=(0, 14))
        c = card(f)
        c.pack(fill="x")
        self.t_name = tk.Label(c, text="", fg=ACC, bg=PANEL, font=(UI_FONT, 16, "bold"))
        self.t_name.pack(pady=(22, 0))
        self.t_clock = tk.Label(c, text="", fg=TEXT, bg=PANEL, font=(self.mono, 60, "bold"))
        self.t_clock.pack(pady=4)
        self.t_next = tk.Label(c, text="", fg=MUTED, bg=PANEL, font=(UI_FONT, 10))
        self.t_next.pack(pady=(0, 10))
        row = tk.Frame(c, bg=PANEL)
        row.pack(pady=(4, 22))
        self.t_btn = gold_button(row, "Start", self.timer_toggle, width=110, height=38)
        self.t_btn.pack(side="left", padx=6)
        dark_button(row, "Skip block", lambda: self.timer_skip(False), width=110, height=38).pack(side="left", padx=6)
        dark_button(row, "Reset", self.timer_reset, width=110, height=38).pack(side="left", padx=6)
        self.timer_show()

    def timer_show(self):
        t = self.timer
        key, lab, mins = BLOCKS[t["step"]]
        m, s = divmod(t["left"], 60)
        self.t_name.configure(text=f"{lab}  ({mins} min)")
        self.t_clock.configure(text=f"{m:02d}:{s:02d}")
        self.t_next.configure(text=f"Next: {BLOCKS[t['step'] + 1][1]}" if t["step"] + 1 < len(BLOCKS)
                              else "Last block of the day")

    def timer_toggle(self):
        t = self.timer
        t["running"] = not t["running"]
        self.t_btn.configure(text="Pause" if t["running"] else "Resume")
        if t["running"]:
            self.timer_tick()

    def timer_tick(self):
        t = self.timer
        if not t["running"]:
            return
        if t["left"] <= 0:
            try:
                import winsound
                winsound.Beep(880, 400)
            except Exception:
                self.bell()
            self.timer_skip(True)
            return
        t["left"] -= 1
        self.timer_show()
        t["job"] = self.after(1000, self.timer_tick)

    def timer_skip(self, auto):
        t = self.timer
        if t["step"] + 1 >= len(BLOCKS):
            t["running"] = False
            self.t_btn.configure(text="Start")
            self.t_name.configure(text="Day complete \u2713")
            return
        t["step"] += 1
        t["left"] = BLOCKS[t["step"]][2] * 60
        self.timer_show()
        if auto and t["running"]:
            t["job"] = self.after(1000, self.timer_tick)

    def timer_reset(self):
        t = self.timer
        t["running"] = False
        if t["job"]:
            self.after_cancel(t["job"])
        t.update(step=0, left=BLOCKS[0][2] * 60)
        self.t_btn.configure(text="Start")
        self.timer_show()

    # ---------- import ----------
    def build_import(self):
        f = self.pages["Import"] = tk.Frame(self.main, bg=BG)
        tl(f, "Import your course", 26, TEXT, True).pack(fill="x")
        tl(f, "Pick the whole course document, or paste it below. Importing again keeps your notes and ticks.",
           11, MUTED).pack(fill="x", pady=(0, 12))
        row = tk.Frame(f, bg=BG)
        row.pack(fill="x", pady=(0, 10))
        gold_button(row, "Choose document (.docx / .txt / .md)", self.import_file, height=40).pack(side="left")
        dark_button(row, "Import pasted text", self.import_pasted, height=40).pack(side="left", padx=8)
        self.imp_status = tl(f, "", 11, ACC_HI, True)
        self.imp_status.pack(fill="x", pady=(0, 6))
        c = card(f)
        c.pack(fill="both", expand=True)
        tl(c, "OR PASTE THE COURSE HERE", 8, MUTED, True, PANEL).pack(fill="x", padx=16, pady=(10, 0))
        self.paste = tk.Text(c, bg=EDBG, fg=TEXT, insertbackground=ACC, relief="flat", wrap="word",
                             font=(self.mono, 10), padx=10, pady=8, bd=0, highlightthickness=0,
                             selectbackground=SEL)
        self.paste.pack(fill="both", expand=True, padx=10, pady=(4, 10))

    def import_file(self):
        p = filedialog.askopenfilename(title="Choose your course document", filetypes=[
            ("Course documents", "*.docx *.txt *.md"), ("All files", "*.*")])
        if not p:
            return
        try:
            text = read_course_file(p)
        except Exception as e:
            messagebox.showerror(APP, f"Could not read that file:\n{e}")
            return
        self.do_import(text)

    def import_pasted(self):
        self.do_import(self.paste.get("1.0", "end"))

    def do_import(self, text):
        days = parse_course(text)
        if not days:
            self.imp_status.configure(text="No 'Day 1: Title' lines found. Use the full course document.",
                                      fg="#E5483B")
            return
        merge_days(self.state_, days)
        self.mark_dirty(activity=False)
        self.flush()
        nchecks = sum(len(d["checks"]) for d in days.values())
        self.imp_status.configure(text=f"Imported {len(days)} days and {nchecks} checklist items.", fg=OK)
        self.cur = None
        self.panel_week = None
        self.build_week_cards()
        self.refresh_overall()
        self.open_day(self.today_target(), go=False)
        self.after(700, lambda: self.show_page("Dashboard"))

    # ---------- settings ----------
    def build_settings(self):
        f = self.pages["Settings"] = tk.Frame(self.main, bg=BG)
        tl(f, "Settings", 26, TEXT, True).pack(fill="x")
        tl(f, "Make it yours and keep your progress safe.", 11, MUTED).pack(fill="x", pady=(0, 14))
        c = card(f)
        c.pack(fill="x", pady=(0, 10))
        tl(c, "Notes text size", 12, ACC, True, PANEL).pack(fill="x", padx=16, pady=(12, 0))
        row = tk.Frame(c, bg=PANEL)
        row.pack(fill="x", padx=16, pady=(4, 12))
        self.font_lbl = tl(row, f"{self.ef.cget('size')} pt", 11, ACC_HI, True, PANEL, width=6)
        self.font_lbl.pack(side="right")
        sl = ctk.CTkSlider(row, from_=9, to=24, number_of_steps=15, progress_color=ACC, button_color=ACC_HI,
                           button_hover_color="#ffe08a", fg_color="#2a2417", command=self.on_font_slide)
        sl.set(int(self.ef.cget("size")))
        sl.pack(side="left", fill="x", expand=True)
        c2 = card(f)
        c2.pack(fill="x", pady=(0, 10))
        tl(c2, "Colour theme (applies next time you open the app)", 12, ACC, True, PANEL).pack(fill="x", padx=16, pady=(12, 4))
        self.pal_var = tk.StringVar(value=self.state_.get("palette", "Gold"))
        ctk.CTkOptionMenu(c2, values=list(PALETTES), variable=self.pal_var, command=self.set_pal, width=220,
                          height=34, corner_radius=10, fg_color=FIELD, button_color=FIELD2,
                          button_hover_color=ACC_DIM, text_color=TEXT, dropdown_fg_color=FIELD,
                          dropdown_hover_color=ACC_DIM, dropdown_text_color=TEXT, anchor="w"
                          ).pack(anchor="w", padx=16, pady=(0, 12))
        c3 = card(f)
        c3.pack(fill="x")
        tl(c3, "Backup", 12, ACC, True, PANEL).pack(fill="x", padx=16, pady=(12, 2))
        tl(c3, f"Your data lives in: {DATA_FILE}", 9, MUTED, False, PANEL).pack(fill="x", padx=16)
        row = tk.Frame(c3, bg=PANEL)
        row.pack(fill="x", padx=16, pady=12)
        gold_button(row, "Export progress", self.export_progress, height=36).pack(side="left")
        dark_button(row, "Import progress", self.import_progress, height=36).pack(side="left", padx=8)

    def on_font_slide(self, v):
        size = int(v)
        self.font_lbl.configure(text=f"{size} pt")
        if self.font_job:
            self.after_cancel(self.font_job)
        self.font_job = self.after(180, lambda: self.apply_font(size))   # only apply when the slider settles

    def set_pal(self, name):
        self.state_["palette"] = name
        self.mark_dirty(activity=False)
        self.set_status("Theme saved. Restart the app to see it.")

    def export_progress(self):
        self.flush()
        p = filedialog.asksaveasfilename(defaultextension=".json", initialfile="course-progress.json",
                                         filetypes=[("JSON", "*.json")])
        if p:
            shutil.copy(DATA_FILE, p)
            self.set_status(f"Exported to {p}")

    def import_progress(self):
        p = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if not p or not messagebox.askyesno(APP, "This replaces everything currently in the app. Continue?"):
            return
        try:
            with open(p, "r", encoding="utf-8") as fh:
                data = json.load(fh)
            base = default_state()
            base.update(data)
            self.state_ = base
        except Exception as e:
            messagebox.showerror(APP, f"Could not read that file:\n{e}")
            return
        self.cur = None
        self.panel_week = None
        self.build_week_cards()
        self.refresh_overall()
        if self.state_["days"]:
            self.open_day(self.today_target(), go=False)
        self.mark_dirty(activity=False)
        self.show_page("Dashboard")

    # ---------- persistence ----------
    def mark_dirty(self, activity=True):
        if activity:
            today = dt.date.today().isoformat()
            if today not in self.state_["activity"]:
                self.state_["activity"].append(today)
        if self.save_job:
            self.after_cancel(self.save_job)
        self.save_job = self.after(1200, self.flush)

    def flush(self):
        if self.save_job:
            try:
                self.after_cancel(self.save_job)
            except Exception:
                pass
            self.save_job = None
        try:
            save_state(self.state_)
        except Exception as e:
            self.set_status(f"SAVE FAILED: {e}")

    def set_status(self, msg):
        self.streak_lbl.configure(text=msg[:30])
        self.after(2500, self.refresh_overall)

    def on_close(self):
        self.flush()
        self.destroy()


def main():
    App().mainloop()


if __name__ == "__main__":
    main()
