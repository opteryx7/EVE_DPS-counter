import tkinter as tk
from tkinter import messagebox, ttk
import json
import os
import re
import time
import threading
import glob
from datetime import datetime

try:
    import win32gui
except Exception:
    win32gui = None


CONFIG_FILE = os.path.join(os.path.expanduser("~"), "eve_dps_config.json")
TRANSPARENT = "#010101"
APP_TITLE = "EVE DPS"
APP_SUBTITLE = "v.1.2 by Leffe Brown"

WINDOW_MIN_W = 400
WINDOW_DEFAULT_W = 400
WINDOW_MAX_W = 400
HEADER_H = 44
ROW_H = 28
BOTTOM_PAD = 10
RADIUS = 14

C_BG = "#0f1923"
C_BORDER = "#1e3a5f"
C_DPS_ON = "#4fc3f7"
C_DPS_OFF = "#e8f4fd"
C_IDLE = "#546e7a"
C_TEXT = "#d7e3ea"
C_MUTED = "#7e97aa"
C_HEADER = "#90caf9"
C_ICON = "#6f8aa3"
C_ICON_H = "#90caf9"
C_HIDE = "#e06060"
C_TOOLTIP_BG = "#0b1118"
C_TOOLTIP_FG = "#d7e3ea"
C_EDIT_BG = "#13263a"
C_EDIT_FG = "#ffffff"
C_GRAPH_AXIS = "#54718f"
C_GRAPH_TEXT = "#d7e3ea"
C_GRAPH_GRID = "#13263a"
C_PANEL = "#102131"
C_PANEL_2 = "#13263a"

GRAPH_COLORS = [
    "#4fc3f7", "#81c784", "#ffd54f", "#ff8a65", "#ba68c8",
    "#f06292", "#4dd0e1", "#aed581", "#ffb74d", "#9575cd",
]

WINDOW_SCAN_INTERVAL_MS = 1500
UI_UPDATE_MS = 400
SESSION_TIMEOUT = 30.0
TITLE_PREFIX = "EVE - "
MAX_HISTORY_BATTLES = 10000

OUTGOING_RE = re.compile(
    r"\[\s*(\d{4}\.\d{2}\.\d{2}\s+\d{2}:\d{2}:\d{2})\s*\]"
    r".*?\(combat\).*?"
    r"<color=0xff00ffff><b>(\d+)</b>",
    re.IGNORECASE,
)

INCOMING_RE = re.compile(
    r"\[\s*(\d{4}\.\d{2}\.\d{2}\s+\d{2}:\d{2}:\d{2})\s*\]"
    r".*?\(combat\).*?"
    r"<color=0xffcc0000><b>(\d+)</b>",
    re.IGNORECASE,
)

# Universal target extraction — language-independent, uses EVE color codes.
# 0xff00ffff (cyan)  = outgoing damage value
# 0xffffffff (white) = target bold block (may contain <localized> tags inside)
TARGET_RE_UNIVERSAL = re.compile(
    r"<color=0xff00ffff><b>\d+</b>.+?<b><color=0xffffffff>(.+?)</b>",
    re.IGNORECASE,
)
# KO PvP ship: Name[CORP](<localized hint="ShipEN">한글*)  → hint is the EN ship name
_SHIP_HINT_RE = re.compile(r'\(<localized\s+hint="([^"]+)">', re.IGNORECASE)
# EN PvP ship: Name[CORP](ShipName)
_PLAIN_SHIP_RE = re.compile(r'\(([^)<>]+)\)\s*$')

LISTENER_RE_KO = re.compile(r"청취자\s*:\s*(.+)")
LISTENER_RE_EN = re.compile(r"Listener\s*:\s*(.+)")


def default_config():
    return {
        "window": {"x": 120, "y": 120, "w": WINDOW_DEFAULT_W, "alpha": 92},
        "aliases": {},
        "hidden_chars": [],
        "inc_expanded": True,
        "graph_window": {"x": 180, "y": 180, "w": 1020, "h": 580},
    }


def load_config():
    cfg = default_config()
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                loaded = json.load(f)
            if isinstance(loaded, dict):
                if isinstance(loaded.get("window"), dict):
                    cfg["window"].update(loaded["window"])
                if isinstance(loaded.get("aliases"), dict):
                    cfg["aliases"] = {str(k): str(v) for k, v in loaded["aliases"].items()}
                if isinstance(loaded.get("hidden_chars"), list):
                    cfg["hidden_chars"] = [str(x) for x in loaded["hidden_chars"]]
                if isinstance(loaded.get("inc_expanded"), bool):
                    cfg["inc_expanded"] = loaded["inc_expanded"]
                if isinstance(loaded.get("graph_window"), dict):
                    cfg["graph_window"].update(loaded["graph_window"])
        except Exception:
            pass
    return cfg


def save_config(cfg):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def parse_ts(s: str):
    try:
        return datetime.strptime(s.strip(), "%Y.%m.%d %H:%M:%S")
    except Exception:
        return None


def clamp(n, lo, hi):
    return max(lo, min(hi, n))


def fmt_num(v):
    if v >= 1_000_000:
        return f"{v / 1_000_000:.1f}M"
    if v >= 1000:
        return f"{v / 1000:.1f}k"
    return f"{v:.1f}"


def fmt_int(v):
    if v >= 1_000_000:
        return f"{v / 1_000_000:.1f}M"
    if v >= 1000:
        return f"{v / 1000:.1f}k"
    return str(int(v))


def fmt_dt(ts):
    if ts is None:
        return "-"
    return ts.strftime("%H:%M:%S")


def rrect(x1, y1, x2, y2, r):
    return [
        x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r,
        x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2,
        x1, y2, x1, y2 - r, x1, y1 + r, x1, y1,
    ]


def natural_sort_key(text):
    parts = re.split(r"(\d+)", text.lower())
    out = []
    for p in parts:
        if p.isdigit():
            out.append((0, int(p)))
        else:
            out.append((1, p))
    return out


def fit_text_binary(text, max_width, measure_func):
    if measure_func(text) <= max_width:
        return text
    lo, hi = 0, len(text)
    best = "…"
    while lo <= hi:
        mid = (lo + hi) // 2
        cand = text[:mid].rstrip() + "…"
        if measure_func(cand) <= max_width:
            best = cand
            lo = mid + 1
        else:
            hi = mid - 1
    return best


def short_name(text: str, limit: int = 18):
    if not text:
        return ""
    if len(text) <= limit:
        return text
    return text[: max(1, limit - 1)] + "…"


def extract_target(line: str) -> str:
    """Extract the target ship/name from a combat log line, any client language.

    4 cases handled:
      KO PvP: Name[CORP](<localized hint="ShipEN">한글*)  → hint attr = EN ship name
      EN PvP: Name[CORP](ShipName)                        → plain parens = ship name
      KO NPC: <localized hint="KR_name">EN_Name*          → strip tag, take EN text
      EN NPC: Plain NPC name                              → return as-is
    """
    m = TARGET_RE_UNIVERSAL.search(line)
    if not m:
        return ""
    raw = m.group(1).strip()

    # KO PvP: ship name is inside hint attr of a localised tag within parens
    sh = _SHIP_HINT_RE.search(raw)
    if sh:
        return sh.group(1).strip()

    # EN PvP: plain (ShipName) at end of string
    ps = _PLAIN_SHIP_RE.search(raw)
    if ps:
        return ps.group(1).strip()

    # KO/EN NPC: strip any <localized hint="..."> tags, then strip [brackets] and *
    clean = re.sub(r'<localized\s+hint="[^"]*">', "", raw).strip().rstrip("*").strip()
    return re.sub(r"\[.*", "", clean).strip()


def guess_log_dirs():
    home = os.path.expanduser("~")
    userprofile = os.environ.get("USERPROFILE", "")
    onedrive = os.environ.get("OneDrive", "")
    roots = [p for p in [home, userprofile, onedrive] if p]
    candidates = []
    for root in roots:
        candidates.extend([
            os.path.join(root, "Documents", "EVE", "logs", "Gamelogs"),
            os.path.join(root, "문서", "EVE", "logs", "Gamelogs"),
            os.path.join(root, "EVE", "logs", "Gamelogs"),
            os.path.join(root, "OneDrive", "Documents", "EVE", "logs", "Gamelogs"),
            os.path.join(root, "OneDrive", "문서", "EVE", "logs", "Gamelogs"),
        ])
    seen = set()
    out = []
    for p in candidates:
        norm = os.path.normpath(p)
        low = norm.lower()
        if low not in seen:
            seen.add(low)
            out.append(norm)
    return out


def find_log_base():
    for p in guess_log_dirs():
        if os.path.isdir(p):
            return p
    users_root = r"C:\Users"
    if os.path.isdir(users_root):
        try:
            for user_name in os.listdir(users_root):
                base = os.path.join(users_root, user_name)
                for sub in [
                    os.path.join(base, "Documents", "EVE", "logs", "Gamelogs"),
                    os.path.join(base, "문서", "EVE", "logs", "Gamelogs"),
                    os.path.join(base, "OneDrive", "Documents", "EVE", "logs", "Gamelogs"),
                    os.path.join(base, "OneDrive", "문서", "EVE", "logs", "Gamelogs"),
                ]:
                    if os.path.isdir(sub):
                        return sub
        except Exception:
            pass
    return None


def list_running_eve_characters_from_windows():
    chars = set()
    if win32gui is None:
        return chars

    def callback(hwnd, _):
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return
            title = (win32gui.GetWindowText(hwnd) or "").strip()
            if not title.startswith(TITLE_PREFIX):
                return
            name = title[len(TITLE_PREFIX):].strip()
            if name:
                chars.add(name)
        except Exception:
            return

    try:
        win32gui.EnumWindows(callback, None)
    except Exception:
        pass
    return chars


class DPSEngine:
    def __init__(self, char_name):
        self.char_name = char_name
        self.log_base = find_log_base()
        self.log_file = None
        self.file_pos = 0

        self.current_battle = None
        self.battles = []
        self._lock = threading.Lock()
        self._stop = False
        self._thread = None

    def start(self):
        if self._thread and self._thread.is_alive():
            return
        self._stop = False
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def stop(self):
        self._stop = True

    def hard_reset(self):
        with self._lock:
            self.current_battle = None
            self.battles = []
        self.log_file = None
        self.file_pos = 0

    def _run(self):
        while not self._stop:
            try:
                self._tick()
            except Exception:
                pass
            time.sleep(0.5)

    def _find_log_for_char(self):
        if not self.log_base or not os.path.isdir(self.log_base):
            self.log_base = find_log_base()
        if not self.log_base or not os.path.isdir(self.log_base):
            return None

        files = sorted(glob.glob(os.path.join(self.log_base, "*.txt")), reverse=True)
        for fp in files:
            try:
                with open(fp, "r", encoding="utf-8", errors="ignore") as f:
                    hdr = f.read(3000)
                m = LISTENER_RE_KO.search(hdr) or LISTENER_RE_EN.search(hdr)
                if m and m.group(1).strip() == self.char_name:
                    return fp
            except Exception:
                continue
        return None

    def _new_battle(self, ts: datetime):
        return {
            "start_ts": ts,
            "last_event_ts": ts,
            "last_event_mono": time.monotonic(),
            "last_outgoing_ts": None,
            "total_dmg": 0,
            "active_elapsed": 0.0,
            "hits": [],
            "first_target": "",
            "inc_total_dmg": 0,
            "inc_hits": [],
        }

    def _finalize_current_battle(self):
        if not self.current_battle:
            return
        battle = self.current_battle
        inc_first = battle.get("inc_first_ts")
        inc_last  = battle.get("inc_last_ts")
        inc_elapsed = (inc_last - inc_first).total_seconds() if (inc_first and inc_last and inc_first != inc_last) else 0.0
        inc_dmg = battle.get("inc_total_dmg", 0)
        record = {
            "start_ts": battle["start_ts"],
            "end_ts": battle["last_event_ts"],
            "total_dmg": battle["total_dmg"],
            "active_elapsed": battle["active_elapsed"],
            "dps": (battle["total_dmg"] / battle["active_elapsed"]) if battle["active_elapsed"] > 0 else 0.0,
            "hits": list(battle["hits"]),
            "first_target": battle["first_target"],
            "inc_total_dmg": inc_dmg,
            "inc_elapsed": inc_elapsed,
            "inc_dps": (inc_dmg / inc_elapsed) if inc_elapsed > 0 else 0.0,
            "inc_hits": list(battle.get("inc_hits", [])),
        }
        self.battles.append(record)
        self.battles = self.battles[-30:]
        self.current_battle = None

    def _register_event(self, ts: datetime):
        if self.current_battle is None:
            self.current_battle = self._new_battle(ts)
            return

        gap = (ts - self.current_battle["last_event_ts"]).total_seconds()
        if gap > SESSION_TIMEOUT:
            self._finalize_current_battle()
            self.current_battle = self._new_battle(ts)
        else:
            self.current_battle["last_event_ts"] = ts
            self.current_battle["last_event_mono"] = time.monotonic()

    def _register_outgoing(self, ts: datetime, dmg: int, target: str):
        if self.current_battle is None:
            self.current_battle = self._new_battle(ts)

        if self.current_battle["first_target"] == "":
            self.current_battle["first_target"] = target

        if self.current_battle["last_outgoing_ts"] is not None:
            gap = (ts - self.current_battle["last_outgoing_ts"]).total_seconds()
            self.current_battle["active_elapsed"] += max(0.0, gap)

        self.current_battle["last_outgoing_ts"] = ts
        self.current_battle["total_dmg"] += dmg
        self.current_battle["hits"].append((ts, dmg))

    def _register_incoming(self, ts: datetime, dmg: int):
        if self.current_battle is None:
            self.current_battle = self._new_battle(ts)
        b = self.current_battle
        b["inc_total_dmg"] = b.get("inc_total_dmg", 0) + dmg
        b.setdefault("inc_hits", []).append((ts, dmg))
        # Track first/last incoming hit timestamp for elapsed calculation
        if b.get("inc_first_ts") is None:
            b["inc_first_ts"] = ts
        b["inc_last_ts"] = ts

    def _check_timeout(self):
        if self.current_battle is None:
            return
        last_mono = self.current_battle["last_event_mono"]
        if (time.monotonic() - last_mono) >= SESSION_TIMEOUT:
            self._finalize_current_battle()

    def _tick(self):
        nf = self._find_log_for_char()
        if nf != self.log_file:
            self.log_file = nf
            if self.log_file and os.path.exists(self.log_file):
                try:
                    with open(self.log_file, "r", encoding="utf-8", errors="ignore") as f:
                        f.seek(0, os.SEEK_END)
                        self.file_pos = f.tell()
                except Exception:
                    self.file_pos = 0
            else:
                self.file_pos = 0

        if not self.log_file:
            return

        try:
            with open(self.log_file, "r", encoding="utf-8", errors="ignore") as f:
                f.seek(self.file_pos)
                chunk = f.read()
                self.file_pos = f.tell()
        except Exception:
            return

        events = []

        # Split chunk into logical log entries by timestamp header.
        # Korean (and some other) clients write one combat event across
        # multiple physical lines, so splitting on \n alone loses the target.
        # Splitting on the timestamp pattern keeps each event intact.
        _ENTRY_RE = re.compile(
            r"(?=\[\s*\d{4}\.\d{2}\.\d{2}\s+\d{2}:\d{2}:\d{2}\s*\])"
        )
        for entry in _ENTRY_RE.split(chunk):
            if not entry.strip():
                continue
            # Flatten newlines so regexes see the full entry as one string
            flat = entry.replace("\n", " ").replace("\r", " ")

            m = OUTGOING_RE.search(flat)
            if m:
                ts = parse_ts(m.group(1))
                if ts:
                    try:
                        dmg = int(m.group(2))
                        target = extract_target(flat)
                        events.append((ts, "out", dmg, target))
                    except Exception:
                        pass
                continue

            m = INCOMING_RE.search(flat)
            if m:
                ts = parse_ts(m.group(1))
                if ts:
                    try:
                        dmg = int(m.group(2))
                        events.append((ts, "in", dmg, ""))
                    except Exception:
                        pass

        events.sort(key=lambda x: (x[0], 0 if x[1] == "out" else 1))

        with self._lock:
            for ts, kind, dmg, target in events:
                self._register_event(ts)
                if kind == "out":
                    self._register_outgoing(ts, dmg, target)
                elif kind == "in":
                    self._register_incoming(ts, dmg)
                if self.current_battle:
                    self.current_battle["last_event_ts"] = ts
                    self.current_battle["last_event_mono"] = time.monotonic()
            self._check_timeout()

    def get_status(self):
        with self._lock:
            finished = list(self.battles)
            current = dict(self.current_battle) if self.current_battle else None

        total_damage = 0
        total_secs = 0.0
        total_inc_dmg = 0
        total_inc_elapsed = 0.0
        for b in finished:
            total_damage += b["total_dmg"]
            total_secs += b["active_elapsed"]
            total_inc_dmg += b.get("inc_total_dmg", 0)
            total_inc_elapsed += b.get("inc_elapsed", 0.0)

        recent_dps = 0.0
        recent_inc_dps = 0.0
        if current:
            total_damage += current["total_dmg"]
            total_secs += current["active_elapsed"]
            total_inc_dmg += current.get("inc_total_dmg", 0)
            if current["active_elapsed"] > 0:
                recent_dps = current["total_dmg"] / current["active_elapsed"]
            # Inc DPS: use time between first and last incoming hit
            inc_first = current.get("inc_first_ts")
            inc_last  = current.get("inc_last_ts")
            if inc_first and inc_last and inc_first != inc_last:
                inc_elapsed = (inc_last - inc_first).total_seconds()
                if inc_elapsed > 0:
                    recent_inc_dps = current.get("inc_total_dmg", 0) / inc_elapsed
                    total_inc_elapsed += inc_elapsed
            elif current.get("inc_total_dmg", 0) > 0 and current["active_elapsed"] > 0:
                recent_inc_dps = current.get("inc_total_dmg", 0) / current["active_elapsed"]
                total_inc_elapsed += current["active_elapsed"]

        total_dps = (total_damage / total_secs) if total_secs > 0 else 0.0
        total_inc_dps = (total_inc_dmg / total_inc_elapsed) if total_inc_elapsed > 0 else 0.0

        return {
            "dps": recent_dps,
            "tdps": total_dps,
            "tdam": total_damage,
            "in_combat": current is not None,
            "inc_dps": recent_inc_dps,
            "inc_tdps": total_inc_dps,
            "inc_tdam": total_inc_dmg,
        }

    def get_battle_records(self):
        with self._lock:
            records = list(self.battles)
            if self.current_battle:
                c = self.current_battle
                records.append({
                    "start_ts": c["start_ts"],
                    "end_ts": c["last_event_ts"],
                    "total_dmg": c["total_dmg"],
                    "active_elapsed": c["active_elapsed"],
                    "dps": (c["total_dmg"] / c["active_elapsed"]) if c["active_elapsed"] > 0 else 0.0,
                    "hits": list(c["hits"]),
                    "first_target": c["first_target"],
                })
            return records


class Tooltip:
    def __init__(self, root):
        self.root = root
        self.win = None

    def show(self, x, y, text):
        self.hide()
        self.win = tk.Toplevel(self.root)
        self.win.overrideredirect(True)
        self.win.attributes("-topmost", True)
        self.win.configure(bg=C_TOOLTIP_BG)
        tk.Label(self.win, text=text, bg=C_TOOLTIP_BG, fg=C_TOOLTIP_FG, font=("Consolas", 9), padx=7, pady=4).pack()
        self.win.geometry(f"+{x + 12}+{y + 12}")

    def hide(self):
        if self.win and self.win.winfo_exists():
            try:
                self.win.destroy()
            except Exception:
                pass
        self.win = None


class GraphWindow(tk.Toplevel):
    TIME_COL_W = 133
    ENEMY_COL_W = 120
    STAT_COL_W = 70
    HEADER_H = 40

    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.attributes("-alpha", 0.96)
        self.config(bg=TRANSPARENT)
        self.attributes("-transparentcolor", TRANSPARENT)

        self._dx = 0
        self._dy = 0
        gw = self.app.cfg.get("graph_window", {})
        self._win_w = int(gw.get("w", 1020))
        self._win_h = int(gw.get("h", 580))
        gx = int(gw.get("x", 180))
        gy = int(gw.get("y", 180))
        self.geometry(f"{self._win_w}x{self._win_h}+{gx}+{gy}")

        self.bind("<ButtonPress-1>", self._drag_start)
        self.bind("<B1-Motion>", self._drag_move)
        self.bind("<ButtonRelease-1>", self._drag_end)

        outer = tk.Frame(self, bg=TRANSPARENT)
        outer.pack(fill="both", expand=True)

        self.bg_canvas = tk.Canvas(outer, bg=TRANSPARENT, highlightthickness=0)
        self.bg_canvas.pack(fill="both", expand=True)
        self.bg_canvas.bind("<Configure>", self._on_resize)
        self.bg_canvas.bind("<ButtonPress-1>", self._drag_start)
        self.bg_canvas.bind("<B1-Motion>", self._drag_move)
        self.bg_canvas.bind("<ButtonRelease-1>", self._drag_end)

        self.title_bar = tk.Frame(self.bg_canvas, bg=C_PANEL)
        self.body = tk.Frame(self.bg_canvas, bg=C_BG)

        self.title_label = tk.Label(self.title_bar, text="Recent Actual Damage Dealt", bg=C_PANEL, fg=C_HEADER, font=("Consolas", 11, "bold"))
        self.title_label.pack(side="left", padx=10, pady=6)

        self.btn_close = tk.Label(self.title_bar, text="✕", bg=C_PANEL, fg=C_ICON, font=("Segoe UI Symbol", 10))
        self.btn_close.pack(side="right", padx=(0, 8))

        self.btn_close.bind("<Button-1>", lambda _e: self.destroy())

        self.title_bar.bind("<ButtonPress-1>", self._drag_start)
        self.title_bar.bind("<B1-Motion>", self._drag_move)
        self.title_bar.bind("<ButtonRelease-1>", self._drag_end)
        self.title_label.bind("<ButtonPress-1>", self._drag_start)
        self.title_label.bind("<B1-Motion>", self._drag_move)
        self.title_label.bind("<ButtonRelease-1>", self._drag_end)

        style = ttk.Style(self)
        try:
            style.theme_use("default")
        except Exception:
            pass
        style.configure("Blue.TNotebook", background=C_BG, borderwidth=0)
        style.configure("Blue.TNotebook.Tab", background=C_PANEL_2, foreground=C_TEXT, padding=(12, 6))
        style.map("Blue.TNotebook.Tab", background=[("selected", C_PANEL)], foreground=[("selected", C_HEADER)])
        style.configure("History.Treeview", rowheight=24, background=C_PANEL, fieldbackground=C_PANEL, foreground=C_TEXT)
        style.configure("History.Treeview.Heading", background=C_PANEL_2, foreground=C_HEADER)
        style.configure("Blue.Vertical.TScrollbar",
                        background=C_BORDER, troughcolor=C_BG,
                        arrowcolor=C_HEADER, bordercolor=C_BG, lightcolor=C_BORDER, darkcolor=C_BORDER)
        style.configure("Blue.Horizontal.TScrollbar",
                        background=C_BORDER, troughcolor=C_BG,
                        arrowcolor=C_HEADER, bordercolor=C_BG, lightcolor=C_BORDER, darkcolor=C_BORDER)
        style.map("Blue.Vertical.TScrollbar",   background=[("active", C_HEADER)])
        style.map("Blue.Horizontal.TScrollbar", background=[("active", C_HEADER)])

        self.notebook = ttk.Notebook(self.body, style="Blue.TNotebook")
        self.notebook.pack(fill="both", expand=True, padx=10, pady=(8, 10))

        self.graph_frame   = tk.Frame(self.notebook, bg=C_BG)
        self.history_frame = tk.Frame(self.notebook, bg=C_BG)
        self.notebook.add(self.graph_frame,   text="Recent Battle DPS")
        self.notebook.add(self.history_frame, text="DPS History")

        self.canvas = tk.Canvas(self.graph_frame, bg=C_BG, highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)

        self.history_header_canvas = None
        self.tree = None
        self.tree_cols = []
        self.h_scroll = None

        self._build_history_tree()

        # Defer first render until the window is fully laid out
        self.after(50, self._first_render)

    def _first_render(self):
        self.update_idletasks()
        self._on_resize(type("E", (), {"width": self._win_w, "height": self._win_h})())
        self.after(950, self._update_loop)

    def _drag_start(self, e):
        self._dx = e.x_root - self.winfo_x()
        self._dy = e.y_root - self.winfo_y()

    def _drag_move(self, e):
        self.geometry(f"+{e.x_root - self._dx}+{e.y_root - self._dy}")

    def _drag_end(self, _e):
        # Save graph window position
        try:
            self.app.cfg["graph_window"] = {
                "x": self.winfo_x(), "y": self.winfo_y(),
                "w": self._win_w,    "h": self._win_h,
            }
            save_config(self.app.cfg)
        except Exception:
            pass

    def _on_resize(self, e):
        w, h = getattr(e, "width", self._win_w), getattr(e, "height", self._win_h)
        if w < 10 or h < 10:
            return
        self._win_w, self._win_h = w, h
        self.bg_canvas.delete("all")
        self.bg_canvas.create_polygon(
            rrect(1, 1, self._win_w - 1, self._win_h - 1, 16),
            smooth=True,
            fill=C_BG,
            outline=C_BORDER,
            width=1.4,
        )
        self.bg_canvas.create_window(1, 1, anchor="nw", window=self.title_bar, width=self._win_w - 2, height=34)
        self.bg_canvas.create_window(1, 35, anchor="nw", window=self.body, width=self._win_w - 2, height=self._win_h - 36)
        # Re-draw whichever graph canvases are currently visible
        self.render()

    def _get_joint_battles(self):
        chars = self.app.sorted_characters()
        intervals = []

        for char in chars:
            eng = self.app.engines.get(char)
            if not eng:
                continue
            recs = eng.get_battle_records()
            for rec in recs:
                intervals.append({
                    "char": char,
                    "start_ts": rec["start_ts"],
                    "end_ts": rec["end_ts"],
                    "record": rec,
                })

        if not intervals:
            return []

        intervals.sort(key=lambda x: x["start_ts"])
        merged = []
        cur_start = intervals[0]["start_ts"]
        cur_end = intervals[0]["end_ts"]
        cur_items = [intervals[0]]

        for item in intervals[1:]:
            if item["start_ts"] <= cur_end:
                if item["end_ts"] > cur_end:
                    cur_end = item["end_ts"]
                cur_items.append(item)
            else:
                merged.append((cur_start, cur_end, cur_items))
                cur_start = item["start_ts"]
                cur_end = item["end_ts"]
                cur_items = [item]
        merged.append((cur_start, cur_end, cur_items))

        joint = []
        visible_chars = self.app.sorted_characters()
        for start_ts, end_ts, items in merged:
            battle = {
                "start_ts": start_ts,
                "end_ts": end_ts,
                "enemy": "",
                "chars": {},
            }
            for char in visible_chars:
                battle["chars"][char] = {
                    "hits": [],
                    "total_dmg": 0,
                    "active_elapsed": 0.0,
                    "dps": 0.0,
                    "inc_hits": [],
                    "inc_total_dmg": 0,
                    "inc_elapsed": 0.0,
                    "inc_dps": 0.0,
                }

            first_targets = []
            for item in items:
                char = item["char"]
                rec = item["record"]
                entry = battle["chars"][char]
                entry["hits"].extend(rec["hits"])
                entry["total_dmg"] += rec["total_dmg"]
                entry["active_elapsed"] += rec["active_elapsed"]
                entry["inc_hits"].extend(rec.get("inc_hits", []))
                entry["inc_total_dmg"] += rec.get("inc_total_dmg", 0)
                entry["inc_elapsed"] += rec.get("inc_elapsed", 0.0)
                if rec.get("first_target"):
                    first_targets.append((rec["start_ts"], rec["first_target"]))

            if first_targets:
                first_targets.sort(key=lambda x: x[0])
                battle["enemy"] = short_name(first_targets[0][1], 22)

            for char, entry in battle["chars"].items():
                if entry["active_elapsed"] > 0:
                    entry["dps"] = entry["total_dmg"] / entry["active_elapsed"]
                if entry["inc_elapsed"] > 0:
                    entry["inc_dps"] = entry["inc_total_dmg"] / entry["inc_elapsed"]

            joint.append(battle)

        joint.sort(key=lambda x: x["end_ts"], reverse=True)
        return joint[:MAX_HISTORY_BATTLES]

    def _history_total_width(self):
        chars = self.app.sorted_characters()
        return self.TIME_COL_W + self.ENEMY_COL_W + (len(chars) * 2 * self.STAT_COL_W)

    def _sync_xview(self, *args):
        if self.tree:
            self.tree.xview(*args)
        if self.history_header_canvas:
            self.history_header_canvas.xview(*args)

    def _on_tree_xscroll(self, first, last):
        if self.h_scroll:
            self.h_scroll.set(first, last)
        if self.history_header_canvas:
            try:
                self.history_header_canvas.xview_moveto(first)
            except Exception:
                pass

    def _draw_history_header(self):
        if not self.history_header_canvas:
            return

        self.history_header_canvas.delete("all")
        total_w = self._history_total_width()
        self.history_header_canvas.configure(scrollregion=(0, 0, total_w, self.HEADER_H))

        x = 0

        def rect(xx, yy, ww, hh, fill):
            self.history_header_canvas.create_rectangle(xx, yy, xx + ww, yy + hh, fill=fill, outline=C_BORDER, width=1)

        def text(cx, cy, s, color, font):
            self.history_header_canvas.create_text(cx, cy, text=s, fill=color, font=font)

        rect(x, 0, self.TIME_COL_W, 20, C_PANEL_2)
        rect(x, 20, self.TIME_COL_W, 20, C_PANEL)
        text(x + self.TIME_COL_W / 2, 10, "Battle", C_HEADER, ("Consolas", 10, "bold"))
        text(x + self.TIME_COL_W / 2, 30, "Time", C_TEXT, ("Consolas", 9))
        x += self.TIME_COL_W

        rect(x, 0, self.ENEMY_COL_W, 20, C_PANEL_2)
        rect(x, 20, self.ENEMY_COL_W, 20, C_PANEL)
        text(x + self.ENEMY_COL_W / 2, 10, "First Target", C_HEADER, ("Consolas", 10, "bold"))
        text(x + self.ENEMY_COL_W / 2, 30, "Enemy", C_TEXT, ("Consolas", 9))
        x += self.ENEMY_COL_W

        for char in self.app.sorted_characters():
            group_w = self.STAT_COL_W * 2
            rect(x, 0, group_w, 20, C_PANEL_2)
            rect(x, 20, self.STAT_COL_W, 20, C_PANEL)
            rect(x + self.STAT_COL_W, 20, self.STAT_COL_W, 20, C_PANEL)
            text(x + group_w / 2, 10, char, C_HEADER, ("Consolas", 10, "bold"))
            text(x + self.STAT_COL_W / 2, 30, "DPS", C_TEXT, ("Consolas", 9))
            text(x + self.STAT_COL_W + self.STAT_COL_W / 2, 30, "DMG", C_TEXT, ("Consolas", 9))
            x += group_w

    def _build_history_tree(self):
        for child in self.history_frame.winfo_children():
            child.destroy()

        chars = self.app.sorted_characters()
        columns = ["time", "enemy"]
        for char in chars:
            columns.append(f"{char}_dps")
            columns.append(f"{char}_dmg")

        self.history_header_canvas = tk.Canvas(
            self.history_frame,
            bg=C_BG,
            height=self.HEADER_H,
            highlightthickness=0,
            bd=0,
        )
        self.history_header_canvas.pack(fill="x", side="top")

        grid_wrap = tk.Frame(self.history_frame, bg=C_BG)
        grid_wrap.pack(fill="both", expand=True)

        tree = ttk.Treeview(grid_wrap, columns=columns, show="headings", style="History.Treeview")
        vsb = ttk.Scrollbar(grid_wrap, orient="vertical",   command=tree.yview,        style="Blue.Vertical.TScrollbar")
        hsb = ttk.Scrollbar(grid_wrap, orient="horizontal", command=self._sync_xview,  style="Blue.Horizontal.TScrollbar")
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=self._on_tree_xscroll)

        tree.heading("time", text="")
        tree.column("time", width=self.TIME_COL_W, anchor="w", stretch=False)
        tree.heading("enemy", text="")
        tree.column("enemy", width=self.ENEMY_COL_W, anchor="w", stretch=False)

        for char in chars:
            tree.heading(f"{char}_dps", text="")
            tree.column(f"{char}_dps", width=self.STAT_COL_W, anchor="e", stretch=False)
            tree.heading(f"{char}_dmg", text="")
            tree.column(f"{char}_dmg", width=self.STAT_COL_W, anchor="e", stretch=False)

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        grid_wrap.grid_rowconfigure(0, weight=1)
        grid_wrap.grid_columnconfigure(0, weight=1)

        self.tree = tree
        self.tree_cols = columns
        self.h_scroll = hsb
        self._draw_history_header()

    def _update_history_rows(self):
        chars = self.app.sorted_characters()
        desired_cols = ["time", "enemy"]
        for char in chars:
            desired_cols.append(f"{char}_dps")
            desired_cols.append(f"{char}_dmg")

        if desired_cols != self.tree_cols:
            self._build_history_tree()
        else:
            self._draw_history_header()

        battles = self._get_joint_battles()

        for item in self.tree.get_children():
            self.tree.delete(item)

        for battle in battles:
            row = [
                f"{fmt_dt(battle['start_ts'])} ~ {fmt_dt(battle['end_ts'])}",
                battle.get("enemy", ""),
            ]
            for char in chars:
                entry = battle["chars"][char]
                row.append(fmt_num(entry["dps"]))
                row.append(fmt_int(entry["total_dmg"]))
            self.tree.insert("", "end", values=row)

    def _render_graph(self):
        self.canvas.delete("all")
        w = max(10, self.canvas.winfo_width())
        h = max(10, self.canvas.winfo_height())

        pad_l, pad_r, pad_t, pad_b = 70, 20, 95, 60
        x1, y1, x2, y2 = pad_l, pad_t, w - pad_r, h - pad_b

        self.canvas.create_rectangle(0, 0, w, h, fill=C_BG, outline=C_BG)
        self.canvas.create_text(18, 16, anchor="w", text="Recent Actual Damage Dealt", fill=C_HEADER, font=("Consolas", 12, "bold"))

        battles = self._get_joint_battles()
        if not battles:
            self.canvas.create_text(w // 2, h // 2, text="No recent battle data", fill=C_MUTED, font=("Consolas", 12))
            return

        battle = battles[0]
        self.canvas.create_text(
            w - 18, 16, anchor="e",
            text=f"{fmt_dt(battle['start_ts'])} ~ {fmt_dt(battle['end_ts'])}",
            fill=C_MUTED, font=("Consolas", 10),
        )

        total_secs = max(1.0, (battle["end_ts"] - battle["start_ts"]).total_seconds())
        max_y = 1.0
        series = []

        for idx, char in enumerate(self.app.sorted_characters()):
            entry = battle["chars"][char]
            points = []
            for ts, dmg in entry["hits"]:
                offset = max(0.0, (ts - battle["start_ts"]).total_seconds())
                points.append((offset, dmg))
                max_y = max(max_y, float(dmg))
            series.append({"char": char, "color": GRAPH_COLORS[idx % len(GRAPH_COLORS)], "points": points})

        max_y *= 1.15
        legend_x, legend_y = x1, 44
        per_row = max(1, (w - 100) // 180)
        for idx, item in enumerate(series):
            lx = legend_x + (idx % per_row) * 180
            ly = legend_y + (idx // per_row) * 20
            self.canvas.create_line(lx, ly, lx + 18, ly, fill=item["color"], width=3)
            self.canvas.create_text(lx + 24, ly, anchor="w", text=item["char"], fill=C_GRAPH_TEXT, font=("Consolas", 9))

        self.canvas.create_line(x1, y2, x2, y2, fill=C_GRAPH_AXIS, width=1)
        self.canvas.create_line(x1, y1, x1, y2, fill=C_GRAPH_AXIS, width=1)

        for i in range(5):
            yy  = y2 - ((y2 - y1) * i / 4.0)
            val = max_y * i / 4.0
            self.canvas.create_line(x1, yy, x2, yy, fill=C_GRAPH_GRID, width=1)
            self.canvas.create_text(x1 - 8, yy, anchor="e", text=f"{int(val)}", fill=C_GRAPH_TEXT, font=("Consolas", 9))

        x_ticks = min(8, max(2, int(total_secs) + 1))
        for i in range(x_ticks):
            ratio = i / (x_ticks - 1 if x_ticks > 1 else 1)
            xx  = x1 + ((x2 - x1) * ratio)
            val = total_secs * ratio
            self.canvas.create_line(xx, y2, xx, y2 + 4, fill=C_GRAPH_AXIS, width=1)
            self.canvas.create_text(xx, y2 + 14, anchor="n", text=f"{val:.0f}s", fill=C_GRAPH_TEXT, font=("Consolas", 9))

        self.canvas.create_text(18, (y1 + y2) / 2, text="Actual Damage", angle=90, fill=C_MUTED, font=("Consolas", 9))

        for item in series:
            coords = []
            for sec, dmg in item["points"]:
                px = x1 + ((sec / total_secs) * (x2 - x1))
                py = y2 - ((dmg / max_y) * (y2 - y1))
                coords.extend([px, py])
            if len(coords) >= 4:
                self.canvas.create_line(*coords, fill=item["color"], width=2, smooth=False)
            for sec, dmg in item["points"]:
                px = x1 + ((sec / total_secs) * (x2 - x1))
                py = y2 - ((dmg / max_y) * (y2 - y1))
                self.canvas.create_oval(px - 3, py - 3, px + 3, py + 3, fill=item["color"], outline=item["color"])
                self.canvas.create_text(px + 6, py - 8, anchor="w", text=str(int(dmg)), fill=item["color"], font=("Consolas", 8, "bold"))

    def render(self):
        self._render_graph()
        self._update_history_rows()

    def _update_loop(self):
        if self.winfo_exists():
            self.render()
            self.after(1000, self._update_loop)


class EVEUnifiedWindow(tk.Toplevel):

    INC_ROW_H = 24   # height of each incoming row
    INC_HDR_H = 20   # height of the "Incoming" section header bar

    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self._dx = 0
        self._dy = 0
        self._cur_w = WINDOW_DEFAULT_W
        self._cur_h = 220
        self.cv = None
        self.tooltip = Tooltip(self)
        self.hover_char = None
        self.inline_editor = None
        self.inline_editor_char = None
        self.font_name = ("Consolas", 10)
        self.font_num = ("Consolas", 10)
        self.font_num_b = ("Consolas", 10, "bold")
        self.row_info = {}
        self.ui_items = {}
        self.header_click_boxes = {}
        self.inc_expanded = self.app.cfg.get("inc_expanded", True)

        self._setup_window()
        self.bind("<Map>", self._on_map_restore)
        self._restore_pos()
        self._build_static_ui()
        self.render_all(force_layout=True)
        self.after(UI_UPDATE_MS, self._update_loop)

    def _setup_window(self):
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.config(bg=TRANSPARENT)
        self.attributes("-transparentcolor", TRANSPARENT)
        self.bind("<ButtonPress-1>", self._drag_start)
        self.bind("<B1-Motion>", self._drag_move)
        self.bind("<ButtonRelease-1>", self._drag_end)

    def _on_map_restore(self, _event):
        try:
            if self.state() == "normal":
                self.after(10, lambda: self.overrideredirect(True))
                self.after(10, lambda: self.attributes("-topmost", True))
                self.after(10, lambda: self.lift())
        except Exception:
            pass

    def _restore_pos(self):
        wc = self.app.cfg.get("window", {})
        self._cur_w = int(clamp(wc.get("w", WINDOW_DEFAULT_W), WINDOW_MIN_W, WINDOW_MAX_W))
        self.geometry(f"{self._cur_w}x{self._cur_h}+{wc.get('x', 120)}+{wc.get('y', 120)}")
        self.attributes("-alpha", clamp(wc.get("alpha", 92), 10, 100) / 100.0)

    def _save_pos(self):
        self.app.cfg["window"] = {
            "x": self.winfo_x(),
            "y": self.winfo_y(),
            "w": self._cur_w,
            "alpha": int(float(self.attributes("-alpha")) * 100),
        }
        save_config(self.app.cfg)

    def _drag_start(self, e):
        if self.inline_editor_char:
            return
        self._dx, self._dy = e.x, e.y
        self.app.pause_updates = True

    def _drag_move(self, e):
        if self.inline_editor_char:
            return
        self.geometry(f"+{self.winfo_x() + (e.x - self._dx)}+{self.winfo_y() + (e.y - self._dy)}")

    def _drag_end(self, _e):
        self._save_pos()
        self.app.pause_updates = False

    def _calc_height(self, rows):
        n = max(1, rows)
        h = HEADER_H + n * ROW_H
        if self.inc_expanded:
            # label row (16px) + data rows
            h += BOTTOM_PAD // 2 + 16 + n * self.INC_ROW_H
        h += 16   # space for the bottom toggle button
        return h

    def _build_static_ui(self):
        for child in self.winfo_children():
            child.destroy()
        self.cv = tk.Canvas(self, width=self._cur_w, height=self._cur_h, bg=TRANSPARENT, highlightthickness=0)
        self.cv.pack()
        self.cv.bind("<Motion>", self._on_motion)
        self.cv.bind("<Leave>", lambda _e: self._on_leave())
        self.cv.bind("<Button-1>", self._on_left_click)
        self.cv.bind("<Button-3>", self._on_right_click)

    def _clear_canvas(self):
        self.cv.delete("all")
        self.row_info = {}
        self.ui_items = {}
        self.header_click_boxes = {}

    def _layout_columns(self):
        left = 10
        right = self._cur_w - 10
        usable = right - left
        name_col = int(usable * 0.40)
        dps_col = int(usable * 0.20)
        tdps_col = int(usable * 0.20)
        return {
            "name_l": left,
            "name_r": left + name_col,
            "dps_r": left + name_col + dps_col,
            "tdps_r": left + name_col + dps_col + tdps_col,
            "tdam_r": right,
        }

    def _measure_text(self, text, font):
        temp = self.cv.create_text(-9999, -9999, text=text, font=font, anchor="nw")
        bbox = self.cv.bbox(temp)
        self.cv.delete(temp)
        return 0 if not bbox else bbox[2] - bbox[0]

    def render_all(self, force_layout=False):
        chars = self.app.sorted_characters()
        n = max(1, len(chars))
        target_h = self._calc_height(len(chars))
        if force_layout or target_h != self._cur_h:
            self._cur_h = target_h
            self.geometry(f"{self._cur_w}x{self._cur_h}+{self.winfo_x()}+{self.winfo_y()}")
            self.cv.config(width=self._cur_w, height=self._cur_h)

        self._clear_canvas()
        self.cv.create_polygon(rrect(1, 1, self._cur_w - 1, self._cur_h - 1, RADIUS), smooth=True, fill=C_BG, outline=C_BORDER, width=1.4)
        cols = self._layout_columns()
        self.cols = cols

        # ── Header bar ──────────────────────────────────────────────
        self.cv.create_text(10, 12, anchor="w", text=APP_TITLE, fill=C_HEADER, font=("Consolas", 11, "bold"))
        self.cv.create_text(83, 13, anchor="w", text=APP_SUBTITLE, fill=C_MUTED, font=("Consolas", 7))

        x_quit  = self._cur_w - 10
        x_min   = self._cur_w - 28
        x_refresh = self._cur_w - 46
        x_graph = self._cur_w - 64
        x_show  = self._cur_w - 82

        self.ui_items["quit"]     = self.cv.create_text(x_quit,    12, anchor="e", text="✕", fill=C_ICON, font=("Segoe UI Symbol", 10))
        self.ui_items["min"]      = self.cv.create_text(x_min,     12, anchor="e", text="—", fill=C_ICON, font=("Segoe UI Symbol", 11))
        self.ui_items["refresh"]  = self.cv.create_text(x_refresh, 12, anchor="e", text="↻", fill=C_ICON, font=("Segoe UI Symbol", 13))
        self.ui_items["graph"]    = self.cv.create_text(x_graph,   12, anchor="e", text="📈", fill=C_ICON, font=("Segoe UI Symbol", 10))
        self.ui_items["show_all"] = self.cv.create_text(x_show,    12, anchor="e", text="👁", fill=C_ICON, font=("Segoe UI Symbol", 13))

        self.header_click_boxes["quit"]     = (x_quit - 20,    0, x_quit + 4,    24)
        self.header_click_boxes["min"]      = (x_min - 20,     0, x_min + 4,     24)
        self.header_click_boxes["refresh"]  = (x_refresh - 20, 0, x_refresh + 4, 24)
        self.header_click_boxes["graph"]    = (x_graph - 20,   0, x_graph + 4,   24)
        self.header_click_boxes["show_all"] = (x_show - 20,    0, x_show + 4,    24)

        header_y = HEADER_H - 15
        self.cv.create_text(cols["name_l"],    header_y, anchor="w", text="Name",  fill=C_HEADER, font=("Consolas", 10, "bold"))
        self.cv.create_text(cols["dps_r"]  - 4, header_y, anchor="e", text="DPS",   fill=C_HEADER, font=("Consolas", 10, "bold"))
        self.cv.create_text(cols["tdps_r"] - 4, header_y, anchor="e", text="T.DPS", fill=C_HEADER, font=("Consolas", 10, "bold"))
        self.cv.create_text(cols["tdam_r"] - 4, header_y, anchor="e", text="T.Dam", fill=C_HEADER, font=("Consolas", 10, "bold"))

        if not chars:
            self.cv.create_text(self._cur_w // 2, HEADER_H + ROW_H // 2,
                                anchor="center", text="Waiting for EVE logs...",
                                fill=C_MUTED, font=("Consolas", 10))
            return

        # ── Outgoing DPS rows ────────────────────────────────────────
        for idx, char in enumerate(chars):
            y1 = HEADER_H + idx * ROW_H
            y2 = y1 + ROW_H
            yc = (y1 + y2) // 2
            hover = (char == self.hover_char)
            row_fill = "#15304a" if hover else C_BG

            self.cv.create_rectangle(6, y1, self._cur_w - 6, y2, outline=C_BORDER, width=0.6, fill=row_fill)

            status = self.app.get_status(char)
            cur_fill = C_DPS_ON if (status["in_combat"] and status["dps"] > 0) else (C_DPS_OFF if status["dps"] > 0 else C_IDLE)

            display = self.app.display_name(char)
            icon_reserve = 22 if hover else 1
            max_name_w = (cols["name_r"] - cols["name_l"]) - icon_reserve
            fitted = fit_text_binary(display, max_name_w, lambda s: self._measure_text(s, self.font_name))

            self.cv.create_text(cols["name_l"],    yc, anchor="w", text=fitted,                   fill=C_TEXT,    font=self.font_name)
            self.cv.create_text(cols["dps_r"]  - 4, yc, anchor="e", text=fmt_num(status["dps"]),   fill=cur_fill,  font=self.font_num_b)
            self.cv.create_text(cols["tdps_r"] - 4, yc, anchor="e", text=fmt_num(status["tdps"]),  fill=C_TEXT,    font=self.font_num)
            self.cv.create_text(cols["tdam_r"] - 4, yc, anchor="e", text=fmt_int(status["tdam"]),  fill=C_TEXT,    font=self.font_num)

            edit_box = None
            hide_box = None
            if hover:
                edit_x = cols["name_r"] - 20
                hide_x = cols["name_r"] + 2
                self.cv.create_text(edit_x, yc, anchor="center", text="✎", fill=C_ICON_H, font=("Segoe UI Symbol", 10, "bold"))
                self.cv.create_text(hide_x, yc, anchor="center", text="⊘", fill=C_HIDE,   font=("Segoe UI Symbol", 12, "bold"))
                edit_box = (edit_x - 9, y1 + 2, edit_x + 9, y2 - 2)
                hide_box = (hide_x - 13, y1 + 1, hide_x + 13, y2 - 1)

            self.row_info[char] = {
                "row_box":  (6, y1, self._cur_w - 6, y2),
                "name_box": (cols["name_l"], y1, cols["name_r"], y2),
                "edit_box": edit_box,
                "hide_box": hide_box,
                "display":  display,
                "fitted":   fitted,
                "y1":       y1,
                "edit_window_id": None,
            }

        if self.inline_editor_char and self.inline_editor_char in self.row_info:
            self._place_inline_editor(recreate_window=True)

        # ── Incoming panel (expanded rows, no header bar) ─────────────
        C_INC       = "#e06060"
        C_INC_MUTED = "#a04040"

        if self.inc_expanded and chars:
            inc_top = HEADER_H + n * ROW_H + BOTTOM_PAD // 2
            # Sub-headers
            self.cv.create_text(cols["name_l"],    inc_top + 8, anchor="w", text="Incoming",      fill=C_INC_MUTED, font=("Consolas", 8, "bold"))
            self.cv.create_text(cols["dps_r"]  - 4, inc_top + 8, anchor="e", text="Inc.DPS",   fill=C_INC_MUTED, font=("Consolas", 8))
            self.cv.create_text(cols["tdps_r"] - 4, inc_top + 8, anchor="e", text="Inc.T.DPS", fill=C_INC_MUTED, font=("Consolas", 8))
            self.cv.create_text(cols["tdam_r"] - 4, inc_top + 8, anchor="e", text="Inc.T.Dam", fill=C_INC_MUTED, font=("Consolas", 8))

            for idx, char in enumerate(chars):
                ry1 = inc_top + 16 + idx * self.INC_ROW_H
                ry2 = ry1 + self.INC_ROW_H
                ryc = (ry1 + ry2) // 2
                self.cv.create_rectangle(6, ry1, self._cur_w - 6, ry2,
                                         outline="#2a1010", width=0.6, fill="#130a0a")
                status     = self.app.get_status(char)
                display    = self.app.display_name(char)
                max_name_w = cols["name_r"] - cols["name_l"] - 1
                fitted     = fit_text_binary(display, max_name_w,
                                             lambda s: self._measure_text(s, self.font_name))
                self.cv.create_text(cols["name_l"],     ryc, anchor="w", text=fitted,                        fill="#c08080", font=self.font_name)
                self.cv.create_text(cols["dps_r"]  - 4, ryc, anchor="e", text=fmt_num(status["inc_dps"]),    fill=C_INC,    font=self.font_num_b)
                self.cv.create_text(cols["tdps_r"] - 4, ryc, anchor="e", text=fmt_num(status["inc_tdps"]),   fill=C_INC,    font=self.font_num)
                self.cv.create_text(cols["tdam_r"] - 4, ryc, anchor="e", text=fmt_int(status["inc_tdam"]),   fill=C_INC,    font=self.font_num)

        # ── Bottom toggle button ──────────────────────────────────────
        btn_w, btn_h = 60, 12
        btn_x = (self._cur_w - btn_w) // 2
        btn_y = self._cur_h - btn_h - 2
        sym = "▲ Inc" if self.inc_expanded else "▼ Inc"
        self.cv.create_rectangle(btn_x, btn_y, btn_x + btn_w, btn_y + btn_h,
                                 fill="#1a1010", outline="#5a2020", width=0.8)
        self.cv.create_text(btn_x + btn_w // 2, btn_y + btn_h // 2, anchor="center",
                            text=sym, fill="#a04040", font=("Consolas", 7))
        self.header_click_boxes["inc_toggle"] = (btn_x, btn_y, btn_x + btn_w, btn_y + btn_h)

    def _update_loop(self):
        if not self.app.pause_updates:
            self.render_all(force_layout=False)
        self.after(UI_UPDATE_MS, self._update_loop)

    def _on_right_click(self, e):
        menu = tk.Menu(self, tearoff=0, bg="#0f1923", fg="#90caf9", activebackground="#1e3a5f", activeforeground="#fff", font=("Consolas", 9), bd=0, relief="flat")
        menu.add_command(label="Graph", command=self.app.open_graph)
        menu.add_command(label="Minimize", command=self.app.minimize_all)
        menu.add_command(label="Reset", command=self.app.reset_everything)
        menu.add_separator()
        menu.add_command(label="Exit", command=self.app.quit_all)
        menu.tk_popup(e.x_root, e.y_root)

    def _hit_test_header(self, x, y):
        for key, box in self.header_click_boxes.items():
            x1, y1, x2, y2 = box
            if x1 <= x <= x2 and y1 <= y <= y2:
                return key
        return None

    def _hit_test_char(self, x, y):
        for char, info in self.row_info.items():
            x1, y1, x2, y2 = info["row_box"]
            if x1 <= x <= x2 and y1 <= y <= y2:
                return char
        return None

    def _on_left_click(self, e):
        header = self._hit_test_header(e.x, e.y)
        if header == "quit":
            self.app.quit_all()
            return
        if header == "min":
            self.app.minimize_all()
            return
        if header == "refresh":
            self.app.reset_dps_only()
            return
        if header == "show_all":
            self.app.show_all_hidden()
            return
        if header == "graph":
            self.app.open_graph()
            return
        if header == "inc_toggle":
            self.inc_expanded = not self.inc_expanded
            self.app.cfg["inc_expanded"] = self.inc_expanded
            save_config(self.app.cfg)
            self.render_all(force_layout=True)
            return

        char = self._hit_test_char(e.x, e.y)
        if not char:
            return

        info = self.row_info.get(char)
        if not info:
            return

        if info["edit_box"]:
            x1, y1, x2, y2 = info["edit_box"]
            if x1 <= e.x <= x2 and y1 <= e.y <= y2:
                self._start_inline_edit(char)
                return

        if info["hide_box"]:
            x1, y1, x2, y2 = info["hide_box"]
            if x1 <= e.x <= x2 and y1 <= e.y <= y2:
                self.app.hide_character(char)
                return

    def _on_motion(self, e):
        if self.inline_editor_char:
            return

        header = self._hit_test_header(e.x, e.y)
        if header:
            text = {
                "show_all": "Show all",
                "graph": "Recent battle",
                "refresh": "Refresh",
                "min": "Minimize",
                "quit": "Quit",
                "inc_toggle": "Toggle incoming",
            }.get(header)
            self.hover_char = None
            if text:
                self.tooltip.show(e.x_root, e.y_root, text)
            else:
                self.tooltip.hide()
            self.render_all(force_layout=False)
            return

        old_hover = self.hover_char
        self.hover_char = self._hit_test_char(e.x, e.y)
        if old_hover != self.hover_char:
            self.tooltip.hide()
            self.render_all(force_layout=False)

        if not self.hover_char:
            self.tooltip.hide()
            return

        info = self.row_info.get(self.hover_char)
        if not info:
            self.tooltip.hide()
            return

        if info.get("edit_box"):
            x1, y1, x2, y2 = info["edit_box"]
            if x1 <= e.x <= x2 and y1 <= e.y <= y2:
                self.tooltip.show(e.x_root, e.y_root, "Set name")
                return

        if info.get("hide_box"):
            x1, y1, x2, y2 = info["hide_box"]
            if x1 <= e.x <= x2 and y1 <= e.y <= y2:
                self.tooltip.show(e.x_root, e.y_root, "Hide")
                return

        x1, y1, x2, y2 = info["name_box"]
        if x1 <= e.x <= x2 and y1 <= e.y <= y2 and info["fitted"] != info["display"]:
            self.tooltip.show(e.x_root, e.y_root, info["display"])
        else:
            self.tooltip.hide()

    def _on_leave(self):
        if self.inline_editor_char:
            return
        changed = self.hover_char is not None
        self.hover_char = None
        self.tooltip.hide()
        if changed:
            self.render_all(force_layout=False)

    def _start_inline_edit(self, char):
        self.app.pause_updates = True
        self.hover_char = char
        self.render_all(force_layout=False)

        if self.inline_editor and self.inline_editor.winfo_exists():
            try:
                self.inline_editor.destroy()
            except Exception:
                pass

        self.inline_editor_char = char
        entry = tk.Entry(self.cv, bg=C_EDIT_BG, fg=C_EDIT_FG, insertbackground="#ffffff", relief="flat", font=self.font_name)
        entry.insert(0, self.app.cfg.get("aliases", {}).get(char, ""))
        entry.focus_set()
        entry.select_range(0, tk.END)
        entry.bind("<Return>", lambda _e: self._commit_inline_edit())
        entry.bind("<Escape>", lambda _e: self._cancel_inline_edit())
        entry.bind("<FocusOut>", lambda _e: self._commit_inline_edit())
        self.inline_editor = entry
        self._place_inline_editor(recreate_window=True)

    def _place_inline_editor(self, recreate_window=False):
        if not self.inline_editor_char or not self.inline_editor or not self.inline_editor.winfo_exists():
            return

        info = self.row_info.get(self.inline_editor_char)
        if not info:
            return

        x = self.cols["name_l"]
        y = info["y1"] + 3
        w = (self.cols["name_r"] - self.cols["name_l"]) - 1
        h = ROW_H - 6

        if recreate_window or info.get("edit_window_id") is None:
            info["edit_window_id"] = self.cv.create_window(x, y, anchor="nw", window=self.inline_editor, width=w, height=h)
        else:
            self.cv.coords(info["edit_window_id"], x, y)
            self.cv.itemconfigure(info["edit_window_id"], width=w, height=h)

    def _commit_inline_edit(self):
        if not self.inline_editor_char or not self.inline_editor or not self.inline_editor.winfo_exists():
            self.app.pause_updates = False
            return

        char = self.inline_editor_char
        text = self.inline_editor.get().strip()
        self.app.set_alias(char, text)

        try:
            self.inline_editor.destroy()
        except Exception:
            pass

        self.inline_editor = None
        self.inline_editor_char = None
        self.app.pause_updates = False
        self.render_all(force_layout=False)

    def _cancel_inline_edit(self):
        if self.inline_editor and self.inline_editor.winfo_exists():
            try:
                self.inline_editor.destroy()
            except Exception:
                pass
        self.inline_editor = None
        self.inline_editor_char = None
        self.app.pause_updates = False
        self.render_all(force_layout=False)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.withdraw()
        self.cfg = load_config()
        self.hidden_chars = set(self.cfg.get("hidden_chars", []))
        self.engines = {}
        self.detected_chars = set()
        self.pause_updates = False
        self.graph_win = None
        self.window = EVEUnifiedWindow(self, self)
        self.protocol("WM_DELETE_WINDOW", self.quit_all)
        self._scan_running_windows()

    def sorted_characters(self):
        visible = [c for c in self.detected_chars if c not in self.hidden_chars]
        return sorted(visible, key=natural_sort_key)

    def display_name(self, char):
        alias = self.cfg.get("aliases", {}).get(char, "").strip()
        return alias if alias else char

    def get_or_create_engine(self, char):
        eng = self.engines.get(char)
        if eng is None:
            eng = DPSEngine(char)
            eng.start()
            self.engines[char] = eng
        return eng

    def get_status(self, char):
        return self.get_or_create_engine(char).get_status()

    def set_alias(self, char, alias_text):
        aliases = self.cfg.setdefault("aliases", {})
        if alias_text:
            aliases[char] = alias_text
        else:
            aliases.pop(char, None)
        save_config(self.cfg)

    def _save_hidden(self):
        self.cfg["hidden_chars"] = sorted(self.hidden_chars, key=natural_sort_key)
        save_config(self.cfg)

    def hide_character(self, char):
        self.hidden_chars.add(char)
        self._save_hidden()
        self.window.hover_char = None
        self.window.tooltip.hide()
        self.window.render_all(force_layout=True)

    def show_all_hidden(self):
        self.hidden_chars.clear()
        self._save_hidden()
        self.window.render_all(force_layout=True)

    def reset_dps_only(self):
        for eng in list(self.engines.values()):
            try:
                eng.hard_reset()
            except Exception:
                pass
        if self.graph_win and self.graph_win.winfo_exists():
            self.graph_win.render()
        self.window.render_all(force_layout=False)

    def reset_everything(self):
        ok = messagebox.askyesno("Reset", "Reset everything to first-run state?")
        if not ok:
            return

        for eng in list(self.engines.values()):
            try:
                eng.stop()
            except Exception:
                pass

        self.engines.clear()
        self.detected_chars.clear()
        self.hidden_chars.clear()
        self.cfg = default_config()
        save_config(self.cfg)

        self.window._cur_w = WINDOW_DEFAULT_W
        self.window.geometry(f"{WINDOW_DEFAULT_W}x220+120+120")
        self.window.attributes("-alpha", 0.92)
        self.window._save_pos()
        self.window.hover_char = None
        self.window.tooltip.hide()
        self.window.render_all(force_layout=True)

        if self.graph_win and self.graph_win.winfo_exists():
            self.graph_win.destroy()
            self.graph_win = None

    def minimize_all(self):
        try:
            self.window._save_pos()
        except Exception:
            pass
        try:
            self.window.overrideredirect(False)
        except Exception:
            pass
        try:
            self.window.iconify()
        except Exception:
            pass

    def open_graph(self):
        if self.graph_win and self.graph_win.winfo_exists():
            self.graph_win.lift()
            self.graph_win.focus_force()
            self.graph_win.render()
            return
        self.graph_win = GraphWindow(self, self)

    def _scan_running_windows(self):
        if not self.pause_updates:
            current = list_running_eve_characters_from_windows()
            self.detected_chars = set(current)

            for char in current:
                self.get_or_create_engine(char)

            for char in list(self.engines.keys()):
                if char not in current:
                    eng = self.engines.pop(char, None)
                    if eng:
                        eng.stop()

            self.window.render_all(force_layout=True)

        self.after(WINDOW_SCAN_INTERVAL_MS, self._scan_running_windows)

    def quit_all(self):
        if self.graph_win and self.graph_win.winfo_exists():
            try:
                self.cfg["graph_window"] = {
                    "x": self.graph_win.winfo_x(), "y": self.graph_win.winfo_y(),
                    "w": self.graph_win._win_w,    "h": self.graph_win._win_h,
                }
            except Exception:
                pass
            try:
                self.graph_win.destroy()
            except Exception:
                pass

        for eng in list(self.engines.values()):
            try:
                eng.stop()
            except Exception:
                pass

        self.engines.clear()

        try:
            self.window._save_pos()
        except Exception:
            pass

        self.destroy()


if __name__ == "__main__":
    App().mainloop()
