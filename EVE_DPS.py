import tkinter as tk
from tkinter import messagebox
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
APP_SUBTITLE = "v.1.1 by Leffe Brown"

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

GRAPH_COLORS = [
    "#4fc3f7", "#81c784", "#ffd54f", "#ff8a65", "#ba68c8",
    "#f06292", "#4dd0e1", "#aed581", "#ffb74d", "#9575cd",
]

WINDOW_SCAN_INTERVAL_MS = 1500
UI_UPDATE_MS = 400
SESSION_TIMEOUT = 30.0
TITLE_PREFIX = "EVE - "
GRAPH_BUCKET_SECONDS = 1.0

# 내가 준 피해(청록색)
OUTGOING_RE = re.compile(
    r"\[\s*(\d{4}\.\d{2}\.\d{2}\s+\d{2}:\d{2}:\d{2})\s*\]"
    r".*?\(combat\).*?"
    r"<color=0xff00ffff><b>(\d+)</b>",
    re.IGNORECASE,
)

# 내가 받은 피해(빨간색)
INCOMING_RE = re.compile(
    r"\[\s*(\d{4}\.\d{2}\.\d{2}\s+\d{2}:\d{2}:\d{2})\s*\]"
    r".*?\(combat\).*?"
    r"<color=0xffcc0000><b>(\d+)</b>",
    re.IGNORECASE,
)

LISTENER_RE_KO = re.compile(r"청취자\s*:\s*(.+)")
LISTENER_RE_EN = re.compile(r"Listener\s*:\s*(.+)")


def default_config():
    return {
        "window": {"x": 120, "y": 120, "w": WINDOW_DEFAULT_W, "alpha": 92},
        "aliases": {},
        "hidden_chars": [],
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
        except Exception:
            pass
    return cfg


def save_config(cfg):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def parse_ts(s):
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

        # DPS 계산용: 내가 준 피해만 사용
        self.sessions = []
        self.cur = None
        self.last_outgoing_mono = None

        # 전투중 판정용: 내가 주거나 받는 피해 모두 사용
        self.last_combat_mono = None

        # 그래프용
        self.last_session_points = []
        self.last_session_start_ts = None
        self.last_session_end_ts = None

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
            self.sessions = []
            self.cur = None
            self.last_outgoing_mono = None
            self.last_combat_mono = None
            self.last_session_points = []
            self.last_session_start_ts = None
            self.last_session_end_ts = None
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

        for m in OUTGOING_RE.finditer(chunk):
            ts = parse_ts(m.group(1))
            if ts:
                try:
                    events.append((ts, "out", int(m.group(2))))
                except Exception:
                    pass

        for m in INCOMING_RE.finditer(chunk):
            ts = parse_ts(m.group(1))
            if ts:
                events.append((ts, "in", 0))

        events.sort(key=lambda x: (x[0], 0 if x[1] == "out" else 1))

        now_mono = time.monotonic()
        with self._lock:
            for ts, kind, value in events:
                self.last_combat_mono = now_mono
                if kind == "out":
                    self._add_outgoing(ts, value, now_mono)
            self._check_end(now_mono)

    def _build_graph_points(self, segment_hits):
        if not segment_hits:
            return []
        buckets = {}
        for active_sec, dmg in segment_hits:
            idx = int(active_sec // GRAPH_BUCKET_SECONDS)
            buckets[idx] = buckets.get(idx, 0) + dmg

        points = []
        for idx in sorted(buckets.keys()):
            points.append((idx * GRAPH_BUCKET_SECONDS, buckets[idx] / GRAPH_BUCKET_SECONDS))
        return points

    def _finalize_current_session(self):
        if not self.cur:
            return
        duration = max(0.0, self.cur["active_elapsed"])
        self.sessions.append({
            "total_dmg": self.cur["total_dmg"],
            "duration": duration,
        })
        self.last_session_points = self._build_graph_points(self.cur["segment_hits"])
        self.last_session_start_ts = self.cur["start_ts"]
        self.last_session_end_ts = self.cur["last_ts"]
        self.cur = None
        self.last_outgoing_mono = None

    def _add_outgoing(self, ts, dmg, now_mono):
        if self.cur is None:
            self.cur = {
                "start_ts": ts,
                "last_ts": ts,
                "total_dmg": dmg,
                "active_elapsed": 0.0,
                "segment_hits": [(0.0, dmg)],
            }
            self.last_outgoing_mono = now_mono
            return

        gap = (ts - self.cur["last_ts"]).total_seconds()
        if gap <= SESSION_TIMEOUT:
            active_elapsed = self.cur["active_elapsed"] + max(0.0, gap)
            self.cur["last_ts"] = ts
            self.cur["total_dmg"] += dmg
            self.cur["active_elapsed"] = active_elapsed
            self.cur["segment_hits"].append((active_elapsed, dmg))
            self.last_outgoing_mono = now_mono
        else:
            self._finalize_current_session()
            self.cur = {
                "start_ts": ts,
                "last_ts": ts,
                "total_dmg": dmg,
                "active_elapsed": 0.0,
                "segment_hits": [(0.0, dmg)],
            }
            self.last_outgoing_mono = now_mono

    def _check_end(self, now_mono):
        if self.cur is not None and self.last_outgoing_mono is not None:
            if (now_mono - self.last_outgoing_mono) >= SESSION_TIMEOUT:
                self._finalize_current_session()

        if self.last_combat_mono is not None:
            if (now_mono - self.last_combat_mono) >= SESSION_TIMEOUT:
                self.last_combat_mono = None

    def get_status(self):
        with self._lock:
            all_sessions = list(self.sessions)
            cur = dict(self.cur) if self.cur else None
            combat_active = self.last_combat_mono is not None

        total_damage = 0
        total_secs = 0.0
        for s in all_sessions:
            total_damage += s["total_dmg"]
            if s["duration"] > 0:
                total_secs += s["duration"]

        dps_damage = 0
        dps_secs = 0.0
        if cur:
            dps_damage = cur["total_dmg"]
            dps_secs = max(0.0, cur["active_elapsed"])
            total_damage += dps_damage
            if dps_secs > 0:
                total_secs += dps_secs

        dps = (dps_damage / dps_secs) if dps_secs > 0 else 0.0
        tdps = (total_damage / total_secs) if total_secs > 0 else 0.0

        return {
            "dps": dps,
            "tdps": tdps,
            "tdam": total_damage,
            "in_combat": combat_active,
        }

    def get_graph_payload(self):
        with self._lock:
            if self.cur and self.cur.get("segment_hits"):
                return {
                    "points": self._build_graph_points(self.cur["segment_hits"]),
                    "start_ts": self.cur["start_ts"],
                    "end_ts": self.cur["last_ts"],
                }
            return {
                "points": list(self.last_session_points),
                "start_ts": self.last_session_start_ts,
                "end_ts": self.last_session_end_ts,
            }


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
        tk.Label(
            self.win,
            text=text,
            bg=C_TOOLTIP_BG,
            fg=C_TOOLTIP_FG,
            font=("Consolas", 9),
            padx=7,
            pady=4,
        ).pack()
        self.win.geometry(f"+{x + 12}+{y + 12}")

    def hide(self):
        if self.win and self.win.winfo_exists():
            try:
                self.win.destroy()
            except Exception:
                pass
        self.win = None


class GraphWindow(tk.Toplevel):
    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self.title("Recent DPS Graph")
        self.geometry("760x420+180+180")
        self.configure(bg=C_BG)
        self.attributes("-topmost", True)

        self.canvas = tk.Canvas(self, bg=C_BG, highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)

        self.bind("<Configure>", lambda _e: self.render())
        self.after(500, self._update_loop)

    def _collect_series(self):
        series = []
        chars = self.app.sorted_characters()
        for idx, char in enumerate(chars):
            eng = self.app.engines.get(char)
            if not eng:
                continue
            payload = eng.get_graph_payload()
            start_ts = payload.get("start_ts")
            end_ts = payload.get("end_ts")
            points = payload.get("points") or []
            if start_ts is None or end_ts is None:
                continue
            series.append({
                "char": char,
                "color": GRAPH_COLORS[idx % len(GRAPH_COLORS)],
                "start_ts": start_ts,
                "end_ts": end_ts,
                "points": points,
            })
        return series

    def _build_aligned_points(self, item, global_start):
        start_offset = max(0.0, (item["start_ts"] - global_start).total_seconds())
        pts = []
        for local_sec, dps in item["points"]:
            pts.append((start_offset + local_sec, dps))
        return pts

    def render(self):
        self.canvas.delete("all")
        w = max(10, self.winfo_width())
        h = max(10, self.winfo_height())
        pad_l, pad_r, pad_t, pad_b = 56, 20, 28, 58
        x1, y1, x2, y2 = pad_l, pad_t, w - pad_r, h - pad_b

        self.canvas.create_rectangle(0, 0, w, h, fill=C_BG, outline=C_BG)
        self.canvas.create_text(
            18, 14,
            anchor="w",
            text="Recent DPS Graph",
            fill=C_HEADER,
            font=("Consolas", 11, "bold"),
        )

        raw_series = self._collect_series()
        if not raw_series:
            self.canvas.create_text(
                w // 2, h // 2,
                text="No recent combat graph data",
                fill=C_MUTED,
                font=("Consolas", 12),
            )
            return

        global_start = min(x["start_ts"] for x in raw_series)
        global_end = max(x["end_ts"] for x in raw_series)
        total_secs = max(1.0, (global_end - global_start).total_seconds())

        series = []
        max_y = 1.0
        for item in raw_series:
            pts = self._build_aligned_points(item, global_start)
            if pts:
                max_y = max(max_y, max(p[1] for p in pts))
            series.append({**item, "aligned_points": pts})

        max_y *= 1.15

        self.canvas.create_line(x1, y2, x2, y2, fill=C_GRAPH_AXIS, width=1)
        self.canvas.create_line(x1, y1, x1, y2, fill=C_GRAPH_AXIS, width=1)

        for i in range(5):
            yy = y2 - ((y2 - y1) * i / 4.0)
            val = max_y * i / 4.0
            self.canvas.create_line(x1, yy, x2, yy, fill="#13263a", width=1)
            self.canvas.create_text(
                x1 - 8, yy,
                anchor="e",
                text=f"{int(val)}",
                fill=C_GRAPH_TEXT,
                font=("Consolas", 9),
            )

        x_ticks = min(7, max(2, int(total_secs) + 1))
        for i in range(x_ticks):
            ratio = i / (x_ticks - 1 if x_ticks > 1 else 1)
            xx = x1 + ((x2 - x1) * ratio)
            val = total_secs * ratio
            self.canvas.create_line(xx, y2, xx, y2 + 4, fill=C_GRAPH_AXIS, width=1)
            self.canvas.create_text(
                xx, y2 + 14,
                anchor="n",
                text=f"{val:.0f}s",
                fill=C_GRAPH_TEXT,
                font=("Consolas", 9),
            )

        self.canvas.create_text(
            18, (y1 + y2) / 2,
            text="DPS",
            angle=90,
            fill=C_MUTED,
            font=("Consolas", 9),
        )

        for item in series:
            coords = []
            for sec, dps in item["aligned_points"]:
                px = x1 + ((sec / total_secs) * (x2 - x1))
                py = y2 - ((dps / max_y) * (y2 - y1))
                coords.extend([px, py])

            if len(coords) >= 4:
                self.canvas.create_line(*coords, fill=item["color"], width=2, smooth=True)
            elif len(coords) == 2:
                px, py = coords
                self.canvas.create_oval(px - 2, py - 2, px + 2, py + 2, fill=item["color"], outline=item["color"])

        legend_x = x1
        legend_y = h - 28
        per_row = max(1, (w - 100) // 120)
        for idx, item in enumerate(series):
            lx = legend_x + (idx % per_row) * 120
            ly = legend_y - (idx // per_row) * 18
            self.canvas.create_line(lx, ly, lx + 18, ly, fill=item["color"], width=3)
            self.canvas.create_text(
                lx + 24, ly,
                anchor="w",
                text=item["char"],
                fill=C_GRAPH_TEXT,
                font=("Consolas", 9),
            )

    def _update_loop(self):
        if self.winfo_exists():
            self.render()
            self.after(800, self._update_loop)


class EVEUnifiedWindow(tk.Toplevel):
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
        return HEADER_H + max(1, rows) * ROW_H + BOTTOM_PAD

    def _build_static_ui(self):
        for child in self.winfo_children():
            child.destroy()
        self.cv = tk.Canvas(
            self,
            width=self._cur_w,
            height=self._cur_h,
            bg=TRANSPARENT,
            highlightthickness=0,
        )
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
        target_h = self._calc_height(len(chars))
        if force_layout or target_h != self._cur_h:
            self._cur_h = target_h
            self.geometry(f"{self._cur_w}x{self._cur_h}+{self.winfo_x()}+{self.winfo_y()}")
            self.cv.config(width=self._cur_w, height=self._cur_h)

        self._clear_canvas()
        self.cv.create_polygon(
            rrect(1, 1, self._cur_w - 1, self._cur_h - 1, RADIUS),
            smooth=True,
            fill=C_BG,
            outline=C_BORDER,
            width=1.4,
        )
        cols = self._layout_columns()
        self.cols = cols

        self.cv.create_text(10, 12, anchor="w", text=APP_TITLE, fill=C_HEADER, font=("Consolas", 11, "bold"))
        self.cv.create_text(83, 13, anchor="w", text=APP_SUBTITLE, fill=C_MUTED, font=("Consolas", 7))

        x_quit = self._cur_w - 10
        x_refresh = self._cur_w - 28
        x_graph = self._cur_w - 46
        x_show = self._cur_w - 64

        self.ui_items["quit"] = self.cv.create_text(
            x_quit, 12, anchor="e", text="✕", fill=C_ICON, font=("Segoe UI Symbol", 10)
        )
        self.ui_items["refresh"] = self.cv.create_text(
            x_refresh, 12, anchor="e", text="↻", fill=C_ICON, font=("Segoe UI Symbol", 13)
        )
        self.ui_items["graph"] = self.cv.create_text(
            x_graph, 12, anchor="e", text="📈", fill=C_ICON, font=("Segoe UI Symbol", 10)
        )
        self.ui_items["show_all"] = self.cv.create_text(
            x_show, 12, anchor="e", text="👁", fill=C_ICON, font=("Segoe UI Symbol", 13)
        )

        self.header_click_boxes["quit"] = (x_quit - 20, 0, x_quit + 4, 24)
        self.header_click_boxes["refresh"] = (x_refresh - 20, 0, x_refresh + 4, 24)
        self.header_click_boxes["graph"] = (x_graph - 20, 0, x_graph + 4, 24)
        self.header_click_boxes["show_all"] = (x_show - 20, 0, x_show + 4, 24)

        header_y = HEADER_H - 15
        self.cv.create_text(cols["name_l"], header_y, anchor="w", text="Name", fill=C_HEADER, font=("Consolas", 10, "bold"))
        self.cv.create_text(cols["dps_r"] - 4, header_y, anchor="e", text="DPS", fill=C_HEADER, font=("Consolas", 10, "bold"))
        self.cv.create_text(cols["tdps_r"] - 4, header_y, anchor="e", text="T.DPS", fill=C_HEADER, font=("Consolas", 10, "bold"))
        self.cv.create_text(cols["tdam_r"] - 4, header_y, anchor="e", text="T.Dam", fill=C_HEADER, font=("Consolas", 10, "bold"))

        if not chars:
            self.cv.create_text(
                self._cur_w // 2,
                HEADER_H + ROW_H // 2,
                anchor="center",
                text="Waiting for EVE logs...",
                fill=C_MUTED,
                font=("Consolas", 10),
            )
            return

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

            self.cv.create_text(cols["name_l"], yc, anchor="w", text=fitted, fill=C_TEXT, font=self.font_name)
            self.cv.create_text(cols["dps_r"] - 4, yc, anchor="e", text=fmt_num(status["dps"]), fill=cur_fill, font=self.font_num_b)
            self.cv.create_text(cols["tdps_r"] - 4, yc, anchor="e", text=fmt_num(status["tdps"]), fill=C_TEXT, font=self.font_num)
            self.cv.create_text(cols["tdam_r"] - 4, yc, anchor="e", text=fmt_int(status["tdam"]), fill=C_TEXT, font=self.font_num)

            edit_box = None
            hide_box = None
            if hover:
                edit_x = cols["name_r"] - 20
                hide_x = cols["name_r"] + 2
                self.cv.create_text(edit_x, yc, anchor="center", text="✎", fill=C_ICON_H, font=("Segoe UI Symbol", 10, "bold"))
                self.cv.create_text(hide_x, yc, anchor="center", text="⊘", fill=C_HIDE, font=("Segoe UI Symbol", 12, "bold"))
                edit_box = (edit_x - 9, y1 + 2, edit_x + 9, y2 - 2)
                hide_box = (hide_x - 13, y1 + 1, hide_x + 13, y2 - 1)

            self.row_info[char] = {
                "row_box": (6, y1, self._cur_w - 6, y2),
                "name_box": (cols["name_l"], y1, cols["name_r"], y2),
                "edit_box": edit_box,
                "hide_box": hide_box,
                "display": display,
                "fitted": fitted,
                "y1": y1,
                "edit_window_id": None,
            }

        if self.inline_editor_char and self.inline_editor_char in self.row_info:
            self._place_inline_editor(recreate_window=True)

    def _update_loop(self):
        if not self.app.pause_updates:
            self.render_all(force_layout=False)
        self.after(UI_UPDATE_MS, self._update_loop)

    def _on_right_click(self, e):
        menu = tk.Menu(
            self,
            tearoff=0,
            bg="#0f1923",
            fg="#90caf9",
            activebackground="#1e3a5f",
            activeforeground="#fff",
            font=("Consolas", 9),
            bd=0,
            relief="flat",
        )
        menu.add_command(label="Graph", command=self.app.open_graph)
        menu.add_command(label="Reset", command=self.app.reset_everything)
        menu.add_command(label="Minimize", command=self.app.minimize_all)
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
        if header == "refresh":
            self.app.reset_dps_only()
            return
        if header == "show_all":
            self.app.show_all_hidden()
            return
        if header == "graph":
            self.app.open_graph()
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
                "graph": "Recent graph",
                "refresh": "Refresh",
                "quit": "Quit",
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
        entry = tk.Entry(
            self.cv,
            bg=C_EDIT_BG,
            fg=C_EDIT_FG,
            insertbackground="#ffffff",
            relief="flat",
            font=self.font_name,
        )
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
            info["edit_window_id"] = self.cv.create_window(
                x, y, anchor="nw", window=self.inline_editor, width=w, height=h
            )
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
            self.graph_win.render()

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