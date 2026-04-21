import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import json
import os
import re
import time
import threading
import glob
import math
import struct
import wave
import tempfile
from datetime import datetime
from collections import deque

try:
    import win32gui
except Exception:
    win32gui = None

try:
    import winsound as _winsound
    WINSOUND_OK = True
except Exception:
    _winsound = None
    WINSOUND_OK = False


DATA_DIR = os.path.join(os.path.expanduser("~"), ".eve_multi_tools")
CONFIG_FILE = os.path.join(DATA_DIR, "eve_dps_config.json")
HISTORY_DIR = os.path.join(DATA_DIR, "eve_dps_history")
ICON_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "EVE_DPS_icon.ico")
CONTROL_FILE = os.path.join(DATA_DIR, "dps_window_control.json")
TRANSPARENT = "#010101"
APP_TITLE = "EVE DPS"
APP_SUBTITLE = "v.2.0 by Leffe Brown"

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
MAX_HISTORY_FILES = 300
HISTORY_CURRENT_LABEL = "Current Session"

# ── Incoming DPS 설정 ────────────────────────────────────────────────────────
INC_SLIDING_WINDOW_SEC = 30   # Inc.DPS(30s): 최근 30초 슬라이딩 윈도우
INC_NO_HIT_EXPIRE_SEC  = 5    # 마지막 피격 후 5초 경과 시 경보 해제

SORT_NAME_ASC   = "name_asc"
SORT_NAME_DESC  = "name_desc"
SORT_TOP_DPS    = "top_dps"
SORT_BOT_DPS    = "bot_dps"
SORT_MANUAL     = "manual"

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
TARGET_RE_UNIVERSAL = re.compile(
    r"<color=0xff00ffff><b>\d+</b>.+?<b><color=0xffffffff>(.+?)</b>",
    re.IGNORECASE,
)
_SHIP_HINT_RE  = re.compile(r'\(<localized\s+hint="([^"]+)">', re.IGNORECASE)
_PLAIN_SHIP_RE = re.compile(r'\(([^)<>]+)\)\s*$')
LISTENER_RE_KO = re.compile(r"청취자\s*:\s*(.+)")
LISTENER_RE_EN = re.compile(r"Listener\s*:\s*(.+)")


# ── WAV generation ────────────────────────────────────────────────────────────
def _generate_alert_wav(path):
    sr = 44100; dur = 0.4; n = int(sr * dur); data = []
    for i in range(n):
        t = i / sr
        env = math.sin(math.pi * t / dur) ** 0.5
        sample = env * 0.35 * (
            math.sin(2 * math.pi * 660 * t) * 0.6 +
            math.sin(2 * math.pi * 880 * t) * 0.4
        )
        v = max(-32768, min(32767, int(sample * 32767)))
        data.append(struct.pack("<h", v))
    with wave.open(path, "w") as wf:
        wf.setnchannels(1); wf.setsampwidth(2); wf.setframerate(sr)
        wf.writeframes(b"".join(data))


_DEFAULT_WAV_PATH = os.path.join(tempfile.gettempdir(), "eve_dps_alert.wav")
if not os.path.exists(_DEFAULT_WAV_PATH):
    try: _generate_alert_wav(_DEFAULT_WAV_PATH)
    except Exception: _DEFAULT_WAV_PATH = None


def default_config():
    return {
        "window":       {"x": 120, "y": 120, "w": WINDOW_DEFAULT_W, "alpha": 92},
        "aliases":      {},
        "hidden_chars": [],
        "inc_expanded": True,
        "graph_window": {"x": 180, "y": 180, "w": 1020, "h": 580},
        "sort_mode":    SORT_NAME_ASC,
        "manual_order": [],
        "alarm": {
            "enabled_visual": True,
            "enabled_sound":  False,
            "threshold":      500.0,
            "volume":         0.7,
            "repeat":         False,
            "wav_path":       "",
        },
    }


def load_config():
    cfg = default_config()
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                loaded = json.load(f)
            if isinstance(loaded, dict):
                if isinstance(loaded.get("window"), dict):        cfg["window"].update(loaded["window"])
                if isinstance(loaded.get("aliases"), dict):       cfg["aliases"] = {str(k): str(v) for k, v in loaded["aliases"].items()}
                if isinstance(loaded.get("hidden_chars"), list):  cfg["hidden_chars"] = [str(x) for x in loaded["hidden_chars"]]
                if isinstance(loaded.get("inc_expanded"), bool):  cfg["inc_expanded"] = loaded["inc_expanded"]
                if isinstance(loaded.get("graph_window"), dict):  cfg["graph_window"].update(loaded["graph_window"])
                if isinstance(loaded.get("sort_mode"), str):      cfg["sort_mode"] = loaded["sort_mode"]
                if isinstance(loaded.get("manual_order"), list):  cfg["manual_order"] = [str(x) for x in loaded["manual_order"]]
                if isinstance(loaded.get("alarm"), dict):         cfg["alarm"].update(loaded["alarm"])
        except Exception:
            pass
    return cfg


def save_config(cfg):
    try:
        os.makedirs(DATA_DIR, exist_ok=True)
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def parse_ts(s):
    try: return datetime.strptime(s.strip(), "%Y.%m.%d %H:%M:%S")
    except Exception: return None


def clamp(n, lo, hi): return max(lo, min(hi, n))


def fmt_num(v):
    if v >= 1_000_000: return f"{v/1_000_000:.1f}M"
    if v >= 1000: return f"{v/1000:.1f}k"
    return f"{v:.1f}"


def fmt_int(v):
    if v >= 1_000_000: return f"{v/1_000_000:.1f}M"
    if v >= 1000: return f"{v/1000:.1f}k"
    return str(int(v))


def fmt_dt(ts):
    if ts is None: return "-"
    return ts.strftime("%H:%M:%S")


def fmt_dt_full(ts):
    if ts is None: return "-"
    return ts.strftime("%Y-%m-%d %H:%M:%S")


def fmt_time_range(start_ts, end_ts):
    if start_ts is None or end_ts is None:
        return "-"
    if start_ts.date() == end_ts.date():
        return f"{start_ts.strftime('%H:%M')} ~ {end_ts.strftime('%H:%M')}"
    return f"{start_ts.strftime('%m-%d %H:%M')} ~ {end_ts.strftime('%m-%d %H:%M')}"


def fmt_duration(start_ts, end_ts):
    if start_ts is None or end_ts is None:
        return "-"
    seconds = max(0, int((end_ts - start_ts).total_seconds()))
    hours, rem = divmod(seconds, 3600)
    minutes, secs = divmod(rem, 60)
    if hours:
        return f"{hours}h {minutes:02d}m"
    if minutes:
        return f"{minutes}m {secs:02d}s"
    return f"{secs}s"


def apply_app_icon(win):
    if os.path.exists(ICON_FILE):
        try:
            win.iconbitmap(ICON_FILE)
        except Exception:
            pass


def rrect(x1, y1, x2, y2, r):
    return [
        x1+r,y1, x2-r,y1, x2,y1, x2,y1+r,
        x2,y2-r, x2,y2, x2-r,y2, x1+r,y2,
        x1,y2, x1,y2-r, x1,y1+r, x1,y1,
    ]


def natural_sort_key(text):
    parts = re.split(r"(\d+)", text.lower())
    return [(0, int(p)) if p.isdigit() else (1, p) for p in parts]


def _dt_to_json(value):
    if isinstance(value, datetime):
        return value.strftime("%Y.%m.%d %H:%M:%S")
    return None


def _dt_from_json(value):
    if not value:
        return None
    return parse_ts(str(value))


def _serialize_hits(hits):
    out = []
    for ts, dmg in hits or []:
        if isinstance(ts, datetime):
            out.append([_dt_to_json(ts), int(dmg)])
    return out


def _deserialize_hits(hits):
    out = []
    for item in hits or []:
        try:
            ts = _dt_from_json(item[0])
            if ts:
                out.append((ts, int(item[1])))
        except Exception:
            pass
    return out


def serialize_battle_record(rec):
    return {
        "start_ts":       _dt_to_json(rec.get("start_ts")),
        "end_ts":         _dt_to_json(rec.get("end_ts")),
        "total_dmg":      int(rec.get("total_dmg", 0)),
        "active_elapsed": float(rec.get("active_elapsed", 0.0)),
        "dps":            float(rec.get("dps", 0.0)),
        "hits":           _serialize_hits(rec.get("hits", [])),
        "first_target":   str(rec.get("first_target", "")),
        "inc_total_dmg":  int(rec.get("inc_total_dmg", 0)),
        "inc_elapsed":    float(rec.get("inc_elapsed", 0.0)),
        "inc_dps":        float(rec.get("inc_dps", 0.0)),
        "inc_hits":       _serialize_hits(rec.get("inc_hits", [])),
    }


def deserialize_battle_record(rec):
    start_ts = _dt_from_json(rec.get("start_ts"))
    end_ts = _dt_from_json(rec.get("end_ts"))
    if not start_ts or not end_ts:
        return None
    active_elapsed = float(rec.get("active_elapsed", 0.0))
    inc_elapsed = float(rec.get("inc_elapsed", 0.0))
    total_dmg = int(rec.get("total_dmg", 0))
    inc_total_dmg = int(rec.get("inc_total_dmg", 0))
    return {
        "start_ts":       start_ts,
        "end_ts":         end_ts,
        "total_dmg":      total_dmg,
        "active_elapsed": active_elapsed,
        "dps":            float(rec.get("dps", (total_dmg / active_elapsed) if active_elapsed > 0 else 0.0)),
        "hits":           _deserialize_hits(rec.get("hits", [])),
        "first_target":   str(rec.get("first_target", "")),
        "inc_total_dmg":  inc_total_dmg,
        "inc_elapsed":    inc_elapsed,
        "inc_dps":        float(rec.get("inc_dps", (inc_total_dmg / inc_elapsed) if inc_elapsed > 0 else 0.0)),
        "inc_hits":       _deserialize_hits(rec.get("inc_hits", [])),
    }


def _history_session_label(saved_at, path, used_labels):
    ts = parse_ts(str(saved_at or ""))
    if not ts:
        try:
            ts = datetime.fromtimestamp(os.path.getmtime(path))
        except Exception:
            ts = None
    label = ts.strftime("%Y-%m-%d %H:%M") if ts else os.path.splitext(os.path.basename(path))[0]
    base = label
    idx = 2
    while label in used_labels:
        label = f"{base} #{idx}"
        idx += 1
    used_labels.add(label)
    return label


def load_history_sessions():
    sessions = []
    if not os.path.isdir(HISTORY_DIR):
        return sessions
    files = sorted(glob.glob(os.path.join(HISTORY_DIR, "*.json")), reverse=True)[:MAX_HISTORY_FILES]
    used_labels = set()
    for path in files:
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            records = data.get("records", {})
            if not isinstance(records, dict):
                continue
            records_by_char = {}
            for char, items in records.items():
                if not isinstance(items, list):
                    continue
                bucket = records_by_char.setdefault(str(char), [])
                for item in items:
                    rec = deserialize_battle_record(item)
                    if rec:
                        bucket.append(rec)
            if records_by_char:
                sessions.append({
                    "key": path,
                    "label": _history_session_label(data.get("saved_at"), path, used_labels),
                    "records": records_by_char,
                })
        except Exception:
            pass
    return sessions


def merge_history_sessions(sessions):
    records_by_char = {}
    for session in reversed(sessions or []):
        for char, records in session.get("records", {}).items():
            records_by_char.setdefault(char, []).extend(records)
    return records_by_char


def load_history_archives():
    return merge_history_sessions(load_history_sessions())


def save_history_archive(records_by_char):
    payload_records = {}
    for char, records in (records_by_char or {}).items():
        clean = []
        for rec in records or []:
            if rec.get("total_dmg", 0) > 0 or rec.get("inc_total_dmg", 0) > 0:
                clean.append(serialize_battle_record(rec))
        if clean:
            payload_records[str(char)] = clean
    if not payload_records:
        return None
    try:
        os.makedirs(HISTORY_DIR, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = os.path.join(HISTORY_DIR, f"eve_dps_history_{stamp}.json")
        txt_path = os.path.join(HISTORY_DIR, f"eve_dps_history_{stamp}.txt")
        payload = {
            "saved_at": datetime.now().strftime("%Y.%m.%d %H:%M:%S"),
            "records": payload_records,
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        text = format_history_records_text(records_by_char)
        if text:
            with open(txt_path, "w", encoding="utf-8") as f:
                f.write(text)
        return path
    except Exception:
        return None


def format_history_records_text(records_by_char):
    lines = []
    saved_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines.append(f"EVE DPS History - saved at {saved_at}")
    lines.append("")
    for char in sorted((records_by_char or {}).keys(), key=natural_sort_key):
        records = records_by_char.get(char) or []
        if not records:
            continue
        lines.append(f"[{char}]")
        lines.append("Battle Time\tFirst Target\tDPS\tDamage\tInc.DPS\tInc.Damage")
        for rec in sorted(records, key=lambda r: r.get("end_ts") or datetime.min, reverse=True):
            lines.append("\t".join([
                f"{fmt_dt_full(rec.get('start_ts'))} ~ {fmt_dt_full(rec.get('end_ts'))}",
                str(rec.get("first_target", "")),
                fmt_num(float(rec.get("dps", 0.0))),
                fmt_int(int(rec.get("total_dmg", 0))),
                fmt_num(float(rec.get("inc_dps", 0.0))),
                fmt_int(int(rec.get("inc_total_dmg", 0))),
            ]))
        lines.append("")
    return "\n".join(lines).rstrip() + "\n" if len(lines) > 2 else ""


def latest_history_text_file():
    if not os.path.isdir(HISTORY_DIR):
        return None
    files = sorted(glob.glob(os.path.join(HISTORY_DIR, "*.txt")), reverse=True)
    return files[0] if files else None


def fit_text_binary(text, max_width, measure_func):
    if measure_func(text) <= max_width: return text
    lo, hi, best = 0, len(text), "…"
    while lo <= hi:
        mid = (lo + hi) // 2
        cand = text[:mid].rstrip() + "…"
        if measure_func(cand) <= max_width: best = cand; lo = mid + 1
        else: hi = mid - 1
    return best


def short_name(text, limit=18):
    if not text: return ""
    if len(text) <= limit: return text
    return text[:max(1, limit-1)] + "…"


def extract_target(line):
    m = TARGET_RE_UNIVERSAL.search(line)
    if not m: return ""
    raw = m.group(1).strip()
    sh = _SHIP_HINT_RE.search(raw)
    if sh: return sh.group(1).strip()
    ps = _PLAIN_SHIP_RE.search(raw)
    if ps: return ps.group(1).strip()
    clean = re.sub(r'<localized\s+hint="[^"]*">', "", raw).strip().rstrip("*").strip()
    return re.sub(r"\[.*", "", clean).strip()


def guess_log_dirs():
    home = os.path.expanduser("~")
    userprofile = os.environ.get("USERPROFILE", "")
    onedrive    = os.environ.get("OneDrive", "")
    roots = [p for p in [home, userprofile, onedrive] if p]
    candidates = []
    for root in roots:
        candidates.extend([
            os.path.join(root, "Documents", "EVE", "logs", "Gamelogs"),
            os.path.join(root, "문서",       "EVE", "logs", "Gamelogs"),
            os.path.join(root, "EVE",        "logs", "Gamelogs"),
            os.path.join(root, "OneDrive", "Documents", "EVE", "logs", "Gamelogs"),
            os.path.join(root, "OneDrive", "문서",       "EVE", "logs", "Gamelogs"),
        ])
    seen, out = set(), []
    for p in candidates:
        norm = os.path.normpath(p)
        if norm.lower() not in seen:
            seen.add(norm.lower()); out.append(norm)
    return out


def find_log_base():
    for p in guess_log_dirs():
        if os.path.isdir(p): return p
    users_root = r"C:\Users"
    if os.path.isdir(users_root):
        try:
            for user_name in os.listdir(users_root):
                base = os.path.join(users_root, user_name)
                for sub in [
                    os.path.join(base, "Documents", "EVE", "logs", "Gamelogs"),
                    os.path.join(base, "문서",       "EVE", "logs", "Gamelogs"),
                    os.path.join(base, "OneDrive", "Documents", "EVE", "logs", "Gamelogs"),
                    os.path.join(base, "OneDrive", "문서",       "EVE", "logs", "Gamelogs"),
                ]:
                    if os.path.isdir(sub): return sub
        except Exception:
            pass
    return None


def list_running_eve_characters_from_windows():
    chars = set()
    if win32gui is None: return chars
    def callback(hwnd, _):
        try:
            if not win32gui.IsWindowVisible(hwnd): return
            title = (win32gui.GetWindowText(hwnd) or "").strip()
            if not title.startswith(TITLE_PREFIX): return
            name = title[len(TITLE_PREFIX):].strip()
            if name: chars.add(name)
        except Exception:
            return
    try: win32gui.EnumWindows(callback, None)
    except Exception: pass
    return chars


def _scale_wav(src_path: str, volume: float) -> str:
    volume = clamp(volume, 0.0, 1.0)
    if volume >= 0.99: return src_path
    try:
        with wave.open(src_path, "rb") as wf:
            params = wf.getparams(); frames = wf.readframes(params.nframes)
        sw = params.sampwidth; n = params.nframes * params.nchannels
        if sw == 2:
            samples = list(struct.unpack(f"<{n}h", frames))
            out_frames = struct.pack(f"<{n}h", *[int(s * volume) for s in samples])
        elif sw == 1:
            samples = list(struct.unpack(f"{n}B", frames))
            out_frames = struct.pack(f"{n}B", *[int((s-128)*volume)+128 for s in samples])
        else:
            return src_path
        tmp = tempfile.NamedTemporaryFile(suffix=".wav", delete=False)
        with wave.open(tmp.name, "wb") as wf:
            wf.setparams(params); wf.writeframes(out_frames)
        return tmp.name
    except Exception:
        return src_path


# ── Alarm Manager ─────────────────────────────────────────────────────────────
class AlarmManager:
    def __init__(self, cfg_ref):
        self.cfg_ref = cfg_ref
        self._lock   = threading.Lock()
        self._alarming     = {}   # char -> bool
        self._blink_state  = {}   # char -> bool
        self._last_hit_mono = {}  # char -> monotonic time of last incoming hit
        self._repeat_thread = None
        self._repeat_stop   = threading.Event()
        self._sound_playing = False

    def _acfg(self): return self.cfg_ref.get("alarm", {})

    def check(self, char: str, inc_dps_30s: float, last_hit_mono: float):
        """
        inc_dps_30s: 최근 30초 슬라이딩 윈도우 DPS
        last_hit_mono: 마지막 피격 시각 (monotonic). 피격 없으면 0.
        경보 조건:
          - inc_dps_30s >= threshold
          - AND 마지막 피격 후 INC_NO_HIT_EXPIRE_SEC 이내
        """
        acfg      = self._acfg()
        threshold = float(acfg.get("threshold", 500))
        now       = time.monotonic()

        # 5초 이상 피격 없으면 경보 강제 해제
        hit_age   = (now - last_hit_mono) if last_hit_mono > 0 else 9999
        triggered = (inc_dps_30s >= threshold) and (hit_age <= INC_NO_HIT_EXPIRE_SEC)

        with self._lock:
            was = self._alarming.get(char, False)
            self._alarming[char] = triggered
            if last_hit_mono > 0:
                self._last_hit_mono[char] = last_hit_mono

        acfg = self._acfg()
        if acfg.get("enabled_sound", False):
            repeat = acfg.get("repeat", False)
            if repeat:
                if triggered and not self._is_repeat_running(): self._start_repeat()
                elif not triggered and self._is_repeat_running(): self._stop_repeat()
            else:
                if triggered and not was: self._play_once()

    def _is_repeat_running(self):
        return self._repeat_thread is not None and self._repeat_thread.is_alive()

    def _start_repeat(self):
        if self._is_repeat_running(): return
        self._repeat_stop.clear()
        self._repeat_thread = threading.Thread(target=self._repeat_loop, daemon=True)
        self._repeat_thread.start()

    def _stop_repeat(self): self._repeat_stop.set()

    def _repeat_loop(self):
        while not self._repeat_stop.is_set():
            if not self._sound_playing: self._play_once()
            self._repeat_stop.wait(1.5)

    def _play_once(self):
        if not WINSOUND_OK or self._sound_playing: return
        def _run():
            self._sound_playing = True
            try:
                acfg = self._acfg()
                wav  = acfg.get("wav_path", "") or _DEFAULT_WAV_PATH
                if wav and os.path.exists(wav):
                    scaled = _scale_wav(wav, float(acfg.get("volume", 0.7)))
                    _winsound.PlaySound(scaled, _winsound.SND_FILENAME | _winsound.SND_NODEFAULT)
            except Exception:
                pass
            finally:
                self._sound_playing = False
        threading.Thread(target=_run, daemon=True).start()

    def tick_blink(self):
        with self._lock:
            for char, alarming in self._alarming.items():
                if alarming: self._blink_state[char] = not self._blink_state.get(char, False)
                else:        self._blink_state[char] = False

    def stop_all(self):
        self._stop_repeat()
        with self._lock:
            self._alarming.clear(); self._blink_state.clear()

    def is_alarming(self, char):
        with self._lock: return self._alarming.get(char, False)

    def blink_state(self, char):
        with self._lock: return self._blink_state.get(char, False)

    def any_alarming(self):
        with self._lock: return any(self._alarming.values())


# ── DPS Engine ────────────────────────────────────────────────────────────────
class DPSEngine:
    def __init__(self, char_name):
        self.char_name  = char_name
        self.log_base   = find_log_base()
        self.log_file   = None
        self.file_pos   = 0
        self.current_battle = None
        self.battles    = []
        self._lock      = threading.Lock()
        self._stop      = False
        self._thread    = None

        # incoming 히트 버퍼 (슬라이딩 윈도우용): deque of (monotonic_time, dmg)
        self._inc_hits_mono: deque = deque()
        # 마지막 incoming 히트 시각 (monotonic)
        self._last_inc_mono: float = 0.0

    def start(self):
        if self._thread and self._thread.is_alive(): return
        self._stop = False
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def stop(self): self._stop = True

    def hard_reset(self):
        with self._lock:
            self.current_battle = None
            self.battles        = []
            self._inc_hits_mono.clear()
            self._last_inc_mono = 0.0
        self.log_file = None; self.file_pos = 0

    def _run(self):
        while not self._stop:
            try: self._tick()
            except Exception: pass
            time.sleep(0.5)

    def _find_log_for_char(self):
        if not self.log_base or not os.path.isdir(self.log_base):
            self.log_base = find_log_base()
        if not self.log_base or not os.path.isdir(self.log_base): return None
        files = sorted(glob.glob(os.path.join(self.log_base, "*.txt")), reverse=True)
        for fp in files:
            try:
                with open(fp, "r", encoding="utf-8", errors="ignore") as f:
                    hdr = f.read(3000)
                m = LISTENER_RE_KO.search(hdr) or LISTENER_RE_EN.search(hdr)
                if m and m.group(1).strip() == self.char_name: return fp
            except Exception:
                continue
        return None

    def _new_battle(self, ts):
        return {
            "start_ts": ts, "last_event_ts": ts,
            "last_event_mono": time.monotonic(),
            "last_outgoing_ts": None,
            "total_dmg": 0, "active_elapsed": 0.0,
            "hits": [], "first_target": "",
            "inc_total_dmg": 0, "inc_hits": [],
            "inc_first_ts": None, "inc_last_ts": None,
        }

    def _finalize_current_battle(self):
        if not self.current_battle: return
        battle = self.current_battle
        inc_first = battle.get("inc_first_ts")
        inc_last  = battle.get("inc_last_ts")
        inc_elapsed = (inc_last - inc_first).total_seconds() \
            if (inc_first and inc_last and inc_first != inc_last) else 0.0
        inc_dmg = battle.get("inc_total_dmg", 0)
        record = {
            "start_ts":      battle["start_ts"],
            "end_ts":        battle["last_event_ts"],
            "total_dmg":     battle["total_dmg"],
            "active_elapsed":battle["active_elapsed"],
            "dps":           (battle["total_dmg"] / battle["active_elapsed"]) if battle["active_elapsed"] > 0 else 0.0,
            "hits":          list(battle["hits"]),
            "first_target":  battle["first_target"],
            "inc_total_dmg": inc_dmg,
            "inc_elapsed":   inc_elapsed,
            "inc_dps":       (inc_dmg / inc_elapsed) if inc_elapsed > 0 else 0.0,
            "inc_hits":      list(battle.get("inc_hits", [])),
        }
        self.battles.append(record)
        self.battles = self.battles[-MAX_HISTORY_BATTLES:]
        self.current_battle = None

    def _register_event(self, ts):
        if self.current_battle is None:
            self.current_battle = self._new_battle(ts); return
        gap = (ts - self.current_battle["last_event_ts"]).total_seconds()
        if gap > SESSION_TIMEOUT:
            self._finalize_current_battle()
            self.current_battle = self._new_battle(ts)
        else:
            self.current_battle["last_event_ts"]  = ts
            self.current_battle["last_event_mono"] = time.monotonic()

    def _register_outgoing(self, ts, dmg, target):
        if self.current_battle is None:
            self.current_battle = self._new_battle(ts)
        if self.current_battle["first_target"] == "":
            self.current_battle["first_target"] = target
        if self.current_battle["last_outgoing_ts"] is not None:
            gap = (ts - self.current_battle["last_outgoing_ts"]).total_seconds()
            self.current_battle["active_elapsed"] += max(0.0, gap)
        self.current_battle["last_outgoing_ts"] = ts
        self.current_battle["total_dmg"]        += dmg
        self.current_battle["hits"].append((ts, dmg))

    def _register_incoming(self, ts, dmg, mono_now: float):
        if self.current_battle is None:
            self.current_battle = self._new_battle(ts)
        b = self.current_battle
        b["inc_total_dmg"] = b.get("inc_total_dmg", 0) + dmg
        b.setdefault("inc_hits", []).append((ts, dmg))
        if b.get("inc_first_ts") is None: b["inc_first_ts"] = ts
        b["inc_last_ts"] = ts

        # 슬라이딩 윈도우 버퍼에 추가 (monotonic 기준)
        self._inc_hits_mono.append((mono_now, dmg))
        self._last_inc_mono = mono_now

    def _check_timeout(self):
        if self.current_battle is None: return
        if (time.monotonic() - self.current_battle["last_event_mono"]) >= SESSION_TIMEOUT:
            self._finalize_current_battle()

    def _tick(self):
        nf = self._find_log_for_char()
        if nf != self.log_file:
            self.log_file = nf
            if self.log_file and os.path.exists(self.log_file):
                try:
                    with open(self.log_file, "r", encoding="utf-8", errors="ignore") as f:
                        f.seek(0, os.SEEK_END); self.file_pos = f.tell()
                except Exception:
                    self.file_pos = 0
            else:
                self.file_pos = 0
        if not self.log_file: return
        try:
            with open(self.log_file, "r", encoding="utf-8", errors="ignore") as f:
                f.seek(self.file_pos); chunk = f.read(); self.file_pos = f.tell()
        except Exception:
            return
        events = []
        _ENTRY_RE = re.compile(r"(?=\[\s*\d{4}\.\d{2}\.\d{2}\s+\d{2}:\d{2}:\d{2}\s*\])")
        for entry in _ENTRY_RE.split(chunk):
            if not entry.strip(): continue
            flat = entry.replace("\n", " ").replace("\r", " ")
            m = OUTGOING_RE.search(flat)
            if m:
                ts = parse_ts(m.group(1))
                if ts:
                    try: events.append((ts, "out", int(m.group(2)), extract_target(flat)))
                    except Exception: pass
                continue
            m = INCOMING_RE.search(flat)
            if m:
                ts = parse_ts(m.group(1))
                if ts:
                    try: events.append((ts, "in", int(m.group(2)), ""))
                    except Exception: pass
        events.sort(key=lambda x: (x[0], 0 if x[1] == "out" else 1))
        mono_now = time.monotonic()
        with self._lock:
            for ts, kind, dmg, target in events:
                self._register_event(ts)
                if   kind == "out": self._register_outgoing(ts, dmg, target)
                elif kind == "in":  self._register_incoming(ts, dmg, mono_now)
                if self.current_battle:
                    self.current_battle["last_event_ts"]  = ts
                    self.current_battle["last_event_mono"] = mono_now
            self._check_timeout()

    def _sliding_inc_dps(self) -> float:
        """최근 INC_SLIDING_WINDOW_SEC 초 내 incoming 피해의 DPS 계산"""
        cutoff = time.monotonic() - INC_SLIDING_WINDOW_SEC
        # 만료된 항목 제거
        while self._inc_hits_mono and self._inc_hits_mono[0][0] < cutoff:
            self._inc_hits_mono.popleft()
        if not self._inc_hits_mono:
            return 0.0
        total_dmg = sum(d for _, d in self._inc_hits_mono)
        # 창 길이: 첫 히트 시각 ~ 지금 (최소 1초)
        elapsed = max(1.0, time.monotonic() - self._inc_hits_mono[0][0])
        return total_dmg / elapsed

    def get_status(self):
        with self._lock:
            finished = list(self.battles)
            current  = dict(self.current_battle) if self.current_battle else None

        total_damage = total_secs = 0.0
        total_inc_dmg = total_inc_elapsed = 0.0
        for b in finished:
            total_damage     += b["total_dmg"]
            total_secs       += b["active_elapsed"]
            total_inc_dmg    += b.get("inc_total_dmg", 0)
            total_inc_elapsed += b.get("inc_elapsed", 0.0)

        recent_dps = 0.0
        in_combat  = current is not None
        recent_inc_dps = 0.0  # 최근 전투 전체 받는 DPS

        if current:
            total_damage += current["total_dmg"]
            total_secs   += current["active_elapsed"]
            total_inc_dmg += current.get("inc_total_dmg", 0)
            if current["active_elapsed"] > 0:
                recent_dps = current["total_dmg"] / current["active_elapsed"]
            inc_first = current.get("inc_first_ts")
            inc_last  = current.get("inc_last_ts")
            if inc_first and inc_last and inc_first != inc_last:
                inc_elapsed = (inc_last - inc_first).total_seconds()
                if inc_elapsed > 0:
                    recent_inc_dps    = current.get("inc_total_dmg", 0) / inc_elapsed
                    total_inc_elapsed += inc_elapsed
            elif current.get("inc_total_dmg", 0) > 0 and current["active_elapsed"] > 0:
                recent_inc_dps    = current.get("inc_total_dmg", 0) / current["active_elapsed"]
                total_inc_elapsed += current["active_elapsed"]

        # 슬라이딩 윈도우 DPS (lock 없이 읽어도 대략 OK)
        sliding_inc_dps = self._sliding_inc_dps()

        return {
            "dps":            recent_dps,
            "tdps":           (total_damage / total_secs) if total_secs > 0 else 0.0,
            "tdam":           total_damage,
            "in_combat":      in_combat,
            # ── Incoming 3종 ──
            "inc_dps_30s":    sliding_inc_dps,          # 최근 30초 슬라이딩 DPS
            "inc_dps":        recent_inc_dps,            # 최근 전투 전체 받는 DPS
            "inc_tdam":       total_inc_dmg,             # 받은 총 피해량
            # 경보용
            "last_inc_mono":  self._last_inc_mono,
        }

    def get_battle_records(self):
        with self._lock:
            records = list(self.battles)
            if self.current_battle:
                c = self.current_battle
                records.append({
                    "start_ts":      c["start_ts"],
                    "end_ts":        c["last_event_ts"],
                    "total_dmg":     c["total_dmg"],
                    "active_elapsed":c["active_elapsed"],
                    "dps":           (c["total_dmg"] / c["active_elapsed"]) if c["active_elapsed"] > 0 else 0.0,
                    "hits":          list(c["hits"]),
                    "first_target":  c["first_target"],
                    "inc_total_dmg": c.get("inc_total_dmg", 0),
                    "inc_elapsed":   0.0, "inc_dps": 0.0,
                    "inc_hits":      list(c.get("inc_hits", [])),
                })
            return records


# ── Tooltip ───────────────────────────────────────────────────────────────────
class Tooltip:
    def __init__(self, root):
        self.root = root; self.win = None

    def show(self, x, y, text):
        self.hide()
        self.win = tk.Toplevel(self.root)
        self.win.overrideredirect(True)
        self.win.attributes("-topmost", True)
        self.win.configure(bg=C_TOOLTIP_BG)
        tk.Label(self.win, text=text, bg=C_TOOLTIP_BG, fg=C_TOOLTIP_FG,
                 font=("Consolas", 9), padx=7, pady=4).pack()
        self.win.geometry(f"+{x+12}+{y+12}")

    def hide(self):
        if self.win and self.win.winfo_exists():
            try: self.win.destroy()
            except Exception: pass
        self.win = None


# ── Alarm Settings Popup ──────────────────────────────────────────────────────
class AlarmSettingsPopup(tk.Toplevel):
    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        apply_app_icon(self)
        self.overrideredirect(False)
        self.attributes("-topmost", True)
        self.title("Alarm Settings")
        self.configure(bg=C_BG)
        self.resizable(False, False)
        acfg = app.cfg.get("alarm", {})
        pad  = dict(padx=12, pady=6)

        self.var_visual = tk.BooleanVar(value=acfg.get("enabled_visual", True))
        tk.Checkbutton(self, text="Visual alarm (blink)", variable=self.var_visual,
                       bg=C_BG, fg=C_TEXT, selectcolor=C_PANEL, activebackground=C_BG,
                       activeforeground=C_HEADER, font=("Consolas", 10)).grid(
            row=0, column=0, columnspan=2, sticky="w", **pad)

        self.var_sound = tk.BooleanVar(value=acfg.get("enabled_sound", False))
        tk.Checkbutton(self, text="Sound alarm", variable=self.var_sound,
                       bg=C_BG, fg=C_TEXT, selectcolor=C_PANEL, activebackground=C_BG,
                       activeforeground=C_HEADER, font=("Consolas", 10),
                       command=self._toggle_sound_opts).grid(
            row=1, column=0, columnspan=2, sticky="w", **pad)

        tk.Label(self, text="Inc.DPS(30s) threshold:", bg=C_BG, fg=C_MUTED,
                 font=("Consolas", 9)).grid(row=2, column=0, sticky="e", **pad)
        self.var_thresh = tk.StringVar(value=str(int(acfg.get("threshold", 500))))
        tk.Entry(self, textvariable=self.var_thresh, width=8,
                 bg=C_EDIT_BG, fg=C_EDIT_FG, insertbackground="#fff",
                 relief="flat", font=("Consolas", 10)).grid(row=2, column=1, sticky="w", **pad)

        tk.Label(self, text="(피격 5초 후 자동 해제)", bg=C_BG, fg=C_MUTED,
                 font=("Consolas", 8)).grid(row=3, column=0, columnspan=2, sticky="w", padx=12, pady=2)

        tk.Label(self, text="Volume:", bg=C_BG, fg=C_MUTED,
                 font=("Consolas", 9)).grid(row=4, column=0, sticky="e", **pad)
        self.var_vol = tk.DoubleVar(value=acfg.get("volume", 0.7))
        vol_frame = tk.Frame(self, bg=C_BG)
        vol_frame.grid(row=4, column=1, sticky="w", **pad)
        self.vol_slider = tk.Scale(vol_frame, from_=0.0, to=1.0, resolution=0.05,
                                   orient="horizontal", variable=self.var_vol,
                                   bg=C_BG, fg=C_TEXT, troughcolor=C_PANEL,
                                   highlightthickness=0, length=120, activebackground=C_HEADER)
        self.vol_slider.pack(side="left")

        self.var_repeat = tk.BooleanVar(value=acfg.get("repeat", False))
        self.chk_repeat = tk.Checkbutton(self, text="Repeat while Inc.DPS(30s) exceeds threshold",
                                          variable=self.var_repeat, bg=C_BG, fg=C_TEXT,
                                          selectcolor=C_PANEL, activebackground=C_BG,
                                          activeforeground=C_HEADER, font=("Consolas", 9))
        self.chk_repeat.grid(row=5, column=0, columnspan=2, sticky="w", **pad)

        tk.Label(self, text="Alert sound:", bg=C_BG, fg=C_MUTED,
                 font=("Consolas", 9)).grid(row=6, column=0, sticky="e", **pad)
        wav_frame = tk.Frame(self, bg=C_BG)
        wav_frame.grid(row=6, column=1, sticky="w", **pad)
        self.var_wav = tk.StringVar(value=acfg.get("wav_path", "") or "(built-in)")
        tk.Label(wav_frame, textvariable=self.var_wav, bg=C_BG, fg=C_MUTED,
                 font=("Consolas", 8), width=22, anchor="w").pack(side="left")
        tk.Button(wav_frame, text="Browse", command=self._browse_wav,
                  bg=C_PANEL, fg=C_HEADER, activebackground=C_BORDER,
                  relief="flat", font=("Consolas", 8), padx=6).pack(side="left", padx=4)
        tk.Button(wav_frame, text="Reset", command=self._reset_wav,
                  bg=C_PANEL, fg=C_MUTED, activebackground=C_BORDER,
                  relief="flat", font=("Consolas", 8), padx=6).pack(side="left")

        btn_frame = tk.Frame(self, bg=C_BG)
        btn_frame.grid(row=7, column=0, columnspan=2, pady=(4, 10))
        tk.Button(btn_frame, text="Test sound", command=self._test_sound,
                  bg=C_PANEL, fg=C_HEADER, activebackground=C_BORDER,
                  relief="flat", font=("Consolas", 9), padx=10).pack(side="left", padx=6)
        tk.Button(btn_frame, text="Save", command=self._save,
                  bg=C_PANEL, fg=C_HEADER, activebackground=C_BORDER,
                  relief="flat", font=("Consolas", 9), padx=10).pack(side="left", padx=6)
        tk.Button(btn_frame, text="Close", command=self.destroy,
                  bg=C_PANEL, fg=C_MUTED, activebackground=C_BORDER,
                  relief="flat", font=("Consolas", 9), padx=10).pack(side="left", padx=6)

        self._toggle_sound_opts()
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w, h   = self.winfo_reqwidth(), self.winfo_reqheight()
        self.geometry(f"+{(sw-w)//2}+{(sh-h)//2}")

    def _toggle_sound_opts(self):
        state = "normal" if self.var_sound.get() else "disabled"
        self.vol_slider.config(state=state); self.chk_repeat.config(state=state)

    def _browse_wav(self):
        path = filedialog.askopenfilename(title="Select alert WAV file",
            filetypes=[("WAV files", "*.wav"), ("All files", "*.*")])
        if path: self.var_wav.set(path)

    def _reset_wav(self): self.var_wav.set("(built-in)")

    def _test_sound(self):
        if not WINSOUND_OK:
            messagebox.showinfo("Sound", "winsound not available."); return
        wav = self.var_wav.get()
        if wav == "(built-in)": wav = _DEFAULT_WAV_PATH
        if not wav or not os.path.exists(wav):
            messagebox.showwarning("Sound", "WAV file not found."); return
        def _run():
            try:
                scaled = _scale_wav(wav, float(self.var_vol.get()))
                _winsound.PlaySound(scaled, _winsound.SND_FILENAME | _winsound.SND_NODEFAULT)
            except Exception: pass
        threading.Thread(target=_run, daemon=True).start()

    def _save(self):
        try: thresh = float(self.var_thresh.get())
        except ValueError: thresh = 500.0
        wav = self.var_wav.get()
        if wav == "(built-in)": wav = ""
        self.app.cfg["alarm"] = {
            "enabled_visual": self.var_visual.get(),
            "enabled_sound":  self.var_sound.get(),
            "threshold":      thresh,
            "volume":         float(self.var_vol.get()),
            "repeat":         self.var_repeat.get(),
            "wav_path":       wav,
        }
        save_config(self.app.cfg)
        if not self.var_sound.get():
            self.app.alarm_mgr._stop_repeat()
        self.destroy()


# ── Sort Menu Popup ───────────────────────────────────────────────────────────
class SortMenuPopup(tk.Toplevel):
    def __init__(self, master, app, x, y):
        super().__init__(master)
        self.app = app
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.configure(bg=C_BORDER)
        options = [
            (SORT_TOP_DPS,  "Top dealer first"),
            (SORT_BOT_DPS,  "Bottom dealer first"),
            (SORT_NAME_ASC, "Name A → Z"),
            (SORT_NAME_DESC,"Name Z → A"),
            (SORT_MANUAL,   "Manual order"),
        ]
        cur = app.cfg.get("sort_mode", SORT_NAME_ASC)
        for mode, label in options:
            fg  = C_HEADER if mode == cur else C_TEXT
            btn = tk.Label(self, text=label, bg=C_PANEL, fg=fg,
                           font=("Consolas", 9), padx=14, pady=5, anchor="w", width=22)
            btn.pack(fill="x", pady=1)
            btn.bind("<Button-1>",  lambda _e, m=mode: self._pick(m))
            btn.bind("<Enter>",     lambda _e, b=btn: b.config(bg=C_BORDER))
            btn.bind("<Leave>",     lambda _e, b=btn: b.config(bg=C_PANEL))
        self.bind("<FocusOut>", lambda _e: self.destroy())
        self.geometry(f"+{x}+{y}")
        self.focus_set()

    def _pick(self, mode):
        self.app.cfg["sort_mode"] = mode
        save_config(self.app.cfg)
        self.app.window.render_all(force_layout=True)
        self.destroy()


# ── Graph Window ──────────────────────────────────────────────────────────────
class GraphWindow(tk.Toplevel):
    TIME_COL_W  = 110
    DUR_COL_W   = 76
    ENEMY_COL_W = 120
    STAT_COL_W  = 70
    HEADER_H    = 40
    RESIZE_GRIP = 10

    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        apply_app_icon(self)
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.attributes("-alpha", 0.96)
        self.config(bg=TRANSPARENT)
        self.attributes("-transparentcolor", TRANSPARENT)
        self._dx = self._dy = 0
        self._resizing = False; self._resize_edge = None; self._resize_start = None
        gw = self.app.cfg.get("graph_window", {})
        self._win_w = int(gw.get("w", 1020))
        self._win_h = int(gw.get("h", 580))
        gx, gy = int(gw.get("x", 180)), int(gw.get("y", 180))
        self.geometry(f"{self._win_w}x{self._win_h}+{gx}+{gy}")

        outer = tk.Frame(self, bg=TRANSPARENT)
        outer.pack(fill="both", expand=True)
        self.bg_canvas = tk.Canvas(outer, bg=TRANSPARENT, highlightthickness=0)
        self.bg_canvas.pack(fill="both", expand=True)
        self.bg_canvas.bind("<Configure>",       self._on_resize)
        self.bg_canvas.bind("<ButtonPress-1>",   self._on_press)
        self.bg_canvas.bind("<B1-Motion>",        self._on_drag)
        self.bg_canvas.bind("<ButtonRelease-1>",  self._on_release)
        self.bg_canvas.bind("<Motion>",           self._on_mouse_move)
        self.bind("<ButtonPress-1>",   self._on_toplevel_press)
        self.bind("<B1-Motion>",       self._on_toplevel_drag)
        self.bind("<ButtonRelease-1>", self._on_toplevel_release)
        self.bind("<Motion>",          self._on_toplevel_motion)
        self.body = tk.Frame(self.bg_canvas, bg=C_BG)
        for w in (self.body, self.bg_canvas):
            w.bind("<ButtonPress-1>",  self._drag_start)
            w.bind("<B1-Motion>",      self._drag_move)
            w.bind("<ButtonRelease-1>",self._drag_end)
        style = ttk.Style(self)
        try: style.theme_use("default")
        except Exception: pass
        style.configure("Blue.TNotebook",     background=C_BG, borderwidth=0)
        style.configure("Blue.TNotebook.Tab", background=C_PANEL_2, foreground=C_TEXT, padding=(12, 6))
        style.map("Blue.TNotebook.Tab", background=[("selected", C_PANEL)], foreground=[("selected", C_HEADER)])
        style.configure("History.Treeview",         rowheight=24, background=C_PANEL, fieldbackground=C_PANEL, foreground=C_TEXT)
        style.configure("History.Treeview.Heading", background=C_PANEL_2, foreground=C_HEADER)
        style.configure("Blue.Vertical.TScrollbar",   background=C_BORDER, troughcolor=C_BG, arrowcolor=C_HEADER, bordercolor=C_BG, lightcolor=C_BORDER, darkcolor=C_BORDER)
        style.configure("Blue.Horizontal.TScrollbar", background=C_BORDER, troughcolor=C_BG, arrowcolor=C_HEADER, bordercolor=C_BG, lightcolor=C_BORDER, darkcolor=C_BORDER)
        self.drag_bar = tk.Frame(self.body, bg=C_PANEL, height=6, cursor="fleur")
        self.drag_bar.pack(fill="x", side="top")
        for w in (self.drag_bar,):
            w.bind("<ButtonPress-1>",  self._drag_start)
            w.bind("<B1-Motion>",      self._drag_move)
            w.bind("<ButtonRelease-1>",self._drag_end)
        nb_wrap = tk.Frame(self.body, bg=C_BG)
        nb_wrap.pack(fill="both", expand=True)
        self.notebook = ttk.Notebook(nb_wrap, style="Blue.TNotebook")
        self.notebook.pack(fill="both", expand=True)
        self.btn_close = tk.Label(nb_wrap, text="✕", bg=C_PANEL_2, fg=C_ICON,
                                  font=("Segoe UI Symbol", 10), padx=6, pady=4, cursor="hand2")
        self.btn_close.place(relx=1.0, rely=0.0, anchor="ne")
        self.btn_close.bind("<Button-1>", lambda _e: self.destroy())
        self.btn_close.bind("<Enter>",    lambda _e: self.btn_close.config(fg=C_HEADER))
        self.btn_close.bind("<Leave>",    lambda _e: self.btn_close.config(fg=C_ICON))
        self.graph_frame   = tk.Frame(self.notebook, bg=C_BG)
        self.history_frame = tk.Frame(self.notebook, bg=C_BG)
        self.notebook.add(self.graph_frame,   text="Recent Battle")
        self.notebook.add(self.history_frame, text="History")
        self.canvas = tk.Canvas(self.graph_frame, bg=C_BG, highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)
        self.history_header_canvas = None
        self.tree = self.tree_cols = self.h_scroll = None
        self.history_select_var = tk.StringVar(value=self.app.selected_history_label)
        self.history_date_var = tk.StringVar(value="Date: -")
        self.history_combo = None
        self._build_history_tree()
        self.after(50, self._first_render)

    def _history_chars(self):
        return self.app.history_characters()

    def _history_source_records(self):
        return self.app.selected_history_records()

    def _include_live_history(self):
        return self.app.selected_history_label == HISTORY_CURRENT_LABEL

    def _get_resize_edge(self, event):
        g = self.RESIZE_GRIP; w, h = self._win_w, self._win_h
        try:
            x = event.x_root - self.winfo_rootx()
            y = event.y_root - self.winfo_rooty()
        except Exception:
            return None
        edges = []
        if x < g: edges.append("w")
        if x > w - g: edges.append("e")
        if y < g: edges.append("n")
        if y > h - g: edges.append("s")
        return "".join(edges) if edges else None

    def _cursor_for_edge(self, edge):
        if not edge: return ""
        cursors = {"n":"top_side","s":"bottom_side","e":"right_side","w":"left_side",
                   "ne":"top_right_corner","nw":"top_left_corner",
                   "se":"bottom_right_corner","sw":"bottom_left_corner"}
        return cursors.get(edge, "")

    def _on_mouse_move(self, e):
        if self._resizing: return
        try: self.config(cursor=self._cursor_for_edge(self._get_resize_edge(e)))
        except Exception: pass

    def _on_press(self, e):
        edge = self._get_resize_edge(e)
        if edge:
            self._resizing = True; self._resize_edge = edge
            self._resize_start = (e.x_root, e.y_root, self._win_w, self._win_h,
                                  self.winfo_x(), self.winfo_y())
        else:
            self._resizing = False; self._drag_start(e)

    def _on_drag(self, e):
        if self._resizing:
            sx, sy, sw, sh, ox, oy = self._resize_start
            dx, dy = e.x_root-sx, e.y_root-sy; edge = self._resize_edge
            new_w, new_h, new_x, new_y = sw, sh, ox, oy
            if "e" in edge: new_w = max(400, sw+dx)
            if "s" in edge: new_h = max(300, sh+dy)
            if "w" in edge: new_w = max(400, sw-dx); new_x = ox+dx if new_w>400 else ox
            if "n" in edge: new_h = max(300, sh-dy); new_y = oy+dy if new_h>300 else oy
            self._win_w, self._win_h = int(new_w), int(new_h)
            self.geometry(f"{self._win_w}x{self._win_h}+{int(new_x)}+{int(new_y)}")
        elif self._dx is not None:
            self._drag_move(e)

    def _on_release(self, e):
        if self._resizing: self._resizing = False; self._save_pos()
        else: self._drag_end(e); self._resizing = False

    def _drag_start(self, e): self._dx = e.x_root-self.winfo_x(); self._dy = e.y_root-self.winfo_y()
    def _drag_move(self, e):  self.geometry(f"+{e.x_root-self._dx}+{e.y_root-self._dy}")
    def _drag_end(self, _e):  self._save_pos()
    def _on_toplevel_press(self, e):   self._on_press(e)
    def _on_toplevel_drag(self, e):    self._on_drag(e)
    def _on_toplevel_release(self, e): self._on_release(e)
    def _on_toplevel_motion(self, e):
        if not self._resizing:
            try: self.config(cursor=self._cursor_for_edge(self._get_resize_edge(e)))
            except Exception: pass

    def _save_pos(self):
        try:
            self.app.cfg["graph_window"] = {
                "x": self.winfo_x(), "y": self.winfo_y(),
                "w": self._win_w,    "h": self._win_h}
            save_config(self.app.cfg)
        except Exception: pass

    def _first_render(self):
        self.update_idletasks()
        self._on_resize(type("E", (), {"width": self._win_w, "height": self._win_h})())
        self.after(950, self._update_loop)

    def _on_resize(self, e):
        w = getattr(e, "width",  self._win_w); h = getattr(e, "height", self._win_h)
        if w < 10 or h < 10: return
        self._win_w, self._win_h = w, h
        self.bg_canvas.delete("all")
        self.bg_canvas.create_polygon(rrect(1,1,self._win_w-1,self._win_h-1,16),
                                      smooth=True, fill=C_BG, outline=C_BORDER, width=1.4)
        g = self.RESIZE_GRIP; rx, ry = self._win_w-2, self._win_h-2
        self.bg_canvas.create_polygon(rx, ry-g, rx, ry, rx-g, ry, fill=C_BORDER, outline="")
        self.bg_canvas.create_window(1, 1, anchor="nw", window=self.body,
                                     width=self._win_w-2, height=self._win_h-2)
        self.render()

    def _get_joint_battles(self):
        chars = self._history_chars()
        intervals = []
        source_records = self._history_source_records()
        for char in chars:
            for rec in source_records.get(char, []):
                intervals.append({"char": char, "start_ts": rec["start_ts"],
                                   "end_ts": rec["end_ts"], "record": rec})
        if self._include_live_history():
            for char in chars:
                eng = self.app.engines.get(char)
                if not eng: continue
                for rec in eng.get_battle_records():
                    intervals.append({"char": char, "start_ts": rec["start_ts"],
                                       "end_ts": rec["end_ts"], "record": rec})
        if not intervals: return []
        intervals.sort(key=lambda x: x["start_ts"])
        merged, cur_start, cur_end, cur_items = [], intervals[0]["start_ts"], intervals[0]["end_ts"], [intervals[0]]
        for item in intervals[1:]:
            if item["start_ts"] <= cur_end:
                if item["end_ts"] > cur_end: cur_end = item["end_ts"]
                cur_items.append(item)
            else:
                merged.append((cur_start, cur_end, cur_items))
                cur_start, cur_end, cur_items = item["start_ts"], item["end_ts"], [item]
        merged.append((cur_start, cur_end, cur_items))
        joint = []
        for start_ts, end_ts, items in merged:
            battle = {"start_ts": start_ts, "end_ts": end_ts, "enemy": "", "chars": {}}
            for char in chars:
                battle["chars"][char] = {"hits": [], "total_dmg": 0, "active_elapsed": 0.0,
                                          "dps": 0.0, "inc_hits": [], "inc_total_dmg": 0,
                                          "inc_elapsed": 0.0, "inc_dps": 0.0}
            first_targets = []
            for item in items:
                char, rec = item["char"], item["record"]
                e = battle["chars"][char]
                e["hits"].extend(rec["hits"]); e["total_dmg"] += rec["total_dmg"]
                e["active_elapsed"] += rec["active_elapsed"]
                e["inc_hits"].extend(rec.get("inc_hits", []))
                e["inc_total_dmg"] += rec.get("inc_total_dmg", 0)
                e["inc_elapsed"]   += rec.get("inc_elapsed", 0.0)
                if rec.get("first_target"):
                    first_targets.append((rec["start_ts"], rec["first_target"]))
            if first_targets:
                first_targets.sort(key=lambda x: x[0])
                battle["enemy"] = short_name(first_targets[0][1], 22)
            for char, e in battle["chars"].items():
                if e["active_elapsed"] > 0: e["dps"]     = e["total_dmg"]    / e["active_elapsed"]
                if e["inc_elapsed"]    > 0: e["inc_dps"] = e["inc_total_dmg"]/ e["inc_elapsed"]
            joint.append(battle)
        joint.sort(key=lambda x: x["end_ts"], reverse=True)
        return joint[:MAX_HISTORY_BATTLES]

    def _history_date_label(self, battles):
        dates = []
        for battle in battles:
            ts = battle.get("start_ts")
            if ts is not None:
                dates.append(ts.strftime("%Y-%m-%d"))
        unique = sorted(set(dates))
        if not unique:
            return "Date: -"
        if len(unique) == 1:
            return f"Date: {unique[0]}"
        return f"Date: {unique[0]} ~ {unique[-1]}"

    def _history_total_width(self):
        return self.TIME_COL_W + self.DUR_COL_W + self.ENEMY_COL_W + len(self._history_chars()) * 2 * self.STAT_COL_W

    def _sync_xview(self, *args):
        if self.tree: self.tree.xview(*args)
        if self.history_header_canvas: self.history_header_canvas.xview(*args)

    def _on_tree_xscroll(self, first, last):
        if self.h_scroll: self.h_scroll.set(first, last)
        if self.history_header_canvas:
            try: self.history_header_canvas.xview_moveto(first)
            except Exception: pass

    def _draw_history_header(self):
        if not self.history_header_canvas: return
        self.history_header_canvas.delete("all")
        total_w = self._history_total_width()
        self.history_header_canvas.configure(scrollregion=(0, 0, total_w, self.HEADER_H))
        x = 0
        def rect(xx, yy, ww, hh, fill):
            self.history_header_canvas.create_rectangle(xx, yy, xx+ww, yy+hh, fill=fill, outline=C_BORDER, width=1)
        def text(cx, cy, s, color, font):
            self.history_header_canvas.create_text(cx, cy, text=s, fill=color, font=font)
        rect(x, 0, self.TIME_COL_W, 20, C_PANEL_2); rect(x, 20, self.TIME_COL_W, 20, C_PANEL)
        text(x+self.TIME_COL_W/2, 10, "Battle", C_HEADER, ("Consolas",10,"bold"))
        text(x+self.TIME_COL_W/2, 30, "Time",   C_TEXT,   ("Consolas",9))
        x += self.TIME_COL_W
        rect(x, 0, self.DUR_COL_W, 20, C_PANEL_2); rect(x, 20, self.DUR_COL_W, 20, C_PANEL)
        text(x+self.DUR_COL_W/2, 10, "Total",    C_HEADER, ("Consolas",10,"bold"))
        text(x+self.DUR_COL_W/2, 30, "Duration", C_TEXT,   ("Consolas",9))
        x += self.DUR_COL_W
        rect(x, 0, self.ENEMY_COL_W, 20, C_PANEL_2); rect(x, 20, self.ENEMY_COL_W, 20, C_PANEL)
        text(x+self.ENEMY_COL_W/2, 10, "First Target", C_HEADER, ("Consolas",10,"bold"))
        text(x+self.ENEMY_COL_W/2, 30, "Enemy",        C_TEXT,   ("Consolas",9))
        x += self.ENEMY_COL_W
        for char in self._history_chars():
            gw = self.STAT_COL_W * 2
            rect(x, 0, gw, 20, C_PANEL_2)
            rect(x, 20, self.STAT_COL_W, 20, C_PANEL)
            rect(x+self.STAT_COL_W, 20, self.STAT_COL_W, 20, C_PANEL)
            text(x+gw/2, 10, char, C_HEADER, ("Consolas",10,"bold"))
            text(x+self.STAT_COL_W/2,             30, "DPS", C_TEXT, ("Consolas",9))
            text(x+self.STAT_COL_W*1.5,            30, "DMG", C_TEXT, ("Consolas",9))
            x += gw

    def _build_history_tree(self):
        for child in self.history_frame.winfo_children(): child.destroy()
        chars = self._history_chars()
        columns = ["time", "duration", "enemy"]
        for c in chars: columns += [f"{c}_dps", f"{c}_dmg"]
        toolbar = tk.Frame(self.history_frame, bg=C_BG)
        toolbar.pack(fill="x", side="top", padx=4, pady=(3, 2))
        btn_kw = dict(bg=C_PANEL, fg=C_HEADER, activebackground=C_BORDER,
                      activeforeground="#ffffff", relief="flat",
                      font=("Consolas", 8), padx=8, pady=2, width=13)
        self.history_select_var.set(self.app.selected_history_label)
        self.history_combo = ttk.Combobox(
            toolbar,
            textvariable=self.history_select_var,
            values=self.app.history_option_labels(),
            state="readonly",
            width=18,
            font=("Consolas", 8),
        )
        self.history_combo.pack(side="left", padx=(0, 5))
        self.history_combo.bind("<<ComboboxSelected>>", self._on_history_selected)
        tk.Button(toolbar, text="Copy All", command=self.copy_history_text,
                  **btn_kw).pack(side="left", padx=(0, 5))
        tk.Button(toolbar, text="Open Full Log", command=self.open_full_history_log,
                  **btn_kw).pack(side="left")
        tk.Label(toolbar, textvariable=self.history_date_var,
                 bg=C_BG, fg=C_MUTED, font=("Consolas", 8)).pack(side="left", padx=(8, 0))
        self.history_header_canvas = tk.Canvas(self.history_frame, bg=C_BG,
                                                height=self.HEADER_H, highlightthickness=0, bd=0)
        self.history_header_canvas.pack(fill="x", side="top")
        grid_wrap = tk.Frame(self.history_frame, bg=C_BG)
        grid_wrap.pack(fill="both", expand=True)
        tree = ttk.Treeview(grid_wrap, columns=columns, show="headings", style="History.Treeview")
        sb_kw = dict(bg=C_BORDER, activebackground=C_HEADER, troughcolor=C_BG,
                     highlightthickness=0, bd=0, relief="flat")
        vsb  = tk.Scrollbar(grid_wrap, orient="vertical",   command=tree.yview, **sb_kw)
        hsb  = tk.Scrollbar(grid_wrap, orient="horizontal", command=self._sync_xview, **sb_kw)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=self._on_tree_xscroll)
        tree.heading("time", text="");  tree.column("time",  width=self.TIME_COL_W,  anchor="w", stretch=False)
        tree.heading("duration", text=""); tree.column("duration", width=self.DUR_COL_W, anchor="center", stretch=False)
        tree.heading("enemy", text=""); tree.column("enemy", width=self.ENEMY_COL_W, anchor="w", stretch=False)
        for c in chars:
            tree.heading(f"{c}_dps", text=""); tree.column(f"{c}_dps", width=self.STAT_COL_W, anchor="e", stretch=False)
            tree.heading(f"{c}_dmg", text=""); tree.column(f"{c}_dmg", width=self.STAT_COL_W, anchor="e", stretch=False)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        grid_wrap.grid_rowconfigure(0, weight=1); grid_wrap.grid_columnconfigure(0, weight=1)
        tree.bind("<MouseWheel>", lambda e: (tree.yview_scroll(-3*(1 if e.delta>0 else -1), "units"), "break"))
        self.tree, self.tree_cols, self.h_scroll = tree, columns, hsb
        self._draw_history_header()

    def _on_history_selected(self, _event=None):
        self.app.selected_history_label = self.history_select_var.get() or HISTORY_CURRENT_LABEL
        self._build_history_tree()
        self.render()

    def refresh_history_selector(self):
        if self.history_combo and self.history_combo.winfo_exists():
            self.history_combo.configure(values=self.app.history_option_labels())
            self.history_select_var.set(self.app.selected_history_label)

    def _update_history_rows(self):
        chars = self._history_chars()
        desired = ["time", "duration", "enemy"]
        for c in chars: desired += [f"{c}_dps", f"{c}_dmg"]
        if desired != self.tree_cols: self._build_history_tree()
        else: self._draw_history_header()
        battles = self._get_joint_battles()
        self.history_date_var.set(self._history_date_label(battles))
        for item in self.tree.get_children(): self.tree.delete(item)
        for battle in battles:
            row = [
                fmt_time_range(battle["start_ts"], battle["end_ts"]),
                fmt_duration(battle["start_ts"], battle["end_ts"]),
                battle.get("enemy",""),
            ]
            for char in chars:
                e = battle["chars"][char]
                row += [fmt_num(e["dps"]), fmt_int(e["total_dmg"])]
            self.tree.insert("", "end", values=row)

    def history_text(self):
        chars = self._history_chars()
        header = ["Battle Time", "Duration", "First Target"]
        for char in chars:
            header += [f"{char} DPS", f"{char} DMG"]
        battles = self._get_joint_battles()
        lines = [self._history_date_label(battles), "\t".join(header)]
        for battle in battles:
            row = [
                fmt_time_range(battle["start_ts"], battle["end_ts"]),
                fmt_duration(battle["start_ts"], battle["end_ts"]),
                battle.get("enemy", ""),
            ]
            for char in chars:
                e = battle["chars"][char]
                row += [fmt_num(e["dps"]), fmt_int(e["total_dmg"])]
            lines.append("\t".join(row))
        return "\n".join(lines)

    def copy_history_text(self):
        text = self.history_text()
        try:
            self.clipboard_clear()
            self.clipboard_append(text)
            self.update()
        except Exception:
            pass

    def open_full_history_log(self):
        path = None
        try:
            os.makedirs(HISTORY_DIR, exist_ok=True)
            text = self.history_text()
            if text.strip():
                path = os.path.join(HISTORY_DIR, "eve_dps_history_current.txt")
                with open(path, "w", encoding="utf-8") as f:
                    f.write(text)
        except Exception:
            path = None
        if not path:
            path = latest_history_text_file()
        if path and os.path.exists(path):
            try:
                os.startfile(path)
            except Exception:
                pass

    def _render_graph(self):
        self.canvas.delete("all")
        w = max(10, self.canvas.winfo_width()); h = max(10, self.canvas.winfo_height())
        pad_l, pad_r, pad_t, pad_b = 70, 20, 95, 60
        x1, y1, x2, y2 = pad_l, pad_t, w-pad_r, h-pad_b
        self.canvas.create_rectangle(0, 0, w, h, fill=C_BG, outline=C_BG)
        self.canvas.create_text(18, 16, anchor="w", text="Recent Actual Damage Dealt",
                                fill=C_HEADER, font=("Consolas",12,"bold"))
        battles = self._get_joint_battles()
        if not battles:
            self.canvas.create_text(w//2, h//2, text="No recent battle data", fill=C_MUTED, font=("Consolas",12))
            return
        battle = battles[0]
        self.canvas.create_text(w-18, 16, anchor="e",
                                text=f"{fmt_dt(battle['start_ts'])} ~ {fmt_dt(battle['end_ts'])}",
                                fill=C_MUTED, font=("Consolas",10))
        total_secs = max(1.0, (battle["end_ts"]-battle["start_ts"]).total_seconds())
        max_y, series = 1.0, []
        for idx, char in enumerate(self._history_chars()):
            e = battle["chars"][char]
            points = []
            for ts, dmg in e["hits"]:
                offset = max(0.0, (ts-battle["start_ts"]).total_seconds())
                points.append((offset, dmg)); max_y = max(max_y, float(dmg))
            series.append({"char": char, "color": GRAPH_COLORS[idx%len(GRAPH_COLORS)], "points": points})
        max_y *= 1.15
        legend_x, legend_y = x1, 44
        per_row = max(1, (w-100)//180)
        for idx, item in enumerate(series):
            lx = legend_x+(idx%per_row)*180; ly = legend_y+(idx//per_row)*20
            self.canvas.create_line(lx, ly, lx+18, ly, fill=item["color"], width=3)
            self.canvas.create_text(lx+24, ly, anchor="w", text=item["char"], fill=C_GRAPH_TEXT, font=("Consolas",9))
        self.canvas.create_line(x1, y2, x2, y2, fill=C_GRAPH_AXIS, width=1)
        self.canvas.create_line(x1, y1, x1, y2, fill=C_GRAPH_AXIS, width=1)
        for i in range(5):
            yy = y2-((y2-y1)*i/4.0); val = max_y*i/4.0
            self.canvas.create_line(x1, yy, x2, yy, fill=C_GRAPH_GRID, width=1)
            self.canvas.create_text(x1-8, yy, anchor="e", text=f"{int(val)}", fill=C_GRAPH_TEXT, font=("Consolas",9))
        x_ticks = min(8, max(2, int(total_secs)+1))
        for i in range(x_ticks):
            ratio = i/(x_ticks-1 if x_ticks>1 else 1)
            xx = x1+((x2-x1)*ratio); val = total_secs*ratio
            self.canvas.create_line(xx, y2, xx, y2+4, fill=C_GRAPH_AXIS, width=1)
            self.canvas.create_text(xx, y2+14, anchor="n", text=f"{val:.0f}s", fill=C_GRAPH_TEXT, font=("Consolas",9))
        for item in series:
            coords = []
            for sec, dmg in item["points"]:
                px = x1+((sec/total_secs)*(x2-x1)); py = y2-((dmg/max_y)*(y2-y1))
                coords.extend([px, py])
            if len(coords) >= 4:
                self.canvas.create_line(*coords, fill=item["color"], width=2, smooth=False)
            for sec, dmg in item["points"]:
                px = x1+((sec/total_secs)*(x2-x1)); py = y2-((dmg/max_y)*(y2-y1))
                self.canvas.create_oval(px-3, py-3, px+3, py+3, fill=item["color"], outline=item["color"])
                self.canvas.create_text(px+6, py-8, anchor="w", text=str(int(dmg)),
                                        fill=item["color"], font=("Consolas",8,"bold"))

    def render(self):
        self._render_graph(); self._update_history_rows()

    def _update_loop(self):
        if self.winfo_exists():
            self.render(); self.after(1000, self._update_loop)


# ── Main Window ───────────────────────────────────────────────────────────────
class EVEUnifiedWindow(tk.Toplevel):
    # Incoming 패널: 3컬럼 (Inc.DPS(30s) / Inc.DPS / T.Inc.Dmg)
    INC_ROW_H = 24
    INC_HDR_H = 20

    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        apply_app_icon(self)
        self._dx = self._dy = 0
        self._cur_w = WINDOW_DEFAULT_W
        self._cur_h = 220
        self.cv = None
        self.tooltip = Tooltip(self)
        self.hover_char = None
        self.inline_editor = None
        self.inline_editor_char = None
        self.font_name  = ("Consolas", 10)
        self.font_num   = ("Consolas", 10)
        self.font_num_b = ("Consolas", 10, "bold")
        self.row_info = {}
        self.ui_items = {}
        self.header_click_boxes = {}
        self.inc_expanded = self.app.cfg.get("inc_expanded", True)
        self._drag_char = None; self._drag_y = None; self._drag_target = None
        self._setup_window()
        self.protocol("WM_DELETE_WINDOW", self.app.quit_all)
        self.bind("<Map>", self._on_map_restore)
        self._restore_pos()
        self._build_static_ui()
        self.render_all(force_layout=True)
        self.after(UI_UPDATE_MS, self._update_loop)
        self.after(1000, self._blink_tick)

    def _blink_tick(self):
        self.app.alarm_mgr.tick_blink(); self.after(1000, self._blink_tick)

    def _setup_window(self):
        self.overrideredirect(True); self.attributes("-topmost", True)
        self.config(bg=TRANSPARENT); self.attributes("-transparentcolor", TRANSPARENT)
        self.bind("<ButtonPress-1>",  self._drag_start)
        self.bind("<B1-Motion>",      self._drag_move)
        self.bind("<ButtonRelease-1>",self._drag_end)

    def _on_map_restore(self, _event):
        try:
            if self.state() == "normal":
                self.after(10, lambda: self.overrideredirect(True))
                self.after(10, lambda: self.attributes("-topmost", True))
                self.after(10, lambda: self.lift())
        except Exception: pass

    def _restore_pos(self):
        wc = self.app.cfg.get("window", {})
        self._cur_w = int(clamp(wc.get("w", WINDOW_DEFAULT_W), WINDOW_MIN_W, WINDOW_MAX_W))
        self.geometry(f"{self._cur_w}x{self._cur_h}+{wc.get('x',120)}+{wc.get('y',120)}")
        self.update_idletasks()
        self.attributes("-alpha", clamp(wc.get("alpha", 92), 10, 100)/100.0)

    def _save_pos(self):
        self.app.cfg["window"] = {
            "x": self.winfo_x(), "y": self.winfo_y(),
            "w": self._cur_w,
            "alpha": int(float(self.attributes("-alpha"))*100)}
        save_config(self.app.cfg)

    def _drag_start(self, e):
        if self.inline_editor_char or self._drag_char: return
        self._dx, self._dy = e.x, e.y; self.app.pause_updates = True

    def _drag_move(self, e):
        if self.inline_editor_char or self._drag_char: return
        self.geometry(f"+{self.winfo_x()+(e.x-self._dx)}+{self.winfo_y()+(e.y-self._dy)}")

    def _drag_end(self, _e):
        self._save_pos(); self.app.pause_updates = False

    def _calc_height(self, rows):
        n = max(1, rows)
        h = HEADER_H + n * ROW_H
        if self.inc_expanded:
            h += BOTTOM_PAD//2 + 16 + n * self.INC_ROW_H
        h += 16
        return h

    def _build_static_ui(self):
        for child in self.winfo_children(): child.destroy()
        self.cv = tk.Canvas(self, width=self._cur_w, height=self._cur_h,
                            bg=TRANSPARENT, highlightthickness=0)
        self.cv.pack()
        self.cv.bind("<Motion>",          self._on_motion)
        self.cv.bind("<Leave>",           lambda _e: self._on_leave())
        self.cv.bind("<Button-1>",        self._on_left_click)
        self.cv.bind("<Button-3>",        self._on_right_click)
        self.cv.bind("<ButtonRelease-1>", self._on_mouse_release)
        self.cv.bind("<B1-Motion>",       self._on_mouse_drag)

    def _clear_canvas(self):
        self.cv.delete("all"); self.row_info = {}; self.ui_items = {}; self.header_click_boxes = {}

    def _layout_columns(self):
        left, right = 10, self._cur_w - 10
        usable   = right - left
        name_col = int(usable * 0.40)
        dps_col  = int(usable * 0.20)
        tdps_col = int(usable * 0.20)
        return {
            "name_l":  left,
            "name_r":  left + name_col,
            "dps_r":   left + name_col + dps_col,
            "tdps_r":  left + name_col + dps_col + tdps_col,
            "tdam_r":  right,
        }

    def _measure_text(self, text, font):
        temp = self.cv.create_text(-9999, -9999, text=text, font=font, anchor="nw")
        bbox  = self.cv.bbox(temp); self.cv.delete(temp)
        return 0 if not bbox else bbox[2]-bbox[0]

    def render_all(self, force_layout=False):
        chars = self.app.sorted_characters()
        n = max(1, len(chars))
        target_h = self._calc_height(len(chars))
        if force_layout or target_h != self._cur_h:
            self._cur_h = target_h
            self.geometry(f"{self._cur_w}x{self._cur_h}+{self.winfo_x()}+{self.winfo_y()}")
            self.cv.config(width=self._cur_w, height=self._cur_h)
        self._clear_canvas()
        self.cv.create_polygon(rrect(1,1,self._cur_w-1,self._cur_h-1,RADIUS),
                               smooth=True, fill=C_BG, outline=C_BORDER, width=1.4)
        cols = self._layout_columns(); self.cols = cols

        # ── Header ──
        self.cv.create_text(10, 12, anchor="w", text=APP_TITLE,    fill=C_HEADER, font=("Consolas",11,"bold"))
        self.cv.create_text(83, 13, anchor="w", text=APP_SUBTITLE, fill=C_MUTED,  font=("Consolas",7))
        x_quit=self._cur_w-10; x_min=self._cur_w-28; x_refresh=self._cur_w-46
        x_graph=self._cur_w-64; x_show=self._cur_w-82; x_alarm=self._cur_w-100; x_sort=self._cur_w-115
        self.ui_items["quit"]     = self.cv.create_text(x_quit,    12, anchor="e", text="✕",  fill=C_ICON, font=("Segoe UI Symbol", 10))
        self.ui_items["min"]      = self.cv.create_text(x_min,     12, anchor="e", text="—",  fill=C_ICON, font=("Segoe UI Symbol", 11))
        self.ui_items["refresh"]  = self.cv.create_text(x_refresh, 12, anchor="e", text="↻",  fill=C_ICON, font=("Segoe UI Symbol", 13))
        self.ui_items["graph"]    = self.cv.create_text(x_graph,   12, anchor="e", text="📈", fill=C_ICON, font=("Segoe UI Symbol", 10))
        self.ui_items["show_all"] = self.cv.create_text(x_show,    12, anchor="e", text="👁", fill=C_ICON, font=("Segoe UI Symbol", 13))
        self.ui_items["alarm"]    = self.cv.create_text(x_alarm,   12, anchor="e", text="🔔", fill=C_ICON, font=("Segoe UI Symbol", 9))
        self.ui_items["sort"]     = self.cv.create_text(x_sort,    12, anchor="e", text="👤", fill=C_ICON, font=("Segoe UI Symbol", 8))
        def _btn_box(x): return (x-18, 0, x+2, 24)
        self.header_click_boxes["quit"]     = _btn_box(x_quit)
        self.header_click_boxes["min"]      = _btn_box(x_min)
        self.header_click_boxes["refresh"]  = _btn_box(x_refresh)
        self.header_click_boxes["graph"]    = _btn_box(x_graph)
        self.header_click_boxes["show_all"] = _btn_box(x_show)
        self.header_click_boxes["alarm"]    = _btn_box(x_alarm)
        self.header_click_boxes["sort"]     = _btn_box(x_sort)

        header_y = HEADER_H - 15
        header_font = ("Consolas",10,"bold")
        dps_c  = (cols["name_r"] + cols["dps_r"]) / 2
        tdps_c = (cols["dps_r"] + cols["tdps_r"]) / 2
        tdam_c = (cols["tdps_r"] + cols["tdam_r"]) / 2
        self.cv.create_text(cols["name_l"], header_y, anchor="w",      text="Name",  fill=C_HEADER, font=header_font)
        self.cv.create_text(dps_c,          header_y, anchor="center", text="DPS",   fill=C_HEADER, font=header_font)
        self.cv.create_text(tdps_c,         header_y, anchor="center", text="T.DPS", fill=C_HEADER, font=header_font)
        self.cv.create_text(tdam_c,         header_y, anchor="center", text="T.Dam", fill=C_HEADER, font=header_font)

        if not chars:
            self.cv.create_text(self._cur_w//2, HEADER_H+ROW_H//2, anchor="center",
                                text="Waiting for EVE logs...", fill=C_MUTED, font=("Consolas",10))
            return

        statuses = {char: self.app.get_status(char) for char in chars}
        in_combat_any = any(statuses[c]["in_combat"] for c in chars)
        bar_values = {c: statuses[c]["dps"] for c in chars} if in_combat_any \
                     else {c: statuses[c]["tdps"] for c in chars}
        max_bar = max((v for v in bar_values.values()), default=0.0)
        is_manual = self.app.cfg.get("sort_mode") == SORT_MANUAL

        for idx, char in enumerate(chars):
            y1 = HEADER_H + idx * ROW_H; y2 = y1 + ROW_H; yc = (y1+y2)//2
            hover = (char == self.hover_char)
            row_fill = "#15304a" if hover else C_BG
            self.cv.create_rectangle(6, y1, self._cur_w-6, y2, outline=C_BORDER, width=0.6, fill=row_fill)
            if max_bar > 0:
                ratio = bar_values[char]/max_bar; bar_w = int((self._cur_w-12)*ratio)
                bar_col = "#1a3a5a" if in_combat_any else "#1a3520"
                self.cv.create_rectangle(6, y1, 6+bar_w, y2, fill=bar_col, outline="", stipple="gray50")
            status   = statuses[char]
            cur_fill = C_DPS_ON if (status["in_combat"] and status["dps"]>0) else (C_DPS_OFF if status["dps"]>0 else C_IDLE)
            display  = self.app.display_name(char)
            icon_reserve = (42 if is_manual else 22) if hover else 1
            max_name_w = (cols["name_r"]-cols["name_l"])-icon_reserve
            fitted = fit_text_binary(display, max_name_w, lambda s: self._measure_text(s, self.font_name))
            self.cv.create_text(cols["name_l"],     yc, anchor="w", text=fitted,                  fill=C_TEXT,   font=self.font_name)
            self.cv.create_text(cols["dps_r"]  - 4, yc, anchor="e", text=fmt_num(status["dps"]),  fill=cur_fill, font=self.font_num_b)
            self.cv.create_text(cols["tdps_r"] - 4, yc, anchor="e", text=fmt_num(status["tdps"]), fill=C_TEXT,   font=self.font_num)
            self.cv.create_text(cols["tdam_r"] - 4, yc, anchor="e", text=fmt_int(status["tdam"]), fill=C_TEXT,   font=self.font_num)
            edit_box = hide_box = drag_box = None
            if hover:
                _BW = 14
                edit_x = cols["name_r"] - _BW; hide_x = cols["name_r"] + _BW//2
                _FONT_BTN = ("Segoe UI Symbol", 10)
                self.cv.create_text(edit_x, yc, anchor="center", text="✎", fill=C_ICON_H, font=_FONT_BTN)
                self.cv.create_text(hide_x, yc, anchor="center", text="⊘", fill=C_HIDE,   font=_FONT_BTN)
                edit_box = (edit_x-8, y1+2, edit_x+8, y2-2); hide_box = (hide_x-8, y1+2, hide_x+8, y2-2)
                if is_manual:
                    drag_x = cols["name_r"]+_BW*2
                    for dy_off in [-4, 0, 4]:
                        self.cv.create_line(drag_x-6, yc+dy_off, drag_x+6, yc+dy_off, fill=C_ICON_H, width=1.5)
                    drag_box = (drag_x-8, y1+2, drag_x+8, y2-2)
            if self._drag_target==char and self._drag_char and self._drag_char!=char:
                self.cv.create_line(6, y1, self._cur_w-6, y1, fill=C_HEADER, width=2)
            if self._drag_target=="__END__" and self._drag_char and idx==len(chars)-1:
                self.cv.create_line(6, y2, self._cur_w-6, y2, fill=C_HEADER, width=2)
            self.row_info[char] = {
                "row_box":  (6, y1, self._cur_w-6, y2),
                "name_box": (cols["name_l"], y1, cols["name_r"], y2),
                "edit_box": edit_box, "hide_box": hide_box, "drag_box": drag_box,
                "display":  display, "fitted": fitted,
                "y1": y1, "y2": y2, "edit_window_id": None,
            }
        if self.inline_editor_char and self.inline_editor_char in self.row_info:
            self._place_inline_editor(recreate_window=True)

        # ── Incoming 패널 (3컬럼) ────────────────────────────────────────────
        C_INC        = "#e06060"
        C_INC_MUTED  = "#a04040"
        alarm_cfg    = self.app.cfg.get("alarm", {})
        visual_alarm = alarm_cfg.get("enabled_visual", True)
        alarm_mgr    = self.app.alarm_mgr

        # Incoming 컬럼 레이아웃 (이름 40% / 각 20%)
        inc_name_r  = cols["name_r"]
        inc_col1_r  = cols["dps_r"]    # Inc.DPS(30s)
        inc_col2_r  = cols["tdps_r"]   # Inc.DPS
        inc_col3_r  = cols["tdam_r"]   # T.Inc.Dmg

        if self.inc_expanded and chars:
            inc_top = HEADER_H + n * ROW_H + BOTTOM_PAD // 2
            # 헤더
            self.cv.create_text(cols["name_l"], inc_top+8, anchor="w",
                                text="Incoming", fill=C_INC_MUTED, font=header_font)
            self.cv.create_text(dps_c,          inc_top+8, anchor="center",
                                text="DPS(30s)", fill=C_INC_MUTED, font=header_font)
            self.cv.create_text(tdps_c,         inc_top+8, anchor="center",
                                text="DPS",      fill=C_INC_MUTED, font=header_font)
            self.cv.create_text(tdam_c,         inc_top+8, anchor="center",
                                text="T.Dmg",    fill=C_INC_MUTED, font=header_font)

            for idx, char in enumerate(chars):
                ry1 = inc_top + 16 + idx * self.INC_ROW_H
                ry2 = ry1 + self.INC_ROW_H; ryc = (ry1+ry2)//2
                status = statuses[char]

                # 경보 체크: 30초 DPS + 마지막 피격 시각 전달
                alarm_mgr.check(char, status["inc_dps_30s"], status["last_inc_mono"])
                alarming = visual_alarm and alarm_mgr.is_alarming(char)
                blink_on = alarm_mgr.blink_state(char)

                if alarming and blink_on: row_bg="#3a0a0a"; txt_fill="#ffffff"
                elif alarming:            row_bg="#3a0a0a"; txt_fill="#ff4040"
                else:                     row_bg="#130a0a"; txt_fill=C_INC

                self.cv.create_rectangle(6, ry1, self._cur_w-6, ry2,
                                         outline="#2a1010", width=0.6, fill=row_bg)

                # Inc DPS(30s) 바
                inc_max = max((statuses[c]["inc_dps_30s"] for c in chars), default=0.0)
                if inc_max > 0:
                    ratio = status["inc_dps_30s"]/inc_max
                    inc_bar_w = int((self._cur_w-12)*ratio)
                    self.cv.create_rectangle(6, ry1, 6+inc_bar_w, ry2,
                                             fill="#3a0a0a", outline="", stipple="gray50")

                display    = self.app.display_name(char)
                max_name_w = cols["name_r"]-cols["name_l"]-1
                fitted     = fit_text_binary(display, max_name_w, lambda s: self._measure_text(s, self.font_name))
                self.cv.create_text(cols["name_l"],  ryc, anchor="w", text=fitted,                        fill="#c08080", font=self.font_name)
                self.cv.create_text(inc_col1_r - 4,  ryc, anchor="e", text=fmt_num(status["inc_dps_30s"]),fill=txt_fill,  font=self.font_num_b)
                self.cv.create_text(inc_col2_r - 4,  ryc, anchor="e", text=fmt_num(status["inc_dps"]),    fill=C_INC,    font=self.font_num)
                self.cv.create_text(inc_col3_r - 4,  ryc, anchor="e", text=fmt_int(status["inc_tdam"]),   fill=C_INC,    font=self.font_num)

        # ── Toggle button ────────────────────────────────────────────────────
        btn_w, btn_h = 60, 12; btn_x = (self._cur_w-btn_w)//2; btn_y = self._cur_h-btn_h-2
        sym = "▲ Inc" if self.inc_expanded else "▼ Inc"
        self.cv.create_rectangle(btn_x, btn_y, btn_x+btn_w, btn_y+btn_h, fill="#1a1010", outline="#5a2020", width=0.8)
        self.cv.create_text(btn_x+btn_w//2, btn_y+btn_h//2, anchor="center", text=sym, fill="#a04040", font=("Consolas",7))
        self.header_click_boxes["inc_toggle"] = (btn_x, btn_y, btn_x+btn_w, btn_y+btn_h)

    def _update_loop(self):
        if not self.app.pause_updates: self.render_all(force_layout=False)
        self.after(UI_UPDATE_MS, self._update_loop)

    def _on_right_click(self, e):
        menu = tk.Menu(self, tearoff=0, bg="#0f1923", fg="#90caf9",
                       activebackground="#1e3a5f", activeforeground="#fff",
                       font=("Consolas",9), bd=0, relief="flat")
        menu.add_command(label="Graph",    command=self.app.open_graph)
        menu.add_command(label="Minimize", command=self.app.minimize_all)
        menu.add_command(label="Reset",    command=self.app.reset_everything)
        menu.add_separator()
        menu.add_command(label="Exit",     command=self.app.quit_all)
        menu.tk_popup(e.x_root, e.y_root)

    def _hit_test_header(self, x, y):
        for key, box in self.header_click_boxes.items():
            x1,y1,x2,y2 = box
            if x1<=x<=x2 and y1<=y<=y2: return key
        return None

    def _hit_test_char(self, x, y):
        for char, info in self.row_info.items():
            x1,y1,x2,y2 = info["row_box"]
            if x1<=x<=x2 and y1<=y<=y2: return char
        return None

    def _on_left_click(self, e):
        header = self._hit_test_header(e.x, e.y)
        if header == "quit":       self.app.quit_all(); return
        if header == "min":        self.app.minimize_all(); return
        if header == "refresh":    self.app.reset_dps_only(); return
        if header == "show_all":   self.app.show_all_hidden(); return
        if header == "graph":      self.app.open_graph(); return
        if header == "alarm":      AlarmSettingsPopup(self, self.app); return
        if header == "sort":
            SortMenuPopup(self, self.app, self.winfo_rootx()+e.x, self.winfo_rooty()+e.y+10); return
        if header == "inc_toggle":
            self.inc_expanded = not self.inc_expanded
            self.app.cfg["inc_expanded"] = self.inc_expanded
            save_config(self.app.cfg)
            self.render_all(force_layout=True); return
        char = self._hit_test_char(e.x, e.y)
        if not char: return
        info = self.row_info.get(char)
        if not info: return
        if info.get("drag_box"):
            x1,y1,x2,y2 = info["drag_box"]
            if x1<=e.x<=x2 and y1<=e.y<=y2:
                self._drag_char = char; self._drag_y = e.y; self.app.pause_updates = True; return
        if info["edit_box"]:
            x1,y1,x2,y2 = info["edit_box"]
            if x1<=e.x<=x2 and y1<=e.y<=y2: self._start_inline_edit(char); return
        if info["hide_box"]:
            x1,y1,x2,y2 = info["hide_box"]
            if x1<=e.x<=x2 and y1<=e.y<=y2: self.app.hide_character(char); return

    def _on_mouse_drag(self, e):
        if not self._drag_char: return
        found=None; last_char=None; last_y2=0
        for char, info in self.row_info.items():
            y1,y2 = info["y1"],info["y2"]
            if y2>last_y2: last_y2=y2; last_char=char
            if y1<=e.y<=y2 and char!=self._drag_char: found=char; break
        if found is None and e.y>last_y2 and last_char and last_char!=self._drag_char:
            found="__END__"
        self._drag_target=found; self.render_all(force_layout=False)

    def _on_mouse_release(self, e):
        if self._drag_char and self._drag_target and self._drag_char!=self._drag_target:
            self._apply_manual_reorder(self._drag_char, self._drag_target)
        self._drag_char=None; self._drag_target=None
        self.app.pause_updates=False; self.render_all(force_layout=False)

    def _apply_manual_reorder(self, moving_char, target_char):
        chars = list(self.app.sorted_characters())
        if moving_char not in chars: return
        chars.remove(moving_char)
        if target_char=="__END__": chars.append(moving_char)
        else:
            if target_char not in chars: return
            chars.insert(chars.index(target_char), moving_char)
        self.app.cfg["manual_order"] = chars; save_config(self.app.cfg)

    def _on_motion(self, e):
        if self.inline_editor_char: return
        header = self._hit_test_header(e.x, e.y)
        if header:
            tooltips = {"show_all":"Show all","graph":"Recent battle","refresh":"Refresh",
                        "min":"Minimize","quit":"Quit","inc_toggle":"Toggle incoming",
                        "alarm":"Alarm settings","sort":"Sort order"}
            self.hover_char = None
            text = tooltips.get(header)
            if text: self.tooltip.show(e.x_root, e.y_root, text)
            else:    self.tooltip.hide()
            self.render_all(force_layout=False); return
        old = self.hover_char; self.hover_char = self._hit_test_char(e.x, e.y)
        if old != self.hover_char: self.tooltip.hide(); self.render_all(force_layout=False)
        if not self.hover_char: self.tooltip.hide(); return
        info = self.row_info.get(self.hover_char)
        if not info: self.tooltip.hide(); return
        for key in ("edit_box", "hide_box", "drag_box"):
            if info.get(key):
                x1,y1,x2,y2 = info[key]
                if x1<=e.x<=x2 and y1<=e.y<=y2:
                    tips = {"edit_box":"Set name","hide_box":"Hide","drag_box":"Drag to reorder"}
                    self.tooltip.show(e.x_root, e.y_root, tips[key]); return
        x1,y1,x2,y2 = info["name_box"]
        if x1<=e.x<=x2 and y1<=e.y<=y2 and info["fitted"]!=info["display"]:
            self.tooltip.show(e.x_root, e.y_root, info["display"])
        else:
            self.tooltip.hide()

    def _on_leave(self):
        if self.inline_editor_char: return
        changed = self.hover_char is not None
        self.hover_char = None; self.tooltip.hide()
        if changed: self.render_all(force_layout=False)

    def _start_inline_edit(self, char):
        self.app.pause_updates = True; self.hover_char = char
        self.render_all(force_layout=False)
        if self.inline_editor and self.inline_editor.winfo_exists():
            try: self.inline_editor.destroy()
            except Exception: pass
        self.inline_editor_char = char
        entry = tk.Entry(self.cv, bg=C_EDIT_BG, fg=C_EDIT_FG, insertbackground="#ffffff",
                         relief="flat", font=self.font_name)
        entry.insert(0, self.app.cfg.get("aliases", {}).get(char, ""))
        entry.focus_set(); entry.select_range(0, tk.END)
        entry.bind("<Return>",   lambda _e: self._commit_inline_edit())
        entry.bind("<Escape>",   lambda _e: self._cancel_inline_edit())
        entry.bind("<FocusOut>", lambda _e: self._commit_inline_edit())
        self.inline_editor = entry
        self._place_inline_editor(recreate_window=True)

    def _place_inline_editor(self, recreate_window=False):
        if not self.inline_editor_char or not self.inline_editor or not self.inline_editor.winfo_exists(): return
        info = self.row_info.get(self.inline_editor_char)
        if not info: return
        x = self.cols["name_l"]; y = info["y1"]+3
        w = (self.cols["name_r"]-self.cols["name_l"])-1; h = ROW_H-6
        if recreate_window or info.get("edit_window_id") is None:
            info["edit_window_id"] = self.cv.create_window(x, y, anchor="nw", window=self.inline_editor, width=w, height=h)
        else:
            self.cv.coords(info["edit_window_id"], x, y)
            self.cv.itemconfigure(info["edit_window_id"], width=w, height=h)

    def _commit_inline_edit(self):
        if not self.inline_editor_char or not self.inline_editor or not self.inline_editor.winfo_exists():
            self.app.pause_updates = False; return
        char = self.inline_editor_char; text = self.inline_editor.get().strip()
        self.app.set_alias(char, text)
        try: self.inline_editor.destroy()
        except Exception: pass
        self.inline_editor = None; self.inline_editor_char = None
        self.app.pause_updates = False; self.render_all(force_layout=False)

    def _cancel_inline_edit(self):
        if self.inline_editor and self.inline_editor.winfo_exists():
            try: self.inline_editor.destroy()
            except Exception: pass
        self.inline_editor = None; self.inline_editor_char = None
        self.app.pause_updates = False; self.render_all(force_layout=False)


# ── App ───────────────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__(); self.withdraw()
        apply_app_icon(self)
        self.cfg = load_config()
        self.hidden_chars   = set(self.cfg.get("hidden_chars", []))
        self.engines        = {}
        self.detected_chars = set()
        self.history_sessions = load_history_sessions()
        self.archived_records = merge_history_sessions(self.history_sessions)
        self.selected_history_label = HISTORY_CURRENT_LABEL
        self.pause_updates  = False
        self.graph_win      = None
        self.alarm_mgr      = AlarmManager(self.cfg)
        self._suite_control_seq = self._current_suite_control_seq()
        self.window         = EVEUnifiedWindow(self, self)
        self.protocol("WM_DELETE_WINDOW", self.quit_all)
        self._scan_running_windows()
        self._poll_suite_control()

    def sorted_characters(self):
        visible = [c for c in self.detected_chars if c not in self.hidden_chars]
        mode    = self.cfg.get("sort_mode", SORT_NAME_ASC)
        if mode == SORT_NAME_ASC:  return sorted(visible, key=natural_sort_key)
        if mode == SORT_NAME_DESC: return sorted(visible, key=natural_sort_key, reverse=True)
        if mode in (SORT_TOP_DPS, SORT_BOT_DPS):
            def dps_key(char):
                st = self.get_or_create_engine(char).get_status()
                return st["dps"] if st["in_combat"] else st["tdps"]
            return sorted(visible, key=dps_key, reverse=(mode==SORT_TOP_DPS))
        if mode == SORT_MANUAL:
            order   = self.cfg.get("manual_order", [])
            order_map = {c: i for i, c in enumerate(order)}
            known   = [c for c in order if c in visible]
            unknown = sorted([c for c in visible if c not in order_map], key=natural_sort_key)
            return known + unknown
        return sorted(visible, key=natural_sort_key)

    def history_option_labels(self):
        return [HISTORY_CURRENT_LABEL] + [s["label"] for s in self.history_sessions]

    def selected_history_records(self):
        if self.selected_history_label == HISTORY_CURRENT_LABEL:
            return {}
        for session in self.history_sessions:
            if session.get("label") == self.selected_history_label:
                return session.get("records", {})
        self.selected_history_label = HISTORY_CURRENT_LABEL
        return {}

    def history_characters(self):
        if self.selected_history_label == HISTORY_CURRENT_LABEL:
            chars = set(self.detected_chars)
        else:
            chars = set(self.selected_history_records().keys())
        chars = chars - self.hidden_chars
        mode = self.cfg.get("sort_mode", SORT_NAME_ASC)
        if mode == SORT_NAME_DESC:
            return sorted(chars, key=natural_sort_key, reverse=True)
        if mode == SORT_MANUAL:
            order = self.cfg.get("manual_order", [])
            order_map = {c: i for i, c in enumerate(order)}
            known = [c for c in order if c in chars]
            unknown = sorted([c for c in chars if c not in order_map], key=natural_sort_key)
            return known + unknown
        return sorted(chars, key=natural_sort_key)

    def display_name(self, char):
        alias = self.cfg.get("aliases", {}).get(char, "").strip()
        return alias if alias else char

    def get_or_create_engine(self, char):
        eng = self.engines.get(char)
        if eng is None:
            eng = DPSEngine(char); eng.start(); self.engines[char] = eng
        return eng

    def get_status(self, char):
        return self.get_or_create_engine(char).get_status()

    def set_alias(self, char, alias_text):
        aliases = self.cfg.setdefault("aliases", {})
        if alias_text: aliases[char] = alias_text
        else:          aliases.pop(char, None)
        save_config(self.cfg)

    def _save_hidden(self):
        self.cfg["hidden_chars"] = sorted(self.hidden_chars, key=natural_sort_key)
        save_config(self.cfg)

    def hide_character(self, char):
        self.hidden_chars.add(char); self._save_hidden()
        self.window.hover_char = None; self.window.tooltip.hide()
        self.window.render_all(force_layout=True)

    def show_all_hidden(self):
        self.hidden_chars.clear(); self._save_hidden()
        self.window.render_all(force_layout=True)

    def reset_dps_only(self):
        for eng in list(self.engines.values()):
            try: eng.hard_reset()
            except Exception: pass
        self.alarm_mgr.stop_all()
        if self.graph_win and self.graph_win.winfo_exists(): self.graph_win.render()
        self.window.render_all(force_layout=False)

    def reset_everything(self):
        if not messagebox.askyesno("Reset", "Reset everything to first-run state?"): return
        for eng in list(self.engines.values()):
            try: eng.stop()
            except Exception: pass
        self.engines.clear(); self.detected_chars.clear(); self.hidden_chars.clear()
        self.alarm_mgr.stop_all()
        self.cfg = default_config(); save_config(self.cfg)
        self.window._cur_w = WINDOW_DEFAULT_W
        self.window.geometry(f"{WINDOW_DEFAULT_W}x220+120+120")
        self.window.attributes("-alpha", 0.92)
        self.window._save_pos(); self.window.hover_char=None; self.window.tooltip.hide()
        self.window.render_all(force_layout=True)
        if self.graph_win and self.graph_win.winfo_exists():
            self.graph_win.destroy(); self.graph_win = None

    def minimize_all(self):
        try: self.window._save_pos()
        except Exception: pass
        try: self.window.overrideredirect(False)
        except Exception: pass
        try: self.window.iconify()
        except Exception: pass

    def restore_from_suite(self):
        try:
            self.window.deiconify()
            self.window.overrideredirect(True)
            self.window.attributes("-topmost", True)
            self.window.lift()
        except Exception:
            pass

    def _current_suite_control_seq(self):
        try:
            if os.path.exists(CONTROL_FILE):
                with open(CONTROL_FILE, "r", encoding="utf-8") as f:
                    return json.load(f).get("seq")
        except Exception:
            pass
        return None

    def _poll_suite_control(self):
        try:
            if os.path.exists(CONTROL_FILE):
                with open(CONTROL_FILE, "r", encoding="utf-8") as f:
                    cmd = json.load(f)
                seq = cmd.get("seq")
                if seq != self._suite_control_seq:
                    self._suite_control_seq = seq
                    action = cmd.get("action")
                    if action == "minimize":
                        self.minimize_all()
                    elif action == "restore":
                        self.restore_from_suite()
        except Exception:
            pass
        self.after(250, self._poll_suite_control)

    def open_graph(self):
        if self.graph_win and self.graph_win.winfo_exists():
            self.graph_win.lift(); self.graph_win.focus_force(); self.graph_win.render(); return
        self.graph_win = GraphWindow(self, self)

    def _archive_records(self, records_by_char):
        clean = {}
        for char, records in (records_by_char or {}).items():
            rows = [
                rec for rec in (records or [])
                if rec.get("total_dmg", 0) > 0 or rec.get("inc_total_dmg", 0) > 0
            ]
            if rows:
                clean[char] = rows
        if not clean:
            return None
        path = save_history_archive(clean)
        if path:
            self.history_sessions = load_history_sessions()
            self.archived_records = merge_history_sessions(self.history_sessions)
            if self.graph_win and self.graph_win.winfo_exists():
                self.graph_win.refresh_history_selector()
                self.graph_win.render()
        return path

    def _archive_current_engines(self):
        records = {}
        for char, eng in list(self.engines.items()):
            try:
                rows = eng.get_battle_records()
                if rows:
                    records[char] = rows
            except Exception:
                pass
        return self._archive_records(records)

    def _scan_running_windows(self):
        if not self.pause_updates:
            current = list_running_eve_characters_from_windows()
            self.detected_chars = set(current)
            for char in current: self.get_or_create_engine(char)
            for char in list(self.engines.keys()):
                if char not in current:
                    eng = self.engines.pop(char, None)
                    if eng:
                        try: self._archive_records({char: eng.get_battle_records()})
                        except Exception: pass
                        eng.stop()
            self.window.render_all(force_layout=True)
        self.after(WINDOW_SCAN_INTERVAL_MS, self._scan_running_windows)

    def quit_all(self):
        if self.graph_win and self.graph_win.winfo_exists():
            try:
                self.cfg["graph_window"] = {
                    "x": self.graph_win.winfo_x(), "y": self.graph_win.winfo_y(),
                "w": self.graph_win._win_w,    "h": self.graph_win._win_h}
            except Exception: pass
            try: self.graph_win.destroy()
            except Exception: pass
        self._archive_current_engines()
        self.alarm_mgr.stop_all()
        for eng in list(self.engines.values()):
            try: eng.stop()
            except Exception: pass
        self.engines.clear()
        try: self.window._save_pos()
        except Exception: pass
        self.destroy()


if __name__ == "__main__":
    App().mainloop()
