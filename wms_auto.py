"""
WMS Automation Tool v5.0
========================
Чистый Win32 API — мгновенное определение окон, надёжные нажатия.

Цикл (скрины 1→12):
  1. Главное меню              → F2
  2. Взятие работы             → F2
  3→4. Перемещение к источнику → Место → Контроль → Ок
  5→6. Поиск паллеты           → Паллета → Контроль → Ок
  7→8. Поиск коробки           → Коробка → Контроль → Ок
  9→10. Поиск места (коробка)  → Зона из XLS → Контроль → Ок (ошибка → D-KM-1)
  11. Размещение в место       → Ок
  12. Поиск места (паллета)    → ВСЕГДА D-KM-1 → Ок
  → цикл заново

pip install openpyxl keyboard
Запуск от имени администратора!
"""

import time, sys, os, re, glob, logging, threading, ctypes, subprocess
from ctypes import wintypes
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
from datetime import datetime

try:
    import openpyxl
except ImportError:
    print("pip install openpyxl"); sys.exit(1)

try:
    import keyboard
except ImportError:
    print("pip install keyboard"); sys.exit(1)

# ───────────────────────────────────────────────────────────────────
#  Win32 API
# ───────────────────────────────────────────────────────────────────
user32  = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

WM_GETTEXT       = 0x000D
WM_GETTEXTLENGTH = 0x000E
WM_SETTEXT       = 0x000C
WM_COMMAND       = 0x0111
WM_KEYDOWN       = 0x0100
WM_KEYUP         = 0x0101
WM_LBUTTONDOWN   = 0x0201
WM_LBUTTONUP     = 0x0202
BM_CLICK         = 0x00F5
VK_RETURN        = 0x0D
VK_F2            = 0x71

WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)


def _get_text(hwnd):
    """Прочитать текст контрола через Win32 API (кросс-процесс)."""
    try:
        ln = user32.SendMessageW(hwnd, WM_GETTEXTLENGTH, 0, 0)
        if ln <= 0:
            ln = 512
        buf = ctypes.create_unicode_buffer(ln + 1)
        user32.SendMessageW(hwnd, WM_GETTEXT, ln + 1, ctypes.addressof(buf))
        return buf.value
    except:
        return ""


def _set_text(hwnd, text):
    """Записать текст в Edit контрол."""
    buf = ctypes.create_unicode_buffer(text)
    user32.SendMessageW(hwnd, WM_SETTEXT, 0, ctypes.addressof(buf))


def _get_class(hwnd):
    buf = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, buf, 256)
    return buf.value


def _get_rect(hwnd):
    r = wintypes.RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(r))
    return r


def _get_pid(hwnd):
    pid = wintypes.DWORD()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return pid.value


def _is_visible(hwnd):
    return bool(user32.IsWindowVisible(hwnd))


# ── Перечисление окон ──

def _enum_top():
    """Все окна верхнего уровня — МГНОВЕННО."""
    res = []
    def cb(h, _):
        res.append(h)
        return True
    user32.EnumWindows(WNDENUMPROC(cb), 0)
    return res


def _enum_children(parent):
    res = []
    def cb(h, _):
        res.append(h)
        return True
    user32.EnumChildWindows(parent, WNDENUMPROC(cb), 0)
    return res


def _find_hwnd(title_part, pid=None):
    """Найти окно по части заголовка. Быстро — через EnumWindows."""
    tl = title_part.lower()
    for h in _enum_top():
        if not _is_visible(h):
            continue
        t = _get_text(h)
        if t and tl in t.lower():
            if pid and _get_pid(h) != pid:
                continue
            return h
    return None


def _children_by_class(parent, cls):
    """Дочерние контролы по части имени класса."""
    cl = cls.lower()
    return [h for h in _enum_children(parent) if cl in _get_class(h).lower()]


# ── Работа с контролами ──

def _get_edits(hwnd):
    return _children_by_class(hwnd, "Edit")


def _get_statics_text(hwnd):
    return [_get_text(s) for s in _children_by_class(hwnd, "Static") if _get_text(s)]


def _edit_near_label(parent, label_part):
    """Найти Edit рядом с меткой (±40px по вертикали)."""
    statics = _children_by_class(parent, "Static")
    edits   = _children_by_class(parent, "Edit")
    ll = label_part.lower()

    lbl = None
    for s in statics:
        if ll in _get_text(s).lower():
            lbl = s
            break
    if not lbl:
        return None

    lr = _get_rect(lbl)
    lcy = (lr.top + lr.bottom) // 2

    best, best_d = None, 9999
    for e in edits:
        er = _get_rect(e)
        ecy = (er.top + er.bottom) // 2
        d = abs(ecy - lcy)
        if d < 40 and d < best_d:
            best_d = d
            best = e
    return best


def _find_empty_edit(parent):
    """Первый пустой Edit."""
    for e in _get_edits(parent):
        if not _get_text(e).strip():
            return e
    return None


def _last_filled_edit(parent):
    """Последний заполненный Edit."""
    edits = _get_edits(parent)
    for e in reversed(edits):
        v = _get_text(e).strip()
        if v:
            return e, v
    return None, ""


# ── Нажатия ──

def _click_btn(btn_hwnd):
    """Нажать кнопку через BM_CLICK + WM_COMMAND."""
    user32.SendMessageW(btn_hwnd, BM_CLICK, 0, 0)


def _click_btn_text(parent, *texts):
    """Нажать кнопку по тексту."""
    for b in _children_by_class(parent, "Button"):
        bt = _get_text(b).strip().lower()
        for t in texts:
            if t.lower() in bt:
                _click_btn(b)
                return True
    return False


def _press_ok(hwnd):
    """Нажать Ок."""
    if _click_btn_text(hwnd, "ок", "ok", "оk", "oк"):
        return True
    # Fallback: Enter
    user32.PostMessageW(hwnd, WM_KEYDOWN, VK_RETURN, 0)
    time.sleep(0.03)
    user32.PostMessageW(hwnd, WM_KEYUP, VK_RETURN, 0)
    return True


def _press_f2(hwnd):
    """Нажать F2 / кнопку с F2."""
    if _click_btn_text(hwnd, "f2", "запросить", "взять"):
        return True
    user32.SetForegroundWindow(hwnd)
    time.sleep(0.1)
    user32.PostMessageW(hwnd, WM_KEYDOWN, VK_F2, 0)
    time.sleep(0.03)
    user32.PostMessageW(hwnd, WM_KEYUP, VK_F2, 0)
    return True


def _activate(hwnd):
    user32.SetForegroundWindow(hwnd)


# ── Ожидание ──

def _wait_gone(hwnd, timeout=3.0):
    """Ждать пока окно исчезнет."""
    t0 = time.time()
    while time.time() - t0 < timeout:
        if not _is_visible(hwnd):
            return True
        time.sleep(0.15)
    return False


def _proc_name(pid):
    try:
        r = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV", "/NH"],
            capture_output=True, text=True, timeout=3)
        for line in r.stdout.strip().split("\n"):
            p = line.replace('"', '').split(",")
            if p:
                return p[0].strip()
    except:
        pass
    return "?"


# ───────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("wms_auto.log", encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger("WMS")

FALLBACK = "D-KM-1"

# UI
BG      = "#1a1a2e"
BG2     = "#16213e"
BGI     = "#1c2a4a"
FG      = "#e0e0e0"
FGD     = "#7a7a9a"
BLUE    = "#2563eb"
GREEN   = "#16a34a"
GREENL  = "#22c55e"
RED     = "#dc2626"
REDL    = "#ef4444"
YELLOW  = "#eab308"
CYAN    = "#06b6d4"
ORANGE  = "#ea580c"
BAR     = "#111827"

# Заголовки WMS для автопоиска
WMS_TITLES = [
    "главное меню", "взятие работы", "перемещение к источнику",
    "поиск паллеты", "поиск палеты", "поиск коробки",
    "поиск места-приёмника", "поиск места-приемника",
    "размещение в место",
]


# ═══════════════════════════════════════════════════════════════════
#  ZoneLookup
# ═══════════════════════════════════════════════════════════════════
class ZoneLookup:
    def __init__(self):
        self.m = {}
        self.d = {}
        self.path = None

    def load(self, path):
        if not path or not os.path.exists(path):
            return False
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            ws = wb.active
            self.m.clear(); self.d.clear()
            for row in ws.iter_rows(min_row=1, values_only=True):
                if row and len(row) >= 2 and row[0] and row[1]:
                    k, v = str(row[0]).strip(), str(row[1]).strip()
                    self.m[k.lower()] = v
                    self.d[k] = v
            wb.close()
            self.path = path
            return True
        except:
            return False

    def lookup(self, zone):
        z = zone.lower().strip()
        if z in self.m: return self.m[z]
        for k, v in self.m.items():
            if k in z: return v
        for k, v in self.m.items():
            if z in k: return v
        return None

    def auto(self):
        d = os.path.dirname(os.path.abspath(__file__))
        for n in ["zones.xlsx", "зоны.xlsx", "зона.xlsx", "zone.xlsx"]:
            p = os.path.join(d, n)
            if os.path.exists(p) and self.load(p): return p
        for f in sorted(glob.glob(os.path.join(d, "*.xlsx"))):
            if self.load(f): return f
        return None


# ═══════════════════════════════════════════════════════════════════
#  WMS Bot — ЯДРО
# ═══════════════════════════════════════════════════════════════════
class WMSBot:

    def __init__(self, zones, fallback, pid=None, log_cb=None, status_cb=None):
        self.zones    = zones
        self.fallback = fallback or FALLBACK
        self.pid      = pid
        self.log_cb   = log_cb
        self.status_cb = status_cb

        self.running = False
        self.paused  = False
        self.poll    = 0.4
        self._last_h = None
        self._last_t = 0

    def log(self, msg, lvl="INFO"):
        logger.log(getattr(logging, lvl, 20), msg)
        if self.log_cb:
            self.log_cb(f"[{lvl}] {msg}")

    # ─── Определение окна ─────────────────────────────────────────
    def detect(self):
        """Что сейчас на экране? Возвращает (тип, hwnd)."""
        pid = self.pid

        # Модальные — приоритет
        for t in ("Ошибка", "Error", "Предупреждение", "Warning"):
            h = _find_hwnd(t, pid)
            if h: return "error", h

        for t in ("Подтверждение", "Подтвердить", "Confirm",
                   "Информация", "Information", "Сообщение", "Message"):
            h = _find_hwnd(t, pid)
            if h: return "confirm", h

        # Размещение в место (скрин 11)
        h = _find_hwnd("Размещение в место", pid)
        if h: return "place", h

        # Поиск места-приёмника (скрин 9/10/12)
        h = _find_hwnd("Поиск места-при", pid)
        if not h: h = _find_hwnd("Поиск места-приёмника", pid)
        if not h: h = _find_hwnd("Поиск места-приемника", pid)
        if h:
            txt = " ".join(_get_statics_text(h)).lower()
            if "паллет" in txt or "содержимого" in txt:
                return "dest_pallet", h    # скрин 12
            return "dest_box", h           # скрин 9

        # Поиск коробки (скрин 7)
        h = _find_hwnd("Поиск коробки", pid)
        if h: return "box", h

        # Поиск паллеты (скрин 5)
        h = _find_hwnd("Поиск паллеты", pid)
        if not h: h = _find_hwnd("Поиск палеты", pid)
        if h: return "pallet", h

        # Перемещение к источнику (скрин 3)
        h = _find_hwnd("Перемещение к источнику", pid)
        if h: return "source", h

        # Взятие работы (скрин 2)
        h = _find_hwnd("Взятие работы", pid)
        if h: return "take", h

        # Главное меню (скрин 1)
        h = _find_hwnd("Главное меню", pid)
        if h: return "menu", h

        return None, None

    # ─── Защита от повторной обработки ────────────────────────────
    def _should_skip(self, hwnd):
        """Пропустить если только что обработали это же окно."""
        if hwnd == self._last_h and time.time() - self._last_t < 1.5:
            return True
        return False

    def _mark_done(self, hwnd):
        self._last_h = hwnd
        self._last_t = time.time()

    # ─── Обработчики каждого шага ─────────────────────────────────

    def step_menu(self, h):
        """Скрин 1: Главное меню → F2"""
        self.log("→ [1] Главное меню → F2")
        _activate(h)
        time.sleep(0.15)
        _press_f2(h)
        self._mark_done(h)
        time.sleep(0.6)

    def step_take(self, h):
        """Скрин 2: Взятие работы → F2"""
        self.log("→ [2] Взятие работы → F2")
        _activate(h)
        time.sleep(0.15)
        _press_f2(h)
        self._mark_done(h)
        time.sleep(0.6)

    def step_source(self, h):
        """Скрин 3-4: Перемещение к источнику → Место → Контроль → Ок"""
        self.log("→ [3] Перемещение к источнику")

        e_mesto = _edit_near_label(h, "Место")
        e_ctrl  = _edit_near_label(h, "Контроль")

        if not e_mesto or not e_ctrl:
            # Фоллбэк: первый Edit = Место, последний пустой = Контроль
            edits = _get_edits(h)
            vals = [(e, _get_text(e).strip()) for e in edits]
            self.log(f"   label-поиск не удался, edits={[v for _,v in vals]}")
            for e, v in vals:
                if v and not e_mesto: e_mesto = e
            e_ctrl = _find_empty_edit(h)

        if e_mesto and e_ctrl:
            val = _get_text(e_mesto).strip()
            cv  = _get_text(e_ctrl).strip()
            if val and not cv:
                self.log(f"   Место='{val}' → Контроль")
                _set_text(e_ctrl, val)
                time.sleep(0.15)
            elif val and cv:
                self.log(f"   Уже заполнено='{cv}'")
            _press_ok(h)
            self._mark_done(h)
            time.sleep(0.6)
        else:
            self.log("   ⚠ Не нашёл поля!", "WARNING")

    def step_pallet(self, h):
        """Скрин 5-6: Поиск паллеты → Паллета → Контроль → Ок"""
        self.log("→ [5] Поиск паллеты")

        e_pal  = _edit_near_label(h, "Паллет")
        e_ctrl = _edit_near_label(h, "Контроль")

        if not e_pal or not e_ctrl:
            edits = _get_edits(h)
            vals = [(e, _get_text(e).strip()) for e in edits]
            self.log(f"   label-поиск не удался, edits={[v for _,v in vals]}")
            # Паллета — предпоследний заполненный, Контроль — последний пустой
            e_ctrl = _find_empty_edit(h)
            for e, v in reversed(vals):
                if v and e != e_ctrl:
                    e_pal = e
                    break

        if e_pal and e_ctrl:
            val = _get_text(e_pal).strip()
            cv  = _get_text(e_ctrl).strip()
            if val and not cv:
                self.log(f"   Паллета='{val}' → Контроль")
                _set_text(e_ctrl, val)
                time.sleep(0.15)
            elif val and cv:
                self.log(f"   Уже заполнено='{cv}'")
            _press_ok(h)
            self._mark_done(h)
            time.sleep(0.6)
        else:
            self.log("   ⚠ Не нашёл поля!", "WARNING")

    def step_box(self, h):
        """Скрин 7-8: Поиск коробки → Коробка → Контроль → Ок"""
        self.log("→ [7] Поиск коробки")

        e_box  = _edit_near_label(h, "Коробк")
        e_ctrl = _edit_near_label(h, "Контроль")

        if not e_box or not e_ctrl:
            edits = _get_edits(h)
            vals = [(e, _get_text(e).strip()) for e in edits]
            self.log(f"   label-поиск не удался, edits={[v for _,v in vals]}")
            e_ctrl = _find_empty_edit(h)
            for e, v in reversed(vals):
                if v and e != e_ctrl:
                    e_box = e
                    break

        if e_box and e_ctrl:
            val = _get_text(e_box).strip()
            cv  = _get_text(e_ctrl).strip()
            if val and not cv:
                self.log(f"   Коробка='{val}' → Контроль")
                _set_text(e_ctrl, val)
                time.sleep(0.15)
            elif val and cv:
                self.log(f"   Уже заполнено='{cv}'")
            _press_ok(h)
            self._mark_done(h)
            time.sleep(0.6)
        else:
            self.log("   ⚠ Не нашёл поля!", "WARNING")

    def step_dest_box(self, h):
        """Скрин 9-10: Поиск места-приёмника (коробка) → Зона → Контроль → Ок"""
        self.log("→ [9] Поиск места-приёмника (коробка)")

        all_txt = " ".join(_get_statics_text(h))
        self.log(f"   Текст: {all_txt}")

        # Извлекаем зону
        zone = ""
        m = re.search(r"[Зз]она[:\s]+(.+)", all_txt)
        if m:
            zone = m.group(1).strip().split("\n")[0].strip()
        self.log(f"   Зона: '{zone}'")

        # Ищем код
        loc = None
        if zone:
            loc = self.zones.lookup(zone)
        if loc:
            self.log(f"   XLS → '{loc}'")
        else:
            loc = self.fallback
            self.log(f"   Не найдено → fallback '{loc}'")

        e_ctrl = _edit_near_label(h, "Контроль")
        if not e_ctrl:
            e_ctrl = _find_empty_edit(h)

        if e_ctrl:
            _set_text(e_ctrl, loc)
            time.sleep(0.15)
            _press_ok(h)
            self._mark_done(h)

            # Ждём результат
            time.sleep(1.0)
            typ, h2 = self.detect()
            if typ == "error":
                self.log(f"   ❌ Ошибка! Закрываю → fallback {self.fallback}")
                _press_ok(h2)
                time.sleep(0.6)
                typ2, h3 = self.detect()
                if typ2 in ("dest_box", "dest_pallet"):
                    ec = _edit_near_label(h3, "Контроль")
                    if not ec:
                        ec = _find_empty_edit(h3)
                    if ec:
                        _set_text(ec, self.fallback)
                        time.sleep(0.15)
                        _press_ok(h3)
                        self._mark_done(h3)
                        time.sleep(0.6)
            elif typ == "confirm":
                _press_ok(h2)
                time.sleep(0.4)
        else:
            self.log("   ⚠ Контроль не найден!", "WARNING")

    def step_dest_pallet(self, h):
        """Скрин 12: Поиск места (паллета) → ВСЕГДА D-KM-1 → Ок"""
        self.log(f"→ [12] Поиск места-приёмника (паллета) → {self.fallback}")

        e_ctrl = _edit_near_label(h, "Контроль")
        if not e_ctrl:
            e_ctrl = _find_empty_edit(h)

        if e_ctrl:
            _set_text(e_ctrl, self.fallback)
            time.sleep(0.15)
            _press_ok(h)
            self._mark_done(h)
            time.sleep(0.6)
        else:
            self.log("   ⚠ Контроль не найден!", "WARNING")

    def step_place(self, h):
        """Скрин 11: Размещение в место → Ок"""
        self.log("→ [11] Размещение в место → Ок")
        _press_ok(h)
        self._mark_done(h)
        time.sleep(0.6)

    def step_error(self, h):
        txt = " ".join(_get_statics_text(h))
        self.log(f"→ Ошибка: {txt}")
        _press_ok(h)
        self._mark_done(h)
        time.sleep(0.5)

    def step_confirm(self, h):
        self.log("→ Подтверждение → Ок")
        _press_ok(h)
        self._mark_done(h)
        time.sleep(0.4)
        # Может быть цепочка
        t2, h2 = self.detect()
        if t2 == "confirm":
            _press_ok(h2)
            time.sleep(0.4)

    # ─── Главный цикл ────────────────────────────────────────────
    def run(self):
        self.running = True
        self.log("=" * 50)
        self.log("🚀 ЗАПУЩЕНО")
        self.log(f"   Fallback={self.fallback}  Зон={len(self.zones.m)}  PID={self.pid or 'все'}")
        self.log("=" * 50)

        H = {
            "menu":        self.step_menu,
            "take":        self.step_take,
            "source":      self.step_source,
            "pallet":      self.step_pallet,
            "box":         self.step_box,
            "dest_box":    self.step_dest_box,
            "dest_pallet": self.step_dest_pallet,
            "place":       self.step_place,
            "error":       self.step_error,
            "confirm":     self.step_confirm,
        }

        while self.running:
            if self.paused:
                if self.status_cb: self.status_cb("paused")
                time.sleep(0.3)
                continue

            try:
                if self.status_cb: self.status_cb("running")

                typ, hwnd = self.detect()

                if typ is None:
                    time.sleep(self.poll)
                    continue

                # Защита от повтора
                if self._should_skip(hwnd):
                    time.sleep(0.3)
                    continue

                fn = H.get(typ)
                if fn:
                    self.log(f"── Обнаружено: {typ} (hwnd={hwnd}) ──")
                    fn(hwnd)
                else:
                    time.sleep(self.poll)

            except Exception as e:
                self.log(f"Ошибка: {e}", "ERROR")
                time.sleep(1)

        if self.status_cb: self.status_cb("stopped")
        self.log("⏹ ОСТАНОВЛЕНО")

    def stop(self):
        self.running = False

    def toggle_pause(self):
        self.paused = not self.paused
        self.log("⏸ ПАУЗА" if self.paused else "▶ ПРОДОЛЖЕНИЕ")


# ═══════════════════════════════════════════════════════════════════
#  Сканер процессов
# ═══════════════════════════════════════════════════════════════════
def scan_wms():
    found = {}
    for h in _enum_top():
        if not _is_visible(h): continue
        t = _get_text(h).lower()
        for k in WMS_TITLES:
            if k in t:
                pid = _get_pid(h)
                if pid and pid not in found:
                    found[pid] = (pid, _proc_name(pid), _get_text(h))
                break
    return list(found.values())

def scan_all():
    found = {}
    for h in _enum_top():
        if not _is_visible(h): continue
        t = _get_text(h)
        if not t.strip(): continue
        pid = _get_pid(h)
        if pid and pid > 0 and pid not in found:
            found[pid] = (pid, _proc_name(pid), t)
    return sorted(found.values(), key=lambda x: x[1].lower())


# ═══════════════════════════════════════════════════════════════════
#  GUI
# ═══════════════════════════════════════════════════════════════════
class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("WMS Автоматизация v5.0")
        self.root.geometry("960x750")
        self.root.configure(bg=BG)
        self.root.resizable(True, True)

        self.bot    = None
        self.thread = None
        self.zones  = ZoneLookup()
        self._pid   = None
        self._procs = []
        self._ax    = -100
        self._anim  = False
        self._mode  = "stopped"

        self._build()
        self.root.after(200, self._auto_zones)
        self.root.after(400, self._auto_scan)

        # Пауза по клавише
        self._kb = keyboard.on_press(self._on_key)

    def _build(self):
        # ── Status bar ──
        self.cv = tk.Canvas(self.root, height=52, bg=BLUE, highlightthickness=0)
        self.cv.pack(fill=tk.X)
        self.cv.bind("<Configure>", lambda e: self._bar())

        m = tk.Frame(self.root, bg=BG)
        m.pack(fill=tk.BOTH, expand=True, padx=12, pady=6)

        # ── Приложение ──
        f1 = tk.LabelFrame(m, text="  🖥  WMS-приложение  ",
                            bg=BG2, fg=CYAN, font=("Segoe UI", 10, "bold"))
        f1.pack(fill=tk.X, pady=(0,5))
        r = tk.Frame(f1, bg=BG2); r.pack(fill=tk.X, padx=8, pady=5)
        tk.Label(r, text="Процесс:", bg=BG2, fg=FG, font=("Segoe UI", 10)).pack(side=tk.LEFT)
        self.va = tk.StringVar(value="(сканирование...)")
        self.cb = ttk.Combobox(r, textvariable=self.va, width=55,
                                state="readonly", font=("Segoe UI", 9))
        self.cb.pack(side=tk.LEFT, padx=8)
        self.cb.bind("<<ComboboxSelected>>", self._on_cb)
        tk.Button(r, text="🔍 WMS", command=self._sc_wms,
                  bg=BLUE, fg="white", relief="flat", font=("Segoe UI", 9, "bold"),
                  padx=10, cursor="hand2").pack(side=tk.LEFT, padx=3)
        tk.Button(r, text="📋 Все", command=self._sc_all,
                  bg="#374151", fg=FG, relief="flat", font=("Segoe UI", 9),
                  padx=10, cursor="hand2").pack(side=tk.LEFT, padx=3)
        self.la = tk.Label(f1, text="", bg=BG2, fg=FGD, font=("Segoe UI", 9))
        self.la.pack(anchor=tk.W, padx=12, pady=(0,4))

        # ── Настройки ──
        f2 = tk.LabelFrame(m, text="  ⚙  Настройки  ",
                            bg=BG2, fg=CYAN, font=("Segoe UI", 10, "bold"))
        f2.pack(fill=tk.X, pady=(0,5))
        r1 = tk.Frame(f2, bg=BG2); r1.pack(fill=tk.X, padx=8, pady=4)
        tk.Label(r1, text="📁 Файл зон:", bg=BG2, fg=FG,
                 font=("Segoe UI", 10)).pack(side=tk.LEFT)
        self.vx = tk.StringVar()
        tk.Entry(r1, textvariable=self.vx, width=48, bg=BGI, fg=FG,
                 insertbackground=FG, relief="flat", font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=8)
        tk.Button(r1, text="Обзор…", command=self._browse,
                  bg=BLUE, fg="white", relief="flat", font=("Segoe UI", 9, "bold"),
                  padx=12, cursor="hand2").pack(side=tk.LEFT)

        r2 = tk.Frame(f2, bg=BG2); r2.pack(fill=tk.X, padx=8, pady=4)
        tk.Label(r2, text="📍 Fallback:", bg=BG2, fg=FG,
                 font=("Segoe UI", 10)).pack(side=tk.LEFT)
        self.vf = tk.StringVar(value=FALLBACK)
        tk.Entry(r2, textvariable=self.vf, width=12, bg=BGI, fg=FG,
                 insertbackground=FG, relief="flat", font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=8)
        tk.Label(r2, text="⏱ Интервал:", bg=BG2, fg=FG,
                 font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=(20,0))
        self.vp = tk.StringVar(value="0.4")
        tk.Entry(r2, textvariable=self.vp, width=6, bg=BGI, fg=FG,
                 insertbackground=FG, relief="flat", font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=8)
        self.lz = tk.Label(r2, text="Зон: 0", bg=BG2, fg=FG, font=("Segoe UI", 10))
        self.lz.pack(side=tk.RIGHT, padx=10)

        # ── Зоны ──
        f3 = tk.LabelFrame(m, text="  📋  Зоны  ",
                            bg=BG2, fg=CYAN, font=("Segoe UI", 10, "bold"))
        f3.pack(fill=tk.X, pady=(0,5))
        fi = tk.Frame(f3, bg=BG2); fi.pack(fill=tk.X, padx=5, pady=5)
        self.tz = tk.Text(fi, height=3, bg=BGI, fg=FG, font=("Consolas", 9),
                           relief="flat", state=tk.DISABLED)
        sc = ttk.Scrollbar(fi, command=self.tz.yview)
        self.tz.configure(yscrollcommand=sc.set)
        sc.pack(side=tk.RIGHT, fill=tk.Y)
        self.tz.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # ── Лог ──
        f4 = tk.LabelFrame(m, text="  📝  Лог  ",
                            bg=BG2, fg=CYAN, font=("Segoe UI", 10, "bold"))
        f4.pack(fill=tk.BOTH, expand=True, pady=(0,2))
        self.tl = scrolledtext.ScrolledText(
            f4, height=10, state=tk.DISABLED, font=("Consolas", 9),
            bg="#0d1117", fg="#c9d1d9", relief="flat", selectbackground=BLUE)
        self.tl.pack(fill=tk.BOTH, expand=True, padx=5, pady=(5,2))
        tk.Button(f4, text="🗑 Очистить", command=self._clr,
                  bg="#374151", fg=FG, relief="flat", font=("Segoe UI", 8),
                  cursor="hand2").pack(anchor=tk.W, padx=5, pady=(0,5))

        # ── Кнопки ──
        bot = tk.Frame(self.root, bg=BAR); bot.pack(fill=tk.X, side=tk.BOTTOM)
        tk.Label(bot, text="💡 Любая клавиша = пауза   |   PostMessage-нажатия не вызывают паузу",
                 bg=BAR, fg=FGD, font=("Segoe UI", 9)).pack(pady=(8,3))
        bf = tk.Frame(bot, bg=BAR); bf.pack(pady=(0,10))

        self.bs = tk.Button(bf, text="▶  Старт", command=self._start,
                             bg=GREEN, fg="white", relief="flat",
                             font=("Segoe UI", 12, "bold"), padx=24, pady=8, cursor="hand2")
        self.bs.pack(side=tk.LEFT, padx=10)
        self.bp = tk.Button(bf, text="⏸  Пауза", command=self._pause,
                             bg=YELLOW, fg="#1a1a2e", relief="flat",
                             font=("Segoe UI", 12, "bold"), padx=24, pady=8,
                             cursor="hand2", state=tk.DISABLED)
        self.bp.pack(side=tk.LEFT, padx=10)
        self.bt = tk.Button(bf, text="⏹  Стоп", command=self._stop,
                             bg=ORANGE, fg="white", relief="flat",
                             font=("Segoe UI", 12, "bold"), padx=24, pady=8,
                             cursor="hand2", state=tk.DISABLED)
        self.bt.pack(side=tk.LEFT, padx=10)
        tk.Button(bf, text="✕  Выход", command=self._close,
                  bg=RED, fg="white", relief="flat",
                  font=("Segoe UI", 12, "bold"), padx=24, pady=8,
                  cursor="hand2").pack(side=tk.LEFT, padx=10)

    # ── Status bar ──
    def _bar(self):
        c = self.cv; c.delete("all")
        w, h = c.winfo_width(), c.winfo_height()
        if self._mode == "running":
            c.configure(bg=GREEN)
            for i in range(10):
                ax = self._ax + i*50
                if -30 < ax < w+30:
                    c.create_text(ax, h//2, text="›", fill=GREENL, font=("Segoe UI", 28, "bold"))
            c.create_text(w//2, h//2, text="▶  РАБОТАЕТ", fill="white", font=("Segoe UI", 16, "bold"))
        elif self._mode == "paused":
            c.configure(bg=BLUE)
            c.create_text(w//2, h//2, text="⏸  ПАУЗА", fill="white", font=("Segoe UI", 16, "bold"))
        else:
            c.configure(bg=BLUE)
            c.create_text(w//2, h//2, text="⏹  СТОП", fill="white", font=("Segoe UI", 16, "bold"))

    def _tick(self):
        if not self._anim: return
        if self._mode == "running":
            w = self.cv.winfo_width() or 960
            self._ax += 4
            if self._ax > w+50: self._ax = -500
        self._bar()
        self.root.after(50, self._tick)

    def _setm(self, m):
        self._mode = m
        if m == "running" and not self._anim:
            self._anim = True; self._tick()
        else:
            self._bar()

    # ── Scan ──
    def _auto_scan(self):
        self._log("🔍 Поиск WMS...")
        threading.Thread(target=lambda: self._apply(scan_wms()), daemon=True).start()

    def _sc_wms(self):
        self._log("🔍 Поиск WMS...")
        threading.Thread(target=lambda: self.root.after(0, lambda: self._apply(scan_wms())),
                         daemon=True).start()

    def _sc_all(self):
        self._log("📋 Все окна...")
        threading.Thread(target=lambda: self.root.after(0, lambda: self._apply(scan_all())),
                         daemon=True).start()

    def _apply(self, procs):
        self._procs = procs
        vs = ["(Все — без фильтра)"] + [f"[PID:{p}] {n} — «{t}»" for p,n,t in procs]
        self.cb["values"] = vs
        if not procs:
            self.va.set(vs[0]); self._pid = None
            self.la.config(text="⚠ WMS не найден. Откройте WMS → 🔍", fg=YELLOW)
            self._log("⚠ WMS не найден")
        elif len(procs) == 1:
            self._pid = procs[0][0]; self.va.set(vs[1])
            self.la.config(text=f"✅ {procs[0][1]} PID:{procs[0][0]}", fg=GREENL)
            self._log(f"✅ WMS: {procs[0][1]} PID:{procs[0][0]}")
        else:
            self._pid = procs[0][0]; self.va.set(vs[1])
            self.la.config(text=f"✅ Найдено {len(procs)}. Выбрано: {procs[0][1]}", fg=GREENL)

    def _on_cb(self, e=None):
        i = self.cb.current()
        if i <= 0:
            self._pid = None
            self.la.config(text="ℹ Без фильтра", fg=CYAN)
        else:
            self._pid = self._procs[i-1][0]
            self.la.config(text=f"✅ PID:{self._pid}", fg=GREENL)

    # ── Zones ──
    def _auto_zones(self):
        p = self.zones.auto()
        if p:
            self.vx.set(p)
            self._log(f"✅ Зоны: {os.path.basename(p)}")
            self._show_z()
        else:
            self._log("ℹ Файл зон не найден")

    def _browse(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if p:
            self.vx.set(p); self.zones.load(p)
            self._log(f"📁 {os.path.basename(p)}")
            self._show_z()

    def _show_z(self):
        self.lz.config(text=f"Зон: {len(self.zones.m)}")
        self.tz.config(state=tk.NORMAL); self.tz.delete("1.0", tk.END)
        for k, v in self.zones.d.items():
            self.tz.insert(tk.END, f"  {k}  →  {v}\n")
        if not self.zones.d:
            self.tz.insert(tk.END, "  (пусто)\n")
        self.tz.config(state=tk.DISABLED)

    # ── Log ──
    def _log(self, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        def i():
            self.tl.config(state=tk.NORMAL)
            self.tl.insert(tk.END, f"[{ts}] {msg}\n")
            self.tl.see(tk.END)
            self.tl.config(state=tk.DISABLED)
        self.root.after(0, i)

    def _clr(self):
        self.tl.config(state=tk.NORMAL)
        self.tl.delete("1.0", tk.END)
        self.tl.config(state=tk.DISABLED)

    # ── Keyboard ──
    def _on_key(self, ev):
        if self.bot and self.bot.running and not self.bot.paused:
            self.bot.paused = True
            self.bot.log(f"⏸ ПАУЗА (клавиша: {ev.name})")
            self.root.after(0, lambda: self._setm("paused"))
            self.root.after(0, lambda: self.bp.config(text="▶  Продолжить"))

    # ── Controls ──
    def _start(self):
        xls = self.vx.get().strip()
        if xls and os.path.exists(xls):
            self.zones.load(xls); self._show_z()

        self.bot = WMSBot(
            self.zones, self.vf.get().strip() or FALLBACK, self._pid,
            log_cb=self._log,
            status_cb=lambda m: self.root.after(0, lambda: self._setm(m)))
        try: self.bot.poll = float(self.vp.get())
        except: pass

        self.thread = threading.Thread(target=self.bot.run, daemon=True)
        self.thread.start()
        self.bs.config(state=tk.DISABLED, bg="#374151")
        self.bp.config(state=tk.NORMAL, text="⏸  Пауза")
        self.bt.config(state=tk.NORMAL)
        self._setm("running")

    def _pause(self):
        if self.bot:
            self.bot.toggle_pause()
            if self.bot.paused:
                self._setm("paused"); self.bp.config(text="▶  Продолжить")
            else:
                self._setm("running"); self.bp.config(text="⏸  Пауза")

    def _stop(self):
        if self.bot: self.bot.stop()
        self.bs.config(state=tk.NORMAL, bg=GREEN)
        self.bp.config(state=tk.DISABLED, text="⏸  Пауза")
        self.bt.config(state=tk.DISABLED)
        self._anim = False; self._setm("stopped")

    def _close(self):
        if self.bot: self.bot.stop()
        try: keyboard.unhook(self._kb)
        except: pass
        self._anim = False; self.root.destroy()

    def run(self):
        self.root.protocol("WM_DELETE_WINDOW", self._close)
        self._bar(); self.root.mainloop()


if __name__ == "__main__":
    App().run()
