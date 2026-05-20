"""
WMS Automation Tool v4.0
========================
Циклическая автоматизация WMS.
Может начать с ЛЮБОГО шага цикла.

Цикл (по скриншотам 1-12):
  1.  Главное меню         → жмём F2 (Запросить работу)
  2.  Взятие работы        → жмём F2 (Взять работу)
  3-4. Перемещение к источнику → копируем Место → Контроль → Ок
  5-6. Поиск паллеты       → копируем Паллета → Контроль → Ок
  7-8. Поиск коробки       → копируем Коробка → Контроль → Ок
  9-10. Поиск места-приёмника (коробка) → зона из XLS → Контроль → Ок
       (если ошибка → D-KM-1)
  11. Размещение в место   → жмём Ок
  12. Поиск места-приёмника (паллета) → ВСЕГДА D-KM-1 → Ок
  → цикл заново с шага 1

Установка:  pip install pywinauto openpyxl keyboard
Запуск:     python wms_auto.py  (от имени администратора)
"""

import time
import sys
import os
import re
import glob
import logging
import threading
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
from datetime import datetime

try:
    import pywinauto
    from pywinauto import Desktop
except ImportError:
    print("ОШИБКА: pip install pywinauto")
    sys.exit(1)

try:
    import openpyxl
except ImportError:
    print("ОШИБКА: pip install openpyxl")
    sys.exit(1)

try:
    import keyboard
except ImportError:
    print("ОШИБКА: pip install keyboard")
    sys.exit(1)

import subprocess

# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("wms_auto.log", encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger("WMS")

# ---------------------------------------------------------------------------
#  Константы
# ---------------------------------------------------------------------------
FALLBACK_LOCATION = "D-KM-1"

# UI
BG_DARK    = "#1a1a2e"
BG_CARD    = "#16213e"
BG_INPUT   = "#1c2a4a"
FG_TEXT    = "#e0e0e0"
FG_DIM     = "#7a7a9a"
C_BLUE     = "#2563eb"
C_BLUE_D   = "#1d4ed8"
C_GREEN    = "#16a34a"
C_GREEN_L  = "#22c55e"
C_RED      = "#dc2626"
C_RED_L    = "#ef4444"
C_YELLOW   = "#eab308"
C_YELLOW_L = "#facc15"
C_CYAN     = "#06b6d4"
C_ORANGE   = "#ea580c"
BAR_BG     = "#111827"


# ===================================================================
#  ZoneLookup — чтение XLS
# ===================================================================
class ZoneLookup:
    def __init__(self):
        self.mapping = {}
        self.display = {}
        self.path = None

    def load(self, path):
        if not path or not os.path.exists(path):
            return False
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            ws = wb.active
            self.mapping.clear()
            self.display.clear()
            for row in ws.iter_rows(min_row=1, values_only=True):
                if row and len(row) >= 2 and row[0] and row[1]:
                    k = str(row[0]).strip()
                    v = str(row[1]).strip()
                    self.mapping[k.lower()] = v
                    self.display[k] = v
            wb.close()
            self.path = path
            return True
        except Exception as e:
            logger.error("Ошибка загрузки XLS: %s", e)
            return False

    def lookup(self, zone_text):
        zt = zone_text.lower().strip()
        # точное
        if zt in self.mapping:
            return self.mapping[zt]
        # ключ содержится в тексте
        for k, v in self.mapping.items():
            if k in zt:
                return v
        # текст содержится в ключе
        for k, v in self.mapping.items():
            if zt in k:
                return v
        return None

    def auto_find(self):
        d = os.path.dirname(os.path.abspath(__file__))
        for name in ["zones.xlsx", "зоны.xlsx", "зона.xlsx", "zone.xlsx"]:
            p = os.path.join(d, name)
            if os.path.exists(p) and self.load(p):
                return p
        for f in sorted(glob.glob(os.path.join(d, "*.xlsx"))):
            if self.load(f):
                return f
        return None


# ===================================================================
#  Поиск окон — ПРОСТОЙ И НАДЁЖНЫЙ
# ===================================================================

def _get_all_windows():
    """Получить все окна через оба бэкенда."""
    windows = []
    for backend in ("win32", "uia"):
        try:
            for w in Desktop(backend=backend).windows():
                try:
                    t = w.window_text()
                    if t:
                        windows.append((t, w, backend))
                except:
                    pass
        except:
            pass
    return windows


def find_window(title_part, pid=None):
    """Найти окно по части заголовка. Возвращает wrapper или None."""
    title_lower = title_part.lower()
    for t, w, be in _get_all_windows():
        if title_lower in t.lower():
            if pid:
                try:
                    if w.process_id() != pid:
                        continue
                except:
                    continue
            return w
    return None


def find_wms_processes():
    """Найти процессы с WMS-окнами."""
    known = ["главное меню", "взятие работы", "перемещение к источнику",
             "поиск паллеты", "поиск палеты", "поиск коробки",
             "поиск места-приёмника", "поиск места-приемника",
             "размещение в место"]
    found = {}
    for t, w, be in _get_all_windows():
        tl = t.lower()
        for k in known:
            if k in tl:
                try:
                    p = w.process_id()
                    if p and p not in found:
                        name = _proc_name(p)
                        found[p] = (p, name, t)
                except:
                    pass
                break
    return list(found.values())


def find_all_visible():
    """Все видимые окна для ручного выбора."""
    found = {}
    for t, w, be in _get_all_windows():
        try:
            if not w.is_visible():
                continue
            p = w.process_id()
            if p and p > 0 and p not in found:
                found[p] = (p, _proc_name(p), t)
        except:
            pass
    return sorted(found.values(), key=lambda x: x[1].lower())


def _proc_name(pid):
    try:
        r = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV", "/NH"],
            capture_output=True, text=True, timeout=3)
        for line in r.stdout.strip().split("\n"):
            parts = line.replace('"', '').split(",")
            if parts:
                return parts[0].strip()
    except:
        pass
    return "?"


# ===================================================================
#  Работа с контролами окна
# ===================================================================

def get_edits(win):
    """Все Edit-контролы окна."""
    result = []
    try:
        for c in win.children():
            try:
                if "edit" in c.class_name().lower():
                    result.append(c)
            except:
                pass
    except:
        pass
    return result


def get_statics(win):
    """Все тексты Static-контролов."""
    texts = []
    try:
        for c in win.children():
            try:
                if "static" in c.class_name().lower():
                    t = c.window_text()
                    if t:
                        texts.append(t)
            except:
                pass
    except:
        pass
    return texts


def get_buttons(win):
    """Все кнопки."""
    result = []
    try:
        for c in win.children():
            try:
                if "button" in c.class_name().lower():
                    result.append(c)
            except:
                pass
    except:
        pass
    return result


def read_edit_value(edit):
    """Прочитать значение из Edit (точное чтение через Win32 API)."""
    try:
        return edit.window_text().strip()
    except:
        return ""


def write_edit_value(edit, value):
    """Записать значение в Edit."""
    try:
        edit.set_focus()
        time.sleep(0.05)
        edit.set_edit_text("")
        time.sleep(0.05)
        edit.set_edit_text(value)
        time.sleep(0.05)
    except Exception:
        try:
            edit.set_focus()
            time.sleep(0.05)
            keyboard.send("ctrl+a")
            time.sleep(0.05)
            keyboard.write(value, delay=0.02)
        except:
            pass


def click_button_text(win, *texts):
    """Нажать кнопку, текст которой содержит одну из строк."""
    for btn in get_buttons(win):
        try:
            bt = btn.window_text().strip().lower()
            for t in texts:
                if t.lower() in bt:
                    btn.click()
                    return True
        except:
            pass
    return False


def click_ok(win):
    """Нажать Ок."""
    return click_button_text(win, "ок", "ok", "oк", "оk")


def click_f2_button(win):
    """Нажать кнопку F2 / Запросить / Взять."""
    return click_button_text(win, "f2", "запросить", "взять")


def find_edit_near_label(win, label_part):
    """
    Найти Edit рядом с меткой (по вертикали ±30px).
    Это КЛЮЧЕВОЙ метод — читает данные из конкретного поля.
    """
    statics = []
    edits = []
    try:
        for c in win.children():
            try:
                cn = c.class_name().lower()
                if "static" in cn:
                    statics.append(c)
                elif "edit" in cn:
                    edits.append(c)
            except:
                pass
    except:
        pass

    # Ищем нужный Static
    target = None
    for s in statics:
        try:
            if label_part.lower() in s.window_text().lower():
                target = s
                break
        except:
            pass

    if not target:
        return None

    # Ищем ближайший Edit по вертикали
    try:
        sr = target.rectangle()
        sy = (sr.top + sr.bottom) // 2
    except:
        return None

    best = None
    best_d = 9999
    for e in edits:
        try:
            er = e.rectangle()
            ey = (er.top + er.bottom) // 2
            d = abs(ey - sy)
            if d < 30 and d < best_d:
                best_d = d
                best = e
        except:
            pass
    return best


# ===================================================================
#  ГЛАВНАЯ ЛОГИКА — Автоматизация
# ===================================================================
class WMSBot:

    def __init__(self, zones, fallback, pid=None, log_cb=None, status_cb=None):
        self.zones = zones
        self.fallback = fallback or FALLBACK_LOCATION
        self.pid = pid
        self.log_cb = log_cb
        self.status_cb = status_cb

        self.running = False
        self.paused = False
        self.poll = 0.5       # интервал опроса
        self.delay = 0.3      # задержка после действия
        self._suppress = False  # подавление паузы при программных нажатиях

    def log(self, msg, lvl="INFO"):
        logger.log(getattr(logging, lvl, 20), msg)
        if self.log_cb:
            self.log_cb(f"[{lvl}] {msg}")

    def _send_key(self, key):
        """Отправить клавишу БЕЗ срабатывания паузы."""
        self._suppress = True
        try:
            keyboard.send(key)
        finally:
            time.sleep(0.05)
            self._suppress = False

    # ------------------------------------------------------------------
    #  Определение текущего окна — ЧТО СЕЙЧАС НА ЭКРАНЕ?
    # ------------------------------------------------------------------
    def detect(self):
        """
        Возвращает (тип_окна, окно) или ("unknown", None).
        Порядок проверки важен: сначала модальные, потом основные.
        """
        pid = self.pid

        # Ошибка / предупреждение
        for t in ("Ошибка", "Error", "Предупреждение", "Warning"):
            w = find_window(t, pid)
            if w:
                return "error", w

        # Подтверждение / информация / сообщение
        for t in ("Подтверждение", "Подтвердить", "Confirm",
                   "Информация", "Information", "Сообщение", "Message"):
            w = find_window(t, pid)
            if w:
                return "confirm", w

        # Размещение в место (скрин 11)
        w = find_window("Размещение в место", pid)
        if w:
            return "place", w

        # Поиск места-приёмника (скрин 9/10/12)
        w = find_window("Поиск места-при", pid)
        if not w:
            w = find_window("Поиск места-приёмника", pid)
        if not w:
            w = find_window("Поиск места-приемника", pid)
        if w:
            # Определяем: коробка (скрин 9) или паллета (скрин 12)?
            all_text = " ".join(get_statics(w)).lower()
            if "паллет" in all_text or "содержимого" in all_text:
                return "dest_pallet", w   # скрин 12
            else:
                return "dest_box", w       # скрин 9

        # Поиск коробки (скрин 7)
        w = find_window("Поиск коробки", pid)
        if w:
            return "box", w

        # Поиск паллеты (скрин 5)
        w = find_window("Поиск паллеты", pid)
        if not w:
            w = find_window("Поиск палеты", pid)
        if w:
            return "pallet", w

        # Перемещение к источнику (скрин 3)
        w = find_window("Перемещение к источнику", pid)
        if w:
            return "move_source", w

        # Взятие работы (скрин 2)
        w = find_window("Взятие работы", pid)
        if w:
            return "take_work", w

        # Главное меню (скрин 1)
        w = find_window("Главное меню", pid)
        if w:
            return "main_menu", w

        return "unknown", None

    # ------------------------------------------------------------------
    #  Обработчики каждого шага
    # ------------------------------------------------------------------

    def do_main_menu(self, w):
        """Скрин 1: Главное меню → жмём F2 (Запросить работу)"""
        self.log("→ [1] Главное меню: жму F2 (Запросить работу)")
        if not click_f2_button(w):
            try:
                w.set_focus()
                time.sleep(0.1)
                self._send_key("F2")
            except:
                pass

    def do_take_work(self, w):
        """Скрин 2: Взятие работы → жмём F2 (Взять работу)"""
        self.log("→ [2] Взятие работы: жму F2 (Взять работу)")
        if not click_f2_button(w):
            try:
                w.set_focus()
                time.sleep(0.1)
                self._send_key("F2")
            except:
                pass

    def do_move_source(self, w):
        """Скрин 3-4: Перемещение к источнику → Место → Контроль → Ок"""
        self.log("→ [3] Перемещение к источнику")

        # Читаем значение из поля «Место»
        edit_mesto = find_edit_near_label(w, "Место")
        edit_ctrl  = find_edit_near_label(w, "Контроль")

        if edit_mesto and edit_ctrl:
            val = read_edit_value(edit_mesto)
            ctrl_val = read_edit_value(edit_ctrl)

            if val and not ctrl_val:
                self.log(f"   Место='{val}' → Контроль")
                write_edit_value(edit_ctrl, val)
                time.sleep(0.15)
                click_ok(w)
            elif val and ctrl_val:
                # Уже заполнено (скрин 4) — просто жмём Ок
                self.log(f"   Уже заполнено: Контроль='{ctrl_val}' → Ок")
                click_ok(w)
            else:
                self.log("   Место пустое!", "WARNING")
        else:
            # Фоллбэк: по индексу
            edits = get_edits(w)
            self.log(f"   Edits: {[read_edit_value(e) for e in edits]}")
            if len(edits) >= 2:
                val = read_edit_value(edits[0])
                if val:
                    write_edit_value(edits[1], val)
                    time.sleep(0.15)
                    click_ok(w)

    def do_pallet(self, w):
        """Скрин 5-6: Поиск паллеты → Паллета → Контроль → Ок"""
        self.log("→ [5] Поиск паллеты")

        edit_pallet = find_edit_near_label(w, "Паллет")
        edit_ctrl   = find_edit_near_label(w, "Контроль")

        if edit_pallet and edit_ctrl:
            val = read_edit_value(edit_pallet)
            ctrl_val = read_edit_value(edit_ctrl)

            if val and not ctrl_val:
                self.log(f"   Паллета='{val}' → Контроль")
                write_edit_value(edit_ctrl, val)
                time.sleep(0.15)
                click_ok(w)
            elif val and ctrl_val:
                self.log(f"   Уже заполнено → Ок")
                click_ok(w)
            else:
                self.log("   Паллета пустая!", "WARNING")
        else:
            # Фоллбэк по всем Edit
            edits = get_edits(w)
            vals = [read_edit_value(e) for e in edits]
            self.log(f"   Fallback edits: {vals}")
            # Ищем первое заполненное поле и первое пустое после него
            src_val = None
            target = None
            for i, e in enumerate(edits):
                v = read_edit_value(e)
                if v and not src_val:
                    # Это может быть "Из места" - пропускаем, берём следующее
                    continue
                if v and src_val is None:
                    src_val = v
                if not v and target is None:
                    target = e
            # Проще: берём предпоследнее заполненное и последнее пустое
            for i in range(len(edits)-1, -1, -1):
                if not read_edit_value(edits[i]):
                    target = edits[i]
                    break
            for i in range(len(edits)-1, -1, -1):
                v = read_edit_value(edits[i])
                if v:
                    src_val = v
                    break
            if src_val and target:
                self.log(f"   Fallback: '{src_val}' → Контроль")
                write_edit_value(target, src_val)
                time.sleep(0.15)
                click_ok(w)

    def do_box(self, w):
        """Скрин 7-8: Поиск коробки → Коробка → Контроль → Ок"""
        self.log("→ [7] Поиск коробки")

        edit_box  = find_edit_near_label(w, "Коробк")
        edit_ctrl = find_edit_near_label(w, "Контроль")

        if edit_box and edit_ctrl:
            val = read_edit_value(edit_box)
            ctrl_val = read_edit_value(edit_ctrl)

            if val and not ctrl_val:
                self.log(f"   Коробка='{val}' → Контроль")
                write_edit_value(edit_ctrl, val)
                time.sleep(0.15)
                click_ok(w)
            elif val and ctrl_val:
                self.log(f"   Уже заполнено → Ок")
                click_ok(w)
            else:
                self.log("   Коробка пустая!", "WARNING")
        else:
            edits = get_edits(w)
            vals = [read_edit_value(e) for e in edits]
            self.log(f"   Fallback edits: {vals}")
            # Ищем последнее пустое
            src_val = None
            target = None
            for i in range(len(edits)-1, -1, -1):
                if not read_edit_value(edits[i]):
                    target = edits[i]
                    break
            for i in range(len(edits)-1, -1, -1):
                v = read_edit_value(edits[i])
                if v:
                    src_val = v
                    break
            if src_val and target:
                write_edit_value(target, src_val)
                time.sleep(0.15)
                click_ok(w)

    def do_dest_box(self, w):
        """
        Скрин 9-10: Поиск места-приёмника (КОРОБКА)
        Читаем зону → ищем в XLS → вставляем код → Ок
        Если ошибка → D-KM-1
        """
        self.log("→ [9] Поиск места-приёмника (коробка)")

        all_text = " ".join(get_statics(w))
        self.log(f"   Текст: {all_text}")

        # Извлекаем зону: ищем текст после "Зона:" или "Док"
        zone = ""
        m = re.search(r"[Зз]она[:\s]+(.+)", all_text)
        if m:
            zone = m.group(1).strip()
            # Убираем лишнее (кнопки и т.д.)
            zone = zone.split("\n")[0].strip()

        self.log(f"   Зона: '{zone}'")

        # Ищем код в XLS
        location = None
        if zone:
            location = self.zones.lookup(zone)
        if location:
            self.log(f"   XLS → {location}")
        else:
            location = self.fallback
            self.log(f"   Не найдено → fallback {location}")

        # Вставляем в Контроль
        edit_ctrl = find_edit_near_label(w, "Контроль")
        if not edit_ctrl:
            edits = get_edits(w)
            for e in edits:
                if not read_edit_value(e):
                    edit_ctrl = e
                    break
            if not edit_ctrl and edits:
                edit_ctrl = edits[-1]

        if edit_ctrl:
            write_edit_value(edit_ctrl, location)
            time.sleep(0.15)
            click_ok(w)

            # Ждём — может появиться ошибка
            time.sleep(0.8)
            typ, ew = self.detect()
            if typ == "error":
                self.log(f"   ❌ Ошибка! Закрываю → вставляю {self.fallback}")
                click_ok(ew)
                time.sleep(0.5)
                # Снова появится окно места-приёмника
                typ2, w2 = self.detect()
                if typ2 in ("dest_box", "dest_pallet"):
                    ec2 = find_edit_near_label(w2, "Контроль")
                    if not ec2:
                        eds = get_edits(w2)
                        for e in eds:
                            if not read_edit_value(e):
                                ec2 = e
                                break
                    if ec2:
                        write_edit_value(ec2, self.fallback)
                        time.sleep(0.15)
                        click_ok(w2)
            elif typ == "confirm":
                click_ok(ew)
        else:
            self.log("   Не найден Контроль!", "WARNING")

    def do_dest_pallet(self, w):
        """
        Скрин 12: Поиск места-приёмника (ПАЛЛЕТА)
        ВСЕГДА вставляем D-KM-1 (fallback) → Ок
        """
        self.log(f"→ [12] Поиск места-приёмника (паллета): ВСЕГДА {self.fallback}")

        edit_ctrl = find_edit_near_label(w, "Контроль")
        if not edit_ctrl:
            edits = get_edits(w)
            for e in edits:
                if not read_edit_value(e):
                    edit_ctrl = e
                    break
            if not edit_ctrl and edits:
                edit_ctrl = edits[-1]

        if edit_ctrl:
            write_edit_value(edit_ctrl, self.fallback)
            time.sleep(0.15)
            click_ok(w)
        else:
            self.log("   Не найден Контроль!", "WARNING")

    def do_place(self, w):
        """Скрин 11: Размещение в место → просто Ок"""
        self.log("→ [11] Размещение в место: жму Ок")
        click_ok(w)

    def do_error(self, w):
        """Окно ошибки → закрываем (Ок)"""
        txt = " ".join(get_statics(w))
        self.log(f"→ Ошибка: {txt}")
        click_ok(w)

    def do_confirm(self, w):
        """Подтверждение → Ок"""
        self.log("→ Подтверждение: жму Ок")
        click_ok(w)
        # Может вылезти ещё одно
        time.sleep(0.3)
        t2, w2 = self.detect()
        if t2 == "confirm":
            click_ok(w2)

    # ------------------------------------------------------------------
    #  Главный цикл
    # ------------------------------------------------------------------
    def run(self):
        self.running = True
        self.log("=" * 50)
        self.log("🚀 Автоматизация ЗАПУЩЕНА")
        self.log(f"   Fallback: {self.fallback}")
        self.log(f"   Зон: {len(self.zones.mapping)}")
        self.log(f"   PID: {self.pid or 'все'}")
        self.log("=" * 50)

        HANDLERS = {
            "main_menu":   self.do_main_menu,
            "take_work":   self.do_take_work,
            "move_source": self.do_move_source,
            "pallet":      self.do_pallet,
            "box":         self.do_box,
            "dest_box":    self.do_dest_box,
            "dest_pallet": self.do_dest_pallet,
            "place":       self.do_place,
            "error":       self.do_error,
            "confirm":     self.do_confirm,
        }

        while self.running:
            if self.paused:
                if self.status_cb:
                    self.status_cb("paused")
                time.sleep(0.3)
                continue

            try:
                if self.status_cb:
                    self.status_cb("running")

                typ, win = self.detect()

                if typ == "unknown":
                    time.sleep(self.poll)
                    continue

                handler = HANDLERS.get(typ)
                if handler:
                    handler(win)

                time.sleep(self.delay)

            except Exception as e:
                self.log(f"Ошибка: {e}", "ERROR")
                time.sleep(1)

        if self.status_cb:
            self.status_cb("stopped")
        self.log("=" * 50)
        self.log("⏹ Автоматизация ОСТАНОВЛЕНА")
        self.log("=" * 50)

    def stop(self):
        self.running = False

    def toggle_pause(self):
        self.paused = not self.paused
        self.log("⏸ ПАУЗА" if self.paused else "▶ ПРОДОЛЖЕНИЕ")


# ===================================================================
#  GUI
# ===================================================================
class App:

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("WMS Автоматизация v4.0")
        self.root.geometry("960x750")
        self.root.configure(bg=BG_DARK)
        self.root.resizable(True, True)

        self.bot = None
        self.thread = None
        self.zones = ZoneLookup()

        self._sel_pid = None
        self._procs = []

        self._arrow_x = -100
        self._anim = False
        self._mode = "stopped"

        self._build()
        self.root.after(200, self._auto_zones)
        self.root.after(500, self._auto_scan)

        self._kb_hook = keyboard.on_press(self._on_key)

    # ---------- UI ----------
    def _build(self):
        # Status bar
        self.canvas = tk.Canvas(self.root, height=54, bg=C_BLUE, highlightthickness=0)
        self.canvas.pack(fill=tk.X)
        self.canvas.bind("<Configure>", lambda e: self._draw_bar())

        main = tk.Frame(self.root, bg=BG_DARK)
        main.pack(fill=tk.BOTH, expand=True, padx=12, pady=6)

        # --- Приложение ---
        f_app = tk.LabelFrame(main, text="  🖥  WMS-приложение  ",
                               bg=BG_CARD, fg=C_CYAN, font=("Segoe UI", 10, "bold"))
        f_app.pack(fill=tk.X, pady=(0,6))

        r = tk.Frame(f_app, bg=BG_CARD)
        r.pack(fill=tk.X, padx=8, pady=6)
        tk.Label(r, text="Приложение:", bg=BG_CARD, fg=FG_TEXT,
                 font=("Segoe UI", 10)).pack(side=tk.LEFT)
        self.var_app = tk.StringVar(value="(сканирование...)")
        self.cmb = ttk.Combobox(r, textvariable=self.var_app, width=55,
                                 state="readonly", font=("Segoe UI", 9))
        self.cmb.pack(side=tk.LEFT, padx=8)
        self.cmb.bind("<<ComboboxSelected>>", self._on_cmb)

        tk.Button(r, text="🔍 WMS", command=self._scan_wms,
                  bg=C_BLUE, fg="white", relief="flat",
                  font=("Segoe UI", 9, "bold"), padx=10, cursor="hand2"
                  ).pack(side=tk.LEFT, padx=3)
        tk.Button(r, text="📋 Все", command=self._scan_all,
                  bg="#374151", fg=FG_TEXT, relief="flat",
                  font=("Segoe UI", 9), padx=10, cursor="hand2"
                  ).pack(side=tk.LEFT, padx=3)

        self.lbl_app = tk.Label(f_app, text="", bg=BG_CARD, fg=FG_DIM,
                                 font=("Segoe UI", 9))
        self.lbl_app.pack(anchor=tk.W, padx=12, pady=(0,4))

        # --- Настройки ---
        f_set = tk.LabelFrame(main, text="  ⚙  Настройки  ",
                               bg=BG_CARD, fg=C_CYAN, font=("Segoe UI", 10, "bold"))
        f_set.pack(fill=tk.X, pady=(0,6))

        r1 = tk.Frame(f_set, bg=BG_CARD)
        r1.pack(fill=tk.X, padx=8, pady=4)
        tk.Label(r1, text="📁 Файл зон:", bg=BG_CARD, fg=FG_TEXT,
                 font=("Segoe UI", 10)).pack(side=tk.LEFT)
        self.var_xls = tk.StringVar()
        tk.Entry(r1, textvariable=self.var_xls, width=48,
                 bg=BG_INPUT, fg=FG_TEXT, insertbackground=FG_TEXT,
                 relief="flat", font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=8)
        tk.Button(r1, text="Обзор…", command=self._browse,
                  bg=C_BLUE, fg="white", relief="flat",
                  font=("Segoe UI", 9, "bold"), padx=12, cursor="hand2"
                  ).pack(side=tk.LEFT)

        r2 = tk.Frame(f_set, bg=BG_CARD)
        r2.pack(fill=tk.X, padx=8, pady=4)
        tk.Label(r2, text="📍 Fallback:", bg=BG_CARD, fg=FG_TEXT,
                 font=("Segoe UI", 10)).pack(side=tk.LEFT)
        self.var_fb = tk.StringVar(value=FALLBACK_LOCATION)
        tk.Entry(r2, textvariable=self.var_fb, width=12,
                 bg=BG_INPUT, fg=FG_TEXT, insertbackground=FG_TEXT,
                 relief="flat", font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=8)
        tk.Label(r2, text="⏱ Интервал:", bg=BG_CARD, fg=FG_TEXT,
                 font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=(20,0))
        self.var_poll = tk.StringVar(value="0.5")
        tk.Entry(r2, textvariable=self.var_poll, width=6,
                 bg=BG_INPUT, fg=FG_TEXT, insertbackground=FG_TEXT,
                 relief="flat", font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=8)
        self.lbl_zn = tk.Label(r2, text="Зон: 0", bg=BG_CARD, fg=FG_TEXT,
                                font=("Segoe UI", 10))
        self.lbl_zn.pack(side=tk.RIGHT, padx=10)

        # --- Зоны ---
        f_z = tk.LabelFrame(main, text="  📋  Зоны  ",
                             bg=BG_CARD, fg=C_CYAN, font=("Segoe UI", 10, "bold"))
        f_z.pack(fill=tk.X, pady=(0,6))
        fi = tk.Frame(f_z, bg=BG_CARD)
        fi.pack(fill=tk.X, padx=5, pady=5)
        self.txt_z = tk.Text(fi, height=3, bg=BG_INPUT, fg=FG_TEXT,
                              font=("Consolas", 9), relief="flat",
                              state=tk.DISABLED)
        sc = ttk.Scrollbar(fi, command=self.txt_z.yview)
        self.txt_z.configure(yscrollcommand=sc.set)
        sc.pack(side=tk.RIGHT, fill=tk.Y)
        self.txt_z.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # --- Лог ---
        f_l = tk.LabelFrame(main, text="  📝  Лог  ",
                             bg=BG_CARD, fg=C_CYAN, font=("Segoe UI", 10, "bold"))
        f_l.pack(fill=tk.BOTH, expand=True, pady=(0,2))
        self.txt_log = scrolledtext.ScrolledText(
            f_l, height=10, state=tk.DISABLED,
            font=("Consolas", 9), bg="#0d1117", fg="#c9d1d9",
            relief="flat", selectbackground=C_BLUE)
        self.txt_log.pack(fill=tk.BOTH, expand=True, padx=5, pady=(5,2))
        tk.Button(f_l, text="🗑 Очистить", command=self._clear_log,
                  bg="#374151", fg=FG_TEXT, relief="flat",
                  font=("Segoe UI", 8), cursor="hand2"
                  ).pack(anchor=tk.W, padx=5, pady=(0,5))

        # --- Кнопки ---
        bot = tk.Frame(self.root, bg=BAR_BG)
        bot.pack(fill=tk.X, side=tk.BOTTOM)
        tk.Label(bot, text="💡 Любая клавиша = пауза",
                 bg=BAR_BG, fg=FG_DIM, font=("Segoe UI", 9)).pack(pady=(8,3))

        bf = tk.Frame(bot, bg=BAR_BG)
        bf.pack(pady=(0,10))

        self.btn_start = tk.Button(
            bf, text="▶  Старт", command=self._start,
            bg=C_GREEN, fg="white", relief="flat",
            font=("Segoe UI", 12, "bold"), padx=24, pady=8, cursor="hand2")
        self.btn_start.pack(side=tk.LEFT, padx=10)

        self.btn_pause = tk.Button(
            bf, text="⏸  Пауза", command=self._pause,
            bg=C_YELLOW, fg="#1a1a2e", relief="flat",
            font=("Segoe UI", 12, "bold"), padx=24, pady=8,
            cursor="hand2", state=tk.DISABLED)
        self.btn_pause.pack(side=tk.LEFT, padx=10)

        self.btn_stop = tk.Button(
            bf, text="⏹  Стоп", command=self._stop,
            bg=C_ORANGE, fg="white", relief="flat",
            font=("Segoe UI", 12, "bold"), padx=24, pady=8,
            cursor="hand2", state=tk.DISABLED)
        self.btn_stop.pack(side=tk.LEFT, padx=10)

        tk.Button(
            bf, text="✕  Выход", command=self._close,
            bg=C_RED, fg="white", relief="flat",
            font=("Segoe UI", 12, "bold"), padx=24, pady=8,
            cursor="hand2").pack(side=tk.LEFT, padx=10)

    # ---------- Status bar ----------
    def _draw_bar(self):
        c = self.canvas
        c.delete("all")
        w = c.winfo_width()
        h = c.winfo_height()
        if self._mode == "running":
            c.configure(bg=C_GREEN)
            for i in range(10):
                ax = self._arrow_x + i * 50
                if -30 < ax < w+30:
                    c.create_text(ax, h//2, text="›", fill=C_GREEN_L,
                                  font=("Segoe UI", 30, "bold"))
            c.create_text(w//2, h//2, text="▶  РАБОТАЕТ",
                          fill="white", font=("Segoe UI", 16, "bold"))
        elif self._mode == "paused":
            c.configure(bg=C_BLUE)
            c.create_text(w//2, h//2, text="⏸  ПАУЗА",
                          fill="white", font=("Segoe UI", 16, "bold"))
        else:
            c.configure(bg=C_BLUE)
            c.create_text(w//2, h//2, text="⏹  ОСТАНОВЛЕНО",
                          fill="white", font=("Segoe UI", 16, "bold"))

    def _anim_tick(self):
        if not self._anim:
            return
        if self._mode == "running":
            w = self.canvas.winfo_width() or 960
            self._arrow_x += 4
            if self._arrow_x > w+50:
                self._arrow_x = -500
        self._draw_bar()
        self.root.after(50, self._anim_tick)

    def _set_mode(self, m):
        self._mode = m
        if m == "running" and not self._anim:
            self._anim = True
            self._anim_tick()
        else:
            self._draw_bar()

    # ---------- Scanning ----------
    def _auto_scan(self):
        self._log("🔍 Автопоиск WMS...")
        def do():
            ps = find_wms_processes()
            self.root.after(0, lambda: self._apply_scan(ps))
        threading.Thread(target=do, daemon=True).start()

    def _scan_wms(self):
        self._log("🔍 Поиск WMS...")
        def do():
            ps = find_wms_processes()
            self.root.after(0, lambda: self._apply_scan(ps))
        threading.Thread(target=do, daemon=True).start()

    def _scan_all(self):
        self._log("📋 Все окна...")
        def do():
            ps = find_all_visible()
            self.root.after(0, lambda: self._apply_scan(ps))
        threading.Thread(target=do, daemon=True).start()

    def _apply_scan(self, procs):
        self._procs = procs
        vals = ["(Все процессы — без фильтра)"]
        for p, n, t in procs:
            vals.append(f"[PID:{p}] {n} — «{t}»")
        self.cmb["values"] = vals

        if len(procs) == 0:
            self.var_app.set("(Все процессы — без фильтра)")
            self._sel_pid = None
            self.lbl_app.config(text="⚠ WMS не найден. Откройте WMS и нажмите 🔍",
                                 fg=C_YELLOW)
            self._log("⚠ WMS-окна не найдены")
        elif len(procs) == 1:
            self._sel_pid = procs[0][0]
            self.var_app.set(vals[1])
            self.lbl_app.config(text=f"✅ Найден: {procs[0][1]} (PID:{procs[0][0]})",
                                 fg=C_GREEN_L)
            self._log(f"✅ WMS: {procs[0][1]} PID:{procs[0][0]}")
        else:
            self._sel_pid = procs[0][0]
            self.var_app.set(vals[1])
            self.lbl_app.config(
                text=f"✅ Найдено {len(procs)} прил. Выбрано: {procs[0][1]}",
                fg=C_GREEN_L)
            self._log(f"✅ Найдено {len(procs)} WMS-окон")

    def _on_cmb(self, e=None):
        idx = self.cmb.current()
        if idx <= 0:
            self._sel_pid = None
            self.lbl_app.config(text="ℹ Без фильтра", fg=C_CYAN)
        else:
            p = self._procs[idx-1]
            self._sel_pid = p[0]
            self.lbl_app.config(text=f"✅ {p[1]} PID:{p[0]}", fg=C_GREEN_L)

    # ---------- Zones ----------
    def _auto_zones(self):
        p = self.zones.auto_find()
        if p:
            self.var_xls.set(p)
            self._log(f"✅ Зоны: {os.path.basename(p)}")
            self._show_zones()
        else:
            self._log("ℹ Файл зон не найден. Загрузите через «Обзор»")

    def _browse(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if p:
            self.var_xls.set(p)
            self.zones.load(p)
            self._log(f"📁 Загружен: {os.path.basename(p)}")
            self._show_zones()

    def _show_zones(self):
        self.lbl_zn.config(text=f"Зон: {len(self.zones.mapping)}")
        self.txt_z.config(state=tk.NORMAL)
        self.txt_z.delete("1.0", tk.END)
        if self.zones.display:
            for k, v in self.zones.display.items():
                self.txt_z.insert(tk.END, f"  {k}  →  {v}\n")
        else:
            self.txt_z.insert(tk.END, "  (пусто)\n")
        self.txt_z.config(state=tk.DISABLED)

    # ---------- Log ----------
    def _log(self, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        def ins():
            self.txt_log.config(state=tk.NORMAL)
            self.txt_log.insert(tk.END, f"[{ts}] {msg}\n")
            self.txt_log.see(tk.END)
            self.txt_log.config(state=tk.DISABLED)
        self.root.after(0, ins)

    def _clear_log(self):
        self.txt_log.config(state=tk.NORMAL)
        self.txt_log.delete("1.0", tk.END)
        self.txt_log.config(state=tk.DISABLED)

    # ---------- Keyboard ----------
    def _on_key(self, ev):
        if (self.bot and self.bot.running
                and not self.bot.paused
                and not self.bot._suppress):
            self.bot.paused = True
            self.bot.log(f"⏸ ПАУЗА (клавиша: {ev.name})")
            self.root.after(0, lambda: self._set_mode("paused"))
            self.root.after(0, lambda: self.btn_pause.config(text="▶  Продолжить"))

    # ---------- Controls ----------
    def _start(self):
        xls = self.var_xls.get().strip()
        if xls and os.path.exists(xls):
            self.zones.load(xls)
            self._show_zones()

        fb = self.var_fb.get().strip() or FALLBACK_LOCATION

        self.bot = WMSBot(
            self.zones, fb, self._sel_pid,
            log_cb=self._log,
            status_cb=lambda m: self.root.after(0, lambda: self._set_mode(m))
        )
        try:
            self.bot.poll = float(self.var_poll.get())
        except:
            pass

        self.thread = threading.Thread(target=self.bot.run, daemon=True)
        self.thread.start()

        self.btn_start.config(state=tk.DISABLED, bg="#374151")
        self.btn_pause.config(state=tk.NORMAL, text="⏸  Пауза")
        self.btn_stop.config(state=tk.NORMAL)
        self._set_mode("running")

    def _pause(self):
        if self.bot:
            self.bot.toggle_pause()
            if self.bot.paused:
                self._set_mode("paused")
                self.btn_pause.config(text="▶  Продолжить")
            else:
                self._set_mode("running")
                self.btn_pause.config(text="⏸  Пауза")

    def _stop(self):
        if self.bot:
            self.bot.stop()
        self.btn_start.config(state=tk.NORMAL, bg=C_GREEN)
        self.btn_pause.config(state=tk.DISABLED, text="⏸  Пауза")
        self.btn_stop.config(state=tk.DISABLED)
        self._anim = False
        self._set_mode("stopped")

    def _close(self):
        if self.bot:
            self.bot.stop()
        try:
            keyboard.unhook(self._kb_hook)
        except:
            pass
        self._anim = False
        self.root.destroy()

    def run(self):
        self.root.protocol("WM_DELETE_WINDOW", self._close)
        self._draw_bar()
        self.root.mainloop()


# ===================================================================
if __name__ == "__main__":
    App().run()
