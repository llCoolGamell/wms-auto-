"""
WMS Automation Tool v3.0
========================
Автоматизация складской WMS-системы.
Определяет текущее окно и выполняет нужное действие.
Может начать с любого шага цикла.

Управление:
  Любая клавиша — пауза
  Кнопки интерфейса — управление

Установка:
  pip install pywinauto openpyxl keyboard

Запуск (от имени администратора!):
  python wms_auto.py
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


# ---------------------------------------------------------------------------
#  Logging
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
#  Constants & Colors
# ---------------------------------------------------------------------------
S_UNKNOWN            = "UNKNOWN"
S_MAIN_MENU          = "ГЛАВНОЕ_МЕНЮ"
S_TAKE_WORK          = "ВЗЯТИЕ_РАБОТЫ"
S_MOVE_TO_SOURCE     = "ПЕРЕМЕЩЕНИЕ_К_ИСТОЧНИКУ"
S_SEARCH_PALLET      = "ПОИСК_ПАЛЕТЫ"
S_SEARCH_BOX         = "ПОИСК_КОРОБКИ"
S_SEARCH_DEST_BOX    = "ПОИСК_МЕСТА_КОРОБКА"
S_SEARCH_DEST_PALLET = "ПОИСК_МЕСТА_ПАЛЕТА"
S_PLACE_IN_LOC       = "РАЗМЕЩЕНИЕ_В_МЕСТО"
S_CONFIRMATION       = "ПОДТВЕРЖДЕНИЕ"
S_ERROR              = "ОШИБКА"

FALLBACK_LOCATION = "D-KM-1"

# --- UI Colors ---
BG_DARK       = "#1a1a2e"
BG_CARD       = "#16213e"
BG_INPUT      = "#1c2a4a"
FG_TEXT        = "#e0e0e0"
FG_DIM         = "#7a7a9a"
COLOR_BLUE     = "#2563eb"
COLOR_BLUE_D   = "#1d4ed8"
COLOR_GREEN    = "#16a34a"
COLOR_GREEN_L  = "#22c55e"
COLOR_GREEN_D  = "#15803d"
COLOR_RED      = "#dc2626"
COLOR_RED_L    = "#ef4444"
COLOR_YELLOW   = "#eab308"
COLOR_YELLOW_L = "#facc15"
COLOR_CYAN     = "#06b6d4"
BOTTOM_BAR     = "#111827"


# ===================================================================
#  Zone Lookup from Excel
# ===================================================================
class ZoneLookup:
    """
    Читает Excel-файл (первый лист):
      столбец A — ключевое слово / название зоны
      столбец B — код места
    """

    def __init__(self, path=None):
        self.mapping = {}           # {lower_key: code}
        self.display_mapping = {}   # {original_key: code}  для отображения
        self.source_path = None
        if path:
            self.load(path)

    def load(self, path):
        if not os.path.exists(path):
            logger.warning("Файл зон не найден: %s", path)
            return False
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            ws = wb.active
            self.mapping.clear()
            self.display_mapping.clear()
            for row in ws.iter_rows(min_row=1, values_only=True):
                if row and len(row) >= 2 and row[0] and row[1]:
                    key_orig = str(row[0]).strip()
                    val = str(row[1]).strip()
                    self.mapping[key_orig.lower()] = val
                    self.display_mapping[key_orig] = val
            wb.close()
            self.source_path = path
            logger.info("Загружено %d зон из %s", len(self.mapping), path)
            return True
        except Exception as exc:
            logger.error("Ошибка загрузки зон: %s", exc)
            return False

    def lookup(self, zone_text):
        """Ищет подходящий код места для текста зоны."""
        zt = zone_text.lower().strip()

        # 1. Точное совпадение
        if zt in self.mapping:
            return self.mapping[zt]

        # 2. Ключ содержится в zone_text
        for key, code in self.mapping.items():
            if key in zt:
                return code

        # 3. zone_text содержится в ключе
        for key, code in self.mapping.items():
            if zt in key:
                return code

        return None

    def auto_find_and_load(self):
        """Автоматически ищет и загружает xlsx-файл зон из папки скрипта."""
        script_dir = os.path.dirname(os.path.abspath(__file__))

        # Сначала известные имена
        for name in ["zones.xlsx", "зоны.xlsx", "зона.xlsx", "zone.xlsx",
                      "файл зон.xlsx", "Файл зон.xlsx"]:
            path = os.path.join(script_dir, name)
            if os.path.exists(path):
                if self.load(path):
                    return path

        # Любой xlsx в папке
        for f in sorted(glob.glob(os.path.join(script_dir, "*.xlsx"))):
            if self.load(f):
                return f

        return None


# ===================================================================
#  Window helpers
# ===================================================================
def find_window_by_title(pattern):
    """Ищет окно, в заголовке которого содержится pattern."""
    try:
        desktop = Desktop(backend="win32")
        for w in desktop.windows():
            try:
                if pattern in w.window_text():
                    return w
            except Exception:
                continue
    except Exception:
        pass
    return None


def get_children_by_class(window, cls_name):
    """Возвращает дочерние контролы определённого класса."""
    result = []
    try:
        for child in window.children():
            try:
                cn = child.class_name()
                if cls_name.lower() in cn.lower():
                    result.append(child)
            except Exception:
                continue
    except Exception:
        pass
    return result


def get_edits(window):
    return get_children_by_class(window, "Edit")


def get_statics(window):
    texts = []
    for ctrl in get_children_by_class(window, "Static"):
        try:
            texts.append(ctrl.window_text())
        except Exception:
            pass
    return texts


def find_labeled_edit(window, label_text):
    """
    Находит Edit-контрол, расположенный на той же строке,
    что и Static-метка содержащая label_text.
    Сопоставление по вертикальной близости (±25 px).
    """
    statics = get_children_by_class(window, "Static")
    edits = get_children_by_class(window, "Edit")

    target_static = None
    for s in statics:
        try:
            txt = s.window_text().strip()
            if label_text.lower() in txt.lower():
                target_static = s
                break
        except Exception:
            continue

    if not target_static:
        return None

    try:
        sr = target_static.rectangle()
        static_cy = (sr.top + sr.bottom) // 2
    except Exception:
        return None

    best_edit = None
    best_dist = 999999

    for ed in edits:
        try:
            er = ed.rectangle()
            edit_cy = (er.top + er.bottom) // 2
            dist = abs(edit_cy - static_cy)
            if dist < 25 and dist < best_dist:
                best_dist = dist
                best_edit = ed
        except Exception:
            continue

    return best_edit


def click_ok(window):
    """Нажимает кнопку Ок / Ok / OK в окне."""
    for ctrl in get_children_by_class(window, "Button"):
        try:
            txt = ctrl.window_text().strip().lower()
            if txt in ("ок", "ok", "оk", "oк"):
                ctrl.click()
                return True
        except Exception:
            continue
    # Fallback — Enter
    try:
        window.set_focus()
        time.sleep(0.05)
        keyboard.send("enter")
        return True
    except Exception:
        return False


def click_button(window, text_part):
    for ctrl in get_children_by_class(window, "Button"):
        try:
            if text_part.lower() in ctrl.window_text().lower():
                ctrl.click()
                return True
        except Exception:
            continue
    return False


def set_edit_value(edit_ctrl, value):
    """Очищает поле и вводит значение."""
    try:
        edit_ctrl.set_focus()
        time.sleep(0.05)
        edit_ctrl.set_edit_text("")
        time.sleep(0.05)
        edit_ctrl.set_edit_text(value)
    except Exception as exc:
        logger.error("set_edit_value error: %s", exc)
        try:
            edit_ctrl.set_focus()
            keyboard.send("ctrl+a")
            time.sleep(0.05)
            keyboard.write(value, delay=0.02)
        except Exception:
            pass


def find_source_and_control(edits):
    """Fallback: первый пустой Edit = контроль, предыдущий непустой = источник."""
    for i, ed in enumerate(edits):
        val = ed.window_text().strip()
        if not val:
            source_val = None
            for j in range(i - 1, -1, -1):
                sv = edits[j].window_text().strip()
                if sv:
                    source_val = sv
                    break
            return source_val, ed
    return None, None


# ===================================================================
#  Core Automation
# ===================================================================
class WMSAutomation:

    def __init__(self, zone_lookup, log_cb=None, status_cb=None):
        self.zone_lookup = zone_lookup
        self.log_cb = log_cb
        self.status_cb = status_cb
        self.running = False
        self.paused = False
        self.poll_interval = 0.5
        self.action_delay = 0.4
        self.fallback = FALLBACK_LOCATION
        self._suppress_kb = False   # подавление паузы при программном нажатии

    def log(self, msg, level="INFO"):
        logger.log(getattr(logging, level, logging.INFO), msg)
        if self.log_cb:
            self.log_cb("[{}] {}".format(level, msg))

    def _safe_send(self, keys):
        """Отправить клавишу с подавлением авто-паузы."""
        self._suppress_kb = True
        try:
            keyboard.send(keys)
        finally:
            time.sleep(0.05)
            self._suppress_kb = False

    # ---------- state detection ----------
    def detect_state(self):
        """Определяет текущее состояние по заголовку окна."""

        # 0. Подтверждения
        for title in ("Подтверждение", "Подтвердить", "Confirm",
                       "Информация", "Information", "Сообщение", "Message"):
            w = find_window_by_title(title)
            if w:
                return S_CONFIRMATION, w

        # 1. Ошибки
        for title in ("Ошибка", "Error", "Предупреждение", "Warning"):
            w = find_window_by_title(title)
            if w:
                return S_ERROR, w

        # 2. Размещение в место
        w = find_window_by_title("Размещение в место")
        if w:
            return S_PLACE_IN_LOC, w

        # 3. Поиск места-приёмника
        w = find_window_by_title("Поиск места-приёмника")
        if w:
            statics = " ".join(get_statics(w)).lower()
            if "паллет" in statics or "содержимого" in statics:
                return S_SEARCH_DEST_PALLET, w
            return S_SEARCH_DEST_BOX, w

        # 4. Поиск коробки
        w = find_window_by_title("Поиск коробки")
        if w:
            return S_SEARCH_BOX, w

        # 5. Поиск палеты / паллеты
        for t in ("Поиск палеты", "Поиск паллеты"):
            w = find_window_by_title(t)
            if w:
                return S_SEARCH_PALLET, w

        # 6. Перемещение к источнику
        w = find_window_by_title("Перемещение к источнику")
        if w:
            return S_MOVE_TO_SOURCE, w

        # 7. Взятие работы
        w = find_window_by_title("Взятие работы")
        if w:
            return S_TAKE_WORK, w

        # 8. Главное меню
        w = find_window_by_title("Главное меню")
        if w:
            return S_MAIN_MENU, w

        return S_UNKNOWN, None

    # ---------- helpers ----------

    def _press_f2(self, window):
        try:
            clicked = click_button(window, "F2")
            if not clicked:
                clicked = click_button(window, "Запросить")
            if not clicked:
                clicked = click_button(window, "Взять")
            if not clicked:
                window.set_focus()
                time.sleep(0.05)
                self._safe_send("F2")
        except Exception as exc:
            self.log("_press_f2 error: {}".format(exc), "ERROR")

    def _check_confirmation(self):
        """После действия — проверяем, не вылезло ли подтверждение."""
        time.sleep(0.4)
        st, w = self.detect_state()
        if st == S_CONFIRMATION:
            self.handle_confirmation(w)
            return True
        if st == S_ERROR:
            self.handle_error(w)
            return True
        return False

    # ---------- handlers ----------

    def handle_main_menu(self, w):
        self.log("→ Главное меню: нажимаю F2 (Запросить работу)")
        self._press_f2(w)

    def handle_take_work(self, w):
        self.log("→ Взятие работы: нажимаю F2 (Взять работу)")
        self._press_f2(w)

    def handle_move_to_source(self, w):
        self.log("→ Перемещение к источнику")
        edits = get_edits(w)
        edit_vals = [e.window_text().strip() for e in edits]
        self.log("   Поля: {}".format(edit_vals))

        # Labeled: «Место» / «Из места» → «Контроль»
        source_edit = find_labeled_edit(w, "Место")
        control_edit = find_labeled_edit(w, "Контроль")

        if source_edit and control_edit:
            src_val = source_edit.window_text().strip()
            if src_val:
                self.log("   Место='{}' → Контроль".format(src_val))
                set_edit_value(control_edit, src_val)
                time.sleep(0.1)
                click_ok(w)
                self._check_confirmation()
                return

        # Fallback
        src, ctrl = find_source_and_control(edits)
        if src and ctrl:
            self.log("   (fallback) '{}' → Контроль".format(src))
            set_edit_value(ctrl, src)
            time.sleep(0.1)
            click_ok(w)
            self._check_confirmation()
        else:
            self.log("   Не удалось найти Место/Контроль", "WARNING")

    def handle_search_pallet(self, w):
        """
        ИСПРАВЛЕНО (v3):
        Копируем именно поле «Паллета» (не «Из места»!) в «Контроль».
        Используем find_labeled_edit для точного определения по метке.
        """
        self.log("→ Поиск палеты")
        edits = get_edits(w)
        edit_vals = [e.window_text().strip() for e in edits]
        self.log("   Поля: {}".format(edit_vals))

        # *** КЛЮЧЕВОЕ ИСПРАВЛЕНИЕ: ищем поле именно по метке «Паллет» ***
        pallet_edit = find_labeled_edit(w, "Паллет")
        control_edit = find_labeled_edit(w, "Контроль")

        if pallet_edit and control_edit:
            pallet_val = pallet_edit.window_text().strip()
            if pallet_val:
                self.log("   Паллета='{}' → Контроль  ✓".format(pallet_val))
                set_edit_value(control_edit, pallet_val)
                time.sleep(0.1)
                click_ok(w)
                self._check_confirmation()
                return
            else:
                self.log("   Поле 'Паллета' пустое!", "WARNING")
        else:
            self.log("   Labeled поиск не удался (Паллет={}, Контроль={})".format(
                pallet_edit is not None, control_edit is not None), "WARNING")

        # Fallback — старый метод (берёт предыдущий непустой перед первым пустым)
        self.log("   Использую fallback-поиск...")
        src, ctrl = find_source_and_control(edits)
        if src and ctrl:
            self.log("   (fallback) '{}' → Контроль".format(src))
            set_edit_value(ctrl, src)
            time.sleep(0.1)
            click_ok(w)
            self._check_confirmation()
        else:
            self.log("   Не удалось найти Паллета/Контроль", "WARNING")

    def handle_search_box(self, w):
        self.log("→ Поиск коробки")
        edits = get_edits(w)
        edit_vals = [e.window_text().strip() for e in edits]
        self.log("   Поля: {}".format(edit_vals))

        # Labeled: «Коробка» → «Контроль»
        box_edit = find_labeled_edit(w, "Коробк")
        control_edit = find_labeled_edit(w, "Контроль")

        if box_edit and control_edit:
            box_val = box_edit.window_text().strip()
            if box_val:
                self.log("   Коробка='{}' → Контроль".format(box_val))
                set_edit_value(control_edit, box_val)
                time.sleep(0.1)
                click_ok(w)
                self._check_confirmation()
                return

        # Fallback
        src, ctrl = find_source_and_control(edits)
        if src and ctrl:
            self.log("   (fallback) '{}' → Контроль".format(src))
            set_edit_value(ctrl, src)
            time.sleep(0.1)
            click_ok(w)
            self._check_confirmation()
        else:
            self.log("   Не удалось найти Коробка/Контроль", "WARNING")

    def handle_search_dest_box(self, w):
        """Поиск места-приёмника (коробка). Зона → код из XLS → fallback."""
        self.log("→ Поиск места-приёмника (коробка)")
        statics = get_statics(w)
        full_text = " ".join(statics)
        self.log("   Текст: {}".format(full_text))

        edits = get_edits(w)
        edit_vals = [e.window_text().strip() for e in edits]
        self.log("   Поля: {}".format(edit_vals))

        # Извлекаем зону
        zone_text = ""
        m = re.search(r"[Зз]она[:\s]+(.+)", full_text)
        if m:
            zone_text = m.group(1).strip()
        else:
            m2 = re.search(r"(Док\s+.+)", full_text)
            if m2:
                zone_text = m2.group(1).strip()

        self.log("   Зона: '{}'".format(zone_text))

        # Ищем код в справочнике
        location = None
        if zone_text:
            location = self.zone_lookup.lookup(zone_text)

        if location:
            self.log("   Найден код: {}".format(location))
        else:
            location = self.fallback
            self.log("   Зона не найдена → fallback {}".format(location), "WARNING")

        # Вставляем в Контроль
        ctrl = find_labeled_edit(w, "Контроль")
        if not ctrl:
            for ed in edits:
                if not ed.window_text().strip():
                    ctrl = ed
                    break
        if not ctrl and edits:
            ctrl = edits[-1]

        if ctrl:
            set_edit_value(ctrl, location)
            time.sleep(0.15)
            click_ok(w)

            # Проверяем ошибку / подтверждение
            time.sleep(0.8)
            st, ew = self.detect_state()
            if st == S_ERROR:
                self.log("   Ошибка! Закрываю и пробую fallback", "WARNING")
                click_ok(ew)
                time.sleep(0.5)
                st2, w2 = self.detect_state()
                if st2 in (S_SEARCH_DEST_BOX, S_SEARCH_DEST_PALLET):
                    ctrl2 = find_labeled_edit(w2, "Контроль")
                    if not ctrl2:
                        edits2 = get_edits(w2)
                        for ed in edits2:
                            if not ed.window_text().strip():
                                ctrl2 = ed
                                break
                        if not ctrl2 and edits2:
                            ctrl2 = edits2[-1]
                    if ctrl2:
                        self.log("   Вставляю fallback: {}".format(self.fallback))
                        set_edit_value(ctrl2, self.fallback)
                        time.sleep(0.15)
                        click_ok(w2)
                        self._check_confirmation()
            elif st == S_CONFIRMATION:
                self.handle_confirmation(ew)
        else:
            self.log("   Не найдено поле Контроль", "WARNING")

    def handle_search_dest_pallet(self, w):
        """Поиск места-приёмника (паллета): всегда fallback."""
        self.log("→ Поиск места-приёмника (паллета): {}".format(self.fallback))
        edits = get_edits(w)

        ctrl = find_labeled_edit(w, "Контроль")
        if not ctrl:
            for ed in edits:
                if not ed.window_text().strip():
                    ctrl = ed
                    break
        if not ctrl and edits:
            ctrl = edits[-1]

        if ctrl:
            set_edit_value(ctrl, self.fallback)
            time.sleep(0.15)
            click_ok(w)
            self._check_confirmation()
        else:
            self.log("   Не найдено поле Контроль", "WARNING")

    def handle_place_in_loc(self, w):
        self.log("→ Размещение в место: нажимаю Ок")
        click_ok(w)
        self._check_confirmation()

    def handle_confirmation(self, w):
        """Обработка окна подтверждения — сразу Ок."""
        self.log("→ Подтверждение: нажимаю Ок")
        click_ok(w)
        # Цепочка: может вылезти ещё одно
        time.sleep(0.3)
        st, w2 = self.detect_state()
        if st == S_CONFIRMATION:
            self.log("→ Ещё подтверждение: нажимаю Ок")
            click_ok(w2)

    def handle_error(self, w):
        try:
            statics = get_statics(w)
            self.log("→ Ошибка: {}".format(" ".join(statics)))
        except Exception:
            self.log("→ Ошибка (текст не прочитан)")
        click_ok(w)

    # ---------- main loop ----------
    def run(self):
        self.running = True
        self.log("=" * 50)
        self.log("Автоматизация запущена")
        self.log("Fallback: {}  |  Зон: {}".format(
            self.fallback, len(self.zone_lookup.mapping)))
        self.log("=" * 50)

        handlers = {
            S_MAIN_MENU:          self.handle_main_menu,
            S_TAKE_WORK:          self.handle_take_work,
            S_MOVE_TO_SOURCE:     self.handle_move_to_source,
            S_SEARCH_PALLET:      self.handle_search_pallet,
            S_SEARCH_BOX:         self.handle_search_box,
            S_SEARCH_DEST_BOX:    self.handle_search_dest_box,
            S_SEARCH_DEST_PALLET: self.handle_search_dest_pallet,
            S_PLACE_IN_LOC:       self.handle_place_in_loc,
            S_CONFIRMATION:       self.handle_confirmation,
            S_ERROR:              self.handle_error,
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

                state, window = self.detect_state()

                if state == S_UNKNOWN:
                    time.sleep(self.poll_interval)
                    continue

                handler = handlers.get(state)
                if handler:
                    handler(window)

                time.sleep(self.action_delay)

            except Exception as exc:
                self.log("Ошибка в цикле: {}".format(exc), "ERROR")
                time.sleep(1.0)

        if self.status_cb:
            self.status_cb("stopped")
        self.log("=" * 50)
        self.log("Автоматизация остановлена")
        self.log("=" * 50)

    def stop(self):
        self.running = False

    def toggle_pause(self):
        self.paused = not self.paused
        self.log("⏸ ПАУЗА" if self.paused else "▶ ПРОДОЛЖЕНИЕ")


# ===================================================================
#  GUI — Modern Dark Theme
# ===================================================================
class AppGUI:

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("WMS Автоматизация v3.0")
        self.root.geometry("920x680")
        self.root.configure(bg=BG_DARK)
        self.root.resizable(True, True)

        # Иконка (если доступна)
        try:
            self.root.iconbitmap(default="")
        except Exception:
            pass

        self.automation = None
        self.thread = None
        self.zone_lookup = ZoneLookup()

        self._arrow_x = -100
        self._anim_running = False
        self._status_mode = "stopped"  # stopped | paused | running

        self._configure_styles()
        self._build_ui()

        # Автозагрузка зон
        self.root.after(200, self._auto_load_zones)

        # Хук: любая клавиша → пауза
        self._kb_hook = keyboard.on_press(self._on_any_key_press)

    # ----------------------------------------------------------------
    def _configure_styles(self):
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("Dark.TFrame", background=BG_DARK)
        style.configure("Card.TFrame", background=BG_CARD)

        style.configure("Dark.TLabel", background=BG_DARK,
                         foreground=FG_TEXT, font=("Segoe UI", 10))
        style.configure("Card.TLabel", background=BG_CARD,
                         foreground=FG_TEXT, font=("Segoe UI", 10))
        style.configure("Dim.TLabel", background=BG_DARK,
                         foreground=FG_DIM, font=("Segoe UI", 9))

        style.configure("Dark.TLabelframe", background=BG_CARD,
                         foreground=COLOR_CYAN)
        style.configure("Dark.TLabelframe.Label", background=BG_CARD,
                         foreground=COLOR_CYAN, font=("Segoe UI", 10, "bold"))

    # ----------------------------------------------------------------
    def _build_ui(self):

        # ===== ВЕРХНЯЯ ПАНЕЛЬ СТАТУСА =====
        self.status_canvas = tk.Canvas(
            self.root, height=54, bg=COLOR_BLUE, highlightthickness=0
        )
        self.status_canvas.pack(fill=tk.X, side=tk.TOP)
        self.status_canvas.bind("<Configure>", lambda e: self._draw_status())

        # ===== ОСНОВНАЯ ОБЛАСТЬ =====
        main_frame = ttk.Frame(self.root, style="Dark.TFrame")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=6)

        # --- Настройки ---
        frm_set = ttk.LabelFrame(main_frame, text="⚙  Настройки",
                                   style="Dark.TLabelframe")
        frm_set.pack(fill=tk.X, pady=(0, 6))

        row1 = ttk.Frame(frm_set, style="Card.TFrame")
        row1.pack(fill=tk.X, padx=8, pady=4)
        ttk.Label(row1, text="📁 Файл зон:", style="Card.TLabel").pack(side=tk.LEFT)
        self.var_xls = tk.StringVar()
        tk.Entry(row1, textvariable=self.var_xls, width=52,
                 bg=BG_INPUT, fg=FG_TEXT, insertbackground=FG_TEXT,
                 relief="flat", font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=8)
        tk.Button(row1, text="Обзор…", command=self._browse_xls,
                  bg=COLOR_BLUE, fg="white", activebackground=COLOR_BLUE_D,
                  relief="flat", font=("Segoe UI", 9, "bold"),
                  padx=14, cursor="hand2").pack(side=tk.LEFT)

        row2 = ttk.Frame(frm_set, style="Card.TFrame")
        row2.pack(fill=tk.X, padx=8, pady=4)
        ttk.Label(row2, text="📍 Fallback:", style="Card.TLabel").pack(side=tk.LEFT)
        self.var_fallback = tk.StringVar(value=FALLBACK_LOCATION)
        tk.Entry(row2, textvariable=self.var_fallback, width=12,
                 bg=BG_INPUT, fg=FG_TEXT, insertbackground=FG_TEXT,
                 relief="flat", font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=8)
        ttk.Label(row2, text="⏱ Интервал (сек):", style="Card.TLabel") \
            .pack(side=tk.LEFT, padx=(20, 0))
        self.var_interval = tk.StringVar(value="0.5")
        tk.Entry(row2, textvariable=self.var_interval, width=6,
                 bg=BG_INPUT, fg=FG_TEXT, insertbackground=FG_TEXT,
                 relief="flat", font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=8)
        self.lbl_zones_count = ttk.Label(row2, text="Зон: 0", style="Card.TLabel")
        self.lbl_zones_count.pack(side=tk.RIGHT, padx=10)

        # --- Таблица зон (для вставки образцы) ---
        frm_zones = ttk.LabelFrame(
            main_frame,
            text="📋  Для вставки образцы (загруженные зоны)",
            style="Dark.TLabelframe"
        )
        frm_zones.pack(fill=tk.X, pady=(0, 6))

        zones_inner = ttk.Frame(frm_zones, style="Card.TFrame")
        zones_inner.pack(fill=tk.X, padx=5, pady=5)
        self.zones_text = tk.Text(
            zones_inner, height=4, bg=BG_INPUT, fg=FG_TEXT,
            font=("Consolas", 9), relief="flat",
            insertbackground=FG_TEXT, state=tk.DISABLED,
            selectbackground=COLOR_BLUE
        )
        z_scroll = ttk.Scrollbar(zones_inner, command=self.zones_text.yview)
        self.zones_text.configure(yscrollcommand=z_scroll.set)
        z_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.zones_text.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # --- Лог ---
        frm_log = ttk.LabelFrame(main_frame, text="📝  Лог",
                                   style="Dark.TLabelframe")
        frm_log.pack(fill=tk.BOTH, expand=True, pady=(0, 2))

        self.txt_log = scrolledtext.ScrolledText(
            frm_log, height=10, state=tk.DISABLED,
            font=("Consolas", 9), bg="#0d1117", fg="#c9d1d9",
            insertbackground="#c9d1d9", relief="flat",
            selectbackground=COLOR_BLUE
        )
        self.txt_log.pack(fill=tk.BOTH, expand=True, padx=5, pady=(5, 2))

        tk.Button(frm_log, text="🗑 Очистить лог", command=self._clear_log,
                  bg="#374151", fg=FG_TEXT, activebackground="#4b5563",
                  relief="flat", font=("Segoe UI", 8),
                  cursor="hand2").pack(anchor=tk.W, padx=5, pady=(0, 5))

        # ===== НИЖНЯЯ ПАНЕЛЬ =====
        bottom = tk.Frame(self.root, bg=BOTTOM_BAR)
        bottom.pack(fill=tk.X, side=tk.BOTTOM)

        # Подсказка
        tk.Label(
            bottom,
            text="💡 Любая клавиша = пауза   |   Кнопки ниже для управления",
            bg=BOTTOM_BAR, fg=FG_DIM, font=("Segoe UI", 9)
        ).pack(side=tk.TOP, pady=(8, 3))

        btn_frame = tk.Frame(bottom, bg=BOTTOM_BAR)
        btn_frame.pack(pady=(0, 10))

        self.btn_start = tk.Button(
            btn_frame, text="▶  Начать перемещение", command=self._on_start,
            bg=COLOR_GREEN, fg="white", activebackground=COLOR_GREEN_L,
            relief="flat", font=("Segoe UI", 12, "bold"),
            padx=24, pady=8, cursor="hand2"
        )
        self.btn_start.pack(side=tk.LEFT, padx=10)

        self.btn_pause = tk.Button(
            btn_frame, text="⏸  Пауза", command=self._on_pause,
            bg=COLOR_YELLOW, fg="#1a1a2e", activebackground=COLOR_YELLOW_L,
            relief="flat", font=("Segoe UI", 12, "bold"),
            padx=24, pady=8, cursor="hand2", state=tk.DISABLED
        )
        self.btn_pause.pack(side=tk.LEFT, padx=10)

        self.btn_exit = tk.Button(
            btn_frame, text="✕  Выход", command=self._on_close,
            bg=COLOR_RED, fg="white", activebackground=COLOR_RED_L,
            relief="flat", font=("Segoe UI", 12, "bold"),
            padx=24, pady=8, cursor="hand2"
        )
        self.btn_exit.pack(side=tk.LEFT, padx=10)

    # ================================================================
    #  Status bar animation
    # ================================================================
    def _draw_status(self):
        c = self.status_canvas
        c.delete("all")
        w = c.winfo_width()
        h = c.winfo_height()

        if self._status_mode == "running":
            c.configure(bg=COLOR_GREEN)
            # Бегущие стрелки
            for i in range(10):
                ax = self._arrow_x + i * 50
                if -30 < ax < w + 30:
                    # Полупрозрачные «›» (более тёмный зелёный)
                    c.create_text(
                        ax, h // 2, text="›",
                        fill=COLOR_GREEN_L,
                        font=("Segoe UI", 30, "bold")
                    )
            # Надпись всегда по центру
            c.create_text(
                w // 2, h // 2, text="▶  ИДЁТ ПЕРЕМЕЩЕНИЕ",
                fill="white", font=("Segoe UI", 16, "bold")
            )

        elif self._status_mode == "paused":
            c.configure(bg=COLOR_BLUE)
            c.create_text(
                w // 2, h // 2, text="⏸  ПАУЗА",
                fill="white", font=("Segoe UI", 16, "bold")
            )

        else:  # stopped
            c.configure(bg=COLOR_BLUE)
            c.create_text(
                w // 2, h // 2, text="⏹  ПРОГРАММА НЕ ЗАПУЩЕНА",
                fill="white", font=("Segoe UI", 16, "bold")
            )

    def _start_animation(self):
        if self._anim_running:
            return
        self._anim_running = True
        self._animate_tick()

    def _stop_animation(self):
        self._anim_running = False

    def _animate_tick(self):
        if not self._anim_running:
            return
        if self._status_mode == "running":
            w = self.status_canvas.winfo_width() or 920
            self._arrow_x += 4
            if self._arrow_x > w + 50:
                self._arrow_x = -500
        self._draw_status()
        self.root.after(50, self._animate_tick)

    def _set_status_mode(self, mode):
        """mode: 'stopped' | 'paused' | 'running'"""
        self._status_mode = mode
        if mode == "running":
            self._start_animation()
        else:
            self._draw_status()

    # ================================================================
    #  Auto-load zones
    # ================================================================
    def _auto_load_zones(self):
        path = self.zone_lookup.auto_find_and_load()
        if path:
            self.var_xls.set(path)
            self._add_log("✅ Автозагрузка зон: {}".format(os.path.basename(path)))
            self._update_zones_display()
        else:
            self._add_log("ℹ️  Файл зон не найден автоматически. Загрузите через «Обзор».")

    def _update_zones_display(self):
        n = len(self.zone_lookup.mapping)
        self.lbl_zones_count.config(text="Зон: {}".format(n))
        self.zones_text.config(state=tk.NORMAL)
        self.zones_text.delete("1.0", tk.END)
        if self.zone_lookup.display_mapping:
            for key, val in self.zone_lookup.display_mapping.items():
                self.zones_text.insert(tk.END, "  {}  →  {}\n".format(key, val))
        else:
            self.zones_text.insert(
                tk.END, "  (пусто — загрузите файл зон через «Обзор»)\n"
            )
        self.zones_text.config(state=tk.DISABLED)

    # ================================================================
    #  Callbacks
    # ================================================================
    def _browse_xls(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel", "*.xlsx *.xls"), ("All", "*.*")]
        )
        if path:
            self.var_xls.set(path)
            self.zone_lookup.load(path)
            self._add_log("📁 Загружен: {}".format(os.path.basename(path)))
            self._update_zones_display()

    def _add_log(self, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        def _insert():
            self.txt_log.config(state=tk.NORMAL)
            self.txt_log.insert(tk.END, "[{}] {}\n".format(ts, msg))
            self.txt_log.see(tk.END)
            self.txt_log.config(state=tk.DISABLED)
        self.root.after(0, _insert)

    def _clear_log(self):
        self.txt_log.config(state=tk.NORMAL)
        self.txt_log.delete("1.0", tk.END)
        self.txt_log.config(state=tk.DISABLED)

    def _status_callback(self, mode):
        """Вызывается из потока автоматизации."""
        self.root.after(0, lambda m=mode: self._set_status_mode(m))

    def _on_any_key_press(self, event):
        """Любая физическая клавиша → пауза."""
        if (self.automation
                and self.automation.running
                and not self.automation.paused
                and not self.automation._suppress_kb):
            self.automation.paused = True
            self.automation.log("⏸ ПАУЗА (клавиша: {})".format(event.name))
            self.root.after(0, lambda: self._set_status_mode("paused"))
            self.root.after(0, lambda: self.btn_pause.config(text="▶  Продолжить"))

    def _on_start(self):
        fb = self.var_fallback.get().strip() or FALLBACK_LOCATION

        xls = self.var_xls.get().strip()
        if xls and os.path.exists(xls):
            self.zone_lookup.load(xls)
            self._update_zones_display()

        self.automation = WMSAutomation(
            self.zone_lookup,
            log_cb=self._add_log,
            status_cb=self._status_callback
        )
        self.automation.fallback = fb
        try:
            self.automation.poll_interval = float(self.var_interval.get())
        except ValueError:
            pass

        self.thread = threading.Thread(target=self.automation.run, daemon=True)
        self.thread.start()

        self.btn_start.config(state=tk.DISABLED, bg="#374151")
        self.btn_pause.config(state=tk.NORMAL, text="⏸  Пауза")
        self._set_status_mode("running")

    def _on_pause(self):
        if self.automation:
            self.automation.toggle_pause()
            if self.automation.paused:
                self._set_status_mode("paused")
                self.btn_pause.config(text="▶  Продолжить")
            else:
                self._set_status_mode("running")
                self.btn_pause.config(text="⏸  Пауза")

    def _on_stop(self):
        if self.automation:
            self.automation.stop()
        self.btn_start.config(state=tk.NORMAL, bg=COLOR_GREEN)
        self.btn_pause.config(state=tk.DISABLED, text="⏸  Пауза")
        self._stop_animation()
        self._set_status_mode("stopped")

    def run(self):
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._draw_status()
        self.root.mainloop()

    def _on_close(self):
        if self.automation:
            self.automation.stop()
        try:
            keyboard.unhook(self._kb_hook)
        except Exception:
            pass
        self._stop_animation()
        self.root.destroy()


# ===================================================================
#  Entry point
# ===================================================================
if __name__ == "__main__":
    app = AppGUI()
    app.run()
