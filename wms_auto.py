"""
WMS Automation Tool v2.0
========================
Автоматизация складской WMS-системы.
Определяет текущее окно и выполняет нужное действие.
Может начать с любого шага цикла.

Управление:
  Ctrl+Q  — остановить
  Ctrl+P  — пауза / продолжить

Установка:
  pip install -r requirements.txt

Запуск (от имени администратора!):
  python wms_auto.py
"""

import time
import sys
import os
import re
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
#  Constants
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
S_ERROR              = "ОШИБКА"

FALLBACK_LOCATION = "D-KM-1"


# ===================================================================
#  Zone Lookup from Excel
# ===================================================================
class ZoneLookup:
    """
    Читает Excel-файл:
      столбец A — ключевое слово / название зоны
      столбец B — код места
    Пример:
      Док Аша Дальняя  |  D-KM-1
      Отгрузка          |  D-KM-1
    """

    def __init__(self, path=None):
        self.mapping = {}
        if path:
            self.load(path)

    def load(self, path):
        if not os.path.exists(path):
            logger.warning("Файл зон не найден: %s", path)
            return
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            ws = wb.active
            self.mapping.clear()
            for row in ws.iter_rows(min_row=1, values_only=True):
                if row and len(row) >= 2 and row[0] and row[1]:
                    key = str(row[0]).strip().lower()
                    val = str(row[1]).strip()
                    self.mapping[key] = val
            wb.close()
            logger.info("Загружено %d зон из %s", len(self.mapping), path)
        except Exception as exc:
            logger.error("Ошибка загрузки зон: %s", exc)

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
    """
    Находит пару (source_value, control_edit).
    Контроль — первый пустой Edit.
    Source — непустой Edit прямо перед ним.
    """
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

    def __init__(self, zone_lookup, log_cb=None):
        self.zone_lookup = zone_lookup
        self.log_cb = log_cb
        self.running = False
        self.paused = False
        self.poll_interval = 0.5
        self.action_delay = 0.4
        self.fallback = FALLBACK_LOCATION

    def log(self, msg, level="INFO"):
        logger.log(getattr(logging, level, logging.INFO), msg)
        if self.log_cb:
            self.log_cb("[{}] {}".format(level, msg))

    # ---------- state detection ----------
    def detect_state(self):
        """Определяет текущее состояние по заголовку окна."""

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

    # ---------- handlers ----------

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
                keyboard.send("F2")
        except Exception as exc:
            self.log("_press_f2 error: {}".format(exc), "ERROR")

    def handle_main_menu(self, w):
        self.log("-> Главное меню: нажимаю F2 (Запросить работу)")
        self._press_f2(w)

    def handle_take_work(self, w):
        self.log("-> Взятие работы: нажимаю F2 (Взять работу)")
        self._press_f2(w)

    def handle_move_to_source(self, w):
        self.log("-> Перемещение к источнику")
        edits = get_edits(w)
        edit_vals = [e.window_text().strip() for e in edits]
        self.log("   Поля: {}".format(edit_vals))

        src, ctrl = find_source_and_control(edits)
        if src and ctrl:
            self.log("   Место='{}' -> Контроль".format(src))
            set_edit_value(ctrl, src)
            time.sleep(0.1)
            click_ok(w)
        else:
            self.log("   Не удалось найти Место/Контроль", "WARNING")

    def handle_search_pallet(self, w):
        self.log("-> Поиск палеты")
        edits = get_edits(w)
        edit_vals = [e.window_text().strip() for e in edits]
        self.log("   Поля: {}".format(edit_vals))

        src, ctrl = find_source_and_control(edits)
        if src and ctrl:
            self.log("   Палета='{}' -> Контроль".format(src))
            set_edit_value(ctrl, src)
            time.sleep(0.1)
            click_ok(w)
        else:
            self.log("   Не удалось найти Палета/Контроль", "WARNING")

    def handle_search_box(self, w):
        self.log("-> Поиск коробки")
        edits = get_edits(w)
        edit_vals = [e.window_text().strip() for e in edits]
        self.log("   Поля: {}".format(edit_vals))

        src, ctrl = find_source_and_control(edits)
        if src and ctrl:
            self.log("   Коробка='{}' -> Контроль".format(src))
            set_edit_value(ctrl, src)
            time.sleep(0.1)
            click_ok(w)
        else:
            self.log("   Не удалось найти Коробка/Контроль", "WARNING")

    def handle_search_dest_box(self, w):
        """
        Поиск места-приёмника (коробка).
        1. Извлечь зону из текста
        2. Найти код в XLS
        3. Если не найден — fallback D-KM-1
        4. Если ошибка после Ок — тоже fallback
        """
        self.log("-> Поиск места-приёмника (коробка)")
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
            self.log("   Зона не найдена -> fallback {}".format(location), "WARNING")

        # Вставляем в Контроль
        ctrl = None
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

            # Проверяем ошибку
            time.sleep(0.8)
            st, ew = self.detect_state()
            if st == S_ERROR:
                self.log("   Ошибка! Закрываю и пробую fallback", "WARNING")
                click_ok(ew)
                time.sleep(0.5)
                st2, w2 = self.detect_state()
                if st2 in (S_SEARCH_DEST_BOX, S_SEARCH_DEST_PALLET):
                    edits2 = get_edits(w2)
                    ctrl2 = None
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
        else:
            self.log("   Не найдено поле Контроль", "WARNING")

    def handle_search_dest_pallet(self, w):
        """Screen 12: Всегда D-KM-1."""
        self.log("-> Поиск места-приёмника (паллета): всегда {}".format(self.fallback))
        edits = get_edits(w)
        edit_vals = [e.window_text().strip() for e in edits]
        self.log("   Поля: {}".format(edit_vals))

        ctrl = None
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
        else:
            self.log("   Не найдено поле Контроль", "WARNING")

    def handle_place_in_loc(self, w):
        self.log("-> Размещение в место: нажимаю Ок")
        click_ok(w)

    def handle_error(self, w):
        try:
            statics = get_statics(w)
            self.log("-> Ошибка: {}".format(" ".join(statics)))
        except Exception:
            self.log("-> Ошибка (не удалось прочитать текст)")
        click_ok(w)

    # ---------- main loop ----------
    def run(self):
        self.running = True
        self.log("========== Автоматизация запущена ==========")
        self.log("Ctrl+Q = стоп  |  Ctrl+P = пауза")
        self.log("Fallback место: {}".format(self.fallback))
        self.log("Зон в справочнике: {}".format(len(self.zone_lookup.mapping)))

        handlers = {
            S_MAIN_MENU:          self.handle_main_menu,
            S_TAKE_WORK:          self.handle_take_work,
            S_MOVE_TO_SOURCE:     self.handle_move_to_source,
            S_SEARCH_PALLET:      self.handle_search_pallet,
            S_SEARCH_BOX:         self.handle_search_box,
            S_SEARCH_DEST_BOX:    self.handle_search_dest_box,
            S_SEARCH_DEST_PALLET: self.handle_search_dest_pallet,
            S_PLACE_IN_LOC:       self.handle_place_in_loc,
            S_ERROR:              self.handle_error,
        }

        while self.running:
            if self.paused:
                time.sleep(0.3)
                continue

            try:
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

        self.log("========== Автоматизация остановлена ==========")

    def stop(self):
        self.running = False

    def toggle_pause(self):
        self.paused = not self.paused
        self.log("ПАУЗА" if self.paused else "ПРОДОЛЖЕНИЕ")


# ===================================================================
#  GUI
# ===================================================================
class AppGUI:

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("WMS Автоматизация v2.0")
        self.root.geometry("750x550")
        self.root.resizable(True, True)

        self.automation = None
        self.thread = None
        self.zone_lookup = ZoneLookup()

        self._build_ui()

        keyboard.add_hotkey("ctrl+q", self._on_stop)
        keyboard.add_hotkey("ctrl+p", self._on_pause)

    def _build_ui(self):
        # --- Settings ---
        frm_set = ttk.LabelFrame(self.root, text="Настройки")
        frm_set.pack(fill=tk.X, padx=10, pady=5)

        r1 = ttk.Frame(frm_set)
        r1.pack(fill=tk.X, padx=5, pady=3)
        ttk.Label(r1, text="Файл зон (XLSX):").pack(side=tk.LEFT)
        self.var_xls = tk.StringVar()
        ttk.Entry(r1, textvariable=self.var_xls, width=50).pack(side=tk.LEFT, padx=5)
        ttk.Button(r1, text="Обзор...", command=self._browse_xls).pack(side=tk.LEFT)

        r2 = ttk.Frame(frm_set)
        r2.pack(fill=tk.X, padx=5, pady=3)
        ttk.Label(r2, text="Fallback место:").pack(side=tk.LEFT)
        self.var_fallback = tk.StringVar(value=FALLBACK_LOCATION)
        ttk.Entry(r2, textvariable=self.var_fallback, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Label(r2, text="Интервал (сек):").pack(side=tk.LEFT, padx=(20, 0))
        self.var_interval = tk.StringVar(value="0.5")
        ttk.Entry(r2, textvariable=self.var_interval, width=6).pack(side=tk.LEFT, padx=5)

        # --- Buttons ---
        frm_btn = ttk.Frame(self.root)
        frm_btn.pack(fill=tk.X, padx=10, pady=5)

        self.btn_start = ttk.Button(frm_btn, text="  Старт", command=self._on_start)
        self.btn_start.pack(side=tk.LEFT, padx=5)
        self.btn_pause = ttk.Button(frm_btn, text="  Пауза", command=self._on_pause,
                                     state=tk.DISABLED)
        self.btn_pause.pack(side=tk.LEFT, padx=5)
        self.btn_stop = ttk.Button(frm_btn, text="  Стоп", command=self._on_stop,
                                    state=tk.DISABLED)
        self.btn_stop.pack(side=tk.LEFT, padx=5)

        self.var_status = tk.StringVar(value="Остановлено")
        ttk.Label(frm_btn, textvariable=self.var_status,
                  font=("", 10, "bold")).pack(side=tk.RIGHT, padx=10)

        # --- Log ---
        frm_log = ttk.LabelFrame(self.root, text="Лог")
        frm_log.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.txt_log = scrolledtext.ScrolledText(
            frm_log, height=18, state=tk.DISABLED, font=("Consolas", 9)
        )
        self.txt_log.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        frm_lb = ttk.Frame(frm_log)
        frm_lb.pack(fill=tk.X, padx=5, pady=2)
        ttk.Button(frm_lb, text="Очистить лог", command=self._clear_log).pack(side=tk.LEFT)
        ttk.Label(frm_lb, text="Ctrl+Q=стоп | Ctrl+P=пауза",
                  foreground="gray").pack(side=tk.RIGHT)

    def _browse_xls(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel", "*.xlsx *.xls"), ("All", "*.*")]
        )
        if path:
            self.var_xls.set(path)
            self.zone_lookup.load(path)
            self._add_log("Загружен файл зон: {}".format(path))
            self._add_log("Зон: {}".format(len(self.zone_lookup.mapping)))
            for k, v in self.zone_lookup.mapping.items():
                self._add_log("   '{}' -> '{}'".format(k, v))

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

    def _on_start(self):
        fb = self.var_fallback.get().strip() or FALLBACK_LOCATION

        xls = self.var_xls.get().strip()
        if xls and os.path.exists(xls):
            self.zone_lookup.load(xls)

        self.automation = WMSAutomation(self.zone_lookup, log_cb=self._add_log)
        self.automation.fallback = fb
        try:
            self.automation.poll_interval = float(self.var_interval.get())
        except ValueError:
            pass

        self.thread = threading.Thread(target=self.automation.run, daemon=True)
        self.thread.start()

        self.btn_start.config(state=tk.DISABLED)
        self.btn_pause.config(state=tk.NORMAL)
        self.btn_stop.config(state=tk.NORMAL)
        self.var_status.set("Работает")

    def _on_pause(self):
        if self.automation:
            self.automation.toggle_pause()
            self.var_status.set("Пауза" if self.automation.paused else "Работает")

    def _on_stop(self):
        if self.automation:
            self.automation.stop()
        self.btn_start.config(state=tk.NORMAL)
        self.btn_pause.config(state=tk.DISABLED)
        self.btn_stop.config(state=tk.DISABLED)
        self.var_status.set("Остановлено")

    def run(self):
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.mainloop()

    def _on_close(self):
        if self.automation:
            self.automation.stop()
        self.root.destroy()


# ===================================================================
#  Entry point
# ===================================================================
if __name__ == "__main__":
    app = AppGUI()
    app.run()
