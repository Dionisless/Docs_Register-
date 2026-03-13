#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Регистратор таблиц из LanDocs
==============================
Извлекает данные из регистрационной карточки LanDocs (через Tab-навигацию
и буфер обмена) и записывает строку в журнал Excel.

Запуск: через горячую клавишу (см. hotkey.ahk) или ярлык.
При запуске окно LanDocs должно быть активным, курсор — в поле 0 (первое поле формы).

Позиции полей в форме LanDocs (количество нажатий Tab от начала):
  0  — № вх
  4  — ссылка на файл письма
  5  — Корреспондент (часть "За подписью")
  6  — Дата
  7  — № письма
  9  — Подписант (часть "За подписью")
  11 — Тема письма
  16 — Связанное письмо
"""

import os
import re
import sys
import time
import getpass
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

# ── Зависимости с мягкой обработкой отсутствия ──────────────────────────────

try:
    import win32api
    import win32con
    import win32clipboard
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

try:
    import openpyxl
    from openpyxl.styles import Alignment
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# ── Конфигурация ──────────────────────────────────────────────────────────────

EXCEL_PATH = r"\\Prim-fs-serv\primrdu\СРЗА\Журнал регистрации корреспонденции\Журнал регистрации входящей документации.xlsx"
DEFAULT_SAVE_FOLDER = r"\\Prim-fs-serv\primrdu\СРЗА\Дела СРЗА\19 Переписка"

# Задержка между нажатиями Tab (сек) — увеличьте если LanDocs тормозит
TAB_DELAY = 0.07
# Задержка после Ctrl+C перед чтением буфера обмена (сек)
COPY_DELAY = 0.12


# ── Работа с буфером обмена и клавиатурой ────────────────────────────────────

def _clear_clipboard():
    if not HAS_WIN32:
        return
    try:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
    except Exception:
        pass
    finally:
        try:
            win32clipboard.CloseClipboard()
        except Exception:
            pass


def _get_clipboard() -> str:
    if not HAS_WIN32:
        return ""
    try:
        win32clipboard.OpenClipboard()
        try:
            if win32clipboard.IsClipboardFormatAvailable(win32con.CF_UNICODETEXT):
                data = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
                return (data or "").strip()
        finally:
            win32clipboard.CloseClipboard()
    except Exception:
        pass
    return ""


def _send_tab():
    if not HAS_WIN32:
        return
    win32api.keybd_event(win32con.VK_TAB, 0, 0, 0)
    win32api.keybd_event(win32con.VK_TAB, 0, win32con.KEYEVENTF_KEYUP, 0)


def _send_ctrl_a():
    if not HAS_WIN32:
        return
    win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
    win32api.keybd_event(ord('A'), 0, 0, 0)
    win32api.keybd_event(ord('A'), 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)


def _send_ctrl_c():
    if not HAS_WIN32:
        return
    win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
    win32api.keybd_event(ord('C'), 0, 0, 0)
    win32api.keybd_event(ord('C'), 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)


def read_current_field() -> str:
    """Читает содержимое текущего поля через буфер обмена."""
    _clear_clipboard()
    time.sleep(0.04)
    _send_ctrl_a()
    time.sleep(0.04)
    _send_ctrl_c()
    time.sleep(COPY_DELAY)
    return _get_clipboard()


def navigate_tabs(n: int):
    """Посылает n нажатий Tab."""
    for _ in range(n):
        _send_tab()
        time.sleep(TAB_DELAY)


# ── Извлечение данных из LanDocs ──────────────────────────────────────────────

def extract_landocs_data() -> dict:
    """
    Извлекает поля из открытой регистрационной карточки LanDocs.
    Курсор должен быть установлен в первое поле формы (позиция 0).
    Возвращает словарь с ключами:
      incoming_num, file_link, correspondent, date,
      letter_num, signatory, subject, related
    """
    data = {}
    current = 0

    # Позиция 0 — № вх
    data['incoming_num'] = read_current_field()

    # → Позиция 4 — ссылка на файл письма
    navigate_tabs(4 - current)
    current = 4
    data['file_link'] = read_current_field()

    # → Позиция 5 — Корреспондент
    navigate_tabs(5 - current)
    current = 5
    data['correspondent'] = read_current_field()

    # → Позиция 6 — Дата
    navigate_tabs(6 - current)
    current = 6
    data['date'] = read_current_field()

    # → Позиция 7 — № письма
    navigate_tabs(7 - current)
    current = 7
    data['letter_num'] = read_current_field()

    # → Позиция 9 — Подписант
    navigate_tabs(9 - current)
    current = 9
    data['signatory'] = read_current_field()

    # → Позиция 11 — Тема письма
    navigate_tabs(11 - current)
    current = 11
    data['subject'] = read_current_field()

    # → Позиция 16 — Связанное письмо
    navigate_tabs(16 - current)
    data['related'] = read_current_field()

    return data


# ── Вспомогательные функции ───────────────────────────────────────────────────

def sanitize_for_filename(text: str) -> str:
    """Заменяет запрещённые символы в имени файла на '_'."""
    return re.sub(r'[<>:"/\\|?*\r\n\t]', '_', text)


def parse_date(date_str: str):
    """Пробует распознать дату из строки. Возвращает datetime или None."""
    for fmt in ('%d.%m.%Y', '%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y', '%Y.%m.%d'):
        try:
            return datetime.strptime(date_str.strip(), fmt)
        except ValueError:
            continue
    return None


def fmt_date_ymd(date_str: str) -> str:
    """Форматирует дату в yyyy-mm-dd."""
    dt = parse_date(date_str)
    return dt.strftime('%Y-%m-%d') if dt else date_str


def fmt_date_dmy_underscore(date_str: str) -> str:
    """Форматирует дату в dd_mm_yyyy."""
    dt = parse_date(date_str)
    return dt.strftime('%d_%m_%Y') if dt else re.sub(r'[.\-/]', '_', date_str)


def build_default_filename(date_str: str, incoming_num: str, letter_num: str) -> str:
    """
    Формирует имя файла по шаблону:
      [yyyy-mm-dd] [№вх]_[№письма_очищенный]_[dd_mm_yyyy]
    """
    date_ymd = fmt_date_ymd(date_str)
    date_dmy = fmt_date_dmy_underscore(date_str)
    letter_clean = sanitize_for_filename(letter_num)
    return f"{date_ymd} {incoming_num}_{letter_clean}_{date_dmy}"


def calc_folder_num(full_path: str) -> str:
    """
    Вычисляет "№ папки" — часть пути после базовой папки DEFAULT_SAVE_FOLDER.
    Если путь не начинается с базовой, возвращает full_path целиком.
    """
    base = DEFAULT_SAVE_FOLDER.rstrip('\\/')
    norm = full_path.replace('/', '\\')
    if norm.lower().startswith(base.lower()):
        return norm[len(base):].lstrip('\\/')
    return norm


# ── Запись в Excel ────────────────────────────────────────────────────────────

def write_to_excel(row_data: dict):
    """
    Добавляет строку в конец последнего листа журнала Excel.
    row_data — словарь с полями:
      date, incoming_num, letter_num, subject, author,
      signed_by, folder_num, keywords, related, hyperlink_path
    """
    if not HAS_OPENPYXL:
        raise RuntimeError("Библиотека openpyxl не установлена.")

    if not os.path.exists(EXCEL_PATH):
        raise FileNotFoundError(f"Файл журнала не найден:\n{EXCEL_PATH}")

    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb.worksheets[-1]  # последний лист

    # Найти первую пустую строку (по столбцу A)
    last_row = ws.max_row
    while last_row > 1 and ws.cell(row=last_row, column=1).value is None:
        last_row -= 1
    new_row = last_row + 1

    # A — Дата (yyyy-mm-dd)
    ws.cell(row=new_row, column=1).value = row_data['date']

    # B — № вх
    ws.cell(row=new_row, column=2).value = row_data['incoming_num']

    # C — № письма (с гиперссылкой на файл письма)
    cell_letter = ws.cell(row=new_row, column=3)
    cell_letter.value = row_data['letter_num']
    if row_data.get('hyperlink_path'):
        cell_letter.hyperlink = row_data['hyperlink_path']
        cell_letter.style = 'Hyperlink'

    # D — Тема письма
    ws.cell(row=new_row, column=4).value = row_data['subject']

    # E — Автор письма
    ws.cell(row=new_row, column=5).value = row_data['author']

    # F — За подписью (Подписант + перенос строки + Корреспондент)
    cell_signed = ws.cell(row=new_row, column=6)
    cell_signed.value = row_data['signed_by']
    cell_signed.alignment = Alignment(wrap_text=True)

    # G — № папки
    ws.cell(row=new_row, column=7).value = row_data['folder_num']

    # H — Кто регистрировал (имя пользователя Windows)
    ws.cell(row=new_row, column=8).value = row_data['who_registered']

    # I — Ключевые слова
    ws.cell(row=new_row, column=9).value = row_data['keywords']

    # J — Связанное письмо
    ws.cell(row=new_row, column=10).value = row_data['related']

    wb.save(EXCEL_PATH)
    # Файл не закрывается — по требованию ТЗ


# ── Диалоговое окно ───────────────────────────────────────────────────────────

class RegistrationApp(tk.Tk):
    """Главное окно программы."""

    def __init__(self, landocs_data: dict):
        super().__init__()
        self.landocs_data = landocs_data

        date_str = landocs_data.get('date', '')
        incoming = landocs_data.get('incoming_num', '')
        letter = landocs_data.get('letter_num', '')
        self._default_filename = build_default_filename(date_str, incoming, letter)

        self.title("Регистрация входящей корреспонденции")
        self.resizable(True, False)
        self._build_ui()
        self._center_window()

    # ── Построение интерфейса ──────────────────────────────────────────────

    def _build_ui(self):
        root_frame = ttk.Frame(self, padding=12)
        root_frame.grid(row=0, column=0, sticky='nsew')
        self.columnconfigure(0, weight=1)
        root_frame.columnconfigure(0, weight=1)

        # --- Данные из LanDocs (только для просмотра) ---
        info_frame = ttk.LabelFrame(root_frame, text="Данные из LanDocs", padding=8)
        info_frame.grid(row=0, column=0, sticky='ew', pady=(0, 8))
        info_frame.columnconfigure(1, weight=1)

        preview_fields = [
            ("Дата:",             'date'),
            ("№ вх:",             'incoming_num'),
            ("№ письма:",         'letter_num'),
            ("Тема письма:",      'subject'),
            ("Подписант:",        'signatory'),
            ("Корреспондент:",    'correspondent'),
            ("Связанное письмо:", 'related'),
        ]
        for i, (label, key) in enumerate(preview_fields):
            ttk.Label(info_frame, text=label, anchor='e').grid(
                row=i, column=0, sticky='e', padx=(0, 6), pady=2)
            ttk.Label(info_frame, text=self.landocs_data.get(key, ''),
                      anchor='w', wraplength=460).grid(
                row=i, column=1, sticky='w', pady=2)

        # --- Поля для заполнения пользователем ---
        input_frame = ttk.LabelFrame(root_frame, text="Данные для регистрации", padding=8)
        input_frame.grid(row=1, column=0, sticky='ew', pady=(0, 8))
        input_frame.columnconfigure(1, weight=1)

        # Автор письма
        ttk.Label(input_frame, text="Автор письма:").grid(
            row=0, column=0, sticky='e', padx=(0, 6), pady=4)
        self.author_var = tk.StringVar(value="-")
        ttk.Entry(input_frame, textvariable=self.author_var, width=48).grid(
            row=0, column=1, sticky='ew', pady=4)

        # Ключевые слова
        ttk.Label(input_frame, text="Ключевые слова:").grid(
            row=1, column=0, sticky='e', padx=(0, 6), pady=4)
        self.keywords_var = tk.StringVar(value="")
        ttk.Entry(input_frame, textvariable=self.keywords_var, width=48).grid(
            row=1, column=1, sticky='ew', pady=4)

        # Название файла
        ttk.Label(input_frame, text="Название файла:").grid(
            row=2, column=0, sticky='e', padx=(0, 6), pady=4)
        self.filename_var = tk.StringVar(value=self._default_filename)
        ttk.Entry(input_frame, textvariable=self.filename_var, width=48).grid(
            row=2, column=1, sticky='ew', pady=4)

        # Папка сохранения + кнопка
        ttk.Label(input_frame, text="Папка сохранения:").grid(
            row=3, column=0, sticky='e', padx=(0, 6), pady=4)
        folder_row = ttk.Frame(input_frame)
        folder_row.grid(row=3, column=1, sticky='ew', pady=4)
        folder_row.columnconfigure(0, weight=1)

        self.save_path_var = tk.StringVar(value="")
        ttk.Entry(folder_row, textvariable=self.save_path_var,
                  state='readonly', width=38).grid(row=0, column=0, sticky='ew')
        ttk.Button(folder_row, text="Выбрать папку…",
                   command=self._choose_save_folder).grid(
            row=0, column=1, padx=(6, 0))

        # № папки (вычисляется автоматически)
        ttk.Label(input_frame, text="№ папки:").grid(
            row=4, column=0, sticky='e', padx=(0, 6), pady=4)
        self.folder_num_var = tk.StringVar(value="")
        ttk.Label(input_frame, textvariable=self.folder_num_var,
                  anchor='w', wraplength=460, foreground='navy').grid(
            row=4, column=1, sticky='w', pady=4)

        # --- Кнопки ---
        btn_frame = ttk.Frame(root_frame)
        btn_frame.grid(row=2, column=0, pady=4)

        ttk.Button(btn_frame, text="Зарегистрировать в журнал",
                   command=self._on_register).pack(side='left', padx=6)
        ttk.Button(btn_frame, text="Отмена",
                   command=self.destroy).pack(side='left', padx=6)

    # ── Логика кнопок ─────────────────────────────────────────────────────

    def _choose_save_folder(self):
        """Открывает диалог выбора папки (начиная с базовой директории)."""
        initial = (DEFAULT_SAVE_FOLDER
                   if os.path.isdir(DEFAULT_SAVE_FOLDER)
                   else os.path.expanduser("~"))

        # Получаем расширение из ссылки на файл в LanDocs
        file_link = self.landocs_data.get('file_link', '')
        ext = os.path.splitext(file_link)[1].lower() if file_link else '.pdf'
        if not ext:
            ext = '.pdf'

        filename_base = self.filename_var.get() or self._default_filename

        selected = filedialog.asksaveasfilename(
            title="Выберите папку и имя файла для сохранения письма",
            initialdir=initial,
            initialfile=filename_base + ext,
            defaultextension=ext,
            filetypes=[
                ("PDF файлы", "*.pdf"),
                ("Все файлы", "*.*"),
            ],
        )
        if not selected:
            return

        # Нормализуем разделители
        selected = selected.replace('/', '\\')
        self.save_path_var.set(selected)

        # Обновляем "Название файла" из имени выбранного файла (без расширения)
        base_name = os.path.splitext(os.path.basename(selected))[0]
        if base_name:
            self.filename_var.set(base_name)

        # Вычисляем № папки
        folder_dir = os.path.dirname(selected)
        folder_num = calc_folder_num(folder_dir)
        self.folder_num_var.set(folder_num)

    def _on_register(self):
        """Валидация и запись в журнал."""
        if not self.save_path_var.get():
            messagebox.showwarning(
                "Внимание",
                "Выберите папку и имя файла для сохранения письма.",
                parent=self,
            )
            return

        if not HAS_OPENPYXL:
            messagebox.showerror(
                "Ошибка",
                "Библиотека openpyxl не установлена.\n"
                "Установите её: pip install openpyxl",
                parent=self,
            )
            return

        try:
            self._do_register()
            messagebox.showinfo(
                "Готово",
                "Запись успешно добавлена в журнал регистрации!",
                parent=self,
            )
            self.destroy()
        except Exception as exc:
            messagebox.showerror(
                "Ошибка",
                f"Не удалось записать в журнал:\n{exc}",
                parent=self,
            )

    def _do_register(self):
        date_str = self.landocs_data.get('date', '')
        signatory = self.landocs_data.get('signatory', '')
        correspondent = self.landocs_data.get('correspondent', '')

        # "За подписью" = Подписант + перенос строки внутри ячейки + Корреспондент
        if correspondent:
            signed_by = f"{signatory}\n{correspondent}"
        else:
            signed_by = signatory

        # Путь для гиперссылки = выбранный путь сохранения файла
        hyperlink_path = self.save_path_var.get()

        row_data = {
            'date':           fmt_date_ymd(date_str),
            'incoming_num':   self.landocs_data.get('incoming_num', ''),
            'letter_num':     self.landocs_data.get('letter_num', ''),
            'subject':        self.landocs_data.get('subject', ''),
            'author':         self.author_var.get(),
            'signed_by':      signed_by,
            'folder_num':     self.folder_num_var.get(),
            'who_registered': getpass.getuser(),
            'keywords':       self.keywords_var.get(),
            'related':        self.landocs_data.get('related', ''),
            'hyperlink_path': hyperlink_path,
        }
        write_to_excel(row_data)

    # ── Утилиты ───────────────────────────────────────────────────────────

    def _center_window(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"+{(sw - w) // 2}+{(sh - h) // 2}")


# ── Точка входа ───────────────────────────────────────────────────────────────

def main():
    # Проверяем зависимости
    missing = []
    if not HAS_WIN32:
        missing.append("pywin32")
    if not HAS_OPENPYXL:
        missing.append("openpyxl")

    if missing:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Не хватает зависимостей",
            "Не установлены библиотеки:\n  " + "\n  ".join(missing) +
            "\n\nУстановите через pip или пересоберите .exe.\n"
            "Продолжить в демо-режиме (без LanDocs)?",
        )
        if not HAS_WIN32:
            # Демо-режим: показываем окно с пустыми полями
            demo_data = {
                'incoming_num': 'вх-XXXX',
                'file_link':    '',
                'correspondent': '',
                'date':         datetime.today().strftime('%d.%m.%Y'),
                'letter_num':   '',
                'signatory':    '',
                'subject':      '',
                'related':      '',
            }
            app = RegistrationApp(demo_data)
            app.mainloop()
            return

    # Небольшая пауза, чтобы пользователь успел переключиться в LanDocs
    # (актуально при запуске через горячую клавишу из AutoHotkey)
    time.sleep(0.4)

    # Извлекаем данные из LanDocs
    try:
        landocs_data = extract_landocs_data()
    except Exception as exc:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Ошибка извлечения данных",
            f"Не удалось прочитать данные из LanDocs:\n{exc}\n\n"
            "Убедитесь, что окно регистрационной карточки активно\n"
            "и курсор стоит в первом поле формы.",
        )
        root.destroy()
        sys.exit(1)

    app = RegistrationApp(landocs_data)
    app.mainloop()


if __name__ == '__main__':
    main()
