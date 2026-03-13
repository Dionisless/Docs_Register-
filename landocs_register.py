#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Регистратор корреспонденции LanDocs
=====================================
Извлекает данные из регистрационной карточки LanDocs (через Tab-навигацию
и буфер обмена) и записывает строку в журнал Excel.

Позиции полей ВХОДЯЩИХ (Tab от начала):
  0  — № вх
  3  — ссылка на файл письма
  4  — Корреспондент
  5  — Дата
  6  — № письма
  8  — Подписант
  10 — Тема письма
  15 — Связанное письмо

Позиции полей ИСХОДЯЩИХ (Tab от начала):
  0  — № письма
  1  — Дата
  5  — Тема письма
  6  — Исполнитель
  9  — ФИО получателей (через ;)
  10 — Компании получателей (через ;)
  16 — Связанное письмо
"""

import os
import re
import sys
import time
import shutil
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

EXCEL_PATH_IN  = r"\\Prim-fs-serv\primrdu\СРЗА\Журнал регистрации корреспонденции\Журнал регистрации входящей документации.xlsx"
EXCEL_PATH_OUT = r"\\Prim-fs-serv\primrdu\СРЗА\Журнал регистрации корреспонденции\Журнал регистрации исходящей документации.xlsx"
DEFAULT_SAVE_FOLDER = r"\\Prim-fs-serv\primrdu\СРЗА\Дела СРЗА\19 Переписка"

# Задержка между нажатиями Tab (сек) — увеличьте если LanDocs тормозит
TAB_DELAY  = 0.07
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


def _send_shift_tab():
    if not HAS_WIN32:
        return
    win32api.keybd_event(win32con.VK_SHIFT, 0, 0, 0)
    win32api.keybd_event(win32con.VK_TAB, 0, 0, 0)
    win32api.keybd_event(win32con.VK_TAB, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(win32con.VK_SHIFT, 0, win32con.KEYEVENTF_KEYUP, 0)


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

def extract_landocs_data_in() -> dict:
    """
    Извлекает поля из карточки ВХОДЯЩЕГО письма.
    Курсор должен быть в первом поле формы (позиция 0).
    """
    data = {}
    current = 0

    # Позиция 0 — № вх (Tab вперёд + Shift+Tab для активации выделения)
    _send_tab()
    time.sleep(TAB_DELAY)
    _send_shift_tab()
    time.sleep(TAB_DELAY)
    data['incoming_num'] = read_current_field()

    navigate_tabs(3 - current); current = 3
    data['file_link'] = read_current_field()

    navigate_tabs(4 - current); current = 4
    data['correspondent'] = read_current_field()

    navigate_tabs(5 - current); current = 5
    data['date'] = read_current_field()

    navigate_tabs(6 - current); current = 6
    data['letter_num'] = read_current_field()

    navigate_tabs(8 - current); current = 8
    data['signatory'] = read_current_field()

    navigate_tabs(10 - current); current = 10
    data['subject'] = read_current_field()

    navigate_tabs(15 - current)
    data['related'] = read_current_field()

    return data


def extract_landocs_data_out() -> dict:
    """
    Извлекает поля из карточки ИСХОДЯЩЕГО письма.
    Курсор должен быть в первом поле формы (позиция 0).
    """
    data = {}
    current = 0

    # Позиция 0 — № письма
    _send_tab()
    time.sleep(TAB_DELAY)
    _send_shift_tab()
    time.sleep(TAB_DELAY)
    data['letter_num'] = read_current_field()

    navigate_tabs(1 - current); current = 1
    data['date'] = read_current_field()

    navigate_tabs(5 - current); current = 5
    data['subject'] = read_current_field()

    navigate_tabs(6 - current); current = 6
    data['executor'] = read_current_field()

    navigate_tabs(9 - current); current = 9
    data['recipient_names'] = read_current_field()

    navigate_tabs(10 - current); current = 10
    data['recipient_companies'] = read_current_field()

    navigate_tabs(16 - current)
    data['related'] = read_current_field()

    return data


def build_recipient_string(names_str: str, companies_str: str) -> str:
    """
    Объединяет списки ФИО и компаний в одну строку для Excel.
    Формат: "1. ФИО1 Компания1;\n2. ФИО2 Компания2"
    """
    names     = [n.strip() for n in names_str.split(';') if n.strip()]
    companies = [c.strip() for c in companies_str.split(';') if c.strip()]
    parts = []
    for i, (n, c) in enumerate(zip(names, companies), 1):
        parts.append(f"{i}. {n} {c}")
    if not parts:
        return names_str  # fallback если компании не распознались
    return ';\n'.join(parts)


# ── Поиск файла письма в ViewDir ─────────────────────────────────────────────

def find_latest_in_viewdir() -> str:
    """
    Возвращает полный путь к самому новому файлу в папке
    %LOCALAPPDATA%\\Temp\\ViewDir (и её подпапках).
    """
    local_app = os.environ.get('LOCALAPPDATA') or os.path.join(
        os.environ.get('USERPROFILE', ''), 'AppData', 'Local')
    view_dir = os.path.join(local_app, 'Temp', 'ViewDir')
    if not os.path.isdir(view_dir):
        return ''
    latest_path, latest_mtime = '', 0.0
    for dirpath, _, filenames in os.walk(view_dir):
        for fname in filenames:
            fpath = os.path.join(dirpath, fname)
            try:
                mtime = os.path.getmtime(fpath)
                if mtime > latest_mtime:
                    latest_mtime, latest_path = mtime, fpath
            except OSError:
                pass
    return latest_path


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


def fmt_date_dmy(date_str: str) -> str:
    """Форматирует дату в dd.mm.yyyy."""
    dt = parse_date(date_str)
    return dt.strftime('%d.%m.%Y') if dt else date_str


def fmt_date_dmy_underscore(date_str: str) -> str:
    """Форматирует дату в dd_mm_yyyy."""
    dt = parse_date(date_str)
    return dt.strftime('%d_%m_%Y') if dt else re.sub(r'[.\-/]', '_', date_str)


def build_default_filename_in(date_str: str, incoming_num: str, letter_num: str) -> str:
    """Имя файла для входящего: [yyyy-mm-dd] [№вх]_[№письма]_[dd_mm_yyyy]"""
    date_ymd = fmt_date_ymd(date_str)
    date_dmy = fmt_date_dmy_underscore(date_str)
    letter_clean = sanitize_for_filename(letter_num)
    return f"{date_ymd} {incoming_num}_{letter_clean}_{date_dmy}"


def build_default_filename_out(date_str: str, letter_num: str) -> str:
    """Имя файла для исходящего: [yyyy-mm-dd] [№письма]_[dd_mm_yyyy]"""
    date_ymd = fmt_date_ymd(date_str)
    date_dmy = fmt_date_dmy_underscore(date_str)
    letter_clean = sanitize_for_filename(letter_num)
    return f"{date_ymd} {letter_clean}_{date_dmy}"


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


# ── Запись в Excel — Входящие ─────────────────────────────────────────────────

def write_to_excel_in(row_data: dict):
    """
    Добавляет строку входящего письма в журнал.
    Колонки: Дата | № вх | № письма | Тема | Автор | За подписью | № папки | Кто рег. | Ключ. слова | Связанное
    """
    if not HAS_OPENPYXL:
        raise RuntimeError("Библиотека openpyxl не установлена.")
    if not os.path.exists(EXCEL_PATH_IN):
        raise FileNotFoundError(f"Файл журнала не найден:\n{EXCEL_PATH_IN}")

    wb = openpyxl.load_workbook(EXCEL_PATH_IN)
    ws = wb.worksheets[-1]

    last_row = ws.max_row
    while last_row > 1 and ws.cell(row=last_row, column=1).value is None:
        last_row -= 1
    new_row = last_row + 1

    ws.cell(row=new_row, column=1).value = row_data['date']

    ws.cell(row=new_row, column=2).value = row_data['incoming_num']

    cell_letter = ws.cell(row=new_row, column=3)
    cell_letter.value = row_data['letter_num']
    if row_data.get('hyperlink_path'):
        cell_letter.hyperlink = row_data['hyperlink_path']
        cell_letter.style = 'Hyperlink'

    ws.cell(row=new_row, column=4).value = row_data['subject']

    ws.cell(row=new_row, column=5).value = row_data['author']

    cell_signed = ws.cell(row=new_row, column=6)
    cell_signed.value = row_data['signed_by']
    cell_signed.alignment = Alignment(wrap_text=True)

    ws.cell(row=new_row, column=7).value = row_data['folder_num']

    ws.cell(row=new_row, column=8).value = row_data['who_registered']

    ws.cell(row=new_row, column=9).value = row_data['keywords']

    ws.cell(row=new_row, column=10).value = row_data['related']

    wb.save(EXCEL_PATH_IN)


# ── Запись в Excel — Исходящие ────────────────────────────────────────────────

def write_to_excel_out(row_data: dict):
    """
    Добавляет строку исходящего письма в журнал.
    Колонки: Дата | № письма | Тема | Получатель | Исполнитель | Ключ. слова | Связанное | Контроль
    """
    if not HAS_OPENPYXL:
        raise RuntimeError("Библиотека openpyxl не установлена.")
    if not os.path.exists(EXCEL_PATH_OUT):
        raise FileNotFoundError(f"Файл журнала не найден:\n{EXCEL_PATH_OUT}")

    wb = openpyxl.load_workbook(EXCEL_PATH_OUT)
    ws = wb.worksheets[-1]

    last_row = ws.max_row
    while last_row > 1 and ws.cell(row=last_row, column=1).value is None:
        last_row -= 1
    new_row = last_row + 1

    # A — Дата (dd.mm.yyyy)
    ws.cell(row=new_row, column=1).value = row_data['date']

    # B — № письма (с гиперссылкой)
    cell_letter = ws.cell(row=new_row, column=2)
    cell_letter.value = row_data['letter_num']
    if row_data.get('hyperlink_path'):
        cell_letter.hyperlink = row_data['hyperlink_path']
        cell_letter.style = 'Hyperlink'

    # C — Тема письма
    ws.cell(row=new_row, column=3).value = row_data['subject']

    # D — Получатель (ФИО + Компания)
    cell_recip = ws.cell(row=new_row, column=4)
    cell_recip.value = row_data['recipient']
    cell_recip.alignment = Alignment(wrap_text=True)

    # E — Исполнитель
    ws.cell(row=new_row, column=5).value = row_data['executor']

    # F — Ключевые слова
    ws.cell(row=new_row, column=6).value = row_data['keywords']

    # G — Связанное письмо
    ws.cell(row=new_row, column=7).value = row_data['related']

    # H — Контроль
    ws.cell(row=new_row, column=8).value = row_data['control']

    wb.save(EXCEL_PATH_OUT)


# ── Диалоговое окно ───────────────────────────────────────────────────────────

class RegistrationApp(tk.Tk):
    """Главное окно программы с двумя вкладками: Входящие / Исходящие."""

    def __init__(self):
        super().__init__()
        self.in_data  = {}
        self.out_data = {}
        self._in_preview_vars  = {}
        self._out_preview_vars = {}
        self._default_filename_in  = ''
        self._default_filename_out = ''

        self.title("Регистрация корреспонденции")
        self.resizable(True, False)
        self._build_ui()
        self._center_window()

    # ── Построение интерфейса ──────────────────────────────────────────────

    def _build_ui(self):
        root_frame = ttk.Frame(self, padding=12)
        root_frame.grid(row=0, column=0, sticky='nsew')
        self.columnconfigure(0, weight=1)
        root_frame.columnconfigure(0, weight=1)

        # Notebook с двумя вкладками
        self._notebook = ttk.Notebook(root_frame)
        self._notebook.grid(row=0, column=0, sticky='ew', pady=(0, 8))

        in_tab  = ttk.Frame(self._notebook, padding=8)
        out_tab = ttk.Frame(self._notebook, padding=8)
        self._notebook.add(in_tab,  text="  Входящие  ")
        self._notebook.add(out_tab, text="  Исходящие  ")

        self._build_incoming_tab(in_tab)
        self._build_outgoing_tab(out_tab)

        # Статус парсинга (общий)
        self._reparse_status = tk.StringVar(value="")
        ttk.Label(root_frame, textvariable=self._reparse_status,
                  foreground='gray').grid(row=1, column=0, pady=(0, 2))

        # Кнопки (общие)
        btn_frame = ttk.Frame(root_frame)
        btn_frame.grid(row=2, column=0, pady=4)

        ttk.Button(btn_frame, text="Зарегистрировать в журнал",
                   command=self._on_register).pack(side='left', padx=6)
        self._reparse_btn = ttk.Button(btn_frame, text="Запустить парсинг",
                                       command=self._start_reparse)
        self._reparse_btn.pack(side='left', padx=6)
        ttk.Button(btn_frame, text="Отмена",
                   command=self.destroy).pack(side='left', padx=6)

    def _build_incoming_tab(self, frame):
        frame.columnconfigure(1, weight=1)

        # Данные из LanDocs
        info = ttk.LabelFrame(frame, text="Данные из LanDocs", padding=8)
        info.grid(row=0, column=0, sticky='ew', pady=(0, 8))
        info.columnconfigure(1, weight=1)

        fields = [
            ("Дата:",             'date'),
            ("№ вх:",             'incoming_num'),
            ("№ письма:",         'letter_num'),
            ("Тема письма:",      'subject'),
            ("Подписант:",        'signatory'),
            ("Корреспондент:",    'correspondent'),
            ("Связанное письмо:", 'related'),
        ]
        for i, (label, key) in enumerate(fields):
            ttk.Label(info, text=label, anchor='e').grid(
                row=i, column=0, sticky='e', padx=(0, 6), pady=2)
            var = tk.StringVar()
            self._in_preview_vars[key] = var
            ttk.Label(info, textvariable=var, anchor='w', wraplength=460).grid(
                row=i, column=1, sticky='w', pady=2)

        # Данные для регистрации
        inp = ttk.LabelFrame(frame, text="Данные для регистрации", padding=8)
        inp.grid(row=1, column=0, sticky='ew')
        inp.columnconfigure(1, weight=1)

        ttk.Label(inp, text="Автор письма:").grid(
            row=0, column=0, sticky='e', padx=(0, 6), pady=4)
        self.in_author_var = tk.StringVar(value="-")
        ttk.Entry(inp, textvariable=self.in_author_var, width=48).grid(
            row=0, column=1, sticky='ew', pady=4)

        ttk.Label(inp, text="Ключевые слова:").grid(
            row=1, column=0, sticky='e', padx=(0, 6), pady=4)
        self.in_keywords_var = tk.StringVar(value="")
        ttk.Entry(inp, textvariable=self.in_keywords_var, width=48).grid(
            row=1, column=1, sticky='ew', pady=4)

        ttk.Label(inp, text="Название файла:").grid(
            row=2, column=0, sticky='e', padx=(0, 6), pady=4)
        self.in_filename_var = tk.StringVar(value="")
        ttk.Entry(inp, textvariable=self.in_filename_var, width=48).grid(
            row=2, column=1, sticky='ew', pady=4)

        ttk.Label(inp, text="Папка сохранения:").grid(
            row=3, column=0, sticky='e', padx=(0, 6), pady=4)
        row3 = ttk.Frame(inp)
        row3.grid(row=3, column=1, sticky='ew', pady=4)
        row3.columnconfigure(0, weight=1)
        self.in_save_path_var = tk.StringVar(value="")
        ttk.Entry(row3, textvariable=self.in_save_path_var,
                  state='readonly', width=38).grid(row=0, column=0, sticky='ew')
        ttk.Button(row3, text="Выбрать папку…",
                   command=self._choose_save_folder_in).grid(
            row=0, column=1, padx=(6, 0))

        ttk.Label(inp, text="№ папки:").grid(
            row=4, column=0, sticky='e', padx=(0, 6), pady=4)
        self.in_folder_num_var = tk.StringVar(value="")
        ttk.Label(inp, textvariable=self.in_folder_num_var,
                  anchor='w', wraplength=460, foreground='navy').grid(
            row=4, column=1, sticky='w', pady=4)

    def _build_outgoing_tab(self, frame):
        frame.columnconfigure(1, weight=1)

        # Данные из LanDocs
        info = ttk.LabelFrame(frame, text="Данные из LanDocs", padding=8)
        info.grid(row=0, column=0, sticky='ew', pady=(0, 8))
        info.columnconfigure(1, weight=1)

        fields = [
            ("Дата:",             'date'),
            ("№ письма:",         'letter_num'),
            ("Тема письма:",      'subject'),
            ("Исполнитель:",      'executor'),
            ("Получатели (ФИО):", 'recipient_names'),
            ("Получатели (орг):", 'recipient_companies'),
            ("Связанное письмо:", 'related'),
        ]
        for i, (label, key) in enumerate(fields):
            ttk.Label(info, text=label, anchor='e').grid(
                row=i, column=0, sticky='e', padx=(0, 6), pady=2)
            var = tk.StringVar()
            self._out_preview_vars[key] = var
            ttk.Label(info, textvariable=var, anchor='w', wraplength=460).grid(
                row=i, column=1, sticky='w', pady=2)

        # Данные для регистрации
        inp = ttk.LabelFrame(frame, text="Данные для регистрации", padding=8)
        inp.grid(row=1, column=0, sticky='ew')
        inp.columnconfigure(1, weight=1)

        ttk.Label(inp, text="Ключевые слова:").grid(
            row=0, column=0, sticky='e', padx=(0, 6), pady=4)
        self.out_keywords_var = tk.StringVar(value="")
        ttk.Entry(inp, textvariable=self.out_keywords_var, width=48).grid(
            row=0, column=1, sticky='ew', pady=4)

        ttk.Label(inp, text="Контроль:").grid(
            row=1, column=0, sticky='e', padx=(0, 6), pady=4)
        self.out_control_var = tk.StringVar(value="")
        ttk.Entry(inp, textvariable=self.out_control_var, width=48).grid(
            row=1, column=1, sticky='ew', pady=4)

        ttk.Label(inp, text="Название файла:").grid(
            row=2, column=0, sticky='e', padx=(0, 6), pady=4)
        self.out_filename_var = tk.StringVar(value="")
        ttk.Entry(inp, textvariable=self.out_filename_var, width=48).grid(
            row=2, column=1, sticky='ew', pady=4)

        ttk.Label(inp, text="Папка сохранения:").grid(
            row=3, column=0, sticky='e', padx=(0, 6), pady=4)
        row3 = ttk.Frame(inp)
        row3.grid(row=3, column=1, sticky='ew', pady=4)
        row3.columnconfigure(0, weight=1)
        self.out_save_path_var = tk.StringVar(value="")
        ttk.Entry(row3, textvariable=self.out_save_path_var,
                  state='readonly', width=38).grid(row=0, column=0, sticky='ew')
        ttk.Button(row3, text="Выбрать папку…",
                   command=self._choose_save_folder_out).grid(
            row=0, column=1, padx=(6, 0))

        ttk.Label(inp, text="№ папки:").grid(
            row=4, column=0, sticky='e', padx=(0, 6), pady=4)
        self.out_folder_num_var = tk.StringVar(value="")
        ttk.Label(inp, textvariable=self.out_folder_num_var,
                  anchor='w', wraplength=460, foreground='navy').grid(
            row=4, column=1, sticky='w', pady=4)

    # ── Применение данных из LanDocs ───────────────────────────────────────

    def _apply_incoming_data(self):
        d = self.in_data
        for key, var in self._in_preview_vars.items():
            var.set(d.get(key, ''))
        self._default_filename_in = build_default_filename_in(
            d.get('date', ''), d.get('incoming_num', ''), d.get('letter_num', ''))
        self.in_filename_var.set(self._default_filename_in)
        self.in_save_path_var.set('')
        self.in_folder_num_var.set('')

    def _apply_outgoing_data(self):
        d = self.out_data
        for key, var in self._out_preview_vars.items():
            var.set(d.get(key, ''))
        self._default_filename_out = build_default_filename_out(
            d.get('date', ''), d.get('letter_num', ''))
        self.out_filename_var.set(self._default_filename_out)
        # Ключевые слова по умолчанию = тема письма
        self.out_keywords_var.set(d.get('subject', ''))
        self.out_save_path_var.set('')
        self.out_folder_num_var.set('')

    # ── Парсинг ────────────────────────────────────────────────────────────

    def _active_tab(self) -> str:
        """Возвращает 'in' или 'out' в зависимости от активной вкладки."""
        return 'out' if self._notebook.index(self._notebook.select()) == 1 else 'in'

    def _start_reparse(self):
        self._reparse_btn.config(state='disabled')
        self.iconify()
        self._reparse_countdown(3)

    def _reparse_countdown(self, n: int):
        if n > 0:
            self._reparse_status.set(f"Переключитесь в LanDocs… парсинг через {n} сек.")
            self.after(1000, self._reparse_countdown, n - 1)
        else:
            self._reparse_status.set("Выполняется парсинг…")
            self.after(50, self._do_reparse)

    def _do_reparse(self):
        tab = self._active_tab()
        try:
            if tab == 'in':
                self.in_data = extract_landocs_data_in()
                self._apply_incoming_data()
            else:
                self.out_data = extract_landocs_data_out()
                self._apply_outgoing_data()
            self._reparse_status.set("Парсинг завершён успешно.")
        except Exception as exc:
            self._reparse_status.set(f"Ошибка парсинга: {exc}")
        finally:
            self._reparse_btn.config(state='normal')
            self.deiconify()
            self.lift()

    # ── Выбор папки сохранения ────────────────────────────────────────────

    def _choose_save_folder_in(self):
        self._choose_save_folder(
            file_link=self.in_data.get('file_link', ''),
            filename_var=self.in_filename_var,
            default_filename=self._default_filename_in,
            save_path_var=self.in_save_path_var,
            folder_num_var=self.in_folder_num_var,
        )

    def _choose_save_folder_out(self):
        self._choose_save_folder(
            file_link=self.out_data.get('file_link', ''),
            filename_var=self.out_filename_var,
            default_filename=self._default_filename_out,
            save_path_var=self.out_save_path_var,
            folder_num_var=self.out_folder_num_var,
        )

    def _choose_save_folder(self, file_link, filename_var, default_filename,
                            save_path_var, folder_num_var):
        initial = (DEFAULT_SAVE_FOLDER
                   if os.path.isdir(DEFAULT_SAVE_FOLDER)
                   else os.path.expanduser("~"))

        ext = os.path.splitext(file_link)[1].lower() if file_link else '.pdf'
        if not ext:
            ext = '.pdf'

        filename_base = filename_var.get() or default_filename

        selected = filedialog.asksaveasfilename(
            title="Выберите папку и имя файла для сохранения письма",
            initialdir=initial,
            initialfile=filename_base + ext,
            defaultextension=ext,
            filetypes=[("PDF файлы", "*.pdf"), ("Все файлы", "*.*")],
        )
        if not selected:
            return

        selected = selected.replace('/', '\\')
        save_path_var.set(selected)

        base_name = os.path.splitext(os.path.basename(selected))[0]
        if base_name:
            filename_var.set(base_name)

        folder_dir = os.path.dirname(selected)
        folder_num_var.set(calc_folder_num(folder_dir))

    # ── Регистрация ────────────────────────────────────────────────────────

    def _on_register(self):
        if self._active_tab() == 'in':
            self._on_register_in()
        else:
            self._on_register_out()

    def _on_register_in(self):
        if not self.in_save_path_var.get():
            messagebox.showwarning("Внимание",
                                   "Выберите папку и имя файла для сохранения письма.",
                                   parent=self)
            return
        if not HAS_OPENPYXL:
            messagebox.showerror("Ошибка", "Библиотека openpyxl не установлена.", parent=self)
            return
        try:
            self._do_register_in()
            messagebox.showinfo("Готово", "Запись успешно добавлена в журнал регистрации!",
                                parent=self)
            self.destroy()
        except Exception as exc:
            messagebox.showerror("Ошибка", f"Не удалось записать в журнал:\n{exc}", parent=self)

    def _do_register_in(self):
        d = self.in_data
        date_str      = d.get('date', '')
        signatory     = d.get('signatory', '')
        correspondent = d.get('correspondent', '')
        signed_by = f"{signatory}\n{correspondent}" if correspondent else signatory

        hyperlink_path = self.in_save_path_var.get()
        if hyperlink_path:
            src = find_latest_in_viewdir()
            if src:
                shutil.copy2(src, hyperlink_path)
            else:
                raise FileNotFoundError(
                    "Не найден файл письма в папке ViewDir.\n"
                    r"Убедитесь, что письмо открыто в LanDocs: "
                    r"%LOCALAPPDATA%\Temp\ViewDir"
                )

        write_to_excel_in({
            'date':           fmt_date_ymd(date_str),
            'incoming_num':   d.get('incoming_num', ''),
            'letter_num':     d.get('letter_num', ''),
            'subject':        d.get('subject', ''),
            'author':         self.in_author_var.get(),
            'signed_by':      signed_by,
            'folder_num':     self.in_folder_num_var.get(),
            'who_registered': getpass.getuser(),
            'keywords':       self.in_keywords_var.get(),
            'related':        d.get('related', ''),
            'hyperlink_path': hyperlink_path,
        })

    def _on_register_out(self):
        if not self.out_save_path_var.get():
            messagebox.showwarning("Внимание",
                                   "Выберите папку и имя файла для сохранения письма.",
                                   parent=self)
            return
        if not HAS_OPENPYXL:
            messagebox.showerror("Ошибка", "Библиотека openpyxl не установлена.", parent=self)
            return
        try:
            self._do_register_out()
            messagebox.showinfo("Готово", "Запись успешно добавлена в журнал регистрации!",
                                parent=self)
            self.destroy()
        except Exception as exc:
            messagebox.showerror("Ошибка", f"Не удалось записать в журнал:\n{exc}", parent=self)

    def _do_register_out(self):
        d = self.out_data
        date_str = d.get('date', '')
        recipient = build_recipient_string(
            d.get('recipient_names', ''),
            d.get('recipient_companies', ''),
        )

        hyperlink_path = self.out_save_path_var.get()
        if hyperlink_path:
            src = find_latest_in_viewdir()
            if src:
                shutil.copy2(src, hyperlink_path)
            else:
                raise FileNotFoundError(
                    "Не найден файл письма в папке ViewDir.\n"
                    r"Убедитесь, что письмо открыто в LanDocs: "
                    r"%LOCALAPPDATA%\Temp\ViewDir"
                )

        write_to_excel_out({
            'date':           fmt_date_dmy(date_str),
            'letter_num':     d.get('letter_num', ''),
            'subject':        d.get('subject', ''),
            'recipient':      recipient,
            'executor':       d.get('executor', ''),
            'keywords':       self.out_keywords_var.get(),
            'related':        d.get('related', ''),
            'control':        self.out_control_var.get(),
            'hyperlink_path': hyperlink_path,
        })

    # ── Утилиты ───────────────────────────────────────────────────────────

    def _center_window(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"+{(sw - w) // 2}+{(sh - h) // 2}")


# ── Точка входа ───────────────────────────────────────────────────────────────

def main():
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
            "\n\nПересоберите .exe с нужными зависимостями.",
        )
        root.destroy()
        sys.exit(1)

    app = RegistrationApp()
    app.mainloop()


if __name__ == '__main__':
    main()
