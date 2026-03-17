#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
1_letter_register.py — Регистратор входящих/исходящих писем
============================================================
Регистрирует письма в журналы Excel.

ИСПРАВЛЕНИЯ по сравнению со старой версией:
  - Импорт писем в форматах PDF, Word (.doc/.docx) и Excel (.xls/.xlsx)
  - Все поля из LanDocs стали редактируемыми (Entry вместо Label)

ИНТЕРФЕЙС ДЛЯ СШИВАНИЯ:
  - get_letter_data() → dict  — вернуть данные текущего письма
  - LetterRegisterApp.in_data  / .out_data  — данные входящего/исходящего
  - Данные автоматически сохраняются в session_data.json
    (ключ "letter"  — для использования другими программами)
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

from shared_lib import (
    EXCEL_PATH_IN, EXCEL_PATH_OUT, DEFAULT_SAVE_FOLDER,
    sanitize_for_filename, parse_date,
    fmt_date_ymd, fmt_date_dmy,
    load_session, save_session,
)

# ── Зависимости ──────────────────────────────────────────────────────────────

try:
    import win32api, win32con, win32clipboard
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

try:
    import openpyxl
    from openpyxl.styles import Alignment
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
    _BASE_CLASS = TkinterDnD.Tk
except Exception:
    HAS_DND = False
    _BASE_CLASS = tk.Tk

TAB_DELAY  = 0.07
COPY_DELAY = 0.12

# ── Поддерживаемые расширения писем ──────────────────────────────────────────
LETTER_EXTS = ('.pdf', '.doc', '.docx', '.xls', '.xlsx', '.msg', '.eml')
LETTER_FILETYPES = [
    ("Письма (PDF, Word, Excel)", "*.pdf *.doc *.docx *.xls *.xlsx *.msg *.eml"),
    ("PDF файлы",                 "*.pdf"),
    ("Word файлы",                "*.doc *.docx"),
    ("Excel файлы",               "*.xls *.xlsx"),
    ("Все файлы",                 "*.*"),
]

# ── Клавиатура / буфер обмена ────────────────────────────────────────────────

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
    if not HAS_WIN32: return
    win32api.keybd_event(win32con.VK_TAB, 0, 0, 0)
    win32api.keybd_event(win32con.VK_TAB, 0, win32con.KEYEVENTF_KEYUP, 0)


def _send_shift_tab():
    if not HAS_WIN32: return
    win32api.keybd_event(win32con.VK_SHIFT, 0, 0, 0)
    win32api.keybd_event(win32con.VK_TAB, 0, 0, 0)
    win32api.keybd_event(win32con.VK_TAB, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(win32con.VK_SHIFT, 0, win32con.KEYEVENTF_KEYUP, 0)


def _send_ctrl_a():
    if not HAS_WIN32: return
    win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
    win32api.keybd_event(ord('A'), 0, 0, 0)
    win32api.keybd_event(ord('A'), 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)


def _send_ctrl_c():
    if not HAS_WIN32: return
    win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
    win32api.keybd_event(ord('C'), 0, 0, 0)
    win32api.keybd_event(ord('C'), 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)


def read_current_field() -> str:
    _clear_clipboard()
    time.sleep(0.04)
    _send_ctrl_a()
    time.sleep(0.04)
    _send_ctrl_c()
    time.sleep(COPY_DELAY)
    return _get_clipboard()


def navigate_tabs(n: int):
    for _ in range(n):
        _send_tab()
        time.sleep(TAB_DELAY)


# ── Парсинг LanDocs ──────────────────────────────────────────────────────────

def extract_landocs_data_in() -> dict:
    """Считывает данные входящего письма из открытой карточки LanDocs."""
    data = {}
    current = 0
    _send_tab(); time.sleep(TAB_DELAY)
    _send_shift_tab(); time.sleep(TAB_DELAY)
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
    """Считывает данные исходящего письма из открытой карточки LanDocs."""
    data = {}
    current = 0
    _send_tab(); time.sleep(TAB_DELAY)
    _send_shift_tab(); time.sleep(TAB_DELAY)
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


# ── Утилиты ──────────────────────────────────────────────────────────────────

def build_recipient_string(names_str: str, companies_str: str) -> str:
    names     = [n.strip() for n in names_str.split(';') if n.strip()]
    companies = [c.strip() for c in companies_str.split(';') if c.strip()]
    parts = []
    for i, (n, c) in enumerate(zip(names, companies), 1):
        parts.append(f"{i}. {n} {c}")
    return ';\n'.join(parts) if parts else names_str


def find_latest_in_viewdir() -> str:
    """Ищет последний открытый файл письма в папке ViewDir LanDocs."""
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


def build_default_filename_in(date_str: str, incoming_num: str, letter_num: str) -> str:
    d = fmt_date_dmy(date_str).replace('.', '_') if date_str else ''
    parts = [p for p in [incoming_num, d, sanitize_for_filename(letter_num)] if p]
    return '_'.join(parts) or 'входящее'


def build_default_filename_out(date_str: str, letter_num: str) -> str:
    d = fmt_date_dmy(date_str).replace('.', '_') if date_str else ''
    parts = [p for p in [d, sanitize_for_filename(letter_num)] if p]
    return '_'.join(parts) or 'исходящее'


def calc_folder_num(full_path: str) -> str:
    """Вычисляет номер папки из пути (последний числовой компонент)."""
    for part in reversed(full_path.replace('\\', '/').split('/')):
        m = re.match(r'^(\d+)', part)
        if m:
            return m.group(1)
    return ''


def lookup_incoming_journal(incoming_num: str, year: str) -> dict | None:
    """Ищет запись в журнале входящих по № вх и году."""
    if not HAS_OPENPYXL:
        return None
    if not os.path.exists(EXCEL_PATH_IN):
        return None
    try:
        wb = openpyxl.load_workbook(EXCEL_PATH_IN, read_only=True, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None:
                continue
            num_cell = str(row[0]).strip()
            if num_cell != str(incoming_num).strip():
                continue
            date_cell = str(row[1]).strip() if len(row) > 1 else ''
            if year and year not in date_cell and not date_cell.endswith(year):
                continue
            return {
                'incoming_num': num_cell,
                'date':         date_cell,
                'letter_num':   str(row[2]).strip() if len(row) > 2 else '',
                'subject':      str(row[3]).strip() if len(row) > 3 else '',
                'author':       str(row[4]).strip() if len(row) > 4 else '',
                'signed_by':    str(row[5]).strip() if len(row) > 5 else '',
                'keywords':     str(row[6]).strip() if len(row) > 6 else '',
                'related':      str(row[7]).strip() if len(row) > 7 else '',
                'hyperlink':    '',
            }
        wb.close()
    except Exception:
        pass
    return None


def write_to_excel_in(row_data: dict):
    """Добавляет строку в журнал входящей корреспонденции."""
    if not HAS_OPENPYXL:
        raise RuntimeError("openpyxl не установлен")
    wb = openpyxl.load_workbook(EXCEL_PATH_IN)
    ws = wb.active
    next_row = ws.max_row + 1
    ws.cell(next_row, 1,  row_data.get('incoming_num', ''))
    ws.cell(next_row, 2,  row_data.get('date', ''))
    ws.cell(next_row, 3,  row_data.get('letter_num', ''))
    ws.cell(next_row, 4,  row_data.get('subject', ''))
    ws.cell(next_row, 5,  row_data.get('author', ''))
    ws.cell(next_row, 6,  row_data.get('signed_by', ''))
    ws.cell(next_row, 7,  row_data.get('folder_num', ''))
    ws.cell(next_row, 8,  row_data.get('who_registered', ''))
    ws.cell(next_row, 9,  row_data.get('keywords', ''))
    ws.cell(next_row, 10, row_data.get('related', ''))
    hp = row_data.get('hyperlink_path', '')
    if hp:
        cell = ws.cell(next_row, 11)
        cell.value = os.path.basename(hp)
        cell.hyperlink = hp
    wb.save(EXCEL_PATH_IN)
    wb.close()


def write_to_excel_out(row_data: dict):
    """Добавляет строку в журнал исходящей корреспонденции."""
    if not HAS_OPENPYXL:
        raise RuntimeError("openpyxl не установлен")
    wb = openpyxl.load_workbook(EXCEL_PATH_OUT)
    ws = wb.active
    next_row = ws.max_row + 1
    ws.cell(next_row, 1,  row_data.get('letter_num', ''))
    ws.cell(next_row, 2,  row_data.get('date', ''))
    ws.cell(next_row, 3,  row_data.get('subject', ''))
    ws.cell(next_row, 4,  row_data.get('recipient', ''))
    ws.cell(next_row, 5,  row_data.get('executor', ''))
    ws.cell(next_row, 6,  row_data.get('keywords', ''))
    ws.cell(next_row, 7,  row_data.get('related', ''))
    ws.cell(next_row, 8,  row_data.get('control', ''))
    hp = row_data.get('hyperlink_path', '')
    if hp:
        cell = ws.cell(next_row, 9)
        cell.value = os.path.basename(hp)
        cell.hyperlink = hp
    wb.save(EXCEL_PATH_OUT)
    wb.close()


# ── Главное окно ─────────────────────────────────────────────────────────────

class LetterRegisterApp(_BASE_CLASS):
    """
    Программа 1: Регистрация входящих/исходящих писем.

    Публичный интерфейс (для сшивания):
      .in_data   dict  — данные входящего письма
      .out_data  dict  — данные исходящего письма
      .get_letter_data() → dict  — возвращает текущее письмо (входящее)
    """

    def __init__(self):
        super().__init__()
        self.in_data  = {}
        self.out_data = {}
        self._default_filename_in  = ''
        self._default_filename_out = ''
        self._in_field_vars  = {}   # key → StringVar (Entry — редактируемые)
        self._out_field_vars = {}

        # Загрузить прошлую сессию
        session = load_session()
        if 'letter' in session:
            self.in_data = session['letter'].get('in_data', {})
            self.out_data = session['letter'].get('out_data', {})

        self.title("Регистрация корреспонденции v2")
        self.resizable(True, True)
        self._build_ui()
        self._apply_incoming_data()
        self._apply_outgoing_data()
        self._center_window()

    # ── Публичный интерфейс ────────────────────────────────────────────────

    def get_letter_data(self) -> dict:
        """Возвращает данные текущего входящего письма (для Program 2/3/4)."""
        return dict(self.in_data)

    # ── Построение UI ─────────────────────────────────────────────────────

    def _build_ui(self):
        root_frame = ttk.Frame(self, padding=8)
        root_frame.grid(row=0, column=0, sticky='nsew')
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        root_frame.columnconfigure(0, weight=1)
        root_frame.rowconfigure(0, weight=1)

        nb = ttk.Notebook(root_frame)
        nb.grid(row=0, column=0, sticky='nsew')
        self._nb = nb

        in_tab  = ttk.Frame(nb, padding=8)
        out_tab = ttk.Frame(nb, padding=8)
        nb.add(in_tab,  text="  Входящие  ")
        nb.add(out_tab, text="  Исходящие  ")
        self._build_incoming_tab(in_tab)
        self._build_outgoing_tab(out_tab)

        # Кнопки действий
        btn_frame = ttk.Frame(root_frame)
        btn_frame.grid(row=1, column=0, pady=6)
        ttk.Button(btn_frame, text="Зарегистрировать в журнал",
                   command=self._on_register).pack(side='left', padx=6)
        self._reparse_btn = ttk.Button(btn_frame, text="Считать из LanDocs (F5)",
                                       command=self._start_reparse)
        self._reparse_btn.pack(side='left', padx=6)
        ttk.Button(btn_frame, text="Сохранить сессию",
                   command=self._save_current_session).pack(side='left', padx=6)
        ttk.Button(btn_frame, text="Закрыть",
                   command=self.destroy).pack(side='left', padx=6)

        self._status_var = tk.StringVar(value="")
        ttk.Label(root_frame, textvariable=self._status_var,
                  foreground='gray').grid(row=2, column=0, pady=(0, 4))

        self.bind('<F5>', lambda _: self._start_reparse())

    def _build_field_row(self, parent, row, label_text, key, store, width=52):
        """Строит строку: Label + Entry (редактируемое) для поля из LanDocs."""
        ttk.Label(parent, text=label_text, anchor='e').grid(
            row=row, column=0, sticky='e', padx=(0, 6), pady=3)
        var = tk.StringVar()
        store[key] = var
        ent = ttk.Entry(parent, textvariable=var, width=width)
        ent.grid(row=row, column=1, sticky='ew', pady=3)

    def _build_incoming_tab(self, frame):
        frame.columnconfigure(1, weight=1)

        # Блок импорта из журнала
        imp = ttk.LabelFrame(frame, text="Импорт из журнала входящих", padding=8)
        imp.grid(row=0, column=0, columnspan=2, sticky='ew', pady=(0, 6))
        imp.columnconfigure(3, weight=1)
        ttk.Label(imp, text="№ вх:").grid(row=0, column=0, sticky='e', padx=(0,4))
        self._in_import_num = tk.StringVar()
        ttk.Entry(imp, textvariable=self._in_import_num, width=18).grid(row=0, column=1)
        ttk.Label(imp, text="  Год:").grid(row=0, column=2, sticky='e', padx=(8,4))
        self._in_import_year = tk.StringVar(value=str(datetime.now().year))
        ttk.Entry(imp, textvariable=self._in_import_year, width=6).grid(row=0, column=3, sticky='w')
        ttk.Button(imp, text="Найти в журнале",
                   command=self._on_import_journal).grid(row=0, column=4, padx=(8,0))
        self._import_status = tk.StringVar(value="")
        ttk.Label(imp, textvariable=self._import_status,
                  foreground='navy').grid(row=1, column=0, columnspan=5, pady=(4,0), sticky='w')

        # Данные из LanDocs — ВСЕ ПОЛЯ РЕДАКТИРУЕМЫЕ
        info = ttk.LabelFrame(frame, text="Данные письма (все поля редактируемые)", padding=8)
        info.grid(row=1, column=0, columnspan=2, sticky='ew', pady=(0, 8))
        info.columnconfigure(1, weight=1)
        fields = [
            ("Дата:",             'date'),
            ("№ вх:",             'incoming_num'),
            ("№ письма:",         'letter_num'),
            ("Тема письма:",      'subject'),
            ("Подписант:",        'signatory'),
            ("Корреспондент:",    'correspondent'),
            ("Ссылка на файл:",   'file_link'),
            ("Связанное письмо:", 'related'),
        ]
        for i, (lbl, key) in enumerate(fields):
            self._build_field_row(info, i, lbl, key, self._in_field_vars)

        # Данные для регистрации
        inp = ttk.LabelFrame(frame, text="Данные для регистрации", padding=8)
        inp.grid(row=2, column=0, columnspan=2, sticky='ew')
        inp.columnconfigure(1, weight=1)
        ttk.Label(inp, text="Автор письма:").grid(row=0, column=0, sticky='e', padx=(0,6), pady=4)
        self.in_author_var = tk.StringVar(value="-")
        ttk.Entry(inp, textvariable=self.in_author_var, width=48).grid(row=0, column=1, sticky='ew', pady=4)
        ttk.Label(inp, text="Ключевые слова:").grid(row=1, column=0, sticky='e', padx=(0,6), pady=4)
        self.in_keywords_var = tk.StringVar()
        ttk.Entry(inp, textvariable=self.in_keywords_var, width=48).grid(row=1, column=1, sticky='ew', pady=4)
        ttk.Label(inp, text="Название файла:").grid(row=2, column=0, sticky='e', padx=(0,6), pady=4)
        self.in_filename_var = tk.StringVar()
        ttk.Entry(inp, textvariable=self.in_filename_var, width=48).grid(row=2, column=1, sticky='ew', pady=4)
        ttk.Label(inp, text="Папка сохранения:").grid(row=3, column=0, sticky='e', padx=(0,6), pady=4)
        row3 = ttk.Frame(inp); row3.grid(row=3, column=1, sticky='ew', pady=4)
        row3.columnconfigure(0, weight=1)
        self.in_save_path_var = tk.StringVar()
        ttk.Entry(row3, textvariable=self.in_save_path_var, width=40).grid(row=0, column=0, sticky='ew')
        ttk.Button(row3, text="Выбрать…",
                   command=self._choose_save_folder_in).grid(row=0, column=1, padx=(6,0))
        ttk.Label(inp, text="№ папки:").grid(row=4, column=0, sticky='e', padx=(0,6), pady=4)
        self.in_folder_num_var = tk.StringVar()
        ttk.Label(inp, textvariable=self.in_folder_num_var,
                  anchor='w', foreground='navy').grid(row=4, column=1, sticky='w', pady=4)

    def _build_outgoing_tab(self, frame):
        frame.columnconfigure(1, weight=1)
        # Данные из LanDocs — ВСЕ ПОЛЯ РЕДАКТИРУЕМЫЕ
        info = ttk.LabelFrame(frame, text="Данные письма (все поля редактируемые)", padding=8)
        info.grid(row=0, column=0, columnspan=2, sticky='ew', pady=(0, 8))
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
        for i, (lbl, key) in enumerate(fields):
            self._build_field_row(info, i, lbl, key, self._out_field_vars)

        inp = ttk.LabelFrame(frame, text="Данные для регистрации", padding=8)
        inp.grid(row=1, column=0, columnspan=2, sticky='ew')
        inp.columnconfigure(1, weight=1)
        ttk.Label(inp, text="Ключевые слова:").grid(row=0, column=0, sticky='e', padx=(0,6), pady=4)
        self.out_keywords_var = tk.StringVar()
        ttk.Entry(inp, textvariable=self.out_keywords_var, width=48).grid(row=0, column=1, sticky='ew', pady=4)
        ttk.Label(inp, text="Контроль:").grid(row=1, column=0, sticky='e', padx=(0,6), pady=4)
        self.out_control_var = tk.StringVar()
        ttk.Entry(inp, textvariable=self.out_control_var, width=48).grid(row=1, column=1, sticky='ew', pady=4)
        ttk.Label(inp, text="Название файла:").grid(row=2, column=0, sticky='e', padx=(0,6), pady=4)
        self.out_filename_var = tk.StringVar()
        ttk.Entry(inp, textvariable=self.out_filename_var, width=48).grid(row=2, column=1, sticky='ew', pady=4)
        ttk.Label(inp, text="Папка сохранения:").grid(row=3, column=0, sticky='e', padx=(0,6), pady=4)
        row3 = ttk.Frame(inp); row3.grid(row=3, column=1, sticky='ew', pady=4)
        row3.columnconfigure(0, weight=1)
        self.out_save_path_var = tk.StringVar()
        ttk.Entry(row3, textvariable=self.out_save_path_var, width=40).grid(row=0, column=0, sticky='ew')
        ttk.Button(row3, text="Выбрать…",
                   command=self._choose_save_folder_out).grid(row=0, column=1, padx=(6,0))
        ttk.Label(inp, text="№ папки:").grid(row=4, column=0, sticky='e', padx=(0,6), pady=4)
        self.out_folder_num_var = tk.StringVar()
        ttk.Label(inp, textvariable=self.out_folder_num_var,
                  anchor='w', foreground='navy').grid(row=4, column=1, sticky='w', pady=4)

    # ── Применение данных ─────────────────────────────────────────────────

    def _apply_incoming_data(self):
        d = self.in_data
        for key, var in self._in_field_vars.items():
            var.set(d.get(key, ''))
        self._default_filename_in = build_default_filename_in(
            d.get('date',''), d.get('incoming_num',''), d.get('letter_num',''))
        self.in_filename_var.set(self._default_filename_in)

    def _apply_outgoing_data(self):
        d = self.out_data
        for key, var in self._out_field_vars.items():
            var.set(d.get(key, ''))
        self._default_filename_out = build_default_filename_out(
            d.get('date',''), d.get('letter_num',''))
        self.out_filename_var.set(self._default_filename_out)
        self.out_keywords_var.set(d.get('subject', ''))

    def _collect_in_data_from_ui(self):
        """Читает данные из Entry-полей обратно в self.in_data."""
        for key, var in self._in_field_vars.items():
            self.in_data[key] = var.get()

    def _collect_out_data_from_ui(self):
        for key, var in self._out_field_vars.items():
            self.out_data[key] = var.get()

    # ── Импорт из журнала ─────────────────────────────────────────────────

    def _on_import_journal(self):
        num  = self._in_import_num.get().strip()
        year = self._in_import_year.get().strip()
        if not num:
            self._import_status.set("Введите № вх")
            return
        try:
            result = lookup_incoming_journal(num, year)
        except Exception as exc:
            self._import_status.set(f"Ошибка: {exc}")
            return
        if result is None:
            self._import_status.set(f"Не найдено: {num} за {year} год")
            return
        self.in_data = result
        self._apply_incoming_data()
        self.in_author_var.set(result.get('author', '-'))
        self.in_keywords_var.set(result.get('keywords', ''))
        self._import_status.set(f"Импортировано: вх-{num}")

    # ── Парсинг LanDocs ───────────────────────────────────────────────────

    def _active_tab(self) -> str:
        return 'out' if self._nb.index(self._nb.select()) == 1 else 'in'

    def _start_reparse(self):
        self._reparse_btn.config(state='disabled')
        self.iconify()
        self._reparse_countdown(3)

    def _reparse_countdown(self, n: int):
        if n > 0:
            self._status_var.set(f"Переключитесь в LanDocs… считывание через {n} сек.")
            self.after(1000, self._reparse_countdown, n - 1)
        else:
            self._status_var.set("Считывание…")
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
            self._status_var.set("Считывание завершено.")
        except Exception as exc:
            self._status_var.set(f"Ошибка: {exc}")
        finally:
            self._reparse_btn.config(state='normal')
            self.deiconify()
            self.lift()

    # ── Выбор файла сохранения ────────────────────────────────────────────

    def _choose_save_folder_in(self):
        self._collect_in_data_from_ui()
        file_link = self.in_data.get('file_link', '')
        self._choose_save_folder(
            file_link=file_link,
            filename_var=self.in_filename_var,
            default_filename=self._default_filename_in,
            save_path_var=self.in_save_path_var,
            folder_num_var=self.in_folder_num_var,
        )

    def _choose_save_folder_out(self):
        self._collect_out_data_from_ui()
        file_link = self.out_data.get('file_link', '')
        self._choose_save_folder(
            file_link=file_link,
            filename_var=self.out_filename_var,
            default_filename=self._default_filename_out,
            save_path_var=self.out_save_path_var,
            folder_num_var=self.out_folder_num_var,
        )

    def _choose_save_folder(self, file_link, filename_var, default_filename,
                            save_path_var, folder_num_var):
        initial = DEFAULT_SAVE_FOLDER if os.path.isdir(DEFAULT_SAVE_FOLDER) else os.path.expanduser("~")
        # Определить расширение: сначала из ссылки, потом — из ViewDir
        ext = os.path.splitext(file_link)[1].lower() if file_link else ''
        if not ext:
            latest = find_latest_in_viewdir()
            ext = os.path.splitext(latest)[1].lower() if latest else '.pdf'
        if not ext or ext not in LETTER_EXTS:
            ext = '.pdf'
        filename_base = filename_var.get() or default_filename
        selected = filedialog.asksaveasfilename(
            title="Выберите папку и имя файла для сохранения письма",
            initialdir=initial,
            initialfile=filename_base + ext,
            defaultextension=ext,
            filetypes=LETTER_FILETYPES,
            parent=self,
        )
        if not selected:
            return
        selected = selected.replace('/', '\\')
        save_path_var.set(selected)
        base_name = os.path.splitext(os.path.basename(selected))[0]
        if base_name:
            filename_var.set(base_name)
        folder_num_var.set(calc_folder_num(os.path.dirname(selected)))

    # ── Регистрация в журнал ──────────────────────────────────────────────

    def _on_register(self):
        if self._active_tab() == 'in':
            self._on_register_in()
        else:
            self._on_register_out()

    def _on_register_in(self):
        self._collect_in_data_from_ui()
        save_path = self.in_save_path_var.get().strip()
        if not save_path:
            messagebox.showwarning("Внимание", "Выберите папку и имя файла.", parent=self)
            return
        d = self.in_data
        # Подтверждение перед записью
        confirm_msg = (
            f"Зарегистрировать входящее письмо?\n\n"
            f"  № вх:      {d.get('incoming_num','')}\n"
            f"  Дата:      {d.get('date','')}\n"
            f"  № письма:  {d.get('letter_num','')}\n"
            f"  Тема:      {d.get('subject','')[:60]}\n\n"
            f"  Сохранить файл как:\n  {save_path}\n\n"
            f"  Журнал: {EXCEL_PATH_IN}"
        )
        if not messagebox.askyesno("Подтверждение", confirm_msg, parent=self):
            return
        try:
            # Скопировать файл письма
            src = find_latest_in_viewdir()
            if src:
                shutil.copy2(src, save_path)
            else:
                if not messagebox.askyesno(
                    "Файл не найден",
                    "Файл письма не найден в ViewDir LanDocs.\n"
                    "Записать в журнал без копирования файла?",
                    parent=self,
                ):
                    return
            signatory     = d.get('signatory','')
            correspondent = d.get('correspondent','')
            signed_by = f"{signatory}\n{correspondent}".strip('\n') if correspondent else signatory
            write_to_excel_in({
                'date':           fmt_date_ymd(d.get('date','')),
                'incoming_num':   d.get('incoming_num',''),
                'letter_num':     d.get('letter_num',''),
                'subject':        d.get('subject',''),
                'author':         self.in_author_var.get(),
                'signed_by':      signed_by,
                'folder_num':     self.in_folder_num_var.get(),
                'who_registered': getpass.getuser(),
                'keywords':       self.in_keywords_var.get(),
                'related':        d.get('related',''),
                'hyperlink_path': save_path,
            })
            self._save_current_session()
            messagebox.showinfo("Готово", "Запись добавлена в журнал входящих!", parent=self)
        except Exception as exc:
            messagebox.showerror("Ошибка", str(exc), parent=self)

    def _on_register_out(self):
        self._collect_out_data_from_ui()
        save_path = self.out_save_path_var.get().strip()
        if not save_path:
            messagebox.showwarning("Внимание", "Выберите папку и имя файла.", parent=self)
            return
        d = self.out_data
        confirm_msg = (
            f"Зарегистрировать исходящее письмо?\n\n"
            f"  № письма:  {d.get('letter_num','')}\n"
            f"  Дата:      {d.get('date','')}\n"
            f"  Тема:      {d.get('subject','')[:60]}\n\n"
            f"  Сохранить файл как:\n  {save_path}\n\n"
            f"  Журнал: {EXCEL_PATH_OUT}"
        )
        if not messagebox.askyesno("Подтверждение", confirm_msg, parent=self):
            return
        try:
            src = find_latest_in_viewdir()
            if src:
                shutil.copy2(src, save_path)
            else:
                if not messagebox.askyesno(
                    "Файл не найден",
                    "Файл письма не найден в ViewDir.\n"
                    "Записать без копирования?",
                    parent=self,
                ):
                    return
            write_to_excel_out({
                'date':           fmt_date_dmy(d.get('date','')),
                'letter_num':     d.get('letter_num',''),
                'subject':        d.get('subject',''),
                'recipient':      build_recipient_string(d.get('recipient_names',''),
                                                         d.get('recipient_companies','')),
                'executor':       d.get('executor',''),
                'keywords':       self.out_keywords_var.get(),
                'related':        d.get('related',''),
                'control':        self.out_control_var.get(),
                'hyperlink_path': save_path,
            })
            messagebox.showinfo("Готово", "Запись добавлена в журнал исходящих!", parent=self)
        except Exception as exc:
            messagebox.showerror("Ошибка", str(exc), parent=self)

    # ── Сессия ────────────────────────────────────────────────────────────

    def _save_current_session(self):
        self._collect_in_data_from_ui()
        self._collect_out_data_from_ui()
        session = load_session()
        session['letter'] = {
            'in_data':  self.in_data,
            'out_data': self.out_data,
        }
        save_session(session)
        self._status_var.set("Сессия сохранена.")

    # ── Центровка ─────────────────────────────────────────────────────────

    def _center_window(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{max(w,700)}x{max(h,600)}+{(sw-max(w,700))//2}+{(sh-max(h,600))//2}")


# ── Точка входа ───────────────────────────────────────────────────────────────

def main():
    if not HAS_OPENPYXL:
        root = tk.Tk(); root.withdraw()
        messagebox.showerror("Ошибка", "Не установлен openpyxl.\npip install openpyxl")
        root.destroy(); sys.exit(1)
    app = LetterRegisterApp()
    app.mainloop()


if __name__ == '__main__':
    main()
