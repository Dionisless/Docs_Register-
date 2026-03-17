#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
1_letter_register.py — Регистратор входящей корреспонденции из LanDocs
=======================================================================
Восстановлено по версии от 13.03.2026 ~16:00 VL (landocs_register.py).
Дополнено поддержкой форматов Word (.doc/.docx) и Excel (.xls/.xlsx)
в дополнение к PDF.

Запуск: через горячую клавишу (см. hotkey.ahk) или ярлык.
При запуске окно LanDocs должно быть активным, курсор — в поле 0.

Позиции полей в форме LanDocs (количество нажатий Tab от начала):
  0  — № вх
  3  — ссылка на файл письма
  4  — Корреспондент
  5  — Дата
  6  — № письма
  8  — Подписант
  10 — Тема письма
  15 — Связанное письмо
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

# ── Зависимости ──────────────────────────────────────────────────────────────

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

TAB_DELAY  = 0.07
COPY_DELAY = 0.12

# ── Поддерживаемые форматы писем ─────────────────────────────────────────────

LETTER_EXTS = ('.pdf', '.doc', '.docx', '.xls', '.xlsx')
LETTER_FILETYPES = [
    ("Письма (PDF, Word, Excel)", "*.pdf *.doc *.docx *.xls *.xlsx"),
    ("PDF файлы",                 "*.pdf"),
    ("Word файлы",                "*.doc *.docx"),
    ("Excel файлы",               "*.xls *.xlsx"),
    ("Все файлы",                 "*.*"),
]

# ── Буфер обмена и клавиатура ─────────────────────────────────────────────────

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


# ── Извлечение данных из LanDocs ──────────────────────────────────────────────

def extract_landocs_data() -> dict:
    """Считывает поля из открытой регистрационной карточки LanDocs."""
    data = {}
    current = 0

    # Позиция 0 — № вх (Tab вперёд + Shift+Tab назад для активации выделения)
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


# ── Поиск файла письма в ViewDir ─────────────────────────────────────────────

def find_latest_in_viewdir() -> str:
    """Возвращает путь к самому новому файлу в %LOCALAPPDATA%\\Temp\\ViewDir."""
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
    return re.sub(r'[<>:"/\\|?*\r\n\t]', '_', text)


def parse_date(date_str: str):
    for fmt in ('%d.%m.%Y', '%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y', '%Y.%m.%d'):
        try:
            return datetime.strptime(date_str.strip(), fmt)
        except ValueError:
            continue
    return None


def fmt_date_ymd(date_str: str) -> str:
    dt = parse_date(date_str)
    return dt.strftime('%Y-%m-%d') if dt else date_str


def fmt_date_dmy_underscore(date_str: str) -> str:
    dt = parse_date(date_str)
    return dt.strftime('%d_%m_%Y') if dt else re.sub(r'[.\-/]', '_', date_str)


def build_default_filename(date_str: str, incoming_num: str, letter_num: str) -> str:
    date_ymd = fmt_date_ymd(date_str)
    date_dmy = fmt_date_dmy_underscore(date_str)
    letter_clean = sanitize_for_filename(letter_num)
    return f"{date_ymd} {incoming_num}_{letter_clean}_{date_dmy}"


def calc_folder_num(full_path: str) -> str:
    base = DEFAULT_SAVE_FOLDER.rstrip('\\/')
    norm = full_path.replace('/', '\\')
    if norm.lower().startswith(base.lower()):
        return norm[len(base):].lstrip('\\/')
    return norm


# ── Запись в Excel ────────────────────────────────────────────────────────────

def write_to_excel(row_data: dict):
    if not HAS_OPENPYXL:
        raise RuntimeError("Библиотека openpyxl не установлена.")
    if not os.path.exists(EXCEL_PATH):
        raise FileNotFoundError(f"Файл журнала не найден:\n{EXCEL_PATH}")

    wb = openpyxl.load_workbook(EXCEL_PATH)
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

    wb.save(EXCEL_PATH)


# ── Диалоговое окно ───────────────────────────────────────────────────────────

class RegistrationApp(tk.Tk):

    def __init__(self, landocs_data: dict):
        super().__init__()
        self.landocs_data = landocs_data
        self._preview_vars = {}

        self.title("Регистрация входящей корреспонденции")
        self.resizable(True, False)
        self._build_ui()
        self._apply_landocs_data()
        self._center_window()

    def _build_ui(self):
        root_frame = ttk.Frame(self, padding=12)
        root_frame.grid(row=0, column=0, sticky='nsew')
        self.columnconfigure(0, weight=1)
        root_frame.columnconfigure(0, weight=1)

        # Данные из LanDocs (только для просмотра)
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
            var = tk.StringVar()
            self._preview_vars[key] = var
            ttk.Label(info_frame, textvariable=var,
                      anchor='w', wraplength=460).grid(
                row=i, column=1, sticky='w', pady=2)

        # Данные для заполнения пользователем
        input_frame = ttk.LabelFrame(root_frame, text="Данные для регистрации", padding=8)
        input_frame.grid(row=1, column=0, sticky='ew', pady=(0, 8))
        input_frame.columnconfigure(1, weight=1)

        ttk.Label(input_frame, text="Автор письма:").grid(
            row=0, column=0, sticky='e', padx=(0, 6), pady=4)
        self.author_var = tk.StringVar(value="-")
        ttk.Entry(input_frame, textvariable=self.author_var, width=48).grid(
            row=0, column=1, sticky='ew', pady=4)

        ttk.Label(input_frame, text="Ключевые слова:").grid(
            row=1, column=0, sticky='e', padx=(0, 6), pady=4)
        self.keywords_var = tk.StringVar(value="")
        ttk.Entry(input_frame, textvariable=self.keywords_var, width=48).grid(
            row=1, column=1, sticky='ew', pady=4)

        ttk.Label(input_frame, text="Название файла:").grid(
            row=2, column=0, sticky='e', padx=(0, 6), pady=4)
        self.filename_var = tk.StringVar(value="")
        ttk.Entry(input_frame, textvariable=self.filename_var, width=48).grid(
            row=2, column=1, sticky='ew', pady=4)

        ttk.Label(input_frame, text="Папка сохранения:").grid(
            row=3, column=0, sticky='e', padx=(0, 6), pady=4)
        folder_row = ttk.Frame(input_frame)
        folder_row.grid(row=3, column=1, sticky='ew', pady=4)
        folder_row.columnconfigure(0, weight=1)
        self.save_path_var = tk.StringVar(value="")
        ttk.Entry(folder_row, textvariable=self.save_path_var,
                  state='readonly', width=38).grid(row=0, column=0, sticky='ew')
        ttk.Button(folder_row, text="Выбрать папку…",
                   command=self._choose_save_folder).grid(row=0, column=1, padx=(6, 0))

        ttk.Label(input_frame, text="№ папки:").grid(
            row=4, column=0, sticky='e', padx=(0, 6), pady=4)
        self.folder_num_var = tk.StringVar(value="")
        ttk.Label(input_frame, textvariable=self.folder_num_var,
                  anchor='w', wraplength=460, foreground='navy').grid(
            row=4, column=1, sticky='w', pady=4)

        # Статус и кнопки
        self._reparse_status = tk.StringVar(value="")
        ttk.Label(root_frame, textvariable=self._reparse_status,
                  foreground='gray').grid(row=2, column=0, pady=(0, 2))

        btn_frame = ttk.Frame(root_frame)
        btn_frame.grid(row=3, column=0, pady=4)

        ttk.Button(btn_frame, text="Зарегистрировать в журнал",
                   command=self._on_register).pack(side='left', padx=6)
        self._reparse_btn = ttk.Button(btn_frame, text="Считать из LanDocs заново",
                                       command=self._start_reparse)
        self._reparse_btn.pack(side='left', padx=6)
        ttk.Button(btn_frame, text="Отмена",
                   command=self.destroy).pack(side='left', padx=6)

    def _apply_landocs_data(self):
        d = self.landocs_data
        for key, var in self._preview_vars.items():
            var.set(d.get(key, ''))
        date_str = d.get('date', '')
        incoming = d.get('incoming_num', '')
        letter   = d.get('letter_num', '')
        self._default_filename = build_default_filename(date_str, incoming, letter)
        self.filename_var.set(self._default_filename)
        self.save_path_var.set('')
        self.folder_num_var.set('')

    def _start_reparse(self):
        self._reparse_btn.config(state='disabled')
        self.iconify()
        self._reparse_countdown(3)

    def _reparse_countdown(self, n: int):
        if n > 0:
            self._reparse_status.set(f"Переключитесь в LanDocs… считывание через {n} сек.")
            self.after(1000, self._reparse_countdown, n - 1)
        else:
            self._reparse_status.set("Выполняется считывание…")
            self.after(50, self._do_reparse)

    def _do_reparse(self):
        try:
            self.landocs_data = extract_landocs_data()
            self._apply_landocs_data()
            self._reparse_status.set("Считывание завершено.")
        except Exception as exc:
            self._reparse_status.set(f"Ошибка считывания: {exc}")
        finally:
            self._reparse_btn.config(state='normal')
            self.deiconify()
            self.lift()

    def _choose_save_folder(self):
        initial = (DEFAULT_SAVE_FOLDER
                   if os.path.isdir(DEFAULT_SAVE_FOLDER)
                   else os.path.expanduser("~"))

        # Определяем расширение: сначала из ссылки LanDocs, потом из ViewDir
        file_link = self.landocs_data.get('file_link', '')
        ext = os.path.splitext(file_link)[1].lower() if file_link else ''
        if not ext:
            latest = find_latest_in_viewdir()
            ext = os.path.splitext(latest)[1].lower() if latest else ''
        if not ext or ext not in LETTER_EXTS:
            ext = '.pdf'

        filename_base = self.filename_var.get() or self._default_filename

        selected = filedialog.asksaveasfilename(
            title="Выберите папку и имя файла для сохранения письма",
            initialdir=initial,
            initialfile=filename_base + ext,
            defaultextension=ext,
            filetypes=LETTER_FILETYPES,
        )
        if not selected:
            return

        selected = selected.replace('/', '\\')
        self.save_path_var.set(selected)

        base_name = os.path.splitext(os.path.basename(selected))[0]
        if base_name:
            self.filename_var.set(base_name)

        folder_dir = os.path.dirname(selected)
        self.folder_num_var.set(calc_folder_num(folder_dir))

    def _on_register(self):
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
                "Библиотека openpyxl не установлена.\npip install openpyxl",
                parent=self,
            )
            return
        try:
            self._do_register()
            messagebox.showinfo("Готово", "Запись добавлена в журнал!", parent=self)
            self.destroy()
        except Exception as exc:
            messagebox.showerror("Ошибка", f"Не удалось записать в журнал:\n{exc}", parent=self)

    def _do_register(self):
        d = self.landocs_data
        signatory    = d.get('signatory', '')
        correspondent = d.get('correspondent', '')
        signed_by = f"{signatory}\n{correspondent}".strip('\n') if correspondent else signatory

        hyperlink_path = self.save_path_var.get()
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

        write_to_excel({
            'date':           fmt_date_ymd(d.get('date', '')),
            'incoming_num':   d.get('incoming_num', ''),
            'letter_num':     d.get('letter_num', ''),
            'subject':        d.get('subject', ''),
            'author':         self.author_var.get(),
            'signed_by':      signed_by,
            'folder_num':     self.folder_num_var.get(),
            'who_registered': getpass.getuser(),
            'keywords':       self.keywords_var.get(),
            'related':        d.get('related', ''),
            'hyperlink_path': hyperlink_path,
        })

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
            "\n\nУстановите через pip или пересоберите .exe.",
        )
        root.destroy()
        if not HAS_WIN32:
            # Демо-режим
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
        sys.exit(1)

    time.sleep(0.4)

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
