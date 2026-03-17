#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
4_deb_selenium.py — Регистрация карт уставок в ДЭБ (Selenium)
==============================================================
Выполняет шаг 7: загрузка обновлённых карт уставок (.vsdx + .pdf) в ДЭБ
через браузер Chrome + Selenium + pyautogui.

ВХОДНЫЕ ДАННЫЕ:
  - session_data.json → ключ "ustavki": список UstavkiEntry
    (short_name, dispatch_name, current_path — из программ 2 и 3)
  - Или ввести вручную: диспетчерское наименование + пути к файлам

ИНТЕРФЕЙС ДЛЯ СШИВАНИЯ:
  - DebApp.ustavki_entries   list[dict]  — обрабатываемые записи
  - session_data.json → ключ "ustavki" обновляется после загрузки
"""

import os
import re
import sys
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

from shared_lib import (
    DEB_BASE_URL, DEB_MAPS_URL, MAPS_FOLDER, MAPS_PDF_FOLDER,
    EMPTY_USTAVKI_ENTRY, match_object_to_short_name,
    load_session, save_session,
)

# ── Зависимости ──────────────────────────────────────────────────────────────

try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.chrome.options import Options as ChromeOptions
    HAS_SELENIUM = True
except ImportError:
    HAS_SELENIUM = False

try:
    import pyautogui
    HAS_PYAUTOGUI = True
except ImportError:
    HAS_PYAUTOGUI = False

try:
    from tkinterdnd2 import TkinterDnD
    _BASE_CLASS = TkinterDnD.Tk
except Exception:
    _BASE_CLASS = tk.Tk

# ── Загрузка в ДЭБ ───────────────────────────────────────────────────────────

def upload_to_deb_entry(dispatch_name: str, visio_path: str,
                        pdf_path: str, log_fn=None) -> tuple:
    """
    Загружает обновлённые файлы карты уставок в ДЭБ через Selenium + pyautogui.

    log_fn — опциональная функция(str) для вывода прогресса.
    Возвращает (success: bool, message: str).
    """
    def _log(msg):
        if log_fn:
            log_fn(msg)

    if not HAS_SELENIUM:
        return False, "selenium не установлен (pip install selenium)"
    if not HAS_PYAUTOGUI:
        return False, "pyautogui не установлен (pip install pyautogui)"

    _log(f"Старт: {dispatch_name}")
    driver = None
    try:
        opts = ChromeOptions()
        driver = webdriver.Chrome(options=opts)
        wait = WebDriverWait(driver, 30)

        # 1. Открываем каталог карт
        _log(f"  Открываю каталог: {DEB_MAPS_URL}")
        driver.get(DEB_MAPS_URL)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'table tbody tr')))

        # 2. Ищем строку объекта
        _log(f"  Ищу строку: {dispatch_name}")
        rows = driver.find_elements(By.CSS_SELECTOR, 'table tbody tr')
        target_row = None
        dn_low = dispatch_name.lower()
        for row in rows:
            try:
                tds = row.find_elements(By.TAG_NAME, 'td')
                for td in tds:
                    if dn_low in td.text.lower():
                        target_row = row
                        break
            except Exception:
                pass
            if target_row:
                break

        if not target_row:
            driver.quit()
            return False, f"Объект не найден в ДЭБ: {dispatch_name}"

        # 3. Переходим по ссылке на карточку
        link_btn = target_row.find_element(By.CSS_SELECTOR, 'a[title="Перейти по ссылке"]')
        card_url = DEB_BASE_URL + link_btn.get_attribute('href').lstrip('.')
        _log(f"  Карточка: {card_url}")
        driver.get(card_url)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '.edit-mode-button')))

        # 4. Жмём «Редактировать» — подтверждение через UI выше
        driver.find_element(By.CSS_SELECTOR, 'button.edit-mode-button').click()
        time.sleep(1.5)

        # 5. Загружаем файлы
        file_items = driver.find_elements(By.CSS_SELECTOR, '.list-group-item')
        uploaded = 0
        for item in file_items:
            try:
                filename_el = item.find_element(By.CSS_SELECTOR, 'a.dz-filename')
                fname_text  = filename_el.text.lower()
                is_visio = any(ext in fname_text for ext in ('.vsd', '.vsdx'))
                is_pdf   = '.pdf' in fname_text
                upload_path = visio_path if is_visio else (pdf_path if is_pdf else None)
                if not upload_path or not os.path.exists(upload_path):
                    continue

                _log(f"  Загружаю: {os.path.basename(upload_path)}")
                change_btn = item.find_element(By.CSS_SELECTOR, 'a.change-file-button')
                change_btn.click()
                time.sleep(0.8)

                try:
                    confirm_btn = wait.until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.btn.btn-success')))
                    confirm_btn.click()
                    time.sleep(0.8)
                except Exception:
                    pass

                time.sleep(1.5)
                pyautogui.hotkey('ctrl', 'a')
                pyautogui.typewrite(upload_path, interval=0.02)
                pyautogui.press('enter')
                time.sleep(2.0)
                uploaded += 1

            except Exception as exc:
                _log(f"  Пропуск файла: {exc}")
                continue

        if uploaded == 0:
            driver.quit()
            return False, "Не удалось найти файловые поля на странице"

        # 6. Сохранить
        save_btn = driver.find_element(By.CSS_SELECTOR, 'button.confirm-button.btn-outline-success')
        save_btn.click()
        time.sleep(2.0)
        driver.quit()
        _log(f"  Загружено файлов: {uploaded}")
        return True, f"Загружено файлов: {uploaded}"

    except Exception as exc:
        try:
            if driver:
                driver.quit()
        except Exception:
            pass
        return False, str(exc)


# ── Главное окно ─────────────────────────────────────────────────────────────

class DebApp(_BASE_CLASS):
    """
    Программа 4: Регистрация карт уставок в ДЭБ.

    Публичный интерфейс:
      .ustavki_entries   list[dict]  — обрабатываемые записи
    """

    def __init__(self):
        super().__init__()
        self.ustavki_entries: list = []

        session = load_session()
        if 'ustavki' in session:
            self.ustavki_entries = session['ustavki']

        self.title("Таблицы уставок — Регистрация в ДЭБ  v2")
        self.resizable(True, True)
        self._build_ui()
        self._center_window()

    # ── UI ────────────────────────────────────────────────────────────────

    def _build_ui(self):
        root = ttk.Frame(self, padding=6)
        root.grid(row=0, column=0, sticky='nsew')
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(1, weight=1)

        # Настройки
        cfg = ttk.LabelFrame(root, text="Параметры ДЭБ", padding=6)
        cfg.grid(row=0, column=0, sticky='ew', pady=(0,4))
        cfg.columnconfigure(1, weight=1)
        ttk.Label(cfg, text="URL каталога:").grid(row=0, column=0, sticky='e', padx=(0,6))
        self._url_var = tk.StringVar(value=DEB_MAPS_URL)
        ttk.Entry(cfg, textvariable=self._url_var, width=60).grid(row=0, column=1, sticky='ew')
        ttk.Label(cfg, text="Папка карт Visio:").grid(row=1, column=0, sticky='e', padx=(0,6))
        self._maps_folder_var = tk.StringVar(value=MAPS_FOLDER)
        ttk.Entry(cfg, textvariable=self._maps_folder_var, width=60).grid(row=1, column=1, sticky='ew')
        ttk.Button(cfg, text="…", command=lambda: self._browse_folder(self._maps_folder_var),
                   width=3).grid(row=1, column=2)
        ttk.Label(cfg, text="Папка PDF:").grid(row=2, column=0, sticky='e', padx=(0,6))
        self._pdf_folder_var = tk.StringVar(value=MAPS_PDF_FOLDER)
        ttk.Entry(cfg, textvariable=self._pdf_folder_var, width=60).grid(row=2, column=1, sticky='ew')
        ttk.Button(cfg, text="…", command=lambda: self._browse_folder(self._pdf_folder_var),
                   width=3).grid(row=2, column=2)

        # Таблица записей
        lf = ttk.LabelFrame(root, text="Записи для загрузки", padding=4)
        lf.grid(row=1, column=0, sticky='nsew', pady=4)
        lf.columnconfigure(0, weight=1)
        lf.rowconfigure(0, weight=1)

        cols = ('dispatch', 'short', 'visio', 'pdf', 'status')
        self._tree = ttk.Treeview(lf, columns=cols, show='headings', height=10)
        self._tree.heading('dispatch', text='Дисп. наим.'); self._tree.column('dispatch', width=220)
        self._tree.heading('short',    text='Объект (папка)'); self._tree.column('short',    width=120)
        self._tree.heading('visio',    text='Visio файл');  self._tree.column('visio',    width=160)
        self._tree.heading('pdf',      text='PDF файл');    self._tree.column('pdf',      width=160)
        self._tree.heading('status',   text='Статус');      self._tree.column('status',   width=120)
        vsb = ttk.Scrollbar(lf, orient='vertical',   command=self._tree.yview)
        hsb = ttk.Scrollbar(lf, orient='horizontal', command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self._tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        self._tree.bind('<Double-1>', self._on_tree_dclick)

        # Ручное добавление
        add_frame = ttk.LabelFrame(root, text="Добавить запись вручную", padding=6)
        add_frame.grid(row=2, column=0, sticky='ew', pady=4)
        add_frame.columnconfigure(1, weight=1)
        add_frame.columnconfigure(3, weight=1)
        ttk.Label(add_frame, text="Дисп. наим.:").grid(row=0, column=0, sticky='e', padx=(0,6))
        self._manual_dispatch = tk.StringVar()
        ttk.Entry(add_frame, textvariable=self._manual_dispatch, width=30).grid(row=0, column=1, sticky='ew')
        ttk.Label(add_frame, text="Visio (.vsdx):").grid(row=0, column=2, sticky='e', padx=(10,6))
        self._manual_visio = tk.StringVar()
        ttk.Entry(add_frame, textvariable=self._manual_visio, width=30).grid(row=0, column=3, sticky='ew')
        ttk.Button(add_frame, text="…", command=self._browse_visio, width=3).grid(row=0, column=4)
        ttk.Label(add_frame, text="PDF:").grid(row=1, column=2, sticky='e', padx=(10,6))
        self._manual_pdf = tk.StringVar()
        ttk.Entry(add_frame, textvariable=self._manual_pdf, width=30).grid(row=1, column=3, sticky='ew')
        ttk.Button(add_frame, text="…", command=self._browse_pdf, width=3).grid(row=1, column=4)
        ttk.Button(add_frame, text="Добавить запись",
                   command=self._add_manual_entry).grid(row=1, column=0, columnspan=2, pady=4)

        # Лог
        lf2 = ttk.LabelFrame(root, text="Лог", padding=4)
        lf2.grid(row=3, column=0, sticky='ew', pady=4)
        lf2.columnconfigure(0, weight=1)
        self._log_txt = tk.Text(lf2, height=8, wrap='word', state='disabled', font=('Consolas', 9))
        sblog = ttk.Scrollbar(lf2, orient='vertical', command=self._log_txt.yview)
        self._log_txt.configure(yscrollcommand=sblog.set)
        self._log_txt.grid(row=0, column=0, sticky='ew')
        sblog.grid(row=0, column=1, sticky='ns')

        # Кнопки
        bot = ttk.Frame(root)
        bot.grid(row=4, column=0, pady=6, sticky='ew')
        ttk.Button(bot, text="Загрузить сессию",
                   command=self._load_session).pack(side='left', padx=4)
        ttk.Button(bot, text="Обновить таблицу из сессии",
                   command=self._refresh_from_session).pack(side='left', padx=4)
        ttk.Button(bot, text="Загрузить всё в ДЭБ →",
                   command=self._upload_all).pack(side='left', padx=8)
        ttk.Button(bot, text="Загрузить выбранные",
                   command=self._upload_selected).pack(side='left', padx=4)
        ttk.Button(bot, text="Закрыть", command=self.destroy).pack(side='right', padx=4)

        self._refresh_tree()

    def _browse_folder(self, var: tk.StringVar):
        d = filedialog.askdirectory(initialdir=var.get() or os.path.expanduser('~'), parent=self)
        if d:
            var.set(d)

    def _browse_visio(self):
        f = filedialog.askopenfilename(
            title="Выберите .vsdx/.vsd",
            filetypes=[("Visio файлы", "*.vsdx *.vsd"), ("Все файлы", "*.*")],
            parent=self)
        if f:
            self._manual_visio.set(f)

    def _browse_pdf(self):
        f = filedialog.askopenfilename(
            title="Выберите PDF",
            filetypes=[("PDF файлы", "*.pdf"), ("Все файлы", "*.*")],
            parent=self)
        if f:
            self._manual_pdf.set(f)

    # ── Логирование ───────────────────────────────────────────────────────

    def _log(self, text: str):
        self._log_txt.configure(state='normal')
        self._log_txt.insert('end', text + '\n')
        self._log_txt.see('end')
        self._log_txt.configure(state='disabled')
        self.update_idletasks()

    # ── Управление таблицей ───────────────────────────────────────────────

    def _refresh_tree(self):
        self._tree.delete(*self._tree.get_children())
        maps_folder = self._maps_folder_var.get()
        pdf_folder  = self._pdf_folder_var.get()
        for entry in self.ustavki_entries:
            dispatch = entry.get('dispatch_name', '')
            short    = entry.get('short_name', '') or match_object_to_short_name(entry.get('object_name',''))
            if short:
                entry['short_name'] = short
            visio = os.path.join(maps_folder, short + '.vsdx') if short else ''
            pdf   = os.path.join(pdf_folder,  short + '.pdf')  if short else ''
            # Если вручную задано — использовать его
            visio = entry.get('_deb_visio', visio)
            pdf   = entry.get('_deb_pdf', pdf)
            self._tree.insert('', 'end', values=(
                dispatch, short,
                os.path.basename(visio) if visio else '—',
                os.path.basename(pdf)   if pdf   else '—',
                entry.get('status', 'ожидание'),
            ))

    def _add_manual_entry(self):
        dispatch = self._manual_dispatch.get().strip()
        if not dispatch:
            messagebox.showwarning("Внимание", "Введите диспетчерское наименование.", parent=self)
            return
        entry = dict(EMPTY_USTAVKI_ENTRY)
        entry['dispatch_name'] = dispatch
        entry['_deb_visio']    = self._manual_visio.get()
        entry['_deb_pdf']      = self._manual_pdf.get()
        # Найти запись с таким dispatch_name
        for e in self.ustavki_entries:
            if e.get('dispatch_name','') == dispatch:
                e['_deb_visio'] = entry['_deb_visio']
                e['_deb_pdf']   = entry['_deb_pdf']
                self._refresh_tree()
                return
        self.ustavki_entries.append(entry)
        self._refresh_tree()
        self._manual_dispatch.set('')

    def _on_tree_dclick(self, event):
        """Двойной клик — редактировать пути для выбранной записи."""
        row_id = self._tree.identify_row(event.y)
        if not row_id:
            return
        all_ids = list(self._tree.get_children())
        try:
            idx = all_ids.index(row_id)
        except ValueError:
            return
        if idx >= len(self.ustavki_entries):
            return
        entry = self.ustavki_entries[idx]
        maps_folder = self._maps_folder_var.get()
        pdf_folder  = self._pdf_folder_var.get()
        short       = entry.get('short_name','')

        dlg = tk.Toplevel(self)
        dlg.title("Редактировать пути")
        dlg.grab_set()
        dlg.columnconfigure(1, weight=1)

        v_visio = tk.StringVar(value=entry.get('_deb_visio') or
                               os.path.join(maps_folder, short + '.vsdx') if short else '')
        v_pdf   = tk.StringVar(value=entry.get('_deb_pdf') or
                               os.path.join(pdf_folder, short + '.pdf') if short else '')

        ttk.Label(dlg, text=f"Запись: {entry.get('dispatch_name','')}",
                  font=('','10','bold')).grid(row=0, column=0, columnspan=3, padx=12, pady=(10,4), sticky='w')
        ttk.Label(dlg, text="Visio (.vsdx/.vsd):").grid(row=1, column=0, sticky='e', padx=(12,6), pady=4)
        ttk.Entry(dlg, textvariable=v_visio, width=46).grid(row=1, column=1, sticky='ew', pady=4)
        ttk.Button(dlg, text="…", width=3,
                   command=lambda: v_visio.set(filedialog.askopenfilename(
                       filetypes=[("Visio","*.vsdx *.vsd"),("Все","*.*")]) or v_visio.get())
                   ).grid(row=1, column=2)
        ttk.Label(dlg, text="PDF:").grid(row=2, column=0, sticky='e', padx=(12,6), pady=4)
        ttk.Entry(dlg, textvariable=v_pdf, width=46).grid(row=2, column=1, sticky='ew', pady=4)
        ttk.Button(dlg, text="…", width=3,
                   command=lambda: v_pdf.set(filedialog.askopenfilename(
                       filetypes=[("PDF","*.pdf"),("Все","*.*")]) or v_pdf.get())
                   ).grid(row=2, column=2)

        def _save():
            entry['_deb_visio'] = v_visio.get()
            entry['_deb_pdf']   = v_pdf.get()
            self._refresh_tree()
            dlg.destroy()

        ttk.Button(dlg, text="OK", command=_save).grid(row=3, column=1, pady=8)
        dlg.transient(self)
        dlg.wait_window()

    # ── Загрузка в ДЭБ ───────────────────────────────────────────────────

    def _upload_all(self):
        if not self.ustavki_entries:
            messagebox.showwarning("Нет записей", "Список пуст.", parent=self)
            return
        self._upload_entries(list(range(len(self.ustavki_entries))))

    def _upload_selected(self):
        sel = self._tree.selection()
        if not sel:
            messagebox.showwarning("Нет выбора", "Выберите строки для загрузки.", parent=self)
            return
        all_ids = list(self._tree.get_children())
        idxs = [all_ids.index(s) for s in sel if s in all_ids]
        self._upload_entries(idxs)

    def _upload_entries(self, idxs: list):
        maps_folder = self._maps_folder_var.get()
        pdf_folder  = self._pdf_folder_var.get()

        # Подготовить список с показом что будет загружено
        items_to_show = []
        for idx in idxs:
            if idx >= len(self.ustavki_entries):
                continue
            entry = self.ustavki_entries[idx]
            dispatch = entry.get('dispatch_name', '?')
            short    = entry.get('short_name', '') or match_object_to_short_name(entry.get('object_name',''))
            visio    = entry.get('_deb_visio') or (os.path.join(maps_folder, short+'.vsdx') if short else '')
            pdf      = entry.get('_deb_pdf')  or (os.path.join(pdf_folder,  short+'.pdf')  if short else '')
            items_to_show.append((idx, entry, dispatch, visio, pdf))

        if not items_to_show:
            return

        # Подтверждение
        preview = '\n'.join(
            f"  {dispatch}\n"
            f"    Visio: {os.path.basename(visio) if visio else '—'}  "
            f"PDF: {os.path.basename(pdf) if pdf else '—'}"
            for _, _, dispatch, visio, pdf in items_to_show
        )
        msg = (f"Загрузить в ДЭБ {len(items_to_show)} карт(ы) уставок?\n\n"
               f"{preview}\n\n"
               f"URL: {self._url_var.get()}\n\n"
               f"Браузер Chrome откроется автоматически.\nПродолжить?")
        if not messagebox.askyesno("Подтверждение", msg, parent=self):
            return

        # Загрузка
        all_ids = list(self._tree.get_children())
        for idx, entry, dispatch, visio, pdf in items_to_show:
            self._log(f"=== Начало: {dispatch} ===")
            ok, msg_result = upload_to_deb_entry(dispatch, visio, pdf, log_fn=self._log)
            entry['status'] = 'загружено' if ok else f'ошибка: {msg_result[:30]}'
            self._log(f"{'OK' if ok else 'ОШИБКА'}  {msg_result}\n")
            # Обновить статус в таблице
            if idx < len(all_ids):
                row_id = all_ids[idx]
                vals = list(self._tree.item(row_id, 'values'))
                if len(vals) >= 5:
                    vals[4] = entry['status']
                    self._tree.item(row_id, values=vals)

        self._save_session()
        messagebox.showinfo("Завершено",
            f"Загрузка завершена. Проверьте лог.", parent=self)

    # ── Сессия ────────────────────────────────────────────────────────────

    def _load_session(self):
        session = load_session()
        if 'ustavki' in session:
            self.ustavki_entries = session['ustavki']
            self._refresh_tree()
            messagebox.showinfo("Сессия",
                f"Загружено {len(self.ustavki_entries)} записей.", parent=self)
        else:
            messagebox.showinfo("Сессия", "Данные уставок не найдены.", parent=self)

    def _refresh_from_session(self):
        """Обновляет таблицу из сессии без сообщения."""
        session = load_session()
        if 'ustavki' in session:
            self.ustavki_entries = session['ustavki']
        self._refresh_tree()

    def _save_session(self):
        session = load_session()
        session['ustavki'] = self.ustavki_entries
        save_session(session)

    # ── Центровка ─────────────────────────────────────────────────────────

    def _center_window(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{max(w,900)}x{max(h,700)}+{(sw-max(w,900))//2}+{(sh-max(h,700))//2}")


# ── Точка входа ───────────────────────────────────────────────────────────────

def main():
    missing = []
    if not HAS_SELENIUM:
        missing.append("selenium  (pip install selenium)")
    if not HAS_PYAUTOGUI:
        missing.append("pyautogui  (pip install pyautogui)")
    if missing:
        root = tk.Tk(); root.withdraw()
        messagebox.showwarning("Зависимости не установлены",
            "Загрузка в ДЭБ недоступна без:\n  " + '\n  '.join(missing) +
            "\n\nПрограмма запустится в режиме просмотра.")
        root.destroy()

    app = DebApp()
    app.mainloop()


if __name__ == '__main__':
    main()
