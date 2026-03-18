#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
shared_lib.py — общие константы и утилиты для всех модулей.

Каждая из 4 программ импортирует отсюда:
  from shared_lib import *
"""

import os
import re
import sys
import json
import difflib
from datetime import datetime

# ── Пути к файлам на сервере ─────────────────────────────────────────────────

EXCEL_PATH_IN  = r"\\Prim-fs-serv\primrdu\СРЗА\Журнал регистрации корреспонденции\Журнал регистрации входящей документации.xlsx"
EXCEL_PATH_OUT = r"\\Prim-fs-serv\primrdu\СРЗА\Журнал регистрации корреспонденции\Журнал регистрации исходящей документации.xlsx"
DEFAULT_SAVE_FOLDER = r"\\Prim-fs-serv\primrdu\СРЗА\Дела СРЗА\19 Переписка"

REGISTRY_PATH    = r"\\Prim-fs-serv\primrdu\СРЗА\Журнал регистрации корреспонденции\Реестр таблиц уставок.xlsx"
SUMMARY_PATH     = r"\\Prim-fs-serv\primrdu\СРЗА\Журнал регистрации корреспонденции\Регистрация таблиц уставок.xlsx"
USTAVKI_EXEC_BASE         = r"\\Prim-fs-serv\primrdu\СРЗА\Уставки\Таблицы РЗА\Таблицы для исполнения РЗА"
USTAVKI_ARCHIVE_BASE      = r"\\Prim-fs-serv\primrdu\СРЗА\Уставки\Таблицы РЗА\Таблицы уставок РЗА"
USTAVKI_REAL_ARCHIVE_BASE = r"\\Prim-fs-serv\primrdu\СРЗА\Уставки\Таблицы РЗА\Архив таблиц РЗА"
MAPS_FOLDER          = r"\\Prim-fs-serv\primrdu\СРЗА\Уставки\КАРТА УСТАВОК"
MAPS_PDF_FOLDER  = r"\\Prim-fs-serv\primrdu\СРЗА\Уставки\КАРТА УСТАВОК\ДЭБ"
DEB_BASE_URL     = "https://pri-mdeb.oduvs.so"
DEB_MAPS_URL     = "https://pri-mdeb.oduvs.so/?sid=02ab815f-a54e-42a5-8a88-36dee8a5af2e&DataAreaId=1b6fecd6-f813-47ac-aa88-de4f67b7a1ac"

# ── Папка AppData для настроек и сессии ──────────────────────────────────────

def get_appdata_dir() -> str:
    """Возвращает %APPDATA%\\DocsRegister, создаёт если нет."""
    appdata = os.environ.get('APPDATA') or os.path.join(
        os.environ.get('USERPROFILE', ''), 'AppData', 'Roaming')
    d = os.path.join(appdata, 'DocsRegister')
    os.makedirs(d, exist_ok=True)
    return d

# Файл сессии — сюда каждая программа сохраняет/читает текущие данные
SESSION_FILE = os.path.join(get_appdata_dir(), "session_data.json")

# ── Сокращённые названия объектов ─────────────────────────────────────────────

OBJECT_SHORT_NAMES = [
    "ПримГРЭС","ТЭЦ Центральная","АТЭЦ","ВТЭЦ-2","ПартГРЭС","Шкотовская ТЭЦ",
    "Восточная ТЭЦ","ТЭЦ Северная","Варяг","Владивосток","Дальневосточная","Лозовая",
    "Хехцир 2","Чугуевка 2","Арсеньев 2","Аэропорт","Береговая 2","Бикин тяговая",
    "Волна","Высокогорск","Горелое","Губерово тяговая","Западная","Звезда",
    "Зелёный угол","Иман","К","Кировка","Козьмино","Лесозаводск","Минеральная",
    "НПС-36","НПС-38","НПС-40","НПС-41","Находка","Новая","Партизанск","Патрокл",
    "Перевал","Промпарк","Розенгартовка тяговая","Ружино тяговая","Русская",
    "Свиягино тяговая","Спасск","Суходол","Уссурийск 2","Чугуевка","Широкая",
    "Шмаковка тяговая","178Ф","1Р","1Р тяговая","2Р","А","АСБ","Агрокомплекс",
    "Амурская","Анисимовка тяговая","Арсеньев 1","Барабаш","Береговая 1","Бикин",
    "Богополь","Бурная","Бурун","ВТЭЦ-1","Вадимовка","Вокзальная тяговая","Волчанец",
    "Восток 2","Восточная тяговая","Глубинная","Голдобин","Голубинка","Голубовка",
    "Горбуша","Горностай","Гранит","Давыдовка","Дальнереченск тяговая","Де-Фриз",
    "Дмитриевка","Екатериновка","ЖБИ-130","ЖБФ","З","Загородная","Залив","Игнатьевка",
    "Казармы","Камень-Рыболов","Кипарисово","Ключи","Кожзавод","Котельная 2Р",
    "Краскино","Кролевцы","ЛРЗ","Лазурная","Ласточка тяговая","Липовцы","Лучегорск",
    "М","Междуречье","Мингородок","Михайловка","Молодежная","Муравейка","Мучная",
    "Мыс Астафьева","НСРЗ","Надаровская","Надеждинская тяговая","Насосная",
    "Находка 110","Находка тяговая","Николаевка","Новицкое","Новоникольск",
    "Новопокровка","Новый мир","Океан","Ольга","Орлиная","Павловка 1","Павловка 2",
    "Песчаная","Петровичи","Плавзавод","Пластун","Подъяпольск","Полевая",
    "Преображение","Прибой","Приозерная","Прогресс","Промузел","Промысловка",
    "Промышленная","Пушкинская","Раздольное 1","Раздольное 2","Разрез","Ракушка",
    "Реттиховка","Рощино","С-55","Садовая","Седанка","Сибирцево тяговая","Славянка",
    "Смоляниново тяговая","Спасск тяговая","Спутник","Стройиндустрия","Студгородок",
    "Тайфун","Тереховка","Тимофеевка","Топаз","Троица","УКФ","Угольная","Улисс",
    "Уссурийск 1","Уссурийск тяговая","Учебная","Факел","Фридман тяговая","ХФЗ",
    "Хороль","Чайка","Черемшаны","Черниговка","Чуркин","Шахта 7","Штыково","Южная",
    "Ярославка",
]

# ── Структура записи уставки (интерфейс между программами) ───────────────────
# Каждая программа работает со списком таких словарей.
# Это "торчащий конец" для сшивания программ.
#
# UstavkiEntry:
#   file_path        str  — полный путь к .docx
#   form_type        str  — 'new' / 'old'
#   object_name      str  — полное название объекта из таблицы
#   dispatch_name    str  — диспетчерское наименование
#   table_number     str  — номер таблицы
#   outgoing_letter  str  — № исходящего письма из таблицы
#   letter_num       str  — № вх. письма (заполняется из Программы 1)
#   letter_date      str  — дата вх. письма
#   status           str  — текущий статус обработки
#   registry_row     int  — строка в реестре (после Программы 2)
#   archive_candidate str — путь к архивируемому файлу (после Программы 2)
#   current_path     str  — куда скопирован файл (после Программы 2)
#   short_name       str  — краткое имя объекта = имя папки (после Программы 2)

EMPTY_USTAVKI_ENTRY = {
    'file_path': '', 'form_type': '', 'object_name': '', 'dispatch_name': '',
    'table_number': '', 'outgoing_letter': '', 'letter_num': '', 'letter_date': '',
    'status': 'ожидание', 'registry_row': 0, 'archive_candidate': '',
    'current_path': '', 'short_name': '',
}

# ── Сессия ────────────────────────────────────────────────────────────────────

def load_session() -> dict:
    """Загружает данные сессии из JSON-файла."""
    if os.path.exists(SESSION_FILE):
        try:
            with open(SESSION_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_session(data: dict):
    """Сохраняет данные сессии в JSON-файл."""
    try:
        with open(SESSION_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as exc:
        print(f"Ошибка сохранения сессии: {exc}")


# ── Утилиты дат ───────────────────────────────────────────────────────────────

def sanitize_for_filename(text: str) -> str:
    return re.sub(r'[<>:"/\\|?*\r\n\t]', '_', text)


def parse_date(date_str: str):
    for fmt in ('%d.%m.%Y', '%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y', '%Y.%m.%d'):
        try:
            return datetime.strptime(date_str.strip(), fmt)
        except Exception:
            pass
    return None


def fmt_date_ymd(date_str: str) -> str:
    d = parse_date(date_str)
    return d.strftime('%Y-%m-%d') if d else date_str


def fmt_date_dmy(date_str: str) -> str:
    d = parse_date(date_str)
    return d.strftime('%d.%m.%Y') if d else date_str


def fmt_date_dmy_underscore(date_str: str) -> str:
    d = parse_date(date_str)
    return d.strftime('%d_%m_%Y') if d else date_str


def fmt_date_ymd_underscore(date_str: str) -> str:
    d = parse_date(date_str)
    return d.strftime('%Y_%m_%d') if d else date_str


# ── Объекты / папки ───────────────────────────────────────────────────────────

def _normalize_name(s: str) -> str:
    return re.sub(r'\s+', ' ', s.strip().lower())


def match_object_to_short_name(object_name: str) -> str:
    """
    Сопоставляет полное имя объекта с кратким (= имя папки).

    Приоритет:
    1. Точное совпадение (без учёта регистра)
    2. Если object_name — это уже известное краткое имя → возвращаем как есть
    3. Fuzzy-match по SequenceMatcher ratio
    """
    if not object_name:
        return ''
    norm = _normalize_name(object_name)

    # 1. Точное совпадение
    for short in OBJECT_SHORT_NAMES:
        if _normalize_name(short) == norm:
            return short

    # 2. Если уже является кратким именем (= имя папки)
    for short in OBJECT_SHORT_NAMES:
        if norm == _normalize_name(short):
            return short

    # 3. Fuzzy match — берём наилучшее по SequenceMatcher
    best_match, best_score = '', 0.0
    for short in OBJECT_SHORT_NAMES:
        norm_short = _normalize_name(short)
        score = difflib.SequenceMatcher(None, norm_short, norm).ratio()
        if score > best_score:
            best_score, best_match = score, short
    # Возвращаем только если score достаточно высок
    return best_match if best_score >= 0.4 else ''


def get_object_short_name_from_path(file_path: str) -> str:
    """
    Извлекает краткое имя объекта из пути файла.
    Ожидается путь вида:
      ...\\Таблицы для исполнения РЗА\\ОБЪЕКТ\\...
    или просто:
      ...\\ОБЪЕКТ\\файл.docx
    Возвращает имя папки-объекта если оно есть в OBJECT_SHORT_NAMES,
    иначе пустую строку.
    """
    parts = file_path.replace('/', '\\').split('\\')
    # Идём от конца к началу, ищем папку из OBJECT_SHORT_NAMES
    norm_shorts = {_normalize_name(s): s for s in OBJECT_SHORT_NAMES}
    for part in reversed(parts[:-1]):  # пропускаем сам файл
        if _normalize_name(part) in norm_shorts:
            return norm_shorts[_normalize_name(part)]
    return ''


def find_object_exec_folder(short_name: str) -> str | None:
    """Ищет папку объекта в базовой директории уставок."""
    if not short_name or not os.path.isdir(USTAVKI_EXEC_BASE):
        return None
    norm_target = _normalize_name(short_name)
    for entry in os.scandir(USTAVKI_EXEC_BASE):
        if entry.is_dir() and _normalize_name(entry.name) == norm_target:
            return entry.path
    return None


def find_current_and_archive_folders(object_folder: str) -> tuple:
    """Возвращает (current_dir, archive_dir) внутри папки объекта."""
    if not object_folder or not os.path.isdir(object_folder):
        return None, None
    current_dir = archive_dir = None
    for entry in os.scandir(object_folder):
        if not entry.is_dir():
            continue
        n = entry.name.lower()
        if 'текущ' in n:
            current_dir = entry.path
        elif 'архив' in n:
            archive_dir = entry.path
    return current_dir, archive_dir
