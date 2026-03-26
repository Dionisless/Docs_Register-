# Project Contract — LanDocs Registrator

## Build And Test

```bash
# Собрать все exe (Wine + PyInstaller, Linux → Windows)
python3 build_exe.py

# Проверить синтаксис Python без сборки
python3 -m py_compile 1_letter_register.py
python3 -m py_compile 2_ustavki_folders.py
python3 -m py_compile landocs_register.py
```

После изменения любого `.py` файла **обязательно** пересобрать exe и скопировать в корень:

```bash
python3 build_exe.py
cp dist/1_LetterRegister.exe .
cp dist/LanDocs_Registrator.exe .
```

## Architecture

```
Docs_Register-/
├── 1_letter_register.py      # Регистрация входящих/исходящих (standalone)
├── 2_ustavki_folders.py      # Обработка таблиц уставок (7 шагов)
├── landocs_register.py       # Монолит v2: регистрация + уставки
├── shared_lib.py             # Общие константы и утилиты
├── build_exe.py              # Сборка через Wine+PyInstaller
├── settings.json             # Настройки (рядом с exe, не в AppData)
├── dist/                     # Артефакты сборки
└── .claude/
    ├── rules/                # Правила по типам файлов
    └── skills/               # Рабочие процессы
```

### Границы модулей

- `shared_lib.py` — только данные и утилиты без UI-зависимостей
- UI-код живёт только внутри классов `RegistrationApp` и `SettingsWindow`
- `ViewDirWatcher` — изолированный класс, не зависит от UI
- Настройки: всегда `settings.json` рядом с exe (`_exe_dir()`)

## Coding Conventions

- Кодировка файлов: UTF-8, `# -*- coding: utf-8 -*-`
- Строки комментариев и UI на русском языке
- Обёртки зависимостей через `HAS_*` флаги с мягким fallback
- Не добавлять новые глобальные переменные без явной необходимости
- Сохранять совместимость схемы `settings.json` (не переименовывать ключи)

## Safety Rails

### NEVER
- Пушить в `main`/`master` напрямую
- Менять ключи `settings.json` без миграции (ломает настройки пользователей)
- Коммитить `.py` изменения без пересборки exe
- Хардкодить пути `%APPDATA%` для постоянных данных (только для сессионных)

### ALWAYS
- После изменения `.py` → пересобрать exe → скопировать в корень → закоммитить всё вместе
- Ветки называть `claude/<task>-<id>`
- Коммитить `dist/*.exe` и корневые `*.exe` вместе с исходниками

## Verification

- Синтаксис: `python3 -m py_compile <file>`
- Сборка: `python3 build_exe.py` — все три exe должны показать `OK`
- Exe в корне должны совпадать с `dist/` по дате изменения

## Compact Instructions

При сжатии сохранять приоритетно:
- Архитектурные решения (NEVER summarize)
- Изменённые файлы и суть правок
- Статус сборки (OK / FAILED)
- Открытые TODO и заметки по откату
- Вывод инструментов можно удалить, оставить только pass/fail
