---
name: add-feature
description: Use when adding a new feature or fixing a bug in any of the three main programs (1_letter_register, landocs_register, 2_ustavki_folders).
---

## Decision matrix: в какой файл идут изменения

| Задача | Файл |
|--------|------|
| Регистрация писем (standalone, кнопка «→ Таблицы уставок») | `1_letter_register.py` |
| Регистрация писем + таблицы уставок (монолит) | `landocs_register.py` |
| Только таблицы уставок (7 шагов) | `2_ustavki_folders.py` |
| Общие константы/утилиты без UI | `shared_lib.py` |

## Steps

1. **Прочитать** нужный `.py` целиком или найти нужную функцию через Grep
2. **Внести изменения** — только то, что попросили, без рефакторинга вокруг
3. **Проверить синтаксис**: `python3 -m py_compile <файл>`
4. **Собрать exe** через skill `build-exe`
5. **Закоммитить и запушить** исходники + exe

## Правила для UI-изменений

- Новые виджеты добавлять через отдельный метод `_build_<section>()`
- Не трогать виджеты из `threading.Thread` — только через `self.after()`
- Фоновые операции → отдельный daemon-поток

## Правила для settings

- Новый ключ в `_settings` → добавить значение по умолчанию в инициализацию словаря
- Не переименовывать существующие ключи (нарушит настройки пользователей)

## Условия остановки

- Изменение нужно в обоих `1_letter_register.py` **и** `landocs_register.py` — реализовать в обоих, пересобрать оба exe
