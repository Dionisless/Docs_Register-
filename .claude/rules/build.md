# Build Rules

Применяется при любых изменениях `*.py` файлов проекта.

## Обязательный порядок после правок Python

1. Проверить синтаксис: `python3 -m py_compile <изменённый файл>`
2. Собрать exe: `python3 build_exe.py`
3. Скопировать в корень:
   - `cp dist/1_LetterRegister.exe .` (если менялся `1_letter_register.py`)
   - `cp dist/LanDocs_Registrator.exe .` (если менялся `landocs_register.py`)
4. Закоммитить `.py`, `dist/*.exe` и корневые `*.exe` одним коммитом

## build_exe.py — что собирает

| Скрипт | Имя exe | Иконка |
|--------|---------|--------|
| `1_letter_register.py` | `1_LetterRegister.exe` | `icons/1_letter.ico` |
| `2_ustavki_folders.py` | `2_UstavkiFolders.exe` | `icons/2_folders.ico` |
| `landocs_register.py`  | `LanDocs_Registrator.exe` | `icons/1_letter.ico` |

## Что НЕ нужно собирать

- `shared_lib.py` — библиотека, не точка входа
- Изменение только `build_exe.py` или `.claude/` не требует пересборки

## Артефакты

- `dist/` — все собранные exe (коммитить)
- `_landocs_cache/` — кэш ViewDirWatcher, **не коммитить** (в .gitignore)
- `settings.json` — пользовательские настройки, **не коммитить** (в .gitignore)
