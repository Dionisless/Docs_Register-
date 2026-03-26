# Python Rules

Применяется ко всем файлам `*.py` в проекте.

## Стиль

- Заголовок файла: `#!/usr/bin/env python3` + `# -*- coding: utf-8 -*-`
- Импорты stdlib сначала, потом сторонние — каждая группа через пустую строку
- Мягкий импорт зависимостей через `try/except ImportError` с флагом `HAS_*`
- Комментарии к секциям: `# ── Название ──────────` (с разделителями)

## Зависимости

- `pywin32` — клавиатура и буфер обмена; всегда проверять `HAS_WIN32`
- `openpyxl` — Excel; всегда проверять `HAS_OPENPYXL`
- `python-docx` — Word; всегда проверять `HAS_DOCX`
- `tkinterdnd2` — drag-and-drop; всегда проверять `HAS_DND`

## Настройки

- Путь к `settings.json` определяется через `_exe_dir()`:
  ```python
  def _exe_dir() -> str:
      if getattr(sys, 'frozen', False):
          return os.path.dirname(sys.executable)
      return os.path.dirname(os.path.abspath(__file__))
  ```
- Никогда не писать настройки в `%APPDATA%` — только рядом с exe
- Сессионные данные (межпрограммная передача) — в `%APPDATA%/DocsRegister/`

## UI

- Базовый класс: `tk.Tk` (или `TkinterDnD.Tk` при наличии `HAS_DND`)
- Grid-менеджер для всех виджетов, pack только в горизонтальных button-frame
- Все строки UI на русском языке
- Не трогать виджеты из фоновых потоков — только через `self.after()`
