---
name: build-exe
description: Use after modifying any .py file to rebuild Windows exe artifacts and commit them together with the source changes.
---

## When to use

После любого изменения `.py` файла проекта. Никогда не коммитить исходники без пересборки exe.

## Steps

1. **Проверить синтаксис** изменённых файлов:
   ```bash
   python3 -m py_compile 1_letter_register.py
   python3 -m py_compile landocs_register.py
   python3 -m py_compile 2_ustavki_folders.py
   ```

2. **Собрать все exe**:
   ```bash
   python3 build_exe.py
   ```
   Ожидаемый результат: три строки `OK → dist/<name>.exe`

3. **Скопировать в корень** (только изменённые программы):
   ```bash
   cp dist/1_LetterRegister.exe .       # если менялся 1_letter_register.py
   cp dist/LanDocs_Registrator.exe .    # если менялся landocs_register.py
   ```
   `2_UstavkiFolders.exe` в корень не копируется — только `dist/`.

4. **Закоммитить всё вместе**:
   ```bash
   git add <изменённые .py> dist/*.exe *.exe
   git commit -m "..."
   git push -u origin <ветка>
   ```

## Условия остановки

- `build_exe.py` вернул FAILED для любого exe → **остановиться**, разобраться с ошибкой
- Синтаксическая ошибка → **исправить** до сборки

## Rollback

```bash
git restore dist/*.exe *.exe   # восстановить предыдущие exe
```
