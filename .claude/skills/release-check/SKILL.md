---
name: release-check
description: Use before committing or pushing changes to verify all artifacts are consistent and ready.
---

## Pre-flight checklist (все пункты должны быть выполнены)

- [ ] Синтаксис всех изменённых `.py` проходит `python3 -m py_compile`
- [ ] `python3 build_exe.py` завершился: все три exe показали `OK`
- [ ] Корневые `*.exe` обновлены (`cp dist/*.exe .`)
- [ ] `git status` не показывает несохранённых изменений в `.py` и `.exe`
- [ ] Ветка называется `claude/<task>-<id>`, **не** `main`/`master`
- [ ] Коммит содержит и исходники, и exe

## Output

Для каждого пункта: ✓ Пройден / ✗ Не пройден + причина.
При любом ✗ — остановиться и исправить до push.
