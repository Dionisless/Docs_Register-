# Как собрать .exe под Windows на Linux

Полная инструкция по воспроизведению кросс-компиляции Python → Windows .exe
без Windows-машины, без Docker и без wine32.

---

## Проблема

Стандартный PyInstaller собирает exe только под ту платформу, на которой запущен.
Значит с Linux сделать Windows `.exe` напрямую нельзя.

---

## Решение: Wine + NuGet Python + subprocess.Popen(PIPE)

### Ключевые открытия

1. **Wine 9.0 не умеет правильно пробрасывать консольные handles (stdin/stdout) дочерним процессам.**
   `wine cmd.exe /c "python.exe script.py"` падает с:
   ```
   Fatal Python error: init_sys_streams: can't initialize sys standard streams
   OSError: [WinError 6] Invalid handle
   ```
   Но если запустить Windows-процесс через **Linux `subprocess.Popen` с `stdin=PIPE, stdout=PIPE`** — всё работает,
   потому что Wine отлично умеет создавать PIPE-handles, в отличие от консольных.

2. **Python NuGet package не содержит tkinter** — нужно добирать файлы вручную.

3. **pythonw.exe работает через Wine**, потому что оконное приложение не инициализирует консольные streams.

---

## Пошаговая инструкция

### Шаг 1: Установить Wine

```bash
apt-get install wine  # wine 9.0+, wine32 НЕ нужен
```

### Шаг 2: Инициализировать Wine prefix

```bash
wineboot --init
```

### Шаг 3: Скачать Python 3.11 NuGet package (не embeddable, не full installer)

```bash
curl -L -o /tmp/python_nuget.nupkg \
  "https://globalcdn.nuget.org/packages/python.3.11.9.nupkg"

mkdir -p ~/.wine/drive_c/Python311NuGet
unzip -q /tmp/python_nuget.nupkg -d /tmp/python_nuget_extracted
cp -r /tmp/python_nuget_extracted/tools/* ~/.wine/drive_c/Python311NuGet/
```

> **Почему NuGet?** Полный `.exe`-installer требует wine32. Embeddable package не умеет запускать скрипты.
> NuGet package — единственный вариант, который работает под wine64.

### Шаг 4: Установить pip и PyInstaller через Linux subprocess.Popen

```bash
curl -s -o /tmp/get-pip.py https://bootstrap.pypa.io/get-pip.py
```

```python
import subprocess, os

PYWIN = os.path.expanduser("~/.wine/drive_c/Python311NuGet")
PY = os.path.join(PYWIN, "python.exe")

def wine_run(args):
    """Запускает Windows-процесс через Wine с PIPE stdio — единственный рабочий способ."""
    return subprocess.run(
        ["wine"] + args,
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
    )

# Установить pip
wine_run([PY, "Z:\\tmp\\get-pip.py", "--no-warn-script-location"])

# Установить зависимости
wine_run([PY, "-m", "pip", "install",
          "pyinstaller", "openpyxl", "python-docx", "--quiet"])
```

### Шаг 5: Добавить tkinter (NuGet Python его не включает)

Источник — conda-forge Windows пакеты (это обычные tar.bz2-архивы, работают на Linux):

```bash
# Скачать Tcl/Tk DLLs
curl -L -o /tmp/tk_win64.tar.bz2 \
  "https://conda.anaconda.org/conda-forge/win-64/tk-8.6.9-hfa6e2cd_1003.tar.bz2"
mkdir -p /tmp/tk_conda && tar -xjf /tmp/tk_win64.tar.bz2 -C /tmp/tk_conda

# Скачать Python 3.11 conda-пакет (содержит _tkinter.pyd и Lib/tkinter/)
curl -L -o /tmp/python311_conda.tar.bz2 \
  "https://conda.anaconda.org/conda-forge/win-64/python-3.11.0-hcf16a7b_0_cpython.tar.bz2"
mkdir -p /tmp/py311_conda && tar -xjf /tmp/python311_conda.tar.bz2 -C /tmp/py311_conda
```

Разложить файлы по правильным путям:

```bash
PYWIN=~/.wine/drive_c/Python311NuGet

# _tkinter.pyd
cp /tmp/py311_conda/DLLs/_tkinter.pyd "$PYWIN/DLLs/"

# tkinter Python-пакет
cp -r /tmp/py311_conda/Lib/tkinter "$PYWIN/Lib/"

# Tcl/Tk DLLs — в корень Python
cp /tmp/tk_conda/Library/bin/tcl86t.dll "$PYWIN/"
cp /tmp/tk_conda/Library/bin/tk86t.dll  "$PYWIN/"

# Tcl/Tk data-файлы — в подпапку tcl/ (именно так ищет PyInstaller)
mkdir -p "$PYWIN/tcl"
cp -r /tmp/tk_conda/Library/lib/tcl8.6 "$PYWIN/tcl/tcl8.6"
cp -r /tmp/tk_conda/Library/lib/tk8.6  "$PYWIN/tcl/tk8.6"
```

Проверка (должно вывести `tkinter OK 8.6`):

```python
result = subprocess.run(
    ["wine", PYWIN + "/pythonw.exe", "-c",
     "import tkinter; print('tkinter OK', tkinter.TkVersion)"],
    stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.STDOUT
)
print(result.stdout.decode())
```

### Шаг 6: Собрать .exe через PyInstaller

```python
import subprocess, os

PYI = os.path.expanduser("~/.wine/drive_c/Python311NuGet/Scripts/pyinstaller.exe")

def build_exe(script_win_path, name, dist_win_path, workdir_win_path,
              shared_lib_win_path=None):
    """
    script_win_path  — путь к .py в Windows-формате, например: Z:\\path\\to\\script.py
    Все пути с Z:\\ — это /  на Linux-стороне (Wine пробрасывает весь Linux-FS как Z:)
    """
    cmd = [
        "wine", PYI,
        script_win_path,
        "--onefile",          # один файл, не папка
        "--windowed",         # нет консольного окна (GUI-приложение)
        "--clean", "--noconfirm",
        f"--name={name}",
        f"--distpath={dist_win_path}",
        f"--workpath={workdir_win_path}",
        f"--specpath={workdir_win_path}",
        "--hidden-import=tkinter",
        "--hidden-import=tkinter.ttk",
        "--collect-all=tkinter",  # вшить весь tkinter включая ttk, filedialog и т.д.
    ]
    if shared_lib_win_path:
        cmd.append(f"--add-data={shared_lib_win_path}:.")

    result = subprocess.run(
        cmd,
        stdin=subprocess.PIPE,   # КРИТИЧНО: без PIPE wine падает с "Invalid handle"
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
    )
    return result
```

Пример вызова:

```python
def to_win(linux_path):
    return "Z:\\" + linux_path.lstrip("/").replace("/", "\\")

BASE = "/home/user/Docs_Register-"

result = build_exe(
    script_win_path = to_win(f"{BASE}/1_letter_register.py"),
    name            = "1_LetterRegister",
    dist_win_path   = to_win(f"{BASE}/dist"),
    workdir_win_path= "Z:\\tmp\\pyi_build",
    shared_lib_win_path = to_win(f"{BASE}/shared_lib.py"),
)
print(result.stdout.decode("utf-8", errors="replace")[-2000:])
```

---

## Структура итогового Wine Python

```
~/.wine/drive_c/Python311NuGet/
├── python.exe
├── pythonw.exe
├── python3.dll
├── python311.dll
├── vcruntime140.dll
├── vcruntime140_1.dll
├── tcl86t.dll          ← из conda-forge tk-8.6.9
├── tk86t.dll           ← из conda-forge tk-8.6.9
├── DLLs/
│   ├── _tkinter.pyd    ← из conda-forge python-3.11.0
│   └── ...
├── Lib/
│   ├── tkinter/        ← из conda-forge python-3.11.0
│   └── site-packages/
│       ├── PyInstaller/
│       ├── openpyxl/
│       ├── docx/
│       └── ...
├── Scripts/
│   └── pyinstaller.exe
└── tcl/
    ├── tcl8.6/         ← из conda-forge tk-8.6.9
    └── tk8.6/          ← из conda-forge tk-8.6.9
```

---

## Что НЕ вшивается (нужно ставить вручную)

| Модуль | Причина |
|--------|---------|
| `selenium` | Требует ChromeDriver рядом с exe и браузер |
| `pywin32` | Win32 API-зависимость, конфликтует с PyInstaller |
| `pyautogui` | Нужен screen access, не поддаётся статической линковке |

---

## Итоговые размеры

| Файл | Размер | Содержимое |
|------|--------|------------|
| `1_LetterRegister.exe` | 16 MB | Python 3.11 + tkinter/ttk + openpyxl + python-docx |
| `2_UstavkiFolders.exe` | 17 MB | то же |
| `3_UstavkiMap.exe`     | 16 MB | то же |
| `4_DebSelenium.exe`    |  6 MB | Python 3.11 + tkinter (без selenium/pywin32) |
