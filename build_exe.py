"""
Кросс-компиляция Python -> Windows .exe на Linux через Wine.
Запуск: python3 build_exe.py
"""

import subprocess
import os
import sys

PYWIN = os.path.expanduser("~/.wine/drive_c/Python311NuGet")
PYI = os.path.join(PYWIN, "Scripts", "pyinstaller.exe")
BASE = os.path.dirname(os.path.abspath(__file__))


def to_win(linux_path):
    return "Z:\\" + linux_path.lstrip("/").replace("/", "\\")


def build_exe(script, name, extra_data=None, extra_args=None, icon=None):
    dist = to_win(f"{BASE}/dist")
    work = f"Z:\\tmp\\pyi_{name}"
    src  = to_win(f"{BASE}/{script}")

    cmd = [
        "wine", PYI,
        src, "--onefile", "--windowed", "--clean", "--noconfirm",
        f"--name={name}",
        f"--distpath={dist}",
        f"--workpath={work}",
        f"--specpath={work}",
        "--hidden-import=tkinter",
        "--hidden-import=tkinter.ttk",
        "--collect-all=tkinter",
    ]

    if icon:
        cmd.append(f"--icon={to_win(icon)}")
        # Кладём ico в папку icons/ внутри exe, чтобы iconbitmap() нашёл его в _MEIPASS
        cmd.append(f"--add-data={to_win(icon)}:icons")

    if extra_args:
        cmd.extend(extra_args)

    shared = os.path.join(BASE, "shared_lib.py")
    if os.path.exists(shared):
        cmd.append(f"--add-data={to_win(shared)}:.")

    if extra_data:
        for src_path, dest in extra_data:
            cmd.append(f"--add-data={to_win(src_path)}:{dest}")

    print(f"\n>>> Building {name}.exe ...")
    result = subprocess.run(
        cmd,
        stdin=subprocess.PIPE,   # КРИТИЧНО для Wine — без PIPE падает с Invalid handle
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
    )
    out = result.stdout.decode("utf-8", errors="replace")

    # Показываем только важные строки
    for line in out.splitlines():
        if any(x in line for x in ["WARNING", "ERROR", "EXE", "Build complete", "broken"]):
            print(" ", line)

    if result.returncode == 0:
        exe = os.path.join(BASE, "dist", f"{name}.exe")
        size = os.path.getsize(exe) // 1024 // 1024 if os.path.exists(exe) else "?"
        print(f"    OK  →  dist/{name}.exe  ({size} MB)")
    else:
        print(f"    FAILED (rc={result.returncode})")
        print(out[-1000:])

    return result.returncode == 0


WIN32COM_ARGS = [
    "--hidden-import=win32com",
    "--hidden-import=win32com.client",
    "--hidden-import=win32com.server",
    "--hidden-import=win32com.shell",
    "--collect-all=win32com",
    "--collect-all=win32",
    "--collect-all=win32comext",
    "--hidden-import=win32api",
    "--hidden-import=win32con",
    "--hidden-import=pywintypes",
    "--hidden-import=win32clipboard",
]

SELENIUM_ARGS = [
    "--hidden-import=selenium",
    "--hidden-import=selenium.webdriver",
    "--hidden-import=selenium.webdriver.chrome",
    "--hidden-import=selenium.webdriver.chrome.service",
    "--hidden-import=selenium.webdriver.chrome.options",
    "--hidden-import=selenium.webdriver.support.ui",
    "--hidden-import=selenium.webdriver.support.expected_conditions",
    "--collect-all=selenium",
    "--hidden-import=pyautogui",
    "--collect-all=pyautogui",
]


def main():
    icons_dir = os.path.join(BASE, "icons")
    programs = [
        ("1_letter_register.py", "1_LetterRegister", None, None,
         os.path.join(icons_dir, "1_letter.ico")),
        # 2_UstavkiFolders включает шаги 0–7 (Visio + ДЭБ)
        ("2_ustavki_folders.py", "2_UstavkiFolders", None,
         WIN32COM_ARGS + SELENIUM_ARGS,
         os.path.join(icons_dir, "2_folders.ico")),
    ]

    os.makedirs(os.path.join(BASE, "dist"), exist_ok=True)
    results = []
    for script, name, extra_data, extra_args, icon in programs:
        icon_path = icon if (icon and os.path.exists(icon)) else None
        results.append(build_exe(script, name, extra_data, extra_args, icon_path))
    sys.exit(0 if all(results) else 1)


if __name__ == "__main__":
    main()
