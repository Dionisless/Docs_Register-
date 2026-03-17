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


def build_exe(script, name, extra_data=None):
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


def main():
    programs = [
        ("1_letter_register.py", "1_LetterRegister"),
        ("2_ustavki_folders.py", "2_UstavkiFolders"),
        ("3_ustavki_map.py",     "3_UstavkiMap"),
        # 4_deb_selenium.py — требует selenium/pywin32, собирается отдельно
    ]

    os.makedirs(os.path.join(BASE, "dist"), exist_ok=True)
    ok = all(build_exe(script, name) for script, name in programs)
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
