# -*- coding: utf-8 -*-
"""Сборка .exe одним файлом.

Запуск:  python build_exe.py

Что важно: templates/, abbrev.json и config.json НЕ упаковываются внутрь —
они кладутся рядом с .exe, чтобы их можно было править без пересборки.
Программа ищет их рядом с собой (см. _app_dir в perechni_core).

Результат — папка «Перечни сигналов» рядом с проектом:
    Перечни сигналов/
        Perechni-signalov-vX.Y.exe
        templates/
        abbrev.json
"""
import os
import re
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
DIST = os.path.join(HERE, "dist")
WORK = os.path.join(HERE, "build")


def app_version():
    src = open(os.path.join(HERE, "perechni_gui.py"), encoding="utf-8").read()
    m = re.search(r'APP_VER\s*=\s*"([^"]+)"', src)
    return m.group(1) if m else "0.0"


def main():
    ver = app_version()
    name = "Perechni-signalov-v%s" % ver
    out_dir = os.path.join(HERE, "Перечни сигналов")

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm", "--clean",
        "--onefile",
        "--windowed",                       # без консольного окна
        "--name", name,
        "--distpath", DIST,
        "--workpath", WORK,
        "--specpath", WORK,
        "--collect-all", "tkinterdnd2",     # тянет свои tcl-библиотеки
        "--hidden-import", "openpyxl.cell._writer",
        # то, что в программе не используется и только раздувает сборку
        "--exclude-module", "PyQt5",
        "--exclude-module", "PyQt6",
        "--exclude-module", "PySide2",
        "--exclude-module", "PySide6",
        "--exclude-module", "IPython",
        "--exclude-module", "pytest",
        "--exclude-module", "pandas",
        "--exclude-module", "scipy",
    ]
    ico = os.path.join(HERE, "app.ico")
    if os.path.exists(ico):
        cmd += ["--icon", ico]
    cmd.append(os.path.join(HERE, "perechni_gui.py"))

    print("Сборка %s ..." % name, flush=True)
    r = subprocess.run(cmd, cwd=HERE)
    if r.returncode != 0:
        print("СБОРКА НЕ УДАЛАСЬ, код %d" % r.returncode)
        return 1

    exe = os.path.join(DIST, name + ".exe")
    if not os.path.exists(exe):
        print("НЕ НАЙДЕН результат:", exe)
        return 1

    # Раскладываем рядом с .exe то, что должно оставаться правимым.
    # Сначала убеждаемся, что старый файл не занят: если программа запущена,
    # копирование упадёт уже ПОСЛЕ удаления папки, и результат будет разрушен.
    old_exe = os.path.join(out_dir, os.path.basename(exe))
    if os.path.exists(old_exe):
        try:
            with open(old_exe, "ab"):
                pass
        except OSError:
            print()
            print("НЕ МОГУ ПЕРЕЗАПИСАТЬ: %s" % old_exe)
            print("Похоже, программа запущена. Закройте её и повторите сборку.")
            print("Новый файл при этом уже собран: %s" % exe)
            return 1
    if os.path.isdir(out_dir):
        shutil.rmtree(out_dir, ignore_errors=True)
    os.makedirs(out_dir, exist_ok=True)
    shutil.copy2(exe, os.path.join(out_dir, os.path.basename(exe)))
    shutil.copytree(os.path.join(HERE, "templates"),
                    os.path.join(out_dir, "templates"))
    for f in ("abbrev.json", "README.md"):
        p = os.path.join(HERE, f)
        if os.path.exists(p):
            shutil.copy2(p, os.path.join(out_dir, f))
    ex = os.path.join(HERE, "examples")
    if os.path.isdir(ex):
        shutil.copytree(ex, os.path.join(out_dir, "examples"))

    size = os.path.getsize(exe) / 1e6
    print()
    print("ГОТОВО: %s (%.0f МБ)" % (out_dir, size))
    for root, dirs, files in os.walk(out_dir):
        lvl = root[len(out_dir):].count(os.sep)
        print("   " * lvl + os.path.basename(root) + os.sep)
        for f in sorted(files):
            print("   " * (lvl + 1) + f)
    return 0


if __name__ == "__main__":
    sys.exit(main())
