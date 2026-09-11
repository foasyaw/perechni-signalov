# -*- coding: utf-8 -*-
"""Проверка собранного .exe.

Запуск:  python check_build.py

Смотрит то, что чаще всего теряется при сборке: попали ли внутрь тяжёлые
зависимости, лежат ли рядом правимые файлы, запускается ли программа и не
пишет ли она сразу в «ошибка.log».
"""
import os
import subprocess
import sys
import time
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "Перечни сигналов")

BAD = []


def ok(cond, good, bad):
    print(("  ок    " if cond else "  ПЛОХО ") + (good if cond else bad))
    if not cond:
        BAD.append(bad)
    return cond


def find_exe():
    if not os.path.isdir(OUT):
        return ""
    for f in os.listdir(OUT):
        if f.lower().endswith(".exe"):
            return os.path.join(OUT, f)
    return ""


def main():
    print("=" * 66)
    print("Проверка сборки")
    print("=" * 66)

    exe = find_exe()
    if not ok(bool(exe), "exe найден: %s" % os.path.basename(exe or ""),
              "не найден .exe в папке «Перечни сигналов»"):
        return 1
    size = os.path.getsize(exe) / 1e6
    ok(size > 20, "размер %.0f МБ" % size,
       "размер всего %.1f МБ — похоже, зависимости не попали" % size)

    print("\nПравимые файлы рядом с exe")
    ok(os.path.isdir(os.path.join(OUT, "templates")),
       "templates/ на месте", "нет папки templates рядом с exe")
    for name in ("Перечень входных сигналов.docx", "Перечень выходных сигналов.docx"):
        p = os.path.join(OUT, "templates", name)
        ok(os.path.exists(p), "шаблон: " + name, "нет шаблона: " + name)
    ok(os.path.exists(os.path.join(OUT, "abbrev.json")),
       "abbrev.json на месте", "нет abbrev.json рядом с exe")
    ok(not os.path.exists(os.path.join(OUT, "config.json")),
       "config.json не подложен (создастся сам при первом запуске)",
       "config.json уже лежит — в сборку не должен попадать")

    print("\nЧто упаковано внутрь")
    try:
        with open(exe, "rb") as f:
            blob = f.read()
        for mark, what in ((b"matplotlib", "matplotlib (печать PDF)"),
                           (b"ezdxf", "ezdxf (чтение DXF)"),
                           (b"tkinterdnd2", "tkinterdnd2 (перетаскивание)"),
                           (b"openpyxl", "openpyxl (Excel)"),
                           (b"docx", "python-docx (Word)")):
            ok(mark in blob, "внутри есть " + what, "НЕ найдено внутри: " + what)
    except Exception as e:
        print("  не удалось прочитать exe:", e)

    print("\nЗапуск")
    log = os.path.join(OUT, "ошибка.log")
    if os.path.exists(log):
        os.remove(log)
    try:
        p = subprocess.Popen([exe], cwd=OUT)
        time.sleep(18)                      # окну нужно время подняться
        alive = p.poll() is None
        ok(alive, "программа запустилась и держится",
           "программа завершилась сама, код %s" % p.poll())
        if alive:
            # onefile-сборка запускает дочерний процесс: terminate() убивает
            # только родителя, а распакованный exe остаётся жить и держит файл
            subprocess.run(["taskkill", "/F", "/T", "/PID", str(p.pid)],
                           capture_output=True)
            try:
                p.wait(timeout=10)
            except Exception:
                p.kill()
    except Exception as e:
        ok(False, "", "не удалось запустить: %s" % e)
    ok(not os.path.exists(log), "«ошибка.log» не появился",
       "появился «ошибка.log» — смотрите его содержимое")

    print("\n" + "=" * 66)
    if BAD:
        print("ЕСТЬ ЗАМЕЧАНИЯ: %d" % len(BAD))
        for b in BAD:
            print("   • " + b)
        return 1
    print("Сборка выглядит рабочей.")
    print("Дальше стоит проверить руками: печать PDF по чертежу и")
    print("перетаскивание файлов мышью — это не проверить без окна.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
