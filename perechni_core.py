# -*- coding: utf-8 -*-
"""
Ядро: извлечение сигналов из чертежей схем подключения (ШСК, DWG/DXF)
и заполнение Word-перечней входных/выходных сигналов по шаблону.

Логика проверена на проекте шифр проекта (ШСК2/ШСК3/ШСК4).
Ожидаемый стиль чертежа:
  - заголовки аналоговых модулей: TEXT "Модуль X.Y AI|AO|WI ..." ;
  - описания каналов: MTEXT на слое "Текст" с "поз. <ТЕГ>." ;
  - метки БИЗ: слой "mark2" (наличие -> Exia, иначе Exd);
  - дискретные каналы: якоря "X.Y-KLDI<n>" / "X.Y-KLDO<n>"
    на слоях "Реле DI (1-KL)" / "Реле DO (3-KL)".
Шаблон Word: как фирменные перечни (СОДЕРЖАНИЕ-таблица, аналоговая и
дискретная таблицы по 7 колонок, ЛИСТ РЕГИСТРАЦИИ, разрыв секции после
содержания, разрывы страниц между таблицами).
"""
import copy, io, os, re, shutil, subprocess, tempfile, statistics, collections

import ezdxf
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def read_dxf(path):
    """Читает DXF; если файл битый — восстанавливает через ezdxf.recover."""
    try:
        return ezdxf.readfile(path)
    except Exception:
        from ezdxf import recover
        doc, _auditor = recover.readfile(path)
        return doc


# ---------------------------------------------------------------- DWG -> DXF
# Два движка: AutoCAD Core Console (accoreconsole.exe, идёт с AutoCAD)
# и ODA File Converter (бесплатный, opendesign.com). Находится любой из них.

# Путь к конвертеру можно задать переменной окружения PERECHNI_ODA
# или в config.json рядом с программой — иначе ищем в обычных местах.
DEFAULT_ODA = os.environ.get("PERECHNI_ODA", "")

def _config_path():
    """config.json лежит рядом с программой.

    В собранном виде (PyInstaller) __file__ указывает внутрь временной папки
    _MEIPASS, а настройки пользователь кладёт рядом с .exe — берём sys.argv[0],
    так же как это делает GUI.
    """
    import sys
    base = os.path.dirname(os.path.abspath(sys.argv[0] or __file__))
    p = os.path.join(base, "config.json")
    if os.path.exists(p):
        return p
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

def load_config():
    """Настройки рядом с программой. Отсутствие файла — не ошибка."""
    try:
        import json
        with open(_config_path(), encoding="utf-8") as f:
            cfg = json.load(f)
        return cfg if isinstance(cfg, dict) else {}
    except FileNotFoundError:
        return {}
    except Exception:
        return {}

def find_oda():
    for cand in (DEFAULT_ODA, load_config().get("oda_exe", "")):
        if cand and os.path.exists(cand):
            return cand
    roots = [r"C:\Program Files\ODA", r"C:\Program Files (x86)\ODA"]
    for drive in "CDEFGH":                     # конвертер часто кладут не на C:
        roots.append(drive + ":\\ODA")
        roots.append(drive + ":\\Program Files\\ODA")
    for root in roots:
        if os.path.isdir(root):
            try:
                for dirpath, _dirs, files in os.walk(root):
                    if "ODAFileConverter.exe" in files:
                        return os.path.join(dirpath, "ODAFileConverter.exe")
            except OSError:
                continue
    return ""

def find_accore():
    roots = [r"C:\Program Files\Autodesk"]
    # диски C..G, папки вида "AutoCAD 20xx"
    for drive in "CDEFG":
        roots.append(drive + ":\\")
    seen = []
    for root in roots:
        if not os.path.isdir(root):
            continue
        try:
            for name in os.listdir(root):
                if "autocad" in name.lower():
                    cand = os.path.join(root, name, "accoreconsole.exe")
                    if os.path.exists(cand):
                        seen.append(cand)
        except OSError:
            continue
    return sorted(seen)[-1] if seen else ""

def find_converter():
    """(движок, путь): ('acad', ...) | ('oda', ...) | ('', '')."""
    p = find_accore()
    if p:
        return "acad", p
    p = find_oda()
    if p:
        return "oda", p
    return "", ""

def _convert_accore(dwg_paths, exe, out_dir, log):
    """DWG -> DXF через AutoCAD Core Console (по файлу за раз)."""
    tmp = tempfile.mkdtemp(prefix="perechni_acc_")
    try:
        return _convert_accore_inner(dwg_paths, exe, out_dir, log, tmp)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def _convert_accore_inner(dwg_paths, exe, out_dir, log, tmp):
    result = {}
    for i, p in enumerate(dwg_paths):
        src = os.path.join(tmp, f"in{i}.dwg")     # ASCII-пути: accoreconsole
        dst = os.path.join(tmp, f"out{i}.dxf")    # капризен к кириллице в .scr
        shutil.copy2(p, src)
        scr = os.path.join(tmp, f"c{i}.scr")
        with open(scr, "w", encoding="ascii") as f:
            f.write("_.DXFOUT\n%s\n16\n\n" % dst.replace("\\", "/"))
        log(f"  AutoCAD: {os.path.basename(p)} ...")
        # без pipe (иначе accoreconsole зависает); сам он тоже не всегда
        # завершается — ждём появления DXF и стабилизации его размера
        proc = subprocess.Popen([exe, "/i", src, "/s", scr, "/l", "en-US"],
                                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        import time
        stable, last = 0, -1
        for _ in range(300):
            if proc.poll() is not None:
                break
            sz = os.path.getsize(dst) if os.path.exists(dst) else -1
            if sz > 0 and sz == last:
                stable += 1
                if stable >= 3:      # размер не меняется 3 сек — готово
                    break
            else:
                stable = 0
            last = sz
            time.sleep(1)
        if proc.poll() is None:
            proc.kill()
        if not os.path.exists(dst):
            log(f"    ОШИБКА: {os.path.basename(p)} не сконвертировался")
            continue
        if os.path.exists(dst):
            final = os.path.join(out_dir, os.path.splitext(os.path.basename(p))[0] + ".dxf")
            shutil.move(dst, final)
            result[p] = final
            log(f"    готово: {os.path.basename(final)}")
        else:
            log(f"    ОШИБКА: {os.path.basename(p)} не сконвертировался")
    return result

def _convert_oda(dwg_paths, exe, out_dir, log):
    """DWG -> DXF через ODA File Converter (папка за раз)."""
    tmp_in = tempfile.mkdtemp(prefix="perechni_in_")
    tmp_out = tempfile.mkdtemp(prefix="perechni_out_")
    try:
        return _convert_oda_inner(dwg_paths, exe, out_dir, log, tmp_in, tmp_out)
    finally:
        shutil.rmtree(tmp_in, ignore_errors=True)
        shutil.rmtree(tmp_out, ignore_errors=True)


def _convert_oda_inner(dwg_paths, exe, out_dir, log, tmp_in, tmp_out):
    import time
    for p in dwg_paths:
        shutil.copy2(p, os.path.join(tmp_in, os.path.basename(p)))
    proc = subprocess.Popen([exe, tmp_in, tmp_out, "ACAD2018", "DXF", "0", "0", "*.dwg"])
    for _ in range(300):
        time.sleep(1)
        done = [f for f in os.listdir(tmp_out) if f.lower().endswith(".dxf")]
        if len(done) >= len(dwg_paths):
            break
    time.sleep(2)
    try:
        proc.kill()
    except Exception:
        pass
    result = {}
    for p in dwg_paths:
        dxf = os.path.join(tmp_out, os.path.splitext(os.path.basename(p))[0] + ".dxf")
        if os.path.exists(dxf):
            final = os.path.join(out_dir, os.path.basename(dxf))
            shutil.move(dxf, final)
            result[p] = final
            log(f"  готово: {os.path.basename(final)}")
        else:
            log(f"  ОШИБКА: не сконвертировался {os.path.basename(p)}")
    return result

def dwg_to_dxf(dwg_paths, out_dir=None, engine=None, exe=None, log=print):
    """Конвертирует DWG в DXF. Возвращает {dwg: dxf}.
    engine/exe можно не указывать — найдутся сами (AutoCAD, затем ODA)."""
    if not engine or not exe:
        engine, exe = find_converter()
    if not exe:
        raise RuntimeError(
            "Не найден конвертер DWG.\nНужен AutoCAD (accoreconsole.exe) или бесплатный "
            "ODA File Converter (opendesign.com).\nЛибо сохраните чертёж из AutoCAD как DXF.")
    out_dir = out_dir or tempfile.mkdtemp(prefix="perechni_dxf_")
    os.makedirs(out_dir, exist_ok=True)
    log(f"Конвертация {len(dwg_paths)} DWG -> DXF ({'AutoCAD' if engine=='acad' else 'ODA'})...")
    if engine == "acad":
        return _convert_accore(dwg_paths, exe, out_dir, log)
    return _convert_oda(dwg_paths, exe, out_dir, log)

# --------------------------------------------------- показать канал в AutoCAD

def find_acad():
    """Ищет acad.exe (задан в настройках, рядом с accoreconsole, в обычных местах)."""
    import glob as _g
    cfg = load_config().get("acad_exe", "")
    if cfg and os.path.exists(cfg):
        return cfg
    env = os.environ.get("PERECHNI_ACAD", "")
    if env and os.path.exists(env):
        return env
    pats = []
    for drive in "CDEFGH":
        pats.append(drive + r":\Program Files\Autodesk\AutoCAD*\acad.exe")
        pats.append(drive + r":\Autodesk\AutoCAD*\acad.exe")
        pats.append(drive + r":\AutoCAD*\acad.exe")
    for pat in pats:
        try:
            hits = sorted(_g.glob(pat), reverse=True)
        except OSError:
            continue
        if hits:
            return hits[0]
    try:
        eng, exe = find_converter()
        if exe and "accoreconsole" in exe.lower():
            cand = os.path.join(os.path.dirname(exe), "acad.exe")
            if os.path.exists(cand):
                return cand
    except Exception:
        pass
    return ""

def show_in_acad(drawing, x, y, log=print, half_w=130, half_h=90):
    """Открывает чертёж в AutoCAD и зумит на окно вокруг (x, y)."""
    acad = find_acad()
    if not acad:
        raise RuntimeError("Не найден acad.exe — AutoCAD не установлен?")
    if x is None or y is None:
        raise RuntimeError("Для этой строки нет координат (источник — Excel?)")
    scr_dir = os.path.join(tempfile.gettempdir(), "perechni_scr")
    os.makedirs(scr_dir, exist_ok=True)
    scr = os.path.join(scr_dir, "zoom.scr")
    with open(scr, "w", encoding="ascii") as f:
        f.write(f"_.ZOOM\n_W\n{x - half_w:.1f},{y - half_h:.1f}\n"
                f"{x + half_w:.1f},{y + half_h:.1f}\n")
    subprocess.Popen([acad, drawing, "/b", scr])
    log(f"AutoCAD: {os.path.basename(drawing)} -> зум на {x:.0f},{y:.0f}")

# ------------------------------------------------- обновление полей через Word

def _ps_quote(s):
    """Строка как литерал PowerShell в одинарных кавычках (без раскрытия $ и `)."""
    return "'" + (s or "").replace("'", "''") + "'"


def update_fields_word(paths, log=print, pdf=False, pdf_a=None):
    """Открывает Word в фоне (COM), обновляет поля; опционально сохраняет PDF рядом.

    pdf_a: выгружать PDF в формате PDF/A (по умолчанию да, см. config.json,
    ключ "pdf_a"). Это единственный способ заставить Word встроить в PDF все
    шрифты: при обычном экспорте Arial остаётся ссылкой, и на машине без него
    (типография, другой ПК) вёрстка плывёт. Проверено на перечне из 11 страниц:
    обычный экспорт — 3 начертания не встроены, PDF/A — встроены все,
    и файл при этом легче.
    """
    if pdf_a is None:
        pdf_a = bool(load_config().get("pdf_a", True))
    # 14-й позиционный аргумент ExportAsFixedFormat — UseISO19005_1 (PDF/A)
    export = ("$d.ExportAsFixedFormat({pdf}, 17, $false, 0, 0, 0, 0, 0, "
              "$true, $true, 0, $false, $false, $true)") if pdf_a else \
             "$d.ExportAsFixedFormat({pdf}, 17)"
    ps_tpl = r'''
$ErrorActionPreference = "Stop"
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
$d = $null
try {{
  $d = $w.Documents.Open({path}, $false, $false)
  $d.Repaginate()
  $null = $d.Fields.Update()
  foreach ($rng in $d.StoryRanges) {{ $null = $rng.Fields.Update() }}
  $d.EmbedTrueTypeFonts = $true
  $d.SaveSubsetFonts = $true
  $d.Save()
  if ({pdf} -ne '') {{ ''' + export + r''' }}
}} finally {{
  # Закрываем документ и приложение, что бы ни случилось на любом шаге.
  # Раньше стоял голый Quit(): на документе с несохранёнными правками он
  # ждал ответа в диалоге «сохранить?», а окна нет — невидимый Word
  # оставался в памяти и держал файл. Следующая сборка падала с «отказано
  # в доступе», и причину на экране было не видно.
  if ($d -ne $null) {{ try {{ $d.Close(0) }} catch {{}} }}
  try {{ $w.Quit(0) }} catch {{}}
}}
'''
    ok = True
    for p in paths:
        # Word открывает документ только по обычному пути Windows: получив путь
        # с прямыми слэшами, он ищет файл относительно своей рабочей папки и
        # отвечает «файл не найден». os.path.exists прямые слэши принимает, так
        # что до Word ошибка не всплывала и выглядела необъяснимой.
        p = os.path.abspath(p)
        if not os.path.exists(p):
            ok = False
            log(f"  файл не найден: {p}")
            continue
        # Проверяем доступ к файлу ДО запуска Word. Открытый в Word документ
        # он берёт «только для чтения», спотыкается на сохранении и остаётся
        # висеть в памяти — а сообщение об этом тонет в тексте ошибки COM.
        try:
            io.open(p, "r+b").close()
        except OSError:
            ok = False
            log(f"  файл занят другой программой: {os.path.basename(p)}")
            log("      закройте его в Word и повторите")
            continue
        pdf_path = (os.path.splitext(p)[0] + ".pdf") if pdf else ""
        scr = ""
        try:
            # Путь подставляем как строку PowerShell в ОДИНАРНЫХ кавычках:
            # в них не раскрываются ни $переменные, ни обратный апостроф.
            # Внутренние апострофы удваиваются — это правило PowerShell.
            body = ps_tpl.format(path=_ps_quote(p), pdf=_ps_quote(pdf_path))
            fd, scr = tempfile.mkstemp(suffix=".ps1", prefix="perechni_")
            with os.fdopen(fd, "w", encoding="utf-8-sig") as f:
                f.write(body)
            r = subprocess.run(
                ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", scr],
                capture_output=True, timeout=300,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            if r.returncode == 0:
                log(f"  поля обновлены: {os.path.basename(p)}" + (" (+PDF)" if pdf else ""))
                if pdf and not os.path.exists(pdf_path):
                    ok = False
                    log(f"  ⚠ Word отработал, но PDF не появился: {os.path.basename(pdf_path)}")
            else:
                ok = False
                err = (r.stderr or b"").decode("cp866", "replace").strip() \
                      or (r.stdout or b"").decode("cp866", "replace").strip()
                log(f"  не удалось обновить поля: {os.path.basename(p)}")
                for line in [l for l in err.splitlines() if l.strip()][:4]:
                    log(f"      {line[:150]}")
                log("      подсказка: закройте файл в Word, если он открыт, и повторите")
        except subprocess.TimeoutExpired:
            ok = False
            log(f"  Word не ответил за 5 минут: {os.path.basename(p)} — возможно, ждёт ответа в диалоге")
        except FileNotFoundError:
            ok = False
            log("  не найден powershell.exe — обновление полей через Word недоступно")
        except Exception as e:
            ok = False
            log(f"  ошибка при обращении к Word ({type(e).__name__}): {e}")
        finally:
            if scr:
                try:
                    os.remove(scr)
                except OSError:
                    pass
    return ok

# ---------------------------------------------------------------- извлечение

POS = re.compile(r"поз\.\s*([^.]+?)\.")
NCH = {"AI": 4, "AO": 4, "WI": 1}   # запасное число каналов, если в заголовке нет


def mod_channels(text, typ):
    """Число каналов модуля из его заголовка или None.

    Количество написано рядом с типом: «AI8x4...20mA», «8AIx4...20mA»,
    «16DIх24 VDC». Значения по умолчанию (NCH) верны не для всех модулей, а
    из заголовка число точное. Без него хвостовые резервные каналы модуля
    просто не создавались: программа делала столько строк, сколько нашла
    описаний, и в перечне не хватало резерва в конце каждого модуля.
    """
    if not typ:
        return None
    t = re.escape(typ)
    # цифру из номера модуля («Модуль 4.2 AO») за число каналов не принимаем
    m = re.search(t + r"\s*(\d+)", text) or re.search(
        r"(?<![\d.])(\d+)\s*" + t, text)
    if not m:
        return None
    try:
        n = int(m.group(1))
    except ValueError:
        return None
    return n if 1 <= n <= 64 else None

# Типы модулей, у которых есть каналы ввода-вывода. Заголовок вида
# «Модуль 1.0 CPU» — это процессорный модуль, каналов у него нет, и
# раньше разбор придумывал ему восемь строк «Резерв»: в перечень они
# не попадали, но портили сводку свободных каналов и сверку.
SIGNAL_TYPES = ("AI", "AO", "DI", "DO", "WI")

# Позиция прибора в тексте описания. Разделителем внутри обозначения
# бывает и точка («HV23.3-6»), а хвост — буквенным («ZA7/11-1o»):
# прежний шаблон обрывался на точке и терял хвост, из-за чего в перечень
# уходило «HV23» вместо «HV23.3-6». Точка засчитывается только внутри —
# точка в конце предложения обозначением не становится.
TAG_FALLBACK = re.compile(
    r"\b[A-Z]{2,4}[0-9]*(?:[-/.][A-Za-z0-9]{1,6})*(?:\(\w+\))?(?<![-/.])")

# --- как программа опознаёт элементы чертежа --------------------------------
# Значения по умолчанию соответствуют чертежам, на которых программа
# отлаживалась. Любое из них можно переопределить в config.json, раздел
# "extract" — тогда чертежи с другими слоями не придётся переделывать.
# Пример config.json:
#   { "extract": { "layer_desc": "Описание", "layer_biz": "искробезопасность" } }
# Номер модуля может нести чертёжную приставку: там, где РСУ и ПАЗ нарисованы
# в одном комплекте, модули ПАЗ помечены буквой — «Модуль z1.5 DO», якорь
# «z1.5-KLDO3». В позицию по контроллеру приставка не идёт: в выпущенных
# перечнях там «1.5.1». Поэтому букву распознаём, но не запоминаем — иначе
# весь шкаф молча выпадает из разбора.
MOD_PREFIX = r"[A-Za-zА-Яа-яЁё]?"

EXTRACT_DEFAULTS = {
    "layer_desc":    "Текст",             # слой с описаниями каналов
    "layer_desc_extra": (),               # доп. слои описаний (обычно находятся сами)
    "kc_format": "{mod}.{ch}",            # как записывать позицию в контроллере:
                                          # «1.0.1» либо «{mod} {type}{ch}» -> «1.0 AI1»
    "max_tag_wires": 6,                   # сколько жил может подписать один прибор
                                          # (чаще встречается — это не прибор)
    "module_prefix": "drop",              # чертёжная приставка у номера модуля:
                                          # "drop" — отбросить, "keep" — оставить
    "layer_biz":     "mark2",             # слой с метками барьеров искрозащиты
    "biz_marker":    "БИЗ",               # текст метки БИЗ на этом слое
    "layer_di":      "Реле DI (1-KL)",    # слой якорей дискретных входов
    "layer_do":      "Реле DO (3-KL)",    # слой якорей дискретных выходов
    # Заголовок модуля пишут и коротко («Модуль 1.1 AI 8AIx4...20mA»), и
    # развёрнуто («Модуль ввода дискретных сигналов 6.1 DI, 6ES7131-...»).
    # Прежний шаблон знал только короткую форму, и на чертежах, где есть
    # лишь развёрнутая, не находилось НИ ОДНОГО модуля — при том что подписи
    # каналов на листе были. Слова между «Модуль» и номером пропускаем.
    "module_header": (r"^Модуль\s+(?:[А-Яа-яЁё\s]+\s)?" + MOD_PREFIX
                      + r"(\d+)\.(\d+)\s*,?\s*([A-ZА-Я]+)"),
    "band_left":     30.0,    # насколько левее заголовка модуля берём описания
    "band_width":    230.0,   # ширина полосы модуля, если он последний
    "di_span":       460.0,   # вправо от якоря ищем колонку описаний
    "di_col_tol":    10.0,    # разброс X внутри одной колонки описаний
    "di_row_tol":    9.0,     # разброс Y при сопоставлении описания и якоря
    "level_analog":  "4-20 мА",
    "level_discrete": "=24В",
    "ctrl_analog":   "",       # графа «уровень управления» у аналоговых:
                              # в части проектов туда пишут «РСУ»/«ПАЗ»
    "ctrl_discrete": "СК НО",
    "ex_with_biz":   "Exia",
    "ex_without_biz": "Exd",
    "reserve_word":  "Резерв",
}


def extract_settings(log=None):
    """Настройки распознавания: значения по умолчанию + правки из config.json."""
    cfg = dict(EXTRACT_DEFAULTS)
    user = load_config().get("extract") or {}
    unknown = [k for k in user if k not in EXTRACT_DEFAULTS]
    if unknown and log:
        log(f"  ⚠ в config.json незнакомые ключи extract: {', '.join(unknown)}")
        log(f"    допустимые: {', '.join(sorted(EXTRACT_DEFAULTS))}")
    for k, v in user.items():
        if k in cfg:
            cfg[k] = v
    return cfg

def _in_brackets(s, i):
    """Стоит ли место i внутри скобок.

    Позицию в скобках («…вентилятора В1 (поз. В5 из РСУ)») пишут как сноску:
    откуда приходит сигнал. Своя позиция канала пишется без скобок, поэтому
    скобочную в графу «позиция по проекту» брать нельзя, а из описания
    выбрасывать — наоборот, нужно оставить читателю.
    """
    return s.count("(", 0, i) > s.count(")", 0, i)


def _tag(s):
    # «поз. X.» — но в описании таких оборотов может быть несколько:
    # «...из цистерны поз. 12 поз.PY-4» ловилось как «12 поз», потому что
    # разбор обрывался на точке внутри второго «поз.». Берём первый захват,
    # похожий на позицию прибора, а огрызки вроде «12 поз» пропускаем.
    for m in POS.finditer(s):
        t = m.group(1).strip().rstrip(",;")
        if not t or t.lower().rstrip(".").endswith("поз"):
            continue
        if _in_brackets(s, m.start()):
            continue                      # сноска «(поз. В5 из РСУ)»
        # «поз. В5 из РСУ» — это ссылка на чужую позицию: сигнал приходит
        # оттуда, а собственная позиция канала пишется на жиле. Такую сноску
        # в графу «позиция по проекту» ставить нельзя.
        w = t.split()
        if len(w) > 1 and w[1].lower().startswith("из"):
            continue
        # Позиция — одно слово.
        t = w[0].rstrip(",;")
        if any(ch.isalpha() for ch in t) and any(ch.isdigit() for ch in t):
            return t
    # описания без «поз.»: берём первый латинский тег с цифрой (YP9, TIT1, PIT-1(23))
    for m in TAG_FALLBACK.finditer(s):
        t = m.group(0)
        if any(c.isdigit() for c in t):
            return t
    return ""

# Кандидат в позицию прибора: латинские буквы + число, иногда с хвостом
# «-1» или «/2» и с суффиксом «+»/«-» (маркировка жил): «PI 3211 +», «TI-2213».
# Технологические позиции аппаратов («поз. 54б/1») сюда не попадают — там
# после цифр идёт кириллица, а шаблон требует латиницу в начале и только её.
TAG_NEAR = re.compile(r"^([A-Za-z]{1,4}[\s-]?\d{1,5}(?:[-/][A-Za-z0-9]{1,4})?"
                      r"(?:\(\w+\))?)\s*[+\-]?$")

# Обозначения, которые выглядят как позиция, но ею не являются: типы сигналов
# модуля (AO-1, DI-12) и клеммники (XT3). Рядом с описанием их полно, и без
# этого списка «Регулирование давления…» получало позицию «AO-1».
# Обозначения, похожие на позицию, но ею не являющиеся: виды сигналов модуля
# («AO-1»), клеммники («XT3») и номера кабелей («КС3.25/2», «KC101» — буквы
# кириллические или латинские вперемешку).
# Сюда же служебные надписи самого модуля: якорь канала «z1.21-KLDO2» и
# клеммник «z1.8-XTAI». Они стоят у каждого канала и жилами прибора не
# являются — без этого якорь дискретного канала уходил в графу «позиция».
NOT_A_TAG = re.compile(
    r"^(?:(?:AI|AO|DI|DO|WI|XT|XS|XP)[\s-]?\d"
    r"|[КKKК][СCCС][\s.-]?\d"
    r"|[A-Za-zА-Яа-я]?\d+\.\d+-(?:KLD[IO]|XT|XS|XP))", re.I)


def _guess_tags(rows, texts, log=None, reserve="Резерв", desc_layer="Текст"):
    """Дописывает позицию каналам, у которых её нет в описании.

    Во многих чертежах позиция прибора не пишется в описании, а стоит рядом
    подписью жилы («PI 3211 +» и «PI 3211 -»). Такие подписи лежат на одной
    высоте с описанием, поэтому берём только попавшие в узкую полосу по
    вертикали — уже половины шага строк. Если в полосу попало несколько разных
    позиций, не угадываем и оставляем пусто.

    Возвращает число заполненных каналов.
    """
    need = [r for r in rows if not r["tag"] and r["desc"] != reserve]
    if not need:
        return 0
    cands = []
    for x, y, lay, s in texts:
        m = TAG_NEAR.match(s or "")
        if m:
            tag = re.sub(r"\s+", " ", m.group(1)).strip()
            if NOT_A_TAG.match(tag):
                continue
            cands.append((x, y, lay, tag))
    if not cands:
        return 0
    # Слой описаний — подсказка, а не фильтр: в части чертежей позиция подписана
    # на другом слое, и жёсткий отбор по слою терял её вовсе. Даём небольшую
    # фору «своему» слою, чтобы при равном удалении выигрывал он.
    LAYER_BONUS = 8.0
    # шаг строк — по расстояниям между соседними каналами одного модуля
    bymod = collections.defaultdict(list)
    for r in rows:
        bymod[r["mod"]].append(r["y"])
    gaps = []
    for ys in bymod.values():
        ys = sorted({round(v, 2) for v in ys}, reverse=True)
        gaps += [a - b for a, b in zip(ys, ys[1:]) if 0.5 < a - b < 200]
    pitch = statistics.median(gaps) if gaps else 10.0
    ytol = max(1.5, pitch * 0.45)
    xtol = max(150.0, pitch * 12)
    filled = ambiguous = 0
    for r in need:
        best = {}
        for cx, cy, lay, tag in cands:
            if abs(cy - r["y"]) <= ytol:
                dx = abs(cx - r["x"])
                if dx > xtol:
                    continue
                score = dx - (LAYER_BONUS if lay == desc_layer else 0.0)
                if tag not in best or score < best[tag]:
                    best[tag] = score
        if not best:
            continue
        order = sorted(best.items(), key=lambda kv: kv[1])
        # Настоящая позиция подписана ближе всех по горизонтали; дальше идут
        # метки клемм и соседних цепей. Берём ближайшую, но только если она
        # заметно ближе следующей — иначе не гадаем.
        if len(order) == 1 or order[1][1] >= max(order[0][1] * 1.4, order[0][1] + 20):
            r["tag"] = order[0][0]
            filled += 1
        else:
            ambiguous += 1
    if log and (filled or ambiguous):
        log(f"  позиции восстановлены по подписям рядом: {filled} из {len(need)}"
            f" (полоса ±{ytol:.1f} при шаге строк {pitch:.1f})")
        if ambiguous:
            log(f"      неоднозначных, оставлены пустыми: {ambiguous}")
    return filled


def _clean(s):
    s = re.sub(r"\s+", " ", s.replace("\n", " ")).strip()
    s = s.replace("авайриного", "аварийного").replace("авариного", "аварийного")
    # Знак номера в ГОСТ-овских SHX. В основной надписи его рисуют латинской
    # «N», а в тексте описаний — решёткой: «в смесителе #1». В перечень должно
    # уйти «№1». Меняем только перед номером: решётка в другом месте — своя.
    s = re.sub(r"#(?=\s?\d)", "№", s)
    # Позицию из описания убираем — для неё в перечне своя графа. Но сноску
    # в скобках («(поз. В5 из РСУ)») оставляем: она говорит, откуда приходит
    # сигнал, и без неё описание рвалось на середине фразы.
    for m in re.finditer(r"\s*поз\.\s*[^.]*?\.\s*", s):
        if not _in_brackets(s, m.start()):
            s = s[:m.start()] + " " + s[m.end():]
            break
    s = re.sub(r"\s{2,}", " ", s).strip()
    # сокращения применяем здесь, чтобы дальше по цепочке — предпросмотр,
    # Word, Excel, сверка — везде был один и тот же текст
    return apply_abbrev(s)

ANCHOR_RE = re.compile(MOD_PREFIX + r"(\d+)\.(\d+)-KLD([IO])(\d+)$")


# Подпись канала на схеме подключения: «3.2.1+», «z2.2.15-», иногда без знака.
# Она несёт полную позицию по контроллеру и точную координату, поэтому даёт
# раскладку описаний по каналам напрямую — вместо догадки «первое описание
# в полосе модуля принадлежит первому каналу». На схемах, где такие подписи
# есть, это разница между 43 % и 97 % найденных сигналов.
CH_LABEL_RE = re.compile(r"^([A-Za-zА-Яа-я]?)(\d+)\.(\d+)\.(\d+)\s*[+-]?$")
# Второй вид разметки канала: «1.0 AI1+» — номер модуля, тип сигнала и
# номер канала. Встречается в проектах, где и в перечне позиция записана
# так же («1.0 AI1»). Без него на таком чертеже не находилось ни одного
# канала, хотя разметка на листе есть.
CH_LABEL_RE2 = re.compile(
    r"^([A-Za-zА-Яа-я]?)(\d+)\.(\d+)\s+[A-Z]{1,3}(\d+)\s*[+-]?$")


def ch_label(txt):
    """Подпись канала в любом из принятых видов: «4.7.1+» или «1.0 AI1+»."""
    return CH_LABEL_RE.match(txt) or CH_LABEL_RE2.match(txt)


def _row_offset(col, ys, pitch, tol):
    """На сколько строка описаний смещена относительно подписей каналов.

    Медиана «до ближайшей подписи» здесь не работает: описание стоит на строке
    первой жилы канала, а это больше половины шага, и ближайшей оказывается
    подпись СЛЕДУЮЩЕГО канала — вся раскладка съезжает ровно на один канал.
    Поэтому смещение подбираем: проверяем все наблюдаемые разности и берём ту,
    при которой описания ложатся на наибольшее число РАЗНЫХ каналов. Ничего
    не подбирается на глаз и не зашито в код — величина берётся из чертежа.
    """
    if not col or not ys:
        return 0.0
    cands = {0.0}
    for _dx, dy, _s in col:
        for y in ys:
            d = dy - y
            if abs(d) <= pitch:
                cands.add(round(d, 1))
    best, best_n = 0.0, -1
    for off in sorted(cands):
        used = set()
        for _dx, dy, _s in col:
            y = min(ys, key=lambda yy: abs((dy - off) - yy))
            if abs((dy - off) - y) <= tol:
                used.add(y)
        # при равенстве предпочитаем меньшее смещение: оно проще и устойчивее
        if len(used) > best_n:
            best, best_n = off, len(used)
    return best


# Маркер жилы в конце подписи. Записывают по-разному, поэтому вариантов
# несколько; какой верен для этого чертежа, решает подсчёт жил на прибор.
WIRE_MARKS = (
    re.compile(r"^(.+?)[A-Za-zА-Яа-я]?[+\-]$"),          # PT21/1+, TT23/4-1i-
    re.compile(r"^(.+?)-[A-Za-zА-Яа-я]\.?\d{1,2}$"),      # TCV617/1-U.3
    re.compile(r"^(.+?)-\d{1,2}$"),                       # TE-1-1, TE-1-2
    re.compile(r"^(.+?)\.\d{1,2}$"),                      # Y-Н1302-3-0.1
    # Жилу называют буквой через точку: у четырёхпроводного прибора это пара
    # питания и пара токовой петли — «AZ-501/2.U+», «AZ-501/2.I-». Без этого
    # у газоанализаторов не опознавалась ни одна жила, и каналы разъезжались
    # по строкам: описание доставалось соседнему каналу, а свой шёл резервом.
    re.compile(r"^(.+?)\.[A-Za-zА-Яа-я]{1,2}[+\-]?$"),    # AZ-501/2.U+
)
# Обозначение прибора: буквы, дальше цифры, внутри «/» и «-».
# Обозначение прибора: буквы, иногда через дефис обозначение аппарата
# буквой («TZ-Н2301-1», «P-Р1201-2»), затем цифры и хвостовые номера.
WIRE_BASE_RE = re.compile(
    r"^[A-Za-zА-Яа-я]{1,6}(?:-?[A-Za-zА-Яа-я]{1,4})?-?\d{1,5}"
    r"(?:[-/.,]\d{1,3})*"                   # «303/2», «П1.1», «П1,1.1»
    r"(?:-[A-Za-zА-Яа-я0-9]{1,6})*$")           # «-POW», «-K», «-ТЕМР», «-zOFF1»


def _lat_head(s):
    """Латиница в начальных буквах обозначения; остальное не трогаем.

    Позиции приборов набирают латиницей, а обозначения аппаратов кириллицей.
    В чертежах буквы прибора часто набраны кириллическими двойниками («РT2»
    вместо «PT2»), поэтому начало переводим — но только его.
    """
    m = re.match(r"^[A-Za-zА-Яа-я]+", s or "")
    if not m:
        return s
    return m.group(0).translate(_CYR2LAT) + s[m.end():]


def _wire_bases(s):
    """Возможные обозначения прибора для подписи жилы: сама подпись и то,
    что остаётся после отбрасывания маркера жилы."""
    out = []
    if WIRE_BASE_RE.match(s):
        out.append(s)
    for rx in WIRE_MARKS:
        m = rx.match(s)
        if m and WIRE_BASE_RE.match(m.group(1)):
            out.append(m.group(1))
    return out


# Тип прибора у условного обозначения: «TT», «PT», «FT» — только буквы.
DEV_TYPE_RE = re.compile(r"^[A-Za-zА-Яа-я]{2,4}$")
# Обозначение аппарата: буква, три-четыре цифры, иногда через дефис и с
# номером — «Н-2301», «Р1201-2», «Т1451». В перечне дефис после буквы не
# пишут, поэтому его убираем.
# Хвостовых номеров бывает два: «Н-1012/1-2» — второй датчик на насосе
# Н-1012/1. Раньше брался только первый, и оба датчика одного насоса
# получали одну позицию.
EQUIP_RE = re.compile(r"\b([А-ЯA-Z])-?(\d{3,4})((?:[-/]\d{1,2})*)\b")


def _same_head(a, b):
    """Начинается ли обозначение a с обозначения b, если не считать разницы
    кириллицы и латиницы: «P2401-3» и «Р2401» — один аппарат."""
    return (a.translate(_CYR2LAT).upper()
            .startswith(b.translate(_CYR2LAT).upper()))


def equip_from(text):
    """Обозначение аппарата из текста: «насоса Н-2301» -> «Н2301»."""
    m = EQUIP_RE.search(text or "")
    return (m.group(1) + m.group(2) + m.group(3)) if m else ""


# Измеряемая величина по ГОСТ 21.208 — по слову в описании сигнала.
# Нужна там, где прибор на листе не подписан: дискретный сигнал идёт через
# промежуточное реле, и условного обозначения рядом с каналом не рисуют.
MEASURED = (
    ("температур", "T"), ("давлен", "P"), ("разрежен", "P"),
    ("расход", "F"), ("количеств", "F"),
    ("уровн", "L"), ("уровен", "L"),
    ("загазован", "Q"), ("концентрац", "Q"), ("состав", "Q"),
    ("влажн", "M"), ("плотност", "D"), ("вязкост", "V"),
    ("масс", "W"), ("вес", "W"),
    ("оборот", "S"), ("скорост", "S"), ("частот", "S"),
    ("напряжен", "E"), ("ток", "I"), ("мощност", "J"),
    ("положен", "G"), ("наличи", "G"),
)


def measured_letter(desc):
    """Буква измеряемой величины из описания сигнала или ''.

    Берём ту, чьё слово встретилось в описании раньше: «Контроль расхода воды»
    это расход, а не вода.
    """
    low = (desc or "").lower()
    best, at = "", len(low) + 1
    for stem, letter in MEASURED:
        i = low.find(stem)
        if 0 <= i < at:
            best, at = letter, i
    return best


def device_type(T, x0, right, ylo, yhi):
    """Тип прибора буквами в полосе канала: «TT», «PT». Пусто, если нет."""
    best, bd = "", 1e18
    for x, y, _l, s in T:
        if x <= x0 or x > right or not (ylo <= y <= yhi):
            continue
        if not DEV_TYPE_RE.match(s):
            continue
        if x - x0 < bd:
            best, bd = s.translate(_CYR2LAT).upper(), x - x0
    return best


# Хвост подписи жилы: номер жилы, полярность, вид сигнала, назначение.
# Номер жилы пишут через точку: «PCV303/2-U.1». Дефис сюда не берём —
# «TT23/4-1» это обозначение прибора целиком, а не прибор с номером жилы.
_WIRE_CORE = re.compile(r"\.\d{1,2}$")
# Жилу называют и буквой через точку: «AZ-501/2.U+» — пара питания, «.I+» —
# пара токовой петли. Это тоже номер жилы, а не часть обозначения прибора.
_WIRE_DOT = re.compile(r"\.[A-Za-zА-Яа-я]{1,2}$")
# Полярность: знак и, если есть, буква перед ним — «PT21/1-», «TT23/4-1i+».
# Признак жилы пишут строчной буквой («TT23/4-1i+»), а вид сигнала —
# прописными («F-Т1201-AI+»). По регистру их и различаем: иначе от «AI»
# отрывалась «I» и обозначение сигнала принималось за позицию прибора.
_WIRE_SIGN = re.compile(r"[a-zа-я]?[+\-]$")
_WIRE_KIND = re.compile(r"-(?:AI|AO|DI|DO|WI)$", re.I)   # вид сигнала
# Назначение жилы: «-U» управление, «-K» контроль положения. Это не часть
# позиции прибора: у клапана она одна, а жил к нему несколько.
_WIRE_CTRL = re.compile(r"-[UKУК]$", re.I)
# Состояние арматуры: «открыто» на чертежах пишут нулём, «закрыто» —
# кириллической «С». Это тоже назначение жилы, а не отдельный прибор:
# у клапана одна позиция, а положений два.
_WIRE_STATE = frozenset("0ОOСC")
# Позиция прибора или оборудования:
#   «PCV303/2», «LCV301»     — прибор;
#   «H1901-POW», «П2а-R»     — оборудование и вид сигнала от МСС;
#   «Y-Н1302-3»              — арматура на насосе Н1302.
_TAG_LOOKS_LIKE = re.compile(
    r"^[A-Za-zА-Яа-я]{1,6}-?\d|^[A-Za-zА-Яа-я]{1,6}-[A-Za-zА-Яа-я]{1,4}-?\d")
# Слишком короткое обозначение прибором не бывает: «A1», «A2» — это выводы
# реле, они подписаны по всему листу сотнями раз.
_TAG_TOO_SHORT = re.compile(r"^[A-Za-zА-Яа-я]\d{1,2}$")


# Буквы измеряемых величин по ГОСТ 21.208 — те, что стоят первыми в
# обозначении сигнала: «T-Р2301-1» это температура реактора Р2301.
_MEASURED_LETTERS = frozenset(l for _stem, l in MEASURED)


def wire_tag(s):
    """Позиция прибора из подписи жилы или '' — если это не позиция.

    Отбрасываем всё, что относится к жиле, а не к прибору: номер жилы,
    полярность, назначение («-U» управление, «-K» контроль положения) и
    состояние арматуры («-0» открыто, «-С» закрыто). У клапана позиция одна,
    а жил к нему несколько — в перечне она и должна быть одна.
    """
    t = (s or "").strip()
    t = _WIRE_SIGN.sub("", t)
    m = re.search(r"[.\-]([0-9A-Za-zА-Яа-я])$", t)
    if m and _WIRE_CORE.search(t):
        t = _WIRE_CORE.sub("", t)          # «.1» — номер жилы
    elif _WIRE_DOT.search(t) and _TAG_LOOKS_LIKE.match(_WIRE_DOT.sub("", t)):
        t = _WIRE_DOT.sub("", t)           # «.U», «.I» — жила питания и петли
    # состояние арматуры перед номером жилы: «-0», «-С»
    m = re.search(r"-([0-9A-Za-zА-Яа-я])$", t)
    if m and m.group(1) in _WIRE_STATE and not t[:m.start()].endswith("-"):
        t = t[:m.start()]
    if _WIRE_KIND.search(t):
        # «F-Т1201-AI» — подпись обозначает ВИД СИГНАЛА, а не прибор: позицию
        # для такого канала собираем из букв типа прибора и обозначения аппарата
        return ""
    if _WIRE_CTRL.search(t):
        t = _WIRE_CTRL.sub("", t)          # назначение жилы, не часть позиции
    t = _lat_head(t)
    if not _TAG_LOOKS_LIKE.match(t):
        return ""                          # обозначение сигнала, не прибора
    out = t
    # Однобуквенный признак сигнала на конце («-Р» пуск, «-S» стоп, «-А»
    # авария) набирают то кириллицей, то латиницей. На вид «K1450-Р» и
    # «K1450-P» не отличить, а в перечне это две разные позиции — приводим
    # к латинице, как и начальные буквы обозначения.
    out = re.sub(r"-([А-Яа-я])$",
                 lambda m: "-" + m.group(1).translate(_CYR2LAT), out)
    if NOT_A_TAG.match(out) or _TAG_TOO_SHORT.match(out):
        return ""                          # кабель, клеммник, вывод реле
    # «T-Р2301-1», «P-Ф1401-1» — одиночная буква измеряемой величины и
    # обозначение аппарата: это ОБОЗНАЧЕНИЕ СИГНАЛА, а не позиция прибора.
    # Позицию для такого канала собирают из букв прибора и аппарата —
    # «TT-Р2301». А «Y-Н1302-3» позиция: Y это исполнительный механизм,
    # измеряемой величины он не обозначает.
    m = re.match(r"^([A-Za-zА-Яа-я])-[A-Za-zА-Яа-я]*\d", out)
    if m and m.group(1).upper() in _MEASURED_LETTERS:
        return ""
    return out


def wire_groups(T, max_wires=6):
    """Жилы, сгруппированные по обозначению прибора: {обозначение: [(x, y)]}.

    Каждая подпись даёт несколько кандидатов в обозначение — с маркером жилы и
    без него. Выбираем тот, у которого жил 2-4: столько их у прибора. Если ни
    один вариант в эту вилку не попал, берём самый частый — строку он отмечает
    верно, даже если в графу «позиция» не годится.
    """
    cand = []
    for x, y, _l, s in T:
        if ch_label(s):
            continue
        bases = _wire_bases(_lat_head(s))
        if bases:
            cand.append((x, y, s, bases))
    n = collections.Counter(b for _x, _y, _s, bs in cand for b in bs)
    out = collections.defaultdict(list)
    for x, y, s, bases in cand:
        best = None
        for b in bases:
            if 2 <= n[b] <= max_wires:
                if best is None or n[b] > n[best]:
                    best = b
        if best is None:
            best = max(bases, key=lambda b: n[b])
        # третьим значением храним исходную подпись: по ней разбирается позиция
        out[best].append((x, y, s))
    return out


def channel_wire(groups, x0, right, ylo, yhi, half):
    """Прибор канала и строка его верхней жилы: (обозначение, y) или None.

    Жил у прибора две, три или четыре, поэтому пару не ищем: в полосе канала
    берём обозначение с наибольшим числом жил, при равенстве — ближайшее к
    подписям самого канала. На строке верхней жилы стоит описание сигнала.
    """
    lo, hi = ylo - half, yhi + half
    best, best_n, best_d = None, 0, 1e18
    for tag, pts in groups.items():
        rows = [(y, s) for x, y, s in pts if x0 < x <= right and lo <= y <= hi]
        if not rows:
            continue
        d = min(0.0 if ylo <= y <= yhi else min(abs(y - ylo), abs(y - yhi))
                for y, _s in rows)
        if len(rows) > best_n or (len(rows) == best_n and d < best_d):
            top = max(rows)
            # четвёртым значением — исходная подпись верхней жилы: по ней
            # разбирается позиция прибора
            best, best_n, best_d = (tag, top[0], d, top[1]), len(rows), d
    return best


def _next_column(lab, key, ys):
    """Левая граница соседнего модуля справа — по подписям каналов.

    Считать её по порядку заголовков нельзя: на многих схемах модули стоят
    в два ряда, и «следующий по X» заголовок оказывается модулем из другого
    ряда, стоящим почти в той же вертикали. Тогда граница колонки уезжала
    ЛЕВЕЕ собственных каналов модуля, и описаний не находилось вовсе.
    Берём ближайший справа модуль, который перекрывается по высоте с нашим.
    """
    x0 = min(v[0] for v in lab[key].values())
    lo, hi_ = min(ys), max(ys)
    best = float("inf")
    for k, chans in lab.items():
        if k == key:
            continue
        xs = [v[0] for v in chans.values()]
        oy = [v[1] for v in chans.values()] + [v[2] for v in chans.values()]
        if max(oy) < lo or min(oy) > hi_:
            continue                      # другой ряд модулей, не мешает
        for x in xs:
            if x > x0:
                best = min(best, x)
    return best


# Подпись жилы у канала: обозначение прибора и признак жилы — «PT21/1+»,
# «TZT21/1i-». Обозначение — буквы, затем цифра, дальше цифры, «/» и «-»
# внутри; хвостовой знак жилы в обозначение не входит.
WIRE_TAG_RE = re.compile(
    r"^([A-Za-zА-Яа-я]{1,6}[0-9](?:[0-9/]|-(?=[0-9]))*)"
    r"(?:[a-zA-Zа-яА-Я]?[+\-])?$")
# Позиции приборов по ГОСТ 21.208 набирают латиницей, но в чертежах сплошь
# встречаются кириллические двойники: «РT2» вместо «PT2». Приводим к латинице,
# иначе один и тот же прибор в перечне двоится.
_CYR2LAT = str.maketrans("АВСЕНКМОРТХУ", "ABCEHKMOPTXY")


def wire_tags(T, max_wires=6):
    """Подписи жил: [(x, y, обозначение)] — кандидаты в позицию по проекту.

    Отсеиваем служебные обозначения. Прибор подписывает жилы своего канала —
    это две жилы, изредка четыре. А выводы реле «A1»/«A2» и клеммники «XTAI1»
    подписаны по всему листу десятками и сотнями раз: в разобранном чертеже
    «A1» встретился 304 раза. Раньше такие обозначения уходили в графу
    «позиция по проекту» — уверенно и неверно. Порог задаётся ключом
    "max_tag_wires": лучше пустая графа, чем выдуманный прибор.
    """
    cand = []
    for x, y, _l, s in T:
        if ch_label(s):
            continue
        m = WIRE_TAG_RE.match(s)
        if m:
            cand.append((x, y, m.group(1).translate(_CYR2LAT)))
    n = collections.Counter(t for _x, _y, t in cand)
    return [c for c in cand if n[c[2]] <= max_wires]


def tag_on_row(wires, lx, ly, span, ytol):
    """Обозначение прибора на строке канала: справа, на своей строке.

    Берём подпись жилы, стоящую на той же строке, что и подпись канала, —
    у аналогового входа это «минусовая» жила. Если на строке ничего нет,
    графа остаётся пустой: пустая клетка честнее выдуманной позиции.
    """
    best, bd = "", 1e18
    for wx, wy, tag in wires:
        if wx <= lx or wx > lx + span:
            continue
        dy = abs(wy - ly)
        if dy > ytol:
            continue
        d = dy * 1000 + (wx - lx)
        if d < bd:
            best, bd = tag, d
    return best


def format_kc(fmt, mod, ch, typ):
    """Позиция в контроллере в принятом на проекте виде.

    Записывают её по-разному: «1.0.1» в одних проектах и «1.0 AI1» в других —
    там номер канала пишут вместе с типом сигнала. Формат задаётся ключом
    "kc_format"; неверный шаблон не должен ронять разбор, поэтому при ошибке
    возвращаем обычный вид.
    """
    try:
        return fmt.format(mod=mod, ch=ch, type=typ)
    except Exception:
        return f"{mod}.{ch}"


# Позиция прибора в принятой проектом форме: функциональные буквы по
# ГОСТ 21.208, пробел, обозначение аппарата — «TIRSA Н2401-1», «YSA Р1401-1».
EXT_POS_RE = re.compile(r"^[A-ZА-Я]{2,6}\s+[А-ЯA-Z]-?\d{3,4}(?:[-/]\d{1,2})?$")


def cable_base(s):
    """Обозначение кабеля без маркера жилы: «Y-Р1401-1-0.1» -> «Y-Р1401-1».

    Состояние арматуры («-0» открыто, «-С» закрыто) частью обозначения не
    является, а вот «-2» в «TZ-Н2301-2» является: это второй датчик на насосе.
    """
    t = re.sub(r"[+\-]$", "", (s or "").strip())
    t = re.sub(r"\.\d{1,2}$", "", t)
    t = re.sub(r"-[0ОOСC]$", "", t)
    return t


def external_tags(paths, log=print, col=14.0, depth=260.0):
    """Позиции со схем внешних проводок: {обозначение кабеля: позиция}.

    Позиция стоит в строке таблицы, обозначения жил её кабеля — ниже в той же
    колонке. Ширина колонки и глубина поиска заданы с запасом по разобранным
    чертежам; лишнее отсеивается тем, что обозначение кабеля должно содержать
    цифры и не быть клеммником или номером кабельной связи.
    """
    out = {}
    for p in paths:
        try:
            doc = read_dxf(p)
        except Exception as e:
            log(f"  не прочитан {os.path.basename(p)}: {type(e).__name__}: {e}")
            continue
        T = []
        for e in doc.modelspace():
            t = e.dxftype()
            if t not in ("TEXT", "MTEXT"):
                continue
            try:
                s = e.plain_text() if t == "MTEXT" else e.dxf.text
            except Exception:
                s = getattr(e.dxf, "text", "")
            s = (s or "").strip()
            if not s:
                continue
            try:
                i = e.dxf.insert
                T.append((float(i[0]), float(i[1]), s))
            except Exception:
                pass
        pos = [(x, y, s) for x, y, s in T if EXT_POS_RE.match(s)]
        n0 = len(out)
        for px, py, ps in pos:
            for x, y, s in T:
                if abs(x - px) > col or not (-depth < y - py < 0):
                    continue
                b = cable_base(_lat_head(s))
                if len(b) < 4 or not any(c.isdigit() for c in b):
                    continue
                if NOT_A_TAG.match(b) or EXT_POS_RE.match(s):
                    continue
                out.setdefault(b, ps)
        log(f"  внешние проводки {os.path.basename(p)}: позиций {len(pos)}, "
            f"кабелей связано {len(out) - n0}")
    return out


def stray_labels(lab, headed, gap=40.0):
    """Модули-описки: подпись канала, попавшая в колонку чужого модуля.

    У настоящего модуля есть свой заголовок либо своя колонка на листе.
    Одиночная надпись «1.1.2+» посреди колонки модуля 3.26 — это описка
    проектировщика, и модулем её считать нельзя.
    """
    out = set()
    for key, chans in lab.items():
        # Заголовок модуля от описки не спасает: «Модуль 1.1» на листе есть,
        # а подписи каналов у него написаны как «3.26.x», и лишь одна осталась
        # «1.1.2». Решает положение на листе, а не наличие заголовка.
        x0 = min(v[0] for v in chans.values())
        ys = [v[1] for v in chans.values()]
        for other, ochans in lab.items():
            if other == key or other in out:
                continue
            oxs = [v[0] for v in ochans.values()]
            oys = [v[1] for v in ochans.values()] + [v[2] for v in ochans.values()]
            if max(oys) < min(ys) - gap or min(oys) > max(ys) + gap:
                continue
            if any(abs(x - x0) <= gap for x in oxs) and len(ochans) > len(chans):
                out.add(key)
                break
    return out


def stale_heads(heads, lab, share=0.1):
    """Заголовки-призраки: модуль подписан, а каналов у него на листе нет.

    В размеченном чертеже подписи каналов рисуют у каждого модуля — в том
    числе у полностью резервного. Заголовок без единой подписи в таком
    чертеже означает не модуль, а забытую надпись из скопированного блока:
    в схемах РСУ «Модуль 1.1 AI» стоял тремя копиями посреди колонок модулей
    3.29, 3.30 и 3.16 и давал двенадцать выдуманных строк «Резерв».

    Правило работает только там, где размечен почти весь чертёж: если модулей
    без подписей много, значит так оформлен проект, а не описка.
    """
    if not lab:
        return set()
    headed = {(g, n) for _x, g, n, _t, _h in heads}
    miss = headed - set(lab)
    if not miss or len(miss) > share * len(headed):
        return set()
    return miss


def label_channels(T):
    """Каналы по подписям: {(g, num): {ch: (x, ylo, yhi, приставка)}}.

    Канал подписан 1-3 раза («4.7.1», «4.7.1+», «4.7.1-») и занимает на листе
    полосу, а не строку. Запоминаем размах полосы и самую левую подпись:
    жилы и описание идут правее, и мерить надо от неё.
    """
    out = collections.defaultdict(dict)
    for x, y, _l, s in T:
        m = ch_label(s)
        if not m:
            continue
        g, num, ch = int(m.group(2)), int(m.group(3)), int(m.group(4))
        prev = out[(g, num)].get(ch)
        if prev is None:
            out[(g, num)][ch] = (x, y, y, m.group(1))
        else:
            out[(g, num)][ch] = (min(x, prev[0]), min(y, prev[1]),
                                 max(y, prev[2]), prev[3] or m.group(1))
    return out


def _auto_layers(T, S, is_desc, log=None):
    """Определяет слои описаний, БИЗ и якорей DI/DO по содержимому чертежа.

    Заданные в настройках слои остаются в приоритете: если на них что-то есть,
    ничего не меняем. Автоопределение включается только там, где заданный слой
    в этом чертеже пуст — чтобы чертежи из другого источника читались без
    ручной настройки.
    """
    desc_by, di_by, do_by, biz_by = (collections.Counter() for _ in range(4))
    for _x, _y, lay, txt in T:
        if is_desc(txt):
            desc_by[lay] += 1
        m = ANCHOR_RE.match(txt)
        if m:
            (di_by if m.group(3) == "I" else do_by)[lay] += 1
        if S["biz_marker"] and S["biz_marker"] in txt:
            biz_by[lay] += 1

    def pick(current, counter, role):
        if counter.get(current):
            return current                      # заданный слой работает
        # слой «0» — свалка по умолчанию, берём его только если больше нечего
        cands = [k for k, _v in counter.most_common() if k != "0"]
        if not cands:
            cands = [k for k, _v in counter.most_common()]
        if not cands:
            return current
        best = cands[0]
        if log:
            log(f"  слой {role} определён по чертежу: «{best}» "
                f"({counter[best]} шт.), в настройках был «{current}»")
        return best

    # Описания одного чертежа нередко разложены по нескольким слоям: у части
    # проектов половина подписей лежит на слое «0». Берём все слои, где есть
    # текст, похожий на описание канала: отбор делает is_desc, а лишние надписи
    # вне полос модулей всё равно не попадут в перечень.
    main = pick(S["layer_desc"], desc_by, "описаний")
    extra = sorted(k for k in desc_by if k != main and desc_by[k] >= 3)
    if extra and log:
        log("  описания есть и на слоях: " + ", ".join(
            f"«{k}» ({desc_by[k]} шт.)" for k in extra))
    return dict(
        layer_desc=main,
        layer_desc_extra=extra,
        layer_biz=pick(S["layer_biz"], biz_by, "БИЗ"),
        layer_di=pick(S["layer_di"], di_by, "якорей DI"),
        layer_do=pick(S["layer_do"], do_by, "якорей DO"),
    )


def control_level(cab, prefix, default=""):
    """Уровень управления: «РСУ», «ПАЗ» либо прежнее значение.

    Берём из имени шкафа («ШКАФ ШСК1 ПАЗ») и из чертёжной приставки контура
    у номера модуля («z1.5» — ПАЗ). Если ни там, ни там не сказано, ничего не
    выдумываем и оставляем то, что было.
    """
    if prefix:
        return "ПАЗ"
    up = (cab or "").upper()
    if "ПАЗ" in up:
        return "ПАЗ"
    if "РСУ" in up:
        return "РСУ"
    return default


def extract(dxf_path, log=print, ext_tags=None):
    """Возвращает список каналов: dict(mod,type,io,ch,tag,desc,level,ctrl,ex,kc)."""
    doc = read_dxf(dxf_path)
    msp = doc.modelspace()
    T = []
    for e in msp:
        t = e.dxftype()
        if t in ("TEXT", "MTEXT"):
            try:
                s = e.plain_text() if t == "MTEXT" else e.dxf.text
            except Exception:
                s = getattr(e.dxf, "text", "")
            i = e.dxf.insert
            T.append((float(i[0]), float(i[1]), e.dxf.layer, s.strip()))
    S = extract_settings(log)
    hdr = re.compile(S["module_header"])
    heads = []
    head_xy = collections.defaultdict(list)   # где на листе стоит заголовок
    for x, y, l, s in T:
        m = hdr.match(s)
        if m:
            heads.append((x, int(m.group(1)), int(m.group(2)), m.group(3), s))
            head_xy[(int(m.group(1)), int(m.group(2)))].append((x, y))
    heads.sort()
    def _is_desc(txt):
        if "поз." in txt:
            return True
        t = txt.strip()
        if len(t) < 20:
            return False
        if t.startswith(("Модуль", "Примечание", "Шкаф", "Питание", "Таблица",
                         "Вид ", "Обозначение", "Схем")):
            return False
        # Название листа: «Схемы ШСК 3 пн3 (DO)», «… (AIAODIDO)». В скобках
        # перечислены виды сигналов — у описания сигнала такого не бывает.
        if re.search(r"\((?:AI|AO|DI|DO|WI|TM|\+|\s)+\)\s*$", t, re.I):
            return False
        # Описание — фраза, а не обозначение: два русских слова от трёх букв
        # и хотя бы одна строчная буква. Раньше слова требовались ПОДРЯД, и
        # «Насос Н-1901. Отключение» описанием не считалось — обозначение
        # аппарата разрывало пару. Из-за этого модуль 8.29 целиком уходил в
        # резерв, хотя все шестнадцать каналов на листе подписаны.
        # Строчная буква отсекает надписи основной надписи («проект В —
        # ШСК3 ПАНЕЛЬ1»): описания сигналов набирают предложением.
        if not re.search(r"[а-яё]", t):
            return False
        return len(re.findall(r"[А-Яа-яЁё]{3,}", t)) >= 2
    if load_config().get("auto_layers", True):
        S = dict(S)
        S.update(_auto_layers(T, S, _is_desc, log))
    desc_layers = {S["layer_desc"]} | set(S.get("layer_desc_extra") or ())
    descs = [(x, y, s) for x, y, l, s in T if l in desc_layers and _is_desc(s)]
    biz = [(x, y) for x, y, l, s in T if l == S["layer_biz"] and S["biz_marker"] in s]
    if not descs and T:
        layers = collections.Counter(l for _x, _y, l, _s in T)
        log(f"  ⚠ на слое «{S['layer_desc']}» описаний не найдено.")
        log("    слои с текстом в этом чертеже: "
            + ", ".join(f"{n} ({c})" for n, c in layers.most_common(6)))
        log("    если описания лежат на другом слое — укажите его в config.json:")
        log('    { "extract": { "layer_desc": "имя вашего слоя" } }')
    # Жилы: по ним канал находит своё описание, а прибор — свою позицию.
    groups = wire_groups(T, S["max_tag_wires"])
    # Обозначения приборов. Прибор подписывает жилы своих каналов — их две,
    # три или четыре. Выводы реле «A1»/«A2» и клеммники «XTAI1» подписаны по
    # всему листу десятками и сотнями раз («A1» встретился 304 раза в одном
    # чертеже) — это не приборы, и в графу «позиция» они попадать не должны.
    # В графу «позиция по проекту» годятся только обозначения приборов:
    # клеммники («XTDI-18», «XTAI-1») и обозначения каналов модуля («AO-1»)
    # приборами не являются, а жил у них столько же.
    dev_groups = {t: p for t, p in groups.items()
                  if len(p) <= S["max_tag_wires"] and not NOT_A_TAG.match(t)}
    used_desc = set()   # какие описания легли в каналы
    # Чертёжную приставку («z» у ПАЗ) одни проекты в позицию по контроллеру
    # пишут, другие отбрасывают: в перечне проекта А стоит «z2.2.1», в
    # втором — «1.6.1» при той же подписи «z1.6.1». Угадать нельзя,
    # поэтому это настройка; по умолчанию отбрасываем, как было.
    keep_prefix = str(S.get("module_prefix", "drop")).lower() == "keep"
    rows = []
    # Ширина полосы последнего модуля. У всех остальных она задаётся položением
    # следующего заголовка, а у последнего следующего нет. Раньше подставлялась
    # константа band_width, и если реальный шаг модулей был больше, у последнего
    # модуля обрезался хвост описаний — они молча не попадали в перечень.
    # Считаем шаг по самому чертежу: он надёжнее любой константы.
    xs_head = sorted({h[0] for h in heads})
    step = None
    if len(xs_head) >= 2:
        gaps = [b - a for a, b in zip(xs_head, xs_head[1:]) if b - a > 1]
        if gaps:
            step = statistics.median(gaps)
    last_w = step if step else S["band_width"]
    if step and abs(step - S["band_width"]) > 1:
        log(f"  шаг модулей на чертеже: {step:.0f} (для последнего модуля "
            f"вместо запасных {S['band_width']:.0f})")
    # Модуль может занимать не один лист: заголовок «Модуль 1.1 AI» тогда
    # встречается дважды, а каналы на чертеже пронумерованы сквозняком —
    # 1.1.1…1.1.8 на первом листе и 1.1.9…1.1.16 на втором. Раньше нумерация
    # начиналась с единицы у каждого вхождения, и вторая половина получала
    # чужие номера: позиции в контроллере дублировались и были неверны.
    # Ведём смещение по каждому модулю в порядке следования слева направо.
    ch_offset = collections.defaultdict(int)
    skipped_types = collections.Counter()   # модули без каналов (CPU, БП и пр.)
    # Дискретные модули разбираются ниже по якорям «X.Y-KLDI/KLDO<n>», где номер
    # канала берётся из самой метки. Но заголовок «Модуль X.Y DO» попадает и в
    # heads, и аналоговая ветка делала для него второй комплект строк с
    # нумерацией с единицы — позиции задваивались. Пропускаем такие заголовки.
    anchor_mods = set()
    _apat = re.compile(MOD_PREFIX + r"(\d+)\.(\d+)-KLD[IO](\d+)$")
    for _x, _y, _l, _s in T:
        if _l in (S["layer_di"], S["layer_do"]):
            _m = _apat.match(_s)
            if _m:
                anchor_mods.add((int(_m.group(1)), int(_m.group(2))))
    # Подписи каналов на чертеже: если модуль ими размечен, раскладываем
    # описания по ним, а не по порядку следования в полосе модуля.
    lab = label_channels(T)
    lab_done = set()
    if lab:
        log(f"  подписи каналов на чертеже: модулей {len(lab)}, "
            f"каналов {sum(len(v) for v in lab.values())}")
    # Модуль, у которого каналы размечены, в перечень идёт обязательно — даже
    # если заголовка на листе нет или тип в нём незнакомый. Иначе такие модули
    # молча пропадали: «TM» (счёт импульсов) отбрасывался наравне с «TM»
    # терминально-соединительным, у которого каналов действительно нет.
    _known = {(g, n) for _x, g, n, t, _h in heads if t in SIGNAL_TYPES}
    _stray = stray_labels(lab, {(g, n) for _x, g, n, _t, _h in heads})
    lab_all = {k: dict(v) for k, v in lab.items() if k in _stray}
    for k in sorted(_stray):
        del lab[k]
    if _stray:
        log("  ⚠ подписи каналов вне своего модуля (описка на листе): "
            + ", ".join(f"{g}.{n}" for g, n in sorted(_stray)))
        for g, n in sorted(_stray):
            for c, v in sorted(lab_all.get((g, n), {}).items()):
                log(f"      {g}.{n}.{c} — x={v[0]:.0f} y={v[1]:.0f}")
    # Заголовок без единой подписи канала — забытая надпись, а не модуль.
    _stale = stale_heads(heads, lab)
    if _stale:
        heads = [h for h in heads if (h[1], h[2]) not in _stale]
        log("  ⚠ заголовок модуля без единой подписи канала "
            "(забытая надпись на листе): "
            + ", ".join(f"{g}.{n}" for g, n in sorted(_stale)))
        for g, n in sorted(_stale):
            for x, y in head_xy[(g, n)]:
                log(f"      {g}.{n} — x={x:.0f} y={y:.0f}")
    _extra = sorted(k for k in lab if k not in _known)
    for g, n in _extra:
        heads.append((min(v[0] for v in lab[(g, n)].values()), g, n, "", ""))
    if _extra:
        heads.sort()
        log("  ⚠ модули с размеченными каналами, но без опознанного типа: "
            + ", ".join(f"{g}.{n}" for g, n in _extra))
        log("    каналы взяты входными с дискретным уровнем — проверьте графы")
    # аналоговые модули (по заголовкам)
    for i, (x, g, num, typ, htext) in enumerate(heads):
        if typ in ("DI", "DO") and (g, num) in anchor_mods and (g, num) not in lab:
            continue
        if (g, num) in _stray:
            # Подписи каналов этого модуля признаны опиской на листе, значит и
            # заголовок его — тоже: иначе модуль вернётся через старую ветку и
            # даст пачку выдуманных «Резерв».
            continue
        if typ not in SIGNAL_TYPES and (g, num) not in lab:
            skipped_types[typ] += 1
            continue
        x1 = heads[i + 1][0] if i + 1 < len(heads) else x + last_w
        lo = x - S["band_left"]
        hi = x1 - S["band_left"]
        band = sorted([d for d in descs if lo <= d[0] < hi], key=lambda r: -r[1])
        has_biz = any(lo <= bx < hi for bx, by in biz)
        # Направление и параметры канала берём по типу модуля из заголовка, а
        # не по ветке разбора. Дискретные модули обычно разбираются ниже, по
        # якорям «X.Y-KLDI/KLDO», но если якорей на чертеже нет (другое
        # оформление схем), заголовки попадают сюда — и раньше дискретные
        # ВХОДЫ получали направление «выход» и уровень «4-20 мА», то есть
        # уезжали в перечень выходных сигналов с параметрами аналоговых.
        analog = typ in ("AI", "AO", "WI")
        # незнакомый тип считаем входом: датчиков на входе подавляющее
        # большинство, а неверное направление уводит сигнал в чужой перечень
        io_ = "out" if typ in ("AO", "DO") else "in"
        lvl = S["level_analog"] if analog else S["level_discrete"]
        ctl = S["ctrl_analog"] if analog else S["ctrl_discrete"]
        if (g, num) in lab:
            # Модуль на нескольких листах даёт несколько заголовков. Подписи
            # каналов уже несут сквозные номера, поэтому обрабатываем модуль
            # один раз, а повторные заголовки просто пропускаем: иначе они
            # уходили в старую ветку и добавляли дубли каналов.
            if (g, num) in lab_done:
                continue
            lab_done.add((g, num))
            ex_ = S["ex_with_biz"] if has_biz else S["ex_without_biz"]
            items = sorted(lab[(g, num)].items())
            lx0 = min(v[0] for _c, v in items)
            spans = {c: (v[1], v[2]) for c, v in items}
            ys = [(v[1] + v[2]) / 2.0 for _c, v in items]
            # Вправо ищем не дальше своей колонки модуля: окно шире шага
            # модулей утаскивало описание из СОСЕДНЕГО модуля, а вместе с
            # описанием — и позицию, извлечённую из его текста.
            right = min(lx0 + S["di_span"], _next_column(lab, (g, num), ys))
            gaps = [abs(a - b) for a, b in zip(sorted(ys), sorted(ys)[1:])
                    if abs(a - b) > 0.5]
            pitch = statistics.median(gaps) if gaps else S["di_row_tol"] * 2
            # Колонку ограничиваем и по высоте: модули часто стоят в два ряда.
            col = [d for d in descs if lx0 <= d[0] <= right
                   and min(ys) - pitch <= d[1] <= max(ys) + pitch]
            tol = max(pitch / 2.0, 1.0)

            # Строка описания у канала — строка верхней жилы его прибора.
            # Сначала ищем среди обозначений приборов, и только если их нет —
            # среди любых подписей жил: строку они отмечают верно, даже если
            # сами в графу «позиция» не годятся.
            rowof, tagof, w3of = {}, {}, {}
            for ch, (ylo_c, yhi_c) in spans.items():
                w = (channel_wire(dev_groups, lx0, right, ylo_c, yhi_c, tol)
                     or channel_wire(groups, lx0, right, ylo_c, yhi_c, tol))
                tagof[ch] = wire_tag(w[3]) if w else ""
                w3of[ch] = w[3] if w else ""
                rowof[ch] = w[1] if w else yhi_c

            # Каждое описание закрепляем за одним каналом: так одна строка не
            # достанется сразу двум каналам и не уедет к соседнему.
            # Первый проход — по строке жилы прибора: это точная привязка.
            owner, taken = {}, set()
            for i, (dx, dy, ds) in enumerate(col):
                hit, hd = None, 1e18
                for ch, ry in rowof.items():
                    d = abs(dy - ry)
                    if d <= tol and d < hd:
                        hit, hd = ch, d
                if hit is not None and (hit not in owner or hd < owner[hit][1]):
                    if hit in owner:
                        taken.discard(owner[hit][0])
                    owner[hit] = (i, hd)
                    taken.add(i)
            # Второй проход — для каналов, у которых жилы не опознались
            # (дискретные сигналы идут через промежуточные реле, и подписи там
            # другие). Опираемся на подписи самого канала. Без этого такие
            # каналы уходили в «Резерв», хотя описание на листе есть.
            for ch, (ylo_c, yhi_c) in spans.items():
                if ch in owner:
                    continue
                hit, hd = None, 1e18
                for i, (dx, dy, ds) in enumerate(col):
                    if i in taken:
                        continue
                    d = 0.0 if ylo_c <= dy <= yhi_c else min(abs(dy - ylo_c),
                                                             abs(dy - yhi_c))
                    if d <= tol and d < hd:
                        hit, hd = i, d
                if hit is not None:
                    owner[ch] = (hit, hd)
                    taken.add(hit)

            for ch, (lx, ylo_c, yhi_c, pref) in items:
                best = col[owner[ch][0]] if ch in owner else None
                if best is not None:
                    tag, desc = _tag(best[2]), _clean(best[2])
                    cx, cy = best[0], best[1]
                    used_desc.add((round(cx, 2), round(cy, 2)))
                else:
                    tag, desc = "", S["reserve_word"]
                    cx, cy = lx, rowof[ch]
                # Позиция из описания надёжнее всего («…клапана FC23/4 на…»).
                # Если прибор в описании не назван, берём обозначение с жил
                # этого же канала — но только там, где жила идёт прямо к
                # прибору. У дискретных входов между прибором и модулем стоит
                # промежуточное реле, и на жиле обозначение реле, а не прибора:
                # пустая графа честнее выдуманной позиции.
                if not tag and desc != S["reserve_word"]:
                    # Позиция прибора — с подписи жилы этого канала. Если жила
                    # подписана видом сигнала («F-Т1201-AI»), а не прибором,
                    # wire_tag вернёт пусто, и позицию собираем из букв типа
                    # прибора у условного обозначения и обозначения аппарата
                    # из описания: в перечне её пишут так же.
                    # Позиция со схемы внешних проводок — самая точная: там
                    # она написана целиком и в принятой проектом форме.
                    wire = w3of.get(ch, "")
                    base = cable_base(_lat_head(wire))
                    tag = (ext_tags or {}).get(base, "")
                    # Граница предложения — точка, пробел и заглавная буква:
                    # словарь сокращений даёт «Контр.», и по простой точке
                    # описание рвалось на первом слове.
                    first = re.split(r"\.\s+(?=[А-ЯA-Z])", desc)[0]
                    e_desc = equip_from(first)
                    # Аппарат бывает назван только на жиле: «Контр. давления в
                    # т/проводе подачи ОЭ перед расходомером» — в описании его
                    # нет, а на жиле «P-X2301-1» есть. Бывает и наоборот: на
                    # жиле указан соседний аппарат («F-Т1201-AI» при описании
                    # «реактора Р-1201») — тогда верен названный в описании.
                    # Тот же аппарат на жиле бывает написан с номером точки
                    # замера: «F-Р1201-3-AI» — третий расходомер на линии
                    # Р-1201. Номер берём с жилы, а написание — из описания:
                    # аппараты набирают кириллицей, а на жилах сплошь и рядом
                    # стоят латинские двойники («P2401» вместо «Р2401»), и от
                    # них в одном перечне получались две разные позиции.
                    e_wire = equip_from(_lat_head(wire))
                    if e_desc and e_wire and _same_head(e_wire, e_desc):
                        e = e_desc + e_wire[len(e_desc):]
                    else:
                        e = e_desc or e_wire
                    if not tag:
                        # Позиция, написанная прямо на жиле, вернее собранной:
                        # это обозначение прибора, данное проектировщиком.
                        # Собирать из букв и аппарата приходится там, где жила
                        # подписана видом сигнала («P-Р1201-1» — давление
                        # реактора), и тогда wire_tag возвращает пусто.
                        # У газоанализаторов жилы подписаны позициями датчиков
                        # («AZ-501/2», «AZ-502/2»), а описание у обоих одно и
                        # то же место — собранная позиция выходила одинаковой
                        # на два разных прибора.
                        tag = wire_tag(wire)
                    if not tag and e:
                        # буквы прибора с условного обозначения рядом с каналом
                        t = device_type(T, lx, right, ylo_c - pitch / 2.0,
                                        yhi_c + pitch / 2.0)
                        if t:
                            tag = t + "-" + e
                    if not tag and e_desc:
                        # Прибора на листе нет: сигнал идёт через промежуточное
                        # реле, условного обозначения рядом не рисуют. Буквы
                        # выводим по ГОСТ 21.208 — измеряемая величина и вид
                        # сигнала: «Контр. расхода…» даёт «FS» на дискретном
                        # канале и «FT» на аналоговом.
                        m = measured_letter(first)
                        if m:
                            tag = m + ("T" if analog else "S") + "-" + e_desc
                    if not tag and wire:
                        # Обозначение сигнала с жилы («H1301-AO») позицией
                        # прибора не является, но это то, чем канал подписан на
                        # листе, и в перечне оно полезнее пустой графы.
                        b = _lat_head(cable_base(wire))
                        if (_TAG_LOOKS_LIKE.match(b) and not NOT_A_TAG.match(b)
                                and not _TAG_TOO_SHORT.match(b)):
                            tag = b
                mod = f"{pref}{g}.{num}" if keep_prefix else f"{g}.{num}"
                rows.append(dict(sortx=x, x=cx, y=cy, mod=mod, type=typ, io=io_,
                                 ch=ch, tag=tag, desc=desc, level=lvl, ctrl=ctl,
                                 ex=ex_,
                                 kc=format_kc(S["kc_format"], mod, ch, typ)))
            continue
        n = max(mod_channels(htext, typ) or NCH.get(typ, 4), len(band))
        base = ch_offset[(g, num, typ)]
        ch_offset[(g, num, typ)] = base + n
        for ch in range(1, n + 1):
            if ch <= len(band):
                s = band[ch - 1][2]
                tag, desc = _tag(s), _clean(s)
                cx, cy = band[ch - 1][0], band[ch - 1][1]
                used_desc.add((round(cx, 2), round(cy, 2)))
            else:
                tag, desc = "", S["reserve_word"]
                cx = x
                # запасной Y: у соседей по модулю, иначе — у любого описания.
                # Раньше здесь стояло жёстко 3150 (координата с чертежа-образца),
                # и на чертеже с другой раскладкой резерв уезжал за лист.
                cy = band[0][1] if band else (descs[0][1] if descs else 0.0)
            gch = base + ch
            rows.append(dict(sortx=x, x=cx, y=cy, mod=f"{g}.{num}", type=typ, io=io_, ch=gch,
                             tag=tag, desc=desc, level=lvl, ctrl=ctl,
                             ex=S["ex_with_biz"] if has_biz else S["ex_without_biz"],
                             kc=format_kc(S["kc_format"], f"{g}.{num}",
                                          gch, typ)))
    # дискретные модули (по якорям KLDI/KLDO)
    for lay, kind, io_ in ((S["layer_di"], "DI", "in"), (S["layer_do"], "DO", "out")):
        pat = re.compile(MOD_PREFIX + r"(\d+)\.(\d+)-KLD[IO](\d+)$")
        an = [(x, y, s, pat.match(s)) for x, y, l, s in T if l == lay and pat.match(s)]
        bymod = collections.defaultdict(list)
        for a in an:
            bymod[(int(a[3].group(1)), int(a[3].group(2)))].append(a)
        for (g, num), anch in bymod.items():
            # Модуль, разложенный по подписям каналов, здесь не трогаем: иначе
            # он получал второй комплект строк и позиции задваивались.
            if (g, num) in lab_done:
                continue
            ax = statistics.median(a[0] for a in anch)
            col = [d for d in descs if ax < d[0] < ax + S["di_span"]]
            colx = min((d[0] for d in col), default=None)
            col = [d for d in col
                   if colx is not None and abs(d[0] - colx) < S["di_col_tol"]]
            used = set()
            for x, y, s, m in sorted(anch, key=lambda a: -a[1]):
                ch = int(m.group(3))
                best, bd = None, 1e9
                for j, (dx, dy, ds) in enumerate(col):
                    if j in used:
                        continue
                    d = abs(dy - y)
                    if d < bd:
                        bd, best = d, j
                if best is not None and bd < S["di_row_tol"]:
                    used.add(best)
                    ds = col[best][2]
                    used_desc.add((round(col[best][0], 2), round(col[best][1], 2)))
                    tag, desc = _tag(ds), _clean(ds)
                else:
                    tag, desc = "", S["reserve_word"]
                rows.append(dict(sortx=x, x=x, y=y, mod=f"{g}.{num}", type=kind, io=io_, ch=ch,
                                 tag=tag, desc=desc, level=S["level_discrete"],
                                 ctrl=S["ctrl_discrete"], ex=S["ex_without_biz"],
                                 kc=format_kc(S["kc_format"], f"{g}.{num}",
                                              ch, kind)))
    rows.sort(key=lambda r: (r["sortx"], r["ch"]))
    if rows and load_config().get("guess_tags", True):
        _guess_tags(rows, T, log, reserve=S["reserve_word"], desc_layer=S["layer_desc"])
    n_res = sum(1 for r in rows if r["desc"] == "Резерв")
    if not rows:
        log("  ⚠ КАНАЛЫ НЕ НАЙДЕНЫ — файл не похож на схему подключения.")
        log("    Если это «Чертёж общего вида» (габариты шкафа) — возьмите файл")
        log("    «Схема принципиальная электрическая питания и управления» этого шкафа.")
    else:
        log(f"  извлечено каналов: {len(rows)} (из них резерв: {n_res})")
        if skipped_types:
            log("  модули без каналов ввода-вывода пропущены: "
                + ", ".join(f"{t} ({n})" for t, n in skipped_types.most_common()))
        if rows and n_res > len(rows) * 0.7:
            log(f"  ⚠ резерва подозрительно много ({n_res}/{len(rows)}) — проверьте раздел на «Сигналах»")
        # Что программа увидела, но никуда не отнесла. Раньше такие описания
        # выпадали молча: в перечень они не попадали, и заметить это можно было,
        # только пересчитав позиции по чертежу вручную.
        lost = [d for d in descs
                if (round(d[0], 2), round(d[1], 2)) not in used_desc]
        if lost:
            log(f"  ⚠ описаний не попало ни в один канал: {len(lost)} из {len(descs)}")
            for x, y, t in sorted(lost, key=lambda d: -d[1])[:5]:
                log(f"      ({x:.0f}, {y:.0f})  {_clean(t)[:62]}")
            if len(lost) > 5:
                log(f"      ... и ещё {len(lost) - 5}")
            log("      обычно это подписи вне полосы модуля или описание без якоря канала")
        no_tag = [r for r in rows if not r["tag"] and r["desc"] != S["reserve_word"]]
        named = [r for r in rows if r["desc"] != S["reserve_word"]]
        if no_tag and named:
            share = len(no_tag) / len(named)
            if share > 0.5:
                # в чертеже просто не принято писать «поз.» — это не ошибка
                log(f"  позиции в описаниях почти не проставлены: {len(no_tag)} из {len(named)}")
                log("    графа «позиция по проекту» останется пустой — заполните на «Сигналах»")
            else:
                log(f"  ⚠ каналов с описанием, но без позиции: {len(no_tag)} из {len(named)}")
                for r in no_tag[:3]:
                    log(f"      {r['kc']}  {r['desc'][:58]}")
    return rows

# ---------------------------------------------------------------- сборка docx

def _ptext(p):
    return "".join(t.text or "" for t in p.iter(qn("w:t")))

def _para_set_text(p, text):
    rs = p.findall(qn("w:r"))
    if rs:
        first = rs[0]
        for t in first.findall(qn("w:t")):
            first.remove(t)
        for br in first.findall(qn("w:br")):
            first.remove(br)
        for extra in rs[1:]:
            p.remove(extra)
        t = OxmlElement("w:t"); t.set(qn("xml:space"), "preserve"); t.text = text
        first.append(t)
    else:
        r = OxmlElement("w:r")
        t = OxmlElement("w:t"); t.set(qn("xml:space"), "preserve"); t.text = text
        r.append(t); p.append(r)
    for hl in p.findall(qn("w:hyperlink")):
        p.remove(hl)

def _tc_set_text(tc, text):
    for extra in tc.findall(qn("w:p"))[1:]:
        tc.remove(extra)
    _para_set_text(tc.find(qn("w:p")), text)

def _cells(tr):
    return tr.findall(qn("w:tc"))

def _set_title_row(tbl, title):
    tr = tbl.findall(qn("w:tr"))[0]
    seen = set()
    for tc in _cells(tr):
        if id(tc) in seen:
            continue
        seen.add(id(tc)); _tc_set_text(tc, title)

def _page_break_before(p, on=True):
    """Свойство абзаца «с новой страницы».

    Надёжнее отдельного абзаца с разрывом: если предыдущая таблица закончилась
    ровно на нижней кромке листа, отдельный разрыв добавляет ещё один лист —
    и он остаётся пустым, потому что заголовок уходит на следующий. Свойство
    самого абзаца такого листа не создаёт.
    """
    pPr = p.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p.insert(0, pPr)
    have = pPr.find(qn("w:pageBreakBefore"))
    if on and have is None:
        pPr.insert(0, OxmlElement("w:pageBreakBefore"))
    elif not on and have is not None:
        pPr.remove(have)


def _drop_keepnext(el):
    """Снять со строки «не отрывать от следующего абзаца».

    В шаблоне это свойство стоит у каждого абзаца строки-образца. Скопированное
    на триста строк, оно требует от Word удержать вместе всю таблицу; выполнить
    это нельзя, и Word сбрасывает её на новый лист — а над ней остаётся лист с
    одним заголовком. У шапки свойство нужно (она не должна отрываться от
    первой строки) и остаётся на месте.
    """
    for p in el.iter(qn("w:p")):
        pPr = p.find(qn("w:pPr"))
        if pPr is None:
            continue
        kn = pPr.find(qn("w:keepNext"))
        if kn is not None:
            pPr.remove(kn)


def _build_table(proto, title, data):
    t = copy.deepcopy(proto)
    trs = t.findall(qn("w:tr"))
    proto_row = copy.deepcopy(trs[2])
    _drop_keepnext(proto_row)
    for tr in trs[2:]:
        t.remove(tr)
    _set_title_row(t, title)
    for d in data:
        tr = copy.deepcopy(proto_row)
        vals = ["", d["tag"], d["desc"], d["ctrl"], d["level"], d["ex"], d["kc"]]
        for tc, v in zip(_cells(tr), vals):
            _tc_set_text(tc, v)
        t.append(tr)
    return t

_bm = [100]

def _add_bookmark(p, name):
    _bm[0] += 1
    i = str(_bm[0])
    bs = OxmlElement("w:bookmarkStart"); bs.set(qn("w:id"), i); bs.set(qn("w:name"), name)
    be = OxmlElement("w:bookmarkEnd"); be.set(qn("w:id"), i)
    pPr = p.find(qn("w:pPr"))
    if pPr is not None:
        pPr.addnext(bs)
    else:
        p.insert(0, bs)
    p.append(be)

def _cell_pageref(tc, name):
    for extra in tc.findall(qn("w:p"))[1:]:
        tc.remove(extra)
    p = tc.find(qn("w:p"))
    for r in p.findall(qn("w:r")):
        p.remove(r)
    for fs in p.findall(qn("w:fldSimple")):
        p.remove(fs)
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), f" PAGEREF {name} \\h ")
    fld.set(qn("w:dirty"), "true")
    r = OxmlElement("w:r"); t = OxmlElement("w:t"); t.text = "0"
    r.append(t); fld.append(r)
    p.append(fld)

def _enable_update_fields(doc):
    s = doc.settings.element
    if s.find(qn("w:updateFields")) is None:
        uf = OxmlElement("w:updateFields"); uf.set(qn("w:val"), "true")
        s.insert(0, uf)

def _fix_date_font(path):
    """Даты в штампе набраны без явного шрифта -> задать ISOCPEUR."""
    import zipfile
    tmp = path + ".tmp"
    OLD = '<w:rFonts w:cs="Arial"/>'
    NEW = '<w:rFonts w:ascii="ISOCPEUR" w:hAnsi="ISOCPEUR" w:cs="Arial"/>'
    try:
        with zipfile.ZipFile(path, "r") as zin, \
             zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for it in zin.infolist():
                data = zin.read(it.filename)
                if it.filename.startswith("word/") \
                   and ("header" in it.filename or "footer" in it.filename) \
                   and it.filename.endswith(".xml"):
                    data = data.decode("utf-8").replace(OLD, NEW).encode("utf-8")
                zout.writestr(it, data)
        os.replace(tmp, path)
    except Exception:
        # исходный документ остаётся нетронутым, мусор убираем за собой
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass
        raise

_ABBREV_CACHE = None


def load_abbrev():
    """Словарь сокращений: [(что, чем)] от длинных к коротким.

    Берётся из abbrev.json рядом с программой и из раздела "abbrev" в
    config.json (второй дополняет первый). Пустой словарь = сокращения
    выключены, описания идут как в чертеже.
    """
    global _ABBREV_CACHE
    if _ABBREV_CACHE is not None:
        return _ABBREV_CACHE
    table = {}
    try:
        import json
        f = os.path.join(_app_dir(), "abbrev.json")
        if os.path.exists(f):
            with io.open(f, encoding="utf-8") as fh:
                data = json.load(fh)
            src = data.get("замены") or data.get("abbrev") or {}
            for k, v in src.items():
                if not k.startswith("_") and isinstance(v, str):
                    table[k] = v
    except Exception:
        pass
    try:
        for k, v in (load_config().get("abbrev") or {}).items():
            if isinstance(v, str):
                table[k] = v
    except Exception:
        pass
    _ABBREV_CACHE = sorted(table.items(), key=lambda kv: -len(kv[0]))
    return _ABBREV_CACHE


def apply_abbrev(text, table=None):
    """Применяет сокращения к описанию.

    Заменяет целыми словами и без учёта регистра, но сохраняет заглавную
    первую букву исходного текста: «Контроль давления…» -> «Контр. давления…».
    """
    if table is None:
        table = load_abbrev()
    if not table or not text:
        return text
    out = text
    for full, short in table:
        out = re.sub(r"(?<![А-Яа-яЁёA-Za-z])" + re.escape(full) + r"(?![А-Яа-яЁёA-Za-z])",
                     short, out, flags=re.IGNORECASE)
    if text[:1].isupper() and out[:1].islower():
        out = out[:1].upper() + out[1:]
    return re.sub(r"\s{2,}", " ", out).strip()


def _cab_title(cab):
    """«ШКАФ ШСК1 РСУ» -> «Шкаф ШСК1 РСУ».

    Обычный capitalize() опускает регистр всей строки после первой буквы, и
    обозначение шкафа превращалось в «шск1 рсу», а щит — в «2щ». Это часть
    шифра, регистр в нём значим, поэтому опускаем только ведущее слово.
    """
    s = (cab or "").strip()
    if not s:
        return s
    parts = s.split(None, 1)
    head = parts[0]
    head = head[:1].upper() + head[1:].lower()
    return head + (" " + parts[1] if len(parts) > 1 else "")


# Типы модулей, чьи каналы идут в перечень обычными графами.
_STD_TYPES = ("AI", "AO", "WI", "DI", "DO")


def _split_rows(rows, io_, analog_types):
    """Каналы направления io_, разложенные по видам сигнала.

    Импульсные («Модуль счета импульсных сигналов 8.7 TM») выделены отдельно:
    дискретными они не являются, и в общей таблице с ними теряются.
    """
    an, di, im = [], [], []
    for r in rows:
        if r["io"] != io_:
            continue
        if r["type"] in analog_types:
            an.append(r)
        elif r["type"] in _STD_TYPES:
            di.append(r)
        else:
            im.append(r)
    return an, di, im


def build_docx(template, outfile, io_, word, sections, log=print):
    """
    template  - путь к шаблону-перечню (фирменный вид);
    io_       - 'in' | 'out';
    word      - 'ВХОДНЫЕ' | 'ВЫХОДНЫЕ';
    sections  - список (имя_раздела, rows) в нужном порядке, напр. ("ШКАФ ШСК2", [...]).
    """
    analog_types = ("AI", "WI") if io_ == "in" else ("AO",)
    doc = Document(template)
    body = doc.element.body
    tables = doc.tables
    if len(tables) < 4:
        raise RuntimeError("Шаблон не похож на перечень: мало таблиц")
    toc_tbl = tables[0]._tbl
    analog_proto = copy.deepcopy(tables[1]._tbl)
    discrete_proto = copy.deepcopy(tables[2]._tbl)
    reg_tbl = tables[-1]._tbl
    conts = sectbreak = reg_head = sec_proto = sub_proto = None
    for ch in list(body):
        if ch.tag != qn("w:p"):
            continue
        txt = _ptext(ch).strip()
        pPr = ch.find(qn("w:pPr"))
        if conts is None and "СОДЕРЖАНИЕ" in txt.upper():
            conts = ch
        if sectbreak is None and pPr is not None and pPr.find(qn("w:sectPr")) is not None:
            sectbreak = ch
        if reg_head is None and "ЛИСТ РЕГИСТРАЦИИ" in txt.upper():
            reg_head = ch
        if sub_proto is None and re.match(r"^\d+\.\d+[\s\xa0]+\S", txt):
            sub_proto = ch
        if sec_proto is None and re.match(r"^\d+[\s\xa0]+\S", txt) and not re.match(r"^\d+\.\d", txt):
            sec_proto = ch
    if sec_proto is None:
        sec_proto = sub_proto
    for name, el in (("СОДЕРЖАНИЕ", conts), ("разрыв секции", sectbreak),
                     ("ЛИСТ РЕГИСТРАЦИИ", reg_head),
                     ("заголовок раздела", sec_proto), ("подзаголовок", sub_proto)):
        if el is None:
            raise RuntimeError(f"В шаблоне не найден элемент: {name}")
    body_sectPr = body.find(qn("w:sectPr"))

    def head(proto, text, bm, brk=False):
        p = copy.deepcopy(proto)
        _para_set_text(p, text)
        _add_bookmark(p, bm)
        _page_break_before(p, brk)
        return p

    # Что вообще попадёт в этот перечень. Пустых разделов быть не должно:
    # заголовок с таблицей без единой строки занимает целый лист, а в шкафу
    # аналоговых выходов может не быть вовсе. Поэтому сначала собираем состав,
    # выбрасываем пустое и нумеруем уже оставшееся — и в содержании, и в теле.
    plan = []
    for cab, rows in sections:
        an, di, im = _split_rows(rows, io_, analog_types)
        parts = [(w, p, r) for w, p, r in (("Аналоговые", analog_proto, an),
                                           ("Дискретные", discrete_proto, di),
                                           ("Импульсные", discrete_proto, im)) if r]
        if not parts:
            log(f"  {cab}: {word.lower()} сигналов нет — раздел не создаётся")
            continue
        plan.append((cab, parts))

    # содержание
    toc = []
    for i, (cab, parts) in enumerate(plan, 1):
        toc.append((f"{i} {cab.upper()}", f"bmsec{i}"))
        for j, (w, _p, _r) in enumerate(parts, 1):
            toc.append((f"{i}.{j} {w} {word.lower()} сигналы. {_cab_title(cab)}",
                        f"bm{i}_{j}"))
    toc.append(("ЛИСТ РЕГИСТРАЦИИ ИЗМЕНЕНИЙ", "bmreg"))
    trs = toc_tbl.findall(qn("w:tr"))
    proto = copy.deepcopy(trs[0])
    for tr in trs:
        toc_tbl.remove(tr)
    for text, bm in toc:
        tr = copy.deepcopy(proto)
        cs = _cells(tr)
        _tc_set_text(cs[0], text)
        if len(cs) > 1:
            _cell_pageref(cs[1], bm)
        toc_tbl.append(tr)

    _add_bookmark(reg_head, "bmreg")
    new = [conts, toc_tbl, sectbreak]
    for i, (cab, parts) in enumerate(plan, 1):
        # Каждый раздел с нового листа — но свойством самого заголовка, а не
        # отдельным абзацем с разрывом: иначе на стыке, где таблица закончилась
        # ровно внизу листа, вклинивался пустой лист.
        new.append(head(sec_proto, f"{i} {cab.upper()}", f"bmsec{i}", brk=i > 1))
        for j, (w, proto_tbl, rr) in enumerate(parts, 1):
            new.append(head(sub_proto,
                            f"{i}.{j} {w.upper()} {word} СИГНАЛЫ. {cab.upper()}",
                            f"bm{i}_{j}", brk=j > 1))
            new.append(_build_table(
                proto_tbl,
                f"Таблица {i}.{j} {w} {word.lower()} сигналы. {_cab_title(cab)}", rr))
        log("  %s: %s" % (cab, ", ".join(f"{w.lower()} {len(r)}"
                                         for w, _p, r in parts)))
    _page_break_before(reg_head, True)
    new.append(reg_head)
    new.append(reg_tbl)
    # body_sectPr — параметры последней секции (поля, формат листа). В нормальном
    # шаблоне он есть всегда; если его нет, append(None) падал невнятной ошибкой
    # lxml. Остальные шесть элементов шаблона проверяются выше — проверим и этот.
    if body_sectPr is None:
        raise RuntimeError(
            "В шаблоне не найдены параметры страницы в конце документа (sectPr).\n"
            "Похоже, шаблон повреждён — возьмите исходный файл перечня.")
    new.append(body_sectPr)
    for ch in list(body):
        body.remove(ch)
    for e in new:
        body.append(e)
    _enable_update_fields(doc)
    doc.save(outfile)
    try:
        _fix_date_font(outfile)
    except Exception as e:
        # шрифт даты в штампе — оформление; документ уже сохранён и пригоден
        log(f"  ⚠ не удалось задать шрифт даты в штампе ({type(e).__name__}: {e})")
        log("    документ сохранён, проверьте дату в основной надписи вручную")
    log(f"  сохранено: {outfile}")

# ---------------------------------------------------------------- Excel-сводка

def export_xlsx(sections, path):
    """Сводка сигналов в .xlsx: лист «Сигналы» + лист «Сводка».
    sections: [(имя_раздела, rows)]"""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    thin = Border(*[Side(style="thin", color="D0D7E2")] * 4)
    head_fill = PatternFill("solid", fgColor="2563EB")
    head_font = Font(color="FFFFFF", bold=True, size=10)
    alt_fill = PatternFill("solid", fgColor="F3F6FB")
    res_font = Font(color="9CA3AF", italic=True)

    ws = wb.active
    ws.title = "Сигналы"
    cols = ["Шкаф", "Позиция в контроллере", "Тип", "Позиция по проекту",
            "Описание сигнала", "Уровень управления", "Уровень сигнала", "Взрывозащита"]
    ws.append(cols)
    for c in ws[1]:
        c.fill, c.font, c.border = head_fill, head_font, thin
        c.alignment = Alignment(horizontal="center", vertical="center")
    r = 2
    for cab, rows in sections:
        for d in rows:
            ws.append([cab, d["kc"], d["type"], d["tag"], d["desc"],
                       d["ctrl"], d["level"], d["ex"]])
            for c in ws[r]:
                c.border = thin
                if r % 2 == 0:
                    c.fill = alt_fill
                if d["desc"] == "Резерв":
                    c.font = res_font
            r += 1
    widths = [12, 20, 7, 18, 70, 12, 12, 12]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:H{r - 1}"

    sm = wb.create_sheet("Сводка")
    sm.append(["Шкаф", "Тип", "Всего каналов", "Занято", "Резерв"])
    for c in sm[1]:
        c.fill, c.font, c.border = head_fill, head_font, thin
    r = 2
    for cab, rows in sections:
        by = {}
        for d in rows:
            t = by.setdefault(d["type"], [0, 0])
            t[0] += 1
            if d["desc"] != "Резерв":
                t[1] += 1
        for typ in sorted(by):
            tot, used = by[typ]
            sm.append([cab, typ, tot, used, tot - used])
            for c in sm[r]:
                c.border = thin
            r += 1
    for i, w in enumerate([14, 8, 14, 10, 10], 1):
        sm.column_dimensions[get_column_letter(i)].width = w
    wb.save(path)
    return path


# ------------------------------------------------------------ готовность среды

def word_installed():
    """Установлен ли Microsoft Word (нужен для номеров страниц и PDF)."""
    try:
        import winreg
        for root in (winreg.HKEY_CLASSES_ROOT,):
            try:
                k = winreg.OpenKey(root, r"Word.Application\CurVer")
                winreg.CloseKey(k)
                return True
            except OSError:
                pass
    except Exception:
        pass
    return False


def readiness():
    """Что есть на этом компьютере и что из-за отсутствующего не заработает.

    Возвращает список кортежей (ок, что, подробность, чем грозит).
    Нужна, чтобы человек узнавал об ограничениях сразу, а не наткнувшись
    на ошибку посреди работы.
    """
    out = []

    eng, exe = find_converter()
    out.append((bool(exe), "Чертежи DWG",
                ("AutoCAD" if eng == "acad" else "ODA File Converter") if exe
                else "конвертер не найден",
                "" if exe else "DWG открыть не получится — сохраните чертёж из "
                               "AutoCAD как DXF либо поставьте бесплатный "
                               "ODA File Converter"))

    w = word_installed()
    out.append((w, "Microsoft Word", "установлен" if w else "не найден",
                "" if w else "перечни соберутся, но номера страниц в содержании "
                             "останутся нулями — откройте документ и нажмите Ctrl+A, F9"))

    f = pick_subst_font()
    good = os.path.basename(f).lower().startswith("isocp") if f else False
    out.append((bool(f) and good, "Шрифт чертежей",
                os.path.basename(f) if f else "не найден",
                "" if good else ("будет использован %s — чертёж напечатается, но "
                                 "начертание не по ГОСТ" % (os.path.basename(f) or "шрифт по умолчанию"))))

    t_in, t_out = default_template("in"), default_template("out")
    out.append((bool(t_in and t_out), "Шаблоны перечней",
                "встроенные" if (t_in and t_out) else "не найдены",
                "" if (t_in and t_out) else "укажите свой шаблон в «Настройках» — "
                                            "иначе собирать перечни не из чего"))
    return out


# ---------------------------------------------------------- PDF-печать листов

# ГОСТ 2.301: основные и все дополнительные (удлинённые) форматы, мм
GOST_FORMATS = {
    "А0": (841, 1189), "А1": (594, 841), "А2": (420, 594), "А3": (297, 420),
    "А4": (210, 297), "А5": (148, 210),
    "А0х2": (1189, 1682), "А0х3": (1189, 2523),
    "А1х3": (841, 1783), "А1х4": (841, 2378),
    "А2х3": (594, 1261), "А2х4": (594, 1682), "А2х5": (594, 2102),
    "А3х3": (420, 891), "А3х4": (420, 1189), "А3х5": (420, 1486),
    "А3х6": (420, 1783), "А3х7": (420, 2080),
    "А4х3": (297, 630), "А4х4": (297, 841), "А4х5": (297, 1051),
    "А4х6": (297, 1261), "А4х7": (297, 1471), "А4х8": (297, 1682),
    "А4х9": (297, 1892),
}

# Разрешение фигуры matplotlib. Для векторного PDF влияет только на растровые
# вставки и на точность привязки штрихов; 300 dpi здесь стоили секунд на лист.
RENDER_DPI = 150


def _entity_boxes(msp, log=None):
    """[(сущность, (x0, y0, x1, y1))] — габариты всех сущностей за один проход.
    Раньше ezdxf.bbox.extents вызывался отдельно на каждую сущность, и на
    крупном чертеже это давало десятки тысяч вызовов и минуты ожидания."""
    from ezdxf import bbox as _ezbbox
    ents = list(msp)
    out = []
    try:
        cache = _ezbbox.Cache()
        for e in ents:
            bb = _ezbbox.extents([e], fast=True, cache=cache)
            if bb.has_data:
                out.append((e, (bb.extmin.x, bb.extmin.y, bb.extmax.x, bb.extmax.y)))
    except TypeError:                       # старые версии ezdxf без cache=
        for e in ents:
            try:
                bb = _ezbbox.extents([e], fast=True)
                if bb.has_data:
                    out.append((e, (bb.extmin.x, bb.extmin.y, bb.extmax.x, bb.extmax.y)))
            except Exception:
                pass
    except Exception as ex:
        if log:
            log(f"  ⚠ не удалось посчитать габариты сущностей: {ex}")
    return out


def _match_format(w, h, tol=3.0):
    """(имя, ориентация) или None. Ориентация: 'книжная'/'альбомная'."""
    for name, (a, b) in GOST_FORMATS.items():
        if abs(w - a) <= tol and abs(h - b) <= tol:
            return name, "книжная"
        if abs(w - b) <= tol and abs(h - a) <= tol:
            return name, "альбомная"
    return None

# --- подстановка шрифтов вместо SHX -------------------------------------
# SHX-шрифты AutoCAD не рендерятся сторонними движками, их приходится
# заменять на TTF. Порядок предпочтения: чем выше, тем ближе к ГОСТ 2.304.
SUBST_FONTS = ("isocpeur.ttf", "ISOCPEUR.TTF", "gost_type_a.ttf",
               "arial.ttf", "DejaVuSans.ttf")

# ГОСТ-овские SHX (gostw.shx, GOST.shx и родня) рисуют латинскую N как знак
# номера «№» — так его набирали в чертежах десятилетиями. У TTF такого
# соответствия нет, и после подмены в PDF выходит буква N.
# Сплошная замена N -> № недопустима: в тех же чертежах есть законные
# «NB1-63», «L+N», «NDR-240-24». Поэтому чиним только known-надписи штампа.
SHX_STAMP_FIX = {
    "Инв.N подл.":  "Инв.№ подл.",
    "Инв.N дубл.":  "Инв.№ дубл.",
    "Взам. инв.N":  "Взам. инв.№",
    "Взам.инв.N":   "Взам.инв.№",
    "N докум.":     "№ докум.",
    "N док.":       "№ док.",
    "N строки":     "№ строки",
    "Подп. и дата": "Подп. и дата",
}


NEAR_MISS = 0.08   # насколько рамка может отличаться от формата, чтобы счесть её «почти листом»


def _near_format(w, h):
    """Имя формата ГОСТ, к которому рамка близка, но не попала в допуск.

    Нужно, чтобы отличать неточно начерченную рамку листа от обычного
    прямоугольника внутри схемы: про первую предупреждаем, вторую молчим.
    """
    best, bestd = None, NEAR_MISS
    for name, (a, b) in GOST_FORMATS.items():
        for pw, ph, orient in ((a, b, "книжная"), (b, a, "альбомная")):
            if pw <= 0 or ph <= 0:
                continue
            d = max(abs(w - pw) / pw, abs(h - ph) / ph)
            if d < bestd:
                best, bestd = f"{name} {orient} ({pw}x{ph} мм)", d
    return best


def _fonts_dir_candidates():
    win = os.environ.get("WINDIR", r"C:\Windows")
    return [os.path.join(win, "Fonts"),
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts")]


def pick_subst_font(log=None):
    """Первый реально установленный шрифт из SUBST_FONTS ('' — не нашли)."""
    for fname in SUBST_FONTS:
        for d in _fonts_dir_candidates():
            if d and os.path.exists(os.path.join(d, fname)):
                return fname
    try:                                   # вдруг шрифт виден matplotlib'у
        from matplotlib import font_manager
        for fname in SUBST_FONTS:
            stem = os.path.splitext(fname)[0]
            for f in font_manager.fontManager.ttflist:
                if os.path.basename(f.fname).lower() == fname.lower() or \
                   f.name.lower().replace(" ", "") == stem.lower():
                    return f.fname
    except Exception:
        pass
    if log:
        log("  ⚠ не найден ни один шрифт для замены SHX "
            "(искали: " + ", ".join(SUBST_FONTS) + ")")
        log("    текст будет нарисован шрифтом по умолчанию — ширина букв изменится.")
    return ""


def _fix_shx_styles(doc, log=None):
    """Меняет нерендерящиеся шрифты (SHX и имена без расширения) на TTF.
    Возвращает множество имён стилей, которые были подменены."""
    subst = pick_subst_font(log)
    replaced = set()
    for st in doc.styles:
        try:
            name = st.dxf.name
            f = (st.dxf.font or "").strip()
            # TTF/OTF оставляем как есть; всё прочее — .shx, пустое имя
            # и имя без расширения (в старых чертежах пишут просто «RUS13»)
            if not f.lower().endswith((".ttf", ".otf", ".ttc")):
                if subst:
                    st.dxf.font = subst
                replaced.add(name)
            # bigfont чистим только если основной шрифт мы заменили:
            # там, где кириллица держится на bigfont, его удаление всё ломает
            if st.dxf.bigfont and name in replaced:
                st.dxf.bigfont = ""
        except Exception:
            pass
    return replaced


def _fix_stamp_text(doc, styles, log=None):
    """Возвращает «№» в надписях штампа у текста, чей стиль мы подменили.

    Обходит и пространство модели, и определения блоков: основная надпись
    почти всегда вставлена блоком, и без обхода блоков «N докум.» в штампе
    оставалась непоправленной.
    """
    if not styles:
        return 0

    def fix_one(e):
        try:
            if e.dxf.get("style", "") not in styles:
                return 0
            is_mtext = e.dxftype() == "MTEXT"
            old = e.text if is_mtext else e.dxf.get("text", "")
            if not old or "N" not in old:
                return 0
            new = old
            for a, b in SHX_STAMP_FIX.items():
                if a in new:
                    new = new.replace(a, b)
            if new == old:
                return 0
            if is_mtext:
                e.text = new
            else:
                e.dxf.text = new
            return 1
        except Exception:
            return 0

    KINDS = ("TEXT", "MTEXT", "ATTRIB", "ATTDEF")
    n = 0
    for e in doc.modelspace():
        if e.dxftype() in KINDS:
            n += fix_one(e)
    for blk in doc.blocks:
        try:
            for e in blk:
                if e.dxftype() in KINDS:
                    n += fix_one(e)
        except Exception:
            continue
    if n and log:
        log(f"  восстановлен знак «№» в надписях штампа: {n}")
    return n

def detect_sheets(dxf_path, log=None):
    """Находит рамки листов (замкнутые прямоугольники размеров ГОСТ).
    Возвращает (doc, [(x0, y0, w, h, имя, ориентация)]) — по строкам, слева направо.
    Прямоугольники подходящего размера, но не совпавшие с форматом ГОСТ,
    не выбрасываются молча: о них сообщается в log."""
    doc = read_dxf(dxf_path)
    replaced = _fix_shx_styles(doc, log)
    _fix_stamp_text(doc, replaced, log)
    msp = doc.modelspace()
    frames, seen, odd = [], set(), []
    for e in msp:
        if e.dxftype() != "LWPOLYLINE" or not e.closed or len(e) not in (4, 5):
            continue
        pts = [(pt[0], pt[1]) for pt in e.get_points()]
        xs = [pt[0] for pt in pts]
        ys = [pt[1] for pt in pts]
        w, h = max(xs) - min(xs), max(ys) - min(ys)
        if w < 100 or h < 100:
            continue
        ok = all(abs(px - min(xs)) < 0.05 or abs(px - max(xs)) < 0.05 for px in xs) and \
             all(abs(py - min(ys)) < 0.05 or abs(py - max(ys)) < 0.05 for py in ys)
        if not ok:
            continue
        key = (round(min(xs), 1), round(min(ys), 1))
        if key in seen:
            continue
        m = _match_format(w, h)
        if not m:
            # Не любой прямоугольник — потерянный лист: в схемах полно рамок
            # вокруг блоков. Сообщаем только про «почти формат» — когда размер
            # отличается от ГОСТ-овского не более чем на NEAR_MISS,
            # то есть рамку, скорее всего, начертили неточно.
            near = _near_format(w, h)
            if near:
                odd.append((min(xs), min(ys), w, h, near))
            continue
        seen.add(key)
        frames.append((min(xs), min(ys), w, h, m[0], m[1]))
    if odd and log:
        log(f"  ⚠ рамок, почти совпавших с форматом ГОСТ, но не принятых: {len(odd)}")
        for x0, y0, w, h, near in odd[:5]:
            log(f"      {w:.1f} x {h:.1f} мм в точке ({x0:.0f}, {y0:.0f}) "
                f"— похоже на {near}, но за допуском: лист НЕ будет напечатан")
        if len(odd) > 5:
            log(f"      ... и ещё {len(odd) - 5}")
        log("      поправьте размер рамки в чертеже — либо это не лист, и тогда всё в порядке")
    # группировка по строкам: шаг берём от самой низкой рамки, а не константой
    step = min((f[3] for f in frames), default=200.0) * 0.5 or 200.0
    frames.sort(key=lambda f: (round(-f[1] / step), f[0]))
    return doc, frames

def sheets_summary(frames):
    import collections as _c
    cnt = _c.Counter(f"{f[4]} {f[5][:3]}." for f in frames)
    return ", ".join(f"{k}×{v}" for k, v in cnt.most_common())

# Вес линии ezdxf хранит в миллиметрах, matplotlib рисует в типографских
# пунктах: 1 пункт = 0.3527 мм.
MM_TO_POINTS = 1.0 / 0.3527


def _pdf_config(color=False):
    """Настройки отрисовки листа: монохром, как в AutoCAD.

    Комплект печатают со стилем `monochrome.ctb`: цвет слоя на бумагу не идёт
    (рамка из слоя «новые» зелёная на экране и чёрная на листе), а толщины
    линий идут — рамка и основные линии выходят своей, 0.4…0.7 мм, тонкие
    остаются тонкими. Без этого лист печатался одинаково тонким штрихом и
    выглядел не как чертёж, а как набросок.
    """
    from ezdxf.addons.drawing.config import (Configuration, ColorPolicy,
                                             BackgroundPolicy, LineweightPolicy)
    return Configuration(
        color_policy=ColorPolicy.COLOR_SWAP_BW if color else ColorPolicy.BLACK,
        background_policy=BackgroundPolicy.WHITE,
        lineweight_policy=LineweightPolicy.ABSOLUTE,
        lineweight_scaling=MM_TO_POINTS)


def _draw_entities_safe(frontend, ents, log=None):
    """Рисует сущности листа по одной: одна испорченная не стоит листа.

    ACAD_TABLE держит стили своих ячеек мимо таблицы стилей чертежа: подмена
    SHX на TTF их не касается, ezdxf лезет за `gostw.shx` в папку шрифтов
    AutoCAD и на битой копии («GOSTW (1).SHX») падает вместе со всем листом —
    печать обрывалась с FileStructureError на первой же таблице.

    По одной выходит не дороже, чем пакетом (лист в 911 объектов: 3.0 с
    против 3.4 с), зато лист печатается целиком.
    """
    bad = []
    for e in ents:
        try:
            frontend.draw_entities([e])
        except Exception as ex:
            bad.append((e.dxftype(), type(ex).__name__))
    if bad and log:
        kinds = collections.Counter(b[0] for b in bad)
        log("  ⚠ не нарисовано объектов: %d (%s)"
            % (len(bad), ", ".join(f"{k} — {v}" for k, v in kinds.items())))
    return len(bad)


def export_sheets_pdf(dxf_path, out_pdf, color=False, log=print, progress=None):
    """Печатает каждый лист чертежа отдельной страницей PDF своего формата 1:1.
    Возвращает (out_pdf, frames)."""
    import matplotlib
    matplotlib.use("Agg")
    matplotlib.rcParams["pdf.fonttype"] = 42
    matplotlib.rcParams["ps.fonttype"] = 42
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_pdf import PdfPages
    from ezdxf import bbox as _ezbbox
    from ezdxf.addons.drawing import RenderContext, Frontend
    from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
    doc, frames = detect_sheets(dxf_path, log)
    if not frames:
        log(f"  ⚠ {os.path.basename(dxf_path)}: рамки листов не найдены — PDF не создан")
        return None, []
    log(f"  листов: {len(frames)} ({sheets_summary(frames)})")
    msp = doc.modelspace()
    pairs = _entity_boxes(msp, log)
    ents = [p[0] for p in pairs]
    boxes = [p[1] for p in pairs]
    ctx = RenderContext(doc)
    cfg = _pdf_config(color)
    with PdfPages(out_pdf) as pdf:
        for i, (x0, y0, w, h, name, orient) in enumerate(frames, 1):
            x1, y1 = x0 + w, y0 + h
            sub = [e for e, (ax0, ay0, ax1, ay1) in zip(ents, boxes)
                   if ax0 <= x1 + 1 and ax1 >= x0 - 1 and ay0 <= y1 + 1 and ay1 >= y0 - 1]
            fig = plt.figure(figsize=(w / 25.4, h / 25.4), dpi=RENDER_DPI)
            ax = fig.add_axes([0, 0, 1, 1])
            ax.set_axis_off()
            _draw_entities_safe(Frontend(ctx, MatplotlibBackend(ax), config=cfg),
                                sub, log)
            ax.set_xlim(x0, x1)
            ax.set_ylim(y0, y1)
            ax.set_aspect("equal")
            pdf.savefig(fig)
            plt.close(fig)
            if progress:
                progress(i, len(frames))
    log(f"  сохранено: {os.path.basename(out_pdf)}")
    return out_pdf, frames

def export_channel_sheet_pdf(dxf_path, x, y, out_pdf, color=False, log=print):
    """Печатает один лист — тот, на котором лежит точка (x, y).

    Нужно для перехода «показать канал»: полный PDF чертежа делается минуты,
    а один лист — секунды, и AutoCAD для этого не требуется.
    Возвращает (файл, имя_формата) либо (None, причина).
    """
    doc, frames = detect_sheets(dxf_path, log)
    if not frames:
        return None, "в чертеже не найдены рамки листов"
    hit = None
    for i, (x0, y0, w, h, name, orient) in enumerate(frames, 1):
        if x0 <= x <= x0 + w and y0 <= y <= y0 + h:
            hit = (i, x0, y0, w, h, name, orient)
            break
    if hit is None:
        return None, "канал не попал ни в одну рамку листа"
    i, x0, y0, w, h, name, orient = hit
    import matplotlib
    matplotlib.use("Agg")
    matplotlib.rcParams["pdf.fonttype"] = 42
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_pdf import PdfPages
    from ezdxf.addons.drawing import RenderContext, Frontend
    from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
    msp = doc.modelspace()
    x1, y1 = x0 + w, y0 + h
    sub = [e for e, (ax0, ay0, ax1, ay1) in _entity_boxes(msp, log)
           if ax0 <= x1 + 1 and ax1 >= x0 - 1 and ay0 <= y1 + 1 and ay1 >= y0 - 1]
    cfg = _pdf_config(color)
    with PdfPages(out_pdf) as pdf:
        fig = plt.figure(figsize=(w / 25.4, h / 25.4), dpi=RENDER_DPI)
        ax = fig.add_axes([0, 0, 1, 1])
        ax.set_axis_off()
        _draw_entities_safe(Frontend(RenderContext(doc), MatplotlibBackend(ax), config=cfg),
                            sub, log)
        ax.set_xlim(x0, x1)
        ax.set_ylim(y0, y1)
        ax.set_aspect("equal")
        pdf.savefig(fig)
        plt.close(fig)
    log(f"  лист {i} ({name} {orient}) сохранён: {os.path.basename(out_pdf)}")
    return out_pdf, "лист %d, %s %s" % (i, name, orient)


def export_sheets_pdf_multi(dxf_paths, out_pdf, color=False, log=print, progress=None):
    """Несколько чертежей -> один PDF (листы подряд, каждый своего формата)."""
    import matplotlib
    matplotlib.use("Agg")
    matplotlib.rcParams["pdf.fonttype"] = 42
    matplotlib.rcParams["ps.fonttype"] = 42
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_pdf import PdfPages
    from ezdxf import bbox as _ezbbox
    from ezdxf.addons.drawing import RenderContext, Frontend
    from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
    plan = []
    for dp in dxf_paths:
        doc, frames = detect_sheets(dp, log)
        if frames:
            plan.append((dp, doc, frames))
            log(f"  {os.path.basename(dp)}: {len(frames)} л. ({sheets_summary(frames)})")
        else:
            log(f"  ⚠ {os.path.basename(dp)}: рамки листов не найдены — пропущен")
    if not plan:
        return None
    total = sum(len(f) for _p, _d, f in plan)
    done = 0
    cfg = _pdf_config(color)
    with PdfPages(out_pdf) as pdf:
        for _dp, doc, frames in plan:
            msp = doc.modelspace()
            pairs = _entity_boxes(msp, log)
            ents = [p[0] for p in pairs]
            boxes = [p[1] for p in pairs]
            ctx = RenderContext(doc)
            for (x0, y0, w, h, _name, _orient) in frames:
                x1, y1 = x0 + w, y0 + h
                sub = [e for e, (ax0, ay0, ax1, ay1) in zip(ents, boxes)
                       if ax0 <= x1 + 1 and ax1 >= x0 - 1 and ay0 <= y1 + 1 and ay1 >= y0 - 1]
                fig = plt.figure(figsize=(w / 25.4, h / 25.4), dpi=RENDER_DPI)
                ax = fig.add_axes([0, 0, 1, 1])
                ax.set_axis_off()
                _draw_entities_safe(Frontend(ctx, MatplotlibBackend(ax), config=cfg),
                                    sub, log)
                ax.set_xlim(x0, x1)
                ax.set_ylim(y0, y1)
                ax.set_aspect("equal")
                pdf.savefig(fig)
                plt.close(fig)
                done += 1
                if progress:
                    progress(done, total)
    log(f"  сохранено: {os.path.basename(out_pdf)} ({total} л.)")
    return out_pdf

# ------------------------------------------------- разделы из Excel-сводки

def sections_from_xlsx(path, log=print):
    """Читает Excel-сводку (лист «Сигналы»: Шкаф|Тип|Позиция по проекту|Описание|
    Уровень управления|Уровень сигнала|Взрывозащита|Позиция в контроллере).
    Возвращает [(имя_раздела, rows)] — по одному разделу на каждый «Шкаф»."""
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb["Сигналы"] if "Сигналы" in wb.sheetnames else wb.active
    rows_iter = ws.iter_rows(values_only=True)
    header = None
    data = []
    for row in rows_iter:
        vals = ["" if v is None else str(v).strip() for v in row]
        if header is None:
            if any("Позиция" in v for v in vals) and any("Описание" in v for v in vals):
                header = [v.lower() for v in vals]
            continue
        if not any(vals):
            continue
        data.append(vals)
    if header is None:
        raise RuntimeError(f"В {os.path.basename(path)} не найден заголовок таблицы сигналов")

    def col(*keys):
        for i, h in enumerate(header):
            if any(k in h for k in keys):
                return i
        return None
    c_cab, c_typ = col("шкаф", "раздел"), col("тип")
    c_tag, c_desc = col("позиция по проекту", "тег"), col("описание")
    c_ctrl, c_lvl = col("уровень управления"), col("уровень сигнала")
    c_ex, c_kc = col("взрыв"), col("позиция в конт")
    if c_desc is None or c_kc is None:
        raise RuntimeError("Не хватает колонок «Описание» / «Позиция в контроллере»")
    secs, order = {}, []
    for v in data:
        def g(ci, default=""):
            return v[ci] if ci is not None and ci < len(v) else default
        cab = g(c_cab) or cab_name_from_file(path)
        typ = g(c_typ).upper() or ("DI" if "24" in g(c_lvl) else "AI")
        io_ = "in" if typ in ("AI", "WI", "DI") else "out"
        kc = g(c_kc)
        m = re.match(r"(\d+)\.(\d+)\.(\d+)", kc)
        rows = secs.setdefault(cab, [])
        if cab not in order:
            order.append(cab)
        rows.append(dict(sortx=len(rows), mod=(f"{m.group(1)}.{m.group(2)}" if m else ""),
                         type=typ, io=io_, ch=(int(m.group(3)) if m else len(rows) + 1),
                         tag=g(c_tag), desc=g(c_desc) or "Резерв", level=g(c_lvl),
                         ctrl=g(c_ctrl), ex=g(c_ex), kc=kc))
    log(f"  из Excel: {sum(len(r) for r in secs.values())} каналов, разделов: {len(order)}")
    return [(cab, secs[cab]) for cab in order]

# --------------------------------------------------------- свободные каналы

def free_channels(sections):
    """[(шкаф, модуль, тип, свободно, 'позиции')] по строкам «Резерв»."""
    out = []
    for cab, rows in sections:
        bymod = {}
        for d in rows:
            if d.get("desc") == "Резерв" and d.get("mod"):
                bymod.setdefault((d["mod"], d["type"]), []).append(d["kc"])
        for (mod, typ), kcs in sorted(bymod.items(),
                                      key=lambda k: [int(v) for v in k[0][0].split(".")]):
            out.append((cab, mod, typ, len(kcs), ", ".join(kcs)))
    return out

# ------------------------------------------------------ проверки нормоконтроля

def checks_report(sections):
    """Проверки данных перечней. Возвращает список предупреждений (строк)."""
    warns = []
    for cab, rows in sections:
        seen_kc = {}
        tag_desc = {}
        for d in rows:
            kc = d.get("kc", "")
            if kc:
                if kc in seen_kc:
                    warns.append(f"{cab}: дубль позиции {kc} (каналы задвоены)")
                seen_kc[kc] = True
            tag, desc = d.get("tag", ""), d.get("desc", "")
            if desc == "Резерв":
                if tag:
                    warns.append(f"{cab}: {kc} — «Резерв» с непустым тегом «{tag}»")
                continue
            if tag and not desc:
                warns.append(f"{cab}: {kc} — тег «{tag}» без описания")
            if desc and not tag:
                warns.append(f"{cab}: {kc} — описание без тега (позиции по проекту)")
            if tag:
                tag_desc.setdefault(tag, desc)
        # непрерывность каналов внутри модулей
        bymod = {}
        for d in rows:
            if d.get("mod"):
                bymod.setdefault((d["mod"], d["type"]), []).append(d["ch"])
        for (mod, typ), chs in bymod.items():
            chs = sorted(chs)
            missing = [c for c in range(1, max(chs) + 1) if c not in chs]
            if missing:
                warns.append(f"{cab}: модуль {mod} {typ} — пропущены каналы {missing}")
    # один тег — разные описания (между шкафами тоже)
    tagmap = {}
    for cab, rows in sections:
        for d in rows:
            if d.get("tag") and d.get("desc") != "Резерв":
                tagmap.setdefault((d["tag"], d.get("type")), set()).add(d["desc"])
    for (tag, typ), descs in tagmap.items():
        if len(descs) > 2:
            warns.append(f"Тег {tag} ({typ}) имеет {len(descs)} разных описаний — проверьте")
    return warns

# --------------------------------------------------------- спецификация и КИП

_SKIP_BLOCKS = ("штамп", "рамка", "sw_", "формат", "solid", "wipeout", "подпись")


def _is_equipment_block(nm):
    """Годится ли имя вставки в спецификацию как оборудование.

    Отсеиваются рамки, штампы и подписи, а вместе с ними имена без единой
    буквы («23463246», «123») — это след автонумерации AutoCAD, а не прибор;
    такие блоки попадали в спецификацию наравне с барьерами и клеммами.
    Имена с «*» в начале AutoCAD выдаёт анонимным и динамическим блокам.
    """
    if not nm or nm.startswith("*"):
        return False
    low = nm.lower()
    if any(k in low for k in _SKIP_BLOCKS):
        return False
    return bool(re.search(r"[A-Za-zА-Яа-яЁё]", nm))

def extract_equipment(dxf_path):
    """Состав шкафа из чертежа: модули (по заголовкам), блоки/клеммы (по вставкам).
    Возвращает список (элемент, артикул, кол-во, примечание)."""
    import collections
    doc = read_dxf(dxf_path)
    msp = doc.modelspace()
    texts, inserts = [], collections.Counter()
    for e in msp:
        t = e.dxftype()
        if t in ("TEXT", "MTEXT"):
            try:
                s = e.plain_text() if t == "MTEXT" else e.dxf.text
            except Exception:
                s = getattr(e.dxf, "text", "")
            texts.append((e.dxf.layer, s.strip()))
        elif t == "INSERT":
            nm = e.dxf.name
            if _is_equipment_block(nm):
                inserts[nm] += 1
    out = []
    # модули с артикулами из полных заголовков
    # Артикул пишется через точки («K3.AI.14.16.00»), поэтому точка должна
    # входить в набор символов: без неё в спецификацию попадало «K3».
    hdr = re.compile(r"^Модуль\s+(ввода|вывода)\s+([а-яё\s]+?)\s+сигналов\s+"
                     r"(" + MOD_PREFIX + r"[\d.]+)\s+"
                     r"([A-Z]+),\s*([0-9A-Za-z.\-]+)", re.I)
    # Считаем РАЗНЫЕ номера модулей, а не сколько раз встретился заголовок:
    # модуль, нарисованный на нескольких листах, подписан на каждом из них, и
    # подсчёт заголовков завышал количество почти вдвое — 16 модулей AI вместо
    # девяти фактических.
    mods = collections.defaultdict(set)
    for lay, s in texts:
        m = hdr.match(s)
        if m:
            key = (f"Модуль {m.group(1).lower()} {m.group(2).lower()} сигналов {m.group(4)}",
                   m.group(5))
            mods[key].add(m.group(3))
    kinds_in_headers = set()
    for (name, art), nums in sorted(mods.items()):
        out.append((name, art, len(nums), "модули " + ", ".join(sorted(nums))))
        m = re.search(r"\b(AI|AO|DI|DO|WI)\b", name)
        if m:
            kinds_in_headers.add(m.group(1))
    # дискретные модули по якорям каналов (если не посчитаны по заголовкам)
    for lay, kind in (("Реле DI (1-KL)", "DI"), ("Реле DO (3-KL)", "DO")):
        if kind in kinds_in_headers:
            continue
        nums = set()
        pat = re.compile(MOD_PREFIX + r"(\d+\.\d+)-KLD[IO]\d+$")
        for l, s in texts:
            if l == lay:
                m = pat.match(s)
                if m:
                    nums.add(m.group(1))
        if nums:
            out.append((f"Модуль дискретных сигналов {kind} 16 каналов", "", len(nums),
                        "по каналам " + ", ".join(sorted(nums))))
    # интерфейсный модуль и прочее из зоны состава
    for l, s in texts:
        m = re.search(r"интерфейсный модуль\s+(\S+.*?)\s*\(([\w\-]+)\)", s, re.I)
        if m:
            out.append(("Интерфейсный модуль " + m.group(1), m.group(2), 1, "из состава на чертеже"))
            break
    # блоки/клеммы по вставкам
    for nm, n in inserts.most_common():
        note = ""
        if "биз" in nm.lower() or "rpssi" in nm.lower():
            note = "барьер искрозащиты"
        elif "degson" in nm.lower() or "клемма" in nm.lower():
            note = "клемма"
        out.append((nm, "", n, note))
    return out

def read_equipment(drawings, oda_exe=None, log=print):
    """Состав по списку чертежей (DWG конвертируются) -> [(имя_раздела, items)]."""
    items = [(d, None) if isinstance(d, str) else (d[0], d[1]) for d in drawings]
    dwgs = [p for p, _n in items if p.lower().endswith(".dwg")]
    conv = {}
    if dwgs:
        engine, exe = ("oda", oda_exe) if oda_exe else find_converter()
        conv = dwg_to_dxf(dwgs, engine=engine, exe=exe, log=log)
    out = []
    for p, name in items:
        dxf = conv.get(p, p)
        log(f"Состав: {os.path.basename(p)}")
        out.append((name or cab_name_from_file(p), extract_equipment(dxf)))
    return out

def _run_set_text(r, text):
    """Заменяет текст прогона, сохраняя его оформление."""
    for t in r.findall(qn("w:t")):
        r.remove(t)
    for br in r.findall(qn("w:br")):
        r.remove(br)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)


def _spec_contents_line(proto, title, bm):
    """Строка списка «Спецификация содержит:»: название, отбивка, ссылка на лист.

    Номер листа ставится полем PAGEREF, а не числом: разделы разъезжаются по
    листам при любой правке состава, и вписанное число почти сразу врёт.
    """
    p = copy.deepcopy(proto)
    rs = p.findall(qn("w:r"))
    if not rs:
        raise RuntimeError("Шаблон спецификации: строка-образец списка пуста")
    _run_set_text(rs[0], title)
    for extra in rs[1:]:
        p.remove(extra)
    rpr = rs[0].find(qn("w:rPr"))

    def styled():
        r = OxmlElement("w:r")
        if rpr is not None:
            r.append(copy.deepcopy(rpr))
        return r

    r = styled()
    r.append(OxmlElement("w:tab"))
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = "лист "
    r.append(t)
    p.append(r)
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), f" PAGEREF {bm} \\h ")
    fld.set(qn("w:dirty"), "true")
    fr = styled()
    ft = OxmlElement("w:t")
    ft.text = "0"
    fr.append(ft)
    fld.append(fr)
    p.append(fld)
    return p


SPEC_EXCLUDED_NOTE = (
    "Оборудование, установленное внутри шкафов (модули, барьеры, автоматы, реле, "
    "клеммы, вентиляция), в спецификацию не включается: шкаф поставляется "
    "укомплектованным, и повторённая позиция даёт заказчику право требовать её "
    "отдельно — сверх уже установленной. Подробный состав шкафов выгружается "
    "отдельной ведомостью комплектации (Excel)."
)


def spec_sections_for_cabs(cabs):
    """Раздел «Шкафы» спецификации: по одной позиции поставки на шкаф.

    cabs - имена шкафов (как в разделах перечня, «ШКАФ ШСК1 РСУ»).

    Внутреннее наполнение шкафов сюда сознательно не попадает — почему,
    сказано в SPEC_EXCLUDED_NOTE. Графы «тип, марка, обозначение документа» и
    «завод-изготовитель» остаются пустыми: по схемам подключения они не
    определяются, а выдумывать обозначение документа нельзя.
    """
    items = []
    for c in cabs:
        name = (c or "").strip()
        if not name:
            continue
        # «ШКАФ ШСК1 РСУ» -> «Шкаф ШСК1 РСУ»: обозначение остаётся прописным
        parts = name.split(None, 1)
        if len(parts) == 2 and parts[0] in ("ШКАФ", "ЩИТ"):
            name = parts[0].capitalize() + " " + parts[1]
        items.append((name, "", 1, ""))
    return [("Шкафы", items)] if items else []


def build_spec_docx(template, outfile, equip_sections, log=print):
    """Спецификация оборудования и материалов из шаблона с рамкой и штампом.

    template       - шаблон: шапка таблицы, строка-образец заголовка раздела и
                     строка-образец позиции;
    equip_sections - [(раздел, [(наименование, тип/марка, кол-во, примечание)])].

    Нумерация позиций сквозная по всему документу, как в спецификациях по
    ГОСТ 21.110, а не своя внутри каждого раздела.
    """
    doc = Document(template)
    if not doc.tables:
        raise RuntimeError("Шаблон спецификации не похож на спецификацию: нет таблиц")
    tbl = doc.tables[0]._tbl
    trs = tbl.findall(qn("w:tr"))
    if len(trs) < 3:
        raise RuntimeError(
            "Шаблон спецификации: нужны шапка и две строки-образца — заголовок "
            f"раздела и позиция. Найдено строк: {len(trs)}")
    sec_proto = copy.deepcopy(trs[1])
    item_proto = copy.deepcopy(trs[2])
    for tr in trs[1:]:
        tbl.remove(tr)

    body = doc.element.body
    line_proto = anchor = None
    for p in body.findall(qn("w:p")):
        txt = _ptext(p).strip()
        if re.match(r"^\d+\.\s+\S", txt) and "лист" in txt:
            line_proto, anchor = copy.deepcopy(p), p
            break
    if line_proto is None:
        raise RuntimeError(
            "В шаблоне спецификации не найдена строка-образец списка разделов "
            "(вида «1. Шкафы ... лист 3»)")

    pos = 0
    lines = []
    # Раздел без единой позиции — это заголовок ни над чем: и в таблице, и в
    # списке «Спецификация содержит». Такие не выводим и не нумеруем.
    equip_sections = [(n, it) for n, it in equip_sections if it]
    for i, (name, items) in enumerate(equip_sections, 1):
        bm = f"bmspec{i}"
        tr = copy.deepcopy(sec_proto)
        cs = _cells(tr)
        seen = set()
        for tc in cs:
            if id(tc) in seen:
                continue
            seen.add(id(tc))
            _tc_set_text(tc, "")
        title = f"{i}. {name}"
        _tc_set_text(cs[1], title)
        _add_bookmark(cs[1].find(qn("w:p")), bm)
        tbl.append(tr)
        lines.append((title, bm))
        for it in items:
            nm, art, n, note = (list(it) + ["", "", "", ""])[:4]
            pos += 1
            row = copy.deepcopy(item_proto)
            # графы «код оборудования», «завод-изготовитель» и «масса единицы»
            # по чертежам схем подключения не определяются — оставляем пустыми,
            # чтобы их заполнил проектировщик, а не выдумывала программа
            vals = [str(pos), nm, art, "", "", "шт.", str(n), "", note]
            rc = _cells(row)
            for tc, v in zip(rc, vals):
                _tc_set_text(tc, v)
            for tc in rc[len(vals):]:
                _tc_set_text(tc, "")
            tbl.append(row)
        log(f"  {name}: позиций {len(items)}")

    for title, bm in lines:
        anchor.addprevious(_spec_contents_line(line_proto, title, bm))
    body.remove(anchor)

    # «ЛИСТ РЕГИСТРАЦИИ ИЗМЕНЕНИЙ» в списке разделов: в шаблоне номер листа
    # записан числом из исходного документа и после пересборки врёт. Меняем
    # строку на ссылку к самому заголовку — как и остальные разделы.
    reg_line = reg_head = None
    for p in body.findall(qn("w:p")):
        txt = _ptext(p).strip()
        if not txt.upper().startswith("ЛИСТ РЕГИСТРАЦИИ"):
            continue
        if "лист" in txt[len("ЛИСТ РЕГИСТРАЦИИ"):].lower():
            reg_line = reg_line if reg_line is not None else p
        else:
            reg_head = reg_head if reg_head is not None else p
    if reg_line is not None and reg_head is not None:
        _add_bookmark(reg_head, "bmspecreg")
        reg_line.addprevious(
            _spec_contents_line(line_proto, _ptext(reg_head).strip(), "bmspecreg"))
        body.remove(reg_line)

    _enable_update_fields(doc)
    doc.save(outfile)
    log(f"  сохранено: {outfile} (позиций всего {pos})")
    return outfile


def export_spec_xlsx(equip_sections, sections, path):
    """equip_sections: [(шкаф, [(элемент, артикул, кол-во, прим)])];
    sections: [(шкаф, rows)] для ведомости КИП."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
    wb = Workbook()
    thin = Border(*[Side(style="thin", color="D0D7E2")] * 4)
    hf = PatternFill("solid", fgColor="2563EB")
    hfont = Font(color="FFFFFF", bold=True, size=10)

    ws = wb.active
    ws.title = "Состав шкафов"
    ws.append(["Шкаф", "Элемент", "Артикул", "Кол-во", "Примечание"])
    for c in ws[1]:
        c.fill, c.font, c.border = hf, hfont, thin
    r = 2
    for cab, items in equip_sections:
        for name, art, n, note in items:
            ws.append([cab, name, art, n, note])
            for c in ws[r]:
                c.border = thin
            r += 1
    for i, w in enumerate([14, 58, 22, 8, 40], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    kd = wb.create_sheet("Ведомость КИП")
    kd.append(["Позиция", "Описание (по первому каналу)", "Сигналы", "Каналов", "Шкаф"])
    for c in kd[1]:
        c.fill, c.font, c.border = hf, hfont, thin
    tags = {}
    for cab, rows in sections:
        for d in rows:
            if not d["tag"] or d["desc"] == "Резерв":
                continue
            t = tags.setdefault(d["tag"], dict(desc=d["desc"], types=set(), n=0, cabs=set()))
            t["types"].add(d["type"])
            t["n"] += 1
            t["cabs"].add(cab)
    r = 2
    for tag in sorted(tags):
        t = tags[tag]
        kd.append([tag, t["desc"], "/".join(sorted(t["types"])), t["n"],
                   ", ".join(sorted(t["cabs"]))])
        for c in kd[r]:
            c.border = thin
        r += 2 - 1
    for i, w in enumerate([16, 75, 12, 9, 18], 1):
        kd.column_dimensions[get_column_letter(i)].width = w
    kd.freeze_panes = "A2"
    kd.auto_filter.ref = f"A1:E{r - 1}"
    wb.save(path)
    return path

# ---------------------------------------------------------- штампы (даты и пр.)

_P_RX = re.compile(r"<w:p(?=[ >]).*?</w:p>", re.S)
_T_RX = re.compile(r"(<w:t[^>]*>)(.*?)(</w:t>)", re.S)

def _hf_parts(zf):
    return [n for n in zf.namelist()
            if n.startswith("word/") and n.endswith(".xml")
            and ("header" in n or "footer" in n)]

def detect_stamp_value(path, pattern=r"^\d{2}[.,]\d{2}$"):
    """Самое частое значение-дата в штампах (колонтитулы, включая надписи)."""
    import zipfile, collections
    rx = re.compile(pattern)
    cnt = collections.Counter()
    with zipfile.ZipFile(path) as zf:
        for name in _hf_parts(zf):
            xml = zf.read(name).decode("utf-8")
            for m in _P_RX.finditer(xml):
                txt = "".join(t[1] for t in _T_RX.findall(m.group(0))).strip()
                if rx.match(txt):
                    cnt[txt] += 1
    return cnt.most_common(1)[0][0] if cnt else ""

def replace_in_stamps(paths, old, new, log=print, backup=True):
    """Заменяет значения (напр. дату «05.26») в штампах-колонтитулах docx."""
    import zipfile
    if not old or not str(old).strip():
        raise RuntimeError("Пустое старое значение — нечего заменять")
    total = 0
    for path in paths:
        n = 0
        with zipfile.ZipFile(path) as zf:
            data = {name: zf.read(name) for name in zf.namelist()}
        for name in list(data):
            if not (name.startswith("word/") and name.endswith(".xml")
                    and ("header" in name or "footer" in name)):
                continue
            xml = data[name].decode("utf-8")

            def fix_para(m):
                nonlocal n
                frag = m.group(0)
                txt = "".join(t[1] for t in _T_RX.findall(frag)).strip()
                if txt != old:
                    return frag
                first = [True]

                def sub_t(mt):
                    val = new if first[0] else ""
                    first[0] = False
                    return mt.group(1) + val + mt.group(3)
                n += 1
                return _T_RX.sub(sub_t, frag)
            data[name] = _P_RX.sub(fix_para, xml).encode("utf-8")
        if n:
            if backup:
                shutil.copy2(path, path + ".bak")
            with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
                for name, blob in data.items():
                    zout.writestr(name, blob)
        log(f"  {os.path.basename(path)}: замен {n}")
        total += n
    return total

# ------------------------------------------------- ревизии и паспорт выпуска

REVISION_RE = re.compile(r"^Ревизия\s+(\d+)")


def file_fingerprint(path):
    """(размер, дата изменения, md5) файла. При ошибке — пустые значения."""
    try:
        import hashlib, datetime
        st = os.stat(path)
        h = hashlib.md5()
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(1 << 20), b""):
                h.update(chunk)
        return (st.st_size,
                datetime.datetime.fromtimestamp(st.st_mtime).strftime("%d.%m.%Y %H:%M"),
                h.hexdigest()[:10])
    except Exception:
        return (0, "", "")


def next_revision_dir(base, when=None):
    """Путь следующей папки-ревизии внутри base. Папку не создаёт.

    Номер — максимальный существующий плюс один, а не их количество:
    при подсчёте количества удаление ранней ревизии давало повтор номера,
    и сборка молча писала поверх уже существующей папки. Если путь всё же
    занят (две сборки в один день), номер растёт до свободного.
    """
    import datetime
    when = when or datetime.date.today()
    top = 0
    try:
        for name in os.listdir(base):
            if os.path.isdir(os.path.join(base, name)):
                m = REVISION_RE.match(name)
                if m:
                    top = max(top, int(m.group(1)))
    except OSError:
        pass
    n = top + 1
    while True:
        p = os.path.join(base, "Ревизия %02d — %s" % (n, when.strftime("%d.%m.%Y")))
        if not os.path.exists(p):
            return p
        n += 1


def manifest_text(sources=(), sections=(), out_files=(), templates=None,
                  verify=None, checks=None, app_version=""):
    """Паспорт выпуска: из чего собрано, что получилось и сошлось ли.

    sources   — [(имя_раздела, путь_к_чертежу)]
    sections  — [(шкаф, каналы)]
    out_files — пути выпущенных документов

    Смысл — чтобы через полгода можно было ответить на два вопроса: по какой
    редакции чертежа выпущен документ и сходился ли он на момент выпуска.
    Поэтому пишутся хеши и исходных, и выпущенных файлов.
    """
    import datetime
    L = ["ПАСПОРТ ВЫПУСКА",
         "=" * 60,
         "Собрано: %s" % datetime.datetime.now().strftime("%d.%m.%Y %H:%M"),
         "Программа: Перечни сигналов%s" % ((" v" + app_version) if app_version else ""),
         ""]

    L.append("СВЕРКА С ЧЕРТЕЖАМИ")
    if verify is None:
        L.append("  не выполнялась")
    elif verify.get("ok"):
        c = verify.get("counts", {})
        L.append("  ПРОЙДЕНА — сошлось %s из %s по всем графам"
                 % (c.get("совпало", "?"), c.get("в_чертеже", "?")))
    else:
        iss = verify.get("issues", [])
        c = verify.get("counts", {})
        L.append("  НЕ ПРОЙДЕНА — расхождений %d (сошлось %s из %s)"
                 % (len(iss), c.get("совпало", "?"), c.get("в_чертеже", "?")))
        for k, n in collections.Counter(i[0] for i in iss).most_common():
            L.append("    %s: %d" % (k, n))
        for i in iss[:20]:
            L.append("    %-9s %-24s чертёж: %s | документ: %s"
                     % (i[1], i[0], (i[3] or "—")[:40], (i[4] or "—")[:40]))
        if len(iss) > 20:
            L.append("    ... и ещё %d" % (len(iss) - 20))
    L.append("")

    L.append("ИСХОДНЫЕ ЧЕРТЕЖИ")
    if not sources:
        L.append("  не указаны")
    for name, path in sources:
        size, mt, md5 = file_fingerprint(path)
        L.append("  %s" % (name or "—"))
        L.append("      %s" % path)
        L.append("      изменён %s, %.1f МБ, md5 %s" % (mt or "?", size / 1e6, md5 or "?"))
    L.append("")

    L.append("ШАБЛОНЫ")
    if templates:
        for kind, path in templates.items():
            L.append("  %-10s %s" % (kind, path or "встроенный"))
    else:
        L.append("  встроенные")
    L.append("")

    L.append("РАЗДЕЛЫ")
    total = 0
    for cab, rows in sections:
        used = sum(1 for r in rows if r.get("desc") != "Резерв")
        total += len(rows)
        L.append("  %-24s каналов %4d (занято %d, резерв %d)"
                 % (cab, len(rows), used, len(rows) - used))
    L.append("  всего каналов: %d" % total)
    L.append("")

    L.append("ВЫПУЩЕННЫЕ ДОКУМЕНТЫ")
    if not out_files:
        L.append("  нет")
    for p in out_files:
        size, _mt, md5 = file_fingerprint(p)
        L.append("  %s" % os.path.basename(p))
        L.append("      %.0f КБ, md5 %s" % (size / 1e3, md5 or "?"))
    L.append("")

    L.append("ПРОВЕРКИ ОФОРМЛЕНИЯ")
    checks = list(checks or [])
    L.append("  замечаний: %d" % len(checks))
    for w in checks[:30]:
        L.append("    " + str(w))
    if len(checks) > 30:
        L.append("    ... и ещё %d" % (len(checks) - 30))
    return "\n".join(L)


def write_manifest(path, **kw):
    """Пишет паспорт выпуска рядом с документами. Возвращает путь."""
    with io.open(path, "w", encoding="utf-8") as f:
        f.write(manifest_text(**kw) + "\n")
    return path


# ------------------------------------------------------------- заготовка ПЗ

_MM_DXA = 56.6929           # 1 мм в твипах (1/20 пункта)


def _tbl_fit_cells(tr, n):
    """Приводит строку таблицы к n ячейкам, размножая последнюю."""
    cs = tr.findall(qn("w:tc"))
    if not cs:
        raise RuntimeError("Образец таблицы без ячеек")
    while len(cs) < n:
        tr.append(copy.deepcopy(cs[-1]))
        cs = tr.findall(qn("w:tc"))
    for extra in cs[n:]:
        tr.remove(extra)
    return tr.findall(qn("w:tc"))


def _tbl_set_widths(tbl, widths_mm):
    """Задаёт ширины граф: и сетку таблицы, и ячейки каждой строки."""
    old = tbl.find(qn("w:tblGrid"))
    if old is not None:
        tbl.remove(old)
    grid = OxmlElement("w:tblGrid")
    for w in widths_mm:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(int(w * _MM_DXA)))
        grid.append(gc)
    pr = tbl.find(qn("w:tblPr"))
    if pr is not None:
        pr.addnext(grid)
    else:
        tbl.insert(0, grid)
    for tr in tbl.findall(qn("w:tr")):
        for tc, w in zip(tr.findall(qn("w:tc")), widths_mm):
            tcPr = tc.find(qn("w:tcPr"))
            if tcPr is None:
                tcPr = OxmlElement("w:tcPr")
                tc.insert(0, tcPr)
            for o in tcPr.findall(qn("w:tcW")):
                tcPr.remove(o)
            tcw = OxmlElement("w:tcW")
            tcw.set(qn("w:w"), str(int(w * _MM_DXA)))
            tcw.set(qn("w:type"), "dxa")
            tcPr.insert(0, tcw)


def _tc_align(tc, how):
    """Выравнивание текста в ячейке: 'l' по левому краю, 'c' по центру."""
    for p in tc.findall(qn("w:p")):
        pPr = p.find(qn("w:pPr"))
        if pPr is None:
            pPr = OxmlElement("w:pPr")
            p.insert(0, pPr)
        for o in pPr.findall(qn("w:jc")):
            pPr.remove(o)
        jc = OxmlElement("w:jc")
        jc.set(qn("w:val"), "center" if how == "c" else "left")
        pPr.append(jc)


def _build_pz_table(proto, header, rows_data, widths_mm, align=None):
    """Таблица по образцу из шаблона: шапка, строки, ширины, выравнивание.

    Шапка помечается как повторяемая (w:tblHeader): длинная ведомость
    приборов уходит на несколько листов, и без этого графы подписаны
    только на первом.
    """
    t = copy.deepcopy(proto)
    trs = t.findall(qn("w:tr"))
    if not trs:
        raise RuntimeError("Образец таблицы в шаблоне пуст")
    head_proto = copy.deepcopy(trs[0])
    row_proto = copy.deepcopy(trs[1] if len(trs) > 1 else trs[0])
    for tr in trs:
        t.remove(tr)
    n = len(header)
    align = align or ["l"] * n

    tr = copy.deepcopy(head_proto)
    for tc, v, a in zip(_tbl_fit_cells(tr, n), header, align):
        _tc_set_text(tc, v)
        _tc_align(tc, "c")
    trPr = tr.find(qn("w:trPr"))
    if trPr is None:
        trPr = OxmlElement("w:trPr")
        tr.insert(0, trPr)
    trPr.append(OxmlElement("w:tblHeader"))
    t.append(tr)

    for r in rows_data:
        tr = copy.deepcopy(row_proto)
        for tc, v, a in zip(_tbl_fit_cells(tr, n), r, align):
            _tc_set_text(tc, str(v))
            _tc_align(tc, a)
        t.append(tr)
    _tbl_set_widths(t, widths_mm)
    return t


def _para_pageref(p, name):
    """Делает из абзаца ссылку на номер листа закладки, сохраняя оформление."""
    rs = p.findall(qn("w:r"))
    rpr = None
    if rs and rs[0].find(qn("w:rPr")) is not None:
        rpr = copy.deepcopy(rs[0].find(qn("w:rPr")))
    for r in rs:
        p.remove(r)
    for fs in p.findall(qn("w:fldSimple")):
        p.remove(fs)
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), " PAGEREF %s \\h " % name)
    fld.set(qn("w:dirty"), "true")
    r = OxmlElement("w:r")
    if rpr is not None:
        r.append(rpr)
    t = OxmlElement("w:t")
    t.text = "0"
    r.append(t)
    fld.append(r)
    p.append(fld)
    return p


# Заголовок записки: «2.1 Структура системы» либо строка прописными
# («ПЕРЕЧЕНЬ ПРИНЯТЫХ СОКРАЩЕНИЙ»). Само слово «СОДЕРЖАНИЕ» заголовком не
# считаем — это шапка оглавления, в оглавление она не входит.
_PZ_NUM_HEAD = re.compile(r"^(\d+(?:\.\d+)*)\s+\S")


def _pz_headings(body):
    """Заголовки записки по порядку: [(абзац, текст, номер или None)]."""
    out = []
    for p in body.findall(qn("w:p")):
        t = _ptext(p).strip()
        if not t or t.upper() == "СОДЕРЖАНИЕ":
            continue
        m = _PZ_NUM_HEAD.match(t)
        if m:
            out.append((p, t, m.group(1)))
        elif t == t.upper() and len(t) > 6 and re.match(r"^[А-ЯЁ]", t):
            out.append((p, t, None))
    return out


def _pz_counts(sections):
    """Каналы по шкафам: [(шкаф, {тип: [всего, занято]}, всего, занято)]."""
    out = []
    for cab, rows in sections:
        by = {}
        tot = used = 0
        for d in rows:
            t = by.setdefault(d["type"], [0, 0])
            t[0] += 1
            tot += 1
            if d["desc"] != "Резерв":
                t[1] += 1
                used += 1
        out.append((cab, by, tot, used))
    return out


def build_pz_docx(template, outfile, sections, equip_sections, log=print):
    """Пояснительная записка из шаблона с рамкой и штампом.

    Шаблон несёт структуру по ГОСТ 34.201 и текст, который переносится из
    проекта в проект; данные конкретного проекта в нём помечены словом
    «Заменить». Программа дописывает в него раздел со сводкой сигналов,
    составом технических средств и ведомостью приборов, после чего
    пересобирает оглавление: номера листов ставятся полями, а не числами.
    """
    doc = Document(template)
    body = doc.element.body
    if not doc.tables:
        raise RuntimeError("Шаблон ПЗ не похож на записку: в документе нет таблиц")
    toc = doc.tables[0]._tbl
    # Образец оформления таблиц берём из самого шаблона: первая после
    # оглавления таблица с 2-4 графами (в записке это перечень сокращений).
    # Лист регистрации изменений не годится — он в десять граф.
    tbl_proto = None
    for t_ in doc.tables[1:]:
        if 2 <= len(t_.columns) <= 4 and len(t_.rows) >= 2:
            tbl_proto = copy.deepcopy(t_._tbl)
            break
    if tbl_proto is None:
        raise RuntimeError(
            "В шаблоне ПЗ не найден образец оформления таблицы "
            "(таблица на 2-4 графы, например перечень сокращений)")

    heads = _pz_headings(body)
    if not heads:
        raise RuntimeError("В шаблоне ПЗ не найдено ни одного заголовка раздела")
    sec_proto = sub_proto = txt_proto = None
    for p, t, num in heads:
        if num and "." not in num and sec_proto is None:
            sec_proto = copy.deepcopy(p)
        if num and "." in num and sub_proto is None:
            sub_proto = copy.deepcopy(p)
    for p in body.findall(qn("w:p")):
        t = _ptext(p).strip()
        if txt_proto is None and t and not _PZ_NUM_HEAD.match(t) and t != t.upper():
            txt_proto = copy.deepcopy(p)
    if sec_proto is None or sub_proto is None or txt_proto is None:
        raise RuntimeError(
            "В шаблоне ПЗ не найдены образцы заголовка, подзаголовка или абзаца")

    # Свой раздел ставим последним перед перечнем сокращений: нумерацию
    # продолжаем от последнего числового раздела шаблона.
    last = max((int(n.split(".")[0]) for _p, _t, n in heads if n), default=0)
    top = last + 1
    anchor = None
    for p, t, num in heads:
        if num is None:
            anchor = p
            break
    if anchor is None:
        anchor = heads[-1][0]

    def put(el):
        anchor.addprevious(el)

    def head(proto, text):
        p = copy.deepcopy(proto)
        _para_set_text(p, text)
        return p

    def para(text):
        return head(txt_proto, text)

    def caption(text):
        """Название таблицы по ГОСТ 2.105: над таблицей, от левого края."""
        p = copy.deepcopy(txt_proto)
        _para_set_text(p, text)
        pPr = p.find(qn("w:pPr"))
        if pPr is None:
            pPr = OxmlElement("w:pPr")
            p.insert(0, pPr)
        for o in pPr.findall(qn("w:ind")):
            pPr.remove(o)
        for o in pPr.findall(qn("w:jc")):
            pPr.remove(o)
        jc = OxmlElement("w:jc")
        jc.set(qn("w:val"), "left")
        pPr.append(jc)
        kn = OxmlElement("w:keepNext")
        pPr.insert(0, kn)
        return p

    def table(cols, rows_data, widths, align=None):
        return _build_pz_table(tbl_proto, cols, rows_data, widths, align)

    def pagebreak():
        """Пустой абзац с разрывом страницы.

        Собираем его сами, а не ищем образец в шаблоне: разрыв там стоит
        внутри заголовка раздела, и копия «образца» тащила за собой текст
        этого заголовка — он лишним пунктом попадал в оглавление.
        """
        p = copy.deepcopy(txt_proto)
        for r in p.findall(qn("w:r")):
            p.remove(r)
        r = OxmlElement("w:r")
        br = OxmlElement("w:br")
        br.set(qn("w:type"), "page")
        r.append(br)
        p.append(r)
        return p

    put(pagebreak())
    put(head(sec_proto, f"{top} Сводка сигналов и состав технических средств"))

    put(head(sub_proto, f"{top}.1 Сводка сигналов"))
    counts = _pz_counts(sections)
    kinds = sorted({k for _c, by, _t, _u in counts for k in by})
    put(para(f"Объём сигналов по шкафам приведён в таблице {top}.1. "
             "«Занято» — каналы с назначенным сигналом, остальные каналы "
             "свободны под развитие системы."))
    data = []
    for cab, by, tot, used in counts:
        data.append([_cab_title(cab)] + [by.get(k, [0, 0])[0] for k in kinds]
                    + [tot, used, tot - used])
    if len(counts) > 1:
        data.append(["Итого"]
                    + [sum(by.get(k, [0, 0])[0] for _c, by, _t, _u in counts) for k in kinds]
                    + [sum(c[2] for c in counts), sum(c[3] for c in counts),
                       sum(c[2] - c[3] for c in counts)])
    put(caption(f"Таблица {top}.1 – Объём сигналов по шкафам"))
    put(table(["Шкаф"] + list(kinds) + ["Всего", "Занято", "Резерв"], data,
              [46] + [13] * len(kinds) + [15, 16, 16],
              ["l"] + ["c"] * (len(kinds) + 3)))

    # Пустая таблица под заголовком — лишний лист, поэтому смотрим не на
    # наличие разделов, а на наличие самих строк.
    rows_ = []
    for cab, items in equip_sections or ():
        for name, art, n, _note in items:
            rows_.append([_cab_title(cab), name, art, n])
    if rows_:
        put(head(sub_proto, f"{top}.2 Состав технических средств"))
        put(para("Оборудование, установленное в шкафах, по данным чертежей "
                 "схем подключения."))
        put(caption(f"Таблица {top}.2 – Технические средства в шкафах"))
        put(table(["Шкаф", "Элемент", "Тип, марка", "Кол-во"], rows_,
                  [40, 66, 40, 18], ["l", "l", "l", "c"]))
        nsub = 3
    else:
        nsub = 2

    put(head(sub_proto, f"{top}.{nsub} Ведомость приборов"))
    tags = {}
    for cab, rows in sections:
        for d in rows:
            if not d["tag"] or d["desc"] == "Резерв":
                continue
            t = tags.setdefault(d["tag"], dict(desc=d["desc"], types=set(), cabs=set()))
            t["types"].add(d["type"])
            t["cabs"].add(cab)
    put(para("Позиции приборов и исполнительных механизмов, встречающиеся в "
             "схемах подключения, с указанием шкафа подключения."))
    put(caption(f"Таблица {top}.{nsub} – Приборы и исполнительные механизмы"))
    put(table(["Позиция", "Назначение", "Сигналы", "Шкаф"],
              [[k, tags[k]["desc"], "/".join(sorted(tags[k]["types"])),
                ", ".join(_cab_title(c) for c in sorted(tags[k]["cabs"]))]
               for k in sorted(tags)],
              [26, 74, 20, 44], ["l", "l", "c", "l"]))
    log(f"  раздел {top}: шкафов {len(sections)}, приборов {len(tags)}")

    # оглавление: заголовки + ссылки на листы
    heads = _pz_headings(body)
    trs = toc.findall(qn("w:tr"))
    src = _cells(trs[-1])
    if len(src) < 2 or not src[0].findall(qn("w:p")) or not src[1].findall(qn("w:p")):
        raise RuntimeError("Шаблон ПЗ: оглавление устроено не так, как ожидалось")
    name_proto = copy.deepcopy(src[0].findall(qn("w:p"))[0])
    num_proto = copy.deepcopy(src[1].findall(qn("w:p"))[0])
    for tr in trs[1:]:
        toc.remove(tr)
    cs = _cells(trs[0])
    for tc in cs[:2]:
        for p in tc.findall(qn("w:p")):
            tc.remove(p)
    for i, (p, text, _num) in enumerate(heads, 1):
        bm = f"bmpz{i}"
        _add_bookmark(p, bm)
        np = copy.deepcopy(name_proto)
        _para_set_text(np, text)
        cs[0].append(np)
        cs[1].append(_para_pageref(copy.deepcopy(num_proto), bm))
    log(f"  оглавление: строк {len(heads)}")

    _enable_update_fields(doc)
    doc.save(outfile)
    log(f"  сохранено: {outfile}")
    return outfile


def export_pz_docx(sections, equip_sections, path):
    """Заготовка раздела пояснительной записки: сводка каналов, состав ТС, КИП."""
    doc = Document()
    doc.add_heading("Автоматизация. Сигналы и технические средства (заготовка)", 1)
    doc.add_paragraph(
        "Раздел сформирован автоматически по чертежам схем подключения. "
        "Текст подлежит редактированию и нормоконтролю.")
    doc.add_heading("1 Сводка сигналов", 2)
    total = dict()
    for cab, rows in sections:
        by = {}
        for d in rows:
            t = by.setdefault(d["type"], [0, 0])
            t[0] += 1
            if d["desc"] != "Резерв":
                t[1] += 1
            g = total.setdefault(d["type"], [0, 0])
            g[0] += 1
            if d["desc"] != "Резерв":
                g[1] += 1
        parts = [f"{k}: {v[1]} из {v[0]} (резерв {v[0]-v[1]})" for k, v in sorted(by.items())]
        doc.add_paragraph(f"{cab}: " + "; ".join(parts) + ".")
    parts = [f"{k}: {v[1]} из {v[0]}" for k, v in sorted(total.items())]
    doc.add_paragraph("Итого по системе: " + "; ".join(parts) + ".")

    if equip_sections:
        doc.add_heading("2 Состав технических средств", 2)
        tbl = doc.add_table(rows=1, cols=4)
        tbl.style = "Table Grid"
        for i, h in enumerate(("Шкаф", "Элемент", "Артикул", "Кол-во")):
            tbl.rows[0].cells[i].text = h
        for cab, items in equip_sections:
            for name, art, n, _note in items:
                r = tbl.add_row().cells
                r[0].text, r[1].text, r[2].text, r[3].text = cab, name, art, str(n)

    doc.add_heading("3 Ведомость приборов КИП", 2)
    tags = {}
    for cab, rows in sections:
        for d in rows:
            if not d["tag"] or d["desc"] == "Резерв":
                continue
            t = tags.setdefault(d["tag"], dict(desc=d["desc"], types=set(), cabs=set()))
            t["types"].add(d["type"])
            t["cabs"].add(cab)
    tbl = doc.add_table(rows=1, cols=4)
    tbl.style = "Table Grid"
    for i, h in enumerate(("Позиция", "Назначение", "Сигналы", "Шкаф")):
        tbl.rows[0].cells[i].text = h
    for tag in sorted(tags):
        t = tags[tag]
        r = tbl.add_row().cells
        r[0].text, r[1].text = tag, t["desc"]
        r[2].text, r[3].text = "/".join(sorted(t["types"])), ", ".join(sorted(t["cabs"]))
    doc.save(path)
    return path

# ------------------------------------------------------------- сравнение версий

_KC_RE = re.compile(r"^\d+\.\d+\.\d+$")
_PER_HINTS = ("п/п", "позиц", "описание", "уровень", "взрыв", "контрол", "сигнал")


def _tc_text(tc):
    return "".join(t.text or "" for t in tc.iter(qn("w:t"))).strip()


def _tr_texts(tr):
    """Тексты ячеек строки с раскрытием объединённых по горизонтали."""
    out = []
    for tc in tr.findall(qn("w:tc")):
        span = 1
        pr = tc.find(qn("w:tcPr"))
        if pr is not None:
            gs = pr.find(qn("w:gridSpan"))
            if gs is not None:
                try:
                    span = max(1, int(gs.get(qn("w:val"))))
                except (TypeError, ValueError):
                    span = 1
        out += [_tc_text(tc)] * span
    return out


def _per_is_title(cs):
    """Строка-название таблицы: вся строка объединена в одну ячейку."""
    return bool(cs) and bool(cs[0]) and len(set(cs)) == 1


def _per_is_header(cs):
    """Шапка граф: несколько характерных слов и никакой позиции в контроллере."""
    if cs and _KC_RE.match(cs[-1]):
        return False
    j = " ".join(cs).lower()
    return sum(1 for w in _PER_HINTS if w in j) >= 3


def _per_context(text, cab, io_):
    """Обновляет шкаф и направление по тексту заголовка раздела или таблицы."""
    low = text.lower()
    if "входн" in low:
        io_ = "in"
    elif "выходн" in low:
        io_ = "out"
    m = re.search(r"(Шкаф|Щит)\s+([^.,;]+?)\s*$", text.strip(), re.I)
    if m:
        cab = m.group(1).upper() + " " + m.group(2).upper()
    return cab, io_


def parse_perechen(path):
    """Читает существующий перечень .docx -> [{cab, kc, tag, desc, ctrl, level, ex}].

    Вёрстка перечней в разных проектах разная. Где-то название таблицы стоит
    первой строкой самой таблицы («Таблица 1.1 Аналоговые входные сигналы.
    Шкаф ШСК1 РСУ»), где-то — отдельным абзацем над ней, а сама таблица разбита
    на десятки кусков по листам: в одних кусках шапка граф повторена, в других
    строки начинаются сразу с данных. Прежний разбор считал строки по номеру —
    первые две всегда пропускал — и на второй вёрстке терял по строке с каждого
    куска, а шкаф и направление оставались пустыми, из-за чего сверка не
    находила ни одной пары.

    Поэтому идём по документу подряд, запоминая последний заголовок, а строки
    разбираем по содержимому: название, шапка или данные.
    """
    doc = Document(path)
    out = []
    cab = io_ = ""
    # направление по имени файла — запасной вариант: в части проектов его
    # не пишут ни в заголовке раздела, ни в названии таблицы
    base = os.path.basename(path).lower()
    fallback_io = "in" if "входн" in base else ("out" if "выходн" in base else "")
    for ch in doc.element.body:
        if ch.tag == qn("w:p"):
            t = _ptext(ch).strip()
            if t:
                cab, io_ = _per_context(t, cab, io_)
            continue
        if ch.tag != qn("w:tbl"):
            continue
        trs = ch.findall(qn("w:tr"))
        if not trs:
            continue
        # Граф бывает семь (с «№ п/п») и шесть: перечни проекта В с сентября идут
        # без номера строки. Считаем от конца — позиция в контроллере всегда
        # последняя графа. По жёсткому «ровно семь» такой перечень читался
        # как ноль строк, и сверка молча сверяла пустоту.
        ncol = len(_tr_texts(trs[0]))
        if ncol not in (6, 7):
            continue
        for tr in trs:
            cs = _tr_texts(tr)
            if len(cs) != ncol:
                continue
            if _per_is_title(cs):
                cab, io_ = _per_context(cs[0], cab, io_)
                continue
            if _per_is_header(cs) or not cs[-1]:
                continue
            out.append(dict(cab=cab, io=io_ or fallback_io, tag=cs[-6],
                            desc=re.sub(r"\s+", " ", cs[-5]),
                            ctrl=cs[-4], level=cs[-3], ex=cs[-2], kc=cs[-1]))
    return out

def compare(sections, old_rows, log=print):
    """Сравнивает извлечённое из чертежей со старым перечнем.
    Возвращает список (статус, шкаф, kc, было, стало)."""
    new = {}
    for cab, rows in sections:
        for d in rows:
            new[(cab.upper(), d["kc"])] = d
    old = {(o["cab"].upper(), o["kc"]): o for o in old_rows}
    report = []
    for key in sorted(set(new) | set(old), key=lambda k: (k[0], k[1])):
        n, o = new.get(key), old.get(key)
        cab, kc = key
        if n and not o:
            if n["desc"] != "Резерв":
                report.append(("добавлено", cab, kc, "", f"{n['tag']} | {n['desc']}"))
        elif o and not n:
            if o["desc"] != "Резерв":
                report.append(("удалено", cab, kc, f"{o['tag']} | {o['desc']}", ""))
        else:
            was = f"{o['tag']} | {o['desc']}"
            now = f"{n['tag']} | {n['desc']}"
            if (o["tag"] != n["tag"] or o["desc"] != n["desc"]) and not (
                    o["desc"] == "Резерв" and n["desc"] == "Резерв"):
                if o["desc"] == "Резерв" and n["desc"] != "Резерв":
                    report.append(("добавлено", cab, kc, "Резерв", now))
                elif n["desc"] == "Резерв" and o["desc"] != "Резерв":
                    report.append(("удалено", cab, kc, was, "Резерв"))
                else:
                    report.append(("изменено", cab, kc, was, now))
    log(f"Сравнение: изменений {len(report)} "
        f"(добавлено {sum(1 for r in report if r[0]=='добавлено')}, "
        f"удалено {sum(1 for r in report if r[0]=='удалено')}, "
        f"изменено {sum(1 for r in report if r[0]=='изменено')})")
    return report

# ----------------------------------------------------- сверка перечня с чертежом

def _norm(s):
    """Текст для сравнения: без лишних пробелов, мягких переносов и регистра."""
    s = (s or "").replace("­", "").replace("‑", "-").replace("\xa0", " ")
    return re.sub(r"\s+", " ", s).strip().lower()


def verify_perechen(sections, docx_paths, log=print):
    """Сверяет готовый перечень с тем, что извлечено из чертежей.

    В отличие от compare(), где сравниваются две РАЗНЫЕ версии и расхождения
    ожидаемы, здесь источник один и тот же: любое расхождение — дефект выпуска.
    Проверяется не только позиция с описанием, но и уровень сигнала,
    взрывозащита и вид контроля — они тоже формируются программой.

    Возвращает dict: counts, issues, ok.
    """
    doc_rows = []
    for p in docx_paths:
        try:
            rows = parse_perechen(p)
        except Exception as e:
            log(f"  не удалось прочитать {os.path.basename(p)}: {type(e).__name__}: {e}")
            continue
        for r in rows:
            r["src"] = os.path.basename(p)
        doc_rows += rows
        log(f"  прочитано из {os.path.basename(p)}: строк {len(rows)}")
    if not doc_rows:
        log("  ⚠ в документах не найдено ни одной строки перечня —")
        log("    сверять не с чем: это не перечень либо таблицы другого вида")
        return dict(counts={}, issues=[], ok=False)

    # шкаф в документе распознаётся не всегда; если его нет — сверяем по (io, kc)
    # Шкаф учитываем, только если он распознан у ВСЕХ строк документа.
    # При «хотя бы у одной» часть ключей была со шкафом, часть без —
    # и не совпадало вообще ничего.
    use_cab = bool(doc_rows) and all(r["cab"] for r in doc_rows)
    # Мало того, что шкаф назван: он должен называться ТАК ЖЕ, как чертёж.
    # В перечне проекта В все таблицы идут под одним «ШКАФ РСУ», а чертежи зовутся
    # «ШКАФ ШСК1 РСУ»…«ШКАФ ШСК5 РСУ» — шкаф в ключе был, не совпадало ничего
    # из полутора тысяч строк, и сверка объявляла дефектом весь перечень.
    if use_cab:
        draw_cabs = {(cab or "").upper() for cab, _rows in sections}
        hit = sum(1 for r in doc_rows if (r["cab"] or "").upper() in draw_cabs)
        if hit * 2 < len(doc_rows):
            use_cab = False
            log(f"  шкафы в перечне названы не так, как чертежи "
                f"(совпало строк: {hit} из {len(doc_rows)}) — "
                f"сверка по каналу и направлению")
    key_doc = (lambda r: ((r["cab"] or "").upper(), r["io"], r["kc"])) if use_cab \
        else (lambda r: ("", r["io"], r["kc"]))
    if not use_cab and not all(r["cab"] for r in doc_rows):
        log("  в документе не указан шкаф в заголовках таблиц — сверка по каналу и направлению")

    dm, dup = {}, []
    for r in doc_rows:
        k = key_doc(r)
        if k in dm:
            dup.append(k)
        else:
            dm[k] = r

    # Сверяем только направления, представленные в поданных документах: если дан
    # лишь перечень входных, выходные каналы чертежа отсутствующими не считаются.
    io_have = {r["io"] for r in doc_rows if r["io"]}
    dr = {}
    skipped_io = collections.Counter()
    for cab, rows in sections:
        for d in rows:
            if io_have and d["io"] not in io_have:
                skipped_io[d["io"]] += 1
                continue
            k = ((cab or "").upper(), d["io"], d["kc"]) if use_cab else ("", d["io"], d["kc"])
            dr[k] = d
    if skipped_io:
        names = {"in": "входных", "out": "выходных"}
        log("  не сверялись (нет соответствующего перечня): "
            + ", ".join(f"{names.get(k, k)} {v}" for k, v in skipped_io.items()))

    issues = []
    for k in dup:
        issues.append(("дубль", k[2], k[1], "строка встречается в документе дважды", ""))

    FIELDS = (("tag", "позиция"), ("desc", "описание"), ("level", "уровень"),
              ("ex", "взрывозащита"), ("ctrl", "вид контроля"))
    same = 0
    for k in sorted(set(dr) | set(dm), key=lambda t: (t[0], t[1], t[2])):
        a, b = dr.get(k), dm.get(k)
        kc, io_ = k[2], k[1]
        if a and not b:
            issues.append(("нет в документе", kc, io_,
                           f"{a['tag']} | {a['desc']}"[:90], ""))
        elif b and not a:
            issues.append(("нет в чертеже", kc, io_, "",
                           f"{b['tag']} | {b['desc']}"[:90]))
        else:
            diff = [(ru, a.get(f, ""), b.get(f, ""))
                    for f, ru in FIELDS if _norm(a.get(f)) != _norm(b.get(f))]
            if diff:
                for ru, va, vb in diff:
                    issues.append((f"расходится: {ru}", kc, io_,
                                   str(va)[:70], str(vb)[:70]))
            else:
                same += 1

    counts = dict(в_чертеже=len(dr), в_документе=len(dm), совпало=same,
                  расхождений=len(issues))
    log("")
    log(f"  каналов в чертежах: {len(dr)}   в перечнях: {len(dm)}   сошлось полностью: {same}")
    if not issues:
        log("  ✓ СВЕРКА ПРОЙДЕНА: расхождений нет")
    else:
        kinds = collections.Counter(i[0] for i in issues)
        log(f"  ⚠ расхождений: {len(issues)}")
        for kind, n in kinds.most_common():
            log(f"      {kind}: {n}")
        for it in issues[:8]:
            log(f"      {it[1]:<9} {it[0]:<22} чертёж: {it[3][:40]!r}  документ: {it[4][:40]!r}")
        if len(issues) > 8:
            log(f"      ... и ещё {len(issues) - 8}")
    return dict(counts=counts, issues=issues, ok=not issues)


def export_verify_xlsx(result, path):
    """Отчёт сверки в .xlsx: лист «Расхождения» + лист «Итог»."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    wb = Workbook()
    ws = wb.active
    ws.title = "Расхождения"
    head = ["Что не так", "Канал", "Направление", "В чертеже", "В документе"]
    ws.append(head)
    fill = PatternFill("solid", fgColor="B91C1C")
    for i, _h in enumerate(head, 1):
        c = ws.cell(row=1, column=i)
        c.font = Font(color="FFFFFF", bold=True)
        c.fill = fill
    for it in result["issues"]:
        ws.append(list(it))
    for col, w in zip("ABCDE", (26, 12, 13, 52, 52)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    ws2 = wb.create_sheet("Итог")
    ws2.append(["Показатель", "Значение"])
    ws2.cell(row=1, column=1).font = Font(bold=True)
    ws2.cell(row=1, column=2).font = Font(bold=True)
    for k, v in result["counts"].items():
        ws2.append([k.replace("_", " "), v])
    ws2.append(["вердикт", "сошлось" if result["ok"] else "есть расхождения"])
    ws2.column_dimensions["A"].width = 22
    ws2.column_dimensions["B"].width = 20
    wb.save(path)
    return path


def export_compare_xlsx(report, path):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
    wb = Workbook()
    ws = wb.active
    ws.title = "Изменения"
    thin = Border(*[Side(style="thin", color="D0D7E2")] * 4)
    ws.append(["Статус", "Шкаф", "Позиция", "Было", "Стало"])
    fills = {"добавлено": PatternFill("solid", fgColor="DCFCE7"),
             "удалено": PatternFill("solid", fgColor="FEE2E2"),
             "изменено": PatternFill("solid", fgColor="FEF9C3")}
    for c in ws[1]:
        c.font = Font(bold=True)
        c.border = thin
    for i, (st, cab, kc, was, now) in enumerate(report, 2):
        ws.append([st, cab, kc, was, now])
        for c in ws[i]:
            c.border = thin
            c.fill = fills.get(st, PatternFill())
    for i, w in enumerate([12, 14, 12, 60, 60], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    wb.save(path)
    return path

# --------------------------------------------------------- генерация схем (бета)

# Генератор схем вынесен в experimental/generate_scheme.py (19.08.2026):
# функция не вызывалась ни ядром, ни интерфейсом, тестов не имела,
# при этом писала DXF. Подробности — в ИЗМЕНЕНИЯ.md.

# ---------------------------------------------------------------- оркестрация

# С/C и К/K допускаем и кириллицей и латиницей (в именах файлов их путают)
CAB_RE = re.compile(r"(Ш[СC][КK]\s*\d+|Ш[СC]\s*\d+|\d+\s*Щ\b|Щ\s*\d+)", re.I)

def _app_dir():
    """Папка программы: рядом с .exe в сборке, рядом с .py в исходниках."""
    import sys
    for base in (os.path.dirname(os.path.abspath(sys.argv[0] or "")),
                 os.path.dirname(os.path.abspath(__file__))):
        if base and os.path.isdir(os.path.join(base, "templates")):
            return base
    return os.path.dirname(os.path.abspath(__file__))


TEMPLATE_FILES = {
    "in":   "Перечень входных сигналов.docx",
    "out":  "Перечень выходных сигналов.docx",
    "spec": "Спецификация оборудования и материалов.docx",
    "pz":   "Пояснительная записка.docx",
}


def default_template(kind):
    """Встроенный шаблон ('in' | 'out' | 'spec'); '' — если не найден.

    Шаблоны лежат в templates/ рядом с программой и оформлены по образцу
    выпущенного комплекта: лист А4 с рамкой и основной надписью по ГОСТ 2.104,
    шифр и наименование объекта заменены на «Заменить», даты — на «sss».
    Выбирать файл отдельно не нужно; свой шаблон по-прежнему можно указать
    на странице «Сборка».
    """
    name = TEMPLATE_FILES.get(kind)
    if not name:
        return ""
    q = os.path.join(_app_dir(), "templates", name)
    return q if os.path.exists(q) else ""


# Уточнение шкафа: «ШСК1 РСУ», «ШСК1 ПАЗ». Без него два разных шкафа проекта
# получают одно имя «ШКАФ ШСК1» и молча сливаются в один раздел.
CAB_SUFFIX_RE = re.compile(r"\s+([А-ЯЁA-Z]{2,4}(?:/[А-ЯЁA-Z]{2,4})?)(?![а-яёa-z])")
# Шкафы без номера: ШИБП, ШУ. Ищем только прописные — иначе под шаблон попадает
# слово «шкафа» из названия файла.
CAB_BARE_RE = re.compile(r"\bШ[А-ЯЁ]{1,4}\b")
# Латинские двойники кириллицы: в чертежах обозначения набраны вперемешку.
_LAT2CYR = str.maketrans("ACEHKMOPTXBY", "АСЕНКМОРТХВУ")


def _cab_word(s):
    """Слово обозначения в кириллице или None, если это не обозначение.

    Латиница переводится в кириллицу («PCУ» -> «РСУ»); если после перевода
    латинские буквы остались, это не русское обозначение, а случайное слово
    вроде «PDF» — такое уточнением шкафа не считаем.
    """
    v = s.translate(_LAT2CYR)
    return None if re.search(r"[A-Za-z]", v) else v


# Место установки в имени файла: «…шкафа системного ШС1. Здание 192Л».
# Без него три шкафа проекта, стоящие в разных зданиях, получают одно имя
# «ШКАФ ШС1» и сливаются в один раздел перечня.
CAB_PLACE_RE = re.compile(r"(Здание|Корпус|Цех|Этаж|Помещение)\s+([^.,;]+?)\s*$",
                          re.I)


def _cab_place(name):
    m = CAB_PLACE_RE.search(name)
    return (" " + m.group(1).capitalize() + " " + m.group(2).strip()) if m else ""


def cab_name_from_file(path):
    name = os.path.splitext(os.path.basename(path))[0]
    m = CAB_RE.search(name)
    if m:
        val = re.sub(r"\s+", "", m.group(1).upper()).translate(_LAT2CYR)
        s = CAB_SUFFIX_RE.match(name, m.end())
        if s:
            w = _cab_word(s.group(1))
            if w:
                val += " " + w
        return ("ЩИТ " if "Щ" in val else "ШКАФ ") + val + _cab_place(name)
    m = CAB_BARE_RE.search(name)
    if m:
        return "ШКАФ " + m.group(0) + _cab_place(name)
    return name

def _kc_sort_key(r):
    """Порядок строк перечня — по позиции в контроллере, как в выпущенных.

    Раньше строки шли в порядке расположения на листе, и модули в перечне
    перемежались: 3.1.1, 4.1.1, 5.1.1, 3.1.2… Сортируем по номерам модуля и
    канала; чертёжная приставка контура на порядок не влияет.
    """
    mod = re.sub(r"^[A-Za-zА-Яа-я]", "", str(r.get("mod", "")))
    nums = []
    for part in mod.split("."):
        try:
            nums.append(int(part))
        except ValueError:
            nums.append(0)
    return (nums, r.get("ch", 0), r.get("kc", ""))


# Вид арматуры в описании. Порядок важен: длинные основы идут первыми.
_VALVE_KIND = ("запорно-регулирующ", "регулирующ", "отсечн")


def _valve_key(desc):
    """Ключ «вид арматуры и место» из описания или None.

    Действие («контроль положения», «управление») отбрасываем — именно им
    описания пары и отличаются.
    """
    low = _norm(desc)
    kind = next((k for k in _VALVE_KIND if k in low), "")
    parts = re.split(r"\.\s+", (desc or "").strip(), maxsplit=1)
    place = _norm(parts[1]) if len(parts) > 1 else ""
    return (kind, place) if kind and place else None


def _fill_valve_tags(sections, log=print):
    """Дописывает позицию каналам контроля положения арматуры.

    Берём её с канала управления тем же клапаном. Если на такой ключ нашлось
    несколько разных позиций, не угадываем и оставляем графу пустой.
    """
    by_key = collections.defaultdict(set)
    for _cab, rows in sections:
        for r in rows:
            if not r.get("tag") or "правлени" not in _norm(r.get("desc", "")):
                continue
            k = _valve_key(r["desc"])
            if k:
                by_key[k].add(r["tag"])
    n = 0
    for _cab, rows in sections:
        for r in rows:
            if r.get("tag") or "положени" not in _norm(r.get("desc", "")):
                continue
            k = _valve_key(r.get("desc", ""))
            tags = by_key.get(k) if k else None
            if tags and len(tags) == 1:
                r["tag"] = next(iter(tags))
                n += 1
    if n:
        log(f"  позиция арматуры взята с канала управления тем же клапаном: {n}")
    return n


def read_sections(drawings, oda_exe=None, log=print, ext_tags=None):
    """Читает чертежи (DWG конвертируются) -> [(имя_раздела, rows)].
    drawings: пути или пары (путь, имя_раздела)."""
    items = [(d, None) if isinstance(d, str) else (d[0], d[1]) for d in drawings]
    dwgs = [p for p, _n in items if p.lower().endswith(".dwg")]
    conv = {}
    if dwgs:
        engine, exe = ("oda", oda_exe) if oda_exe else find_converter()
        conv = dwg_to_dxf(dwgs, engine=engine, exe=exe, log=log)
        missing = [p for p in dwgs if p not in conv]
        if missing:
            raise RuntimeError("Не сконвертировались: " + ", ".join(os.path.basename(m) for m in missing))
    sections = []
    for p, name in items:
        log(f"Чтение: {os.path.basename(p)}")
        if p.lower().endswith(".xlsx"):
            for cab, rows in sections_from_xlsx(p, log):
                sections.append((name or cab, rows))
            continue
        dxf = conv.get(p, p)
        rows = extract(dxf, log, ext_tags=ext_tags)
        for r in rows:
            r["src"] = p
        seen = {}
        for r in rows:
            if r["kc"] in seen:
                log(f"  ВНИМАНИЕ: позиция {r['kc']} встречается дважды")
            seen[r["kc"]] = True
        cab = name or cab_name_from_file(p)
        for r in rows:
            # приставка контура сохраняется в номере модуля, если она не
            # отброшена настройкой; иначе смотрим на имя шкафа
            pref = re.match(r"^[A-Za-zА-Яа-я]", str(r.get("mod", "")))
            r["ctrl"] = control_level(cab, pref.group(0) if pref else "",
                                      r.get("ctrl", ""))
        sections.append((cab, sorted(rows, key=_kc_sort_key)))
    _fill_valve_tags(sections, log)
    return sections

def run(drawings, template_in, template_out, out_dir, oda_exe=None, log=print,
        update_fields=False, sections=None, make_pdf=False, ext_tags=None):
    """drawings: список путей DWG/DXF или пар (путь, имя_раздела) в порядке разделов.
    sections — готовые данные (например, правленные в предпросмотре); если
    переданы, чертежи не читаются повторно.

    Шаблоны можно не указывать: тогда берутся встроенные из templates/,
    оформленные по образцу выпущенного комплекта."""
    if sections is None:
        sections = read_sections(drawings, oda_exe=oda_exe, log=log,
                                 ext_tags=ext_tags)
    os.makedirs(out_dir, exist_ok=True)
    if not template_in and not template_out:
        template_in = default_template("in")
        template_out = default_template("out")
        if template_in or template_out:
            log("Шаблоны не указаны — использую встроенные из папки templates")
        else:
            raise RuntimeError(
                "Не найдены шаблоны перечней.\n"
                "Ожидались файлы в папке templates рядом с программой:\n"
                "  Перечень входных сигналов.docx\n"
                "  Перечень выходных сигналов.docx\n"
                "Либо укажите свой шаблон на странице «Сборка».")

    def _target(template):
        """Куда писать результат, не затерев сам шаблон.

        Если папка результата совпадает с папкой шаблона, файл назывался бы
        так же, как шаблон, и перезаписал бы его. Такой шаблон обычно один
        на весь отдел, поэтому вместо перезаписи добавляем суффикс.
        """
        out = os.path.join(out_dir, os.path.basename(template))
        if os.path.abspath(out) == os.path.abspath(template):
            stem, ext = os.path.splitext(os.path.basename(template))
            out = os.path.join(out_dir, f"{stem} (заполненный){ext}")
            log(f"  ⚠ папка результата совпадает с папкой шаблона —")
            log(f"    чтобы не затереть шаблон, файл назван: {os.path.basename(out)}")
        return out

    results = []
    if template_in:
        out = _target(template_in)
        log("Сборка перечня входных сигналов...")
        build_docx(template_in, out, "in", "ВХОДНЫЕ", sections, log)
        results.append(out)
    if template_out:
        out = _target(template_out)
        log("Сборка перечня выходных сигналов...")
        build_docx(template_out, out, "out", "ВЫХОДНЫЕ", sections, log)
        results.append(out)
    if update_fields and results:
        log("Обновление номеров страниц через Word...")
        if update_fields_word(results, log, pdf=make_pdf):
            log("ГОТОВО.")
            return results
    log("ГОТОВО. При открытии в Word ответьте «Да» на обновление полей")
    log("(или Ctrl+A, F9) — проставятся номера страниц в содержании.")
    return results
