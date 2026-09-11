# -*- coding: utf-8 -*-
"""Полная проверка программы.

Запуск:  python qa_tests.py
         python qa_tests.py --fast     без печати PDF (быстрее)

В отличие от selftest.py, который проверяет мелкие функции за секунду, здесь
прогоняются целые сценарии: чертёж собирается синтетически прямо в памяти,
из него делаются перечни, сверяются, печатаются в PDF, выгружаются в Excel.
Ничего из рабочих файлов не требуется и не изменяется — всё во временной папке.
"""
import os
import re
import shutil
import sys
import tempfile
import traceback

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import perechni_core as core

PASS, FAIL, SKIP = [], [], []
QUIET = lambda *a, **k: None


def check(name, got, want):
    if got == want:
        PASS.append(name)
        print("  ок    " + name)
    else:
        FAIL.append("%s: получено %r, ожидалось %r" % (name, got, want))
        print("  ПЛОХО %s — получено %r, ожидалось %r" % (name, got, want))


def truthy(name, got, note=""):
    if got:
        PASS.append(name)
        print("  ок    " + name + (("  — " + str(note)) if note else ""))
    else:
        FAIL.append(name)
        print("  ПЛОХО " + name)


def section(t):
    print("\n" + t)
    print("-" * len(t))


# ---------------------------------------------------------------- фикстуры

def make_dxf(path, modules=2, chans=4, with_biz=True, with_di=True, frame=True):
    """Синтетическая схема подключения: рамка А4, модули, описания, якоря.

    Собирается так же, как устроены реальные чертежи: заголовок «Модуль X.Y AI»,
    описания на слое «Текст», метки БИЗ на «mark2», якоря реле на своих слоях.
    """
    import ezdxf
    doc = ezdxf.new(setup=True)
    doc.styles.add("GOSTW", font="gostw.shx")
    msp = doc.modelspace()
    if frame:
        # рамка А4 книжная (210x297) в точке (0,0)
        msp.add_lwpolyline([(0, 0), (210, 0), (210, 297), (0, 297)], close=True)
        msp.add_text("Инв.N подл.", dxfattribs={"style": "GOSTW",
                                                "insert": (2, 10)})
    x = 300.0
    for m in range(1, modules + 1):
        msp.add_text("Модуль 1.%d AI 8AIx4...20mA" % m,
                     dxfattribs={"insert": (x, 500)})
        for c in range(1, chans + 1):
            y = 480.0 - c * 10
            msp.add_text("PT-%d%d +" % (m, c),
                         dxfattribs={"insert": (x - 35, y), "layer": "Текст"})
            mt = msp.add_mtext("Контроль давления в трубопроводе узла %d.%d" % (m, c),
                               dxfattribs={"layer": "Текст"})
            mt.set_location((x, y))
        if with_biz:
            msp.add_text("БИЗ-1", dxfattribs={"insert": (x + 5, 495), "layer": "mark2"})
        x += 230.0
    if with_di:
        for c in range(1, 5):
            y = 480.0 - c * 10
            msp.add_text("1.9-KLDI%d" % c,
                         dxfattribs={"insert": (x, y), "layer": "Реле DI (1-KL)"})
            mt = msp.add_mtext("Положение отсечного клапана поз. HV-%d" % c,
                               dxfattribs={"layer": "Текст"})
            mt.set_location((x + 20, y))
    doc.saveas(path)
    return path


# ------------------------------------------------------------------ сценарии

def t_fixture(tmp):
    section("Синтетический чертёж читается")
    p = make_dxf(os.path.join(tmp, "synt.dxf"))
    truthy("чертёж создан", os.path.exists(p), "%.0f КБ" % (os.path.getsize(p) / 1e3))
    rows = core.extract(p, log=QUIET)
    truthy("каналы извлечены", len(rows) > 0, "каналов: %d" % len(rows))
    types = {r["type"] for r in rows}
    truthy("есть аналоговые", "AI" in types, sorted(types))
    truthy("есть дискретные", "DI" in types, sorted(types))
    kcs = [r["kc"] for r in rows]
    check("нет задвоенных позиций", len(kcs), len(set(kcs)))
    tagged = [r for r in rows if r["tag"]]
    truthy("позиции распознаны", len(tagged) > 0,
           "%d из %d" % (len(tagged), len(rows)))
    biz = {r["ex"] for r in rows if r["type"] == "AI"}
    truthy("взрывозащита проставлена", biz and biz <= {"Exia", "Exd"}, biz)
    return p, rows


def t_sheets(tmp, p):
    section("Рамки листов")
    doc, frames = core.detect_sheets(p, log=QUIET)
    check("найден один лист", len(frames), 1)
    if frames:
        check("формат А4 книжная", (frames[0][4], frames[0][5]), ("А4", "книжная"))
    truthy("сводка формируется", bool(core.sheets_summary(frames)),
           core.sheets_summary(frames))
    # чертёж без рамки
    p2 = make_dxf(os.path.join(tmp, "noframe.dxf"), frame=False)
    _d, fr2 = core.detect_sheets(p2, log=QUIET)
    check("без рамки — ноль листов", len(fr2), 0)


def t_stamp(tmp, p):
    section("Знак номера в штампе")
    doc = core.read_dxf(p)
    st = core._fix_shx_styles(doc)
    n = core._fix_stamp_text(doc, st)
    check("надпись штампа исправлена", n, 1)
    got = [e.dxf.text for e in doc.modelspace()
           if e.dxftype() == "TEXT" and "подл" in e.dxf.get("text", "")]
    check("получилось «Инв.№ подл.»", got[0] if got else "", "Инв.№ подл.")


def t_build_and_verify(tmp, rows):
    section("Сборка перечней и сверка")
    out = os.path.join(tmp, "out")
    secs = [("ШКАФ ПРОВЕРКА", rows)]
    built = core.run([], "", "", out, sections=secs, log=QUIET)
    check("собрано документов", len(built), 2)
    for b in built:
        truthy("файл создан: " + os.path.basename(b),
               os.path.exists(b) and os.path.getsize(b) > 10000,
               "%.0f КБ" % (os.path.getsize(b) / 1e3))
    res = core.verify_perechen(secs, built, log=QUIET)
    truthy("сверка пройдена", res["ok"],
           "сошлось %d из %d" % (res["counts"]["совпало"], res["counts"]["в_чертеже"]))
    check("расхождений нет", len(res["issues"]), 0)
    return built, secs, res


def t_verify_negative(tmp, rows):
    section("Сверка ловит порчу документа")
    secs = [("ШКАФ ПРОВЕРКА", rows)]
    orig = core.parse_perechen
    try:
        good = [dict(cab="", io=r["io"], tag=r["tag"], desc=r["desc"],
                     level=r["level"], ctrl=r["ctrl"], ex=r["ex"], kc=r["kc"])
                for r in rows]
        # строку удалили
        core.parse_perechen = lambda _p: good[1:]
        r1 = core.verify_perechen(secs, ["x.docx"], log=QUIET)
        truthy("замечена пропавшая строка",
               any(i[0] == "нет в документе" for i in r1["issues"]))
        # позицию подменили
        bad = [dict(g) for g in good]
        bad[0]["tag"] = "ПОДМЕНА"
        core.parse_perechen = lambda _p: bad
        r2 = core.verify_perechen(secs, ["x.docx"], log=QUIET)
        truthy("замечена подменённая позиция",
               any(i[0].startswith("расходится: позиция") for i in r2["issues"]))
        # лишняя строка
        extra = good + [dict(good[0], kc="9.9.9", tag="ЛИШНЯЯ")]
        core.parse_perechen = lambda _p: extra
        r3 = core.verify_perechen(secs, ["x.docx"], log=QUIET)
        truthy("замечена лишняя строка",
               any(i[0] == "нет в чертеже" for i in r3["issues"]))
        # порча уровня сигнала
        lvl = [dict(g) for g in good]
        lvl[0]["level"] = "=24В"
        core.parse_perechen = lambda _p: lvl
        r4 = core.verify_perechen(secs, ["x.docx"], log=QUIET)
        truthy("замечен подменённый уровень",
               any("уровень" in i[0] for i in r4["issues"]))
    finally:
        core.parse_perechen = orig


def _doc_rows_like(rows, cab):
    """Строки, какими их читает parse_perechen из готового перечня."""
    return [dict(cab=cab, io=r["io"], tag=r["tag"], desc=r["desc"],
                 level=r["level"], ctrl=r["ctrl"], ex=r["ex"], kc=r["kc"])
            for r in rows]


def t_verify_cab_names(tmp, rows):
    section("Сверка: шкаф в перечне назван иначе, чем чертёж")
    secs = [("ШКАФ ШСК1 РСУ", rows)]
    orig = core.parse_perechen
    try:
        # Перечень ведёт все шкафы одним разделом «ШКАФ РСУ»: по ключу со
        # шкафом не совпадало ничего, и весь перечень выглядел дефектным.
        core.parse_perechen = lambda _p: _doc_rows_like(rows, "ШКАФ РСУ")
        r1 = core.verify_perechen(secs, ["x.docx"], log=QUIET)
        check("чужое имя шкафа не мешает сверке",
              r1["counts"]["совпало"], len(rows))
        # Там, где имена сходятся, шкаф из ключа не выпадает: два шкафа с
        # одинаковыми номерами каналов не должны слиться в один.
        core.parse_perechen = lambda _p: (_doc_rows_like(rows, "ШКАФ ШСК1 РСУ")
                                          + _doc_rows_like(rows, "ШКАФ ШСК2 РСУ"))
        two = [("ШКАФ ШСК1 РСУ", rows), ("ШКАФ ШСК2 РСУ", rows)]
        r2 = core.verify_perechen(two, ["x.docx"], log=QUIET)
        check("одинаковые каналы разных шкафов не слиплись",
              r2["counts"]["совпало"], 2 * len(rows))
    finally:
        core.parse_perechen = orig


def t_perechen_six_columns(tmp, rows):
    section("Перечень без графы «№ п/п»")
    from docx import Document
    out = os.path.join(tmp, "six")
    built = core.run([], "", "", out, sections=[("ШКАФ ПРОВЕРКА", rows)], log=QUIET)
    src = built[0]
    dst = os.path.join(out, "без номера " + os.path.basename(src))
    doc = Document(src)
    for t in doc.tables:
        for tr in t._tbl.findall(core.qn("w:tr")):
            tcs = tr.findall(core.qn("w:tc"))
            if not tcs:
                continue
            pr = tcs[0].find(core.qn("w:tcPr"))
            gs = pr.find(core.qn("w:gridSpan")) if pr is not None else None
            span = int(gs.get(core.qn("w:val"))) if gs is not None else 1
            if span > 1:                      # название таблицы во всю ширину
                gs.set(core.qn("w:val"), str(span - 1))
            else:
                tr.remove(tcs[0])
    doc.save(dst)
    n7 = len(core.parse_perechen(src))
    n6 = len(core.parse_perechen(dst))
    check("строк прочитано столько же, сколько в семи графах", n6, n7)
    truthy("строки не пустые", n6 > 0, "строк: %d" % n6)
    a = core.parse_perechen(src)[0]
    b = core.parse_perechen(dst)[0]
    check("позиция в контроллере на месте", b["kc"], a["kc"])
    check("позиция по проекту на месте", b["tag"], a["tag"])
    check("описание на месте", b["desc"], a["desc"])


def t_error_box():
    section("Ошибка из фонового потока доходит до диалога")
    try:
        import types
        import perechni_gui as gui
    except Exception as e:
        SKIP.append("диалог ошибки (%s)" % type(e).__name__)
        print("  пропуск: перечни_gui не импортируется (%s)" % type(e).__name__)
        return
    shown = []
    orig = gui.messagebox
    gui.messagebox = types.SimpleNamespace(showerror=lambda t, m: shown.append((t, m)))
    try:
        queue = []
        stub = type("Stub", (), dict(after=lambda self, _ms, fn, *a: queue.append((fn, a)),
                                     error_box=gui.App.error_box))()
        try:
            raise RuntimeError("конвертер не найден")
        except Exception as e:
            stub.error_box("Ошибка", e)
        # очередь Tk разбирается позже: к этому моменту имя из except уже удалено
        for fn, a in queue:
            fn(*a)
        check("в диалоге текст ошибки, а не NameError", shown,
              [("Ошибка", "конвертер не найден")])
    finally:
        gui.messagebox = orig


def t_pdf_monochrome():
    section("Печать листов: монохром, как в AutoCAD")
    from ezdxf.addons.drawing.config import ColorPolicy, LineweightPolicy
    cfg = core._pdf_config()
    check("цвет: всё чёрным", cfg.color_policy, ColorPolicy.BLACK)
    check("толщины линий берутся из чертежа",
          cfg.lineweight_policy, LineweightPolicy.ABSOLUTE)
    truthy("миллиметры переведены в пункты",
           abs(cfg.lineweight_scaling - 1 / 0.3527) < 0.01,
           "множитель %.3f" % cfg.lineweight_scaling)
    col = core._pdf_config(color=True)
    check("цветная печать осталась доступной",
          col.color_policy, ColorPolicy.COLOR_SWAP_BW)
    check("толщины линий и в цвете те же",
          col.lineweight_policy, LineweightPolicy.ABSOLUTE)


def t_draw_resilience():
    section("Один битый объект не рвёт лист")

    class FakeFrontend:
        """Падает на «таблице» — так ведёт себя ezdxf на ACAD_TABLE с SHX."""
        def __init__(self):
            self.drawn = []

        def draw_entities(self, ents):
            for e in ents:
                if e.dxftype() == "ACAD_TABLE":
                    raise RuntimeError("invalid last entry in index table")
                self.drawn.append(e)

    class E:
        def __init__(self, t):
            self._t = t
        def dxftype(self):
            return self._t

    ents = [E("LINE"), E("TEXT"), E("ACAD_TABLE"), E("LINE"), E("MTEXT")]
    fe = FakeFrontend()
    bad = core._draw_entities_safe(fe, ents, log=QUIET)
    check("пропущен ровно один объект", bad, 1)
    check("остальные нарисованы", len(fe.drawn), 4)


def t_template_guard(tmp):
    section("Шаблон не затирается результатом")
    d = os.path.join(tmp, "tpl")
    os.makedirs(d, exist_ok=True)
    tpl = os.path.join(d, "Перечень.docx")
    shutil.copy2(core.default_template("in"), tpl)
    before = os.path.getsize(tpl)
    rows = [dict(sortx=0, x=0, y=0, mod="1.1", type="AI", io="in", ch=1,
                 tag="PT-1", desc="Проверка", level="4-20 мА", ctrl="",
                 ex="Exd", kc="1.1.1")]
    core.run([], tpl, None, d, sections=[("ШКАФ", rows)], log=QUIET)
    check("шаблон не изменился", os.path.getsize(tpl), before)
    truthy("результат сохранён отдельно",
           os.path.exists(os.path.join(d, "Перечень (заполненный).docx")))


def t_exports(tmp, secs, res):
    section("Выгрузки")
    rows = secs[0][1]
    p1 = os.path.join(tmp, "signals.xlsx")
    core.export_xlsx(secs, p1)
    truthy("Excel-сводка сигналов", os.path.getsize(p1) > 3000,
           "%.0f КБ" % (os.path.getsize(p1) / 1e3))
    p2 = os.path.join(tmp, "verify.xlsx")
    core.export_verify_xlsx(res, p2)
    truthy("отчёт сверки", os.path.getsize(p2) > 3000,
           "%.0f КБ" % (os.path.getsize(p2) / 1e3))
    rep = core.compare(secs, [], log=QUIET)
    p3 = os.path.join(tmp, "compare.xlsx")
    core.export_compare_xlsx(rep, p3)
    truthy("отчёт сравнения версий", os.path.getsize(p3) > 3000,
           "%.0f КБ" % (os.path.getsize(p3) / 1e3))
    # проверяем, что файлы читаются обратно
    try:
        import openpyxl
        wb = openpyxl.load_workbook(p1)
        truthy("сводка открывается", len(wb.sheetnames) >= 1, wb.sheetnames)
        wb.close()
    except Exception as e:
        FAIL.append("сводка не открывается: %s" % e)
        print("  ПЛОХО сводка не открывается:", e)


def t_dashboard(tmp, secs):
    section("Данные дашборда и нормоконтроль")
    rows = secs[0][1]

    # Проверять на «чистых» данных бессмысленно: такой тест пройдёт и у функции,
    # которая всегда возвращает пустой список. Готовим данные с изъянами.
    spoiled = [dict(r) for r in rows]
    spoiled.append(dict(spoiled[0], kc="1.1.99", ch=99, desc="Резерв", tag=""))
    spoiled.append(dict(spoiled[0], kc="1.1.98", ch=98, desc="Резерв", tag="PT-XX"))
    spoiled.append(dict(spoiled[0], kc="1.1.97", ch=97, tag=spoiled[0]["tag"],
                        desc="Совсем другое описание того же тега"))
    bad = [("ШКАФ ПРОВЕРКА", spoiled)]

    free = core.free_channels(bad)
    truthy("свободные каналы находятся", len(free) > 0, "строк: %d" % len(free))

    rep = core.checks_report(bad)
    truthy("нормоконтроль что-то нашёл", len(rep) > 0, "замечаний: %d" % len(rep))
    text = " | ".join(str(x) for x in rep).lower()
    truthy("замечен «Резерв» с проставленной позицией", "резерв" in text)
    truthy("замечен один тег с разными описаниями",
           "тег" in text or "описан" in text or "PT-" in " ".join(str(x) for x in rep))

    # на чистых данных замечаний быть не должно
    clean = core.checks_report(secs)
    check("на чистых данных замечаний нет", len(clean), 0)


def t_equipment_and_pz(tmp, p, secs):
    section("Состав шкафа, спецификация и заготовка ПЗ")
    try:
        items = core.extract_equipment(p)
        truthy("состав шкафа читается", isinstance(items, list),
               "позиций: %d" % len(items))
    except Exception as e:
        FAIL.append("extract_equipment падает: %s" % e)
        print("  ПЛОХО extract_equipment падает:", e)
        items = []
    equip = [("ШКАФ ПРОВЕРКА", items)]
    p1 = os.path.join(tmp, "spec.xlsx")
    try:
        core.export_spec_xlsx(equip, secs, p1)
        truthy("спецификация выгружается", os.path.getsize(p1) > 3000,
               "%.0f КБ" % (os.path.getsize(p1) / 1e3))
    except Exception as e:
        FAIL.append("export_spec_xlsx падает: %s" % e)
        print("  ПЛОХО export_spec_xlsx падает:", e)
    p2 = os.path.join(tmp, "pz.docx")
    try:
        core.export_pz_docx(secs, equip, p2)
        truthy("заготовка ПЗ создаётся", os.path.getsize(p2) > 5000,
               "%.0f КБ" % (os.path.getsize(p2) / 1e3))
        from docx import Document
        d = Document(p2)
        txt = " ".join(x.text for x in d.paragraphs)
        truthy("в ПЗ попали данные", len(txt.strip()) > 50, "%d знаков" % len(txt))
    except Exception as e:
        FAIL.append("export_pz_docx падает: %s" % e)
        print("  ПЛОХО export_pz_docx падает:", e)


def t_stamps(tmp):
    section("Замена значений в штампах")
    d = os.path.join(tmp, "stamps")
    os.makedirs(d, exist_ok=True)
    doc = os.path.join(d, "Перечень.docx")
    shutil.copy2(core.default_template("in"), doc)
    # Во встроенных шаблонах даты намеренно заменены на «sss», поэтому
    # шаблон по умолчанию под шаблон даты не подходит — проверяем и это.
    val = core.detect_stamp_value(doc)
    truthy("в шаблоне даты нет (стоит «sss»)", not val, repr(val))
    val = core.detect_stamp_value(doc, pattern=r"^sss$")
    truthy("заглушка «sss» находится своим шаблоном", val == "sss", repr(val))
    if val:
        core.replace_in_stamps([doc], val, "07.26", log=QUIET, backup=False)
        check("заглушка заменилась на дату", core.detect_stamp_value(doc), "07.26")
        # обратная замена — документ не должен портиться
        core.replace_in_stamps([doc], "07.26", val, log=QUIET, backup=False)
        check("вернулась заглушка", core.detect_stamp_value(doc, pattern=r"^sss$"), val)
        from docx import Document
        truthy("документ остался читаемым", len(Document(doc).tables) > 0)


def t_multisection(tmp, rows):
    section("Несколько шкафов в одном перечне")
    a = [dict(r) for r in rows]
    b = [dict(r, kc="2." + r["kc"].split(".", 1)[1], mod="2." + r["mod"].split(".", 1)[1])
         for r in rows]
    secs = [("ШКАФ ШСК1 РСУ", a), ("ЩИТ 2Щ", b)]
    out = os.path.join(tmp, "multi")
    # шаблоны не указаны -> берутся встроенные и собираются оба перечня
    built = core.run([], "", "", out, sections=secs, log=QUIET)
    truthy("документы собраны", len(built) == 2,
           ", ".join(os.path.basename(x) for x in built))
    from docx import Document
    doc = Document(built[0])
    toc = [r.cells[0].text.strip() for r in doc.tables[0].rows]
    truthy("в содержании два раздела",
           any(t.startswith("1 ") for t in toc) and any(t.startswith("2 ") for t in toc),
           " / ".join(t[:22] for t in toc[:6]))
    truthy("регистр обозначений сохранён",
           any("ШСК1 РСУ" in t for t in toc) and any("2Щ" in t for t in toc))
    truthy("нет «шск1 рсу» строчными", not any("шск1 рсу" in t for t in toc))
    res = core.verify_perechen(secs, built, log=QUIET)
    truthy("сверка двух шкафов проходит", res["ok"],
           "сошлось %d" % res["counts"]["совпало"])


def t_no_empty_sections(tmp, rows):
    section("Пустых разделов в перечне нет")
    # Шкаф без единого дискретного сигнала: раздел «Дискретные» не должен
    # появиться ни в теле, ни в содержании, а нумерация — сдвинуться.
    only_an = [dict(r) for r in rows if r["type"] in ("AI", "WI")]
    truthy("в наборе есть только аналоговые", bool(only_an), "%d строк" % len(only_an))
    secs = [("ШКАФ ШСК1 РСУ", only_an)]
    out = os.path.join(tmp, "empty")
    built = core.run([], "", "", out, sections=secs, log=QUIET)
    from docx import Document
    doc = Document(built[0])
    heads = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    toc = [r.cells[0].text.strip() for r in doc.tables[0].rows]
    check("дискретного раздела нет в теле",
          any("ДИСКРЕТНЫЕ" in h for h in heads), False)
    check("дискретного раздела нет в содержании",
          any("искретные" in t for t in toc), False)
    truthy("аналоговый раздел на месте",
           any("1.1 АНАЛОГОВЫЕ" in h for h in heads),
           " / ".join(h[:26] for h in heads[:4]))
    # ни одна таблица перечня не должна остаться без строк данных
    empty = [t.rows[0].cells[0].text.strip()[:40] for t in doc.tables[1:-1]
             if len(t.rows) <= 2]
    check("таблиц без строк нет", empty, [])


def t_pagination(tmp, rows):
    section("Разбивка на листы")
    from docx import Document
    from docx.oxml.ns import qn
    a = [dict(r) for r in rows]
    b = [dict(r, kc="2." + r["kc"].split(".", 1)[1], mod="2." + r["mod"].split(".", 1)[1])
         for r in rows]
    out = os.path.join(tmp, "pages")
    built = core.run([], "", "", out, sections=[("ШКАФ А", a), ("ШКАФ Б", b)],
                     log=QUIET)
    doc = Document(built[0])
    body = doc.element.body
    # Разрыв делается свойством абзаца. Отдельный абзац с разрывом на стыке,
    # где таблица кончилась ровно внизу листа, давал лишний пустой лист.
    brs = [core._ptext(p).strip()[:30] for p in body.findall(qn("w:p"))
           if any(x.get(qn("w:type")) == "page" for x in p.iter(qn("w:br")))]
    check("отдельных абзацев с разрывом нет", brs, [])
    heads = [p for p in body.findall(qn("w:p"))
             if re.match(r"^\d+(\.\d+)? \S", core._ptext(p).strip())]
    truthy("заголовки найдены", len(heads) >= 4, "%d шт." % len(heads))
    first, rest = heads[0], heads[1:]

    def brk(p):
        pPr = p.find(qn("w:pPr"))
        return pPr is not None and pPr.find(qn("w:pageBreakBefore")) is not None

    check("первый заголовок без разрыва", brk(first), False)
    # Первый подраздел шкафа идёт сразу под заголовком шкафа — свой разрыв ему
    # не нужен, иначе заголовок шкафа остался бы на листе один.
    sub1 = lambda p: re.match(r"^\d+\.1 ", core._ptext(p).strip())
    check("подраздел «N.1» без своего разрыва",
          [core._ptext(p).strip()[:18] for p in rest if sub1(p) and brk(p)], [])
    check("остальные заголовки с новой страницы",
          [core._ptext(p).strip()[:18] for p in rest
           if not sub1(p) and not brk(p)], [])
    # «Не отрывать от следующего» на каждой строке заставляло Word держать
    # вместе всю таблицу — она уезжала на новый лист, а над ней оставался
    # лист с одним заголовком.
    bad = 0
    for t in body.findall(qn("w:tbl"))[1:]:
        for tr in t.findall(qn("w:tr"))[2:]:
            for p in tr.iter(qn("w:p")):
                pPr = p.find(qn("w:pPr"))
                if pPr is not None and pPr.find(qn("w:keepNext")) is not None:
                    bad += 1
    check("строки данных без «не отрывать»", bad, 0)


def t_naming():
    section("Имена шкафов и заголовки")
    check("латиница в имени файла нормализуется",
          core.cab_name_from_file("08б ... ШCК1 РСУ.dwg").startswith("ШКАФ"), True)
    check("щит опознаётся",
          core.cab_name_from_file("2Щ.dwg"), "ЩИТ 2Щ")
    check("регистр обозначения сохраняется",
          core._cab_title("ШКАФ ШСК1 РСУ"), "Шкаф ШСК1 РСУ")


def t_environment():
    section("Готовность среды")
    rows = core.readiness()
    check("проверяется четыре пункта", len(rows), 4)
    for ok_, name, detail, cons in rows:
        print("     %s %-18s %s" % ("✓" if ok_ else "✗", name, detail))
        if not ok_:
            truthy("для «%s» объяснены последствия" % name, bool(cons))
    truthy("шаблоны найдены", bool(core.default_template("in")))


def t_config(tmp):
    section("Настройки и словарь")
    s = core.extract_settings()
    check("ключей настройки", len(s), len(core.EXTRACT_DEFAULTS))
    truthy("словарь сокращений загружается", isinstance(core.load_abbrev(), list),
           "%d пар" % len(core.load_abbrev()))
    check("сокращение применяется",
          core.apply_abbrev("Контроль давления", [("Контроль давления", "Контр. давления")]),
          "Контр. давления")


def t_pdf(tmp, p):
    section("Печать чертежа в PDF")
    out = os.path.join(tmp, "sheets.pdf")
    res, frames = core.export_sheets_pdf(p, out, log=QUIET)
    truthy("PDF создан", bool(res) and os.path.exists(out),
           "%.0f КБ, листов %d" % (os.path.getsize(out) / 1e3 if res else 0, len(frames)))
    try:
        import pymupdf
        d = pymupdf.open(out)
        check("страниц в PDF", d.page_count, len(frames))
        w = d[0].rect.width / 72 * 25.4
        truthy("размер страницы А4 по ширине", abs(w - 210) < 2, "%.0f мм" % w)
        d.close()
    except ImportError:
        SKIP.append("проверка страниц PDF (нет pymupdf)")
        print("  пропуск  проверка страниц PDF — не установлен pymupdf")


def t_revisions(tmp):
    section("Папки-ревизии")
    import datetime
    d = os.path.join(tmp, "rev")
    os.makedirs(d, exist_ok=True)
    day = datetime.date(2026, 8, 20)
    made = []
    for i in range(3):
        r = core.next_revision_dir(d, day)
        os.makedirs(r)
        open(os.path.join(r, "x.txt"), "w").write(str(i))
        made.append(r)
    check("номера идут по порядку",
          [os.path.basename(x)[:10] for x in made],
          ["Ревизия 01", "Ревизия 02", "Ревизия 03"])
    # Удаление ранней ревизии раньше приводило к повтору номера и молчаливой
    # записи поверх существующей папки — самое опасное, что тут может быть.
    shutil.rmtree(made[0])
    nxt = core.next_revision_dir(d, day)
    check("после удаления номер не повторяется", os.path.basename(nxt)[:10], "Ревизия 04")
    check("следующая папка ещё не существует", os.path.exists(nxt), False)
    check("прежняя ревизия цела", open(os.path.join(made[2], "x.txt")).read(), "2")
    check("на пустом месте — первая", os.path.basename(
        core.next_revision_dir(os.path.join(tmp, "пусто"), day))[:10], "Ревизия 01")


def t_manifest(tmp, secs, res):
    section("Паспорт выпуска")
    src = os.path.join(tmp, "чертёж.dxf")
    open(src, "wb").write(b"x" * 1000)
    out = os.path.join(tmp, "док.docx")
    open(out, "wb").write(b"y" * 2000)
    size, mt, md5 = core.file_fingerprint(src)
    check("размер файла определён", size, 1000)
    truthy("md5 посчитан", len(md5) == 10, md5)
    truthy("дата изменения есть", bool(mt), mt)
    check("отсутствующий файл не роняет", core.file_fingerprint(
        os.path.join(tmp, "нет.dxf")), (0, "", ""))

    t = core.manifest_text(sources=[("ШКАФ", src)], sections=secs, out_files=[out],
                           verify=res, checks=["замечание раз"], app_version="9.9")
    for must in ("ПАСПОРТ ВЫПУСКА", "СВЕРКА С ЧЕРТЕЖАМИ", "ИСХОДНЫЕ ЧЕРТЕЖИ",
                 "ВЫПУЩЕННЫЕ ДОКУМЕНТЫ", "md5", "замечание раз", "v9.9"):
        truthy("в паспорте есть «%s»" % must, must in t)
    truthy("вердикт сверки записан", "ПРОЙДЕНА" in t)

    # сверка не прошла — это должно быть видно сразу
    bad = dict(ok=False, counts={"совпало": 5, "в_чертеже": 9},
               issues=[("нет в документе", "1.1.1", "in", "PT1 | Описание", "")])
    t2 = core.manifest_text(sources=[], sections=secs, out_files=[], verify=bad)
    truthy("непройденная сверка отражена", "НЕ ПРОЙДЕНА" in t2)
    truthy("расхождение перечислено", "1.1.1" in t2)
    t3 = core.manifest_text(sections=secs)
    truthy("без сверки так и написано", "не выполнялась" in t3)

    f = core.write_manifest(os.path.join(tmp, "Паспорт.txt"), sections=secs, verify=res)
    truthy("паспорт записан на диск", os.path.getsize(f) > 100,
           "%d байт" % os.path.getsize(f))


def t_module_types(tmp):
    section("Модули без каналов ввода-вывода")
    import ezdxf
    p = os.path.join(tmp, "cpu.dxf")
    doc = ezdxf.new(setup=True)
    msp = doc.modelspace()
    # процессорный модуль: каналов у него нет
    msp.add_text("Модуль 1.0 CPU", dxfattribs={"insert": (300, 500)})
    # обычный аналоговый — рядом, чтобы было с чем сравнить
    msp.add_text("Модуль 1.1 AI 8AIx4...20mA", dxfattribs={"insert": (530, 500)})
    for c in range(1, 4):
        mt = msp.add_mtext("Контроль давления в узле %d" % c, dxfattribs={"layer": "Текст"})
        mt.set_location((530, 480.0 - c * 10))
    doc.saveas(p)
    rows = core.extract(p, log=QUIET)
    types = {r["type"] for r in rows}
    truthy("каналы аналогового модуля найдены", "AI" in types, sorted(types))
    check("процессорный модуль каналов не даёт", "CPU" in types, False)
    check("нет строк модуля 1.0", [r for r in rows if r["mod"] == "1.0"], [])


def t_robustness(tmp):
    section("Устойчивость к плохим данным")
    import ezdxf
    # пустой чертёж
    e = os.path.join(tmp, "empty.dxf")
    ezdxf.new().saveas(e)
    try:
        rows = core.extract(e, log=QUIET)
        check("пустой чертёж — ноль каналов", len(rows), 0)
    except Exception as ex:
        FAIL.append("пустой чертёж роняет extract: %s" % ex)
        print("  ПЛОХО пустой чертёж роняет extract:", ex)
    # несуществующий файл
    try:
        core.extract(os.path.join(tmp, "нет-такого.dxf"), log=QUIET)
        FAIL.append("несуществующий файл не вызвал ошибку")
        print("  ПЛОХО несуществующий файл не вызвал ошибку")
    except Exception:
        PASS.append("несуществующий файл даёт понятную ошибку")
        print("  ок    несуществующий файл даёт ошибку, а не тихий пустой ответ")
    # не тот формат
    junk = os.path.join(tmp, "junk.dxf")
    with open(junk, "w", encoding="utf-8") as f:
        f.write("это не чертёж")
    try:
        core.extract(junk, log=QUIET)
        print("  ок    мусорный файл обработан без падения")
        PASS.append("мусорный файл обработан")
    except Exception as ex:
        print("  ок    мусорный файл даёт ошибку: %s" % type(ex).__name__)
        PASS.append("мусорный файл даёт ошибку")
    # путь с пробелами и кириллицей
    d = os.path.join(tmp, "путь с пробелами и ёлкой")
    os.makedirs(d, exist_ok=True)
    p = make_dxf(os.path.join(d, "чертёж №1.dxf"))
    rows = core.extract(p, log=QUIET)
    truthy("путь с кириллицей и пробелами", len(rows) > 0, "каналов: %d" % len(rows))


def t_cab_naming_details():
    section("Имя шкафа из имени файла")
    n = core.cab_name_from_file
    check("уточнение РСУ сохраняется",
          n("Схема ... шкафа системно-кроссового ШСК1 РСУ.dxf"), "ШКАФ ШСК1 РСУ")
    check("уточнение ПАЗ сохраняется",
          n("Схема ... шкафа системно-кроссового zШСК1 ПАЗ.dxf"), "ШКАФ ШСК1 ПАЗ")
    check("два шкафа проекта не сливаются",
          n("...ШСК1 РСУ.dxf") != n("...zШСК1 ПАЗ.dxf"), True)
    check("шкаф без номера опознаётся",
          n("10б Чертеж общего вида шкафа бесперебойного питания ШИБП.dwg"),
          "ШКАФ ШИБП")
    check("латиница в уточнении переводится",
          n("Схема ШCК1 PCУ.dxf"), "ШКАФ ШСК1 РСУ")
    check("расширение файла уточнением не считается",
          n("ШСК1 PDF версия.dxf"), "ШКАФ ШСК1")
    check("слово из названия шкафом не становится",
          n("Обычный файл без обозначения.dxf"), "Обычный файл без обозначения")


def t_module_prefix(tmp):
    section("Чертёжная приставка у номера модуля")
    import re
    S = core.extract_settings()
    hdr = re.compile(S["module_header"])
    m = hdr.match("Модуль z1.5 DO 16DOх24 VDC")
    truthy("заголовок «Модуль z1.5 DO» разбирается", bool(m))
    if m:
        check("приставка в номер модуля не попадает",
              (m.group(1), m.group(2), m.group(3)), ("1", "5", "DO"))
    m2 = hdr.match("Модуль 1.5 DO 16DOх24 VDC")
    truthy("обычный заголовок по-прежнему разбирается", bool(m2))
    a = core.ANCHOR_RE.match("z1.5-KLDO3")
    truthy("якорь «z1.5-KLDO3» разбирается", bool(a))
    if a:
        check("номер модуля из якоря без приставки",
              (a.group(1), a.group(2), a.group(4)), ("1", "5", "3"))


def t_equipment_counts():
    section("Подсчёт оборудования в шкафу")
    check("безымянный блок в состав не идёт", core._is_equipment_block("23463246"), False)
    check("числовое имя блока в состав не идёт", core._is_equipment_block("123"), False)
    check("анонимный блок AutoCAD отсеивается", core._is_equipment_block("*U12"), False)
    check("штамп отсеивается", core._is_equipment_block("штамп_ЕСКД_л2"), False)
    check("настоящий блок проходит", core._is_equipment_block("AVIS12-RTD-I-C"), True)
    check("кириллический блок проходит", core._is_equipment_block("НПТ 2 2"), True)


def t_spec_docx(tmp):
    section("Спецификация по ГОСТ")
    tpl = core.default_template("spec")
    truthy("шаблон спецификации найден", bool(tpl))
    if not tpl:
        return
    secs = core.spec_sections_for_cabs(["ШКАФ ШСК1 РСУ", "ШКАФ ШСК1 ПАЗ", "ЩИТ 2Щ"])
    check("раздел один — «Шкафы»", [s[0] for s in secs], ["Шкафы"])
    check("позиций по числу шкафов", len(secs[0][1]), 3)
    check("обозначение шкафа остаётся прописным",
          [i[0] for i in secs[0][1]],
          ["Шкаф ШСК1 РСУ", "Шкаф ШСК1 ПАЗ", "Щит 2Щ"])
    check("пустое имя шкафа пропускается",
          len(core.spec_sections_for_cabs(["ШКАФ А", "", None])[0][1]), 1)
    check("без шкафов разделов нет", core.spec_sections_for_cabs([]), [])

    out = os.path.join(tmp, "Спецификация.docx")
    core.build_spec_docx(tpl, out, secs, log=QUIET)
    truthy("документ создан", os.path.getsize(out) > 20000,
           "%.0f КБ" % (os.path.getsize(out) / 1e3))

    from docx import Document
    d = Document(out)
    s = d.sections[0]
    check("формат листа А4 по ширине", round(s.page_width.mm), 210)
    check("формат листа А4 по высоте", round(s.page_height.mm), 297)

    import zipfile
    z = zipfile.ZipFile(out)
    hf = [x for x in z.namelist()
          if ("header" in x or "footer" in x) and x.endswith(".xml")]
    truthy("рамка и штамп на месте (колонтитулы)", len(hf) >= 2,
           "частей: %d" % len(hf))
    body = z.read("word/document.xml").decode("utf-8")
    stamp = " ".join(z.read(x).decode("utf-8") for x in hf)
    z.close()
    truthy("номера листов — поля, а не числа", "PAGEREF bmspec1" in body)
    truthy("ссылка на лист регистрации изменений",
           "PAGEREF bmspecreg" in body)
    # Шаблон сделан из выпущенного комплекта: шифр, объект, заказчик и
    # фамилии подписавших заменены меткой «Заменить» — иначе чужие данные
    # уедут в новый проект.
    truthy("в штампе стоит метка «Заменить»", "Заменить" in stamp)
    for junk in ("Шифр", "Объект", "Заказчик"):
        check("графа штампа не осталась с чужим значением: " + junk,
              junk.lower() + ":" in stamp.lower(), False)

    tbl = d.tables[0]
    names = [r.cells[1].text.strip() for r in tbl.rows]
    truthy("шкафы попали в таблицу",
           all(n in names for n in ("Шкаф ШСК1 РСУ", "Шкаф ШСК1 ПАЗ", "Щит 2Щ")))
    truthy("начинка шкафов в спецификацию не попала",
           not any("Модуль" in n or "Клемма" in n or "Реле" in n for n in names))
    # первая строка — шапка таблицы, позиции считаем с неё же не начиная
    nums = [r.cells[0].text.strip() for r in tbl.rows[1:] if r.cells[0].text.strip()]
    check("нумерация позиций сквозная", nums, ["1", "2", "3"])

    bad = os.path.join(tmp, "не-шаблон.docx")
    Document().save(bad)
    try:
        core.build_spec_docx(bad, os.path.join(tmp, "нет.docx"), secs, log=QUIET)
        FAIL.append("плохой шаблон спецификации не отвергнут")
        print("  ПЛОХО плохой шаблон спецификации не отвергнут")
    except RuntimeError as e:
        truthy("плохой шаблон отвергается с понятной ошибкой",
               "спецификаци" in str(e).lower(), str(e)[:52])


def t_word_path():
    section("Пути для Word")
    import inspect
    src = inspect.getsource(core.update_fields_word)
    truthy("путь приводится к абсолютному перед передачей в Word",
           "os.path.abspath(p)" in src)


def t_pz_docx(tmp, secs):
    section("Пояснительная записка по ГОСТ")
    tpl = core.default_template("pz")
    truthy("шаблон записки найден", bool(tpl))
    if not tpl:
        return
    equip = [(secs[0][0], [("Модуль ввода аналоговых сигналов AI",
                           "K3.AI.14.16.00", 2, "модули 1.1, 1.2")])]
    out = os.path.join(tmp, "ПЗ.docx")
    core.build_pz_docx(tpl, out, secs, equip, log=QUIET)
    truthy("записка создана", os.path.getsize(out) > 20000,
           "%.0f КБ" % (os.path.getsize(out) / 1e3))

    from docx import Document
    import zipfile
    d = Document(out)
    check("формат листа А4 по ширине", round(d.sections[0].page_width.mm), 210)
    check("формат листа А4 по высоте", round(d.sections[0].page_height.mm), 297)

    z = zipfile.ZipFile(out)
    hf = [x for x in z.namelist()
          if ("header" in x or "footer" in x) and x.endswith(".xml")]
    stamp = " ".join(z.read(x).decode("utf-8") for x in hf)
    body = z.read("word/document.xml").decode("utf-8")
    z.close()
    truthy("рамка и штамп на месте", len(hf) >= 2, "частей: %d" % len(hf))
    for junk in ("проект Б", "дибора", "шифр объекта"):
        check("в записке нет данных исходного проекта: " + junk,
              junk.lower() in (stamp + body).lower(), False)
    truthy("метка «Заменить» на месте", "Заменить" in body)

    toc = [p.text.strip() for p in d.tables[0].rows[0].cells[0].paragraphs
           if p.text.strip()]
    truthy("оглавление собрано", len(toc) >= 10, "строк: %d" % len(toc))
    check("в оглавлении нет повторов", len(toc), len(set(toc)))
    truthy("свой раздел попал в оглавление",
           any("Сводка сигналов" in x for x in toc))
    truthy("лист регистрации изменений в оглавлении",
           any("ЛИСТ РЕГИСТРАЦИИ" in x for x in toc))
    truthy("номера листов — поля, а не числа", "PAGEREF bmpz1" in body)

    heads = [p.text.strip() for p in d.paragraphs if p.text.strip()]
    truthy("раздел не перебивает нумерацию шаблона",
           any(x.startswith("4 Сводка сигналов") for x in heads))

    got = [c.text.strip() for c in d.tables[1].rows[0].cells]
    truthy("сводка сигналов: есть графы «Занято» и «Резерв»",
           "Занято" in got and "Резерв" in got)
    cabs = [r.cells[0].text.strip() for r in d.tables[1].rows[1:]]
    truthy("шкаф назван как в перечне, не прописными",
           all(not c.isupper() for c in cabs if c), cabs[:2])

    # Оформление: таблицы должны быть набраны как таблицы самой записки,
    # а не стилем «Table Grid» — иначе шрифт и рамки выбиваются из документа.
    from docx.oxml.ns import qn as _qn
    proto = d.tables[-2]._tbl        # перечень сокращений — образец из шаблона
    mine = d.tables[1]._tbl          # сводка сигналов — построена программой
    def _borders(t):
        pr = t.find(_qn("w:tblPr"))
        b = pr.find(_qn("w:tblBorders")) if pr is not None else None
        return None if b is None else [(e.tag.split("}")[1], e.get(_qn("w:val")),
                                        e.get(_qn("w:sz"))) for e in b]
    check("рамки таблиц как в шаблоне", _borders(mine), _borders(proto))
    def _italic(t):
        r = t.findall(_qn("w:tr"))[1].findall(_qn("w:tc"))[0]
        return any(x.find(_qn("w:i")) is not None for x in r.iter(_qn("w:rPr")))
    check("шрифт таблиц как в шаблоне (курсив)", _italic(mine), _italic(proto))
    hdr = mine.findall(_qn("w:tr"))[0]
    truthy("шапка таблицы повторяется на каждом листе",
           hdr.find(_qn("w:trPr")) is not None
           and hdr.find(_qn("w:trPr")).find(_qn("w:tblHeader")) is not None)
    caps = [p.text.strip() for p in d.paragraphs
            if p.text.strip().startswith("Таблица ")]
    truthy("у таблиц есть названия по ГОСТ 2.105", len(caps) >= 2, caps[:2])
    truthy("название таблицы оформлено «Таблица N – Имя»",
           all(" – " in c for c in caps), caps[:1])

    bad = os.path.join(tmp, "не-записка.docx")
    Document().save(bad)
    try:
        core.build_pz_docx(bad, os.path.join(tmp, "нет2.docx"), secs, equip, log=QUIET)
        FAIL.append("плохой шаблон ПЗ не отвергнут")
        print("  ПЛОХО плохой шаблон ПЗ не отвергнут")
    except RuntimeError as e:
        truthy("плохой шаблон записки отвергается понятной ошибкой",
               "пз" in str(e).lower() or "записк" in str(e).lower(), str(e)[:50])


def make_dxf_labels(path, keep_prefix=False):
    """Схема с подписями каналов «1.2.N» — как в схемах соединения.

    Описания намеренно разрежены: они есть только у 1-го, 3-го и 6-го каналов.
    Раскладка по порядку отдала бы их каналам 1, 2, 3 — то есть двум каналам
    достались бы чужие сигналы. Правильная раскладка ставит их по подписям.
    """
    import ezdxf
    doc = ezdxf.new(setup=True)
    msp = doc.modelspace()
    x = 300.0
    msp.add_text("Модуль z1.2 AI8x4...20mA", dxfattribs={"insert": (x, 500)})
    for c in (1, 2, 3, 4, 5, 6, 7, 8):
        y = 480.0 - c * 10
        msp.add_text("z1.2.%d+" % c, dxfattribs={"insert": (x, y), "layer": "Надписи"})
        if c in (1, 3, 6):
            mt = msp.add_mtext("Контроль давления в трубопроводе узла номер %d" % c,
                               dxfattribs={"layer": "Схема"})
            mt.set_location((x + 40, y))
    doc.saveas(path)
    return path


def t_channel_labels(tmp):
    section("Раскладка каналов по подписям")
    check("подпись «3.2.1+» разбирается",
          bool(core.CH_LABEL_RE.match("3.2.1+")), True)
    check("подпись с приставкой разбирается",
          bool(core.CH_LABEL_RE.match("z2.2.15-")), True)
    check("подпись без знака разбирается",
          bool(core.CH_LABEL_RE.match("1.6.4")), True)
    check("позиция прибора подписью канала не считается",
          bool(core.CH_LABEL_RE.match("PT-21/2")), False)
    check("номер модуля подписью канала не считается",
          bool(core.CH_LABEL_RE.match("1.6")), False)

    check("число каналов из заголовка «AI8x4...20mA»",
          core.mod_channels("Модуль 3.2 AI8x4...20mA", "AI"), 8)
    check("число каналов из заголовка «8AIx4...20mA»",
          core.mod_channels("Модуль 1.1 AI 8AIx4...20mA", "AI"), 8)
    check("число каналов из заголовка «16DOх24 VDC»",
          core.mod_channels("Модуль z2.5 DO 16DOх24 VDC", "DO"), 16)
    check("цифра из номера модуля числом каналов не считается",
          core.mod_channels("Модуль 4.2 AO", "AO"), None)

    p = make_dxf_labels(os.path.join(tmp, "подписи.dxf"))
    rows = core.extract(p, log=QUIET)
    check("каналов столько, сколько подписей", len(rows), 8)
    by = {r["kc"]: r for r in rows}
    check("позиции по контроллеру сквозные",
          sorted(by, key=lambda k: int(k.split(".")[-1])),
          ["1.2.%d" % i for i in range(1, 9)])
    check("описание легло на свой канал (1)",
          "номер 1" in by["1.2.1"]["desc"], True)
    check("описание легло на свой канал (3)",
          "номер 3" in by["1.2.3"]["desc"], True)
    check("описание легло на свой канал (6)",
          "номер 6" in by["1.2.6"]["desc"], True)
    check("канал без описания остаётся резервом",
          [by["1.2.%d" % i]["desc"] for i in (2, 4, 5, 7, 8)],
          ["Резерв"] * 5)
    check("чертёжная приставка в позицию не идёт",
          any(k.startswith("z") for k in by), False)

    old = core.load_config
    try:
        core.load_config = lambda: {"extract": {"module_prefix": "keep"}}
        rows2 = core.extract(p, log=QUIET)
        check("настройкой приставку можно оставить",
              all(r["kc"].startswith("z") for r in rows2), True)
    finally:
        core.load_config = old

    # На чертеже без подписей работает прежняя раскладка по полосе модуля
    p2 = make_dxf(os.path.join(tmp, "без подписей.dxf"), modules=1, chans=4,
                  with_di=False)
    rows3 = core.extract(p2, log=QUIET)
    truthy("чертёж без подписей разбирается по-прежнему", len(rows3) >= 4,
           "каналов: %d" % len(rows3))
    # «трубопроводе» словарь сокращений превращает в «т/проводе», поэтому
    # проверяем слово, которое он не трогает
    truthy("описания на месте",
           any("давления" in r["desc"] for r in rows3))


def t_direction_by_type(tmp):
    section("Направление и параметры по типу модуля")
    import ezdxf
    p = os.path.join(tmp, "дискретные без якорей.dxf")
    doc = ezdxf.new(setup=True)
    msp = doc.modelspace()
    msp.add_text("Модуль 2.1 DI 16DIх24 VDC", dxfattribs={"insert": (300, 500)})
    for c in (1, 2):
        y = 480.0 - c * 10
        msp.add_text("2.1.%d+" % c, dxfattribs={"insert": (300, y), "layer": "Надписи"})
        mt = msp.add_mtext("Положение отсечного клапана на линии номер %d" % c,
                           dxfattribs={"layer": "Схема"})
        mt.set_location((340, y))
    doc.saveas(p)
    rows = core.extract(p, log=QUIET)
    truthy("дискретный модуль разобран", bool(rows), "каналов: %d" % len(rows))
    if rows:
        S = core.extract_settings()
        check("дискретный ВХОД идёт во входные",
              sorted({r["io"] for r in rows}), ["in"])
        check("уровень дискретный, а не «4-20 мА»",
              sorted({r["level"] for r in rows}), [S["level_discrete"]])
        check("вид контроля дискретный",
              sorted({r["ctrl"] for r in rows}), [S["ctrl_discrete"]])


def t_tag_sources(tmp):
    section("Позиция по проекту")
    check("обозначение с точкой не обрывается",
          core._tag("Управление клапаном HV23.3-6"), "HV23.3-6")
    check("буквенный хвост сохраняется",
          core._tag("Положение отсечного клапана ZA7/11-1o"), "ZA7/11-1o")
    check("дробная часть сохраняется",
          core._tag("Управление клапаном HV110/1"), "HV110/1")
    check("точка в конце предложения позицией не считается",
          core._tag("Регулирование давления. Клапан открыт."), "")
    check("кириллическое обозначение позицией не считается",
          core._tag("Давление в т/проводе нагнетания насоса Н21/1"), "")

    # Служебные обозначения жил не должны попадать в графу «позиция»
    T = []
    y = 100.0
    for i in range(40):
        T.append((10.0, y - i, "Надписи", "A1"))       # вывод реле — по всему листу
        T.append((20.0, y - i, "Надписи", "A2"))
    T.append((30.0, y, "Надписи", "PT21/1+"))          # настоящий прибор
    T.append((30.0, y - 1, "Надписи", "PT21/1-"))
    got = {t for _x, _y, t in core.wire_tags(T, max_wires=6)}
    check("выводы реле отсеиваются", "A1" in got or "A2" in got, False)
    check("обозначение прибора остаётся", "PT21/1" in got, True)
    check("кириллическая «Р» приводится к латинской",
          {t for _x, _y, t in core.wire_tags([(1.0, 1.0, "Надписи", "РT2+")])},
          {"PT2"})

    check("позиция берётся с той же строки",
          core.tag_on_row([(30.0, 100.0, "PT21/1"), (30.0, 108.0, "PT21/2")],
                          10.0, 100.0, 100.0, 4.0), "PT21/1")
    check("если на строке ничего нет — графа пустая",
          core.tag_on_row([(30.0, 200.0, "PT21/1")], 10.0, 100.0, 100.0, 4.0), "")
    check("слева от канала не берём",
          core.tag_on_row([(5.0, 100.0, "PT21/1")], 10.0, 100.0, 100.0, 4.0), "")


def make_dxf_two_columns(path):
    """Два модуля рядом: у каждого своя колонка описаний.

    Проверяет, что описание не утаскивается из соседнего модуля: окно поиска
    шире шага модулей, и раньше канал получал чужой сигнал.
    """
    import ezdxf
    doc = ezdxf.new(setup=True)
    msp = doc.modelspace()
    for k, (x, word) in enumerate(((300.0, "первого"), (528.0, "второго"))):
        msp.add_text("Модуль 1.%d AI8x4...20mA" % (k + 2),
                     dxfattribs={"insert": (x, 500)})
        for c in (1, 2, 3, 4):
            y = 480.0 - c * 16
            msp.add_text("1.%d.%d+" % (k + 2, c),
                         dxfattribs={"insert": (x, y), "layer": "Надписи"})
            mt = msp.add_mtext("Контроль давления в узле %s номер %d" % (word, c),
                               dxfattribs={"layer": "Схема"})
            mt.set_location((x + 112, y))
    doc.saveas(path)
    return path


def t_module_columns(tmp):
    section("Описание берётся из своей колонки модуля")
    p = make_dxf_two_columns(os.path.join(tmp, "две колонки.dxf"))
    rows = core.extract(p, log=QUIET)
    by = {r["kc"]: r["desc"] for r in rows}
    check("каналов обоих модулей", len(rows), 8)
    truthy("первый модуль описан своим текстом",
           all("первого" in by.get("1.2.%d" % c, "") for c in (1, 2, 3, 4)),
           by.get("1.2.1"))
    truthy("второй модуль описан своим текстом",
           all("второго" in by.get("1.3.%d" % c, "") for c in (1, 2, 3, 4)),
           by.get("1.3.1"))


def make_dxf_wires(path, wires_per_device=2):
    """Схема, где канал подключён к прибору 2, 3 или 4 жилами.

    Так устроены настоящие схемы подключения: двухпроводной датчик 4-20 мА,
    термосопротивление на три и на четыре жилы. Описание сигнала стоит на
    строке верхней жилы прибора — по ней канал и находит своё описание.
    """
    import ezdxf
    doc = ezdxf.new(setup=True)
    msp = doc.modelspace()
    x = 300.0
    msp.add_text("Модуль 1.2 AI8x4...20mA", dxfattribs={"insert": (x, 500)})
    for c in (1, 2, 3, 4):
        base = 480.0 - c * 16          # шаг строк модуля
        # подписи канала: сама и «минусовая» жила
        msp.add_text("1.2.%d" % c, dxfattribs={"insert": (x, base), "layer": "Надписи"})
        msp.add_text("1.2.%d-" % c,
                     dxfattribs={"insert": (x + 20, base), "layer": "Надписи"})
        # жилы прибора: верхняя на 8 выше подписи канала
        signs = ["-", "+", "i+", "i-"][:wires_per_device]
        # жилы одного прибора идут вплотную, внутри своей строки модуля,
        # а не разъезжаются на полтора шага — иначе это уже соседний канал
        step = 12.0 / max(len(signs) - 1, 1)
        for k, sg in enumerate(signs):
            msp.add_text("PT2%d%s" % (c, sg),
                         dxfattribs={"insert": (x + 60, base - 6 + step * k),
                                     "layer": "Надписи"})
        top = base - 6 + step * (len(signs) - 1)
        mt = msp.add_mtext("Давление в аппарате номер %d по проекту" % c,
                           dxfattribs={"layer": "Схема"})
        mt.set_location((x + 100, top))
    doc.saveas(path)
    return path


def t_channel_wires(tmp):
    section("Канал и жилы прибора")
    for n in (2, 3, 4):
        p = make_dxf_wires(os.path.join(tmp, "жил%d.dxf" % n), n)
        rows = core.extract(p, log=QUIET)
        by = {r["kc"]: r for r in rows}
        check("жил %d: каналов" % n, len(rows), 4)
        ok = all(("номер %d" % c) in by.get("1.2.%d" % c, {}).get("desc", "")
                 for c in (1, 2, 3, 4))
        check("жил %d: описание легло на свой канал" % n, ok, True)
        tags = [by.get("1.2.%d" % c, {}).get("tag", "") for c in (1, 2, 3, 4)]
        check("жил %d: позиция взята с жил прибора" % n,
              tags, ["PT21", "PT22", "PT23", "PT24"])

    # группировка жил по обозначению
    T = [(10.0, 100.0, "Надписи", "PT21/1+"),
         (10.0, 92.0, "Надписи", "PT21/1-"),
         (10.0, 84.0, "Надписи", "PT21/2+")]
    g = core.wire_groups(T)
    check("жилы сгруппированы по прибору",
          {k: len(v) for k, v in g.items()}, {"PT21/1": 2, "PT21/2": 1})
    w = core.channel_wire(g, 0.0, 100.0, 92.0, 92.0, 8.0)
    check("прибор канала — тот, у кого жил в полосе больше", w[0], "PT21/1")
    check("строка описания — верхняя жила", w[1], 100.0)
    check("за полосой канала прибор не ищется",
          core.channel_wire(g, 0.0, 100.0, 300.0, 300.0, 8.0), None)
    check("слева от канала жилы не берутся",
          core.channel_wire(g, 50.0, 200.0, 92.0, 92.0, 8.0), None)


def main():
    fast = "--fast" in sys.argv
    tmp = tempfile.mkdtemp(prefix="qa_perechni_")
    print("=" * 70)
    print("ПОЛНАЯ ПРОВЕРКА  (временная папка: %s)" % tmp)
    print("=" * 70)
    try:
        p, rows = t_fixture(tmp)
        t_sheets(tmp, p)
        t_stamp(tmp, p)
        built, secs, res = t_build_and_verify(tmp, rows)
        t_verify_negative(tmp, rows)
        t_verify_cab_names(tmp, rows)
        t_perechen_six_columns(tmp, rows)
        t_error_box()
        t_pdf_monochrome()
        t_draw_resilience()
        t_template_guard(tmp)
        t_exports(tmp, secs, res)
        t_dashboard(tmp, secs)
        t_equipment_and_pz(tmp, p, secs)
        t_stamps(tmp)
        t_multisection(tmp, rows)
        t_no_empty_sections(tmp, rows)
        t_pagination(tmp, rows)
        t_revisions(tmp)
        t_manifest(tmp, secs, res)
        t_module_types(tmp)
        t_naming()
        t_cab_naming_details()
        t_module_prefix(tmp)
        t_channel_labels(tmp)
        t_direction_by_type(tmp)
        t_tag_sources(tmp)
        t_channel_wires(tmp)
        t_module_columns(tmp)
        t_equipment_counts()
        t_spec_docx(tmp)
        t_pz_docx(tmp, secs)
        t_word_path()
        t_environment()
        t_config(tmp)
        t_robustness(tmp)
        if fast:
            SKIP.append("печать PDF (--fast)")
            print("\nПечать PDF пропущена (--fast)")
        else:
            t_pdf(tmp, p)
    except Exception:
        FAIL.append("сценарий оборвался")
        print("\nСЦЕНАРИЙ ОБОРВАЛСЯ:")
        traceback.print_exc()
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    print("\n" + "=" * 70)
    print("Пройдено: %d   Не пройдено: %d   Пропущено: %d"
          % (len(PASS), len(FAIL), len(SKIP)))
    for s in SKIP:
        print("   пропуск: " + s)
    if FAIL:
        print("\nНЕ ПРОЙДЕНО:")
        for f in FAIL:
            print("   • " + f)
        return 1
    print("Все сценарии пройдены.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
