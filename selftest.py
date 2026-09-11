# -*- coding: utf-8 -*-
"""Самопроверка программы.

Запуск:  python selftest.py
Ничего не устанавливает и не меняет на диске — только проверяет, что
разбор текста, распознавание форматов листов и настройки работают верно.
Тот же файл понимает pytest:  pytest selftest.py
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import perechni_core as core

FAILED = []


def check(name, got, want):
    ok = got == want
    if not ok:
        FAILED.append(f"{name}\n      получено: {got!r}\n      ожидалось: {want!r}")
    print(("  ок   " if ok else "  ПЛОХО ") + name)
    return ok


# --------------------------------------------------------------- теги позиций
def test_tag():
    print("\nРаспознавание позиции из описания канала")
    check("поз. в описании", core._tag("Давление в линии поз. PIT-1(23)."), "PIT-1(23)")
    check("поз. с пробелами", core._tag("Уровень поз.  LT-5 ."), "LT-5")
    check("без поз., есть тег", core._tag("Сигнализация загазованности YP9 в помещении"), "YP9")
    check("без поз. и без тега", core._tag("Резервный канал"), "")
    check("тег без цифр не берём", core._tag("Контроль ABC давления"), "")


def test_clean():
    print("\nОчистка описания")
    check("убирается поз.", core._clean("Давление поз. PIT-1. в линии"), "Давление в линии")
    check("схлопываются пробелы", core._clean("Уровень    в   баке"), "Уровень в баке")
    check("переносы строк", core._clean("Температура\nв печи"), "Температура в печи")
    check("правка опечатки", core._clean("Сигнал авариного останова"),
          "Сигнал аварийного останова")


# ------------------------------------------------------------ форматы листов
def test_tag_multi_poz():
    print("\nПозиция при нескольких «поз.» в описании")
    # «...из цистерны поз. 12 поз.PY-4» раньше давало огрызок «12 поз»
    check("не берём огрызок",
          core._tag("Контроль давления водорода из цистерны поз. 12 поз.PY-4"), "PY-4")
    check("позиция аппарата — не позиция прибора",
          core._tag("Контроль температуры газов после поз. 54б/1"), "")


def test_tag_near():
    print("\nПозиция из подписи рядом с описанием")
    check("подпись жилы с плюсом", bool(core.TAG_NEAR.match("PI 3211 +")), True)
    check("подпись жилы с минусом", bool(core.TAG_NEAR.match("TI 2213 -")), True)
    check("обычная позиция", bool(core.TAG_NEAR.match("PY-4")), True)
    check("номер канала не позиция", bool(core.TAG_NEAR.match("1.2.3")), False)
    for s in ("AO-1", "AI-3", "DI-12", "DO-5", "XT1"):
        check(f"служебное {s} отбрасывается", bool(core.NOT_A_TAG.match(s)), True)
    check("PI 3211 служебным не считается", bool(core.NOT_A_TAG.match("PI 3211")), False)

    # ближайшая по горизонтали побеждает, если заметно ближе
    rows = [dict(kc="1.1.1", mod="1.1", tag="", desc="Расход азота", x=0.0, y=100.0),
            dict(kc="1.1.2", mod="1.1", tag="", desc="Температура", x=0.0, y=90.0)]
    texts = [(-70.0, 100.0, "Текст", "FSH 5301"),
             (-138.0, 100.0, "Текст", "A1"),
             (-70.0, 90.0, "Текст", "TI 2213"),
             (-138.0, 90.0, "Текст", "A2")]
    n = core._guess_tags(rows, texts, log=None)
    check("заполнено каналов", n, 2)
    check("ближайшая позиция взята", rows[0]["tag"], "FSH 5301")
    check("для второго канала своя", rows[1]["tag"], "TI 2213")

    # два разных кандидата на равном удалении — не гадаем
    rows2 = [dict(kc="1.1.1", mod="1.1", tag="", desc="Что-то", x=0.0, y=100.0),
             dict(kc="1.1.2", mod="1.1", tag="", desc="Ещё", x=0.0, y=90.0)]
    texts2 = [(-70.0, 100.0, "Текст", "PI 1111"), (-72.0, 100.0, "Текст", "TI 2222")]
    core._guess_tags(rows2, texts2, log=None)
    check("спорный случай оставлен пустым", rows2[0]["tag"], "")


def test_formats():
    print("\nОпознание форматов ГОСТ 2.301")
    check("А4 книжная", core._match_format(210, 297), ("А4", "книжная"))
    check("А4 альбомная", core._match_format(297, 210), ("А4", "альбомная"))
    check("А3 альбомная", core._match_format(420, 297), ("А3", "альбомная"))
    check("А0", core._match_format(841, 1189), ("А0", "книжная"))
    check("удлинённый А4х3", core._match_format(297, 630), ("А4х3", "книжная"))
    check("в пределах допуска 3 мм", core._match_format(212, 299), ("А4", "книжная"))
    check("за допуском — не формат", core._match_format(230, 320), None)
    check("не формат вовсе", core._match_format(500, 500), None)


# --------------------------------------------------------- экранирование путей
def test_ps_quote():
    print("\nПодстановка путей в PowerShell")
    check("обычный путь", core._ps_quote(r"C:\work\a.docx"), r"'C:\work\a.docx'")
    check("доллар не раскроется", core._ps_quote(r"C:\цена $100\a.docx"),
          r"'C:\цена $100\a.docx'")
    check("апостроф удваивается", core._ps_quote(r"C:\o'brien\b.docx"),
          r"'C:\o''brien\b.docx'")
    check("пустая строка", core._ps_quote(""), "''")


# ------------------------------------------------------------------ настройки
def test_settings():
    print("\nНастройки распознавания")
    s = core.extract_settings()
    check("есть слой описаний", s["layer_desc"], "Текст")
    check("есть допуск строк DI", s["di_row_tol"], 9.0)
    check("ключей столько же, сколько по умолчанию",
          sorted(s), sorted(core.EXTRACT_DEFAULTS))


# ------------------------------------------------------- знак номера в штампе
def test_stamp_fix():
    print("\nВосстановление знака № в штампе")
    try:
        import ezdxf
    except ImportError:
        print("  пропуск — не установлен ezdxf")
        return
    doc = ezdxf.new()
    doc.styles.add("GOSTW", font="gostw.shx")
    msp = doc.modelspace()
    msp.add_text("Инв.N подл.", dxfattribs={"style": "GOSTW"})
    msp.add_text("Взам. инв.N", dxfattribs={"style": "GOSTW"})
    msp.add_text("NB1-63", dxfattribs={"style": "GOSTW"})       # трогать нельзя
    msp.add_text("L+N", dxfattribs={"style": "GOSTW"})          # и это тоже
    replaced = core._fix_shx_styles(doc)
    check("стиль на SHX опознан как подменяемый", "GOSTW" in replaced, True)
    core._fix_stamp_text(doc, replaced)
    got = [e.dxf.text for e in msp if e.dxftype() == "TEXT"]
    check("Инв.№ подл.", got[0], "Инв.№ подл.")
    check("Взам. инв.№", got[1], "Взам. инв.№")
    check("марка автомата не тронута", got[2], "NB1-63")
    check("нейтраль L+N не тронута", got[3], "L+N")


def test_font_pick():
    print("\nПоиск шрифта для замены SHX")
    f = core.pick_subst_font()
    ok = bool(f)
    print(("  ок    найден: " + f) if ok else
          "  ВНИМАНИЕ  шрифт не найден — чертежи напечатаются шрифтом по умолчанию")


# ------------------------------------------------------------- демо-данные
def test_cab_title():
    print("\nНазвание шкафа в содержании и заголовках")
    # раньше capitalize() гасил обозначение: «Шкаф шск1 рсу», «Щит 2щ»
    check("шкаф", core._cab_title("ШКАФ ШСК1 РСУ"), "Шкаф ШСК1 РСУ")
    check("щит", core._cab_title("ЩИТ 2Щ"), "Щит 2Щ")
    check("шкаф ПАЗ", core._cab_title("ШКАФ ШС2 ПАЗ"), "Шкаф ШС2 ПАЗ")
    check("одно слово", core._cab_title("ШКАФ"), "Шкаф")
    check("пустая строка", core._cab_title(""), "")


def test_default_template():
    print("\nВстроенные шаблоны перечней")
    for io_, name in (("in", "входных"), ("out", "выходных")):
        p = core.default_template(io_)
        ok = bool(p) and os.path.exists(p)
        if not ok:
            FAILED.append(f"нет встроенного шаблона {name} сигналов")
        print(("  ок    %s: %s" % (name, os.path.basename(p))) if ok
              else "  ПЛОХО  не найден шаблон %s сигналов" % name)


def _row(kc, io_, tag, desc, level="4-20 мА", ctrl="", ex="Exd"):
    return dict(kc=kc, io=io_, tag=tag, desc=desc, level=level, ctrl=ctrl, ex=ex,
                mod=kc.rsplit(".", 1)[0], type="AI", ch=1, x=0.0, y=0.0, sortx=0)


class _FakeDoc:
    """Подменяет parse_perechen: сверку проверяем без реального .docx."""
    def __init__(self, rows):
        self.rows = rows


def test_verify():
    print("\nСверка перечня с чертежом")
    orig = core.parse_perechen
    try:
        # 1) документ повторяет чертёж — расхождений быть не должно
        drawing = [_row("1.1.1", "in", "PT1", "Давление"),
                   _row("1.1.2", "in", "TT2", "Температура")]
        doc = [dict(cab="", io="in", tag="PT1", desc="Давление",
                    level="4-20 мА", ctrl="", ex="Exd", kc="1.1.1"),
               dict(cab="", io="in", tag="TT2", desc="Температура",
                    level="4-20 мА", ctrl="", ex="Exd", kc="1.1.2")]
        core.parse_perechen = lambda _p: doc
        r = core.verify_perechen([("ШКАФ", drawing)], ["x.docx"], log=lambda *a: None)
        check("совпадающий документ проходит", r["ok"], True)
        check("сошлось строк", r["counts"]["совпало"], 2)

        # 2) в документе потеряна строка и испорчена позиция
        doc2 = [dict(cab="", io="in", tag="ПТ1", desc="Давление",
                     level="4-20 мА", ctrl="", ex="Exd", kc="1.1.1")]
        core.parse_perechen = lambda _p: doc2
        r2 = core.verify_perechen([("ШКАФ", drawing)], ["x.docx"], log=lambda *a: None)
        check("расхождения найдены", r2["ok"], False)
        kinds = {i[0] for i in r2["issues"]}
        check("замечена потерянная строка", "нет в документе" in kinds, True)
        check("замечена подмена позиции", "расходится: позиция" in kinds, True)

        # 3) выходные каналы не требуются, если подан только перечень входных
        mixed = drawing + [_row("1.9.1", "out", "K1", "Регулирование")]
        core.parse_perechen = lambda _p: doc
        r3 = core.verify_perechen([("ШКАФ", mixed)], ["x.docx"], log=lambda *a: None)
        check("выходные не считаются потерянными", r3["ok"], True)

        # 4) дубль строки в документе
        core.parse_perechen = lambda _p: doc + [doc[0]]
        r4 = core.verify_perechen([("ШКАФ", drawing)], ["x.docx"], log=lambda *a: None)
        check("дубль замечен", any(i[0] == "дубль" for i in r4["issues"]), True)
    finally:
        core.parse_perechen = orig


def test_abbrev():
    print("\nСловарь сокращений")
    tbl = [("технологические трубопроводы", "технол. трубопроводы"),
           ("Контроль давления", "Контр. давления"),
           ("трубопровод", "т/провод")]
    tbl.sort(key=lambda kv: -len(kv[0]))
    check("простая замена",
          core.apply_abbrev("Контроль давления в линии", tbl), "Контр. давления в линии")
    check("длинная пара раньше короткой",
          core.apply_abbrev("технологические трубопроводы", tbl), "технол. трубопроводы")
    check("замена по целым словам",
          core.apply_abbrev("трубопроводчик на месте", tbl), "трубопроводчик на месте")
    check("регистр первой буквы сохраняется",
          core.apply_abbrev("Трубопровод азота", tbl), "Т/провод азота")
    check("пустой словарь ничего не меняет",
          core.apply_abbrev("Контроль давления", []), "Контроль давления")
    check("пустой текст", core.apply_abbrev("", tbl), "")
    real = core.load_abbrev()
    ok = isinstance(real, list)
    if not ok:
        FAILED.append("load_abbrev вернул не список")
    print("  ок    пар в рабочем словаре: %d" % len(real))


def test_auto_layers():
    print("\nАвтоопределение слоёв по чертежу")
    S = core.extract_settings()
    is_desc = lambda t: len(t) > 19 and " " in t
    # слой из настроек пуст, зато описания лежат на «Описания»
    T = [(0, 0, "Описания", "Контроль давления водорода в линии"),
         (0, 1, "Описания", "Контроль температуры газа на выходе"),
         (0, 2, "Реле", "1.2-KLDI3"),
         (0, 3, "Реле вых", "1.3-KLDO5"),
         (0, 4, "искра", "БИЗ-1")]
    got = core._auto_layers(T, S, is_desc, log=None)
    check("слой описаний найден", got["layer_desc"], "Описания")
    check("слой якорей DI найден", got["layer_di"], "Реле")
    check("слой якорей DO найден", got["layer_do"], "Реле вых")
    check("слой БИЗ найден", got["layer_biz"], "искра")
    # если заданный слой работает — он и остаётся
    T2 = [(0, 0, S["layer_desc"], "Контроль давления водорода в линии"),
          (0, 1, "Прочее", "Контроль температуры газа на выходе")]
    got2 = core._auto_layers(T2, S, is_desc, log=None)
    check("заданный слой в приоритете", got2["layer_desc"], S["layer_desc"])


def test_demo_xlsx():
    print("\nЧтение демо-данных")
    p = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                     "examples", "demo-signals.xlsx")
    if not os.path.exists(p):
        print("  пропуск — нет examples/demo-signals.xlsx")
        return
    try:
        sections = core.sections_from_xlsx(p, log=lambda *a: None)
    except Exception as e:
        FAILED.append(f"чтение demo-signals.xlsx: {e}")
        print("  ПЛОХО  не прочитались:", e)
        return
    # sections — список пар (имя шкафа, [каналы])
    total = sum(len(rows) for _name, rows in sections)
    check("разделов в демо-файле", len(sections), 2)
    check("каналов в демо-файле", total, 48)
    types = {r["type"] for _n, rows in sections for r in rows}
    check("типы сигналов распознаны", types >= {"AI", "DI"}, True)
    tagged = [r for _n, rows in sections for r in rows if r["tag"]]
    ok = len(tagged) > 0
    if not ok:
        FAILED.append("demo-signals.xlsx: ни у одного канала нет позиции")
    print("  ок    каналов с позицией: %d из %d" % (len(tagged), total))


def main():
    print("=" * 62)
    print("Самопроверка «Перечни сигналов»")
    print("=" * 62)
    for t in (test_tag, test_tag_multi_poz, test_tag_near, test_clean, test_formats, test_ps_quote, test_settings,
              test_cab_title, test_default_template, test_verify,
              test_abbrev, test_auto_layers,
              test_stamp_fix, test_font_pick, test_demo_xlsx):
        try:
            t()
        except Exception as e:
            FAILED.append(f"{t.__name__}: {type(e).__name__}: {e}")
            print(f"  ПЛОХО  {t.__name__} упал: {e}")
    print("\n" + "=" * 62)
    if FAILED:
        print("НЕ ПРОЙДЕНО проверок: %d\n" % len(FAILED))
        for f in FAILED:
            print("  • " + f)
        return 1
    print("Все проверки пройдены.")
    return 0


# чтобы работало и под pytest
def test_all():
    assert main() == 0


if __name__ == "__main__":
    sys.exit(main())
