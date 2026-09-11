# -*- coding: utf-8 -*-
"""Генератор листов схемы подключения по образцу — НЕЗАВЕРШЁННЫЙ.

Вынесен из perechni_core.py 19.08.2026. Причина: код никем не вызывался —
ни из ядра, ни из интерфейса, — не покрыт тестами и при этом изменяет
геометрию чертежа и пишет DXF. Держать такое в основном модуле рискованно.

Работа не потеряна: файл самодостаточен. Чтобы вернуть в строй, нужно
проверить его на реальном доноре и покрыть тестами.

Использование:
    import sys; sys.path.insert(0, "..")
    from experimental.generate_scheme import generate_scheme
"""
import os
import re

import ezdxf

from perechni_core import read_dxf


def _sheet_frames(msp):
    out = []
    for e in msp.query("LWPOLYLINE"):
        pts = e.get_points("xy")
        xs = [p[0] for p in pts]; ys = [p[1] for p in pts]
        w, h = max(xs) - min(xs), max(ys) - min(ys)
        if 180 < w < 240 and 260 < h < 320:
            out.append((min(xs), min(ys), max(xs), max(ys)))
    out.sort()
    return out

def _anchor(e):
    t = e.dxftype()
    try:
        if t in ("TEXT", "MTEXT", "INSERT"):
            p = e.dxf.insert; return p[0], p[1]
        if t == "LINE":
            return (e.dxf.start[0] + e.dxf.end[0]) / 2, (e.dxf.start[1] + e.dxf.end[1]) / 2
        if t in ("CIRCLE", "ARC", "ELLIPSE"):
            c = e.dxf.center; return c[0], c[1]
        if t == "LWPOLYLINE":
            pts = e.get_points("xy")
            return (sum(p[0] for p in pts) / len(pts), sum(p[1] for p in pts) / len(pts))
        if t == "SPLINE":
            cp = list(e.control_points)
            return (sum(p[0] for p in cp) / len(cp), sum(p[1] for p in cp) / len(cp))
    except Exception:
        pass
    return None

def generate_scheme(donor_dxf, out_dxf, sections, types=("AI", "AO"), log=print):
    """Генерирует листы схемы подключения в стиле донора (бета: аналоговые модули).
    Донор — фирменный чертёж, из которого берётся по одному заполненному листу
    каждого типа; на каждый модуль из sections печатается свой лист."""
    from ezdxf.math import Matrix44
    doc = read_dxf(donor_dxf)
    msp = doc.modelspace()
    frames = _sheet_frames(msp)
    if not frames:
        raise RuntimeError("В доноре не найдены рамки листов А4")
    hdr = re.compile(r"^Модуль\s+(\d+\.\d+)\s+([A-Z]+)")
    kcre = re.compile(r"^(\d+\.\d+\.\d+)([+-])$")

    def sheet_of(x, y):
        for f in frames:
            if f[0] <= x <= f[2] and f[1] <= y <= f[3]:
                return f
        return None

    # донорские листы по типам: первый лист типа с максимумом описаний
    donors = {}
    for t in msp.query("TEXT"):
        m = hdr.match(t.dxf.text.strip())
        if not m or m.group(2) not in types:
            continue
        f = sheet_of(t.dxf.insert[0], t.dxf.insert[1])
        if not f:
            continue
        ndesc = sum(1 for e in msp.query("MTEXT")
                    if e.dxf.layer == "Текст" and "поз." in e.text
                    and f[0] <= e.dxf.insert[0] <= f[2] and f[1] <= e.dxf.insert[1] <= f[3])
        cur = donors.get(m.group(2))
        if cur is None or ndesc > cur[2]:
            donors[m.group(2)] = (f, m.group(1), ndesc)
    PAD = 12
    all_x = [a[0] for e in msp if (a := _anchor(e))]
    gen_x = max(all_x) + 500
    sheet_no = [1]
    biz_no = [1]
    new_ids = set()

    def build_donor(typ):
        f, dmod, _n = donors[typ]
        x0, y0, x1, y1 = f
        ents = [e for e in msp if (a := _anchor(e)) and
                x0 - PAD <= a[0] <= x1 + PAD and y0 - 30 <= a[1] <= y1 + PAD]
        # каналы донора: КС-метки с +/-
        kc = {}
        for e in ents:
            if e.dxftype() == "TEXT":
                m = kcre.match(e.dxf.text.strip())
                if m:
                    kc.setdefault(m.group(1), []).append(e.dxf.insert[1])
        chans = sorted(kc, key=lambda k: -max(kc[k]))
        # теги донора: подписи вида ТЕГ+ / ТЕГ- / ТЕГ-U.1 (не КС)
        sufre = re.compile(r"^(.{2,}?)([+-]|-U\.\d+)$")
        tags = {}
        for e in ents:
            if e.dxftype() == "TEXT":
                t = e.dxf.text.strip()
                m = sufre.match(t)
                if m and not kcre.match(t):
                    tags.setdefault(m.group(1).rstrip("-"), []).append(e.dxf.insert[1])
        dtags = sorted(tags, key=lambda k: -max(tags[k]))
        # описания сверху вниз: текст донора -> номер канала
        descs = sorted([e for e in ents if e.dxftype() == "MTEXT"
                        and e.dxf.layer == "Текст" and "поз." in e.text],
                       key=lambda e: -e.dxf.insert[1])
        dmap = {e.text: i for i, e in enumerate(descs)}
        return dict(f=f, mod=dmod, ents=ents, chans=chans, tags=dtags, dmap=dmap)

    dcache = {t: build_donor(t) for t in donors}
    made = 0
    for cab, rows in sections:
        mods = {}
        for r in rows:
            if r["type"] in dcache:
                mods.setdefault((r["type"], r["mod"]), []).append(r)
        for (typ, mod), chans in sorted(
                mods.items(), key=lambda kv: [int(x) for x in kv[0][1].split(".")]):
            d = dcache[typ]
            x0 = d["f"][0]
            dx = gen_x - x0 + 240.0 * made
            repl = {}
            repl[f"Модуль {d['mod']} {typ}"] = f"Модуль {mod} {typ}"
            repl[f"сигналов {d['mod']} {typ}"] = f"сигналов {mod} {typ}"
            for i, dkc in enumerate(d["chans"]):
                nk = chans[i]["kc"] if i < len(chans) else f"{mod}.{i+1}"
                repl[dkc + "+"] = nk + "+"
                repl[dkc + "-"] = nk + "-"
                repl[dkc] = nk
            tag_map = {}
            for i, dt in enumerate(d["tags"]):
                ch = chans[i] if i < len(chans) else None
                tag_map[dt] = (ch["tag"] if ch and ch["desc"] != "Резерв" else "")
            new = []
            for e in d["ents"]:
                c = e.copy()
                c.transform(Matrix44.translate(dx, 0, 0))
                msp.add_entity(c)
                new.append(c)
                new_ids.add(id(c))
            for e in new:
                if e.dxftype() == "TEXT":
                    t = e.dxf.text.strip()
                    for old, nv in repl.items():
                        if t == old:
                            e.dxf.text = nv
                            break
                    else:
                        hit = None
                        for dt in sorted(tag_map, key=len, reverse=True):
                            if t == dt or (t.startswith(dt) and
                                           re.match(r"^([+-]|-U\.\d+)$", t[len(dt):])):
                                hit = dt
                                break
                        if hit is not None:
                            nt = tag_map[hit]
                            e.dxf.text = (nt + t[len(hit):]) if nt else ""
                        elif "БИЗ-" in t and e.dxf.layer == "mark2":
                            e.dxf.text = f"БИЗ-{typ}-{biz_no[0]}"
                            biz_no[0] += 1
                        else:
                            for old, nv in repl.items():
                                if old in t:
                                    e.dxf.text = t.replace(old, nv)
                                    break
                elif e.dxftype() == "MTEXT":
                    t = e.text
                    if e.dxf.layer == "Текст" and t in d["dmap"]:
                        i = d["dmap"][t]
                        ch = chans[i] if i < len(chans) else None
                        if ch is None or ch["desc"] == "Резерв" or not ch["tag"]:
                            e.text = "Резерв"
                        else:
                            e.text = f"{ch['desc']}. Поз. {ch['tag']}"
                    elif "БИЗ-" in t and "См." in t:
                        e.text = re.sub(r"БИЗ-\w+-\d+\s*", f"БИЗ-{typ}-{biz_no[0]} ", t, count=1)
                    elif e.dxf.layer.startswith("Num_str") and t.strip().isdigit():
                        e.text = str(sheet_no[0])
                    else:
                        for old, nv in repl.items():
                            if old in t:
                                e.text = t.replace(old, nv)
                                break
            made += 1
            sheet_no[0] += 1
            log(f"  лист: {cab} Модуль {mod} {typ} ({min(len(chans), len(d['chans']))} кан.)")
    # удалить всё, кроме новых листов
    for e in list(msp):
        if id(e) not in new_ids:
            try:
                msp.delete_entity(e)
            except Exception:
                pass
    doc.saveas(out_dxf)
    log(f"Схема сохранена: {out_dxf} (листов: {made}) — БЕТА: аналоговые модули")
    return out_dxf
