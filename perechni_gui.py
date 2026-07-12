# -*- coding: utf-8 -*-
"""Перечни сигналов из схем подключения — GUI c боковой навигацией."""
import json, os, sys, threading, traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import perechni_core as core

APP_VER = "2.4"
APP_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))
CFG = os.path.join(APP_DIR, "config.json")
ICON = os.path.join(getattr(sys, "_MEIPASS", APP_DIR), "app.ico")

# ---- палитры (в стиле Claude: тёплые тона, терракотовый акцент)
PALETTES = {
 "light": dict(SIDE="#f0eee6", SIDE_ACT="#d97757", SIDE_HOV="#e8e5da", SIDE_FG="#6c6a60",
     SIDE_TXT="#29261f",
     BG="#faf9f5", CARD="#ffffff", BORDER="#e3e0d5", ACCENT="#d97757", ACCENT2="#c4643f",
     ACCENT_DIS="#ecc3b1", TEXT="#29261f", MUTED="#838073", OK="#5a8a58", WARN="#a6772e",
     ERR="#c14e3d", ROW_ALT="#f7f6f1", FIELD="#faf9f5", BTN2="#f0eee6", BTN2A="#e8e5da",
     SEL="#f0dcd2", DIFF_A="#e2eedd", DIFF_D="#f6dfdb", DIFF_C="#f3ecd4", RES="#a6a39a"),
 "dark": dict(SIDE="#1f1e1d", SIDE_ACT="#d97757", SIDE_HOV="#30302e", SIDE_FG="#a6a39a",
     SIDE_TXT="#f5f4ef",
     BG="#262624", CARD="#30302e", BORDER="#3e3d3a", ACCENT="#d97757", ACCENT2="#c4643f",
     ACCENT_DIS="#6b4a3c", TEXT="#f5f4ef", MUTED="#a6a39a", OK="#7fbf7c", WARN="#d9a648",
     ERR="#e0705c", ROW_ALT="#383836", FIELD="#262624", BTN2="#3e3d3a", BTN2A="#4a4947",
     SEL="#54423a", DIFF_A="#2f4a2c", DIFF_D="#57302a", DIFF_C="#57492a", RES="#78756c"),
}

def set_theme(name):
    globals().update(PALETTES.get(name, PALETTES["light"]))

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    BaseTk = TkinterDnD.Tk
    HAS_DND = True
except Exception:
    BaseTk = tk.Tk
    DND_FILES = None
    HAS_DND = False

FONT      = ("Segoe UI", 10)
FONT_SM   = ("Segoe UI", 9)
FONT_H1   = ("Segoe UI Semibold", 15)
FONT_H2   = ("Segoe UI Semibold", 11)
FONT_NAV  = ("Segoe UI Semibold", 10)
FONT_BTN  = ("Segoe UI Semibold", 11)
FONT_MONO = ("Consolas", 9)


def load_cfg():
    try:
        with open(CFG, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_cfg(d):
    try:
        with open(CFG, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=1)
    except Exception:
        pass


class RoundBtn(tk.Canvas):
    """Скруглённая кнопка в стиле Claude (Canvas): primary — акцентная."""
    def __init__(self, parent, text, cmd, primary=False, font=None, padx=None,
                 pady=None, anchor=None, **_kw):
        self.primary = primary
        self._font = font or (FONT_BTN if primary else FONT_SM)
        self._text = text
        self._cmd = cmd
        self._state = "normal"
        self._anchor = anchor
        px = padx if padx is not None else (22 if primary else 12)
        py = pady if pady is not None else (9 if primary else 6)
        # измерить текст
        probe = tk.Label(parent, text=text, font=self._font)
        probe.update_idletasks()
        w = probe.winfo_reqwidth() + px * 2
        h = probe.winfo_reqheight() + py * 2
        probe.destroy()
        bgp = parent.cget("bg") if isinstance(parent, (tk.Frame, tk.Canvas)) else BG
        super().__init__(parent, width=w, height=h, bg=bgp, highlightthickness=0, bd=0)
        self._bw, self._bh = w, h
        self._fill = ACCENT if primary else BTN2
        self._hover = ACCENT2 if primary else BTN2A
        self._fg = "#ffffff" if primary else TEXT
        self._draw(self._fill)
        self.bind("<Enter>", lambda e: self._state == "normal" and self._draw(self._hover))
        self.bind("<Leave>", lambda e: self._state == "normal" and self._draw(self._fill))
        self.bind("<Button-1>", self._click)
        self.configure(cursor="hand2")

    def _rounded(self, x1, y1, x2, y2, r):
        pts = [x1+r,y1, x2-r,y1, x2,y1, x2,y1+r, x2,y2-r, x2,y2, x2-r,y2,
               x1+r,y2, x1,y2, x1,y2-r, x1,y1+r, x1,y1]
        return self.create_polygon(pts, smooth=True, splinesteps=12)

    def _draw(self, fill):
        self.delete("all")
        r = min(10, self._bh // 2 - 1)
        pid = self._rounded(1, 1, self._bw - 2, self._bh - 2, r)
        self.itemconfigure(pid, fill=fill,
                           outline=fill if self.primary else BORDER)
        anchor_x = 12 if self._anchor == "w" else self._bw // 2
        self.create_text(anchor_x, self._bh // 2, text=self._text, font=self._font,
                         fill=self._fg if self._state == "normal" else MUTED,
                         anchor="w" if self._anchor == "w" else "center")

    def _click(self, _e):
        if self._state == "normal" and self._cmd:
            self._cmd()

    def configure(self, cnf=None, **kw):
        st = kw.pop("state", None)
        bgv = kw.pop("bg", None)
        if st is not None:
            self._state = st
            if st == "disabled":
                self._draw(ACCENT_DIS if self.primary else BTN2)
            else:
                self._draw(self._fill)
        if bgv is not None and st is None:
            self._fill = bgv
            self._draw(bgv)
        if kw or cnf:
            super().configure(cnf, **kw)
    config = configure


def flat_btn(parent, text, cmd, primary=False, danger=False, **kw):
    kw.pop("ipady", None)
    return RoundBtn(parent, text, cmd, primary=primary, **kw)


class Card(tk.Frame):
    def __init__(self, parent, title="", subtitle=""):
        super().__init__(parent, bg=CARD, highlightbackground=BORDER,
                         highlightthickness=1, bd=0)
        if title:
            head = tk.Frame(self, bg=CARD)
            head.pack(fill="x", padx=14, pady=(10, 2))
            tk.Label(head, text=title, font=FONT_H2, bg=CARD, fg=TEXT).pack(side="left")
            if subtitle:
                tk.Label(head, text="   " + subtitle, font=FONT_SM, bg=CARD,
                         fg=MUTED).pack(side="left")
        self.body = tk.Frame(self, bg=CARD)
        self.body.pack(fill="both", expand=True, padx=14, pady=(4, 12))


class PathRow(tk.Frame):
    def __init__(self, parent, label, var, browse, lw=18):
        super().__init__(parent, bg=CARD)
        tk.Label(self, text=label, font=FONT, bg=CARD, fg=TEXT, anchor="w",
                 width=lw).pack(side="left")
        e = tk.Entry(self, textvariable=var, font=FONT, fg=TEXT, bg=FIELD,
                     relief="flat", highlightbackground=BORDER, highlightcolor=ACCENT,
                     highlightthickness=1, insertbackground=TEXT)
        e.pack(side="left", fill="x", expand=True, padx=(8, 6), ipady=5)
        flat_btn(self, "Обзор…", browse, pady=3).pack(side="left")


class FileList(tk.Frame):
    def __init__(self, parent, columns, on_dblclick=None, height=5):
        super().__init__(parent, bg=CARD)
        self.paths = {}
        self.tree = ttk.Treeview(self, style="T.Treeview",
                                 columns=[c[0] for c in columns], show="headings",
                                 selectmode="extended", height=height)
        for cid, title, w in columns:
            self.tree.heading(cid, text=title, anchor="w")
            self.tree.column(cid, width=w, anchor="w")
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.tag_configure("odd", background=ROW_ALT)
        if on_dblclick:
            self.tree.bind("<Double-1>", on_dblclick)
        sb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side="left", fill="y")
        bts = tk.Frame(self, bg=CARD)
        bts.pack(side="left", fill="y", padx=(8, 0))
        for txt, cmd in (("＋  Добавить", self.on_add), ("✕  Удалить", self.remove_sel),
                         ("▲  Вверх", lambda: self.move(-1)), ("▼  Вниз", lambda: self.move(1))):
            flat_btn(bts, txt, cmd, anchor="w").pack(fill="x", pady=2, ipady=3)
        self.add_cb = None

    def on_add(self):
        if self.add_cb:
            self.add_cb()

    def add(self, path, *values):
        iid = self.tree.insert("", "end", values=(os.path.basename(path),) + values)
        self.paths[iid] = path
        self.restripe()
        return iid

    def remove_sel(self):
        for it in self.tree.selection():
            self.paths.pop(it, None)
            self.tree.delete(it)
        self.restripe()

    def move(self, d):
        sel = self.tree.selection()
        items = sel if d < 0 else list(reversed(sel))
        for it in items:
            self.tree.move(it, "", self.tree.index(it) + d)
        self.restripe()

    def restripe(self):
        for i, it in enumerate(self.tree.get_children()):
            self.tree.item(it, tags=("odd",) if i % 2 else ())

    def items(self):
        return [(self.paths[it], self.tree.item(it, "values"))
                for it in self.tree.get_children()]


NAV = [("home", "◈", "Дашборд"),
       ("sep1", None, None),
       ("build", "▣", "Сборка перечней"),
       ("signals", "≣", "Сигналы"),
       ("spec", "▤", "Спецификация"),
       ("conv", "⇄", "Конвертер DWG"),
       ("diff", "Δ", "Сравнение версий"),
       ("sep2", None, None),
       ("settings", "⚙", "Настройки"),
       ("about", "ⓘ", "О программе")]


class App(BaseTk):
    def __init__(self):
        set_theme(load_cfg().get("theme", "light"))
        super().__init__()
        self.title("Перечни сигналов из схем подключения")
        self.configure(bg=BG)
        self.geometry("1120x760")
        self.minsize(1000, 680)
        try:
            self.iconbitmap(ICON)
        except Exception:
            pass
        cfg = load_cfg()
        self.theme = cfg.get("theme", "light")
        self.option_add("*TCombobox*Listbox.background", CARD)
        self.option_add("*TCombobox*Listbox.foreground", TEXT)
        self.option_add("*TCombobox*Listbox.selectBackground", SEL)
        self.option_add("*TCombobox*Listbox.selectForeground", TEXT)
        self.var_in = tk.StringVar(value=cfg.get("template_in", ""))
        self.var_out = tk.StringVar(value=cfg.get("template_out", ""))
        self.var_dst = tk.StringVar(value=cfg.get("out_dir", ""))
        self.var_upd = tk.BooleanVar(value=cfg.get("update_fields", True))
        self.var_pdf = tk.BooleanVar(value=cfg.get("make_pdf", False))
        self.var_rev = tk.BooleanVar(value=cfg.get("revisions", True))
        self.var_cdst = tk.StringVar(value=cfg.get("conv_dir", ""))
        self.var_addlist = tk.BooleanVar(value=cfg.get("conv_add", True))
        self.var_conv = tk.StringVar(value=cfg.get("converter", ""))
        self.var_old_in = tk.StringVar(value=cfg.get("old_in", ""))
        self.var_old_out = tk.StringVar(value=cfg.get("old_out", ""))
        self.var_s_donor = tk.StringVar(value=cfg.get("s_donor", ""))
        self.sections = None            # прочитанные (и правленные) сигналы
        self.sections_files = None      # список файлов, из которых прочитано

        self._styles()
        self._layout()
        self._apply_titlebar()
        for item in cfg.get("drawings", []):
            try:
                p, n = item
            except Exception:
                continue
            if os.path.exists(p):
                self.fl.add(p, n)
        self._refresh_engine()
        self.show_page("home")

    def _apply_titlebar(self):
        """Заголовок окна в цвет темы (Windows 11 DWM)."""
        try:
            import ctypes
            self.update_idletasks()
            hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
            dwm = ctypes.windll.dwmapi

            def set_attr(attr, value):
                v = ctypes.c_int(value)
                return dwm.DwmSetWindowAttribute(hwnd, attr, ctypes.byref(v),
                                                 ctypes.sizeof(v))
            def cref(hexcolor):
                r, g, b = (int(hexcolor[i:i+2], 16) for i in (1, 3, 5))
                return b << 16 | g << 8 | r

            set_attr(20, 1 if self.theme == "dark" else 0)  # dark mode
            set_attr(35, cref(SIDE))    # DWMWA_CAPTION_COLOR — в цвет сайдбара
            set_attr(36, cref(TEXT))    # DWMWA_TEXT_COLOR
            set_attr(34, cref(SIDE))    # DWMWA_BORDER_COLOR
        except Exception:
            pass

    # ------------------------------------------------------------------ стили
    def _styles(self):
        st = ttk.Style(self)
        st.theme_use("clam")
        st.configure("T.Treeview", font=FONT, rowheight=26, background=CARD,
                     bordercolor=BORDER, lightcolor=CARD, darkcolor=CARD,
                     relief="flat",
                     fieldbackground=CARD, foreground=TEXT, borderwidth=0)
        st.configure("T.Treeview.Heading", font=FONT_SM, background=BTN2,
                     foreground=MUTED, relief="flat")
        st.map("T.Treeview", background=[("selected", SEL)],
               foreground=[("selected", TEXT)])
        st.configure("Blue.Horizontal.TProgressbar", troughcolor=BTN2A,
                     background=ACCENT, borderwidth=0, thickness=6)
        st.configure("TCombobox", fieldbackground=FIELD, background=BTN2,
                     foreground=TEXT, arrowcolor=TEXT, bordercolor=BORDER,
                     lightcolor=CARD, darkcolor=CARD, insertcolor=TEXT,
                     selectbackground=SEL, selectforeground=TEXT, padding=3)
        st.map("TCombobox",
               fieldbackground=[("readonly", FIELD), ("disabled", CARD)],
               foreground=[("disabled", MUTED)],
               background=[("active", BTN2A), ("pressed", BTN2A)],
               arrowcolor=[("disabled", MUTED)])
        st.layout("Vertical.TScrollbar",
                  [("Vertical.Scrollbar.trough",
                    {"children": [("Vertical.Scrollbar.thumb",
                                   {"expand": "1", "sticky": "nswe"})],
                     "sticky": "ns"})])
        st.layout("Horizontal.TScrollbar",
                  [("Horizontal.Scrollbar.trough",
                    {"children": [("Horizontal.Scrollbar.thumb",
                                   {"expand": "1", "sticky": "nswe"})],
                     "sticky": "we"})])
        for orient in ("Vertical", "Horizontal"):
            st.configure(f"{orient}.TScrollbar", troughcolor=CARD, background=BTN2A,
                         bordercolor=CARD, lightcolor=CARD, darkcolor=CARD,
                         relief="flat", gripcount=0, width=10)
            st.map(f"{orient}.TScrollbar",
                   background=[("active", RES), ("pressed", RES)])

    # ------------------------------------------------------------------ каркас
    def _layout(self):
        side = tk.Frame(self, bg=SIDE, width=210)
        side.pack(side="left", fill="y")
        side.pack_propagate(False)
        logo = tk.Frame(side, bg=SIDE)
        logo.pack(fill="x", pady=(18, 14), padx=16)
        tk.Label(logo, text="Перечни", font=("Segoe UI Semibold", 14), bg=SIDE,
                 fg=SIDE_TXT).pack(anchor="w")
        tk.Label(logo, text="сигналов АСУ ТП", font=FONT_SM, bg=SIDE,
                 fg=SIDE_FG).pack(anchor="w")
        self.nav = {}
        for key, ic, txt in NAV:
            if ic is None:
                tk.Frame(side, bg=BORDER, height=1).pack(fill="x", padx=16, pady=6)
                continue
            fr = tk.Frame(side, bg=SIDE, cursor="hand2")
            fr.pack(fill="x", pady=1, padx=8)
            bar = tk.Frame(fr, bg=SIDE, width=3, height=26)
            bar.pack(side="left")
            lbl = tk.Label(fr, text=f"  {ic}  {txt}", font=FONT_NAV, bg=SIDE,
                           fg=SIDE_FG, anchor="w", pady=7)
            lbl.pack(side="left", fill="x", expand=True)
            for w in (fr, lbl):
                w.bind("<Button-1>", lambda _e, k=key: self.show_page(k))
                w.bind("<Enter>", lambda _e, f=fr, l=lbl, k=key: self._nav_hover(k, True))
                w.bind("<Leave>", lambda _e, f=fr, l=lbl, k=key: self._nav_hover(k, False))
            self.nav[key] = (fr, bar, lbl)

        main = tk.Frame(self, bg=BG)
        main.pack(side="left", fill="both", expand=True)
        self.page_area = tk.Frame(main, bg=BG)
        self.page_area.pack(fill="both", expand=True)

        logc = Card(main, "Журнал")
        logc.pack(fill="x", padx=16, pady=(2, 12))
        self.txt = tk.Text(logc.body, height=6, state="disabled", font=FONT_MONO,
                           bg=FIELD, fg=TEXT, relief="flat",
                           highlightbackground=BORDER, highlightthickness=1)
        self.txt.pack(fill="both", expand=True)

        self.pages = {
            "home": self._page_home(self.page_area),
            "build": self._page_build(self.page_area),
            "signals": self._page_signals(self.page_area),
            "spec": self._page_spec(self.page_area),
            "conv": self._page_conv(self.page_area),
            "diff": self._page_diff(self.page_area),
            "settings": self._page_settings(self.page_area),
            "about": self._page_about(self.page_area),
        }
        self.current = None
        if HAS_DND:
            self._enable_dnd()


    def _enable_dnd(self):
        def bind_drop(tree, exts, handler):
            tree.drop_target_register(DND_FILES)
            tree.dnd_bind("<<Drop>>", lambda e: handler(
                [f for f in self.tk.splitlist(e.data)
                 if os.path.splitext(f)[1].lower() in exts]))
        bind_drop(self.fl.tree, (".dxf", ".dwg"), self._drop_build)
        bind_drop(self.cl.tree, (".dwg",), self._drop_conv)

    def _drop_build(self, files):
        for f in files:
            self.fl.add(f, core.cab_name_from_file(f))
        if files:
            self.sections = None

    def _drop_conv(self, files):
        for f in files:
            self.cl.add(f)

    def _nav_hover(self, key, on):
        if self.current == key:
            return
        fr, bar, lbl = self.nav[key]
        bg = SIDE_HOV if on else SIDE
        fr.configure(bg=bg); bar.configure(bg=bg); lbl.configure(bg=bg)

    def show_page(self, key):
        for k, fr in self.pages.items():
            fr.pack_forget()
        for k, (fr, bar, lbl) in self.nav.items():
            fr.configure(bg=SIDE); bar.configure(bg=SIDE)
            lbl.configure(bg=SIDE, fg=SIDE_FG)
        fr, bar, lbl = self.nav[key]
        fr.configure(bg=SIDE_HOV); bar.configure(bg=SIDE_ACT)
        lbl.configure(bg=SIDE_HOV, fg=SIDE_TXT)
        self.pages[key].pack(fill="both", expand=True)
        self.current = key

    def _header(self, parent, title, sub=""):
        h = tk.Frame(parent, bg=BG)
        h.pack(fill="x", padx=16, pady=(14, 4))
        tk.Label(h, text=title, font=FONT_H1, bg=BG, fg=TEXT).pack(side="left")
        if sub:
            tk.Label(h, text="  " + sub, font=FONT_SM, bg=BG, fg=MUTED).pack(
                side="left", pady=(5, 0))
        return h

    # ------------------------------------------------------------- стр. Дашборд
    def _page_home(self, parent):
        fr = tk.Frame(parent, bg=BG)
        self._header(fr, "Дашборд", "сводка проекта и быстрые действия")
        bar = tk.Frame(fr, bg=BG)
        bar.pack(fill="x", padx=16, pady=(2, 4))
        flat_btn(bar, "Прочитать чертежи", self.read_signals, primary=True).pack(side="left")
        flat_btn(bar, "⟳ Обновить", self._dash_refresh).pack(side="left", padx=(8, 0))
        flat_btn(bar, "▦ Штампы пакета…", self.stamps_dialog).pack(side="left", padx=(8, 0))
        tk.Label(bar, text="Найти по всем шкафам:", font=FONT_SM, bg=BG,
                 fg=MUTED).pack(side="left", padx=(18, 4))
        self.dash_q = tk.Entry(bar, font=FONT_SM, width=22, bg=FIELD, fg=TEXT,
                               relief="flat", highlightbackground=BORDER,
                               highlightcolor=ACCENT, highlightthickness=1,
                               insertbackground=TEXT)
        self.dash_q.pack(side="left", ipady=3)
        self.dash_q.bind("<Return>", self._dash_search)
        flat_btn(bar, "Найти", lambda: self._dash_search(None)).pack(side="left", padx=(6, 0))

        stats = tk.Frame(fr, bg=BG)
        stats.pack(fill="x", padx=16, pady=(4, 2))
        self.dash_stats = {}
        for key, cap in (("cabs", "Шкафы"), ("total", "Каналы"), ("used", "Занято"),
                         ("res", "Резерв"), ("warns", "Замечания")):
            c = tk.Frame(stats, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
            c.pack(side="left", fill="x", expand=True, padx=(0, 8), ipady=2)
            num = tk.Label(c, text="—", font=("Segoe UI Semibold", 20), bg=CARD, fg=TEXT)
            num.pack(anchor="w", padx=14, pady=(8, 0))
            tk.Label(c, text=cap, font=FONT_SM, bg=CARD, fg=MUTED).pack(
                anchor="w", padx=14, pady=(0, 8))
            self.dash_stats[key] = num

        c2 = Card(fr, "Свободные каналы", "куда посадить новый сигнал")
        c2.pack(fill="both", expand=True, padx=16, pady=(6, 6))
        fbar = tk.Frame(c2.body, bg=CARD)
        fbar.pack(fill="x", pady=(0, 4))
        tk.Label(fbar, text="Тип:", font=FONT_SM, bg=CARD, fg=MUTED).pack(side="left")
        self.free_type = ttk.Combobox(fbar, values=["Все", "AI", "AO", "DI", "DO", "WI"],
                                      width=7, state="readonly", font=FONT_SM)
        self.free_type.set("Все")
        self.free_type.pack(side="left", padx=(4, 0))
        self.free_type.bind("<<ComboboxSelected>>", lambda _e: self._dash_free())
        self.free_stat = tk.Label(fbar, text="", font=FONT_SM, bg=CARD, fg=MUTED)
        self.free_stat.pack(side="right")
        cols = [("cab", "Шкаф", 110), ("mod", "Модуль", 70), ("typ", "Тип", 50),
                ("n", "Свободно", 70), ("kcs", "Позиции", 560)]
        self.fg_tree = ttk.Treeview(c2.body, style="T.Treeview",
                                    columns=[x[0] for x in cols], show="headings", height=7)
        for cid, t, w in cols:
            self.fg_tree.heading(cid, text=t, anchor="w")
            self.fg_tree.column(cid, width=w, anchor="w")
        self.fg_tree.tag_configure("odd", background=ROW_ALT)
        self.fg_tree.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(c2.body, orient="vertical", command=self.fg_tree.yview)
        self.fg_tree.configure(yscrollcommand=sb.set)
        sb.pack(side="left", fill="y")
        self.dash_hint = tk.Label(fr, text="Нажмите «Прочитать чертежи» (список — на «Сборке»), "
                                           "чтобы увидеть сводку.",
                                  font=FONT_SM, bg=BG, fg=MUTED)
        self.dash_hint.pack(anchor="w", padx=18, pady=(0, 6))
        return fr

    def _dash_search(self, _e):
        q = self.dash_q.get().strip()
        if not q:
            return
        if not self.sections:
            messagebox.showwarning("Нет данных", "Сначала «Прочитать чертежи».")
            return
        self.f_text.delete(0, "end")
        self.f_text.insert(0, q)
        self.f_type.set("Все")
        self.show_page("signals")
        self.refresh_signals()

    def _dash_refresh(self):
        st = self.dash_stats
        if not self.sections:
            for lbl in st.values():
                lbl.configure(text="—")
            self.fg_tree.delete(*self.fg_tree.get_children())
            self.free_stat.configure(text="")
            return
        total = used = 0
        for _cab, rows in self.sections:
            for d in rows:
                total += 1
                if d["desc"] != "Резерв":
                    used += 1
        warns = core.checks_report(self.sections)
        st["cabs"].configure(text=str(len(self.sections)))
        st["total"].configure(text=str(total))
        st["used"].configure(text=f"{used} ({used * 100 // max(total, 1)}%)")
        st["res"].configure(text=f"{total - used} ({(total - used) * 100 // max(total, 1)}%)")
        st["warns"].configure(text=str(len(warns)), fg=WARN if warns else OK)
        self.dash_hint.configure(text="")
        self._dash_free()

    def _dash_free(self):
        self.fg_tree.delete(*self.fg_tree.get_children())
        if not self.sections:
            return
        ft = self.free_type.get()
        rows = core.free_channels(self.sections)
        shown = free_total = 0
        for cab, mod, typ, n, kcs in rows:
            free_total += n
            if ft != "Все" and typ != ft:
                continue
            tags = ("odd",) if shown % 2 else ()
            self.fg_tree.insert("", "end", tags=tags, values=(cab, mod, typ, n, kcs))
            shown += 1
        self.free_stat.configure(text=f"всего свободно: {free_total}")

    # ------------------------------------------------------------- штампы пакета
    def stamps_dialog(self):
        win = tk.Toplevel(self)
        win.title("Штампы пакета")
        win.configure(bg=BG)
        win.geometry("640x420")
        win.grab_set()
        tk.Label(win, text="Массовая замена в штампах (даты, стадия)", font=FONT_H2,
                 bg=BG, fg=TEXT).pack(anchor="w", padx=16, pady=(12, 2))
        tk.Label(win, text="Значение меняется в колонтитулах .docx; рядом создаётся .bak",
                 font=FONT_SM, bg=BG, fg=MUTED).pack(anchor="w", padx=16)
        lb = tk.Listbox(win, font=FONT_SM, bg=FIELD, fg=TEXT, relief="flat",
                        highlightbackground=BORDER, highlightthickness=1,
                        selectbackground=SEL, selectforeground=TEXT, height=8)
        lb.pack(fill="both", expand=True, padx=16, pady=8)
        files = []

        def add_files():
            for f in filedialog.askopenfilenames(filetypes=[("Word", "*.docx")]):
                files.append(f)
                lb.insert("end", os.path.basename(f))
            autodetect()

        def add_folder():
            d = filedialog.askdirectory()
            if d:
                for f in sorted(os.listdir(d)):
                    if f.lower().endswith(".docx") and not f.startswith("~"):
                        files.append(os.path.join(d, f))
                        lb.insert("end", f)
                autodetect()

        def autodetect():
            if files and not v_old.get():
                try:
                    v_old.set(core.detect_stamp_value(files[0]))
                except Exception:
                    pass
        row = tk.Frame(win, bg=BG)
        row.pack(fill="x", padx=16)
        flat_btn(row, "＋ Файлы…", add_files).pack(side="left")
        flat_btn(row, "＋ Папка…", add_folder).pack(side="left", padx=(6, 0))
        v_old, v_new = tk.StringVar(), tk.StringVar()
        row2 = tk.Frame(win, bg=BG)
        row2.pack(fill="x", padx=16, pady=8)
        tk.Label(row2, text="Заменить:", font=FONT, bg=BG, fg=TEXT).pack(side="left")
        e1 = tk.Entry(row2, textvariable=v_old, width=10, font=FONT, bg=FIELD, fg=TEXT,
                      relief="flat", highlightbackground=BORDER, highlightcolor=ACCENT,
                      highlightthickness=1, insertbackground=TEXT)
        e1.pack(side="left", padx=6, ipady=3)
        tk.Label(row2, text="на:", font=FONT, bg=BG, fg=TEXT).pack(side="left")
        e2 = tk.Entry(row2, textvariable=v_new, width=10, font=FONT, bg=FIELD, fg=TEXT,
                      relief="flat", highlightbackground=BORDER, highlightcolor=ACCENT,
                      highlightthickness=1, insertbackground=TEXT)
        e2.pack(side="left", padx=6, ipady=3)

        def apply():
            if not files:
                messagebox.showwarning("Штампы", "Добавьте файлы.", parent=win)
                return
            if not v_old.get().strip() or not v_new.get().strip():
                messagebox.showwarning("Штампы", "Заполните оба значения.", parent=win)
                return
            try:
                n = core.replace_in_stamps(files, v_old.get().strip(), v_new.get().strip(),
                                           log=self._logcb())
                self.log(f"Штампы: всего замен {n} в {len(files)} файлах", OK)
                messagebox.showinfo("Штампы", f"Готово: {n} замен.", parent=win)
                win.destroy()
            except Exception as e:
                messagebox.showerror("Ошибка", str(e), parent=win)
        flat_btn(row2, "Применить", apply, primary=True, padx=16, pady=5).pack(
            side="right")

    # ------------------------------------------------------------- стр. Сборка
    def _page_build(self, parent):
        fr = tk.Frame(parent, bg=BG)
        self._header(fr, "Сборка перечней", "чертежи → готовые Word-перечни")

        pr = tk.Frame(fr, bg=BG)
        pr.pack(fill="x", padx=16, pady=(0, 2))
        tk.Label(pr, text="Проект:", font=FONT_SM, bg=BG, fg=MUTED).pack(side="left")
        self.preset_cb = ttk.Combobox(pr, state="readonly", width=28, font=FONT_SM,
                                      values=sorted(load_cfg().get("presets", {})))
        self.preset_cb.pack(side="left", padx=(6, 4))
        self.preset_cb.bind("<<ComboboxSelected>>", self._preset_load)
        flat_btn(pr, "Сохранить как…", self._preset_save).pack(side="left", padx=2)
        flat_btn(pr, "Удалить", self._preset_del).pack(side="left", padx=2)
        c1 = Card(fr, "1. Чертежи схем подключения",
                  "порядок = порядок разделов; раздел переименовывается двойным щелчком")
        c1.pack(fill="both", expand=True, padx=16, pady=5)
        self.fl = FileList(c1.body,
                           [("file", "Чертёж", 470), ("section", "Раздел в перечне", 190)],
                           on_dblclick=self.edit_section)
        self.fl.add_cb = self.add_drawings
        self.fl.pack(fill="both", expand=True)

        c2 = Card(fr, "2. Шаблоны перечней",
                  "любой прежний перечень .docx — таблицы заменятся, рамки сохранятся")
        c2.pack(fill="x", padx=16, pady=5)
        PathRow(c2.body, "Входные сигналы:", self.var_in,
                lambda: self._pick(self.var_in, [("Word", "*.docx")])).pack(fill="x", pady=2)
        PathRow(c2.body, "Выходные сигналы:", self.var_out,
                lambda: self._pick(self.var_out, [("Word", "*.docx")])).pack(fill="x", pady=2)

        c3 = Card(fr, "3. Результат")
        c3.pack(fill="x", padx=16, pady=5)
        PathRow(c3.body, "Папка результата:", self.var_dst,
                lambda: self._pickdir(self.var_dst)).pack(fill="x", pady=2)
        tk.Checkbutton(c3.body, text="Обновить номера страниц содержания через Word (в фоне)",
                       variable=self.var_upd, font=FONT_SM, bg=CARD, fg=TEXT,
                       activebackground=CARD, selectcolor=FIELD).pack(anchor="w")
        tk.Checkbutton(c3.body, text="Сохранить также в PDF (через Word)",
                       variable=self.var_pdf, font=FONT_SM, bg=CARD, fg=TEXT,
                       activebackground=CARD, selectcolor=FIELD).pack(anchor="w")
        tk.Checkbutton(c3.body, text="Пакет — в папку-ревизию с манифестом",
                       variable=self.var_rev, font=FONT_SM, bg=CARD, fg=TEXT,
                       activebackground=CARD, selectcolor=FIELD, activeforeground=TEXT).pack(anchor="w")

        run = tk.Frame(fr, bg=BG)
        run.pack(fill="x", padx=16, pady=(4, 6))
        self.btn = flat_btn(run, "Собрать перечни", self.go, primary=True)
        self.btn.pack(side="left")
        self.btn_pack = flat_btn(run, "Собрать пакет (перечни + Excel + СО + ПЗ)",
                                 self.go_package)
        self.btn_pack.pack(side="left", padx=(8, 0))
        self.btn_open = flat_btn(run, "📂 Открыть папку",
                                 lambda: self._open_dir(self.var_dst.get()))
        self.prog = ttk.Progressbar(run, mode="indeterminate",
                                    style="Blue.Horizontal.TProgressbar", length=200)
        self.status = tk.Label(run, text="", font=FONT_SM, bg=BG, fg=MUTED)
        self.status.pack(side="left", padx=12)
        return fr

    # ------------------------------------------------------------ стр. Сигналы
    def _page_signals(self, parent):
        fr = tk.Frame(parent, bg=BG)
        self._header(fr, "Сигналы", "предпросмотр и правка перед сборкой")
        bar = tk.Frame(fr, bg=BG)
        bar.pack(fill="x", padx=16, pady=(2, 4))
        self.sbtn = flat_btn(bar, "Прочитать чертежи", self.read_signals, primary=True)
        self.sbtn.pack(side="left")
        flat_btn(bar, "⭳ Экспорт в Excel", self.export_xlsx).pack(side="left", padx=(8, 0))
        flat_btn(bar, "✓ Проверки", self.run_checks).pack(side="left", padx=(8, 0))
        flat_btn(bar, "⌖ В AutoCAD", self.show_in_acad).pack(side="left", padx=(8, 0))
        tk.Label(bar, text="Тип:", font=FONT_SM, bg=BG, fg=MUTED).pack(side="left", padx=(16, 4))
        self.f_type = ttk.Combobox(bar, values=["Все", "AI", "AO", "DI", "DO", "WI", "Резерв"],
                                   width=8, state="readonly", font=FONT_SM)
        self.f_type.set("Все")
        self.f_type.pack(side="left")
        self.f_type.bind("<<ComboboxSelected>>", lambda _e: self.refresh_signals())
        tk.Label(bar, text="Поиск:", font=FONT_SM, bg=BG, fg=MUTED).pack(side="left", padx=(12, 4))
        self.f_text = tk.Entry(bar, font=FONT_SM, width=24, bg=FIELD, fg=TEXT,
                               insertbackground=TEXT, relief="flat",
                               highlightbackground=BORDER, highlightcolor=ACCENT,
                               highlightthickness=1)
        self.f_text.pack(side="left", ipady=3)
        self.f_text.bind("<KeyRelease>", lambda _e: self.refresh_signals())
        self.sig_stat = tk.Label(bar, text="", font=FONT_SM, bg=BG, fg=MUTED)
        self.sig_stat.pack(side="right")

        c = Card(fr)
        c.pack(fill="both", expand=True, padx=16, pady=(2, 6))
        cols = [("cab", "Шкаф", 90), ("kc", "Позиция", 80), ("type", "Тип", 45),
                ("tag", "Позиция по проекту", 130), ("desc", "Описание сигнала", 560),
                ("ex", "Ex", 50)]
        self.sg = ttk.Treeview(c.body, style="T.Treeview",
                               columns=[x[0] for x in cols], show="headings")
        for cid, t, w in cols:
            self.sg.heading(cid, text=t, anchor="w")
            self.sg.column(cid, width=w, anchor="w")
        self.sg.tag_configure("odd", background=ROW_ALT)
        self.sg.tag_configure("res", foreground=RES)
        self.sg.tag_configure("edit", foreground=ACCENT)
        self.sg.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(c.body, orient="vertical", command=self.sg.yview)
        self.sg.configure(yscrollcommand=sb.set)
        sb.pack(side="left", fill="y")
        self.sg.bind("<Double-1>", self.edit_signal)
        self.sg_rows = {}   # iid -> dict строки
        tk.Label(fr, text="Двойной щелчок по «Позиция по проекту», «Описание» или «Ex» — правка. "
                          "Правки попадут в перечни при сборке.",
                 font=FONT_SM, bg=BG, fg=MUTED).pack(anchor="w", padx=18, pady=(0, 6))
        return fr


    def run_checks(self):
        if not self.sections:
            messagebox.showwarning("Нет данных", "Сначала «Прочитать чертежи».")
            return
        warns = core.checks_report(self.sections)
        if not warns:
            self.log("Проверки: замечаний нет ✓", OK)
            messagebox.showinfo("Проверки", "Замечаний нет — данные чистые.")
            return
        self.log(f"Проверки: {len(warns)} замечаний", WARN)
        for w in warns:
            self.log("  • " + w, WARN)
        messagebox.showwarning(
            "Проверки", f"Найдено замечаний: {len(warns)}\n\n" +
            "\n".join(warns[:12]) + ("\n…" if len(warns) > 12 else "") +
            "\n\nПолный список — в журнале.")

    def show_in_acad(self):
        sel = self.sg.selection()
        if not sel:
            messagebox.showwarning("AutoCAD", "Выберите строку в таблице сигналов.")
            return
        d = self.sg_rows.get(sel[0])
        if not d:
            return
        try:
            core.show_in_acad(d.get("src") or "", d.get("x"), d.get("y"),
                              log=self._logcb())
        except Exception as e:
            messagebox.showerror("AutoCAD", str(e))

    # -------------------------------------------------------- стр. Спецификация
    def _page_spec(self, parent):
        fr = tk.Frame(parent, bg=BG)
        self._header(fr, "Спецификация", "состав шкафов по чертежам — заготовка СО и ведомость КИП")
        bar = tk.Frame(fr, bg=BG)
        bar.pack(fill="x", padx=16, pady=(2, 4))
        self.qbtn = flat_btn(bar, "Прочитать состав", self.read_spec, primary=True)
        self.qbtn.pack(side="left")
        flat_btn(bar, "⭳ Экспорт в Excel", self.export_spec).pack(side="left", padx=(8, 0))
        flat_btn(bar, "⭳ Заготовка ПЗ (Word)", self.export_pz).pack(side="left", padx=(8, 0))
        self.qprog = ttk.Progressbar(bar, mode="indeterminate",
                                     style="Blue.Horizontal.TProgressbar", length=200)
        self.spec_stat = tk.Label(bar, text="", font=FONT_SM, bg=BG, fg=MUTED)
        self.spec_stat.pack(side="left", padx=12)
        c = Card(fr)
        c.pack(fill="both", expand=True, padx=16, pady=(2, 6))
        cols = [("cab", "Шкаф", 100), ("name", "Элемент", 480), ("art", "Артикул", 150),
                ("n", "Кол-во", 60), ("note", "Примечание", 240)]
        self.qg = ttk.Treeview(c.body, style="T.Treeview",
                               columns=[x[0] for x in cols], show="headings")
        for cid, t, w in cols:
            self.qg.heading(cid, text=t, anchor="w")
            self.qg.column(cid, width=w, anchor="w")
        self.qg.tag_configure("odd", background=ROW_ALT)
        self.qg.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(c.body, orient="vertical", command=self.qg.yview)
        self.qg.configure(yscrollcommand=sb.set)
        sb.pack(side="left", fill="y")
        self.equip = None
        tk.Label(fr, text="Заготовка: количества по чертежу, номенклатуру проверяет инженер. "
                          "В Excel добавляется лист «Ведомость КИП».",
                 font=FONT_SM, bg=BG, fg=MUTED).pack(anchor="w", padx=18, pady=(0, 6))
        return fr

    def read_spec(self):
        if not self._drawing_items():
            messagebox.showwarning("Нет чертежей", "Добавьте чертежи на странице «Сборка перечней».")
            return
        self.qbtn.configure(state="disabled", bg=ACCENT_DIS)
        self.qprog.pack(side="left", padx=12)
        self.qprog.start(12)
        threading.Thread(target=self._spec_work, daemon=True).start()

    def _spec_work(self):
        try:
            eq = core.read_equipment(self._drawing_items(), log=self._logcb())
            self.equip = eq

            def show():
                self.qg.delete(*self.qg.get_children())
                i = 0
                for cab, items in eq:
                    for name, art, n, note in items:
                        self.qg.insert("", "end", tags=("odd",) if i % 2 else (),
                                       values=(cab, name, art, n, note))
                        i += 1
                self.spec_stat.configure(text=f"позиций: {i}")
            self.after(0, show)
        except Exception as e:
            self.after(0, self.log, "ОШИБКА: " + str(e), ERR)
            self.after(0, lambda: messagebox.showerror("Ошибка", str(e)))
        finally:
            self.after(0, self.qprog.stop)
            self.after(0, self.qprog.pack_forget)
            self.after(0, lambda: self.qbtn.configure(state="normal", bg=ACCENT))

    def export_spec(self):
        if not self.equip:
            messagebox.showwarning("Нет данных", "Сначала «Прочитать состав».")
            return
        p = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                         filetypes=[("Excel", "*.xlsx")],
                                         initialfile="Спецификация (заготовка).xlsx")
        if not p:
            return
        try:
            secs = self.sections or []
            core.export_spec_xlsx(self.equip, secs, p)
            self.log(f"Спецификация сохранена: {p}", OK)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def export_pz(self):
        if not self.sections and not self.equip:
            messagebox.showwarning("Нет данных",
                                   "Сначала «Прочитать состав» (и/или «Сигналы» → «Прочитать чертежи»).")
            return
        p = filedialog.asksaveasfilename(defaultextension=".docx",
                                         filetypes=[("Word", "*.docx")],
                                         initialfile="ПЗ раздел Автоматизация (заготовка).docx")
        if not p:
            return
        try:
            core.export_pz_docx(self.sections or [], self.equip or [], p)
            self.log(f"Заготовка ПЗ сохранена: {p}", OK)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    # ---------------------------------------------------------- стр. Конвертер
    def _page_conv(self, parent):
        fr = tk.Frame(parent, bg=BG)
        self._header(fr, "Конвертер DWG → DXF", "пакетная конвертация чертежей")
        c1 = Card(fr, "Файлы DWG")
        c1.pack(fill="both", expand=True, padx=16, pady=5)
        self.cl = FileList(c1.body, [("file", "Чертёж DWG", 640)])
        self.cl.add_cb = self.add_dwgs
        self.cl.pack(fill="both", expand=True)
        c2 = Card(fr, "Параметры")
        c2.pack(fill="x", padx=16, pady=5)
        PathRow(c2.body, "Папка для DXF:", self.var_cdst,
                lambda: self._pickdir(self.var_cdst)).pack(fill="x", pady=2)
        tk.Label(c2.body, text="(пусто — DXF сохранится рядом с DWG)", font=FONT_SM,
                 bg=CARD, fg=MUTED).pack(anchor="w", padx=(150, 0))
        tk.Checkbutton(c2.body, text="После конвертации добавить DXF в «Сборку перечней»",
                       variable=self.var_addlist, font=FONT_SM, bg=CARD, fg=TEXT,
                       activebackground=CARD, selectcolor=FIELD, activeforeground=TEXT).pack(anchor="w", pady=(3, 0))
        run = tk.Frame(fr, bg=BG)
        run.pack(fill="x", padx=16, pady=(4, 6))
        self.cbtn = flat_btn(run, "Конвертировать", self.convert, primary=True)
        self.cbtn.pack(side="left")
        self.cprog = ttk.Progressbar(run, mode="indeterminate",
                                     style="Blue.Horizontal.TProgressbar", length=200)
        self.engine_lbl = tk.Label(run, text="", font=FONT_SM, bg=BG)
        self.engine_lbl.pack(side="left", padx=12)
        return fr

    # ---------------------------------------------------------- стр. Сравнение
    def _page_diff(self, parent):
        fr = tk.Frame(parent, bg=BG)
        self._header(fr, "Сравнение версий",
                     "что изменилось в чертежах относительно прежних перечней")
        c1 = Card(fr, "Прежние перечни (.docx)")
        c1.pack(fill="x", padx=16, pady=5)
        PathRow(c1.body, "Входные сигналы:", self.var_old_in,
                lambda: self._pick(self.var_old_in, [("Word", "*.docx")])).pack(fill="x", pady=2)
        PathRow(c1.body, "Выходные сигналы:", self.var_old_out,
                lambda: self._pick(self.var_old_out, [("Word", "*.docx")])).pack(fill="x", pady=2)
        bar = tk.Frame(fr, bg=BG)
        bar.pack(fill="x", padx=16, pady=(2, 4))
        self.dbtn = flat_btn(bar, "Сравнить", self.run_diff, primary=True)
        self.dbtn.pack(side="left")
        flat_btn(bar, "⭳ Сохранить отчёт Excel", self.export_diff).pack(side="left", padx=(8, 0))
        self.dprog = ttk.Progressbar(bar, mode="indeterminate",
                                     style="Blue.Horizontal.TProgressbar", length=200)
        self.diff_stat = tk.Label(bar, text="", font=FONT_SM, bg=BG, fg=MUTED)
        self.diff_stat.pack(side="left", padx=12)
        c2 = Card(fr)
        c2.pack(fill="both", expand=True, padx=16, pady=(2, 6))
        cols = [("st", "Статус", 90), ("cab", "Шкаф", 95), ("kc", "Позиция", 80),
                ("was", "Было", 380), ("now", "Стало", 380)]
        self.dg = ttk.Treeview(c2.body, style="T.Treeview",
                               columns=[x[0] for x in cols], show="headings")
        for cid, t, w in cols:
            self.dg.heading(cid, text=t, anchor="w")
            self.dg.column(cid, width=w, anchor="w")
        self.dg.tag_configure("добавлено", background=DIFF_A)
        self.dg.tag_configure("удалено", background=DIFF_D)
        self.dg.tag_configure("изменено", background=DIFF_C)
        self.dg.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(c2.body, orient="vertical", command=self.dg.yview)
        self.dg.configure(yscrollcommand=sb.set)
        sb.pack(side="left", fill="y")
        self.diff_report = []
        return fr


    # ---------------------------------------------------------- стр. Настройки
    def _page_settings(self, parent):
        fr = tk.Frame(parent, bg=BG)
        self._header(fr, "Настройки")
        c = Card(fr, "Конвертер DWG → DXF",
                 "ищется автоматически: AutoCAD (accoreconsole), затем ODA File Converter")
        c.pack(fill="x", padx=16, pady=5)
        PathRow(c.body, "Путь (вручную):", self.var_conv,
                lambda: (self._pick(self.var_conv, [("exe", "*.exe")]),
                         self._refresh_engine())).pack(fill="x", pady=2)
        tk.Label(c.body, text="Оставьте пустым для автоматического поиска. "
                              "ODA File Converter — бесплатный (opendesign.com).",
                 font=FONT_SM, bg=CARD, fg=MUTED).pack(anchor="w", pady=(2, 0))
        self.engine_lbl2 = tk.Label(c.body, text="", font=FONT_SM, bg=CARD)
        self.engine_lbl2.pack(anchor="w", pady=(4, 0))
        c2 = Card(fr, "Оформление")
        c2.pack(fill="x", padx=16, pady=5)
        self.var_dark = tk.BooleanVar(value=self.theme == "dark")
        tk.Checkbutton(c2.body, text="Тёмная тема (применится после перезапуска)",
                       variable=self.var_dark, font=FONT, bg=CARD, fg=TEXT,
                       activebackground=CARD, selectcolor=FIELD,
                       command=self._toggle_theme).pack(anchor="w")
        return fr

    def _toggle_theme(self):
        self.theme = "dark" if self.var_dark.get() else "light"
        self._save_cfg()
        if messagebox.askyesno("Тема", "Перезапустить приложение сейчас?"):
            import subprocess
            env = {k: v for k, v in os.environ.items()
                   if not k.startswith("_PYI") and k != "_MEIPASS2"}
            if getattr(sys, "frozen", False):
                subprocess.Popen([sys.executable], env=env, close_fds=True)
            else:
                subprocess.Popen([sys.executable, os.path.abspath(__file__)], env=env)
            self.destroy()

    # ---------------------------------------------------------- стр. О программе
    def _page_about(self, parent):
        fr = tk.Frame(parent, bg=BG)
        self._header(fr, "О программе")
        c = Card(fr)
        c.pack(fill="both", expand=True, padx=16, pady=5)
        b = c.body
        tk.Label(b, text="Перечни сигналов", font=("Segoe UI Semibold", 18), bg=CARD,
                 fg=TEXT).pack(anchor="w", pady=(8, 0))
        tk.Label(b, text=f"версия {APP_VER}", font=FONT_SM, bg=CARD, fg=MUTED).pack(anchor="w")
        txt = (
            "Автоматизация выпуска перечней входных/выходных сигналов АСУ ТП.\n\n"
            "Программа читает чертежи схем подключения (общие виды шкафов, DWG/DXF),\n"
            "извлекает все каналы модулей ввода-вывода (AI/AO/DI/DO/WI) с позициями,\n"
            "тегами КИП, описаниями и видом взрывозащиты — и заполняет фирменные\n"
            "Word-перечни, сохраняя рамки ГОСТ, штампы и содержание с номерами страниц.\n\n"
            "Возможности:\n"
            "  •  сборка перечней по нескольким шкафам за один запуск;\n"
            "  •  предпросмотр и правка сигналов до сборки;\n"
            "  •  пакетный конвертер DWG → DXF (AutoCAD Core Console / ODA);\n"
            "  •  сравнение с прежней версией перечней (отчёт об изменениях);\n"
            "  •  экспорт сводки сигналов в Excel;\n"
            "  •  автообновление номеров страниц через Word.\n\n"
            "Экономит часы ручного переноса данных на каждом проекте\n"
            "и исключает опечатки при копировании."
        )
        tk.Label(b, text=txt, font=FONT, bg=CARD, fg=TEXT, justify="left").pack(anchor="w", pady=8)
        return fr

    # ------------------------------------------------------------------ общее
    def _pick(self, var, types):
        p = filedialog.askopenfilename(filetypes=types)
        if p:
            var.set(p)

    def _pickdir(self, var):
        p = filedialog.askdirectory()
        if p:
            var.set(p)

    def _open_dir(self, d):
        if d and os.path.isdir(d):
            os.startfile(d)

    def add_drawings(self):
        ps = filedialog.askopenfilenames(
            filetypes=[("Чертежи и Excel-сводки", "*.dxf *.dwg *.xlsx"), ("Чертежи", "*.dxf *.dwg"), ("DXF", "*.dxf"), ("DWG", "*.dwg")])
        for p in ps:
            self.fl.add(p, core.cab_name_from_file(p))
        self.sections = None

    def add_dwgs(self):
        ps = filedialog.askopenfilenames(filetypes=[("DWG", "*.dwg")])
        for p in ps:
            self.cl.add(p)

    def edit_section(self, ev):
        tree = self.fl.tree
        it = tree.identify_row(ev.y)
        if not it or tree.identify_column(ev.x) != "#2":
            return
        x, y, w, h = tree.bbox(it, "#2")
        cur = tree.item(it, "values")[1]
        e = tk.Entry(tree, font=FONT)
        e.insert(0, cur)
        e.select_range(0, "end")
        e.place(x=x, y=y, width=w, height=h)
        e.focus_set()

        def done(save):
            if save:
                f, _s = tree.item(it, "values")
                tree.item(it, values=(f, e.get().strip() or cur))
                self.sections = None
            e.destroy()
        e.bind("<Return>", lambda _e: done(True))
        e.bind("<Escape>", lambda _e: done(False))
        e.bind("<FocusOut>", lambda _e: done(True))

    def _refresh_engine(self):
        manual = self.var_conv.get().strip()
        if manual and os.path.exists(manual):
            eng = "AutoCAD" if "accoreconsole" in manual.lower() else "ODA"
            t, c = f"✓ конвертер: {eng} (задан вручную)", OK
        else:
            eng, exe = core.find_converter()
            if eng == "acad":
                t, c = "✓ найден AutoCAD (accoreconsole)", OK
            elif eng == "oda":
                t, c = "✓ найден ODA File Converter", OK
            else:
                t, c = "✗ конвертер не найден (нужен AutoCAD или ODA)", ERR
        self.engine_lbl.configure(text=t, fg=c)
        if hasattr(self, "engine_lbl2"):
            self.engine_lbl2.configure(text=t, fg=c)

    def log(self, s, color=None):
        import time as _t
        s = _t.strftime("%H:%M  ") + s
        self.txt.configure(state="normal")
        self.txt.insert("end", s + "\n")
        if color:
            ln = int(self.txt.index("end-1c").split(".")[0]) - 1
            tag = f"c{ln}"
            self.txt.tag_add(tag, f"{ln}.0", f"{ln}.end")
            self.txt.tag_config(tag, foreground=color)
        self.txt.see("end")
        self.txt.configure(state="disabled")

    def _logcb(self):
        return lambda s: self.after(0, self.log, s, WARN if "ВНИМАНИЕ" in s else None)

    def _save_cfg(self):
        save_cfg(dict(theme=self.theme,
                      template_in=self.var_in.get(), template_out=self.var_out.get(),
                      out_dir=self.var_dst.get(), update_fields=self.var_upd.get(), make_pdf=self.var_pdf.get(), revisions=self.var_rev.get(),
                      conv_dir=self.var_cdst.get(), conv_add=self.var_addlist.get(),
                      converter=self.var_conv.get(),
                      old_in=self.var_old_in.get(), old_out=self.var_old_out.get(),
                      s_donor=self.var_s_donor.get(),
                      drawings=[(p, v[1]) for p, v in self.fl.items()]))

    def _drawing_items(self):
        return [(p, v[1]) for p, v in self.fl.items()]

    def _conv_exe(self):
        p = self.var_conv.get().strip()
        return p if (p and os.path.exists(p)) else None

    # ----------------------------------------------------------- чтение сигналов
    def _ensure_sections(self, force=False):
        """Возвращает sections; читает чертежи при необходимости (в текущем потоке)."""
        items = self._drawing_items()
        if not items:
            raise RuntimeError("Список чертежей пуст (страница «Сборка перечней»).")
        key = json.dumps(items, ensure_ascii=False)
        if not force and self.sections is not None and self.sections_files == key:
            return self.sections
        exe = self._conv_exe()
        secs = core.read_sections(items, oda_exe=exe if exe and "oda" in exe.lower() else None,
                                  log=self._logcb())
        self.sections = secs
        self.sections_files = key
        return secs

    def read_signals(self):
        self.sbtn.configure(state="disabled", bg=ACCENT_DIS)
        threading.Thread(target=self._read_signals_work, daemon=True).start()

    def _read_signals_work(self):
        try:
            self._ensure_sections(force=True)
            self.after(0, self.refresh_signals)
            self.after(0, self._dash_refresh)
        except Exception as e:
            self.after(0, self.log, "ОШИБКА: " + str(e), ERR)
            self.after(0, lambda: messagebox.showerror("Ошибка", str(e)))
        finally:
            self.after(0, lambda: self.sbtn.configure(state="normal", bg=ACCENT))

    def refresh_signals(self):
        self.sg.delete(*self.sg.get_children())
        self.sg_rows.clear()
        if not self.sections:
            self.sig_stat.configure(text="")
            return
        ft = self.f_type.get()
        q = self.f_text.get().strip().lower()
        n = shown = 0
        for cab, rows in self.sections:
            for d in rows:
                n += 1
                if ft == "Резерв" and d["desc"] != "Резерв":
                    continue
                if ft not in ("Все", "Резерв") and d["type"] != ft:
                    continue
                if q and q not in (d["tag"] + " " + d["desc"] + " " + d["kc"]).lower():
                    continue
                tags = []
                if shown % 2:
                    tags.append("odd")
                if d["desc"] == "Резерв":
                    tags.append("res")
                if d.get("_edited"):
                    tags.append("edit")
                iid = self.sg.insert("", "end", tags=tuple(tags),
                                     values=(cab, d["kc"], d["type"], d["tag"],
                                             d["desc"], d["ex"]))
                self.sg_rows[iid] = d
                shown += 1
        self.sig_stat.configure(text=f"показано {shown} из {n}")

    def edit_signal(self, ev):
        it = self.sg.identify_row(ev.y)
        col = self.sg.identify_column(ev.x)
        editable = {"#4": "tag", "#5": "desc", "#6": "ex"}
        if not it or col not in editable:
            return
        field = editable[col]
        x, y, w, h = self.sg.bbox(it, col)
        d = self.sg_rows[it]
        e = tk.Entry(self.sg, font=FONT)
        e.insert(0, d[field])
        e.select_range(0, "end")
        e.place(x=x, y=y, width=w, height=h)
        e.focus_set()

        def done(save):
            if save and e.get().strip() != d[field]:
                d[field] = e.get().strip()
                d["_edited"] = True
                vals = list(self.sg.item(it, "values"))
                vals[{"tag": 3, "desc": 4, "ex": 5}[field]] = d[field]
                self.sg.item(it, values=vals)
                tags = set(self.sg.item(it, "tags")) | {"edit"}
                self.sg.item(it, tags=tuple(tags))
            e.destroy()
        e.bind("<Return>", lambda _e: done(True))
        e.bind("<Escape>", lambda _e: done(False))
        e.bind("<FocusOut>", lambda _e: done(True))

    def export_xlsx(self):
        if not self.sections:
            messagebox.showwarning("Нет данных", "Сначала «Прочитать чертежи».")
            return
        p = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                         filetypes=[("Excel", "*.xlsx")],
                                         initialfile="Сводка сигналов.xlsx")
        if not p:
            return
        try:
            core.export_xlsx(self.sections, p)
            self.log(f"Excel-сводка сохранена: {p}", OK)
        except Exception as e:
            self.log("ОШИБКА экспорта: " + str(e), ERR)
            messagebox.showerror("Ошибка", str(e))

    # ------------------------------------------------------------------ сборка
    def ask_text(self, title, prompt, initial=""):
        """Стилизованный ввод строки (вместо системного simpledialog)."""
        win = tk.Toplevel(self)
        win.title(title)
        win.configure(bg=BG)
        win.resizable(False, False)
        win.grab_set()
        win.transient(self)
        tk.Label(win, text=prompt, font=FONT, bg=BG, fg=TEXT).pack(
            anchor="w", padx=16, pady=(14, 4))
        var = tk.StringVar(value=initial)
        e = tk.Entry(win, textvariable=var, font=FONT, width=34, bg=FIELD, fg=TEXT,
                     relief="flat", highlightbackground=BORDER, highlightcolor=ACCENT,
                     highlightthickness=1, insertbackground=TEXT)
        e.pack(padx=16, ipady=5, fill="x")
        e.focus_set()
        e.select_range(0, "end")
        res = {"v": None}
        row = tk.Frame(win, bg=BG)
        row.pack(fill="x", padx=16, pady=12)

        def ok(_e=None):
            res["v"] = var.get().strip()
            win.destroy()
        flat_btn(row, "OK", ok, primary=True, padx=18, pady=5).pack(side="right")
        flat_btn(row, "Отмена", win.destroy).pack(side="right", padx=(0, 8))
        win.bind("<Return>", ok)
        win.bind("<Escape>", lambda _e: win.destroy())
        self.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 380) // 2
        y = self.winfo_y() + (self.winfo_height() - 160) // 2
        win.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        win.wait_window()
        return res["v"]

    def _preset_save(self):
        name = self.ask_text("Проект", "Название проекта:")
        if not name:
            return
        cfg = load_cfg()
        cfg.setdefault("presets", {})[name] = dict(
            drawings=[(p, v[1]) for p, v in self.fl.items()],
            template_in=self.var_in.get(), template_out=self.var_out.get(),
            out_dir=self.var_dst.get())
        save_cfg(cfg)
        self.preset_cb.configure(values=sorted(cfg["presets"]))
        self.preset_cb.set(name)
        self.log(f"Проект «{name}» сохранён", OK)

    def _preset_load(self, _e=None):
        pr = load_cfg().get("presets", {}).get(self.preset_cb.get())
        if not pr:
            return
        self.fl.tree.delete(*self.fl.tree.get_children())
        self.fl.paths.clear()
        for item in pr.get("drawings", []):
            try:
                path, sec = item
            except Exception:
                continue
            if os.path.exists(path):
                self.fl.add(path, sec)
        self.var_in.set(pr.get("template_in", ""))
        self.var_out.set(pr.get("template_out", ""))
        self.var_dst.set(pr.get("out_dir", ""))
        self.log(f"Проект «{self.preset_cb.get()}» загружен")

    def _preset_del(self):
        name = self.preset_cb.get()
        cfg = load_cfg()
        if name and name in cfg.get("presets", {}):
            del cfg["presets"][name]
            save_cfg(cfg)
            self.preset_cb.configure(values=sorted(cfg.get("presets", {})))
            self.preset_cb.set("")
            self.log(f"Проект «{name}» удалён")

    def go(self):
        items = self._drawing_items()
        t_in, t_out = self.var_in.get().strip(), self.var_out.get().strip()
        dst = self.var_dst.get().strip()
        if not items:
            messagebox.showwarning("Нет чертежей", "Добавьте хотя бы один чертёж.")
            return
        if not (t_in or t_out):
            messagebox.showwarning("Нет шаблонов", "Укажите шаблон входных и/или выходных сигналов.")
            return
        if not dst:
            messagebox.showwarning("Нет папки", "Укажите папку результата.")
            return
        for t in (t_in, t_out):
            if t and os.path.normcase(os.path.dirname(t)) == os.path.normcase(dst):
                if not messagebox.askyesno("Перезапись шаблона",
                        "Папка результата совпадает с папкой шаблона — файл шаблона будет ПЕРЕЗАПИСАН.\nПродолжить?"):
                    return
                break
        self._save_cfg()
        self.btn.configure(state="disabled", bg=ACCENT_DIS)
        self.btn_open.pack_forget()
        self.prog.pack(side="left", padx=12)
        self.prog.start(12)
        self.status.configure(text="Работаю…", fg=MUTED)
        threading.Thread(target=self._work, args=(items, t_in, t_out, dst), daemon=True).start()

    def _work(self, items, t_in, t_out, dst):
        try:
            secs = None
            key = json.dumps(items, ensure_ascii=False)
            if self.sections is not None and self.sections_files == key:
                secs = self.sections
                edits = sum(1 for _c, rows in secs for r in rows if r.get("_edited"))
                if edits:
                    self.after(0, self.log, f"Использую данные предпросмотра (правок: {edits})")
            results = core.run(items, t_in or None, t_out or None, dst,
                               log=self._logcb(), update_fields=self.var_upd.get(),
                               sections=secs, make_pdf=self.var_pdf.get())
            self.after(0, self.log, "ГОТОВО: " + "; ".join(os.path.basename(r) for r in results), OK)
            self.after(0, lambda: self.status.configure(text="Готово ✓", fg=OK))
            self.after(0, lambda: self.btn_open.pack(side="left", padx=(8, 0)))
        except Exception as e:
            self.after(0, self.log, "ОШИБКА: " + str(e), ERR)
            self.after(0, self.log, traceback.format_exc(), MUTED)
            self.after(0, lambda: self.status.configure(text="Ошибка", fg=ERR))
            self.after(0, lambda: messagebox.showerror("Ошибка", str(e)))
        finally:
            self.after(0, self.prog.stop)
            self.after(0, self.prog.pack_forget)
            self.after(0, lambda: self.btn.configure(state="normal", bg=ACCENT))

    def go_package(self):
        items = self._drawing_items()
        t_in, t_out = self.var_in.get().strip(), self.var_out.get().strip()
        dst = self.var_dst.get().strip()
        if not items or not (t_in or t_out) or not dst:
            messagebox.showwarning("Пакетная сборка",
                                   "Нужны: чертежи, шаблон(ы) перечней и папка результата.")
            return
        self._save_cfg()
        self.btn.configure(state="disabled", bg=ACCENT_DIS)
        self.btn_pack.configure(state="disabled")
        self.prog.pack(side="left", padx=12)
        self.prog.start(12)
        self.status.configure(text="Пакетная сборка…", fg=MUTED)
        threading.Thread(target=self._package_work, args=(items, t_in, t_out, dst),
                         daemon=True).start()

    def _package_work(self, items, t_in, t_out, dst):
        try:
            if self.var_rev.get():
                import datetime
                n = 1 + sum(1 for f in os.listdir(dst)
                            if f.startswith("Ревизия") and
                            os.path.isdir(os.path.join(dst, f)))
                dst = os.path.join(dst, f"Ревизия {n:02d} — "
                                   + datetime.date.today().strftime("%d.%m.%Y"))
                os.makedirs(dst, exist_ok=True)
                self.after(0, self.log, f"Пакет собирается в: {dst}")
            results = core.run(items, t_in or None, t_out or None, dst,
                               log=self._logcb(), update_fields=self.var_upd.get(), make_pdf=self.var_pdf.get())
            secs = core.read_sections(items, log=lambda s: None)
            self.sections = secs
            self.sections_files = json.dumps(items, ensure_ascii=False)
            core.export_xlsx(secs, os.path.join(dst, "Сводка сигналов.xlsx"))
            self.after(0, self.log, "Сводка сигналов.xlsx — готово", OK)
            eq = core.read_equipment(items, log=lambda s: None)
            self.equip = eq
            core.export_spec_xlsx(eq, secs, os.path.join(dst, "Спецификация (заготовка).xlsx"))
            self.after(0, self.log, "Спецификация (заготовка).xlsx — готово", OK)
            core.export_pz_docx(secs, eq, os.path.join(dst, "ПЗ Автоматизация (заготовка).docx"))
            self.after(0, self.log, "ПЗ Автоматизация (заготовка).docx — готово", OK)
            if self.var_rev.get():
                import datetime, hashlib
                lines = [f"Пакет собран: {datetime.datetime.now():%d.%m.%Y %H:%M}",
                         f"Приложение: Перечни из схем v{APP_VER}", "", "Чертежи:"]
                for pth, sec_name in items:
                    try:
                        h = hashlib.md5(open(pth, "rb").read()).hexdigest()[:10]
                        mt = datetime.datetime.fromtimestamp(
                            os.path.getmtime(pth)).strftime("%d.%m.%Y %H:%M")
                        lines.append(f"  {sec_name} <- {pth} (изм. {mt}, md5 {h})")
                    except Exception:
                        lines.append(f"  {sec_name} <- {pth}")
                lines += ["", "Шаблоны:", f"  входные:  {t_in}", f"  выходные: {t_out}",
                          "", "Разделы:"]
                for cab, rows in secs:
                    used = sum(1 for r in rows if r["desc"] != "Резерв")
                    lines.append(f"  {cab}: {len(rows)} каналов (занято {used})")
                warns = core.checks_report(secs)
                lines += ["", f"Проверки: {len(warns)} замечаний"]
                lines += ["  " + w for w in warns]
                with open(os.path.join(dst, "манифест.txt"), "w", encoding="utf-8") as f:
                    f.write("\n".join(lines))
                self.after(0, self.log, "манифест.txt — готово", OK)
            self.after(0, self.log, "ПАКЕТ СОБРАН: " + dst, OK)
            self.after(0, lambda: self.status.configure(text="Пакет готов ✓", fg=OK))
            self.after(0, lambda: self.btn_open.pack(side="left", padx=(8, 0)))
        except Exception as e:
            self.after(0, self.log, "ОШИБКА: " + str(e), ERR)
            self.after(0, lambda: self.status.configure(text="Ошибка", fg=ERR))
            self.after(0, lambda: messagebox.showerror("Ошибка", str(e)))
        finally:
            self.after(0, self.prog.stop)
            self.after(0, self.prog.pack_forget)
            self.after(0, lambda: self.btn.configure(state="normal", bg=ACCENT))
            self.after(0, lambda: self.btn_pack.configure(state="normal"))

    # ---------------------------------------------------------------- конвертер
    def convert(self):
        paths = [p for p, _v in self.cl.items()]
        if not paths:
            messagebox.showwarning("Нет файлов", "Добавьте DWG-файлы.")
            return
        out_dir = self.var_cdst.get().strip() or None
        exe = self._conv_exe()
        engine = ("acad" if "accoreconsole" in exe.lower() else "oda") if exe else None
        self._save_cfg()
        self.cbtn.configure(state="disabled", bg=ACCENT_DIS)
        self.cprog.pack(side="left", padx=12)
        self.cprog.start(12)
        threading.Thread(target=self._convert_work, args=(paths, out_dir, engine, exe),
                         daemon=True).start()

    def _convert_work(self, paths, out_dir, engine, exe):
        try:
            outs = {}
            for p in paths:
                od = out_dir or os.path.dirname(p)
                outs.update(core.dwg_to_dxf([p], out_dir=od, engine=engine, exe=exe,
                                            log=self._logcb()))
            self.after(0, self.log, f"ГОТОВО: сконвертировано {len(outs)} из {len(paths)}",
                       OK if len(outs) == len(paths) else WARN)
            if self.var_addlist.get() and outs:
                def add_all():
                    for src in paths:
                        dxf = outs.get(src)
                        if dxf:
                            self.fl.add(dxf, core.cab_name_from_file(dxf))
                    self.sections = None
                    self.show_page("home")
                self.after(0, add_all)
        except Exception as e:
            self.after(0, self.log, "ОШИБКА: " + str(e), ERR)
            self.after(0, lambda: messagebox.showerror("Ошибка", str(e)))
        finally:
            self.after(0, self.cprog.stop)
            self.after(0, self.cprog.pack_forget)
            self.after(0, lambda: self.cbtn.configure(state="normal", bg=ACCENT))

    # ---------------------------------------------------------------- сравнение
    def run_diff(self):
        if not (self.var_old_in.get().strip() or self.var_old_out.get().strip()):
            messagebox.showwarning("Нет перечней", "Укажите прежний перечень (входных и/или выходных).")
            return
        self._save_cfg()
        self.dbtn.configure(state="disabled", bg=ACCENT_DIS)
        self.dprog.pack(side="left", padx=12)
        self.dprog.start(12)
        threading.Thread(target=self._diff_work, daemon=True).start()

    def _diff_work(self):
        try:
            secs = self._ensure_sections()
            old = []
            for p in (self.var_old_in.get().strip(), self.var_old_out.get().strip()):
                if p:
                    old += core.parse_perechen(p)
            rep = core.compare(secs, old, log=self._logcb())
            self.diff_report = rep

            def show():
                self.dg.delete(*self.dg.get_children())
                for st, cab, kc, was, now in rep:
                    self.dg.insert("", "end", tags=(st,), values=(st, cab, kc, was, now))
                a = sum(1 for r in rep if r[0] == "добавлено")
                u = sum(1 for r in rep if r[0] == "удалено")
                c = sum(1 for r in rep if r[0] == "изменено")
                self.diff_stat.configure(
                    text=f"изменений: {len(rep)}  (+{a} / −{u} / ±{c})" if rep
                    else "изменений нет — перечни соответствуют чертежам ✓",
                    fg=WARN if rep else OK)
            self.after(0, show)
        except Exception as e:
            self.after(0, self.log, "ОШИБКА: " + str(e), ERR)
            self.after(0, lambda: messagebox.showerror("Ошибка", str(e)))
        finally:
            self.after(0, self.dprog.stop)
            self.after(0, self.dprog.pack_forget)
            self.after(0, lambda: self.dbtn.configure(state="normal", bg=ACCENT))

    def export_diff(self):
        if not self.diff_report:
            messagebox.showwarning("Нет отчёта", "Сначала выполните сравнение.")
            return
        p = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                         filetypes=[("Excel", "*.xlsx")],
                                         initialfile="Изменения перечней.xlsx")
        if not p:
            return
        try:
            core.export_compare_xlsx(self.diff_report, p)
            self.log(f"Отчёт сохранён: {p}", OK)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))


if __name__ == "__main__":
    App().mainloop()
