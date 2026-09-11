# -*- coding: utf-8 -*-
"""Перечни сигналов из схем подключения — GUI c боковой навигацией."""
import json, os, sys, threading, traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import perechni_core as core

APP_VER = "2.5"
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

def system_theme():
    """Тема оформления Windows: 'dark' | 'light'.

    Нужна, чтобы при первом запуске приложение выглядело так же, как система,
    а не всегда светлым. Свой выбор пользователя сохраняется в настройках
    и дальше имеет приоритет над системным.
    """
    try:
        import winreg
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
        try:
            val, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
        finally:
            winreg.CloseKey(key)
        return "light" if val else "dark"
    except Exception:
        return "light"


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
    """Сохраняет состояние интерфейса, не затирая чужие ключи.

    В этом же файле лежат настройки разбора (`extract`, `abbrev`,
    `auto_layers`, `pdf_a`, пути к конвертерам). Раньше словарь писался
    целиком, и всё, чего не знает интерфейс, пропадало при первом же
    сохранении. Поэтому дописываем поверх прочитанного.
    """
    try:
        cur = load_cfg()
    except Exception:
        cur = {}
    cur.update(d)
    try:
        with open(CFG, "w", encoding="utf-8") as f:
            json.dump(cur, f, ensure_ascii=False, indent=1)
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
       ("pdf", "⎙", "PDF по листам"),
       ("diff", "Δ", "Сравнение версий"),
       ("sep2", None, None),
       ("settings", "⚙", "Настройки"),
       ("about", "ⓘ", "О программе")]


class App(BaseTk):
    def __init__(self):
        set_theme(load_cfg().get("theme") or system_theme())
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
        self.theme = cfg.get("theme") or system_theme()
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
        self.var_manifest = tk.BooleanVar(value=cfg.get("manifest", True))
        self.var_cdst = tk.StringVar(value=cfg.get("conv_dir", ""))
        self.var_pdfdst = tk.StringVar(value=cfg.get("pdf_dir", ""))
        self.var_pdfcolor = tk.BooleanVar(value=cfg.get("pdf_color", False))
        self.var_pdfone = tk.BooleanVar(value=cfg.get("pdf_one", False))
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
        logbar = tk.Frame(logc.body, bg=CARD)
        logbar.pack(fill="x", pady=(0, 2))
        # Замечания тонут в обычных сообщениях, а именно их и надо видеть.
        # Полный текст никуда не девается — переключатель только фильтрует показ.
        self.only_warn = tk.BooleanVar(value=False)
        tk.Checkbutton(logbar, text="только замечания", variable=self.only_warn,
                       command=self._relog, font=FONT_SM, bg=CARD, fg=TEXT,
                       activebackground=CARD, selectcolor=FIELD).pack(side="left")
        self.warn_stat = tk.Label(logbar, text="", font=FONT_SM, bg=CARD, fg=MUTED)
        self.warn_stat.pack(side="left", padx=(10, 0))
        flat_btn(logbar, "⭳ Сохранить журнал…", self.save_log).pack(side="right")
        self.txt = tk.Text(logc.body, height=6, state="disabled", font=FONT_MONO,
                           bg=FIELD, fg=TEXT, relief="flat",
                           highlightbackground=BORDER, highlightthickness=1)
        self.txt.pack(fill="both", expand=True)
        self.log_lines = []          # весь журнал целиком, независимо от фильтра

        self.pages = {
            "home": self._page_home(self.page_area),
            "build": self._page_build(self.page_area),
            "signals": self._page_signals(self.page_area),
            "spec": self._page_spec(self.page_area),
            "conv": self._page_conv(self.page_area),
            "pdf": self._page_pdf(self.page_area),
            "diff": self._page_diff(self.page_area),
            "settings": self._page_settings(self.page_area),
            "about": self._page_about(self.page_area),
        }
        self.current = None
        if HAS_DND:
            self._enable_dnd()
        else:
            # раньше отсутствие библиотеки просто отключало перетаскивание,
            # и пользователь думал, что такой возможности нет вовсе
            self.log("Перетаскивание файлов недоступно: не установлен tkinterdnd2. "
                     "Файлы добавляются кнопкой. Включить: pip install tkinterdnd2", MUTED)


    def _enable_dnd(self):
        def bind_drop(tree, exts, handler):
            tree.drop_target_register(DND_FILES)
            tree.dnd_bind("<<Drop>>", lambda e: handler(
                [f for f in self.tk.splitlist(e.data)
                 if os.path.splitext(f)[1].lower() in exts]))
        bind_drop(self.fl.tree, (".dxf", ".dwg", ".xlsx"), self._drop_build)
        bind_drop(self.cl.tree, (".dwg",), self._drop_conv)
        bind_drop(self.pl.tree, (".dxf", ".dwg"), self._drop_pdf)
        # перечни .docx — на поля страницы «Сравнение версий»
        for widget, var in ((getattr(self, "old_in_row", None), self.var_old_in),
                            (getattr(self, "old_out_row", None), self.var_old_out)):
            if widget is None:
                continue
            try:
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", lambda e, v=var: self._drop_docx(e, v))
            except Exception:
                pass

    def _drop_docx(self, event, var):
        files = [f for f in self.tk.splitlist(event.data)
                 if os.path.splitext(f)[1].lower() == ".docx"]
        if files:
            var.set(files[0])
            self.log("Перетащен перечень: " + os.path.basename(files[0]))

    def _drop_build(self, files):
        for f in files:
            self.fl.add(f, core.cab_name_from_file(f))
        if files:
            self.sections = None

    def _drop_pdf(self, files):
        for f in files:
            self.pl.add(f, "—")
        if files:
            self.log(f"Добавлено на печать: {len(files)}")

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
    def _build_readiness(self, parent):
        """Строка готовности: что есть на этом компьютере, чего не хватает.

        Раньше об ограничениях (нет конвертера DWG, нет Word) человек узнавал,
        только наткнувшись на ошибку посреди работы. Проверка идёт в фоне —
        поиск конвертера ходит по дискам и может занять секунду-другую.
        """
        card = Card(parent, "Готовность", "что доступно на этом компьютере")
        card.pack(fill="x", padx=16, pady=(4, 2))
        self.ready_box = tk.Frame(card.body, bg=CARD)
        self.ready_box.pack(fill="x")
        tk.Label(self.ready_box, text="проверяю…", font=FONT_SM,
                 bg=CARD, fg=MUTED).pack(anchor="w")
        threading.Thread(target=self._readiness_work, daemon=True).start()
        return card

    def _readiness_work(self):
        try:
            rows = core.readiness()
        except Exception as e:
            rows = [(False, "Проверка готовности", "не удалась", str(e))]
        self.after(0, self._readiness_show, rows)

    def _readiness_show(self, rows):
        for w in self.ready_box.winfo_children():
            w.destroy()
        for ok_, name, detail, cons in rows:
            line = tk.Frame(self.ready_box, bg=CARD)
            line.pack(fill="x", anchor="w")
            tk.Label(line, text=("✓" if ok_ else "✗"), font=FONT_SM, bg=CARD,
                     fg=(OK if ok_ else ERR), width=2).pack(side="left")
            tk.Label(line, text=name + ":", font=FONT_SM, bg=CARD, fg=TEXT,
                     width=17, anchor="w").pack(side="left")
            tk.Label(line, text=detail, font=FONT_SM, bg=CARD,
                     fg=(TEXT if ok_ else ERR), anchor="w").pack(side="left")
            if cons:
                tk.Label(line, text="— " + cons, font=FONT_SM, bg=CARD,
                         fg=MUTED, anchor="w", justify="left").pack(side="left", padx=(6, 0))

    def _page_home(self, parent):
        fr = tk.Frame(parent, bg=BG)
        self._header(fr, "Дашборд", "сводка проекта и быстрые действия")
        self._build_readiness(fr)
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
        self.dash_hint = tk.Label(
            fr,
            text="С чего начать:  1) слева откройте «Сборка перечней» и добавьте чертежи "
                 "(DWG или DXF)  →  2) вернитесь сюда и нажмите «Прочитать чертежи».   "
                 "Нет чертежей под рукой — добавьте examples\\demo-signals.xlsx, это демо.",
            font=FONT_SM, bg=BG, fg=MUTED, justify="left")
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

        # Шаблоны переехали в «Настройки»: по умолчанию берутся встроенные,
        # и два поля выбора файла на первом экране только сбивали с толку.
        c3 = Card(fr, "2. Результат")
        c3.pack(fill="x", padx=16, pady=5)
        PathRow(c3.body, "Папка результата:", self.var_dst,
                lambda: self._pickdir(self.var_dst)).pack(fill="x", pady=2)
        tk.Checkbutton(c3.body, text="Обновить номера страниц содержания через Word (в фоне)",
                       variable=self.var_upd, font=FONT_SM, bg=CARD, fg=TEXT,
                       activebackground=CARD, selectcolor=FIELD).pack(anchor="w")
        tk.Checkbutton(c3.body, text="Сохранить также в PDF (через Word)",
                       variable=self.var_pdf, font=FONT_SM, bg=CARD, fg=TEXT,
                       activebackground=CARD, selectcolor=FIELD).pack(anchor="w")
        # Раньше одна галочка управляла двумя разными вещами: и папкой-ревизией,
        # и паспортом выпуска. Паспорт полезен всегда, поэтому он теперь
        # отдельный и включён по умолчанию.
        tk.Checkbutton(c3.body, text="Складывать в папку-ревизию (Ревизия NN — дата)",
                       variable=self.var_rev, font=FONT_SM, bg=CARD, fg=TEXT,
                       activebackground=CARD, selectcolor=FIELD,
                       activeforeground=TEXT).pack(anchor="w")
        tk.Checkbutton(c3.body,
                       text="Писать паспорт выпуска (из каких чертежей собрано и сошлась ли сверка)",
                       variable=self.var_manifest, font=FONT_SM, bg=CARD, fg=TEXT,
                       activebackground=CARD, selectcolor=FIELD,
                       activeforeground=TEXT).pack(anchor="w")

        run = tk.Frame(fr, bg=BG)
        run.pack(fill="x", padx=16, pady=(4, 6))
        self.btn = flat_btn(run, "Собрать перечни", self.go, primary=True)
        self.btn.pack(side="left")
        self.btn_pack = flat_btn(run, "Собрать пакет (перечни + Excel + СО + ПЗ)",
                                 self.go_package)
        self.btn_pack.pack(side="left", padx=(8, 0))
        self.btn_open = flat_btn(run, "📂 Открыть папку",
                                 lambda: self._open_dir(self.var_dst.get()))
        self.btn_verify = flat_btn(run, "⚖ Отчёт сверки…", self.export_verify)
        self.verify_result = None
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
        flat_btn(bar, "⎙ Лист в PDF", self.show_sheet_pdf).pack(side="left", padx=(8, 0))
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

    def show_sheet_pdf(self):
        """Печатает лист, на котором лежит выбранный канал, и открывает его.

        Переход «в AutoCAD» требует установленного CAD, а посмотреть, откуда
        взялась строка, нужно и без него. Печатается только один лист —
        это секунды, а не минуты на весь чертёж.
        """
        sel = self.sg.selection()
        if not sel:
            messagebox.showwarning("Лист в PDF", "Выберите строку в таблице сигналов.")
            return
        d = self.sg_rows.get(sel[0])
        if not d:
            return
        src = d.get("src") or ""
        if not src or not os.path.exists(src):
            messagebox.showwarning(
                "Лист в PDF",
                "Неизвестно, из какого чертежа эта строка. "
                "Так бывает, когда данные прочитаны из Excel-сводки.")
            return
        self.log("Печать листа с каналом %s..." % d.get("kc", ""))
        threading.Thread(target=self._sheet_pdf_work, args=(src, d), daemon=True).start()

    def _sheet_pdf_work(self, src, d):
        import tempfile
        try:
            if src.lower().endswith(".dwg"):
                conv = core.dwg_to_dxf([src], log=self._logcb())
                src = conv.get(src, src)
            out = os.path.join(tempfile.mkdtemp(prefix="perechni_sheet_"),
                               "Канал %s.pdf" % str(d.get("kc", "")).replace(".", "-"))
            p, note = core.export_channel_sheet_pdf(src, d.get("x"), d.get("y"),
                                                    out, log=self._logcb())
            if not p:
                self.after(0, self.log, "не удалось: " + note, ERR)
                self.after(0, lambda: messagebox.showwarning("Лист в PDF", note))
                return
            self.after(0, self.log, "открываю: " + note, OK)
            os.startfile(p)
        except Exception as e:
            self.after(0, self.log, "ОШИБКА: " + str(e), ERR)
            self.error_box("Лист в PDF", e)

    # -------------------------------------------------------- стр. Спецификация
    def _page_spec(self, parent):
        fr = tk.Frame(parent, bg=BG)
        self._header(fr, "Спецификация",
                     "шкафы — в спецификацию по ГОСТ, их начинка — в ведомость комплектации")
        bar = tk.Frame(fr, bg=BG)
        bar.pack(fill="x", padx=16, pady=(2, 4))
        self.qbtn = flat_btn(bar, "Прочитать состав", self.read_spec, primary=True)
        self.qbtn.pack(side="left")
        flat_btn(bar, "⭳ Спецификация (Word)", self.export_spec).pack(side="left", padx=(8, 0))
        flat_btn(bar, "⭳ Ведомость комплектации (Excel)",
                 self.export_vedom).pack(side="left", padx=(8, 0))
        flat_btn(bar, "⭳ Пояснительная записка (Word)",
                 self.export_pz).pack(side="left", padx=(8, 0))
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
            self.error_box("Ошибка", e)
        finally:
            self.after(0, self.qprog.stop)
            self.after(0, self.qprog.pack_forget)
            self.after(0, lambda: self.qbtn.configure(state="normal", bg=ACCENT))

    def _cab_names(self):
        """Имена шкафов проекта: из состава, иначе из разделов перечня."""
        return ([n for n, _i in (self.equip or [])]
                or [n for n, _r in (self.sections or [])])

    def export_spec(self):
        if not self.equip and not self.sections:
            messagebox.showwarning("Нет данных", "Сначала «Прочитать состав».")
            return
        tpl = core.default_template("spec")
        if not tpl:
            messagebox.showerror(
                "Нет шаблона",
                "Не найден шаблон templates/Спецификация оборудования и материалов.docx. "
                "Положите шаблон рядом с программой.")
            return
        p = filedialog.asksaveasfilename(
            defaultextension=".docx", filetypes=[("Word", "*.docx")],
            initialfile="Спецификация оборудования и материалов.docx")
        if not p:
            return
        try:
            core.build_spec_docx(tpl, p, core.spec_sections_for_cabs(self._cab_names()),
                                 log=self.log)
            self.log(f"Спецификация сохранена: {p}", OK)
            self.log(core.SPEC_EXCLUDED_NOTE)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def export_vedom(self):
        """Ведомость комплектации шкафов — рабочий документ, не спецификация."""
        if not self.equip:
            messagebox.showwarning("Нет данных", "Сначала «Прочитать состав».")
            return
        p = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                         filetypes=[("Excel", "*.xlsx")],
                                         initialfile="Ведомость комплектации шкафов.xlsx")
        if not p:
            return
        try:
            core.export_spec_xlsx(self.equip, self.sections or [], p)
            self.log(f"Ведомость комплектации сохранена: {p}", OK)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def export_pz(self):
        if not self.sections and not self.equip:
            messagebox.showwarning("Нет данных",
                                   "Сначала «Прочитать состав» (и/или «Сигналы» → «Прочитать чертежи»).")
            return
        p = filedialog.asksaveasfilename(defaultextension=".docx",
                                         filetypes=[("Word", "*.docx")],
                                         initialfile="Пояснительная записка.docx")
        if not p:
            return
        try:
            tpl = core.default_template("pz")
            if tpl:
                core.build_pz_docx(tpl, p, self.sections or [], self.equip or [],
                                   log=self.log)
            else:
                # без шаблона запиской занимается прежний путь: лист без рамки
                core.export_pz_docx(self.sections or [], self.equip or [], p)
                self.log("шаблон записки не найден — документ без рамки и штампа", ERR)
            self.log(f"Пояснительная записка сохранена: {p}", OK)
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

    # -------------------------------------------------------- стр. PDF по листам
    def _page_pdf(self, parent):
        fr = tk.Frame(parent, bg=BG)
        self._header(fr, "PDF по листам",
                     "каждый лист чертежа — отдельная страница своего формата (А5…А0, А4х3 и др.)")
        c1 = Card(fr, "Чертежи", "DWG сконвертируются автоматически")
        c1.pack(fill="both", expand=True, padx=16, pady=5)
        self.pl = FileList(c1.body, [("file", "Чертёж", 470), ("fmt", "Форматы листов", 200)])
        self.pl.add_cb = self.add_pdf_files
        self.pl.pack(fill="both", expand=True)
        bar = tk.Frame(c1.body, bg=CARD)
        bar.pack(fill="x", pady=(6, 0))
        flat_btn(bar, "⟳ Определить форматы", self.scan_formats).pack(side="left")
        flat_btn(bar, "↧ Взять из «Сборки»", self.pdf_from_build).pack(side="left", padx=(8, 0))
        self.pdf_hint = tk.Label(bar, text="", font=FONT_SM, bg=CARD, fg=MUTED)
        self.pdf_hint.pack(side="left", padx=12)

        c2 = Card(fr, "Параметры")
        c2.pack(fill="x", padx=16, pady=5)
        PathRow(c2.body, "Папка для PDF:", self.var_pdfdst,
                lambda: self._pickdir(self.var_pdfdst)).pack(fill="x", pady=2)
        tk.Label(c2.body, text="(пусто — PDF сохранится рядом с чертежом)", font=FONT_SM,
                 bg=CARD, fg=MUTED).pack(anchor="w", padx=(150, 0))
        tk.Checkbutton(c2.body, text="Цветная печать (по умолчанию — монохром, как в AutoCAD)",
                       variable=self.var_pdfcolor, font=FONT_SM, bg=CARD, fg=TEXT,
                       activebackground=CARD, selectcolor=FIELD,
                       activeforeground=TEXT).pack(anchor="w", pady=(3, 0))
        tk.Checkbutton(c2.body, text="Все чертежи в один PDF-файл",
                       variable=self.var_pdfone, font=FONT_SM, bg=CARD, fg=TEXT,
                       activebackground=CARD, selectcolor=FIELD,
                       activeforeground=TEXT).pack(anchor="w")

        run = tk.Frame(fr, bg=BG)
        run.pack(fill="x", padx=16, pady=(4, 6))
        self.pbtn = flat_btn(run, "Печать в PDF", self.make_pdf, primary=True)
        self.pbtn.pack(side="left")
        self.pprog = ttk.Progressbar(run, mode="determinate",
                                     style="Blue.Horizontal.TProgressbar", length=220)
        self.pstat = tk.Label(run, text="", font=FONT_SM, bg=BG, fg=MUTED)
        self.pstat.pack(side="left", padx=12)
        self.pdf_results = []
        return fr

    def add_pdf_files(self):
        ps = filedialog.askopenfilenames(
            filetypes=[("Чертежи", "*.dxf *.dwg"), ("DXF", "*.dxf"), ("DWG", "*.dwg")])
        for x in ps:
            self.pl.add(x, "—")

    def pdf_from_build(self):
        n = 0
        have = {p for p, _v in self.pl.items()}
        for path, _vals in self.fl.items():
            if path not in have and path.lower().endswith((".dwg", ".dxf")):
                self.pl.add(path, "—")
                n += 1
        self.log(f"Добавлено из «Сборки»: {n}" if n else "Нечего добавлять из «Сборки»")

    def scan_formats(self):
        items = self.pl.items()
        if not items:
            messagebox.showwarning("Нет файлов", "Добавьте чертежи.")
            return
        self.pdf_hint.configure(text="определяю…")
        threading.Thread(target=self._scan_work, args=(items,), daemon=True).start()

    def _scan_work(self, items):
        total = 0
        for iid, (path, _vals) in zip(self.pl.tree.get_children(), items):
            try:
                if path.lower().endswith(".dwg"):
                    conv = core.dwg_to_dxf([path], exe=self._conv_exe() or None,
                                           log=lambda s: None)
                    path = conv.get(path, path)
                _doc, frames = core.detect_sheets(path)
                txt = core.sheets_summary(frames) if frames else "рамки не найдены"
                total += len(frames)
            except Exception as e:
                txt = "ошибка: " + str(e)[:40]
            self.after(0, lambda i=iid, t=txt: self.pl.tree.set(i, "fmt", t))
        self.after(0, lambda: self.pdf_hint.configure(text=f"всего листов: {total}"))
        self.after(0, self.log, f"Форматы определены, листов всего: {total}", OK)

    def make_pdf(self):
        items = self.pl.items()
        if not items:
            messagebox.showwarning("Нет файлов", "Добавьте чертежи.")
            return
        self._save_cfg()
        self.pbtn.configure(state="disabled", bg=ACCENT_DIS)
        self.pprog.pack(side="left", padx=12)
        self.pprog.configure(value=0, maximum=100)
        self.pstat.configure(text="печать…", fg=MUTED)
        threading.Thread(target=self._pdf_work, args=([p for p, _v in items],),
                         daemon=True).start()

    def _pdf_work(self, paths):
        try:
            out_dir = self.var_pdfdst.get().strip()
            color = self.var_pdfcolor.get()
            one = self.var_pdfone.get()
            log = self._logcb()
            exe = self._conv_exe() or None
            dwgs = [x for x in paths if x.lower().endswith(".dwg")]
            conv = core.dwg_to_dxf(dwgs, exe=exe, log=log) if dwgs else {}
            results, pages_all = [], []
            for i, src in enumerate(paths, 1):
                dxf = conv.get(src, src)
                self.after(0, self.log, f"Печать: {os.path.basename(src)}")
                od = out_dir or os.path.dirname(src)
                os.makedirs(od, exist_ok=True)
                if one:
                    pages_all.append(dxf)
                    continue

                def prog(k, n, i=i, tot=len(paths)):
                    pct = ((i - 1) + k / max(n, 1)) / tot * 100
                    self.after(0, lambda: self.pprog.configure(value=pct))
                    self.after(0, lambda: self.pstat.configure(
                        text=f"{os.path.basename(src)[:24]}: лист {k}/{n}"))
                out = os.path.join(od, os.path.splitext(os.path.basename(src))[0] + ".pdf")
                r, _fr = core.export_sheets_pdf(dxf, out, color=color, log=log, progress=prog)
                if r:
                    results.append(r)
            if one and pages_all:
                od = out_dir or os.path.dirname(paths[0])
                out = os.path.join(od, "Чертежи (полистно).pdf")

                def prog1(k, n):
                    self.after(0, lambda: self.pprog.configure(value=k / max(n, 1) * 100))
                    self.after(0, lambda: self.pstat.configure(text=f"лист {k}/{n}"))
                r = core.export_sheets_pdf_multi(pages_all, out, color=color,
                                                 log=log, progress=prog1)
                if r:
                    results.append(r)
            self.pdf_results = results
            self.after(0, self.log, "ГОТОВО: " + "; ".join(os.path.basename(r) for r in results), OK)
            self.after(0, lambda: self.pstat.configure(text="готово ✓", fg=OK))
            if results:
                self.after(0, lambda: messagebox.showinfo(
                    "Готово", "Создано PDF: %d\n\n%s" % (len(results), "\n".join(results))))
        except Exception as e:
            self.after(0, self.log, "ОШИБКА: " + str(e), ERR)
            self.after(0, lambda: self.pstat.configure(text="ошибка", fg=ERR))
            self.error_box("Ошибка", e)
        finally:
            self.after(0, self.pprog.pack_forget)
            self.after(0, lambda: self.pbtn.configure(state="normal", bg=ACCENT))

    # ---------------------------------------------------------- стр. Сравнение
    def _page_diff(self, parent):
        fr = tk.Frame(parent, bg=BG)
        self._header(fr, "Сравнение версий",
                     "что изменилось в чертежах относительно прежних перечней")
        c1 = Card(fr, "Прежние перечни (.docx)")
        c1.pack(fill="x", padx=16, pady=5)
        self.old_in_row = PathRow(c1.body, "Входные сигналы:", self.var_old_in,
                                  lambda: self._pick(self.var_old_in, [("Word", "*.docx")]))
        self.old_in_row.pack(fill="x", pady=2)
        self.old_out_row = PathRow(c1.body, "Выходные сигналы:", self.var_old_out,
                                   lambda: self._pick(self.var_old_out, [("Word", "*.docx")]))
        self.old_out_row.pack(fill="x", pady=2)
        tk.Label(c1.body, text="файлы можно перетащить сюда мышью",
                 font=FONT_SM, bg=CARD, fg=MUTED).pack(anchor="w", pady=(2, 0))
        bar = tk.Frame(fr, bg=BG)
        bar.pack(fill="x", padx=16, pady=(2, 4))
        self.dbtn = flat_btn(bar, "Сравнить", self.run_diff, primary=True)
        self.dbtn.pack(side="left")
        # Сверка отличается от сравнения: там две редакции и расхождения
        # ожидаемы, здесь источник один и любое расхождение — дефект выпуска.
        flat_btn(bar, "⚖ Сверить с чертежами", self.run_verify).pack(side="left", padx=(8, 0))
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
        ct = Card(fr, "Шаблоны перечней",
                  "по умолчанию встроенные; свой шаблон — любой прежний перечень .docx")
        ct.pack(fill="x", padx=16, pady=5)
        PathRow(ct.body, "Входные сигналы:", self.var_in,
                lambda: self._pick(self.var_in, [("Word", "*.docx")])).pack(fill="x", pady=2)
        PathRow(ct.body, "Выходные сигналы:", self.var_out,
                lambda: self._pick(self.var_out, [("Word", "*.docx")])).pack(fill="x", pady=2)
        tk.Label(ct.body, text="Пусто — используются шаблоны из папки templates "
                               "рядом с программой.",
                 font=FONT_SM, bg=CARD, fg=MUTED).pack(anchor="w", pady=(2, 0))

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

    @staticmethod
    def _is_warn(s):
        """Строка-замечание: то, что инженер должен увидеть обязательно."""
        low = s.lower()
        return ("⚠" in s or "ВНИМАНИЕ" in s or "ОШИБКА" in s
                or "не удалось" in low or "не найден" in low
                or "расхожден" in low or "ПЛОХО" in s)

    def _put_line(self, s, color=None):
        self.txt.configure(state="normal")
        self.txt.insert("end", s + "\n")
        if color:
            ln = int(self.txt.index("end-1c").split(".")[0]) - 1
            tag = f"c{ln}"
            self.txt.tag_add(tag, f"{ln}.0", f"{ln}.end")
            self.txt.tag_config(tag, foreground=color)
        self.txt.see("end")
        self.txt.configure(state="disabled")

    def log(self, s, color=None):
        import time as _t
        s = _t.strftime("%H:%M  ") + s
        self.log_lines.append((s, color))
        warn = self._is_warn(s)
        n_warn = sum(1 for t, _c in self.log_lines if self._is_warn(t))
        self.warn_stat.configure(text=("замечаний: %d" % n_warn) if n_warn else "",
                                 fg=(ERR if n_warn else MUTED))
        if self.only_warn.get() and not warn:
            return
        self._put_line(s, color or (WARN if warn else None))

    def _relog(self):
        """Перерисовывает журнал по текущему фильтру."""
        self.txt.configure(state="normal")
        self.txt.delete("1.0", "end")
        self.txt.configure(state="disabled")
        only = self.only_warn.get()
        for s, color in self.log_lines:
            warn = self._is_warn(s)
            if only and not warn:
                continue
            self._put_line(s, color or (WARN if warn else None))

    def save_log(self):
        """Сохраняет журнал целиком — независимо от того, что показано."""
        if not self.log_lines:
            messagebox.showinfo("Журнал", "Журнал пуст.")
            return
        import datetime
        p = filedialog.asksaveasfilename(
            defaultextension=".txt", filetypes=[("Текст", "*.txt")],
            initialfile="Журнал %s.txt" % datetime.datetime.now().strftime("%d.%m.%Y %H-%M"))
        if not p:
            return
        try:
            with open(p, "w", encoding="utf-8") as f:
                f.write("\n".join(s for s, _c in self.log_lines) + "\n")
            self.log("Журнал сохранён: " + p, OK)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def error_box(self, title, exc):
        """Показать ошибку из фонового потока.

        Текст берём сразу: имя из `except ... as exc` живёт только внутри
        своего блока, и лямбда, вызванная позже из очереди Tk, его уже не
        находит — вместо сообщения об ошибке всплывает NameError.
        """
        msg = str(exc) or type(exc).__name__
        self.after(0, lambda: messagebox.showerror(title, msg))

    def _logcb(self):
        return lambda s: self.after(0, self.log, s, WARN if "ВНИМАНИЕ" in s else None)

    def _save_cfg(self):
        save_cfg(dict(theme=self.theme,
                      template_in=self.var_in.get(), template_out=self.var_out.get(),
                      out_dir=self.var_dst.get(), update_fields=self.var_upd.get(), make_pdf=self.var_pdf.get(), revisions=self.var_rev.get(),
                      manifest=self.var_manifest.get(),
                      conv_dir=self.var_cdst.get(), conv_add=self.var_addlist.get(),
                      pdf_dir=self.var_pdfdst.get(), pdf_color=self.var_pdfcolor.get(),
                      pdf_one=self.var_pdfone.get(),
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
            self.error_box("Ошибка", e)
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
        # Резерв стоит считать отдельно: это единственное место, где видно
        # разницу между «канал свободен» и «программа не разобрала описание».
        res = sum(1 for _c, rows in self.sections for d in rows
                  if d["desc"] == "Резерв")
        notag = sum(1 for _c, rows in self.sections for d in rows
                    if d["desc"] != "Резерв" and not d["tag"])
        parts = [f"показано {shown} из {n}"]
        if res:
            parts.append(f"резерв {res} ({res * 100 // n if n else 0}%)")
        if notag:
            parts.append(f"без позиции {notag}")
        self.sig_stat.configure(text="   ·   ".join(parts),
                                fg=(ERR if res > n * 0.7 else MUTED))

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
        if (not (t_in or t_out) and not core.default_template("in")
                and not core.default_template("out")):
            messagebox.showwarning(
                "Нет шаблонов",
                "Не найдены встроенные шаблоны в папке templates рядом с программой. "
                "Укажите шаблон входных и/или выходных сигналов вручную.")
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
            if secs is None:
                # читаем один раз здесь, чтобы одни и те же данные пошли
                # и в сборку, и в сверку — иначе сверять было бы не с чем
                secs = core.read_sections(items, log=self._logcb())
            results = core.run(items, t_in or None, t_out or None, dst,
                               log=self._logcb(), update_fields=self.var_upd.get(),
                               sections=secs, make_pdf=self.var_pdf.get())
            self.after(0, self.log, "ГОТОВО: " + "; ".join(os.path.basename(r) for r in results), OK)
            # Сверка сразу после сборки: инженер должен видеть подтверждение,
            # что в документ попало всё, а не верить на слово.
            self.after(0, self.log, "Сверка перечней с чертежами...")
            try:
                ver = core.verify_perechen(secs, results, log=self._logcb())
                self.verify_result = ver
                if ver["ok"]:
                    self.after(0, self.log, "СВЕРКА ПРОЙДЕНА: расхождений нет", OK)
                else:
                    self.after(0, self.log,
                               "СВЕРКА: расхождений %d — подробности в журнале"
                               % len(ver["issues"]), ERR)
                # Отчёт сохраняем сразу: иначе он живёт до следующей сборки,
                # и потом не восстановить, когда именно разошлось.
                try:
                    import datetime
                    hist = os.path.join(dst, "Сверки")
                    os.makedirs(hist, exist_ok=True)
                    stamp = datetime.datetime.now().strftime("%Y-%m-%d %H-%M")
                    verdict = "сошлось" if ver["ok"] else "расхождений %d" % len(ver["issues"])
                    rp = os.path.join(hist, "Сверка %s — %s.xlsx" % (stamp, verdict))
                    core.export_verify_xlsx(ver, rp)
                    self.after(0, self.log, "отчёт сверки: " + rp, MUTED)
                except Exception as se:
                    self.after(0, self.log, "не удалось сохранить отчёт сверки: " + str(se), MUTED)
                self.after(0, lambda: self.btn_verify.pack(side="left", padx=(8, 0)))
            except Exception as ve:
                self.after(0, self.log, "сверка не выполнена: " + str(ve), MUTED)
            self.after(0, lambda: self.status.configure(text="Готово ✓", fg=OK))
            self.after(0, lambda: self.btn_open.pack(side="left", padx=(8, 0)))
        except Exception as e:
            self.after(0, self.log, "ОШИБКА: " + str(e), ERR)
            self.after(0, self.log, traceback.format_exc(), MUTED)
            self.after(0, lambda: self.status.configure(text="Ошибка", fg=ERR))
            self.error_box("Ошибка", e)
        finally:
            self.after(0, self.prog.stop)
            self.after(0, self.prog.pack_forget)
            self.after(0, lambda: self.btn.configure(state="normal", bg=ACCENT))

    def run_verify(self):
        """Сверяет уже готовые перечни с чертежами, ничего не перевыпуская."""
        docs = [p for p in (self.var_old_in.get().strip(),
                            self.var_old_out.get().strip()) if p]
        items = self._drawing_items()
        if not docs:
            messagebox.showwarning(
                "Сверка", "Укажите перечни (.docx) в полях выше — их и будем сверять.")
            return
        if not items:
            messagebox.showwarning(
                "Сверка", "Добавьте чертежи на странице «Сборка перечней»: "
                          "сверять перечни не с чем.")
            return
        self.dbtn.configure(state="disabled", bg=ACCENT_DIS)
        self.dprog.pack(side="left", padx=(8, 0))
        self.dprog.start(12)
        threading.Thread(target=self._verify_work, args=(items, docs), daemon=True).start()

    def _verify_work(self, items, docs):
        try:
            secs = None
            key = json.dumps(items, ensure_ascii=False)
            if self.sections is not None and self.sections_files == key:
                secs = self.sections
            if secs is None:
                secs = core.read_sections(items, log=self._logcb())
            res = core.verify_perechen(secs, docs, log=self._logcb())
            self.verify_result = res
            n_ok = res["counts"].get("совпало", 0)
            n_all = res["counts"].get("в_чертеже", 0)
            if res["ok"]:
                self.after(0, self.log, f"СВЕРКА ПРОЙДЕНА: сошлось {n_ok} из {n_all}", OK)
                self.after(0, lambda: self.diff_stat.configure(
                    text=f"сверка пройдена: {n_ok} из {n_all}", fg=OK))
            else:
                n = len(res["issues"])
                self.after(0, self.log, f"СВЕРКА: расхождений {n} — подробности в журнале", ERR)
                self.after(0, lambda: self.diff_stat.configure(
                    text=f"расхождений: {n}", fg=ERR))
            self.after(0, self._fill_verify_grid, res)
        except Exception as e:
            self.after(0, self.log, "ОШИБКА: " + str(e), ERR)
            self.error_box("Сверка", e)
        finally:
            self.after(0, self.dprog.stop)
            self.after(0, self.dprog.pack_forget)
            self.after(0, lambda: self.dbtn.configure(state="normal", bg=ACCENT))

    def _fill_verify_grid(self, res):
        """Показывает расхождения сверки в той же таблице, что и сравнение."""
        for i in self.dg.get_children():
            self.dg.delete(i)
        for kind, kc, io_, a, b in res["issues"]:
            tag = "удалено" if kind.startswith("нет в документе") else (
                "добавлено" if kind.startswith("нет в чертеже") else "изменено")
            self.dg.insert("", "end", values=(kind, io_, kc, a, b), tags=(tag,))

    def export_verify(self):
        """Сохраняет отчёт сверки в .xlsx."""
        if not self.verify_result:
            messagebox.showinfo("Сверка", "Сначала соберите перечни — сверка идёт автоматически.")
            return
        p = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                         filetypes=[("Excel", "*.xlsx")],
                                         initialfile="Сверка перечней с чертежами.xlsx")
        if not p:
            return
        try:
            core.export_verify_xlsx(self.verify_result, p)
            self.log(f"Отчёт сверки сохранён: {p}", OK)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def go_package(self):
        items = self._drawing_items()
        t_in, t_out = self.var_in.get().strip(), self.var_out.get().strip()
        dst = self.var_dst.get().strip()
        has_tpl = (t_in or t_out or core.default_template("in")
                   or core.default_template("out"))
        if not items or not has_tpl or not dst:
            messagebox.showwarning("Пакетная сборка",
                                   "Нужны: чертежи и папка результата. "
                                   "Шаблоны берутся встроенные, если не указаны свои.")
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
                dst = core.next_revision_dir(dst)
                os.makedirs(dst)
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
            spec_docx = os.path.join(dst, "Спецификация оборудования и материалов.docx")
            tpl_spec = core.default_template("spec")
            if tpl_spec:
                core.build_spec_docx(tpl_spec, spec_docx,
                                     core.spec_sections_for_cabs([n for n, _i in eq]),
                                     log=lambda t: self.after(0, self.log, t))
                self.after(0, self.log,
                           "Спецификация оборудования и материалов.docx — готово", OK)
                self.after(0, self.log, core.SPEC_EXCLUDED_NOTE)
            else:
                self.after(0, self.log,
                           "шаблон спецификации не найден — документ не собран", ERR)
            vedom = os.path.join(dst, "Ведомость комплектации шкафов.xlsx")
            core.export_spec_xlsx(eq, secs, vedom)
            self.after(0, self.log, "Ведомость комплектации шкафов.xlsx — готово", OK)
            pz_docx = os.path.join(dst, "Пояснительная записка.docx")
            tpl_pz = core.default_template("pz")
            if tpl_pz:
                core.build_pz_docx(tpl_pz, pz_docx, secs, eq,
                                   log=lambda t: self.after(0, self.log, t))
            else:
                core.export_pz_docx(secs, eq, pz_docx)
                self.after(0, self.log,
                           "шаблон записки не найден — документ без рамки и штампа", ERR)
            self.after(0, self.log, "Пояснительная записка.docx — готово", OK)
            # Паспорт выпуска: считает и оформляет его ядро, здесь только вызов.
            if self.var_manifest.get():
                try:
                    outs = list(results) + [
                        os.path.join(dst, "Сводка сигналов.xlsx"),
                        spec_docx, vedom,
                        pz_docx]
                    mp = core.write_manifest(
                        os.path.join(dst, "Паспорт выпуска.txt"),
                        sources=[(sec_name, pth) for pth, sec_name in items],
                        sections=secs,
                        out_files=[o for o in outs if os.path.exists(o)],
                        templates={"входные": t_in, "выходные": t_out},
                        verify=getattr(self, "verify_result", None),
                        checks=core.checks_report(secs),
                        app_version=APP_VER)
                    self.after(0, self.log, "Паспорт выпуска.txt — готово", OK)
                except Exception as me:
                    self.after(0, self.log, "паспорт выпуска не записан: " + str(me), ERR)
            self.after(0, self.log, "ПАКЕТ СОБРАН: " + dst, OK)
            self.after(0, lambda: self.status.configure(text="Пакет готов ✓", fg=OK))
            self.after(0, lambda: self.btn_open.pack(side="left", padx=(8, 0)))
        except Exception as e:
            self.after(0, self.log, "ОШИБКА: " + str(e), ERR)
            self.after(0, lambda: self.status.configure(text="Ошибка", fg=ERR))
            self.error_box("Ошибка", e)
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
            self.error_box("Ошибка", e)
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
            self.error_box("Ошибка", e)
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


REQUIRED = [
    ("ezdxf",      "ezdxf",       "чтение чертежей DXF"),
    ("docx",       "python-docx", "запись перечней в Word"),
    ("openpyxl",   "openpyxl",    "выгрузка в Excel"),
    ("matplotlib", "matplotlib",  "печать чертежей в PDF"),
]


def check_deps():
    """Список отсутствующих библиотек: [(пакет, зачем нужен), ...]."""
    import importlib
    miss = []
    for mod, pkg, why in REQUIRED:
        try:
            importlib.import_module(mod)
        except Exception:
            miss.append((pkg, why))
    return miss


def show_startup_error(title, text):
    """Показывает ошибку окном; если окно не поднять — печатает в консоль."""
    try:
        r = tk.Tk()
        r.withdraw()
        messagebox.showerror(title, text)
        r.destroy()
    except Exception:
        print(title + "\n\n" + text)


def install_crash_handler(app):
    """Необработанная ошибка не должна молча закрывать окно.

    В собранном .exe консоли нет, traceback уходит в никуда, и пользователь
    видит только исчезнувшую программу. Показываем понятное окно и пишем
    подробности в файл рядом с программой.
    """
    def handler(exc, val, tb):
        text = "".join(traceback.format_exception(exc, val, tb))
        path = os.path.join(APP_DIR, "ошибка.log")
        try:
            import datetime
            with open(path, "a", encoding="utf-8") as f:
                f.write("\n=== %s ===\n%s"
                        % (datetime.datetime.now().isoformat(" ", "seconds"), text))
        except Exception:
            path = "(не удалось записать файл лога)"
        messagebox.showerror(
            "Что-то пошло не так",
            "Программа столкнулась с ошибкой, но продолжает работать.\n\n"
            f"{val.__class__.__name__}: {val}\n\n"
            f"Подробности записаны в файл:\n{path}\n\n"
            "Если ошибка повторяется — пришлите этот файл разработчику.")
    app.report_callback_exception = handler


def run_selftest(dxf=None):
    """Проверка собранной программы без окна: --selftest [чертёж.dxf]

    Нужна потому, что запуск окна ещё не доказывает работоспособность: при
    сборке в .exe чаще всего отваливаются данные matplotlib и ezdxf, а видно
    это только когда доходит до дела. Отчёт пишется рядом с программой.
    """
    import tempfile, shutil, traceback, datetime
    lines = []
    def say(t=""):
        lines.append(str(t))
    say("Проверка программы — %s" % datetime.datetime.now().strftime("%d.%m.%Y %H:%M"))
    say("Файл: %s" % os.path.abspath(sys.argv[0]))
    say("=" * 60)
    bad = 0

    def step(name, fn):
        nonlocal bad
        try:
            r = fn()
            say("  ок    %s%s" % (name, ("  — " + str(r)) if r else ""))
        except Exception as e:
            bad += 1
            say("  ПЛОХО %s: %s: %s" % (name, type(e).__name__, e))
            say("        " + traceback.format_exc().strip().splitlines()[-1])

    step("библиотеки на месте", lambda: "нет: " + ", ".join(p for p, _ in check_deps())
         if check_deps() else "все")
    step("встроенные шаблоны", lambda: os.path.basename(core.default_template("in")) or "НЕ НАЙДЕН")
    step("словарь сокращений", lambda: "%d пар" % len(core.load_abbrev()))
    step("шрифт для замены SHX", lambda: core.pick_subst_font() or "не найден")
    step("конвертер DWG", lambda: " ".join(core.find_converter()) or "не найден")

    def matplotlib_ok():
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        fig = plt.figure(figsize=(1, 1)); plt.close(fig)
        return matplotlib.get_data_path()
    step("matplotlib рисует", matplotlib_ok)

    def ezdxf_ok():
        import ezdxf
        d = ezdxf.new(); d.modelspace().add_text("проверка")
        return "версия " + ezdxf.__version__
    step("ezdxf читает и пишет", ezdxf_ok)

    if dxf and os.path.exists(dxf):
        tmp = tempfile.mkdtemp(prefix="selftest_")
        try:
            rows = []
            step("разбор чертежа", lambda: "каналов: %d" % len(
                rows.extend(core.extract(dxf, log=lambda *a: None)) or rows))
            if rows:
                secs = [("ШКАФ ПРОВЕРКА", rows)]
                built = []
                step("сборка перечней", lambda: "документов: %d" % len(
                    built.extend(core.run([], "", "", tmp, sections=secs,
                                          log=lambda *a: None)) or built))
                if built:
                    step("сверка с чертежом", lambda: (
                        "сошлось %d" % core.verify_perechen(secs, built, log=lambda *a: None)
                        ["counts"]["совпало"]))
            step("печать PDF", lambda: "листов: %d" % len(
                core.export_sheets_pdf(dxf, os.path.join(tmp, "t.pdf"),
                                       log=lambda *a: None)[1]))
        finally:
            shutil.rmtree(tmp, ignore_errors=True)
    else:
        say("  (чертёж не передан — разбор, сборка и печать PDF не проверялись)")
        say("  запуск с чертежом:  Perechni-signalov.exe --selftest C:\\путь\\чертёж.dxf")

    say("=" * 60)
    say("ИТОГ: " + ("всё работает" if not bad else "замечаний: %d" % bad))
    text = "\n".join(lines)
    out = os.path.join(APP_DIR, "проверка сборки.txt")
    try:
        with open(out, "w", encoding="utf-8") as f:
            f.write(text + "\n")
    except Exception:
        out = "(не удалось записать файл)"
    print(text)
    # В собранном .exe консоли нет, поэтому отчёт надо показать. Но окно с
    # кнопкой «ОК» блокирует выполнение до нажатия — а проверку сборки часто
    # запускают из скрипта, где нажимать некому. Поэтому просто открываем
    # файл отчёта в блокноте: видно так же, а выполнение не встаёт.
    if getattr(sys, "frozen", False) and os.path.exists(out):
        try:
            os.startfile(out)          # не блокирует
        except Exception:
            pass
    return 1 if bad else 0


def main():
    if "--selftest" in sys.argv:
        i = sys.argv.index("--selftest")
        arg = sys.argv[i + 1] if len(sys.argv) > i + 1 else None
        return run_selftest(arg)
    miss = check_deps()
    if miss:
        lines = "\n".join(f"  - {pkg} — {why}" for pkg, why in miss)
        pkgs = " ".join(pkg for pkg, _why in miss)
        show_startup_error(
            "Не хватает библиотек",
            "Программа не может запуститься: не установлены нужные компоненты.\n\n"
            f"{lines}\n\n"
            "Как исправить — выполнить в командной строке:\n\n"
            f"    pip install {pkgs}\n\n"
            "Если вы пользуетесь готовым .exe — сообщите разработчику, "
            "сборка неполная.")
        return 1
    try:
        app = App()
    except Exception as e:
        show_startup_error(
            "Не удалось запустить программу",
            f"{e.__class__.__name__}: {e}\n\n" + traceback.format_exc()[-1200:])
        return 1
    install_crash_handler(app)
    app.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
