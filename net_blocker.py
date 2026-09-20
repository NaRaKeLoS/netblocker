"""
NetBlocker v1.1.0 - odcina internet wybranej aplikacji na Windows 11
- blokada na stałe
- LAG: krótkie odcięcie internetu
- motywy kolorystyczne
- własne skróty klawiszowe
- animacje
- automatyczne sprawdzanie aktualizacji
- aktualizacja programu jednym kliknięciem

Wymagania:
    pip install customtkinter psutil pywin32 keyboard requests

Uruchomienie:
    python net_blocker.py

Konfiguracja:
    %APPDATA%\\NetBlocker\\config.json
"""

import copy
import ctypes
import json
import os
import queue
import subprocess
import sys
import threading
import time
import tkinter.messagebox as messagebox
import webbrowser

import customtkinter as ctk
import keyboard
import psutil
import win32com.client


# ---------------------------------------------------------------------
# WERSJA I AKTUALIZACJE
# ---------------------------------------------------------------------

VERSION = "1.1.0"

# Plik version.txt na GitHubie
VERSION_CHECK_URL = (
    "https://raw.githubusercontent.com/NaRaKeLoS/netblocker/"
    "refs/heads/main/version.txt"
)

# Strona projektu - awaryjnie, jeśli version.txt nie poda linku
DOWNLOAD_PAGE_URL = "https://github.com/NaRaKeLoS/netblocker"


PREFIX = "NetBlocker_"

DIR_IN, DIR_OUT = 1, 2
ACTION_BLOCK = 0
PROFILES_ALL = 0x7FFFFFFF


# =====================================================================
# WERSJE
# =====================================================================

def parse_version(v):
    try:
        return tuple(int(x) for x in v.strip().split("."))
    except Exception:
        return (0,)


def check_for_update(result_queue):
    """
    Sprawdza version.txt.
    
    Pierwsza linia:
        numer wersji

    Druga linia:
        bezpośredni link do nowego net_blocker.py
    """

    try:
        import requests
        import random

        bust = f"?_={int(time.time())}{random.randint(1000, 9999)}"

        r = requests.get(
            VERSION_CHECK_URL + bust,
            timeout=5,
            headers={
                "Cache-Control": "no-cache",
                "Pragma": "no-cache"
            }
        )

        r.raise_for_status()

        lines = [
            ln.strip()
            for ln in r.text.splitlines()
            if ln.strip()
        ]

        if not lines:
            return

        remote_version = lines[0]

        if len(lines) >= 2:
            download_url = lines[1]
        else:
            download_url = DOWNLOAD_PAGE_URL

        if parse_version(remote_version) > parse_version(VERSION):
            result_queue.put(
                (
                    "update_available",
                    remote_version,
                    download_url
                )
            )

    except Exception:
        # Brak internetu / GitHub niedostępny / błąd pliku itd.
        pass


# =====================================================================
# KONFIGURACJA
# =====================================================================

CONFIG_DIR = os.path.join(
    os.environ.get("APPDATA", "."),
    "NetBlocker"
)

CONFIG_PATH = os.path.join(
    CONFIG_DIR,
    "config.json"
)


DEFAULT_CFG = {
    "theme": "Ciemny",
    "hotkeys": {
        "toggle": "ctrl+alt+b",
        "lag": "ctrl+alt+l"
    },
    "lag_ms": 500,
    "topmost": False,
    "last_app": None,
    "skip_version": None
}


HOTKEY_LABELS = {
    "toggle": "Odetnij / przywróć",
    "lag": "Lag"
}


def load_config():
    cfg = copy.deepcopy(DEFAULT_CFG)

    file_missing_or_broken = True

    try:
        with open(CONFIG_PATH, encoding="utf-8") as f:
            data = json.load(f)

        for k in (
            "theme",
            "lag_ms",
            "topmost",
            "last_app",
            "skip_version"
        ):
            if k in data:
                cfg[k] = data[k]

        cfg["hotkeys"].update(
            data.get("hotkeys", {})
        )

        file_missing_or_broken = False

    except Exception:
        pass

    if file_missing_or_broken:
        save_config(cfg)

    return cfg


def save_config(cfg):
    try:
        os.makedirs(CONFIG_DIR, exist_ok=True)

        tmp_path = CONFIG_PATH + ".tmp"

        with open(
            tmp_path,
            "w",
            encoding="utf-8"
        ) as f:
            json.dump(
                cfg,
                f,
                indent=2,
                ensure_ascii=False
            )

        os.replace(
            tmp_path,
            CONFIG_PATH
        )

        return True

    except Exception:
        return False


def fmt_hotkey(hk):
    return hk.upper().replace("+", " + ") if hk else ""


# =====================================================================
# MOTYWY
# =====================================================================

def _theme(
    mode,
    bg,
    card,
    text,
    muted,
    accent,
    accent_h,
    red="#d64545",
    red_h="#b53737",
    green="#2fa36b",
    green_h="#26855a",
    orange="#e08a1e",
    orange_h="#bd7418"
):
    return dict(
        mode=mode,
        bg=bg,
        card=card,
        text=text,
        muted=muted,
        accent=accent,
        accent_h=accent_h,
        red=red,
        red_h=red_h,
        green=green,
        green_h=green_h,
        orange=orange,
        orange_h=orange_h
    )


THEMES = {
    "Ciemny": _theme(
        "dark",
        "#15171c",
        "#1f2229",
        "#e8eaed",
        "#8b919c",
        "#3b82f6",
        "#2f6bd0"
    ),

    "Jasny": _theme(
        "light",
        "#eef0f4",
        "#ffffff",
        "#1c1f26",
        "#6b7280",
        "#2563eb",
        "#1d4fbf"
    ),

    "Północ": _theme(
        "dark",
        "#120f1f",
        "#1c1830",
        "#ece9f7",
        "#8f88ad",
        "#8b5cf6",
        "#7444d8"
    ),

    "Ocean": _theme(
        "dark",
        "#0b1a24",
        "#12283a",
        "#e2f1f8",
        "#7d9db0",
        "#06b6d4",
        "#0592ab"
    ),

    "Zachód słońca": _theme(
        "dark",
        "#1f1414",
        "#2c1c1c",
        "#f6e9e4",
        "#a58880",
        "#f97316",
        "#d5600e"
    ),

    "Hacker": _theme(
        "dark",
        "#050805",
        "#0d150d",
        "#39ff14",
        "#2e9d2e",
        "#00c853",
        "#00a044",
        red="#ff3b3b",
        red_h="#d62f2f",
        green="#00e676",
        green_h="#00b85e",
        orange="#ffb300",
        orange_h="#d69600"
    )
}


# =====================================================================
# ADMIN
# =====================================================================

def is_admin():
    try:
        return bool(
            ctypes.windll.shell32.IsUserAnAdmin()
        )
    except Exception:
        return False


def relaunch_as_admin():
    params = " ".join(
        f'"{a}"'
        for a in sys.argv
    )

    ctypes.windll.shell32.ShellExecuteW(
        None,
        "runas",
        sys.executable,
        params,
        None,
        1
    )

    sys.exit(0)


# =====================================================================
# WINDOWS FIREWALL
# =====================================================================

policy = win32com.client.Dispatch(
    "HNetCfg.FwPolicy2"
)


def rule_names(path: str, kind: str):
    p = path.lower()

    return [
        f"{PREFIX}{kind}_out_{p}",
        f"{PREFIX}{kind}_in_{p}"
    ]


def rule_exists(name: str):
    try:
        policy.Rules.Item(name)
        return True
    except Exception:
        return False


def add_rule(
    name,
    path,
    direction,
    enabled
):
    r = win32com.client.Dispatch(
        "HNetCfg.FWRule"
    )

    r.Name = name
    r.ApplicationName = path
    r.Action = ACTION_BLOCK
    r.Direction = direction
    r.Profiles = PROFILES_ALL
    r.Grouping = "NetBlocker"
    r.Enabled = enabled

    policy.Rules.Add(r)


def ensure_rules(
    path,
    kind,
    enabled
):
    for name, direction in zip(
        rule_names(path, kind),
        (DIR_OUT, DIR_IN)
    ):
        if rule_exists(name):
            policy.Rules.Item(name).Enabled = enabled
        else:
            add_rule(
                name,
                path,
                direction,
                enabled
            )


def set_enabled(
    path,
    kind,
    enabled
):
    for name in rule_names(path, kind):
        try:
            policy.Rules.Item(name).Enabled = enabled
        except Exception:
            pass


def remove_rules(
    path,
    kind
):
    for name in rule_names(path, kind):
        while rule_exists(name):
            try:
                policy.Rules.Remove(name)
            except Exception:
                break


def is_blocked(path):
    try:
        return bool(
            policy.Rules.Item(
                rule_names(path, "block")[0]
            ).Enabled
        )
    except Exception:
        return False


def cleanup_stale_lag_rules():
    try:
        stale = [
            r.Name
            for r in policy.Rules
            if (
                r.Name or ""
            ).startswith(
                PREFIX + "lag_"
            )
        ]

        for name in stale:
            policy.Rules.Remove(name)

    except Exception:
        pass


# =====================================================================
# PROCESY
# =====================================================================

def list_apps():
    found = {}

    for p in psutil.process_iter(
        ["name", "exe"]
    ):
        exe = p.info.get("exe")
        name = p.info.get("name")

        if (
            not exe
            or not name
            or exe.lower().startswith(
                r"c:\windows"
            )
        ):
            continue

        found[exe.lower()] = (
            name,
            exe
        )

    names = [
        n
        for n, _ in found.values()
    ]

    result = {}

    for name, exe in sorted(
        found.values(),
        key=lambda x: x[0].lower()
    ):
        label = name

        if names.count(name) > 1:
            label = (
                f"{name}  "
                f"({os.path.basename(os.path.dirname(exe))})"
            )

        result[label] = exe

    return result


# =====================================================================
# GUI
# =====================================================================

class App(ctk.CTk):

    def __init__(self):
        super().__init__()

        self.cfg = load_config()

        self.T = THEMES.get(
            self.cfg["theme"],
            THEMES["Ciemny"]
        )

        self.apps = {}
        self.current_label = None

        self.lag_paths = set()
        self.lag_running = False

        self.recording = None
        self.hk_btns = {}

        self._last_toggle = 0.0
        self._pulse_job = None

        self.q = queue.Queue()

        self.title(
            f"NetBlocker v{VERSION}"
        )

        self.geometry(
            "460x680"
        )

        self.resizable(
            False,
            False
        )

        self.attributes(
            "-topmost",
            bool(self.cfg["topmost"])
        )

        self.attributes(
            "-alpha",
            0.0
        )

        self.build_ui("main")

        self.register_hotkeys()

        self.protocol(
            "WM_DELETE_WINDOW",
            self.on_close
        )

        self.after(
            30,
            self.poll_queue
        )

        self.fade_in()

        threading.Thread(
            target=check_for_update,
            args=(self.q,),
            daemon=True
        ).start()


    # -----------------------------------------------------------------
    # ANIMACJE
    # -----------------------------------------------------------------

    def fade_in(self, step=0.0):
        step = min(
            step + 0.08,
            1.0
        )

        try:
            self.attributes(
                "-alpha",
                step
            )
        except Exception:
            return

        if step < 1.0:
            self.after(
                15,
                lambda: self.fade_in(step)
            )


    def pulse_button(
        self,
        btn,
        color_a,
        color_b,
        times=6,
        delay=110
    ):
        if self._pulse_job:
            try:
                self.after_cancel(
                    self._pulse_job
                )
            except Exception:
                pass

            self._pulse_job = None

        def step(i):
            if not btn.winfo_exists():
                return

            try:
                btn.configure(
                    fg_color=(
                        color_a
                        if i % 2 == 0
                        else color_b
                    )
                )
            except Exception:
                return

            if i < times:
                self._pulse_job = self.after(
                    delay,
                    lambda: step(i + 1)
                )
            else:
                self._pulse_job = None

        step(0)


    def slide_status(
        self,
        label,
        text,
        color
    ):
        try:
            label.configure(
                text=text,
                text_color=color,
                font=ctk.CTkFont(
                    size=16,
                    weight="bold"
                )
            )

            self.after(
                160,
                lambda: label.configure(
                    font=ctk.CTkFont(
                        size=14,
                        weight="bold"
                    )
                )
            )

        except Exception:
            pass


    # -----------------------------------------------------------------
    # UI
    # -----------------------------------------------------------------

    def lbl(
        self,
        parent,
        text,
        size=13,
        bold=False,
        color=None,
        **kw
    ):
        return ctk.CTkLabel(
            parent,
            text=text,
            text_color=color or self.T["text"],
            font=ctk.CTkFont(
                size=size,
                weight=(
                    "bold"
                    if bold
                    else "normal"
                )
            ),
            **kw
        )


    def card(
        self,
        parent,
        title
    ):
        c = ctk.CTkFrame(
            parent,
            corner_radius=14,
            fg_color=self.T["card"]
        )

        c.pack(
            fill="x",
            padx=22,
            pady=6
        )

        self.lbl(
            c,
            title,
            11,
            True,
            self.T["muted"]
        ).pack(
            anchor="w",
            padx=16,
            pady=(12, 2)
        )

        return c


    def build_ui(self, page="main"):

        if hasattr(self, "combo"):
            try:
                self.current_label = (
                    self.combo.get()
                )
            except Exception:
                pass

        for w in self.winfo_children():
            w.destroy()

        self.hk_btns = {}

        ctk.set_appearance_mode(
            self.T["mode"]
        )

        self.configure(
            fg_color=self.T["bg"]
        )

        self.main_frame = ctk.CTkFrame(
            self,
            fg_color="transparent"
        )

        self.settings_frame = ctk.CTkFrame(
            self,
            fg_color="transparent"
        )

        self.build_main()
        self.build_settings()

        self.show(page)
        self.refresh()


    def show(self, page):

        self.main_frame.pack_forget()
        self.settings_frame.pack_forget()

        frame = (
            self.main_frame
            if page == "main"
            else self.settings_frame
        )

        frame.pack(
            fill="both",
            expand=True
        )


    # -----------------------------------------------------------------
    # MAIN
    # -----------------------------------------------------------------

    def build_main(self):

        T = self.T
        f = self.main_frame

        header = ctk.CTkFrame(
            f,
            fg_color="transparent"
        )

        header.pack(
            fill="x",
            padx=22,
            pady=(20, 4)
        )

        titles = ctk.CTkFrame(
            header,
            fg_color="transparent"
        )

        titles.pack(side="left")

        self.lbl(
            titles,
            "NetBlocker",
            26,
            True
        ).pack(anchor="w")

        self.lbl(
            titles,
            f"v{VERSION} — odcinaj internet wybranym aplikacjom",
            12,
            False,
            T["muted"]
        ).pack(anchor="w")

        ctk.CTkButton(
            header,
            text="⚙",
            width=42,
            height=42,
            font=ctk.CTkFont(size=18),
            fg_color=T["card"],
            hover_color=T["accent"],
            text_color=T["text"],
            command=lambda: self.show("settings")
        ).pack(side="right")

        self._header_widget = header

        # UPDATE BAR

        self.update_bar = ctk.CTkFrame(
            f,
            corner_radius=10,
            fg_color=T["accent"]
        )

        self.update_lbl = self.lbl(
            self.update_bar,
            "",
            12,
            True,
            "#ffffff"
        )

        self.update_lbl.pack(
            side="left",
            padx=12,
            pady=8
        )

        ctk.CTkButton(
            self.update_bar,
            text="Pobierz",
            width=80,
            height=28,
            fg_color="#ffffff",
            text_color=T["accent"],
            hover_color="#e5e5e5",
            command=self.open_download
        ).pack(
            side="right",
            padx=6,
            pady=6
        )

        ctk.CTkButton(
            self.update_bar,
            text="Pomiń",
            width=70,
            height=28,
            fg_color="transparent",
            text_color="#ffffff",
            hover_color=T["accent_h"],
            command=self.dismiss_update
        ).pack(
            side="right",
            padx=0,
            pady=6
        )

        # APLIKACJA

        c1 = self.card(
            f,
            "APLIKACJA"
        )

        row = ctk.CTkFrame(
            c1,
            fg_color="transparent"
        )

        row.pack(
            fill="x",
            padx=12,
            pady=(0, 4)
        )

        self.combo = ctk.CTkComboBox(
            row,
            values=["(brak)"],
            state="readonly",
            height=36,
            command=self.on_select,
            fg_color=T["bg"],
            border_color=T["bg"],
            text_color=T["text"],
            button_color=T["accent"],
            button_hover_color=T["accent_h"],
            dropdown_fg_color=T["card"],
            dropdown_text_color=T["text"],
            dropdown_hover_color=T["accent"]
        )

        self.combo.pack(
            side="left",
            fill="x",
            expand=True,
            padx=(4, 6)
        )

        ctk.CTkButton(
            row,
            text="⟳",
            width=40,
            height=36,
            fg_color=T["accent"],
            hover_color=T["accent_h"],
            text_color="#ffffff",
            command=self.refresh
        ).pack(
            side="left",
            padx=(0, 4)
        )

        self.path_lbl = self.lbl(
            c1,
            "",
            11,
            False,
            T["muted"],
            wraplength=390,
            justify="left"
        )

        self.path_lbl.pack(
            anchor="w",
            padx=16,
            pady=(0, 12)
        )

        # BLOKADA

        c2 = self.card(
            f,
            "BLOKADA"
        )

        srow = ctk.CTkFrame(
            c2,
            fg_color="transparent"
        )

        srow.pack(
            fill="x",
            padx=16,
            pady=(0, 8)
        )

        self.dot = self.lbl(
            srow,
            "●",
            18,
            False,
            T["muted"]
        )

        self.dot.pack(side="left")

        self.status = self.lbl(
            srow,
            "Wybierz aplikację",
            14,
            True
        )

        self.status.pack(
            side="left",
            padx=8
        )

        self.toggle_btn = ctk.CTkButton(
            c2,
            text="Odetnij internet",
            height=46,
            font=ctk.CTkFont(
                size=15,
                weight="bold"
            ),
            fg_color=T["red"],
            hover_color=T["red_h"],
            text_color="#ffffff",
            command=lambda: self.toggle()
        )

        self.toggle_btn.pack(
            fill="x",
            padx=16,
            pady=(0, 16)
        )

        # LAG

        c3 = self.card(
            f,
            "LAG"
        )

        lrow = ctk.CTkFrame(
            c3,
            fg_color="transparent"
        )

        lrow.pack(
            fill="x",
            padx=16
        )

        self.lbl(
            lrow,
            "Czas odcięcia:"
        ).pack(side="left")

        self.ms_lbl = self.lbl(
            lrow,
            f"{int(self.cfg['lag_ms'])} ms",
            13,
            True
        )

        self.ms_lbl.pack(
            side="right"
        )

        self.slider = ctk.CTkSlider(
            c3,
            from_=100,
            to=3000,
            number_of_steps=29,
            progress_color=T["orange"],
            button_color=T["orange"],
            button_hover_color=T["orange_h"],
            fg_color=T["bg"],
            command=self.on_slider
        )

        self.slider.set(
            self.cfg["lag_ms"]
        )

        self.slider.pack(
            fill="x",
            padx=16,
            pady=(6, 10)
        )

        self.lag_btn = ctk.CTkButton(
            c3,
            text=self.lag_text(),
            height=46,
            font=ctk.CTkFont(
                size=15,
                weight="bold"
            ),
            fg_color=T["orange"],
            hover_color=T["orange_h"],
            text_color="#ffffff",
            command=lambda: self.do_lag()
        )

        self.lag_btn.pack(
            fill="x",
            padx=16,
            pady=(0, 16)
        )

        self.info = self.lbl(
            f,
            "",
            12,
            False,
            T["muted"]
        )

        self.info.pack(
            pady=(6, 0)
        )

        self.lbl(
            f,
            "Blokada na stałe zostaje w zaporze po zamknięciu programu.",
            10,
            False,
            T["muted"]
        ).pack(
            pady=(2, 0)
        )


    # -----------------------------------------------------------------
    # SETTINGS
    # -----------------------------------------------------------------

    def build_settings(self):

        T = self.T
        f = self.settings_frame

        header = ctk.CTkFrame(
            f,
            fg_color="transparent"
        )

        header.pack(
            fill="x",
            padx=22,
            pady=(20, 4)
        )

        ctk.CTkButton(
            header,
            text="← Wróć",
            width=80,
            height=36,
            fg_color=T["card"],
            hover_color=T["accent"],
            text_color=T["text"],
            command=lambda: self.show("main")
        ).pack(side="left")

        self.lbl(
            header,
            "Ustawienia",
            22,
            True
        ).pack(
            side="left",
            padx=14
        )

        # MOTYW

        c1 = self.card(
            f,
            "MOTYW"
        )

        grid = ctk.CTkFrame(
            c1,
            fg_color="transparent"
        )

        grid.pack(
            fill="x",
            padx=12,
            pady=(0, 12)
        )

        grid.columnconfigure(
            (0, 1),
            weight=1
        )

        for i, (name, t) in enumerate(
            THEMES.items()
        ):

            selected = (
                name == self.cfg["theme"]
            )

            ctk.CTkButton(
                grid,
                text=(
                    ("✓ " if selected else "")
                    + name
                ),
                height=38,
                fg_color=t["bg"],
                hover_color=t["card"],
                text_color=t["text"],
                border_width=(
                    3 if selected else 1
                ),
                border_color=t["accent"],
                command=lambda n=name: self.set_theme(n)
            ).grid(
                row=i // 2,
                column=i % 2,
                padx=4,
                pady=4,
                sticky="ew"
            )

        # HOTKEYE

        c2 = self.card(
            f,
            "SKRÓTY KLAWISZOWE (GLOBALNE)"
        )

        for action, label in HOTKEY_LABELS.items():

            row = ctk.CTkFrame(
                c2,
                fg_color="transparent"
            )

            row.pack(
                fill="x",
                padx=16,
                pady=4
            )

            self.lbl(
                row,
                label,
                13
            ).pack(side="left")

            ctk.CTkButton(
                row,
                text="✕",
                width=34,
                height=34,
                fg_color=T["bg"],
                hover_color=T["red"],
                text_color=T["text"],
                command=lambda a=action: self.clear_hotkey(a)
            ).pack(
                side="right"
            )

            btn = ctk.CTkButton(
                row,
                text="",
                width=190,
                height=34,
                fg_color=T["bg"],
                hover_color=T["accent"],
                text_color=T["text"],
                command=lambda a=action: self.start_record(a)
            )

            btn.pack(
                side="right",
                padx=(0, 6)
            )

            self.hk_btns[action] = btn

        self.lbl(
            c2,
            "Kliknij pole skrótu i naciśnij nową kombinację klawiszy.\n"
            "Esc anuluje. Skróty działają na aplikację wybraną na liście.",
            11,
            False,
            T["muted"],
            justify="left"
        ).pack(
            anchor="w",
            padx=16,
            pady=(6, 14)
        )

        self.refresh_hk_buttons()

        # INNE

        c3 = self.card(
            f,
            "INNE"
        )

        sw = ctk.CTkSwitch(
            c3,
            text="Okno zawsze na wierzchu",
            text_color=T["text"],
            progress_color=T["accent"],
            command=self.on_topmost
        )

        if self.cfg["topmost"]:
            sw.select()

        self.topmost_switch = sw

        sw.pack(
            anchor="w",
            padx=16,
            pady=(4, 8)
        )

        # O PROGRAMIE

        c4 = self.card(
            f,
            "O PROGRAMIE"
        )

        self.version_lbl = self.lbl(
            c4,
            f"Wersja: {VERSION}",
            13
        )

        self.version_lbl.pack(
            anchor="w",
            padx=16,
            pady=(0, 4)
        )

        ctk.CTkButton(
            c4,
            text="Sprawdź aktualizacje",
            height=36,
            fg_color=T["accent"],
            hover_color=T["accent_h"],
            text_color="#ffffff",
            command=self.manual_check_update
        ).pack(
            fill="x",
            padx=16,
            pady=(4, 16)
        )


    def set_theme(self, name):

        self.cfg["theme"] = name
        self.T = THEMES[name]

        save_config(
            self.cfg
        )

        self.after(
            10,
            lambda: self.build_ui("settings")
        )


    def on_topmost(self):

        self.cfg["topmost"] = bool(
            self.topmost_switch.get()
        )

        self.attributes(
            "-topmost",
            self.cfg["topmost"]
        )

        save_config(
            self.cfg
        )


    # -----------------------------------------------------------------
    # AKTUALIZACJE
    # -----------------------------------------------------------------

    def manual_check_update(self):

        self.version_lbl.configure(
            text=f"Wersja: {VERSION}  (sprawdzam...)"
        )

        threading.Thread(
            target=self._manual_check_worker,
            daemon=True
        ).start()


    def _manual_check_worker(self):

        local_q = queue.Queue()

        check_for_update(
            local_q
        )

        try:
            item = local_q.get_nowait()
        except queue.Empty:
            item = None

        self.q.put(
            (
                "manual_check_result",
                item
            )
        )


    def show_update_bar(
        self,
        remote_version,
        url
    ):

        self._update_url = url

        self.update_lbl.configure(
            text=(
                f"Dostępna nowa wersja "
                f"{remote_version} "
                f"(masz {VERSION})"
            )
        )

        if not self.update_bar.winfo_ismapped():

            self.update_bar.pack(
                fill="x",
                padx=22,
                pady=(0, 6),
                after=self._header_widget
            )


    def open_download(self):

        url = getattr(
            self,
            "_update_url",
            None
        )

        if not url:
            messagebox.showerror(
                "NetBlocker",
                "Nie znaleziono adresu aktualizacji."
            )
            return

        current_file = os.path.abspath(
            sys.argv[0]
        )

        if not current_file.lower().endswith(
            ".py"
        ):
            messagebox.showerror(
                "NetBlocker",
                "Automatyczna aktualizacja jest "
                "dostępna tylko dla wersji .py."
            )
            return

        self.update_lbl.configure(
            text="Pobieranie aktualizacji..."
        )

        threading.Thread(
            target=self._download_update_worker,
            args=(
                url,
                current_file
            ),
            daemon=True
        ).start()


    def _download_update_worker(
        self,
        url,
        current_file
    ):

        try:
            import requests

            r = requests.get(
                url,
                timeout=15,
                headers={
                    "Cache-Control": "no-cache",
                    "Pragma": "no-cache"
                }
            )

            r.raise_for_status()

            new_code = r.content

            if len(new_code) < 1000:
                raise RuntimeError(
                    "Pobrany plik jest podejrzanie mały."
                )

            try:
                compile(
                    new_code.decode("utf-8"),
                    current_file,
                    "exec"
                )

            except Exception:
                raise RuntimeError(
                    "Pobrany plik nie jest poprawnym "
                    "skryptem Python."
                )

            temp_file = (
                current_file
                + ".update"
            )

            with open(
                temp_file,
                "wb"
            ) as f:
                f.write(new_code)

            self.q.put(
                (
                    "update_downloaded",
                    current_file,
                    temp_file
                )
            )

        except Exception as e:

            self.q.put(
                (
                    "update_error",
                    str(e)
                )
            )


    def install_update(
        self,
        current_file,
        temp_file
    ):

        try:

            if os.path.exists(
                current_file
            ):
                os.remove(
                    current_file
                )

            os.replace(
                temp_file,
                current_file
            )

            subprocess.Popen(
                [
                    sys.executable,
                    current_file
                ],
                cwd=os.path.dirname(
                    current_file
                )
            )

            self.destroy()

        except Exception as e:

            try:
                if os.path.exists(
                    temp_file
                ):
                    os.remove(
                        temp_file
                    )
            except Exception:
                pass

            messagebox.showerror(
                "NetBlocker",
                "Nie udało się zainstalować "
                f"aktualizacji:\n{e}"
            )


    def dismiss_update(self):

        self.cfg["skip_version"] = getattr(
            self,
            "_pending_version",
            None
        )

        save_config(
            self.cfg
        )

        self.update_bar.pack_forget()


    # -----------------------------------------------------------------
    # HOTKEYE
    # -----------------------------------------------------------------

    def register_hotkeys(self):

        try:
            keyboard.clear_all_hotkeys()
        except Exception:
            pass

        for action, hk in self.cfg["hotkeys"].items():

            if not hk:
                continue

            try:
                keyboard.add_hotkey(
                    hk,
                    lambda a=action:
                    self.q.put(a)
                )
            except Exception:
                pass


    def start_record(self, action):

        if self.recording:
            return

        self.recording = action

        try:
            keyboard.clear_all_hotkeys()
        except Exception:
            pass

        self.refresh_hk_buttons()

        def worker():

            try:
                hk = keyboard.read_hotkey(
                    suppress=False
                )
            except Exception:
                hk = None

            self.q.put(
                (
                    "recorded",
                    action,
                    hk
                )
            )

        threading.Thread(
            target=worker,
            daemon=True
        ).start()


    def on_recorded(
        self,
        action,
        hk
    ):

        self.recording = None

        if hk and hk != "esc":

            other = (
                "lag"
                if action == "toggle"
                else "toggle"
            )

            if (
                self.cfg["hotkeys"].get(other)
                == hk
            ):

                messagebox.showwarning(
                    "NetBlocker",
                    f"Skrót {fmt_hotkey(hk)} "
                    "jest już użyty w innej akcji."
                )

            else:

                self.cfg["hotkeys"][action] = hk

                save_config(
                    self.cfg
                )

        self.register_hotkeys()
        self.refresh_hk_buttons()
        self.update_status()
        self.set_lag_btn(False)


    def clear_hotkey(self, action):

        if self.recording:
            return

        self.cfg["hotkeys"][action] = None

        save_config(
            self.cfg
        )

        self.register_hotkeys()
        self.refresh_hk_buttons()
        self.update_status()
        self.set_lag_btn(False)


    def refresh_hk_buttons(self):

        for action, btn in self.hk_btns.items():

            text = (
                "Naciśnij skrót…"
                if self.recording == action
                else (
                    fmt_hotkey(
                        self.cfg["hotkeys"].get(action)
                    )
                    or "— brak —"
                )
            )

            try:
                btn.configure(
                    text=text
                )
            except Exception:
                pass


    # -----------------------------------------------------------------
    # KOLEJKA
    # -----------------------------------------------------------------

    def poll_queue(self):

        try:

            while True:

                item = self.q.get_nowait()

                if (
                    isinstance(item, tuple)
                    and item[0] == "recorded"
                ):

                    self.on_recorded(
                        item[1],
                        item[2]
                    )

                elif (
                    isinstance(item, tuple)
                    and item[0] == "update_available"
                ):

                    _, remote_version, url = item

                    self._pending_version = (
                        remote_version
                    )

                    if (
                        self.cfg.get("skip_version")
                        != remote_version
                    ):
                        self.show_update_bar(
                            remote_version,
                            url
                        )

                elif (
                    isinstance(item, tuple)
                    and item[0] == "manual_check_result"
                ):

                    result = item[1]

                    if result:

                        _, remote_version, url = result

                        self._pending_version = (
                            remote_version
                        )

                        self.show_update_bar(
                            remote_version,
                            url
                        )

                        self.version_lbl.configure(
                            text=(
                                f"Wersja: {VERSION} "
                                f"(nowa: {remote_version})"
                            )
                        )

                    else:

                        self.version_lbl.configure(
                            text=(
                                f"Wersja: {VERSION} "
                                "(to najnowsza wersja)"
                            )
                        )

                elif (
                    isinstance(item, tuple)
                    and item[0] == "update_downloaded"
                ):

                    _, current_file, temp_file = item

                    result = messagebox.askyesno(
                        "NetBlocker",
                        "Aktualizacja została pobrana.\n\n"
                        "Program zostanie zamknięty "
                        "i uruchomiony ponownie "
                        "w nowej wersji.\n\n"
                        "Kontynuować?"
                    )

                    if result:

                        self.install_update(
                            current_file,
                            temp_file
                        )

                    else:

                        try:
                            os.remove(
                                temp_file
                            )
                        except Exception:
                            pass

                        self.update_lbl.configure(
                            text=(
                                "Aktualizacja pobrana, "
                                "ale nie została zainstalowana."
                            )
                        )

                elif (
                    isinstance(item, tuple)
                    and item[0] == "update_error"
                ):

                    _, error = item

                    self.update_lbl.configure(
                        text=(
                            "Nie udało się pobrać aktualizacji."
                        )
                    )

                    messagebox.showerror(
                        "NetBlocker",
                        f"Nie udało się pobrać aktualizacji:\n{error}"
                    )

                elif item == "toggle":

                    now = time.time()

                    if (
                        now - self._last_toggle
                        > 0.4
                    ):

                        self._last_toggle = now

                        self.toggle(
                            silent=True
                        )

                elif item == "lag":

                    self.do_lag(
                        silent=True
                    )

        except queue.Empty:
            pass

        self.after(
            30,
            self.poll_queue
        )


    # -----------------------------------------------------------------
    # LISTA APLIKACJI
    # -----------------------------------------------------------------

    def refresh(self):

        self.apps = list_apps()

        labels = (
            list(self.apps.keys())
            or ["(brak)"]
        )

        self.combo.configure(
            values=labels
        )

        label = (
            self.current_label
            if self.current_label
            in self.apps
            else None
        )

        if (
            label is None
            and self.cfg.get("last_app")
        ):

            for l, p in self.apps.items():

                if (
                    p.lower()
                    == self.cfg["last_app"].lower()
                ):
                    label = l
                    break

        self.combo.set(
            label or labels[0]
        )

        self.update_status()


    def on_select(self, label):

        self.current_label = label

        path = self.apps.get(label)

        if path:

            self.cfg["last_app"] = path

            save_config(
                self.cfg
            )

        self.update_status()


    def selected_path(self):

        return self.apps.get(
            self.combo.get()
        )


    def toggle_text(self, blocked):

        base = (
            "Przywróć internet"
            if blocked
            else "Odetnij internet"
        )

        hk = self.cfg["hotkeys"].get(
            "toggle"
        )

        return (
            f"{base}   [{fmt_hotkey(hk)}]"
            if hk
            else base
        )


    def lag_text(self):

        hk = self.cfg["hotkeys"].get(
            "lag"
        )

        return (
            f"⚡ LAG   [{fmt_hotkey(hk)}]"
            if hk
            else "⚡ LAG"
        )


    def update_status(self):

        T = self.T
        path = self.selected_path()

        try:

            if not path:

                self.path_lbl.configure(
                    text=""
                )

                self.status.configure(
                    text="Wybierz aplikację"
                )

                self.dot.configure(
                    text_color=T["muted"]
                )

                self.toggle_btn.configure(
                    text=self.toggle_text(False),
                    fg_color=T["red"],
                    hover_color=T["red_h"]
                )

                return

            self.path_lbl.configure(
                text=path
            )

            if is_blocked(path):

                self.dot.configure(
                    text_color=T["red"]
                )

                self.status.configure(
                    text="Internet ODCIĘTY"
                )

                self.toggle_btn.configure(
                    text=self.toggle_text(True),
                    fg_color=T["green"],
                    hover_color=T["green_h"]
                )

            else:

                self.dot.configure(
                    text_color=T["green"]
                )

                self.status.configure(
                    text="Internet działa"
                )

                self.toggle_btn.configure(
                    text=self.toggle_text(False),
                    fg_color=T["red"],
                    hover_color=T["red_h"]
                )

        except Exception:
            pass


    def flash(self, msg):

        try:

            self.info.configure(
                text=msg
            )

            self.after(
                1800,
                lambda: self._clear_info(msg)
            )

        except Exception:
            pass


    def _clear_info(self, msg):

        try:

            if (
                self.info.cget("text")
                == msg
            ):
                self.info.configure(
                    text=""
                )

        except Exception:
            pass


    # -----------------------------------------------------------------
    # BLOKADA
    # -----------------------------------------------------------------

    def toggle(
        self,
        silent=False
    ):

        T = self.T
        path = self.selected_path()

        if not path:

            if silent:
                self.flash(
                    "Nie wybrano aplikacji"
                )
            else:
                messagebox.showinfo(
                    "NetBlocker",
                    "Najpierw wybierz aplikację."
                )

            return

        try:

            if is_blocked(path):

                remove_rules(
                    path,
                    "block"
                )

                self.flash(
                    "Internet przywrócony"
                )

                self.slide_status(
                    self.status,
                    "Internet działa",
                    T["green"]
                )

            else:

                ensure_rules(
                    path,
                    "block",
                    True
                )

                self.flash(
                    "Internet odcięty"
                )

                self.slide_status(
                    self.status,
                    "Internet ODCIĘTY",
                    T["red"]
                )

            self.pulse_button(
                self.toggle_btn,
                T["accent"],
                self.toggle_btn.cget("fg_color"),
                times=2,
                delay=90
            )

        except Exception as e:

            messagebox.showerror(
                "NetBlocker",
                "Nie udało się zmienić "
                f"reguły zapory:\n{e}"
            )

        self.update_status()


    # -----------------------------------------------------------------
    # LAG
    # -----------------------------------------------------------------

    def on_slider(self, v):

        self.cfg["lag_ms"] = int(v)

        self.ms_lbl.configure(
            text=f"{int(v)} ms"
        )

        save_config(
            self.cfg
        )


    def set_lag_btn(self, running):

        try:

            self.lag_btn.configure(
                text=(
                    "LAG…"
                    if running
                    else self.lag_text()
                ),
                state=(
                    "disabled"
                    if running
                    else "normal"
                )
            )

        except Exception:
            pass


    def do_lag(
        self,
        silent=False
    ):

        path = self.selected_path()

        if not path:

            if silent:
                self.flash(
                    "Nie wybrano aplikacji"
                )
            else:
                messagebox.showinfo(
                    "NetBlocker",
                    "Najpierw wybierz aplikację."
                )

            return

        if self.lag_running:
            return

        try:

            if path not in self.lag_paths:

                ensure_rules(
                    path,
                    "lag",
                    False
                )

                self.lag_paths.add(
                    path
                )

            self.lag_running = True

            self.set_lag_btn(
                True
            )

            self.pulse_button(
                self.lag_btn,
                self.T["red"],
                self.T["orange"],
                times=int(
                    max(
                        2,
                        self.cfg["lag_ms"] / 120
                    )
                ),
                delay=100
            )

            set_enabled(
                path,
                "lag",
                True
            )

            self.after(
                int(self.cfg["lag_ms"]),
                lambda: self.end_lag(path)
            )

        except Exception as e:

            self.lag_running = False

            self.set_lag_btn(
                False
            )

            messagebox.showerror(
                "NetBlocker",
                f"Nie udało się wywołać lagu:\n{e}"
            )


    def end_lag(self, path):

        set_enabled(
            path,
            "lag",
            False
        )

        self.lag_running = False

        self.set_lag_btn(
            False
        )

        self.flash(
            "Lag zakończony"
        )


    # -----------------------------------------------------------------
    # ZAMYKANIE
    # -----------------------------------------------------------------

    def on_close(self):

        try:
            keyboard.unhook_all()
        except Exception:
            pass

        for path in list(
            self.lag_paths
        ):
            remove_rules(
                path,
                "lag"
            )

        save_config(
            self.cfg
        )

        self.destroy()


# =====================================================================
# START
# =====================================================================

if __name__ == "__main__":

    if sys.platform != "win32":
        sys.exit(
            "Ten skrypt działa tylko na Windows."
        )

    if not is_admin():
        relaunch_as_admin()

    cleanup_stale_lag_rules()

    App().mainloop()