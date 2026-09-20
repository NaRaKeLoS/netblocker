from pathlib import Path

code = r'''"""
NetBlocker v1.3.0 - odcina internet wybranej aplikacji na Windows 11

Nowości:
- blokada na stałe
- LAG: krótkie odcięcie internetu
- motywy kolorystyczne
- własne skróty klawiszowe
- animacje UI
- wybór języka przy pierwszym uruchomieniu
- zmiana języka w ustawieniach
- automatyczne sprawdzanie aktualizacji
- automatyczne pobieranie aktualizacji przy starcie
- aktualizacja programu jako .exe
- wybór procesu z klawiatury: B -> pierwszy proces na B,
  kolejne B -> następny proces na B

Wymagania do uruchomienia jako .py:
    pip install customtkinter psutil pywin32 keyboard requests

Budowanie .exe:
    pyinstaller --noconfirm --onefile --windowed --uac-admin --name NetBlocker net_blocker.py
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
import tkinter as tk
import tkinter.messagebox as messagebox

import customtkinter as ctk
import keyboard
import psutil
import win32com.client


# ---------------------------------------------------------------------
# WERSJA I AKTUALIZACJE
# ---------------------------------------------------------------------

VERSION = "1.3.0"

VERSION_CHECK_URL = (
    "https://raw.githubusercontent.com/NaRaKeLoS/netblocker/"
    "refs/heads/main/version.txt"
)

DOWNLOAD_PAGE_URL = "https://github.com/NaRaKeLoS/netblocker"

UPDATE_EXE_NAME = "NetBlocker.exe"
UPDATE_MIN_SIZE = 100_000

PREFIX = "NetBlocker_"

DIR_IN, DIR_OUT = 1, 2
ACTION_BLOCK = 0
PROFILES_ALL = 0x7FFFFFFF


def parse_version(v):
    try:
        return tuple(int(x) for x in v.strip().split("."))
    except Exception:
        return (0,)


def release_download_url(version):
    """
    Domyślna lokalizacja pliku EXE w GitHub Releases.

    Release:
        v1.3.0

    Asset:
        NetBlocker.exe
    """
    return (
        f"https://github.com/NaRaKeLoS/netblocker/"
        f"releases/download/v{version}/{UPDATE_EXE_NAME}"
    )


def check_for_update(result_queue):
    """
    version.txt:
        linia 1 -> numer wersji
        linia 2 -> opcjonalny bezpośredni URL do EXE

    Jeżeli druga linia nie istnieje, program buduje URL:
        Releases/download/vX.Y.Z/NetBlocker.exe
    """

    try:
        import random
        import requests

        bust = f"?_={int(time.time())}{random.randint(1000, 9999)}"

        r = requests.get(
            VERSION_CHECK_URL + bust,
            timeout=7,
            headers={
                "Cache-Control": "no-cache",
                "Pragma": "no-cache",
            },
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
            download_url = release_download_url(remote_version)

        if parse_version(remote_version) > parse_version(VERSION):
            result_queue.put(
                (
                    "update_available",
                    remote_version,
                    download_url,
                )
            )

    except Exception:
        pass


# ---------------------------------------------------------------------
# JĘZYKI
# ---------------------------------------------------------------------

LANGUAGES = {
    "pl": "Polski",
    "en": "English",
}


TEXT = {
    "pl": {
        "app_subtitle": "odcinaj internet wybranym aplikacjom",
        "settings": "Ustawienia",
        "back": "← Wróć",
        "application": "APLIKACJA",
        "none": "(brak)",
        "refresh": "Odśwież",
        "blocked": "Internet ODCIĘTY",
        "working": "Internet działa",
        "choose_app": "Wybierz aplikację",
        "cut_internet": "Odetnij internet",
        "restore_internet": "Przywróć internet",
        "lag": "⚡ LAG",
        "lag_running": "LAG…",
        "lag_time": "Czas odcięcia:",
        "permanent_note": "Blokada na stałe zostaje w zaporze po zamknięciu programu.",
        "block_section": "BLOKADA",
        "lag_section": "LAG",
        "theme": "MOTYW",
        "hotkeys": "SKRÓTY KLAWISZOWE (GLOBALNE)",
        "other": "INNE",
        "about": "O PROGRAMIE",
        "always_top": "Okno zawsze na wierzchu",
        "auto_update": "Automatycznie pobieraj aktualizacje przy starcie",
        "language": "JĘZYK",
        "language_desc": "Możesz zmienić język programu w dowolnym momencie.",
        "check_updates": "Sprawdź aktualizacje",
        "version": "Wersja",
        "latest": "(to najnowsza wersja)",
        "checking": "(sprawdzam...)",
        "new_version": "(nowa: {version})",
        "available_update": "Dostępna nowa wersja {version} (masz {current})",
        "download": "Pobierz",
        "install": "Zainstaluj",
        "skip": "Pomiń",
        "downloading": "Pobieranie aktualizacji…",
        "downloaded": "Aktualizacja pobrana. Zainstalować teraz?",
        "downloaded_ready": "Aktualizacja pobrana. Kliknij „Zainstaluj”, aby ją wgrać.",
        "download_error_bar": "Nie udało się pobrać aktualizacji.",
        "download_failed": "Nie udało się pobrać aktualizacji:\n{error}",
        "invalid_update": "Pobrany plik nie wygląda na poprawny plik EXE.",
        "update_not_exe": "Automatyczna aktualizacja wymaga wersji EXE.",
        "no_update_url": "Nie znaleziono adresu aktualizacji.",
        "install_error": "Nie udało się przygotować aktualizacji:\n{error}",
        "install_now": "Aktualizacja jest gotowa.\n\nProgram zostanie zamknięty i uruchomiony ponownie w nowej wersji.\n\nKontynuować?",
        "record_hotkey": "Naciśnij skrót…",
        "no_hotkey": "— brak —",
        "hotkey_help": "Kliknij pole skrótu i naciśnij nową kombinację klawiszy.\nEsc anuluje. Skróty działają na aplikację wybraną na liście.",
        "toggle_hotkey": "Odetnij / przywróć",
        "lag_hotkey": "Lag",
        "hotkey_used": "Skrót {hotkey} jest już użyty w innej akcji.",
        "no_app": "Najpierw wybierz aplikację.",
        "no_app_short": "Nie wybrano aplikacji",
        "internet_restored": "Internet przywrócony",
        "internet_cut": "Internet odcięty",
        "lag_finished": "Lag zakończony",
        "lag_error": "Nie udało się wywołać lagu:\n{error}",
        "firewall_error": "Nie udało się zmienić reguły zapory:\n{error}",
        "language_title": "Wybierz język",
        "language_subtitle": "Wybierz język programu.",
        "continue": "Dalej",
        "language_changed": "Język zmieniony",
        "process_found": "Wybrano: {name}",
        "no_process_letter": "Brak procesu zaczynającego się na „{letter}”.",
        "theme_dark": "Ciemny",
        "theme_light": "Jasny",
        "theme_midnight": "Północ",
        "theme_ocean": "Ocean",
        "theme_sunset": "Zachód słońca",
        "theme_hacker": "Hacker",
    },
    "en": {
        "app_subtitle": "cut internet access for selected applications",
        "settings": "Settings",
        "back": "← Back",
        "application": "APPLICATION",
        "none": "(none)",
        "refresh": "Refresh",
        "blocked": "Internet BLOCKED",
        "working": "Internet works",
        "choose_app": "Choose an application",
        "cut_internet": "Block internet",
        "restore_internet": "Restore internet",
        "lag": "⚡ LAG",
        "lag_running": "LAG…",
        "lag_time": "Disconnect time:",
        "permanent_note": "Permanent blocking remains in the firewall after closing the program.",
        "block_section": "BLOCKING",
        "lag_section": "LAG",
        "theme": "THEME",
        "hotkeys": "GLOBAL HOTKEYS",
        "other": "OTHER",
        "about": "ABOUT",
        "always_top": "Always on top",
        "auto_update": "Automatically download updates at startup",
        "language": "LANGUAGE",
        "language_desc": "You can change the program language at any time.",
        "check_updates": "Check for updates",
        "version": "Version",
        "latest": "(up to date)",
        "checking": "(checking...)",
        "new_version": "(new: {version})",
        "available_update": "New version {version} available (you have {current})",
        "download": "Download",
        "install": "Install",
        "skip": "Skip",
        "downloading": "Downloading update…",
        "downloaded": "Update downloaded. Install it now?",
        "downloaded_ready": "Update downloaded. Click “Install” to install it.",
        "download_error_bar": "Failed to download update.",
        "download_failed": "Failed to download update:\n{error}",
        "invalid_update": "The downloaded file does not look like a valid EXE.",
        "update_not_exe": "Automatic updates require the EXE version.",
        "no_update_url": "No update URL was found.",
        "install_error": "Failed to prepare the update:\n{error}",
        "install_now": "The update is ready.\n\nThe program will close and restart in the new version.\n\nContinue?",
        "record_hotkey": "Press a shortcut…",
        "no_hotkey": "— none —",
        "hotkey_help": "Click the shortcut field and press a new key combination.\nEsc cancels. Shortcuts work on the selected application.",
        "toggle_hotkey": "Block / restore",
        "lag_hotkey": "Lag",
        "hotkey_used": "Shortcut {hotkey} is already used by another action.",
        "no_app": "Choose an application first.",
        "no_app_short": "No application selected",
        "internet_restored": "Internet restored",
        "internet_cut": "Internet blocked",
        "lag_finished": "Lag finished",
        "lag_error": "Failed to trigger lag:\n{error}",
        "firewall_error": "Failed to change firewall rule:\n{error}",
        "language_title": "Choose language",
        "language_subtitle": "Choose the program language.",
        "continue": "Continue",
        "language_changed": "Language changed",
        "process_found": "Selected: {name}",
        "no_process_letter": "No process starting with “{letter}”.",
        "theme_dark": "Dark",
        "theme_light": "Light",
        "theme_midnight": "Midnight",
        "theme_ocean": "Ocean",
        "theme_sunset": "Sunset",
        "theme_hacker": "Hacker",
    },
}


def t(lang, key, **kwargs):
    value = TEXT.get(lang, TEXT["pl"]).get(key, key)
    try:
        return value.format(**kwargs)
    except Exception:
        return value


# ---------------------------------------------------------------------
# KONFIGURACJA
# ---------------------------------------------------------------------

CONFIG_DIR = os.path.join(
    os.environ.get("APPDATA", "."),
    "NetBlocker",
)

CONFIG_PATH = os.path.join(
    CONFIG_DIR,
    "config.json",
)

DEFAULT_CFG = {
    "language": None,
    "theme": "Ciemny",
    "hotkeys": {
        "toggle": "ctrl+alt+b",
        "lag": "ctrl+alt+l",
    },
    "lag_ms": 500,
    "topmost": False,
    "auto_update": True,
    "last_app": None,
    "skip_version": None,
}


THEME_KEYS = {
    "Ciemny": "theme_dark",
    "Jasny": "theme_light",
    "Północ": "theme_midnight",
    "Ocean": "theme_ocean",
    "Zachód słońca": "theme_sunset",
    "Hacker": "theme_hacker",
}


def load_config():
    cfg = copy.deepcopy(DEFAULT_CFG)
    file_missing_or_broken = True

    try:
        with open(CONFIG_PATH, encoding="utf-8") as f:
            data = json.load(f)

        for k in (
            "language",
            "theme",
            "lag_ms",
            "topmost",
            "auto_update",
            "last_app",
            "skip_version",
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
            encoding="utf-8",
        ) as f:
            json.dump(
                cfg,
                f,
                indent=2,
                ensure_ascii=False,
            )

        os.replace(
            tmp_path,
            CONFIG_PATH,
        )
        return True

    except Exception:
        return False


def fmt_hotkey(hk):
    return hk.upper().replace("+", " + ") if hk else ""


# ---------------------------------------------------------------------
# MOTYWY
# ---------------------------------------------------------------------

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
    orange_h="#bd7418",
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
        orange_h=orange_h,
    )


THEMES = {
    "Ciemny": _theme(
        "dark",
        "#15171c",
        "#1f2229",
        "#e8eaed",
        "#8b919c",
        "#3b82f6",
        "#2f6bd0",
    ),
    "Jasny": _theme(
        "light",
        "#eef0f4",
        "#ffffff",
        "#1c1f26",
        "#6b7280",
        "#2563eb",
        "#1d4fbf",
    ),
    "Północ": _theme(
        "dark",
        "#120f1f",
        "#1c1830",
        "#ece9f7",
        "#8f88ad",
        "#8b5cf6",
        "#7444d8",
    ),
    "Ocean": _theme(
        "dark",
        "#0b1a24",
        "#12283a",
        "#e2f1f8",
        "#7d9db0",
        "#06b6d4",
        "#0592ab",
    ),
    "Zachód słońca": _theme(
        "dark",
        "#1f1414",
        "#2c1c1c",
        "#f6e9e4",
        "#a58880",
        "#f97316",
        "#d5600e",
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
        orange_h="#d69600",
    ),
}


# ---------------------------------------------------------------------
# ADMIN
# ---------------------------------------------------------------------

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
        for a in sys.argv[1:]
    )

    executable = sys.executable

    ctypes.windll.shell32.ShellExecuteW(
        None,
        "runas",
        executable,
        params,
        None,
        1,
    )

    sys.exit(0)


# ---------------------------------------------------------------------
# WINDOWS FIREWALL
# ---------------------------------------------------------------------

policy = win32com.client.Dispatch(
    "HNetCfg.FwPolicy2"
)


def rule_names(path: str, kind: str):
    p = path.lower()

    return [
        f"{PREFIX}{kind}_out_{p}",
        f"{PREFIX}{kind}_in_{p}",
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
    enabled,
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
    enabled,
):
    for name, direction in zip(
        rule_names(path, kind),
        (DIR_OUT, DIR_IN),
    ):
        if rule_exists(name):
            policy.Rules.Item(name).Enabled = enabled
        else:
            add_rule(
                name,
                path,
                direction,
                enabled,
            )


def set_enabled(
    path,
    kind,
    enabled,
):
    for name in rule_names(path, kind):
        try:
            policy.Rules.Item(name).Enabled = enabled
        except Exception:
            pass


def remove_rules(
    path,
    kind,
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


# ---------------------------------------------------------------------
# PROCESY
# ---------------------------------------------------------------------

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
            or exe.lower().startswith(r"c:\windows")
        ):
            continue

        found[exe.lower()] = (
            name,
            exe,
        )

    names = [
        n
        for n, _ in found.values()
    ]

    result = {}

    for name, exe in sorted(
        found.values(),
        key=lambda x: x[0].lower(),
    ):
        label = name

        if names.count(name) > 1:
            label = (
                f"{name}  "
                f"({os.path.basename(os.path.dirname(exe))})"
            )

        result[label] = exe

    return result


# ---------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------

class App(ctk.CTk):

    def __init__(self):
        super().__init__()

        self.cfg = load_config()

        self.T = THEMES.get(
            self.cfg["theme"],
            THEMES["Ciemny"],
        )

        self.lang = self.cfg.get("language") or "pl"

        self.apps = {}
        self.current_label = None

        self.lag_paths = set()
        self.lag_running = False

        self.recording = None
        self.hk_btns = {}

        self._last_toggle = 0.0
        self._pulse_jobs = set()

        self.q = queue.Queue()

        self._page = "main"
        self._update_url = None
        self._pending_version = None
        self._update_temp = None
        self._update_downloading = False

        self._key_cycle_letter = None
        self._key_cycle_index = -1
        self._key_cycle_time = 0.0

        self.title(
            f"NetBlocker v{VERSION}"
        )

        self.geometry(
            "460x700"
        )

        self.minsize(
            460,
            700,
        )

        self.maxsize(
            460,
            700,
        )

        self.attributes(
            "-topmost",
            bool(self.cfg["topmost"]),
        )

        self.attributes(
            "-alpha",
            0.0,
        )

        if not self.cfg.get("language"):
            self.choose_language()

        self.build_ui("main")
        self.register_hotkeys()

        self.protocol(
            "WM_DELETE_WINDOW",
            self.on_close,
        )

        self.after(
            30,
            self.poll_queue,
        )

        self.fade_in()

        threading.Thread(
            target=check_for_update,
            args=(self.q,),
            daemon=True,
        ).start()

        self.bind_all(
            "<KeyPress>",
            self.on_global_keypress,
            add="+",
        )


    # -----------------------------------------------------------------
    # JĘZYK
    # -----------------------------------------------------------------

    def tr(self, key, **kwargs):
        return t(
            self.lang,
            key,
            **kwargs,
        )


    def choose_language(self):
        dialog = ctk.CTkToplevel(self)

        dialog.title(
            self.tr("language_title")
        )

        dialog.geometry(
            "360x280"
        )

        dialog.resizable(
            False,
            False,
        )

        dialog.transient(self)
        dialog.grab_set()

        self.update_idletasks()

        x = self.winfo_x() + (
            self.winfo_width() - 360
        ) // 2
        y = self.winfo_y() + (
            self.winfo_height() - 280
        ) // 2

        if x < 0:
            x = 100
        if y < 0:
            y = 100

        dialog.geometry(
            f"360x280+{x}+{y}"
        )

        title = self.lbl(
            dialog,
            "NetBlocker",
            28,
            True,
        )
        title.pack(
            pady=(28, 2)
        )

        self.lbl(
            dialog,
            self.tr("language_subtitle"),
            13,
            False,
            self.T["muted"],
        ).pack(
            pady=(0, 14)
        )

        selected = tk.StringVar(
            value="pl"
        )

        frame = ctk.CTkFrame(
            dialog,
            fg_color="transparent",
        )
        frame.pack(
            fill="x",
            padx=34,
        )

        for code, name in LANGUAGES.items():
            ctk.CTkRadioButton(
                frame,
                text=name,
                value=code,
                variable=selected,
                font=ctk.CTkFont(
                    size=14,
                    weight="bold",
                ),
            ).pack(
                anchor="w",
                pady=5,
            )

        def apply():
            self.lang = selected.get()
            self.cfg["language"] = self.lang
            save_config(self.cfg)

            dialog.grab_release()
            dialog.destroy()

        ctk.CTkButton(
            dialog,
            text=self.tr("continue"),
            width=180,
            height=40,
            fg_color=self.T["accent"],
            hover_color=self.T["accent_h"],
            command=apply,
        ).pack(
            pady=(20, 0)
        )

        self.wait_window(dialog)


    # -----------------------------------------------------------------
    # ANIMACJE
    # -----------------------------------------------------------------

    def fade_in(self, step=0.0):
        step = min(
            step + 0.08,
            1.0,
        )

        try:
            self.attributes(
                "-alpha",
                step,
            )
        except Exception:
            return

        if step < 1.0:
            self.after(
                15,
                lambda: self.fade_in(step),
            )


    def fade_to(self, target, step=1.0, callback=None):
        try:
            self.attributes(
                "-alpha",
                max(0.0, min(1.0, step)),
            )
        except Exception:
            if callback:
                callback()
            return

        if abs(step - target) < 0.02:
            try:
                self.attributes(
                    "-alpha",
                    target,
                )
            except Exception:
                pass

            if callback:
                callback()
            return

        direction = (
            -0.08
            if step > target
            else 0.08
        )

        self.after(
            15,
            lambda: self.fade_to(
                target,
                step + direction,
                callback,
            ),
        )


    def pulse_button(
        self,
        btn,
        color_a,
        color_b,
        times=6,
        delay=110,
    ):
        job_id = object()
        self._pulse_jobs.add(job_id)

        def step(i):
            if job_id not in self._pulse_jobs:
                return

            try:
                if not btn.winfo_exists():
                    self._pulse_jobs.discard(job_id)
                    return

                btn.configure(
                    fg_color=(
                        color_a
                        if i % 2 == 0
                        else color_b
                    )
                )

            except Exception:
                self._pulse_jobs.discard(job_id)
                return

            if i < times:
                self.after(
                    delay,
                    lambda: step(i + 1),
                )
            else:
                self._pulse_jobs.discard(job_id)

        step(0)


    def animate_status(
        self,
        label,
        text,
        color,
    ):
        try:
            label.configure(
                text=text,
                text_color=color,
                font=ctk.CTkFont(
                    size=16,
                    weight="bold",
                ),
            )

            self.after(
                150,
                lambda: label.configure(
                    font=ctk.CTkFont(
                        size=14,
                        weight="bold",
                    )
                ),
            )

        except Exception:
            pass


    def show_page(self, page):
        if page == self._page:
            return

        self.fade_to(
            0.72,
            step=float(
                self.attributes("-alpha")
            ),
            callback=lambda: self._finish_page_change(page),
        )


    def _finish_page_change(self, page):
        self.show(page)
        self.fade_in_from(
            float(
                self.attributes("-alpha")
            )
        )


    def fade_in_from(self, step):
        step = min(
            step + 0.07,
            1.0,
        )

        try:
            self.attributes(
                "-alpha",
                step,
            )
        except Exception:
            return

        if step < 1.0:
            self.after(
                15,
                lambda: self.fade_in_from(step),
            )


    # -----------------------------------------------------------------
    # UI HELPERS
    # -----------------------------------------------------------------

    def lbl(
        self,
        parent,
        text,
        size=13,
        bold=False,
        color=None,
        **kw,
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
                ),
            ),
            **kw,
        )


    def card(
        self,
        parent,
        title,
    ):
        c = ctk.CTkFrame(
            parent,
            corner_radius=14,
            fg_color=self.T["card"],
        )

        c.pack(
            fill="x",
            padx=22,
            pady=6,
        )

        self.lbl(
            c,
            title,
            11,
            True,
            self.T["muted"],
        ).pack(
            anchor="w",
            padx=16,
            pady=(12, 2),
        )

        return c


    def build_ui(self, page="main"):

        if hasattr(self, "combo"):
            try:
                self.current_label = self.combo.get()
            except Exception:
                pass

        for w in self.winfo_children():
            if isinstance(w, ctk.CTkToplevel):
                continue
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
            fg_color="transparent",
        )

        self.settings_frame = ctk.CTkFrame(
            self,
            fg_color="transparent",
        )

        self.build_main()
        self.build_settings()

        self._page = page

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
            expand=True,
        )

        self._page = page


    # -----------------------------------------------------------------
    # MAIN
    # -----------------------------------------------------------------

    def build_main(self):

        T = self.T
        f = self.main_frame

        header = ctk.CTkFrame(
            f,
            fg_color="transparent",
        )

        header.pack(
            fill="x",
            padx=22,
            pady=(20, 4),
        )

        titles = ctk.CTkFrame(
            header,
            fg_color="transparent",
        )

        titles.pack(side="left")

        self.lbl(
            titles,
            "NetBlocker",
            26,
            True,
        ).pack(anchor="w")

        self.lbl(
            titles,
            f"v{VERSION} — {self.tr('app_subtitle')}",
            12,
            False,
            T["muted"],
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
            command=lambda: self.show_page("settings"),
        ).pack(side="right")

        self._header_widget = header

        # UPDATE BAR

        self.update_bar = ctk.CTkFrame(
            f,
            corner_radius=10,
            fg_color=T["accent"],
        )

        self.update_lbl = self.lbl(
            self.update_bar,
            "",
            11,
            True,
            "#ffffff",
            wraplength=215,
            justify="left",
        )

        self.update_lbl.pack(
            side="left",
            fill="both",
            expand=True,
            padx=(12, 6),
            pady=8,
        )

        self.update_skip_btn = ctk.CTkButton(
            self.update_bar,
            text=self.tr("skip"),
            width=62,
            height=30,
            fg_color="transparent",
            text_color="#ffffff",
            hover_color=T["accent_h"],
            command=self.dismiss_update,
        )

        self.update_skip_btn.pack(
            side="right",
            padx=(0, 4),
            pady=5,
        )

        self.update_download_btn = ctk.CTkButton(
            self.update_bar,
            text=self.tr("download"),
            width=78,
            height=30,
            fg_color="#ffffff",
            text_color=T["accent"],
            hover_color="#e5e5e5",
            command=self.open_download,
        )

        self.update_download_btn.pack(
            side="right",
            padx=(4, 4),
            pady=5,
        )

        # APLIKACJA

        c1 = self.card(
            f,
            self.tr("application"),
        )

        row = ctk.CTkFrame(
            c1,
            fg_color="transparent",
        )

        row.pack(
            fill="x",
            padx=12,
            pady=(0, 4),
        )

        self.combo = ctk.CTkComboBox(
            row,
            values=[self.tr("none")],
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
            dropdown_hover_color=T["accent"],
        )

        self.combo.pack(
            side="left",
            fill="x",
            expand=True,
            padx=(4, 6),
        )

        ctk.CTkButton(
            row,
            text="⟳",
            width=40,
            height=36,
            fg_color=T["accent"],
            hover_color=T["accent_h"],
            text_color="#ffffff",
            command=self.refresh,
        ).pack(
            side="left",
            padx=(0, 4),
        )

        self.path_lbl = self.lbl(
            c1,
            "",
            11,
            False,
            T["muted"],
            wraplength=390,
            justify="left",
        )

        self.path_lbl.pack(
            anchor="w",
            padx=16,
            pady=(0, 12),
        )

        # BLOKADA

        c2 = self.card(
            f,
            self.tr("block_section"),
        )

        srow = ctk.CTkFrame(
            c2,
            fg_color="transparent",
        )

        srow.pack(
            fill="x",
            padx=16,
            pady=(0, 8),
        )

        self.dot = self.lbl(
            srow,
            "●",
            18,
            False,
            T["muted"],
        )
        self.dot.pack(side="left")

        self.status = self.lbl(
            srow,
            self.tr("choose_app"),
            14,
            True,
        )
        self.status.pack(
            side="left",
            padx=8,
        )

        self.toggle_btn = ctk.CTkButton(
            c2,
            text=self.tr("cut_internet"),
            height=46,
            font=ctk.CTkFont(
                size=15,
                weight="bold",
            ),
            fg_color=T["red"],
            hover_color=T["red_h"],
            text_color="#ffffff",
            command=lambda: self.toggle(),
        )

        self.toggle_btn.pack(
            fill="x",
            padx=16,
            pady=(0, 16),
        )

        # LAG

        c3 = self.card(
            f,
            self.tr("lag_section"),
        )

        lrow = ctk.CTkFrame(
            c3,
            fg_color="transparent",
        )

        lrow.pack(
            fill="x",
            padx=16,
        )

        self.lbl(
            lrow,
            self.tr("lag_time"),
        ).pack(side="left")

        self.ms_lbl = self.lbl(
            lrow,
            f"{int(self.cfg['lag_ms'])} ms",
            13,
            True,
        )

        self.ms_lbl.pack(
            side="right",
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
            command=self.on_slider,
        )

        self.slider.set(
            self.cfg["lag_ms"]
        )

        self.slider.pack(
            fill="x",
            padx=16,
            pady=(6, 10),
        )

        self.lag_btn = ctk.CTkButton(
            c3,
            text=self.lag_text(),
            height=46,
            font=ctk.CTkFont(
                size=15,
                weight="bold",
            ),
            fg_color=T["orange"],
            hover_color=T["orange_h"],
            text_color="#ffffff",
            command=lambda: self.do_lag(),
        )

        self.lag_btn.pack(
            fill="x",
            padx=16,
            pady=(0, 16),
        )

        self.info = self.lbl(
            f,
            "",
            12,
            False,
            T["muted"],
        )

        self.info.pack(
            pady=(6, 0),
        )

        self.lbl(
            f,
            self.tr("permanent_note"),
            10,
            False,
            T["muted"],
        ).pack(
            pady=(2, 0),
        )


    # -----------------------------------------------------------------
    # SETTINGS
    # -----------------------------------------------------------------

    def build_settings(self):

        T = self.T
        f = self.settings_frame

        header = ctk.CTkFrame(
            f,
            fg_color="transparent",
        )

        header.pack(
            fill="x",
            padx=22,
            pady=(20, 4),
        )

        ctk.CTkButton(
            header,
            text=self.tr("back"),
            width=80,
            height=36,
            fg_color=T["card"],
            hover_color=T["accent"],
            text_color=T["text"],
            command=lambda: self.show_page("main"),
        ).pack(side="left")

        self.lbl(
            header,
            self.tr("settings"),
            22,
            True,
        ).pack(
            side="left",
            padx=14,
        )

        # MOTYW

        c1 = self.card(
            f,
            self.tr("theme"),
        )

        grid = ctk.CTkFrame(
            c1,
            fg_color="transparent",
        )

        grid.pack(
            fill="x",
            padx=12,
            pady=(0, 12),
        )

        grid.columnconfigure(
            (0, 1),
            weight=1,
        )

        for i, (name, theme) in enumerate(
            THEMES.items()
        ):
            selected = (
                name == self.cfg["theme"]
            )

            ctk.CTkButton(
                grid,
                text=(
                    ("✓ " if selected else "")
                    + self.tr(
                        THEME_KEYS[name]
                    )
                ),
                height=38,
                fg_color=theme["bg"],
                hover_color=theme["card"],
                text_color=theme["text"],
                border_width=3 if selected else 1,
                border_color=theme["accent"],
                command=lambda n=name: self.set_theme(n),
            ).grid(
                row=i // 2,
                column=i % 2,
                padx=4,
                pady=4,
                sticky="ew",
            )

        # JĘZYK

        c_lang = self.card(
            f,
            self.tr("language"),
        )

        self.language_combo = ctk.CTkComboBox(
            c_lang,
            values=list(LANGUAGES.values()),
            state="readonly",
            height=36,
            command=self.on_language_combo,
            fg_color=T["bg"],
            border_color=T["bg"],
            text_color=T["text"],
            button_color=T["accent"],
            button_hover_color=T["accent_h"],
            dropdown_fg_color=T["card"],
            dropdown_text_color=T["text"],
            dropdown_hover_color=T["accent"],
        )

        self.language_combo.set(
            LANGUAGES.get(
                self.lang,
                "Polski",
            )
        )

        self.language_combo.pack(
            fill="x",
            padx=16,
            pady=(0, 6),
        )

        self.lbl(
            c_lang,
            self.tr("language_desc"),
            11,
            False,
            T["muted"],
            justify="left",
        ).pack(
            anchor="w",
            padx=16,
            pady=(0, 14),
        )

        # HOTKEYE

        c2 = self.card(
            f,
            self.tr("hotkeys"),
        )

        hotkey_labels = {
            "toggle": self.tr("toggle_hotkey"),
            "lag": self.tr("lag_hotkey"),
        }

        for action, label in hotkey_labels.items():
            row = ctk.CTkFrame(
                c2,
                fg_color="transparent",
            )

            row.pack(
                fill="x",
                padx=16,
                pady=4,
            )

            self.lbl(
                row,
                label,
                13,
            ).pack(side="left")

            ctk.CTkButton(
                row,
                text="✕",
                width=34,
                height=34,
                fg_color=T["bg"],
                hover_color=T["red"],
                text_color=T["text"],
                command=lambda a=action: self.clear_hotkey(a),
            ).pack(
                side="right",
            )

            btn = ctk.CTkButton(
                row,
                text="",
                width=190,
                height=34,
                fg_color=T["bg"],
                hover_color=T["accent"],
                text_color=T["text"],
                command=lambda a=action: self.start_record(a),
            )

            btn.pack(
                side="right",
                padx=(0, 6),
            )

            self.hk_btns[action] = btn

        self.lbl(
            c2,
            self.tr("hotkey_help"),
            11,
            False,
            T["muted"],
            justify="left",
        ).pack(
            anchor="w",
            padx=16,
            pady=(6, 14),
        )

        self.refresh_hk_buttons()

        # INNE

        c3 = self.card(
            f,
            self.tr("other"),
        )

        sw = ctk.CTkSwitch(
            c3,
            text=self.tr("always_top"),
            text_color=T["text"],
            progress_color=T["accent"],
            command=self.on_topmost,
        )

        if self.cfg["topmost"]:
            sw.select()

        self.topmost_switch = sw

        sw.pack(
            anchor="w",
            padx=16,
            pady=(4, 8),
        )

        update_sw = ctk.CTkSwitch(
            c3,
            text=self.tr("auto_update"),
            text_color=T["text"],
            progress_color=T["accent"],
            command=self.on_auto_update,
        )

        if self.cfg.get("auto_update", True):
            update_sw.select()

        self.auto_update_switch = update_sw

        update_sw.pack(
            anchor="w",
            padx=16,
            pady=(0, 12),
        )

        # O PROGRAMIE

        c4 = self.card(
            f,
            self.tr("about"),
        )

        self.version_lbl = self.lbl(
            c4,
            f"{self.tr('version')}: {VERSION}",
            13,
        )

        self.version_lbl.pack(
            anchor="w",
            padx=16,
            pady=(0, 4),
        )

        ctk.CTkButton(
            c4,
            text=self.tr("check_updates"),
            height=36,
            fg_color=T["accent"],
            hover_color=T["accent_h"],
            text_color="#ffffff",
            command=self.manual_check_update,
        ).pack(
            fill="x",
            padx=16,
            pady=(4, 16),
        )


    def set_theme(self, name):
        self.cfg["theme"] = name
        self.T = THEMES[name]

        save_config(self.cfg)

        old_page = self._page

        self.fade_to(
            0.72,
            step=float(
                self.attributes("-alpha")
            ),
            callback=lambda: (
                self.build_ui(old_page),
                self.fade_in_from(0.72),
            ),
        )


    def on_language_combo(self, display_name):
        code = next(
            (
                c
                for c, n in LANGUAGES.items()
                if n == display_name
            ),
            self.lang,
        )

        if code == self.lang:
            return

        self.lang = code
        self.cfg["language"] = code
        save_config(self.cfg)

        old_page = self._page

        self.fade_to(
            0.72,
            step=float(
                self.attributes("-alpha")
            ),
            callback=lambda: (
                self.build_ui(old_page),
                self.fade_in_from(0.72),
            ),
        )


    def on_topmost(self):
        self.cfg["topmost"] = bool(
            self.topmost_switch.get()
        )

        self.attributes(
            "-topmost",
            self.cfg["topmost"],
        )

        save_config(self.cfg)


    def on_auto_update(self):
        self.cfg["auto_update"] = bool(
            self.auto_update_switch.get()
        )
        save_config(self.cfg)


    # -----------------------------------------------------------------
    # AKTUALIZACJE
    # -----------------------------------------------------------------

    def manual_check_update(self):
        self.version_lbl.configure(
            text=(
                f"{self.tr('version')}: {VERSION} "
                f"{self.tr('checking')}"
            )
        )

        threading.Thread(
            target=self._manual_check_worker,
            daemon=True,
        ).start()


    def _manual_check_worker(self):
        local_q = queue.Queue()

        check_for_update(local_q)

        try:
            item = local_q.get_nowait()
        except queue.Empty:
            item = None

        self.q.put(
            (
                "manual_check_result",
                item,
            )
        )


    def show_update_bar(
        self,
        remote_version,
        url,
    ):
        self._update_url = url
        self._pending_version = remote_version

        self.update_lbl.configure(
            text=self.tr(
                "available_update",
                version=remote_version,
                current=VERSION,
            )
        )

        self.update_download_btn.configure(
            text=(
                self.tr("install")
                if self._update_temp
                else self.tr("download")
            ),
            state="normal",
        )

        self.update_skip_btn.configure(
            state="normal",
            text=self.tr("skip"),
        )

        if not self.update_bar.winfo_ismapped():
            self.update_bar.pack(
                fill="x",
                padx=22,
                pady=(0, 6),
                after=self._header_widget,
            )

            self.after(
                30,
                lambda: self.pulse_button(
                    self.update_download_btn,
                    "#ffffff",
                    "#dce8ff",
                    times=2,
                    delay=90,
                ),
            )


    def open_download(self):
        if self._update_temp and os.path.exists(
            self._update_temp
        ):
            self.install_update(
                self._update_temp
            )
            return

        url = self._update_url

        if not url:
            messagebox.showerror(
                "NetBlocker",
                self.tr("no_update_url"),
            )
            return

        current_file = os.path.abspath(
            sys.executable
        )

        if not current_file.lower().endswith(
            ".exe"
        ):
            messagebox.showerror(
                "NetBlocker",
                self.tr("update_not_exe"),
            )
            return

        self.begin_update_download(
            url,
            current_file,
            automatic=False,
        )


    def begin_update_download(
        self,
        url,
        current_file,
        automatic=False,
    ):
        if self._update_downloading:
            return

        self._update_downloading = True

        self.update_lbl.configure(
            text=self.tr("downloading")
        )

        self.update_download_btn.configure(
            text="…",
            state="disabled",
        )

        self.update_skip_btn.configure(
            state="disabled",
        )

        threading.Thread(
            target=self._download_update_worker,
            args=(
                url,
                current_file,
                automatic,
            ),
            daemon=True,
        ).start()


    def _download_update_worker(
        self,
        url,
        current_file,
        automatic,
    ):
        try:
            import requests

            r = requests.get(
                url,
                timeout=60,
                headers={
                    "Cache-Control": "no-cache",
                    "Pragma": "no-cache",
                },
            )

            r.raise_for_status()

            new_exe = r.content

            if len(new_exe) < UPDATE_MIN_SIZE:
                raise RuntimeError(
                    self.tr("invalid_update")
                )

            if not new_exe.startswith(
                b"MZ"
            ):
                raise RuntimeError(
                    self.tr("invalid_update")
                )

            temp_file = (
                current_file
                + ".update"
            )

            with open(
                temp_file,
                "wb",
            ) as f:
                f.write(new_exe)

            self.q.put(
                (
                    "update_downloaded",
                    current_file,
                    temp_file,
                    automatic,
                )
            )

        except Exception as e:
            self.q.put(
                (
                    "update_error",
                    str(e),
                )
            )


    def create_updater_script(
        self,
        current_file,
        temp_file,
    ):
        current_file = os.path.abspath(
            current_file
        )
        temp_file = os.path.abspath(
            temp_file
        )

        pid = os.getpid()

        bat_path = os.path.join(
            os.path.dirname(current_file),
            "NetBlockerUpdater.cmd",
        )

        def bat_quote(value):
            return value.replace("%", "%%")

        old_q = bat_quote(current_file)
        new_q = bat_quote(temp_file)

        script = f"""@echo off
setlocal
set "OLD={old_q}"
set "NEW={new_q}"
set "PID={pid}"

:WAIT
tasklist /FI "PID eq %PID%" 2>NUL | findstr /R /C:" %PID% " >NUL
if not errorlevel 1 (
    timeout /t 1 /nobreak >NUL
    goto WAIT
)

set /a TRY=0

:REPLACE
set /a TRY+=1
if exist "%OLD%" (
    move /Y "%NEW%" "%OLD%" >NUL 2>&1
) else (
    move /Y "%NEW%" "%OLD%" >NUL 2>&1
)

if exist "%NEW%" if %TRY% LSS 15 (
    timeout /t 1 /nobreak >NUL
    goto REPLACE
)

if exist "%NEW%" (
    exit /b 1
)

start "" "%OLD%"

del "%~f0" >NUL 2>&1
exit /b 0
"""

        with open(
            bat_path,
            "w",
            encoding="utf-8",
        ) as f:
            f.write(script)

        return bat_path


    def install_update(self, temp_file):
        current_file = os.path.abspath(
            sys.executable
        )

        try:
            if not current_file.lower().endswith(
                ".exe"
            ):
                raise RuntimeError(
                    self.tr("update_not_exe")
                )

            if not os.path.exists(temp_file):
                raise RuntimeError(
                    self.tr("invalid_update")
                )

            updater = self.create_updater_script(
                current_file,
                temp_file,
            )

            subprocess.Popen(
                [
                    "cmd.exe",
                    "/c",
                    updater,
                ],
                cwd=os.path.dirname(
                    current_file
                ),
                creationflags=getattr(
                    subprocess,
                    "CREATE_NO_WINDOW",
                    0,
                ),
            )

            try:
                keyboard.unhook_all()
            except Exception:
                pass

            self.destroy()

        except Exception as e:
            messagebox.showerror(
                "NetBlocker",
                self.tr(
                    "install_error",
                    error=e,
                ),
            )


    def dismiss_update(self):
        self.cfg["skip_version"] = (
            self._pending_version
        )

        save_config(self.cfg)
        self._update_temp = None

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
                    self.q.put(a),
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
                    hk,
                )
            )

        threading.Thread(
            target=worker,
            daemon=True,
        ).start()


    def on_recorded(
        self,
        action,
        hk,
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
                    self.tr(
                        "hotkey_used",
                        hotkey=fmt_hotkey(hk),
                    ),
                )
            else:
                self.cfg["hotkeys"][action] = hk
                save_config(self.cfg)

        self.register_hotkeys()
        self.refresh_hk_buttons()
        self.update_status()
        self.set_lag_btn(False)


    def clear_hotkey(self, action):
        if self.recording:
            return

        self.cfg["hotkeys"][action] = None

        save_config(self.cfg)

        self.register_hotkeys()
        self.refresh_hk_buttons()
        self.update_status()
        self.set_lag_btn(False)


    def refresh_hk_buttons(self):
        for action, btn in self.hk_btns.items():
            text = (
                self.tr("record_hotkey")
                if self.recording == action
                else (
                    fmt_hotkey(
                        self.cfg["hotkeys"].get(action)
                    )
                    or self.tr("no_hotkey")
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
                        item[2],
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
                            url,
                        )

                        if (
                            self.cfg.get(
                                "auto_update",
                                True,
                            )
                            and self._page == "main"
                        ):
                            self.after(
                                120,
                                lambda u=url: self.begin_update_download(
                                    u,
                                    os.path.abspath(
                                        sys.executable
                                    ),
                                    automatic=True,
                                ),
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
                            url,
                        )

                        self.version_lbl.configure(
                            text=(
                                f"{self.tr('version')}: {VERSION} "
                                f"{self.tr('new_version', version=remote_version)}"
                            )
                        )
                    else:
                        self.version_lbl.configure(
                            text=(
                                f"{self.tr('version')}: {VERSION} "
                                f"{self.tr('latest')}"
                            )
                        )

                elif (
                    isinstance(item, tuple)
                    and item[0] == "update_downloaded"
                ):
                    (
                        _,
                        current_file,
                        temp_file,
                        automatic,
                    ) = item

                    self._update_downloading = False
                    self._update_temp = temp_file

                    self.update_download_btn.configure(
                        text=self.tr("install"),
                        state="normal",
                    )

                    self.update_skip_btn.configure(
                        state="normal",
                    )

                    self.update_lbl.configure(
                        text=(
                            self.tr("downloaded_ready")
                        )
                    )

                    self.pulse_button(
                        self.update_download_btn,
                        "#ffffff",
                        "#dce8ff",
                        times=4,
                        delay=80,
                    )

                    install_now = messagebox.askyesno(
                        "NetBlocker",
                        self.tr("install_now"),
                    )

                    if install_now:
                        self.install_update(
                            temp_file
                        )

                elif (
                    isinstance(item, tuple)
                    and item[0] == "update_error"
                ):
                    _, error = item

                    self._update_downloading = False

                    self.update_lbl.configure(
                        text=self.tr(
                            "download_error_bar"
                        )
                    )

                    self.update_download_btn.configure(
                        text=self.tr("download"),
                        state="normal",
                    )

                    self.update_skip_btn.configure(
                        state="normal",
                    )

                    messagebox.showerror(
                        "NetBlocker",
                        self.tr(
                            "download_failed",
                            error=error,
                        ),
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
            self.poll_queue,
        )


    # -----------------------------------------------------------------
    # LISTA APLIKACJI + WYBÓR LITERĄ
    # -----------------------------------------------------------------

    def refresh(self):
        self.apps = list_apps()

        labels = (
            list(self.apps.keys())
            or [self.tr("none")]
        )

        self.combo.configure(
            values=labels
        )

        label = (
            self.current_label
            if self.current_label in self.apps
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

        self.current_label = (
            label or labels[0]
        )

        self.update_status()

        if hasattr(self, "info"):
            self.flash(
                self.tr("refresh")
            )


    def on_global_keypress(self, event):
        if self._page != "main":
            return

        if self.recording:
            return

        if isinstance(event.widget, tk.Entry):
            return

        key = (event.keysym or "").lower()

        if len(key) != 1 or not (
            "a" <= key <= "z"
            or "0" <= key <= "9"
        ):
            return

        matches = [
            label
            for label in self.apps.keys()
            if label.lower().startswith(key)
        ]

        if not matches:
            self.flash(
                self.tr(
                    "no_process_letter",
                    letter=key.upper(),
                )
            )
            return

        now = time.time()

        if (
            self._key_cycle_letter == key
            and now - self._key_cycle_time < 1.2
        ):
            self._key_cycle_index = (
                self._key_cycle_index + 1
            ) % len(matches)
        else:
            self._key_cycle_letter = key
            self._key_cycle_index = 0

        self._key_cycle_time = now

        label = matches[
            self._key_cycle_index
        ]

        self.combo.set(label)
        self.current_label = label

        path = self.apps.get(label)
        if path:
            self.cfg["last_app"] = path
            save_config(self.cfg)

        self.update_status()

        self.pulse_button(
            self.toggle_btn,
            self.T["accent"],
            self.toggle_btn.cget("fg_color"),
            times=2,
            delay=75,
        )

        self.flash(
            self.tr(
                "process_found",
                name=label,
            )
        )


    def on_select(self, label):
        self.current_label = label

        path = self.apps.get(label)

        if path:
            self.cfg["last_app"] = path
            save_config(self.cfg)

        self.update_status()


    def selected_path(self):
        return self.apps.get(
            self.combo.get()
        )


    def toggle_text(self, blocked):
        base = (
            self.tr("restore_internet")
            if blocked
            else self.tr("cut_internet")
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
            f"{self.tr('lag')}   [{fmt_hotkey(hk)}]"
            if hk
            else self.tr("lag")
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
                    text=self.tr("choose_app")
                )

                self.dot.configure(
                    text_color=T["muted"]
                )

                self.toggle_btn.configure(
                    text=self.toggle_text(False),
                    fg_color=T["red"],
                    hover_color=T["red_h"],
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
                    text=self.tr("blocked")
                )

                self.toggle_btn.configure(
                    text=self.toggle_text(True),
                    fg_color=T["green"],
                    hover_color=T["green_h"],
                )
            else:
                self.dot.configure(
                    text_color=T["green"]
                )

                self.status.configure(
                    text=self.tr("working")
                )

                self.toggle_btn.configure(
                    text=self.toggle_text(False),
                    fg_color=T["red"],
                    hover_color=T["red_h"],
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
                lambda: self._clear_info(msg),
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
        silent=False,
    ):
        T = self.T
        path = self.selected_path()

        if not path:
            if silent:
                self.flash(
                    self.tr("no_app_short")
                )
            else:
                messagebox.showinfo(
                    "NetBlocker",
                    self.tr("no_app"),
                )
            return

        try:
            if is_blocked(path):
                remove_rules(
                    path,
                    "block",
                )

                self.flash(
                    self.tr("internet_restored")
                )

                self.animate_status(
                    self.status,
                    self.tr("working"),
                    T["green"],
                )

            else:
                ensure_rules(
                    path,
                    "block",
                    True,
                )

                self.flash(
                    self.tr("internet_cut")
                )

                self.animate_status(
                    self.status,
                    self.tr("blocked"),
                    T["red"],
                )

            self.pulse_button(
                self.toggle_btn,
                T["accent"],
                self.toggle_btn.cget("fg_color"),
                times=2,
                delay=90,
            )

        except Exception as e:
            messagebox.showerror(
                "NetBlocker",
                self.tr(
                    "firewall_error",
                    error=e,
                ),
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
                    self.tr("lag_running")
                    if running
                    else self.lag_text()
                ),
                state=(
                    "disabled"
                    if running
                    else "normal"
                ),
            )
        except Exception:
            pass


    def do_lag(
        self,
        silent=False,
    ):
        path = self.selected_path()

        if not path:
            if silent:
                self.flash(
                    self.tr("no_app_short")
                )
            else:
                messagebox.showinfo(
                    "NetBlocker",
                    self.tr("no_app"),
                )
            return

        if self.lag_running:
            return

        try:
            if path not in self.lag_paths:
                ensure_rules(
                    path,
                    "lag",
                    False,
                )
                self.lag_paths.add(
                    path
                )

            self.lag_running = True

            self.set_lag_btn(True)

            self.pulse_button(
                self.lag_btn,
                self.T["red"],
                self.T["orange"],
                times=int(
                    max(
                        2,
                        self.cfg["lag_ms"] / 120,
                    )
                ),
                delay=100,
            )

            set_enabled(
                path,
                "lag",
                True,
            )

            self.after(
                int(self.cfg["lag_ms"]),
                lambda: self.end_lag(path),
            )

        except Exception as e:
            self.lag_running = False

            self.set_lag_btn(False)

            messagebox.showerror(
                "NetBlocker",
                self.tr(
                    "lag_error",
                    error=e,
                ),
            )


    def end_lag(self, path):
        set_enabled(
            path,
            "lag",
            False,
        )

        self.lag_running = False

        self.set_lag_btn(False)

        self.flash(
            self.tr("lag_finished")
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
                "lag",
            )

        save_config(
            self.cfg
        )

        self.destroy()


# ---------------------------------------------------------------------
# START
# ---------------------------------------------------------------------

if __name__ == "__main__":

    if sys.platform != "win32":
        sys.exit(
            "Ten program działa tylko na Windows."
        )

    if not is_admin():
        relaunch_as_admin()

    cleanup_stale_lag_rules()

    App().mainloop()
'''

path = Path("/mnt/data/net_blocker_v1_3_0.py")
path.write_text(code, encoding="utf-8")
print(path)
print(f"Rozmiar: {path.stat().st_size} B")
