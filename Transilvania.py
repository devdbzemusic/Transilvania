import logging
import os
import shutil
import sys
import threading
import time
import tkinter as tk
import webbrowser
import ctypes
import re
from ctypes import wintypes
from pathlib import Path
from tkinter import messagebox
from tkinter.scrolledtext import ScrolledText

import pyautogui
import pystray
import pytesseract
import requests
from deep_translator import GoogleTranslator
from PIL import Image, ImageDraw, ImageGrab, ImageOps, ImageTk

TESSERACT_URL = "https://github.com/UB-Mannheim/tesseract/wiki"
PROJECT_URL = "https://github.com/devdbzemusic/Transilvania"

logging.basicConfig(
    filename="transilvania.log",
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)


class TranslationApp:
    def __init__(self):
        self._enable_dpi_awareness()
        self.hotkey_key = "d"
        self.hotkey_combo = f"<ctrl>+{self.hotkey_key}"
        self.window_hotkey_combo = f"<ctrl>+<shift>+{self.hotkey_key}"
        self.ocr_languages = ["eng", "rus", "ukr", "ara"]
        self.listener = None
        self.hotkey = None
        self.window_hotkey = None
        self.hotkey_thread = None
        self.hotkey_thread_id = None
        self.icon = None
        self.overlay = None
        self.last_trigger_ts = 0.0
        self.logo_path = self._find_logo_path()
        self.bg_path = self._find_background_path()
        self.local_tessdata_dir = self._local_tessdata_dir()
        self.available_ocr_languages = []
        self.tesseract_path = None
        self.tesseract_ready = False
        self.tk_logo = None
        self.bg_photo = None
        self.about_window = None

        self.root = tk.Tk()
        self.root.title("Transilvania - Einstellungen")
        self.root.geometry("460x620")
        self.root.resizable(False, False)
        self.root.configure(bg="#0b0b0b")
        self.root.protocol("WM_DELETE_WINDOW", self.hide_to_background)
        self._apply_window_icon()

        self._build_settings_ui()
        self.tesseract_ready = self.ensure_tesseract_available()
        if self.tesseract_ready:
            self.ensure_ocr_languages()
            self.requirements_label.config(text="Tesseract: OK", fg="#8ef08e")
        else:
            self.requirements_label.config(text="Tesseract: NICHT installiert", fg="#ff7a7a")

        self.start_listener()
        self.start_tray_icon()
        logging.info(
            "App gestartet. Hotkeys=STRG+%s | STRG+SHIFT+%s",
            self.hotkey_key.upper(),
            self.hotkey_key.upper(),
        )

    def _enable_dpi_awareness(self):
        try:
            user32 = ctypes.windll.user32
            if hasattr(ctypes.windll, "shcore"):
                ctypes.windll.shcore.SetProcessDpiAwareness(2)
            else:
                user32.SetProcessDPIAware()
        except Exception:
            logging.info("DPI-Awareness konnte nicht gesetzt werden (ok).")

    def _resource_dirs(self):
        dirs = []
        if hasattr(sys, "_MEIPASS"):
            dirs.append(Path(sys._MEIPASS))
        dirs.append(Path(__file__).resolve().parent)
        dirs.append(Path.cwd())
        return dirs

    def _find_logo_path(self):
        names = (
            "resources/dbz.ico",
            "resources/logo.ico",
            "logo.ico",
            "logo.png",
            "app_logo.ico",
            "app_logo.png",
            "dbz.ico",
        )
        for base in self._resource_dirs():
            for name in names:
                candidate = base / name
                if candidate.exists():
                    return candidate
        return None

    def _find_background_path(self):
        candidates = (
            Path("resources") / "dbzs_logo_bg.png",
            Path("dbzs_logo_bg.png"),
        )
        for base in self._resource_dirs():
            for rel in candidates:
                candidate = base / rel
                if candidate.exists():
                    return candidate
        return None

    def _find_readme_path(self):
        for base in self._resource_dirs():
            candidate = base / "README.md"
            if candidate.exists():
                return candidate
        return None

    def _local_tessdata_dir(self):
        base = Path(os.getenv("LOCALAPPDATA", str(Path.home())))
        return base / "Transilvania" / "tessdata"

    def _system_tessdata_dir(self):
        if not self.tesseract_path:
            return None
        candidate = Path(self.tesseract_path).parent / "tessdata"
        if candidate.exists():
            return candidate
        return None

    def _resolve_tesseract_path(self):
        candidates = [
            Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
            Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
        ]

        for candidate in candidates:
            if candidate.exists():
                return str(candidate)

        path_hit = shutil.which("tesseract")
        if path_hit:
            return path_hit

        return None

    def ensure_tesseract_available(self):
        self.tesseract_path = self._resolve_tesseract_path()
        if self.tesseract_path:
            pytesseract.pytesseract.tesseract_cmd = self.tesseract_path
            logging.info("Tesseract gefunden: %s", self.tesseract_path)
            return True

        logging.warning("Tesseract nicht gefunden.")
        open_link = messagebox.askyesno(
            "Tesseract fehlt",
            "Tesseract-OCR ist nicht installiert.\n\n"
            "Ohne Tesseract funktioniert OCR auf nicht-markiertem Text nicht.\n"
            "Markierter Text kann weiterhin uebersetzt werden.\n\n"
            "Installationsseite jetzt oeffnen?",
        )
        if open_link:
            webbrowser.open(TESSERACT_URL)
        return False

    def ensure_ocr_languages(self):
        self.local_tessdata_dir.mkdir(parents=True, exist_ok=True)
        missing = []
        system_tess = self._system_tessdata_dir()

        for lang in self.ocr_languages:
            local_file = self.local_tessdata_dir / f"{lang}.traineddata"
            if local_file.exists():
                continue

            if system_tess:
                system_file = system_tess / f"{lang}.traineddata"
                if system_file.exists():
                    shutil.copy2(system_file, local_file)
                    logging.info("Sprache aus System uebernommen: %s", lang)
                    continue

            missing.append(lang)

        for lang in missing:
            self._download_lang(lang)

        available = []
        unresolved = []
        for lang in self.ocr_languages:
            traineddata = self.local_tessdata_dir / f"{lang}.traineddata"
            if traineddata.exists():
                available.append(lang)
            else:
                unresolved.append(lang)

        self.available_ocr_languages = available
        if unresolved:
            logging.warning(
                "OCR Sprachdateien fehlen weiterhin: %s",
                "+".join(unresolved),
            )
        if not self.available_ocr_languages:
            logging.error("Keine OCR Sprachdateien verfuegbar.")

        logging.info(
            "OCR Sprachordner: %s | Verfuegbar: %s",
            self.local_tessdata_dir,
            "+".join(self.available_ocr_languages) or "keine",
        )

    def _download_lang(self, lang):
        url = (
            "https://raw.githubusercontent.com/tesseract-ocr/tessdata_fast/main/"
            f"{lang}.traineddata"
        )
        target = self.local_tessdata_dir / f"{lang}.traineddata"
        try:
            logging.info("Lade Sprache herunter: %s", lang)
            with requests.get(url, timeout=90, stream=True) as resp:
                resp.raise_for_status()
                with open(target, "wb") as fh:
                    for chunk in resp.iter_content(chunk_size=1024 * 256):
                        if chunk:
                            fh.write(chunk)
            logging.info("Sprache installiert: %s", lang)
        except Exception:
            logging.exception("Download fehlgeschlagen fuer Sprache: %s", lang)

    def _apply_window_icon(self):
        if not self.logo_path:
            logging.info("Kein Logo gefunden, nutze Standard-Icon.")
            return
        try:
            if self.logo_path.suffix.lower() == ".ico":
                self.root.iconbitmap(default=str(self.logo_path))
            else:
                self.tk_logo = tk.PhotoImage(file=str(self.logo_path))
                self.root.iconphoto(True, self.tk_logo)
            logging.info("Window-Icon geladen: %s", self.logo_path)
        except Exception:
            logging.exception("Window-Icon konnte nicht geladen werden.")

    def _build_settings_ui(self):
        root_frame = tk.Frame(self.root, bg="#0b0b0b")
        root_frame.pack(fill="both", expand=True)

        if self.bg_path:
            try:
                bg_img = Image.open(self.bg_path).convert("RGB")
                bg_img = ImageOps.contain(bg_img, (436, 320))
                self.bg_photo = ImageTk.PhotoImage(bg_img)
                bg_label = tk.Label(root_frame, image=self.bg_photo, bd=0, bg="#0b0b0b")
                bg_label.pack(padx=12, pady=(12, 8))
                logging.info("Hintergrundbild geladen: %s", self.bg_path)
            except Exception:
                logging.exception("Hintergrundbild konnte nicht geladen werden.")

        panel = tk.Frame(root_frame, padx=14, pady=12, bg="#000000")
        panel.pack(fill="x", padx=12, pady=(0, 12))

        tk.Label(
            panel,
            text="Tastenkombi: STRG + Taste",
            fg="white",
            bg="#000000",
            anchor="w",
        ).pack(fill="x")

        self.hotkey_var = tk.StringVar(value=self.hotkey_key)
        self.hotkey_entry = tk.Entry(panel, textvariable=self.hotkey_var, width=6, justify="center")
        self.hotkey_entry.pack(fill="x", pady=(6, 0))
        self.hotkey_entry.bind("<KeyRelease>", self._on_hotkey_input_change)
        self.hotkey_entry.bind("<FocusOut>", self._on_hotkey_input_change)
        self.hotkey_entry.bind("<Return>", self._on_hotkey_input_change)

        self.status_label = tk.Label(
            panel,
            text=(
                f"Aktiv: STRG + {self.hotkey_key.upper()} "
                f"| Fenster: STRG + SHIFT + {self.hotkey_key.upper()}"
            ),
            fg="#8ef08e",
            bg="#000000",
            anchor="w",
            justify="left",
            wraplength=410,
        )
        self.status_label.pack(fill="x", pady=(6, 0))

        self.requirements_label = tk.Label(
            panel,
            text="Tesseract: pruefe...",
            fg="#d0d0d0",
            bg="#000000",
            anchor="w",
        )
        self.requirements_label.pack(fill="x", pady=(2, 0))

        tk.Button(panel, text="Im Hintergrund laufen", command=self.hide_to_background).pack(
            fill="x", pady=(10, 0)
        )

        tk.Label(
            panel,
            text="Erst markierter Text (ohne Zwischenablage), dann Fenstertext.",
            font=("Arial", 9),
            fg="#d0d0d0",
            bg="#000000",
        ).pack(pady=(8, 0))

        footer = tk.Frame(root_frame, bg="#0b0b0b")
        footer.pack(side="bottom", fill="x", pady=(0, 12))
        tk.Button(
            footer,
            text="About / Projekt",
            command=self.open_about_dialog,
            width=20,
        ).pack(anchor="center")

    def _preprocess_for_ocr(self, image):
        gray = ImageOps.grayscale(image)
        scale = 2
        resampling = getattr(Image, "Resampling", Image)
        enlarged = gray.resize((gray.width * scale, gray.height * scale), resampling.LANCZOS)
        return ImageOps.autocontrast(enlarged)

    def _extract_text_from_image(self, image):
        processed = self._preprocess_for_ocr(image)
        return self._extract_text_multi_config(processed)

    def _extract_text_multi_config(self, processed_image):
        configs = (
            "--oem 1 --psm 6",
            "--oem 1 --psm 11",
            "--oem 1 --psm 3",
        )
        best_text = ""
        lang = "+".join(self.available_ocr_languages)
        tessdata_dir = str(self.local_tessdata_dir)
        for cfg in configs:
            try:
                text = pytesseract.image_to_string(
                    processed_image,
                    lang=lang,
                    config=f"{cfg} --tessdata-dir {tessdata_dir}",
                )
                text = re.sub(r"\s+", " ", (text or "")).strip()
                if len(text) > len(best_text):
                    best_text = text
            except Exception:
                logging.exception("OCR-Konfiguration fehlgeschlagen: %s", cfg)
                try:
                    text = pytesseract.image_to_string(
                        processed_image,
                        lang=lang,
                        config=cfg,
                    )
                    text = re.sub(r"\s+", " ", (text or "")).strip()
                    if len(text) > len(best_text):
                        best_text = text
                except Exception:
                    pass
        return best_text

    def _read_window_text(self, hwnd):
        user32 = ctypes.windll.user32
        length = user32.GetWindowTextLengthW(hwnd)
        if length <= 0:
            return ""
        buf = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, buf, length + 1)
        return (buf.value or "").strip()

    def _extract_text_from_foreground_window(self):
        try:
            user32 = ctypes.windll.user32
            fg = user32.GetForegroundWindow()
            if not fg:
                return ""

            texts = []
            window_title = self._read_window_text(fg)
            if len(window_title) >= 2:
                texts.append(window_title)

            enum_proc_type = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)

            @enum_proc_type
            def _enum_proc(hwnd, _lparam):
                txt = self._read_window_text(hwnd)
                if len(txt) >= 2:
                    texts.append(txt)
                return True

            user32.EnumChildWindows(fg, _enum_proc, 0)
            unique = []
            seen = set()
            for t in texts:
                if t not in seen:
                    seen.add(t)
                    unique.append(t)
            return "\n".join(unique).strip()
        except Exception:
            logging.exception("Foreground-Window-Text konnte nicht gelesen werden.")
            return ""

    def _get_window_bbox_at_point(self, x, y):
        try:
            user32 = ctypes.windll.user32
            point = wintypes.POINT(x=x, y=y)
            hwnd = user32.WindowFromPoint(point)
            if not hwnd:
                return None

            # Auf Top-Level-Fenster hochgehen, damit nicht nur ein kleines Control erfasst wird.
            GA_ROOT = 2
            root_hwnd = user32.GetAncestor(hwnd, GA_ROOT)
            if root_hwnd:
                hwnd = root_hwnd

            rect = wintypes.RECT()
            if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
                return None

            left, top, right, bottom = rect.left, rect.top, rect.right, rect.bottom
            width = max(0, right - left)
            height = max(0, bottom - top)
            if width < 80 or height < 40:
                return None
            return (left, top, right, bottom)
        except Exception:
            logging.exception("Fensterrechteck konnte nicht ermittelt werden.")
            return None

    def _fallback_fullscreen_ocr(self):
        try:
            screenshot = ImageGrab.grab()
            return self._extract_text_from_image(screenshot)
        except Exception:
            logging.exception("Fullscreen-OCR fehlgeschlagen.")
            return ""

    def _on_hotkey_input_change(self, event=None):
        raw = (self.hotkey_var.get() or "").strip().lower()
        if not raw:
            return
        key = raw[0]
        if not key.isalnum():
            self.hotkey_var.set(self.hotkey_key)
            return

        if key == self.hotkey_key:
            return

        self.hotkey_key = key
        self.hotkey_combo = f"<ctrl>+{self.hotkey_key}"
        self.window_hotkey_combo = f"<ctrl>+<shift>+{self.hotkey_key}"
        self.restart_listener()
        self.status_label.config(
            text=(
                f"Aktiv: STRG + {self.hotkey_key.upper()} "
                f"| Fenster: STRG + SHIFT + {self.hotkey_key.upper()}"
            )
        )
        self.hotkey_var.set(self.hotkey_key)
        logging.info(
            "Hotkeys geaendert auf STRG+%s und STRG+SHIFT+%s",
            self.hotkey_key.upper(),
            self.hotkey_key.upper(),
        )

    def _get_selected_text_from_focus_control(self):
        # Liest markierten Text direkt aus dem fokussierten Edit/RichEdit-Control
        # ohne die Zwischenablage zu beruehren.
        EM_GETSEL = 0x00B0
        EM_EXGETSEL = 0x0434
        WM_GETTEXT = 0x000D
        WM_GETTEXTLENGTH = 0x000E

        class CHARRANGE(ctypes.Structure):
            _fields_ = [("cpMin", ctypes.c_long), ("cpMax", ctypes.c_long)]

        try:
            user32 = ctypes.windll.user32
            kernel32 = ctypes.windll.kernel32

            fg = user32.GetForegroundWindow()
            if not fg:
                return ""

            fg_thread = user32.GetWindowThreadProcessId(fg, None)
            cur_thread = kernel32.GetCurrentThreadId()
            attached = bool(user32.AttachThreadInput(cur_thread, fg_thread, True))
            try:
                focus = user32.GetFocus()
            finally:
                if attached:
                    user32.AttachThreadInput(cur_thread, fg_thread, False)

            if not focus:
                return ""

            text_len = int(user32.SendMessageW(focus, WM_GETTEXTLENGTH, 0, 0))
            if text_len <= 0:
                return ""
            buf = ctypes.create_unicode_buffer(text_len + 1)
            user32.SendMessageW(focus, WM_GETTEXT, text_len + 1, ctypes.byref(buf))
            full_text = (buf.value or "")
            if not full_text:
                return ""

            rng = CHARRANGE(0, 0)
            user32.SendMessageW(focus, EM_EXGETSEL, 0, ctypes.byref(rng))
            start, end = int(rng.cpMin), int(rng.cpMax)

            if end <= start:
                packed = int(user32.SendMessageW(focus, EM_GETSEL, 0, 0))
                start = packed & 0xFFFF
                end = (packed >> 16) & 0xFFFF

            if end > start and start >= 0:
                return full_text[start:end].strip()
            return ""
        except Exception:
            logging.exception("Markierter Text konnte ohne Zwischenablage nicht gelesen werden.")
            return ""

    def perform_translate(self, prefer_clipboard, force_window, use_ocr_fallback):
        try:
            x, y = pyautogui.position()
            text = ""

            if prefer_clipboard:
                text = self._get_selected_text_from_focus_control()
                if text:
                    logging.info("Text aus Markierung gelesen: %r", text)

            if not text:
                window_text = self._extract_text_from_foreground_window()
                if window_text:
                    text = window_text
                    logging.info("Text aus aktivem Fenster gelesen: %r", text)

            if not text and use_ocr_fallback:
                if not self.tesseract_ready:
                    self.root.after(
                        0,
                        lambda: self.show_overlay(
                            "Tesseract fehlt. Installiere es, um OCR ohne Markierung zu nutzen.",
                            x,
                            max(10, y - 50),
                        ),
                    )
                    return

                if not self.available_ocr_languages:
                    self.root.after(
                        0,
                        lambda: self.show_overlay(
                        "Keine OCR-Sprachdateien verfuegbar. Pruefe Internet/Tesseract-Setup.",
                            x,
                            max(10, y - 50),
                        ),
                    )
                    return

                if not force_window:
                    bbox = (x - 170, y - 55, x + 170, y + 55)
                    screenshot = ImageGrab.grab(bbox)
                    text = self._extract_text_from_image(screenshot)
                    logging.info("OCR Mausbereich bei (%s,%s), OCR=%r", x, y, text)

                if force_window or len(text) < 8:
                    window_bbox = self._get_window_bbox_at_point(x, y)
                    if window_bbox:
                        window_shot = ImageGrab.grab(window_bbox)
                        window_text = self._extract_text_from_image(window_shot)
                        if len(window_text) > len(text):
                            text = window_text
                        logging.info(
                            "OCR Fensterbereich bei (%s,%s), bbox=%s, OCR=%r",
                            x,
                            y,
                            window_bbox,
                            text,
                        )
                    elif force_window:
                        logging.info("Kein Fenster unter Maus gefunden, nutze Fullscreen-OCR.")

                if len(text) < 8:
                    fullscreen_text = self._fallback_fullscreen_ocr()
                    if len(fullscreen_text) > len(text):
                        text = fullscreen_text
                    logging.info("OCR Fullscreen-Fallback, OCR=%r", text)

            if not text or len(text) < 2:
                msg = "Kein Text erkannt (Markierung/Fenstertext)."
                if use_ocr_fallback:
                    msg = "Kein Text erkannt (weder Markierung noch OCR)."
                self.root.after(
                    0,
                    lambda: self.show_overlay(
                        msg,
                        x,
                        max(10, y - 50),
                    ),
                )
                return

            translation = GoogleTranslator(source="auto", target="de").translate(text)
            logging.info("Uebersetzung=%r", translation)
            self.root.after(0, lambda: self.show_overlay(translation, x, max(10, y - 50)))
        except Exception as exc:
            logging.exception("Fehler bei Translation")
            self.root.after(0, lambda: self.show_overlay(f"Fehler: {exc}", 30, 30))

    def show_overlay(self, text, x, y):
        if self.overlay and self.overlay.winfo_exists():
            self.overlay.destroy()

        self.overlay = tk.Toplevel(self.root)
        self.overlay.overrideredirect(True)
        self.overlay.attributes("-topmost", True)
        self.overlay.attributes("-alpha", 0.92)
        self.overlay.configure(bg="black")
        self.overlay.geometry(f"+{x}+{y}")

        label = tk.Label(
            self.overlay,
            text=text,
            fg="white",
            bg="black",
            font=("Segoe UI", 11, "bold"),
            padx=10,
            pady=6,
            justify="left",
            wraplength=500,
        )
        label.pack()
        self.overlay.after(4500, self.overlay.destroy)

    def on_hotkey_pressed(self):
        now = time.monotonic()
        if now - self.last_trigger_ts < 0.7:
            return
        self.last_trigger_ts = now
        logging.info("Hotkey erkannt: STRG+%s", self.hotkey_key.upper())
        threading.Thread(
            target=self.perform_translate,
            args=(True, False, False),
            daemon=True,
        ).start()

    def on_window_hotkey_pressed(self):
        now = time.monotonic()
        if now - self.last_trigger_ts < 0.7:
            return
        self.last_trigger_ts = now
        logging.info("Fenster-Hotkey erkannt: STRG+SHIFT+%s", self.hotkey_key.upper())
        threading.Thread(
            target=self.perform_translate,
            args=(False, True, True),
            daemon=True,
        ).start()

    def create_tray_icon(self):
        image = None
        if self.logo_path:
            try:
                image = Image.open(self.logo_path).convert("RGBA").resize((64, 64))
                logging.info("Tray-Icon geladen: %s", self.logo_path)
            except Exception:
                logging.exception("Tray-Icon konnte nicht geladen werden.")

        if image is None:
            image = Image.new("RGB", (64, 64), color=(40, 40, 40))
            draw = ImageDraw.Draw(image)
            draw.rectangle([10, 10, 54, 54], fill=(0, 120, 215))
            draw.text((22, 18), "Tr", fill="white")

        menu = pystray.Menu(
            pystray.MenuItem("Einstellungen", self.show_settings_from_tray),
            pystray.MenuItem("Beenden", self.quit_app),
        )
        self.icon = pystray.Icon("Transilvania", image, "Transilvania OCR", menu)
        self.icon.run()

    def start_tray_icon(self):
        threading.Thread(target=self.create_tray_icon, daemon=True).start()

    def start_listener(self):
        def _hotkey_loop():
            WM_HOTKEY = 0x0312
            WM_QUIT = 0x0012
            MOD_CONTROL = 0x0002
            MOD_SHIFT = 0x0004
            HK_ID_NORMAL = 1
            HK_ID_WINDOW = 2

            user32 = ctypes.windll.user32
            kernel32 = ctypes.windll.kernel32
            self.hotkey_thread_id = kernel32.GetCurrentThreadId()

            vk = ord(self.hotkey_key.upper())
            ok_normal = bool(user32.RegisterHotKey(None, HK_ID_NORMAL, MOD_CONTROL, vk))
            ok_window = bool(
                user32.RegisterHotKey(None, HK_ID_WINDOW, MOD_CONTROL | MOD_SHIFT, vk)
            )

            if not ok_normal:
                logging.error("RegisterHotKey fehlgeschlagen fuer STRG+%s", self.hotkey_key.upper())
            if not ok_window:
                logging.error(
                    "RegisterHotKey fehlgeschlagen fuer STRG+SHIFT+%s",
                    self.hotkey_key.upper(),
                )
            if not (ok_normal or ok_window):
                return

            logging.info(
                "Keyboard-Listener gestartet (%s | %s).",
                self.hotkey_combo,
                self.window_hotkey_combo,
            )

            msg = wintypes.MSG()
            while True:
                result = user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
                if result == 0 or msg.message == WM_QUIT:
                    break
                if msg.message == WM_HOTKEY:
                    if msg.wParam == HK_ID_NORMAL:
                        self.on_hotkey_pressed()
                    elif msg.wParam == HK_ID_WINDOW:
                        self.on_window_hotkey_pressed()
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageW(ctypes.byref(msg))

            user32.UnregisterHotKey(None, HK_ID_NORMAL)
            user32.UnregisterHotKey(None, HK_ID_WINDOW)
            self.hotkey_thread_id = None

        self.hotkey_thread = threading.Thread(target=_hotkey_loop, daemon=True)
        self.hotkey_thread.start()

    def restart_listener(self):
        if self.hotkey_thread_id:
            ctypes.windll.user32.PostThreadMessageW(self.hotkey_thread_id, 0x0012, 0, 0)
        self.start_listener()

    def _load_readme_text(self):
        readme_path = self._find_readme_path()
        if not readme_path:
            return "README.md nicht gefunden."
        try:
            return readme_path.read_text(encoding="utf-8", errors="replace")
        except Exception:
            logging.exception("README konnte nicht gelesen werden.")
            return "README.md konnte nicht gelesen werden."

    def open_project_link(self):
        webbrowser.open(PROJECT_URL)

    def open_about_dialog(self):
        if self.about_window and self.about_window.winfo_exists():
            self.about_window.lift()
            self.about_window.focus_force()
            return

        win = tk.Toplevel(self.root)
        win.title("About Transilvania")
        win.geometry("640x500")
        win.configure(bg="#101010")
        win.attributes("-topmost", True)
        self.about_window = win

        title = tk.Label(
            win,
            text="Transilvania",
            fg="white",
            bg="#101010",
            font=("Segoe UI", 15, "bold"),
        )
        title.pack(pady=(12, 4))

        link = tk.Label(
            win,
            text=PROJECT_URL,
            fg="#7fc4ff",
            bg="#101010",
            cursor="hand2",
            font=("Segoe UI", 10, "underline"),
        )
        link.pack(pady=(0, 10))
        link.bind("<Button-1>", lambda _e: self.open_project_link())

        readme_text = ScrolledText(
            win,
            wrap="word",
            bg="#171717",
            fg="#f0f0f0",
            insertbackground="#f0f0f0",
            font=("Consolas", 10),
        )
        readme_text.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        readme_text.insert("1.0", self._load_readme_text())
        readme_text.config(state="disabled")

    def hide_to_background(self):
        self.root.withdraw()

    def show_settings_from_tray(self, icon=None, item=None):
        def _show():
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()

        self.root.after(0, _show)

    def quit_app(self, icon=None, item=None):
        def _shutdown():
            if self.hotkey_thread_id:
                ctypes.windll.user32.PostThreadMessageW(self.hotkey_thread_id, 0x0012, 0, 0)
            if self.icon:
                self.icon.stop()
            self.root.quit()
            self.root.destroy()
            sys.exit(0)

        self.root.after(0, _shutdown)

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = TranslationApp()
    app.run()
