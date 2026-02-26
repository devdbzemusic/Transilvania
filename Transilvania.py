import logging
import os
import shutil
import sys
import threading
import time
import tkinter as tk
import webbrowser
from pathlib import Path
from tkinter import messagebox

import pyautogui
import pyperclip
import pystray
import pytesseract
import requests
from deep_translator import GoogleTranslator
from PIL import Image, ImageDraw, ImageGrab, ImageOps, ImageTk
from pynput import keyboard

TESSERACT_URL = "https://github.com/UB-Mannheim/tesseract/wiki"

logging.basicConfig(
    filename="transilvania.log",
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)


class TranslationApp:
    def __init__(self):
        self.hotkey_key = "d"
        self.hotkey_combo = f"<ctrl>+{self.hotkey_key}"
        self.ocr_languages = ["eng", "rus", "ara"]
        self.listener = None
        self.hotkey = None
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
        logging.info("App gestartet. Hotkey=STRG+%s", self.hotkey_key.upper())

    def _resource_dirs(self):
        dirs = []
        if hasattr(sys, "_MEIPASS"):
            dirs.append(Path(sys._MEIPASS))
        dirs.append(Path(__file__).resolve().parent)
        dirs.append(Path.cwd())
        return dirs

    def _find_logo_path(self):
        names = ("logo.ico", "logo.png", "app_logo.ico", "app_logo.png", "dbz.ico")
        for base in self._resource_dirs():
            for name in names:
                candidate = base / name
                if candidate.exists():
                    return candidate
        return None

    def _find_background_path(self):
        candidates = (
            Path("assets") / "logos" / "dbzs_logo_bg.png",
            Path("dbzs_logo_bg.png"),
        )
        for base in self._resource_dirs():
            for rel in candidates:
                candidate = base / rel
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
            text=f"Aktiv: STRG + {self.hotkey_key.upper()}",
            fg="#8ef08e",
            bg="#000000",
            anchor="w",
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
            text="Erst Markierung (Ctrl+C), sonst OCR an Maus.",
            font=("Arial", 9),
            fg="#d0d0d0",
            bg="#000000",
        ).pack(pady=(8, 0))

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
        self.restart_listener()
        self.status_label.config(text=f"Aktiv: STRG + {self.hotkey_key.upper()}")
        self.hotkey_var.set(self.hotkey_key)
        logging.info("Hotkey geaendert auf STRG+%s", self.hotkey_key.upper())

    def _get_selected_text_from_clipboard(self):
        old_clipboard = ""
        clipboard_read_ok = False
        try:
            old_clipboard = pyperclip.paste()
            clipboard_read_ok = True
        except Exception:
            pass

        try:
            pyautogui.hotkey("ctrl", "c")
            time.sleep(0.12)
            selected = (pyperclip.paste() or "").strip()
            if selected:
                return selected
            return ""
        except Exception:
            logging.exception("Clipboard-Markierung konnte nicht gelesen werden.")
            return ""
        finally:
            if clipboard_read_ok:
                try:
                    pyperclip.copy(old_clipboard)
                except Exception:
                    pass

    def perform_translate(self, prefer_clipboard):
        try:
            x, y = pyautogui.position()
            text = ""

            if prefer_clipboard:
                text = self._get_selected_text_from_clipboard()
                if text:
                    logging.info("Text aus Markierung gelesen: %r", text)

            if not text:
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

                bbox = (x - 150, y - 40, x + 150, y + 40)
                screenshot = ImageGrab.grab(bbox)
                screenshot = ImageOps.grayscale(screenshot)
                text = pytesseract.image_to_string(
                    screenshot,
                    lang="+".join(self.available_ocr_languages),
                    config=f'--tessdata-dir "{self.local_tessdata_dir}"',
                ).strip()
                logging.info("OCR fallback bei Maus=(%s,%s), OCR=%r", x, y, text)

            if not text or len(text) < 2:
                self.root.after(
                    0,
                    lambda: self.show_overlay(
                        "Kein Text erkannt (weder Markierung noch OCR).",
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
        threading.Thread(target=self.perform_translate, args=(True,), daemon=True).start()

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
        self.hotkey = keyboard.HotKey(
            keyboard.HotKey.parse(self.hotkey_combo),
            self.on_hotkey_pressed,
        )
        self.listener = keyboard.Listener(
            on_press=self._on_key_press,
            on_release=self._on_key_release,
        )
        self.listener.start()
        logging.info("Keyboard-Listener gestartet (%s).", self.hotkey_combo)

    def restart_listener(self):
        if self.listener:
            self.listener.stop()
        self.start_listener()

    def _on_key_press(self, key):
        if self.listener and self.hotkey:
            self.hotkey.press(self.listener.canonical(key))

    def _on_key_release(self, key):
        if self.listener and self.hotkey:
            self.hotkey.release(self.listener.canonical(key))

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
            if self.listener:
                self.listener.stop()
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
