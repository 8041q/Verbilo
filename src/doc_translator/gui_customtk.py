# A small CustomTkinter-based GUI front-end that wraps the existing translate_file API.

# This UI provides file/folder selection, a language dropdown, translator selector, output folder,
# start/stop controls, a per-file status list, and a simple log area. Translations run in a
# background thread using `gui_helpers.Worker` and report coarse per-file progress back to the UI.

from __future__ import annotations

import os
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk
import logging

try:
    import customtkinter as ctk
except Exception:
    ctk = None

from .gui_helpers import Worker, list_supported_files, center_window
from .gui_config import load_config, save_config


def _get_language_options() -> list[tuple[str, str]]:
    # Try to probe deep-translator for supported languages, else fallback to a small list
    logger = logging.getLogger(__name__)
    # common fallback to use when the probed list is too large or unavailable
    common_fallback = {
        "en": "English",
        "es": "Spanish",
        "fr": "French",
        "de": "German",
        "pt": "Portuguese",
        "zh": "Chinese",
        "ja": "Japanese",
        "ru": "Russian",
        "it": "Italian",
        "nl": "Dutch",
    }
    try:
        from deep_translator import GoogleTranslator
        langs = None
        # Try class method first, then instance method, then SUPPORTED_LANGUAGES
        getlangs = getattr(GoogleTranslator, "get_supported_languages", None)
        if callable(getlangs):
            try:
                langs = getlangs()
            except TypeError:
                try:
                    langs = GoogleTranslator().get_supported_languages()
                except Exception:
                    langs = None
        elif hasattr(GoogleTranslator, "SUPPORTED_LANGUAGES"):
            langs = getattr(GoogleTranslator, "SUPPORTED_LANGUAGES")

        if isinstance(langs, dict):
            # assume mapping code->name
            try:
                return [(str(code), str(name)) for code, name in langs.items()]
            except Exception:
                logger.exception("Unexpected dict shape from deep_translator supported languages")
        if isinstance(langs, (list, tuple)):
            # list may be codes or names; treat entries as codes
            # If deep-translator returns a very large list, fallback to a small common subset
            try:
                if len(langs) > 40:
                    logger.warning("Deep-translator returned a large language list; using common subset")
                    return list(common_fallback.items())
            except Exception:
                pass
            return [(str(c), str(c)) for c in langs]
    except Exception:
        logger.exception("Failed to probe deep_translator for supported languages")

    # Fallback mapping of common languages
    fallback = {
        "en": "English",
        "es": "Spanish",
        "fr": "French",
        "de": "German",
        "pt": "Portuguese",
        "zh": "Chinese",
        "ja": "Japanese",
        "ru": "Russian",
        "it": "Italian",
        "nl": "Dutch",
        "sv": "Swedish",
        "no": "Norwegian",
        "da": "Danish",
        "fi": "Finnish",
        "pl": "Polish",
    }
    return list(fallback.items())


def _ensure_ctk():
    if ctk is None:
        messagebox.showerror(
            "Missing dependency",
            "customtkinter is required for GUI. Install it via:\n\n    pip install customtkinter",
        )
        raise RuntimeError("customtkinter not installed")


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.worker = Worker()
        self.files: list[str] = []
        self.cfg = load_config() or {}
        self.total_files = 0
        self.completed_files = 0

        if ctk:
            ctk.set_appearance_mode("System")
            ctk.set_default_color_theme("blue")

        self._build_ui()
        # apply defaults from config
        default_out = self.cfg.get("default_output")
        if default_out:
            self.output_entry.insert(0, default_out)
        default_in = self.cfg.get("default_input")
        if default_in:
            found = list_supported_files(default_in)
            for f in found:
                if f not in self.files:
                    self.files.append(f)
                    self.listbox.insert(tk.END, f"{os.path.basename(f)} - pending")

    def _initialdir_for_input(self) -> str:
        # Priority: first file's parent -> config default_input -> cwd
        try:
            if self.files:
                return str(Path(self.files[0]).parent)
        except Exception:
            pass
        if self.cfg.get("default_input"):
            return self.cfg.get("default_input") or ""
        return str(Path.cwd())

    def _initialdir_for_output(self) -> str:
        # Priority: UI output entry -> config default_output -> cwd
        try:
            if hasattr(self, "output_entry"):
                val = self.output_entry.get().strip()
                if val:
                    return val
        except Exception:
            pass
        if self.cfg.get("default_output"):
            return str(self.cfg.get("default_output") or "")
        return str(Path.cwd())

    def _build_ui(self):
        if ctk:
            Frame = ctk.CTkFrame
            Button = ctk.CTkButton
            Label = ctk.CTkLabel
            Entry = ctk.CTkEntry
        else:
            Frame = tk.Frame
            Button = tk.Button
            Label = tk.Label
            Entry = tk.Entry

        self.root.title("Doc Translator GUI")
        # desired starting size
        desired_w, desired_h = 900, 600
        try:
            self.root.minsize(800, 560)
        except Exception:
            pass
        # center on screen and prevent manual resize
        self._desired_width = desired_w
        self._desired_height = desired_h
        center_window(self.root, desired_w, desired_h)
        try:
            self.root.resizable(False, False)
        except Exception:
            pass

        top = Frame(self.root)
        top.pack(fill=tk.X, padx=8, pady=8)

        Label(top, text="Files:").grid(row=0, column=0, sticky="w")
        Button(top, text="Add Files", command=self._add_files).grid(row=0, column=1, padx=4)
        Button(top, text="Select Folder", command=self._select_folder).grid(row=0, column=2, padx=4)
        Button(top, text="Clear", command=self._clear_files).grid(row=0, column=3, padx=4)
        # settings cog (kept here)
        Button(top, text="⚙", command=self._open_settings).grid(row=0, column=4, padx=4)

        mid = Frame(self.root)
        mid.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)

        # Left: file list
        left = Frame(mid)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.listbox = tk.Listbox(left)
        self.listbox.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        # Right: controls and logs
        self.right = Frame(mid)
        self.right.pack(side=tk.RIGHT, fill=tk.Y)

        Label(self.right, text="Target language:").pack(anchor="w", padx=4, pady=(8, 0))
        # language dropdown (show "Name (code)")
        self.lang_frame = Frame(self.right)
        self.lang_frame.pack(fill=tk.X, padx=4)
        self.lang_var = tk.StringVar()
        lang_opts = _get_language_options()
        self._lang_map = {f"{name} ({code})": code for code, name in lang_opts}
        values = list(self._lang_map.keys())
        if not values:
            values = ["English (en)"]
        self.lang_var.set(values[0])
        if ctk:
            self.lang_menu = ctk.CTkOptionMenu(self.lang_frame, values=values, variable=self.lang_var)
            self.lang_menu.pack(fill=tk.X, padx=4)
        else:
            self.lang_menu = tk.OptionMenu(self.lang_frame, self.lang_var, *values)
            self.lang_menu.pack(fill=tk.X, padx=4)

        Label(self.right, text="Translator:").pack(anchor="w", padx=4, pady=(8, 0))
        self.translator_var = tk.StringVar(value="auto")
        if ctk:
            self.translator_menu = ctk.CTkOptionMenu(self.right, values=["auto", "identity", "deep"], variable=self.translator_var)
            self.translator_menu.pack(fill=tk.X, padx=4)
        else:
            self.translator_menu = tk.OptionMenu(self.right, self.translator_var, "auto", "identity", "deep")
            self.translator_menu.pack(fill=tk.X, padx=4)
        # repopulate languages when translator changes
        try:
            self.translator_var.trace_add("write", lambda *_: self._on_translator_change())
        except Exception:
            try:
                self.translator_var.trace("w", lambda *_: self._on_translator_change())
            except Exception:
                pass

        Label(self.right, text="Output folder:").pack(anchor="w", padx=4, pady=(8, 0))
        self.output_entry = Entry(self.right)
        self.output_entry.pack(fill=tk.X, padx=4)
        Button(self.right, text="Browse", command=self._select_output).pack(fill=tk.X, padx=4, pady=(2, 2))

        # Progress bar
        if ctk:
            self.progress = ctk.CTkProgressBar(self.right)
        else:
            self.progress = ttk.Progressbar(self.right, orient="horizontal", mode="determinate")
        self.progress.pack(fill=tk.X, padx=4, pady=(4, 8))
        # ensure progress shows as empty at startup
        try:
            if hasattr(self.progress, "set"):
                if ctk and isinstance(self.progress, ctk.CTkProgressBar):
                    self.progress.set(0.0)
                else:
                    self.progress['value'] = 0
                    self.progress['maximum'] = 1
            else:
                self.progress['value'] = 0
                self.progress['maximum'] = 1
        except Exception:
            pass

        # 'Set Default Input' removed from main page; use settings dialog instead

        self.start_btn = Button(self.right, text="Start", command=self._start)
        self.start_btn.pack(fill=tk.X, padx=4, pady=(4, 2))
        self.stop_btn = Button(self.right, text="Stop", command=self._stop, state=tk.DISABLED)
        self.stop_btn.pack(fill=tk.X, padx=4, pady=(2, 8))

        Label(self.right, text="Log:").pack(anchor="w", padx=4)
        self.log = scrolledtext.ScrolledText(self.right, width=40, height=15)
        self.log.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

    def _open_settings(self):
        win = tk.Toplevel(self.root)
        win.title("Settings")
        win.transient(self.root)
        win.grab_set()

        tk.Label(win, text="Default input folder:").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        in_entry = tk.Entry(win, width=60)
        in_entry.grid(row=0, column=1, padx=8, pady=6)
        in_entry.insert(0, self.cfg.get("default_input", ""))

        # Browse button for input default
        def _browse_default_input():
            init = in_entry.get().strip() or self.cfg.get("default_input") or str(Path.cwd())
            d = filedialog.askdirectory(title="Select default input folder", initialdir=init)
            if d:
                in_entry.delete(0, tk.END)
                in_entry.insert(0, d)

        tk.Button(win, text="Browse", command=_browse_default_input).grid(row=0, column=2, padx=4)

        tk.Label(win, text="Default output folder:").grid(row=1, column=0, sticky="w", padx=8, pady=6)
        out_entry = tk.Entry(win, width=60)
        out_entry.grid(row=1, column=1, padx=8, pady=6)
        out_entry.insert(0, self.cfg.get("default_output", ""))

        # Browse button for output default
        def _browse_default_output():
            init = out_entry.get().strip() or self.cfg.get("default_output") or str(Path.cwd())
            d = filedialog.askdirectory(title="Select default output folder", initialdir=init)
            if d:
                out_entry.delete(0, tk.END)
                out_entry.insert(0, d)

        tk.Button(win, text="Browse", command=_browse_default_output).grid(row=1, column=2, padx=4)

        def _save_and_close():
            if not messagebox.askyesno("Confirm", "Save settings and overwrite defaults?"):
                return
            self.cfg['default_input'] = in_entry.get().strip() or None
            self.cfg['default_output'] = out_entry.get().strip() or None
            save_config(self.cfg)
            messagebox.showinfo("Saved", "Settings saved.")
            win.destroy()

        btn_frame = tk.Frame(win)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=8)
        tk.Button(btn_frame, text="Save", command=_save_and_close).pack(side=tk.LEFT, padx=8)
        tk.Button(btn_frame, text="Cancel", command=win.destroy).pack(side=tk.LEFT)

        # center settings window over main window and prevent manual resize
        try:
            win.update_idletasks()
            req_w = win.winfo_reqwidth()
            req_h = win.winfo_reqheight()
            center_window(win, req_w, req_h)
            try:
                win.resizable(False, False)
            except Exception:
                pass
        except Exception:
            pass

    def _add_files(self):
        init = self._initialdir_for_input()
        paths = filedialog.askopenfilenames(title="Select files", initialdir=init)
        for p in paths:
            if p not in self.files:
                self.files.append(p)
                self.listbox.insert(tk.END, f"{os.path.basename(p)} - pending")

    def _select_folder(self):
        init = self._initialdir_for_input()
        d = filedialog.askdirectory(title="Select folder containing files", initialdir=init)
        if not d:
            return
        found = list_supported_files(d)
        if not found:
            messagebox.showinfo("No files", f"No supported files found in {d}")
            return
        for f in found:
            if f not in self.files:
                self.files.append(f)
                self.listbox.insert(tk.END, f"{os.path.basename(f)} - pending")

    def _clear_files(self):
        self.files.clear()
        self.listbox.delete(0, tk.END)

    def _select_output(self):
        init = self._initialdir_for_output()
        d = filedialog.askdirectory(title="Select output folder", initialdir=init)
        if d:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, d)

    def _start(self):
        sel = self.lang_var.get()
        lang = self._lang_map.get(sel)
        if not lang:
            messagebox.showwarning("Missing language", "Please select a target language from the dropdown.")
            return
        if not self.files:
            messagebox.showwarning("No files", "Please add files or select a folder first.")
            return

        # determine output folder: UI -> config default -> cwd/output
        output_ui = self.output_entry.get().strip()
        if output_ui:
            output = output_ui
        elif self.cfg.get("default_output"):
            output = self.cfg.get("default_output")
        else:
            output = str(Path.cwd() / "output")
        # Ensure output is not None
        if output is None:
            output = str(Path.cwd() / "output")
        Path(str(output)).mkdir(parents=True, exist_ok=True)
        # Normalize translator selection: map "auto" (or empty) -> None so factory auto-detects
        sel_trans = (self.translator_var.get() or "").strip()
        norm = sel_trans.lower()
        if not sel_trans or norm in ("auto", "none"):
            translator = None
        else:
            translator = sel_trans

        # Log selection for easier debugging (safe to ignore failures)
        try:
            self._log(f"Starting translation: lang={lang!r}, translator={translator!r}, output={output!r}")
        except Exception:
            pass

        # disable UI
        self.start_btn.configure(state=tk.DISABLED)
        self.stop_btn.configure(state=tk.NORMAL)

        self._update_listbox_all_status("pending")
        self.total_files = len(self.files)
        self.completed_files = 0
        # initialize progress
        try:
            if ctk and isinstance(self.progress, ctk.CTkProgressBar):
                # customtkinter progress
                self.progress.set(0.0)
            else:
                self.progress['value'] = 0
                self.progress['maximum'] = self.total_files
        except Exception:
            pass
        self.worker.start(self.files, lang, output, translator, self._progress_cb, self._log)

    def _stop(self):
        # request cancellation; do not re-enable Start until worker finishes
        self.worker.stop()
        self.stop_btn.configure(state=tk.DISABLED)
        self._log("Cancellation requested")

    

    def _on_translator_change(self):
        # rebuild language options depending on translator selection
        t = self.translator_var.get() or "auto"
        # if deep requested, attempt to probe deep-translator; otherwise use fallback
        opts = _get_language_options()
        # rebuild map with friendly labels
        new_map = {f"{name} ({code})": code for code, name in opts}
        self._lang_map = new_map
        vals = list(new_map.keys())
        if not vals:
            vals = ["English (en)"]
        # update option menu
        try:
            if ctk:
                # Destroy the old OptionMenu and create a new one with updated values
                try:
                    self.lang_menu.pack_forget()
                    self.lang_menu.destroy()
                except Exception:
                    # ignore errors tearing down old widget
                    pass
                # Recreate using stable parent reference
                self.lang_menu = ctk.CTkOptionMenu(self.lang_frame, values=vals, variable=self.lang_var)
                self.lang_menu.pack(fill=tk.X, padx=4)
            else:
                menu = self.lang_menu["menu"]
                menu.delete(0, "end")
                for v in vals:
                    menu.add_command(label=v, command=lambda value=v: self.lang_var.set(value))
            self.lang_var.set(vals[0])
        except Exception as e:
            logging.exception("Failed to update language menu: %s", e)

    def _progress_cb(self, filepath: str, status: str):
        name = os.path.basename(filepath)
        # find index
        for i, p in enumerate(self.files):
            if p == filepath:
                display = f"{name} - {status}"
                self.listbox.delete(i)
                self.listbox.insert(i, display)
                break
        if status in ("finished", "error"):
            self.completed_files += 1
            try:
                if ctk and isinstance(self.progress, ctk.CTkProgressBar):
                    self.progress.set(self.completed_files / max(1, self.total_files))
                else:
                    self.progress['value'] = self.completed_files
            except Exception:
                pass
            if self.completed_files >= self.total_files:
                self.start_btn.configure(state=tk.NORMAL)
                self.stop_btn.configure(state=tk.DISABLED)

    def _log(self, msg: str):
        def append():
            self.log.insert(tk.END, msg + "\n")
            self.log.see(tk.END)
        self.root.after(0, append)

    def _update_listbox_all_status(self, status: str):
        self.listbox.delete(0, tk.END)
        for f in self.files:
            name = os.path.basename(f)
            self.listbox.insert(tk.END, f"{name} - {status}")

    def _set_default_output(self):
        val = self.output_entry.get().strip()
        if not val:
            messagebox.showwarning("No output", "Please choose an output folder to set as default.")
            return
        self.cfg['default_output'] = val
        save_config(self.cfg)
        messagebox.showinfo("Saved", "Default output saved.")

    def _set_default_input(self):
        # set default input to first file's parent or let user pick
        if self.files:
            parent = str(Path(self.files[0]).parent)
        else:
            init = self._initialdir_for_input()
            d = filedialog.askdirectory(title="Select default input folder", initialdir=init)
            if not d:
                return
            parent = d
        self.cfg['default_input'] = parent
        save_config(self.cfg)
        messagebox.showinfo("Saved", "Default input saved.")


def main():
    if ctk is None:
        tk.Tk().withdraw()
        messagebox.showerror("Missing dependency", "customtkinter is required for GUI. Install with:\n\n    pip install customtkinter")
        return

    root = ctk.CTk()
    app = App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
