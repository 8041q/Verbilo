# Verbilo - Portable
<sub>- beta</sub>

Translate Excel spreadsheets, Word documents, and PDFs into any of 130+ languages while preserving the original formatting, tables, fonts, images, colours, and layout.

---

## Future Implementations

- Change UI/UX framework (Electron, Tauri, Neutralino.js, Wails or Flutter)
- Create Web version
- Add AI translating capabilities and/or detection (via API) 

## Features

- **Multi-Engine Language Detection**: Use Lingua, FastText, or LangDetect to filter what gets translated. Use Auto mode for a majority-vote system between all three for maximum accuracy
- **GUI**: interactive interface with file picker, progress tracking, and log output
- **Formatting preserved**: DOCX run-level styles (bold, italic, font, size) and XLSX cell styles survive translation
- **PDFs edited in-place**: text is replaced inside the original PDF using PyMuPDF, so images, graphics, and page layout are kept
- **Scanned PDF detection**: PDFs that require OCR are automatically detected and skipped with a log message instead of producing broken output
- **Source language filter**: specify a source language so only text in that language is translated
- **130+ languages**: full Google Translate language list with a searchable dropdown
- **Segment-Aware Translation**: Smart splitting logic detects languages within sub-segments of cells (separated by / or \n), ensuring only the intended text is translated while keeping technical strings or other languages untouched

## UI / UX

The GUI are loaded centrally via the open-source Tabler Icons wrapper `pytablericons`; install it (and Pillow) to render icons in the app:

```bash
pip install pytablericons Pillow
```

Credits: Tabler Icons / pytablericons: https://pypi.org/project/pytablericons/

---

## Requirements

- Python 3.12+
- Dependencies:

```bash
pip install -r requirements.txt
```

For the GUI, also install (comes included):

```bash
pip install customtkinter
```

---

## How it Works

Verbilo handles translation differently based on your Source Language setting:

Source = Auto-detect: Every text segment is sent to Google Translate unconditionally.

Source = Specific Language (e.g., English):

The app splits text into segments (using / and \n as delimiters).

Each segment is checked by your chosen Detector.

Only segments matching the source language are translated.

Auto Detector Mode: All three engines (Lingua, FastText, LangDetect) vote. If 2+ agree it's the source language, it's translated.

---

## Quick Start

### 1. Install dependencies

```bash
pip install -r requirements.txt
pip install customtkinter   # if it says not installed, otherwise ignore this command
```

### 2. Launch the GUI

```bash
python -m src.verbilo.cli --gui
```

In the GUI:
1. Click **Add Files** or **Select Folder** to load documents
2. Set **Source language**: pick the language to translate only that or leave on *Auto-detect* to translate everything
3. Set **Target language**: pick the destination language to translate into (type to filter the list)
4. **Language detector**: select your preferred engine (use auto for best results)
5. Choose an **Output folder**: (defaults to `output/`, creates the folder on root)
6. Click **Start**

### 3. Or use the CLI

Translate all supported files in `origin/` to Spanish:

```bash
python -m src.verbilo.cli es
```

Translate only English text to Portuguese (leaves other languages untouched):

```bash
python -m src.verbilo.cli pt --source en
```

---

## CLI Reference

```
python -m src.verbilo.cli [LANG] [--source CODE] [--gui]
```

| Argument | Description |
|---|---|
| `LANG` | Target language code (e.g.,`es`, `pt`) |
| `--source CODE` | Source language code (default: `auto`) |
| `--detector` | Detection engine: auto (default), lingua, fasttext, or langdetect |
| `--gui` | Launch the graphical interface |

---

## Project Structure

```
output/          ← translated files, auto creates if no folder selected
src/
  verbilo/
    cli.py
    main.py         - `translate_file()` core API
    gui/
      app.py        - CustomTkinter GUI
      config.py
      helpers.py
      theme.py
      icons.py
    converters/
      docx_converter.py
      xlsx_converter.py
      pdf_converter.py
    translators/
      base.py
      dummy.py
      google.py
      lang_detect.py
```

---


## Notes

- **API Usage**: To avoid rate limits, Verbilo batches segments. When a specific source language is set, detection happens locally, saving API overhead on non-target text
- **Scanned PDFs** (image-only, no embedded text) are automatically detected and skipped: a message is logged so you know which files were not translated
- **Detection Performance**: Lingua is highly accurate for short strings but heavier; FastText is extremely fast. The auto (vote system) mode provides the best balance


## License
MIT
