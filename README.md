# Verbilo

Translate Word documents, Excel spreadsheets, and PDFs into any of 130+ languages while preserving the original formatting — tables, fonts, images, colours, and layout all stay intact.

---

## Features

- **Batch translation** — drop files in `origin/`, run one command, get translated files in `output/`
- **GUI** — interactive interface with file picker, progress tracking, and log output
- **Formatting preserved** — DOCX run-level styles (bold, italic, font, size) and XLSX cell styles survive translation unchanged
- **PDFs edited in-place** — text is replaced inside the original PDF using PyMuPDF, so images, graphics, and page layout are kept
- **Scanned PDF detection** — PDFs that require OCR are automatically detected and skipped with a log message instead of producing broken output
- **Source language filter** — specify a source language (e.g. English) so only text in that language is translated; columns or cells already in other languages are left untouched
- **130+ languages** — full Google Translate language list with a searchable dropdown in the GUI
- **Fast** — all text is collected and sent in batches rather than one request per word/cell, cutting translation time from ~10 minutes to under a minute for most files

## UI / Visual updates

The GUI has been refreshed for a more consistent, modern appearance.
Icons are now loaded centrally via the open-source Tabler Icons wrapper `pytablericons`; install it (and Pillow) to render icons in the app:

```bash
pip install pytablericons Pillow
```

Credits: Tabler Icons / pytablericons — https://pypi.org/project/pytablericons/

---

## Requirements

- Python 3.12+
- Dependencies:

```bash
pip install -r requirements.txt
```

For the GUI, also install:

```bash
pip install customtkinter
```

---

## Supported File Types

| Format | Extension | Notes |
|---|---|---|
| Word document | `.docx` | Run-level formatting preserved |
| Excel spreadsheet | `.xlsx`, `.xls` | Cell styles, merged cells, formulas preserved |
| PDF | `.pdf` | Text replaced in-place; images/graphics untouched. Scanned/image-only PDFs are skipped. |

---

## Quick Start

### 1. Install dependencies

```bash
pip install -r requirements.txt
pip install customtkinter   # for the GUI
```

### 2. Launch the GUI

```bash
python -m src.cli --gui
```

In the GUI:
1. Click **Add Files** or **Select Folder** to load documents
2. Set **Source language** — pick the language your documents are written in, or leave on *Auto-detect* to translate everything
3. Set **Target language** — the language to translate into (type to filter the list)
4. Choose an **Output folder** (defaults to `output/`)
5. Click **Start**

### 3. Or use the CLI

Translate all supported files in `origin/` to Spanish:

```bash
python -m src.cli es
```

Translate only English text to Portuguese (leaves other languages untouched):

```bash
python -m src.cli pt --source en
```

---

## CLI Reference

```
python -m src.cli [LANG] [--source CODE] [--gui]
```

| Argument | Description |
|---|---|
| `LANG` | Target language code, e.g. `es`, `pt`, `fr` |
| `--source CODE` | Source language code (default: `auto` — translate all text) |
| `--gui` | Launch the graphical interface instead of batch mode |

---

## Project Structure

```
origin/          ← place input files here
output/          ← translated files are written here
src/
  cli.py           — command-line entry point
  main.py          — translate_file() core API
  gui_customtk.py  — CustomTkinter GUI
  converters/
    docx_converter.py
    xlsx_converter.py
    pdf_converter.py
  translators/
    dummy.py       — Google Translate wrapper with batching + source filtering
```

---

## Language Codes

Standard ISO 639-1 codes are used. Common examples:

| Code | Language | Code | Language |
|---|---|---|---|
| `en` | English | `ar` | Arabic |
| `es` | Spanish | `hi` | Hindi |
| `fr` | French | `ja` | Japanese |
| `de` | German | `ko` | Korean |
| `it` | Italian | `ru` | Russian |
| `pt` | Portuguese | `zh-CN` | Chinese (Simplified) |
| `nl` | Dutch | `tr` | Turkish |
| `pl` | Polish | `sv` | Swedish |

The full list of 130+ supported codes is shown in the GUI dropdown.

---

## Notes

- Translation uses the **free Google Translate API** via `deep-translator`. No API key is required, but very large batches may occasionally hit rate limits.
- **Scanned PDFs** (image-only, no embedded text) are automatically detected and skipped — a message is logged so you know which files were not translated.
- The **source language filter** uses local language detection (`langdetect`) — no extra API calls.

---

## Future Implementations

- **Settings:** Auto Update (check for updates, install, etc)

## License
MIT
