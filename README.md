Translator — Simple document translator
=====================================

A small, easy-to-use document translation helper. Convert and translate common document files from the `origin/` folder and write results to `output/`.

What it does
-----------
- Converts and translates documents using the project's converters and translators.
- Keeps the interface minimal so you can integrate the tool into pipelines or call it manually.

Requirements
-----------
- Python 3.8+
- Install dependencies with:

```bash
pip install -r requirements.txt
```

Quick usage (examples)
----------------------
The CLI supports two modes: batch translation (positional language code) or launching the GUI.

Batch (translate all supported files found in `origin/` to a target language):

```bash
# Example: translate all supported files in origin/ to Spanish
python -m src.doc_translator.cli es
```

Launch the GUI (interactive):

```bash
# Optional GUI dependency
pip install customtkinter

# Launch GUI via the CLI flag
python -m src.doc_translator.cli --gui

# Or run the GUI module directly
python -m src.doc_translator.gui_customtk
```

Notes:
- The CLI will create `origin/` and `output/` if they don't exist and process all supported files it finds.
- Supported input extensions: `.docx`, `.pdf`, `.xlsx`, `.xls`.

Supported input formats
----------------------
- .docx (Microsoft Word)
- .pdf (PDF documents)
- .xlsx (Excel spreadsheets)

Supported output
----------------
- Same formats converted and written to `output/` (format depends on converter used).

Available language codes (examples)
---------------------------------
The project accepts standard two-letter language codes. Common supported codes include:

- `en` — English
- `es` — Spanish
- `fr` — French
- `de` — German
- `it` — Italian
- `pt` — Portuguese
- `ru` — Russian
- `zh` — Chinese (Simplified)
- `ja` — Japanese
- `ko` — Korean


GUI
---
- A simple GUI is available under `src/doc_translator/` (see `gui_customtk.py`, `gui_config.py`, and `gui_helpers.py`).
- To run the GUI (example):

```bash
# Optional: install CustomTkinter for the themed UI
pip install customtkinter

# Run the GUI module (example)
python -m src.doc_translator.gui_customtk
```

- The GUI lets you pick input files, choose a target language, and run conversions without the CLI.

License
-----------------------
MIT
