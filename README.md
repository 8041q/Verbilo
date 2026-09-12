<p align="center">
  <img src="./src/verbilo/assets/favicon.jpg" alt="Verbilo logo" width="220" />
</p>

<div align="center">

[![Releases](https://img.shields.io/github/v/release/8041q/Verbilo)](https://github.com/8041q/Verbilo/releases)
[![Release Notes](https://img.shields.io/badge/release-notes-blue)](https://github.com/8041q/Verbilo/releases)
[![Python](https://img.shields.io/badge/python-3.12.x-brightgreen)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/License-AGPL)](https://www.gnu.org/licenses/agpl-3.0.en.html)
[![Stars](https://img.shields.io/github/stars/8041q/Verbilo?style=flat)](https://github.com/8041q/Verbilo/stargazers)
[![Issues](https://img.shields.io/github/issues/8041q/Verbilo)](https://github.com/8041q/Verbilo/issues)

</div>

### Verbilo - Portable

<p align="center"><em>Translate DOCX, XLSX and PDF documents while preserving layout, styles, formatting, and images.</em></p>

---
---

## How it works

Verbilo extracts document content into translation units, sends those units to the selected translation backend, and reconstructs the translated document while preserving the original structure as closely as possible.

* Converts document content into translation units such as runs, spans, rows, and paragraphs.
* Groups translation units into batches to reduce API calls and improve throughput.
* Uses resilient HTTP handling with retries, backoff, timeouts, and sub-batch fallbacks for large requests.
* Uses API-aware inline tagging where supported to preserve run and span formatting across the translation round trip.
* Reconstructs translated text into the original document structure while retaining formatting where feasible.
* Applies layout-aware handling for PDFs, including optional semantic/local-LLM routing for text that may not fit safely inside its original bounding box.

## Feature Highlights

* **Multiple translation engines** - Google Translate (free), Google Cloud Translation API, Baidu, Microsoft Azure Translator, DeepL, Local OPUS-MT, and Ollama-based local models.
* **Broad language coverage** - Google-based translation backends support a large range of languages, including 130+ languages through Google Cloud Translation. Exact language support depends on the selected engine.
* **Proxy & resilience** - Network translation engines use resilient HTTP sessions with retries, backoff, timeouts, and optional HTTPS/HTTP proxy support.
* **Selective Translation** - Translate only text detected as a specified source language, or use automatic source-language handling.
* **Formatting Preservation** - Preserves DOCX run-level formatting, XLSX cell structure, and in-place PDF layout where possible.
* **Multi-Engine Detection** - Choose between Lingua and FastText depending on your preferred accuracy/speed trade-off.
* **Batching for Efficiency** - Translation segments are grouped to reduce API calls and help avoid rate limits.
* **Semantic PDF Translation** - Optional local-LLM assistance can classify and route layout-constrained PDF blocks before translation.
* **Local Translation** - OPUS-MT and Ollama models can run locally without sending document text to a remote translation API.
* **GUI Language** - Interface available in English and Simplified Chinese (`ZH-Hans`), selectable in Settings.

## Known Limitations

* Inline-tag survival is API-dependent; formatting tags are not guaranteed to survive every translation backend.
* The PDF Z-order guard is conservative. Text entirely covered by opaque graphics can be skipped rather than rewriting PDF content streams to change stacking order.
* Very short CJK tokens (1–2 characters) can behave inconsistently across translation APIs. When possible, specify the source language explicitly.
* Extremely complex layouts, including heavy overlays, rotated text, or unusual encodings, can still produce visual artifacts. Manual verification is recommended for critical documents.
* Chinese → English translation can be particularly challenging in layout-constrained documents because translated English text frequently occupies more horizontal space.
* Image-only/scanned PDFs are not OCR-translated. They are detected and skipped rather than producing broken output.
* Language availability varies by translation engine and local model.


## Quick Start - For Developers

### 1. Requirements

Verbilo currently supports:

* **Python 3.12.x**
* Windows is the primary target for the portable GUI/Nuitka build.

A virtual environment is recommended.

### 2. Install dependencies

From the repository root:

```bash
pip install -r requirements.txt
```

### 3. Launch the GUI

```bash
cd src
python -m launch
```


## Translation Engines & Network Settings

The translation engine can be selected from the GUI sidebar.

### Google Translate

* Free.
* Default translation engine.
* No API key required.
* Broad language coverage.

### Google Cloud Translation API

Supports Google Cloud Translation API configurations.

* **v2** - API key.
* **v3** - Project ID and account/service credentials.

Google Cloud Translation supports 130+ languages, although available language combinations and capabilities can depend on the API/model being used.

### Baidu Translate

Requires:

* Baidu App ID
* Baidu App Key

### Microsoft Azure Translator

Requires:

* Subscription Key
* Region

### DeepL

Requires a DeepL API key:

* DeepL API Free
* DeepL API Pro

### Local - OPUS-MT

* Offline.
* No translation API key required.
* Free and unlimited local use.
* Requires a separate downloaded model for each supported source → target language pair.

### Ollama / Local LLM

* Runs locally.
* No translation API key required.
* Free and unlimited local use after the required models are installed.
* Used directly for DOCX/XLSX translation and for layout-aware/semantic PDF translation workflows.

Configure network settings, API credentials, proxy options, and related settings under:

**Settings → Network & API keys**

If an API-based translation method is selected without the required credentials, the GUI displays a warning instead of starting the translation job.


## Local AI Translation via Ollama

Verbilo can use local Ollama models for DOCX, XLSX, and PDF translation.

The role of the local model differs depending on the document type and selected model.

### DOCX / XLSX

When Ollama is selected as the translation engine, the selected Ollama model handles the document text directly.

This does **not** mean Ollama is required for DOCX or XLSX. Other translation engines such as Google, DeepL, Azure, Baidu, and Local OPUS-MT can also be used.

### PDF

For PDF documents, local models can work alongside the selected primary translation engine.

Translated text can be substantially longer than the source text. This is particularly noticeable for languages such as Chinese → English, where translated text may require significantly more physical space.

Verbilo can use the dimensions and available character budget of each PDF text block to decide how translation should be handled.

|             | **HY-MT 1.5 1.8B · Tencent** | **Qwen3.5 4B · Alibaba**  |
| ----------- | ---------------------------- | ------------------------- |
| Size        | ~1.1 GB                      | ~2.4 GB                   |
| Languages   | 33 major languages           | Broad coverage            |
| Speed       | Faster / smaller             | Slower / larger           |
| PDF routing | All routed blocks            | Layout-constrained blocks |
| DOCX / XLSX | ✓                            | ✓                         |

### HY-MT

HY-MT is a purpose-built machine-translation model.

For PDF workflows, HY-MT can translate routed blocks directly.

For DOCX and XLSX, it can operate as the selected primary local translator.

### Qwen

Qwen is a general-purpose local LLM.

For semantic PDF translation, Verbilo can use a smaller companion/advisor stage to classify blocks before translation. This allows layout-constrained content to be routed to the local model while ordinary blocks can continue through the selected primary translation engine.

This reduces unnecessary local-LLM calls while allowing difficult blocks to receive layout-aware translation.

For DOCX and XLSX, Qwen can also operate as the selected local translation engine.

### Ollama installation

Ollama is **not installed automatically during normal Verbilo setup**.

If Ollama is already installed, Verbilo can use it directly.

If Ollama is not installed and the user enables the Semantic Translation workflow and attempts to download a supported local model, Verbilo can download/install Ollama automatically as part of that model-download process.

Downloaded Ollama models are stored outside the Verbilo application directory, for example on Windows:

```text
%USERPROFILE%\.ollama\models\
```


## Language Detection

Verbilo supports multiple local language-detection engines.

### Lingua

* High accuracy.
* Particularly useful for short strings.
* Heavier than FastText.

### FastText

* Very fast.
* Good accuracy/performance balance.
* Requires the FastText language identification model.

To download the required FastText model, run:

```bash
python download_models.py
```

This downloads:

```text
models/lid.176.bin
```

When a specific source language is selected, local language detection can prevent text in other languages from being sent unnecessarily to translation APIs.


## Project Structure

<details>
<summary>Click to expand</summary>

```text
src/
  Origin/
  Output/
  verbilo/
    launch.py
    cli.py
    main.py              - translate_file() core API

    gui/
      app.py             - CustomTkinter GUI
      config.py
      helpers.py
      theme.py
      icons.py

    converters/
      docx_converter.py
      xlsx_converter.py
      pdf_converter.py

    translators/
      azure.py
      baidu.py
      base.py
      cache.py
      deepl.py
      factory.py
      google.py
      http_session.py
      lang_detect.py
      local.py
      usage.py

    assets/
      __init__.py

    utils/
      io.py

pyproject.toml
requirements.txt
README.md
```

</details>

> The project structure above highlights the main modules rather than every internal implementation file.


## Nuitka Build - Windows

### Prerequisites

* Python 3.12.x
* An activated virtual environment is recommended.
* Nuitka installed in the environment.

Install Nuitka if necessary:

```bash
pip install nuitka
```

### Build the GUI executable

From the repository root, with the virtual environment active:

```bash
.venv\Scripts\python.exe scripts\build_nuitka.py --entry gui --output dist/nuitka
```

### Notes

* For a final GUI build without a console window, use the appropriate `--windows-console-mode=disable` behavior exposed by the build helper.
* If a build fails after configuration or dependency changes, clearing the Nuitka cache can help:

```text
%LOCALAPPDATA%\Nuitka\Nuitka\
```

After a successful build, launch the GUI by running `verbilo.exe`.

### Build troubleshooting

FastText language detection requires:

```text
models/lid.176.bin
```

Run:

```bash
python download_models.py
```

For Local OPUS-MT translation, the corresponding source → target translation model must also be downloaded.

Each supported source/target language pair uses its own model.

If paths or build behavior differ, verify that:

* Commands are being executed from the repository root.
* The intended virtual environment is active.
* Nuitka is installed in that environment.
* Runtime dependencies are synchronized between `requirements.txt` and `pyproject.toml`.


## Requirements & Notes

### Python

```text
>=3.12,<3.13
```

In other words, use **Python 3.12.x**.

### Install runtime dependencies

```bash
pip install -r requirements.txt
```

### Important behavior

* Scanned/image-only PDFs are detected and skipped rather than producing malformed translated output.
* Selecting an explicit source language allows local detection to avoid unnecessary translation requests.
* Translation units are batched to reduce API usage and improve throughput.
* Language support depends on the selected translation backend.
* Local models must be downloaded before they can be used.
* API-based engines require their corresponding credentials where applicable.


## Development Dependencies

Runtime dependencies should be kept synchronized between:

```text
requirements.txt
pyproject.toml
poetry.lock
```

When adding or changing a runtime dependency, update the Poetry dependency configuration as well.

For example, `tkinterdnd2` is currently:

```text
tkinterdnd2==0.6.3
```

in `requirements.txt`, so the corresponding Poetry dependency should be:

```toml
tkinterdnd2 = "==0.6.3"
```

## Contributing

Contributions are welcome.

* Found a bug? Open an issue with steps to reproduce it and, if possible, include a sample file.
* Want to add a feature? Fork the repository, create a focused feature branch, and open a pull request.
* Keep changes focused and minimal.
* Follow the existing code style.
* Run the relevant tests locally before submitting a pull request.
* Document user-visible changes, even when they are small.


## Acknowledgments

* [PyMuPDF](https://pymupdf.readthedocs.io/) for PDF processing and in-place PDF text editing.
* [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) for the GUI framework.
* Tabler Icons / `pytablericons` for GUI icons.
* Lingua and FastText for language detection.
* CTranslate2 and SentencePiece for local translation support.
* Ollama and supported local models for local LLM translation.
* Everyone who files issues, reports bugs, and contributes patches.


## License

This project is released under the **GNU Affero General Public License v3 (AGPL-3.0-or-later)**.

See the `LICENSE` file for details.
