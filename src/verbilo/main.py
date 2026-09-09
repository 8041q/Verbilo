import os
from pathlib import Path
import threading
from uuid import uuid4
from typing import Callable
from .advisors import AdvisorBase, NullAdvisor
from .translators.factory import TranslatorFactory
from .translators.base import Translator
from .converters import docx_converter, xlsx_converter, pdf_converter
from .utils.io import resolve_output_path
from .utils import CancelledError

__all__ = ["translate_file", "CancelledError"]

def translate_file(
    input_path: str,
    target_lang: str,
    output_path: str | None = None,
    translator_name: str | None = None,
    source_lang: str = "auto",
    cancel_event: threading.Event | None = None,
    detector: str = "fasttext",
    engine: str = "google",
    proxies: dict | None = None,
    google_api_key: str = "",
    baidu_appid: str = "",
    baidu_appkey: str = "",
    azure_key: str = "",
    azure_region: str = "",
    deepl_api_key: str = "",
    baidu_tier: str = "standard",
    google_project_id: str = "",
    google_sa_json: str = "",
    local_model_dir: str = "",
    progress_callback: Callable[[int, int], None] | None = None,
    advisor: AdvisorBase | None = None,
    semantic_translator: Translator | None = None,
    translator_override: Translator | None = None,
    *,
    overwrite: bool = False,
):
    # source_lang="auto" translates everything. Results are staged and committed
    # atomically, so a cancellation or error never publishes a partial document.
    p = Path(input_path)
    if not p.exists():
        raise FileNotFoundError(input_path)

    if not target_lang or not isinstance(target_lang, str) or not target_lang.strip():
        raise ValueError("target_lang must be a non-empty language code (e.g. 'en', 'pt')")

    suffix = p.suffix.lower()
    if suffix not in (".docx", ".xlsx", ".pdf"):
        if suffix == ".xls":
            raise ValueError("Legacy .xls files are not supported; convert the workbook to .xlsx first.")
        raise ValueError(f"Unsupported file type: {suffix}")

    final_output = Path(resolve_output_path(p, output_path)).resolve()
    if final_output == p.resolve():
        raise ValueError("Output path must not overwrite the source document.")
    if final_output.exists() and not overwrite:
        raise FileExistsError(
            f"Output already exists: {final_output}. Choose another output path or pass overwrite=True."
        )
    final_output.parent.mkdir(parents=True, exist_ok=True)
    staged_output = final_output.with_name(
        f".{final_output.stem}.{uuid4().hex}.staging{final_output.suffix}"
    )

    translator = translator_override or TranslatorFactory.get(
        translator_name,
        source_lang=source_lang or "auto",
        detector=detector,
        engine=engine,
        proxies=proxies,
        google_api_key=google_api_key,
        baidu_appid=baidu_appid,
        baidu_appkey=baidu_appkey,
        azure_key=azure_key,
        azure_region=azure_region,
        deepl_api_key=deepl_api_key,
        baidu_tier=baidu_tier,
        google_project_id=google_project_id,
        google_sa_json=google_sa_json,
        local_model_dir=local_model_dir,
    )

    if advisor is None:
        advisor = NullAdvisor()

    try:
        if suffix == ".docx":
            docx_converter.translate_docx(str(p), str(staged_output), translator, target_lang, cancel_event=cancel_event, source_lang=source_lang, progress_callback=progress_callback)
        elif suffix == ".xlsx":
            xlsx_converter.translate_xlsx(str(p), str(staged_output), translator, target_lang, cancel_event=cancel_event, source_lang=source_lang, progress_callback=progress_callback)
        else:
            result = pdf_converter.translate_pdf(
                str(p), str(staged_output), translator, target_lang,
                cancel_event=cancel_event, source_lang=source_lang,
                progress_callback=progress_callback, advisor=advisor,
                semantic_translator=semantic_translator,
            )
            if result == "skipped-ocr":
                return "skipped-ocr"
        if cancel_event is not None and cancel_event.is_set():
            raise CancelledError("Translation cancelled before publishing output")
        if overwrite:
            os.replace(staged_output, final_output)
        else:
            # Linking is an atomic no-clobber publish when staging and final
            # live in the same directory.  Unlike os.replace(), it cannot
            # overwrite an output created after the earlier existence check.
            try:
                os.link(staged_output, final_output)
            except FileExistsError as exc:
                raise FileExistsError(
                    f"Output was created while translating: {final_output}. "
                    "Choose another output path and try again."
                ) from exc
            staged_output.unlink()
        return str(final_output)
    finally:
        # A failed conversion, cancellation, or OCR skip must not leave an
        # incomplete destination visible to the user.
        try:
            staged_output.unlink(missing_ok=True)
        except OSError:
            pass


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("input")
    parser.add_argument("--to", "-t", required=True, help="target language code (e.g., 'pt')")
    parser.add_argument("--from", "--source", dest="source", default="auto",
                        help="source language code (e.g., 'en'). 'auto' = translate all text")
    parser.add_argument("--out", "-o", default=None, help="output path")
    parser.add_argument("--translator", default=None, help="translator backend (default: auto)")
    args = parser.parse_args()
    translate_file(args.input, args.to, args.out, args.translator, source_lang=args.source)
