from pathlib import Path
import threading
from .translators.factory import TranslatorFactory
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
):
    # source_lang="auto" translates everything; cancel_event raises CancelledError before saving
    p = Path(input_path)
    if not p.exists():
        raise FileNotFoundError(input_path)

    translator = TranslatorFactory.get(
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

    # Validate target language early to avoid silent no-ops downstream
    if not target_lang or not isinstance(target_lang, str) or not target_lang.strip():
        raise ValueError("target_lang must be a non-empty language code (e.g. 'en', 'pt')")

    suffix = p.suffix.lower()

    # Resolve output path after any conversion so the suffix/filename is correct
    output_path = resolve_output_path(p, output_path)

    if suffix == ".docx":
        docx_converter.translate_docx(str(p), str(output_path), translator, target_lang, cancel_event=cancel_event)
    elif suffix in (".xls", ".xlsx"):
        xlsx_converter.translate_xlsx(str(p), str(output_path), translator, target_lang, cancel_event=cancel_event)
    elif suffix == ".pdf":
        result = pdf_converter.translate_pdf(str(p), str(output_path), translator, target_lang, cancel_event=cancel_event)
        if result == "skipped-ocr":
            return "skipped-ocr"
    else:
        raise ValueError(f"Unsupported file type: {suffix}")
    return output_path


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
