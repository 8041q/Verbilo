from pathlib import Path
from .translators.dummy import TranslatorFactory
from .converters import docx_converter, xlsx_converter, pdf_converter
from .utils.io import resolve_output_path


def translate_file(input_path: str, target_lang: str, output_path: str | None = None, translator_name: str | None = None):
    p = Path(input_path)
    if not p.exists():
        raise FileNotFoundError(input_path)

    translator = TranslatorFactory.get(translator_name)

    # Validate target language early to avoid silent no-ops downstream
    if not target_lang or not isinstance(target_lang, str) or not target_lang.strip():
        raise ValueError("target_lang must be a non-empty language code (e.g. 'en', 'pt')")

    suffix = p.suffix.lower()

    # Resolve output path after any conversion so the suffix/filename is correct
    output_path = resolve_output_path(p, output_path)

    if suffix == ".docx":
        docx_converter.translate_docx(str(p), str(output_path), translator, target_lang)
    elif suffix in (".xls", ".xlsx"):
        xlsx_converter.translate_xlsx(str(p), str(output_path), translator, target_lang)
    elif suffix == ".pdf":
        pdf_converter.translate_pdf(str(p), str(output_path), translator, target_lang)
    else:
        raise ValueError(f"Unsupported file type: {suffix}")
    # return the final output path for callers that want to log or inspect it
    return output_path


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("input")
    parser.add_argument("--to", "-t", required=True, help="target language code (e.g., 'pt')")
    parser.add_argument("--out", "-o", default=None, help="output path")
    parser.add_argument("--translator", default=None, help="translator backend (default: auto)")
    args = parser.parse_args()
    translate_file(args.input, args.to, args.out, args.translator)
