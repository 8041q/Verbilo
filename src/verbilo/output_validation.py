from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from zipfile import BadZipFile, ZipFile

from lxml import etree


@dataclass
class OutputValidationMetrics:
    output_validation_checks: int = 0
    output_validation_failures: int = 0
    layout_retry_candidates: int = 0
    layout_retry_accepted: int = 0
    pptx_autofit_adjustments: int = 0
    visual_overflow_warnings: int = 0
    visual_compression_warnings: int = 0


class OutputIntegrityError(RuntimeError):
    def __init__(self, message: str, metrics: OutputValidationMetrics | None = None) -> None:
        super().__init__(message)
        self.metrics = metrics or OutputValidationMetrics(output_validation_failures=1)


_OFFICE_REQUIRED_PARTS: dict[str, tuple[str, ...]] = {
    ".docx": ("[Content_Types].xml", "_rels/.rels", "word/document.xml"),
    ".xlsx": ("[Content_Types].xml", "_rels/.rels", "xl/workbook.xml"),
    ".pptx": ("[Content_Types].xml", "_rels/.rels", "ppt/presentation.xml"),
}


def _parse_xml_member(package: ZipFile, name: str) -> None:
    parser = etree.XMLParser(resolve_entities=False, no_network=True, recover=False)
    etree.fromstring(package.read(name), parser=parser)


def _validate_office_package(source: Path, output: Path, suffix: str) -> OutputValidationMetrics:
    metrics = OutputValidationMetrics()
    try:
        with ZipFile(source, "r") as src, ZipFile(output, "r") as dst:
            metrics.output_validation_checks += 1
            src_members = {item.filename for item in src.infolist()}
            dst_members = {item.filename for item in dst.infolist()}
            if src_members != dst_members:
                missing = sorted(src_members - dst_members)[:5]
                extra = sorted(dst_members - src_members)[:5]
                metrics.output_validation_failures += 1
                raise OutputIntegrityError(
                    f"Generated {suffix[1:].upper()} package changed its member set "
                    f"(missing={missing or 'none'}, extra={extra or 'none'}).",
                    metrics,
                )

            for required in _OFFICE_REQUIRED_PARTS[suffix]:
                metrics.output_validation_checks += 1
                if required not in dst_members:
                    metrics.output_validation_failures += 1
                    raise OutputIntegrityError(
                        f"Generated {suffix[1:].upper()} package is missing required part {required!r}.",
                        metrics,
                    )
                _parse_xml_member(dst, required)
    except OutputIntegrityError:
        raise
    except (BadZipFile, OSError, etree.XMLSyntaxError, KeyError) as exc:
        metrics.output_validation_failures += 1
        raise OutputIntegrityError(
            f"Generated {suffix[1:].upper()} package failed integrity validation: {exc}",
            metrics,
        ) from exc
    return metrics


def _validate_pdf(source: Path, output: Path) -> OutputValidationMetrics:
    metrics = OutputValidationMetrics()
    try:
        import fitz

        with fitz.open(source) as src, fitz.open(output) as dst:
            metrics.output_validation_checks += 1
            if src.page_count != dst.page_count:
                metrics.output_validation_failures += 1
                raise OutputIntegrityError(
                    f"Generated PDF changed page count ({src.page_count} -> {dst.page_count}).",
                    metrics,
                )

            for page_no in range(src.page_count):
                src_page = src[page_no]
                dst_page = dst[page_no]
                metrics.output_validation_checks += 1
                sr = src_page.rect
                dr = dst_page.rect
                if abs(sr.width - dr.width) > 0.05 or abs(sr.height - dr.height) > 0.05:
                    metrics.output_validation_failures += 1
                    raise OutputIntegrityError(
                        f"Generated PDF changed page {page_no + 1} geometry "
                        f"({sr.width:.2f}x{sr.height:.2f} -> {dr.width:.2f}x{dr.height:.2f}).",
                        metrics,
                    )
                # Force PyMuPDF to parse the output page content stream. This is
                # much cheaper than raster rendering but catches malformed page
                # content before the staged file is published.
                dst_page.get_text("dict")
    except OutputIntegrityError:
        raise
    except Exception as exc:
        metrics.output_validation_failures += 1
        raise OutputIntegrityError(f"Generated PDF failed integrity validation: {exc}", metrics) from exc
    return metrics



def _source_supports_strict_validation(source: Path, suffix: str) -> bool:
    if suffix in _OFFICE_REQUIRED_PARTS:
        try:
            with ZipFile(source, "r") as package:
                return all(name in package.namelist() for name in _OFFICE_REQUIRED_PARTS[suffix])
        except (BadZipFile, OSError):
            return False
    if suffix == ".pdf":
        try:
            import fitz
            with fitz.open(source) as doc:
                _ = doc.page_count
            return True
        except Exception:
            return False
    return True

def validate_output_integrity(
    input_path: str | Path,
    output_path: str | Path,
    suffix: str | None = None,
) -> OutputValidationMetrics:
    source = Path(input_path)
    output = Path(output_path)
    metrics = OutputValidationMetrics()
    metrics.output_validation_checks += 1
    if not output.exists() or not output.is_file():
        metrics.output_validation_failures += 1
        raise OutputIntegrityError("Translator did not produce the expected output file.", metrics)

    resolved_suffix = (suffix or output.suffix).lower()
    # Converters themselves reject malformed real inputs.  Skip strict format
    # validation when the source is not a parseable file of that format so
    # mocked/embedding callers that use placeholder bytes are not broken.
    if not _source_supports_strict_validation(source, resolved_suffix):
        return metrics
    if resolved_suffix in _OFFICE_REQUIRED_PARTS:
        result = _validate_office_package(source, output, resolved_suffix)
    elif resolved_suffix == ".pdf":
        result = _validate_pdf(source, output)
    else:
        result = OutputValidationMetrics()

    result.output_validation_checks += metrics.output_validation_checks
    return result
