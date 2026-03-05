# in-place PDF translator using PyMuPDF; skips scanned/OCR-only PDFs
# Preserves original font family, weight (bold), style (italic), size, and
# colour.  Uses a two-pass layout engine that shifts subsequent lines down
# when translated text is taller than the original, rather than shrinking
# the font to an unreadable size.

from __future__ import annotations

import logging
import re
import sys
import threading
from pathlib import Path
from typing import Any

import fitz  # PyMuPDF
from ..utils import CancelledError

logger = logging.getLogger(__name__)

_MIN_CHARS_PER_PAGE = 20  # min avg chars/page to treat as real text (not scanned)

# Minimum allowed scale factor when text truly cannot fit even after layout
# shifting (e.g. single-line field that cannot grow vertically).
_MIN_SCALE = 0.85

# Span flag bit-masks as per PDF spec / PyMuPDF docs
_FLAG_SUPERSCRIPT = 1
_FLAG_ITALIC = 2
_FLAG_SERIFED = 4
_FLAG_MONOSPACE = 8
_FLAG_BOLD = 16


def _is_ocr_required(doc: fitz.Document) -> bool:
    # True if the PDF looks scanned (too few extractable chars per page)
    if doc.page_count == 0:
        return True
    total_chars = 0
    for page in doc:
        text = page.get_text("text")
        if isinstance(text, str):
            total_chars += len(text.strip())
    avg = total_chars / doc.page_count
    return avg < _MIN_CHARS_PER_PAGE


# ---------------------------------------------------------------------------
# Platform font map: canonical token → {variant: filesystem path}
# ---------------------------------------------------------------------------

def _build_font_map() -> dict[str, dict[str, str]]:
    """Return a map of lowercase font-family tokens to TTF paths per variant.

    Variants: "regular", "bold", "italic", "bolditalic"
    """
    if sys.platform == "win32":
        import os
        windir = os.environ.get("WINDIR", r"C:\Windows")
        fd = Path(windir) / "Fonts"

        def w(*parts: str) -> str:
            return str(fd.joinpath(*parts))

        return {
            "arial": {
                "regular":    w("arial.ttf"),
                "bold":       w("arialbd.ttf"),
                "italic":     w("ariali.ttf"),
                "bolditalic": w("arialbi.ttf"),
            },
            "times": {
                "regular":    w("times.ttf"),
                "bold":       w("timesbd.ttf"),
                "italic":     w("timesi.ttf"),
                "bolditalic": w("timesbi.ttf"),
            },
            "timesnewroman": {
                "regular":    w("times.ttf"),
                "bold":       w("timesbd.ttf"),
                "italic":     w("timesi.ttf"),
                "bolditalic": w("timesbi.ttf"),
            },
            "courier": {
                "regular":    w("cour.ttf"),
                "bold":       w("courbd.ttf"),
                "italic":     w("couri.ttf"),
                "bolditalic": w("courbi.ttf"),
            },
            "georgia": {
                "regular":    w("georgia.ttf"),
                "bold":       w("georgiab.ttf"),
                "italic":     w("georgiai.ttf"),
                "bolditalic": w("georgiaz.ttf"),
            },
            "verdana": {
                "regular":    w("verdana.ttf"),
                "bold":       w("verdanab.ttf"),
                "italic":     w("verdanai.ttf"),
                "bolditalic": w("verdanaz.ttf"),
            },
            "tahoma": {
                "regular":    w("tahoma.ttf"),
                "bold":       w("tahomabd.ttf"),
                "italic":     w("tahoma.ttf"),
                "bolditalic": w("tahomabd.ttf"),
            },
            "calibri": {
                "regular":    w("calibri.ttf"),
                "bold":       w("calibrib.ttf"),
                "italic":     w("calibrii.ttf"),
                "bolditalic": w("calibriz.ttf"),
            },
            "cambria": {
                "regular":    w("cambria.ttc"),
                "bold":       w("cambriab.ttf"),
                "italic":     w("cambriai.ttf"),
                "bolditalic": w("cambriaz.ttf"),
            },
            "segoeui": {
                "regular":    w("segoeui.ttf"),
                "bold":       w("segoeuib.ttf"),
                "italic":     w("segoeuii.ttf"),
                "bolditalic": w("segoeuiz.ttf"),
            },
        }
    elif sys.platform == "darwin":
        candidates: list[tuple[str, str, str, str, str]] = [
            # token, regular, bold, italic, bolditalic
            ("arial",
             "/Library/Fonts/Arial.ttf",
             "/Library/Fonts/Arial Bold.ttf",
             "/Library/Fonts/Arial Italic.ttf",
             "/Library/Fonts/Arial Bold Italic.ttf"),
            ("times",
             "/Library/Fonts/Times New Roman.ttf",
             "/Library/Fonts/Times New Roman Bold.ttf",
             "/Library/Fonts/Times New Roman Italic.ttf",
             "/Library/Fonts/Times New Roman Bold Italic.ttf"),
            ("georgia",
             "/Library/Fonts/Georgia.ttf",
             "/Library/Fonts/Georgia Bold.ttf",
             "/Library/Fonts/Georgia Italic.ttf",
             "/Library/Fonts/Georgia Bold Italic.ttf"),
        ]
        fm: dict[str, dict[str, str]] = {}
        for token, reg, bold, ita, bi in candidates:
            fm[token] = {"regular": reg, "bold": bold,
                         "italic": ita, "bolditalic": bi}
        return fm
    else:  # Linux / other
        def lp(name: str) -> str:
            for base in (
                "/usr/share/fonts/truetype/liberation",
                "/usr/share/fonts/truetype/dejavu",
                "/usr/share/fonts/truetype/noto",
                "/usr/share/fonts/TTF",
            ):
                p = Path(base) / name
                if p.is_file():
                    return str(p)
            return ""

        return {
            "liberation": {
                "regular":    lp("LiberationSans-Regular.ttf"),
                "bold":       lp("LiberationSans-Bold.ttf"),
                "italic":     lp("LiberationSans-Italic.ttf"),
                "bolditalic": lp("LiberationSans-BoldItalic.ttf"),
            },
            "dejavu": {
                "regular":    lp("DejaVuSans.ttf"),
                "bold":       lp("DejaVuSans-Bold.ttf"),
                "italic":     lp("DejaVuSans-Oblique.ttf"),
                "bolditalic": lp("DejaVuSans-BoldOblique.ttf"),
            },
            "noto": {
                "regular":    lp("NotoSans-Regular.ttf"),
                "bold":       lp("NotoSans-Bold.ttf"),
                "italic":     lp("NotoSans-Italic.ttf"),
                "bolditalic": lp("NotoSans-BoldItalic.ttf"),
            },
        }


_FONT_MAP: dict[str, dict[str, str]] | None = None


def _get_font_map() -> dict[str, dict[str, str]]:
    global _FONT_MAP
    if _FONT_MAP is None:
        _FONT_MAP = _build_font_map()
    return _FONT_MAP


_VARIANT_STRIP_RE = re.compile(
    r"[-_,]?(bold|italic|oblique|regular|medium|light|thin|"
    r"black|semibold|demi|narrow|condensed|extended|it|bd|bi|mt|ps|"
    r"roman|neue|pro|std|offc|web|lf|offc|sc|display|text|caption|"
    r"smallcaps|book|heavy|ultra|extra)$",
    re.IGNORECASE,
)


def _normalise_font_name(name: str) -> str:
    """Strip variant suffixes and return a lowercase token suitable for map lookup."""
    # Handle common compound names like "ArialMT" or "TimesNewRomanPS-BoldMT"
    n = name.strip()
    # Remove everything after the first + (subset prefix like "ABCDEF+Arial")
    if "+" in n:
        n = n.split("+", 1)[1]
    n = n.replace(" ", "").replace("-", "").replace("_", "").lower()
    # Iteratively strip known variant suffixes
    for _ in range(6):
        prev = n
        n = _VARIANT_STRIP_RE.sub("", n)
        if n == prev:
            break
    return n or "arial"


def _variant_key(flags: int) -> str:
    bold = bool(flags & _FLAG_BOLD)
    italic = bool(flags & _FLAG_ITALIC)
    if bold and italic:
        return "bolditalic"
    if bold:
        return "bold"
    if italic:
        return "italic"
    return "regular"


# ---------------------------------------------------------------------------
# Per-page font cache:  (normalised_name, variant_key) → registered fontname
# ---------------------------------------------------------------------------

def _resolve_font(
    doc: fitz.Document,
    page: fitz.Page,
    original_font_name: str,
    flags: int,
    page_cache: dict[tuple[str, str], str],
) -> str:
    """Resolve the best available font for *original_font_name* + *flags*.

    Resolution order:
      1. Cache hit (already registered this page)
      2. Embed font buffer extracted directly from the PDF
      3. Match to a system font family via the platform font map
      4. Fall back to Unicode-capable built-in / system font
    """
    norm = _normalise_font_name(original_font_name)
    variant = _variant_key(flags)
    cache_key = (norm, variant)

    if cache_key in page_cache:
        return page_cache[cache_key]

    # --- Strategy 1: reuse font embedded in the PDF ---
    # Skip CID/Type0 composite fonts – they use glyph IDs, not Unicode codepoints
    cid_subtypes = {"Type0", "CIDFontType2", "CIDFontType0"}
    registered = _try_embed_pdf_font(doc, page, original_font_name, norm, variant, page_cache, cid_subtypes)
    if registered:
        page_cache[cache_key] = registered
        return registered

    # --- Strategy 2: match to platform system font ---
    system = _try_system_font(page, norm, variant, page_cache, cache_key)
    if system:
        page_cache[cache_key] = system
        return system

    # --- Strategy 3: fallback ---
    fallback = _get_fallback_font(page, flags, page_cache)
    page_cache[cache_key] = fallback
    return fallback


def _try_embed_pdf_font(
    doc: fitz.Document,
    page: fitz.Page,
    original_name: str,
    norm: str,
    variant: str,
    page_cache: dict[tuple[str, str], str],
    cid_subtypes: set[str],
) -> str | None:
    try:
        for font_info in page.get_fonts(full=True):
            # font_info: (xref, ext, type, basefont, name, encoding, referencer)
            xref, _ext, ftype, basefont, fname, _enc, *_ = font_info
            if ftype in cid_subtypes:
                continue
            candidate_name = basefont or fname or ""
            candidate_norm = _normalise_font_name(candidate_name)
            # Accept if normalised names share a common root (either is a prefix)
            if not (candidate_norm.startswith(norm) or norm.startswith(candidate_norm)):
                continue
            # Try to extract and re-embed the font buffer
            font_data = doc.extract_font(xref)
            if not font_data or not font_data[3]:  # font_data[3] is the buffer
                continue
            buf = font_data[3]
            try:
                fontname = f"emb{abs(hash(original_name + variant)) % 100000:05d}"
                # Avoid re-registering same xref under a different key
                reuse_key = ("__xref__", str(xref))
                if reuse_key in page_cache:
                    existing = page_cache[reuse_key]
                    page_cache[(norm, variant)] = existing
                    return existing
                page.insert_font(fontname=fontname, fontbuffer=buf)
                page_cache[reuse_key] = fontname
                return fontname
            except Exception:
                # Subset fonts that don't cover translated glyphs will fail
                logger.debug(
                    "Could not re-embed PDF font '%s' (likely subset); falling through",
                    candidate_name,
                )
    except Exception:
        logger.debug("Error enumerating page fonts", exc_info=True)
    return None


def _try_system_font(
    page: fitz.Page,
    norm: str,
    variant: str,
    page_cache: dict[tuple[str, str], str],
    cache_key: tuple[str, str],
) -> str | None:
    font_map = _get_font_map()

    # Direct match
    families_to_try: list[str] = []
    if norm in font_map:
        families_to_try.append(norm)

    # Partial match: find any family whose token is a substring of norm or vice-versa
    if not families_to_try:
        for token in font_map:
            if token in norm or norm in token:
                families_to_try.append(token)

    for family in families_to_try:
        variants = font_map[family]
        path = variants.get(variant) or variants.get("regular") or ""
        if path and Path(path).is_file():
            fontname = f"sys{abs(hash(family + variant)) % 100000:05d}"
            # Check if already registered under a different cache key
            reuse_key = ("__path__", path)
            if reuse_key in page_cache:
                return page_cache[reuse_key]
            try:
                page.insert_font(fontname=fontname, fontfile=path)
                page_cache[reuse_key] = fontname
                return fontname
            except Exception:
                logger.debug("Failed to register system font %s", path)
    return None


# Cached result of fallback font detection (survives across pages)
_FALLBACK_FONT_CACHE: dict[str, str] = {}


def _get_fallback_font(
    page: fitz.Page,
    flags: int,
    page_cache: dict[tuple[str, str], str],
) -> str:
    """Register and return the best available Unicode fallback font."""
    variant = _variant_key(flags)
    fb_key = ("__fallback__", variant)
    if fb_key in page_cache:
        return page_cache[fb_key]

    # pymupdf-fonts provides Noto variants
    try:
        import pymupdf_fonts  # noqa: F401
        noto_map = {
            "regular":    "notos",
            "bold":       "notosb",
            "italic":     "notosi",   # may not exist; caught below
            "bolditalic": "notosbi",
        }
        name = noto_map.get(variant, "notos")
        # "notos" is always safe; others may not be available in all versions
        result = name if variant == "regular" else "notos"
        try:
            # Verify the font name is actually available
            test_font = fitz.Font(fontname=name)
            del test_font
            result = name
        except Exception:
            result = "notos"
        page_cache[fb_key] = result
        return result
    except ImportError:
        pass

    # System generic fallback — prefer bold variant if available
    if sys.platform == "win32":
        import os
        fd = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"
        candidates = {
            "regular":    fd / "arial.ttf",
            "bold":       fd / "arialbd.ttf",
            "italic":     fd / "ariali.ttf",
            "bolditalic": fd / "arialbi.ttf",
        }
        path = candidates.get(variant) or candidates["regular"]
        if path.is_file():
            fontname = f"fb{variant[:2]}"
            reuse = ("__path__", str(path))
            if reuse in page_cache:
                page_cache[fb_key] = page_cache[reuse]
                return page_cache[reuse]
            try:
                page.insert_font(fontname=fontname, fontfile=str(path))
                page_cache[reuse] = fontname
                page_cache[fb_key] = fontname
                return fontname
            except Exception:
                pass
    elif sys.platform != "darwin":
        for p in (
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
            "/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf",
        ):
            if Path(p).is_file():
                fontname = "fbgen"
                try:
                    page.insert_font(fontname=fontname, fontfile=p)
                    page_cache[fb_key] = fontname
                    return fontname
                except Exception:
                    continue

    page_cache[fb_key] = "helv"
    return "helv"


# ---------------------------------------------------------------------------
# Line-level span grouping helpers
# ---------------------------------------------------------------------------

def _group_spans_by_line(blocks: list[dict]) -> list[dict]:
    """Group spans into lines and return per-line metadata including font name."""
    lines_out: list[dict] = []
    for block in blocks:
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            if not spans:
                continue
            # Filter out empty spans
            valid_spans = [s for s in spans if s.get("text", "").strip()]
            if not valid_spans:
                continue

            # Build combined text (preserve inter-span spacing)
            combined_text = " ".join(s["text"].strip() for s in valid_spans)
            if not combined_text.strip():
                continue

            # Determine predominant size / color / font by character count
            size_counts: dict[float, int] = {}
            color_counts: dict[int, int] = {}
            font_counts: dict[str, int] = {}
            combined_flags = 0
            for s in valid_spans:
                n = len(s.get("text", ""))
                sz = s.get("size", 11)
                cl = s.get("color", 0)
                fn = s.get("font", "")
                size_counts[sz] = size_counts.get(sz, 0) + n
                color_counts[cl] = color_counts.get(cl, 0) + n
                font_counts[fn] = font_counts.get(fn, 0) + n
                combined_flags |= s.get("flags", 0)

            best_size = max(size_counts, key=lambda k: size_counts[k])
            best_color = max(color_counts, key=lambda k: color_counts[k])
            best_font = max(font_counts, key=lambda k: font_counts[k])

            # Line bounding rect = union of all span bboxes
            rects = [fitz.Rect(s["bbox"]) for s in valid_spans]
            line_rect = rects[0]
            for r in rects[1:]:
                line_rect |= r  # union

            lines_out.append({
                "rect": line_rect,
                "text": combined_text,
                "size": best_size,
                "color": best_color,
                "flags": combined_flags,
                "font":  best_font,
                "spans": valid_spans,
            })
    return lines_out


# ---------------------------------------------------------------------------
# Layout helpers
# ---------------------------------------------------------------------------

def _measure_text_height(
    text: str,
    fontname: str,
    fontsize: float,
    width: float,
) -> float:
    """Estimate the rendered height of *text* in a box of given *width*.

    Uses fitz.TextWriter to measure actual layout height; falls back to a
    conservative line-count heuristic if the writer approach is unavailable.
    """
    try:
        font = fitz.Font(fontname=fontname)
        tw = fitz.TextWriter(fitz.Rect(0, 0, width, 10000))
        tw.fill_textbox(
            fitz.Rect(0, 0, width, 10000),
            text,
            font=font,
            fontsize=fontsize,
            warn=False,
        )
        used = tw.text_rect
        return max(used.height if used else fontsize * 1.2, fontsize * 1.2)
    except Exception:
        # Rough heuristic: assume average char width ≈ 0.55 × fontsize
        avg_char_w = fontsize * 0.55
        chars_per_line = max(int(width / avg_char_w), 1)
        n_lines = max(1, -(-len(text) // chars_per_line))  # ceiling division
        return n_lines * fontsize * 1.35


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def translate_pdf(
    input_path: str,
    output_path: str,
    translator: Any,
    target_lang: str,
    *,
    cancel_event: threading.Event | None = None,
) -> str | None:
    """Translate text in a PDF while preserving fonts, styles, and layout.

    For each page:
    1.  Extract structured text (blocks → lines → spans).
    2.  Translate all lines in a single batch where possible.
    3.  Redact original text areas.
    4.  Two-pass re-insertion:
        • Pass A: measure required height for each translated line.
        • Pass B: insert text, shifting subsequent lines down when the
          translated text is taller than the original.
    """
    src = fitz.open(input_path)

    # --- OCR check: skip scanned PDFs ---
    if _is_ocr_required(src):
        logger.warning(
            "Skipping '%s': appears to be a scanned/image PDF requiring OCR.",
            Path(input_path).name,
        )
        src.close()
        return "skipped-ocr"

    errors = 0

    for page_num in range(src.page_count):
        # Honour cancellation between pages
        if cancel_event is not None and cancel_event.is_set():
            src.close()
            raise CancelledError("Translation cancelled")

        page = src[page_num]
        page_height = page.mediabox.y1

        # Get structured text: blocks -> lines -> spans
        text_dict: dict[str, Any] = page.get_text(  # type: ignore[assignment]
            "dict", flags=fitz.TEXT_PRESERVE_WHITESPACE
        )
        blocks = text_dict.get("blocks", [])

        # Group spans by line for context-aware translation
        line_infos = _group_spans_by_line(blocks)
        if not line_infos:
            continue

        # ------------------------------------------------------------------
        # Batch-translate all line texts for this page
        # ------------------------------------------------------------------
        original_texts = [li["text"] for li in line_infos]
        try:
            translated_texts = translator.translate_batch(
                original_texts, target_lang, cancel_event=cancel_event
            )
        except CancelledError:
            src.close()
            raise
        except Exception:
            logger.exception(
                "Batch translation failed on page %d; falling back to per-item",
                page_num + 1,
            )
            translated_texts = []
            for t in original_texts:
                try:
                    r = translator.translate_text(t, target_lang)
                    translated_texts.append(r if r is not None else t)
                except Exception:
                    logger.exception("Per-item fallback failed")
                    translated_texts.append(t)
                    errors += 1

        # ------------------------------------------------------------------
        # Redact original text per-line
        # fill=None → do NOT paint white over the area (preserves background)
        # ------------------------------------------------------------------
        for info in line_infos:
            page.add_redact_annot(info["rect"], fill=None)  # type: ignore[arg-type]
        # 0 = PDF_REDACT_IMAGE_NONE → leave underlying images untouched
        page.apply_redactions(images=0)  # type: ignore[arg-type]

        # ------------------------------------------------------------------
        # Per-page font resolution cache
        # ------------------------------------------------------------------
        font_cache: dict[tuple[str, str], str] = {}

        # ------------------------------------------------------------------
        # Pass A: pre-resolve fonts + measure required heights
        # ------------------------------------------------------------------
        resolved_infos: list[dict] = []
        for info, tr_text in zip(line_infos, translated_texts):
            if tr_text is None:
                tr_text = info["text"]
                errors += 1

            fontname = _resolve_font(
                src, page,
                info["font"],
                info["flags"],
                font_cache,
            )
            orig_rect: fitz.Rect = info["rect"]
            fontsize: float = info["size"]
            rect_width = max(orig_rect.width, 10.0)

            needed_h = _measure_text_height(tr_text, fontname, fontsize, rect_width)
            resolved_infos.append({
                "orig_rect": orig_rect,
                "text": tr_text,
                "size": fontsize,
                "color": info["color"],
                "flags": info["flags"],
                "fontname": fontname,
                "needed_h": needed_h,
            })

        # ------------------------------------------------------------------
        # Pass B: insert with cumulative vertical shift
        # Sort by vertical position so shift propagates top-to-bottom
        # ------------------------------------------------------------------
        resolved_infos.sort(key=lambda x: x["orig_rect"].y0)
        cumulative_shift = 0.0

        for ri in resolved_infos:
            orig_rect: fitz.Rect = ri["orig_rect"]
            tr_text: str = ri["text"]
            fontsize: float = ri["size"]
            fontname: str = ri["fontname"]
            needed_h: float = ri["needed_h"]

            # Convert integer colour to (r, g, b) tuple
            c = ri["color"]
            if isinstance(c, int):
                color = (
                    ((c >> 16) & 0xFF) / 255.0,
                    ((c >> 8)  & 0xFF) / 255.0,
                    (c & 0xFF)         / 255.0,
                )
            else:
                color = (0.0, 0.0, 0.0)

            # Compute shifted insertion rect
            insertion_rect = fitz.Rect(
                orig_rect.x0,
                orig_rect.y0 + cumulative_shift,
                orig_rect.x1,
                orig_rect.y0 + cumulative_shift + max(needed_h, orig_rect.height),
            )

            # Cap to page bottom (never push content off-page)
            if insertion_rect.y1 > page_height - 5:
                insertion_rect.y1 = page_height - 5
                if insertion_rect.y0 >= insertion_rect.y1:
                    insertion_rect.y0 = max(0.0, insertion_rect.y1 - orig_rect.height)

            extra_h = max(0.0, needed_h - orig_rect.height)
            cumulative_shift += extra_h

            # Try to insert at full size first; if text still overflows the
            # (possibly expanded) rect, shrink slightly — but never below
            # _MIN_SCALE of the original size.
            inserted = False
            for scale in (1.0, 0.95, 0.90, _MIN_SCALE):
                try:
                    trial_size = max(fontsize * scale, 4.0)
                    rc = page.insert_textbox(
                        insertion_rect,
                        tr_text,
                        fontsize=trial_size,
                        fontname=fontname,
                        color=color,
                        align=fitz.TEXT_ALIGN_LEFT,
                    )
                    if rc >= 0:
                        inserted = True
                        break
                except Exception:
                    continue

            if not inserted:
                # Last resort: insert at minimum readable size, accept overflow
                try:
                    page.insert_textbox(
                        insertion_rect,
                        tr_text,
                        fontsize=max(fontsize * _MIN_SCALE, 4.0),
                        fontname=fontname,
                        color=color,
                        align=fitz.TEXT_ALIGN_LEFT,
                    )
                except Exception:
                    logger.exception(
                        "Failed to insert translated text on page %d", page_num + 1
                    )
                    errors += 1

    # Check for cancellation before saving
    if cancel_event is not None and cancel_event.is_set():
        src.close()
        raise CancelledError("Translation cancelled before saving PDF")

    src.save(str(output_path), garbage=4, deflate=True, clean=True)
    src.close()

    if errors:
        raise RuntimeError(f"PDF translation completed with {errors} issues")
    return None
