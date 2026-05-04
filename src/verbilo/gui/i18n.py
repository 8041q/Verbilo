from __future__ import annotations

from dataclasses import dataclass
import json
from importlib import resources
from typing import Any, Iterable


DEFAULT_UI_LOCALE = "en"

_UI_LOCALE_ALIASES: dict[str, str] = {
    "en": "en",
    "en-us": "en",
    "en_us": "en",
    "zh": "zh-Hans",
    "zh-cn": "zh-Hans",
    "zh_cn": "zh-Hans",
    "zh-hans": "zh-Hans",
}

_UI_LOCALE_FILES: dict[str, str] = {
    "en": "ui_en.json",
    "zh-Hans": "ui_zh_hans.json",
}


def _assets_locales_dir():
    return resources.files("verbilo.assets").joinpath("locales")


def resolve_ui_locale(locale: str | None) -> str:
    if not locale:
        return DEFAULT_UI_LOCALE
    normalized = locale.strip().replace("_", "-")
    if normalized in _UI_LOCALE_FILES:
        return normalized
    lowered = normalized.lower()
    return _UI_LOCALE_ALIASES.get(lowered, DEFAULT_UI_LOCALE)


def _load_catalog(locale: str) -> dict[str, Any]:
    locale = resolve_ui_locale(locale)
    filename = _UI_LOCALE_FILES[locale]
    payload = _assets_locales_dir().joinpath(filename).read_text(encoding="utf-8")
    return json.loads(payload)


@dataclass(frozen=True)
class UiLocalizer:
    locale: str
    strings: dict[str, str]
    fallback_strings: dict[str, str]
    language_names: dict[str, str]
    fallback_language_names: dict[str, str]
    ui_locale_names: dict[str, str]
    fallback_ui_locale_names: dict[str, str]

    def t(self, key: str, **kwargs: Any) -> str:
        template = self.strings.get(key)
        if template is None:
            template = self.fallback_strings.get(key)
        if template is None:
            return key
        if not kwargs:
            return template
        try:
            return template.format(**kwargs)
        except KeyError as exc:
            missing = exc.args[0]
            raise KeyError(f"Missing format key '{missing}' for UI string '{key}'") from exc

    def language_name(self, code: str) -> str:
        return self.language_names.get(code) or self.fallback_language_names.get(code) or code

    def ui_locale_name(self, code: str) -> str:
        return self.ui_locale_names.get(code) or self.fallback_ui_locale_names.get(code) or code

    def language_label(self, code: str) -> str:
        return f"{self.language_name(code)} ({code})"

    def build_language_options(self, codes: Iterable[str]) -> list[tuple[str, str]]:
        return [(code, self.language_name(code)) for code in codes]


def load_ui_localizer(locale: str | None = None) -> UiLocalizer:
    resolved = resolve_ui_locale(locale)
    fallback_catalog = _load_catalog(DEFAULT_UI_LOCALE)
    catalog = fallback_catalog if resolved == DEFAULT_UI_LOCALE else _load_catalog(resolved)
    return UiLocalizer(
        locale=resolved,
        strings=dict(catalog.get("strings", {})),
        fallback_strings=dict(fallback_catalog.get("strings", {})),
        language_names=dict(catalog.get("language_names", {})),
        fallback_language_names=dict(fallback_catalog.get("language_names", {})),
        ui_locale_names=dict(catalog.get("ui_locale_names", {})),
        fallback_ui_locale_names=dict(fallback_catalog.get("ui_locale_names", {})),
    )


def get_supported_ui_locales(current_locale: str | None = None) -> list[tuple[str, str]]:
    localizer = load_ui_localizer(current_locale)
    return [(code, localizer.ui_locale_name(code)) for code in _UI_LOCALE_FILES]


def get_catalog_key_sets(locale: str | None = None) -> dict[str, set[str]]:
    catalog = _load_catalog(resolve_ui_locale(locale))
    return {
        "strings": set(catalog.get("strings", {}).keys()),
        "language_names": set(catalog.get("language_names", {}).keys()),
        "ui_locale_names": set(catalog.get("ui_locale_names", {}).keys()),
    }