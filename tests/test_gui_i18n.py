from __future__ import annotations

import pytest

from verbilo.gui.i18n import (
    DEFAULT_UI_LOCALE,
    UiLocalizer,
    get_catalog_key_sets,
    get_supported_ui_locales,
    load_ui_localizer,
    resolve_ui_locale,
)


def test_unknown_ui_locale_falls_back_to_english() -> None:
    localizer = load_ui_localizer("xx")
    assert localizer.locale == DEFAULT_UI_LOCALE
    assert localizer.t("sidebar.settings") == "Settings"


def test_formatting_works_for_dynamic_strings() -> None:
    localizer = load_ui_localizer("zh-Hans")
    assert localizer.t("progress.files_complete", completed=2, total=5, percent=40) == "2 / 5 个文件（40%）"


def test_missing_key_falls_back_to_english_catalog() -> None:
    localizer = UiLocalizer(
        locale="zh-Hans",
        strings={},
        fallback_strings={"demo.key": "English fallback"},
        language_names={},
        fallback_language_names={},
        ui_locale_names={},
        fallback_ui_locale_names={},
    )
    assert localizer.t("demo.key") == "English fallback"


def test_missing_format_key_raises_clear_error() -> None:
    localizer = load_ui_localizer("en")
    with pytest.raises(KeyError, match="Missing format key 'count'"):
        localizer.t("sidebar.source_language_count", total=4)


def test_supported_ui_locale_names_are_localized() -> None:
    assert get_supported_ui_locales("zh-Hans") == [
        ("en", "英语"),
        ("zh-Hans", "简体中文"),
    ]


def test_locale_key_sets_match_english_catalog() -> None:
    assert get_catalog_key_sets("zh-Hans") == get_catalog_key_sets("en")


def test_language_names_are_localized() -> None:
    localizer = load_ui_localizer("zh-Hans")
    assert localizer.language_name("en") == "英语"
    assert localizer.language_label("zh") == "中文（简体） (zh)"


def test_locale_alias_resolution() -> None:
    assert resolve_ui_locale("zh") == "zh-Hans"
    assert resolve_ui_locale("zh_CN") == "zh-Hans"
    assert resolve_ui_locale(None) == "en"