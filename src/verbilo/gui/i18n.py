from __future__ import annotations

from dataclasses import dataclass
import json
from importlib import resources
from typing import Any, Iterable


DEFAULT_UI_LOCALE = "en"

_SUPPLEMENTAL_STRINGS: dict[str, dict[str, str]] = {
    "en": {
        "sidebar.terminology": "Terminology",
        "content.drop_hint": "Drag and drop files here",
        "content.drop_hint_active": "Drop files to add them",
        "content.drop_added": "Added {count} file(s)",
        "content.file_list_locked": "The file list cannot be changed while translation is running.",
        "settings.backend_cache.title": "Translation cache",
        "settings.backend_cache.clear": "Clear translation cache",
        "settings.backend_cache.hint": "Recommended and enabled automatically. Reuses results from the same translation service for identical requests, reducing repeated API calls and local-model work. Stored locally in plaintext and safe to clear.",
        "settings.translation_memory.enable": "Enable Translation Memory",
        "settings.translation_memory.hint": "Optional. Reuses validated translations even after switching translation services. Most useful for recurring documents and consistent wording. Most users can leave this off. Source and translated text are stored persistently in plaintext SQLite.",
        "settings.section.translation_memory": "Translation Memory",
        "settings.translation_memory_enabled": "Reuse validated translations across documents",
        "settings.translation_memory_path": "Database location",
        "settings.translation_memory_path_hint": "Stores source and translated text in plaintext SQLite.",
        "settings.translation_memory.entries": "{entries} entries · {size}",
        "settings.translation_memory.manage": "Inspect",
        "settings.translation_memory.clear": "Clear",
        "settings.translation_memory.confirm_clear_title": "Clear Translation Memory?",
        "settings.translation_memory.confirm_clear_body": "Delete all saved translations from this Translation Memory? This cannot be undone.",
        "settings.translation_memory.running_note": "Translation Memory cannot be cleared while a translation is running.",
        "settings.error.translation_memory": "Translation Memory error: {error}",
        "settings.error.save_failed": "Settings could not be saved. Check the application log and permissions.",
        "dialog.select_translation_memory": "Choose Translation Memory database",
        "terminology.title": "Terminology",
        "terminology.intro": "Preferred terms are stored locally in plaintext JSON and applied to future translations for the matching language pair.",
        "terminology.running_note": "A translation is running. Changes are safe, but apply to the next job only.",
        "terminology.add": "Add",
        "terminology.edit": "Edit",
        "terminology.delete": "Delete",
        "terminology.import": "Import",
        "terminology.export": "Export",
        "terminology.close": "Close",
        "terminology.enabled": "Enabled",
        "terminology.source_language": "Source language",
        "terminology.target_language": "Target language",
        "terminology.source_term": "Source term",
        "terminology.target_term": "Preferred translation",
        "terminology.column.enabled": "On",
        "terminology.column.source_language": "From",
        "terminology.column.target_language": "To",
        "terminology.column.source": "Source term",
        "terminology.column.target": "Preferred translation",
        "terminology.error.required": "Choose both languages and enter both terms.",
        "terminology.error.duplicate": "That source term already exists for this language pair.",
        "terminology.confirm_delete_title": "Delete terminology entry?",
        "terminology.confirm_delete_body": "Delete the selected terminology entry?",
        "terminology.import_success": "Imported {added} new and updated {updated} terminology entries.",
        "terminology.export_success": "Exported {count} terminology entries.",
        "terminology.import_error": "Could not import terminology: {error}",
        "terminology.export_error": "Could not export terminology: {error}",
        "terminology.load_error_title": "Terminology unavailable",
        "terminology.load_error_body": "The terminology file could not be read: {error}\n\nFix or re-import it before starting this translation so required terms are not silently ignored.",
        "tm_manager.title": "Translation Memory",
        "tm_manager.search": "Search source or translation",
        "tm_manager.refresh": "Refresh",
        "tm_manager.delete": "Delete selected",
        "tm_manager.close": "Close",
        "tm_manager.column.source": "Source",
        "tm_manager.column.translation": "Translation",
        "tm_manager.column.languages": "Languages",
        "tm_manager.column.uses": "Uses",
        "tm_manager.empty": "No Translation Memory entries found.",
        "tm_manager.showing": "Showing {shown} of {total} entries · {size}",
        "report.column.details": "Details",
        "tm_manager.confirm_delete_title": "Delete Translation Memory entry?",
        "tm_manager.confirm_delete_body": "Delete the selected saved translation?",
        "report.view": "View report",
        "report.title": "Translation report",
        "report.no_report": "No completed translation report is available yet.",
        "report.summary.files": "Files: {finished} finished, {failed} failed, {skipped} skipped, {cancelled} cancelled",
        "report.summary.units": "Text units: {units} total · {translated} translated · {skipped} preserved/skipped",
        "report.summary.cache": "Translation cache: {hits} reused · {writes} saved · {errors} errors",
        "report.summary.tm": "Translation Memory: {hits} reused · {writes} new entries · {errors} errors",
        "report.summary.terminology": "Terminology: {entries} active entries · {mismatches} mismatches",
        "report.summary.quality": "Quality/layout: {fallbacks} fallbacks · {retries} retries · {warnings} layout warnings · {failed} failed items",
        "report.summary.selective": "Selective translation: {target} already-target · {nonling} non-linguistic · {other} non-source",
        "report.summary.consistency": "Document consistency: {reuses} presentation variants reused across {families} group(s)",
        "report.terminology_conflicts": "Ignored ambiguous auto-source terminology: {terms}",
        "report.completed_with_warnings": "Completed with {count} warning(s) — report available",
        "report.completed": "Translation complete — report available",
    },
    "zh-Hans": {
        "sidebar.terminology": "术语",
        "content.drop_hint": "将文件拖放到此处",
        "content.drop_hint_active": "松开以添加文件",
        "content.drop_added": "已添加 {count} 个文件",
        "content.file_list_locked": "翻译进行中时无法更改文件列表。",
        "settings.backend_cache.title": "翻译缓存",
        "settings.backend_cache.clear": "清除翻译缓存",
        "settings.backend_cache.hint": "推荐使用并自动启用。相同翻译服务遇到相同请求时会复用结果，从而减少重复 API 调用和本地模型工作。内容以明文保存在本地，可安全清除。",
        "settings.translation_memory.enable": "启用翻译记忆库",
        "settings.translation_memory.hint": "可选。即使更换翻译服务，也可继续复用已验证译文。最适合重复文档和需要一致措辞的场景。大多数用户可以保持关闭。源文本和译文会长期以明文 SQLite 保存。",
        "settings.section.translation_memory": "翻译记忆库",
        "settings.translation_memory_enabled": "在文档之间复用已验证的翻译",
        "settings.translation_memory_path": "数据库位置",
        "settings.translation_memory_path_hint": "源文本和译文会以明文形式存储在 SQLite 中。",
        "settings.translation_memory.entries": "{entries} 条 · {size}",
        "settings.translation_memory.manage": "查看",
        "settings.translation_memory.clear": "清空",
        "settings.translation_memory.confirm_clear_title": "清空翻译记忆库？",
        "settings.translation_memory.confirm_clear_body": "删除此翻译记忆库中的所有已保存翻译？此操作无法撤销。",
        "settings.translation_memory.running_note": "翻译进行中时无法清空翻译记忆库。",
        "settings.error.translation_memory": "翻译记忆库错误：{error}",
        "settings.error.save_failed": "设置无法保存。请检查应用日志和文件权限。",
        "dialog.select_translation_memory": "选择翻译记忆库数据库",
        "terminology.title": "术语",
        "terminology.intro": "首选术语以明文 JSON 保存在本地，并应用于之后匹配语言对的翻译。",
        "terminology.running_note": "当前正在翻译。更改是安全的，但仅对下一次任务生效。",
        "terminology.add": "添加",
        "terminology.edit": "编辑",
        "terminology.delete": "删除",
        "terminology.import": "导入",
        "terminology.export": "导出",
        "terminology.close": "关闭",
        "terminology.enabled": "启用",
        "terminology.source_language": "源语言",
        "terminology.target_language": "目标语言",
        "terminology.source_term": "源术语",
        "terminology.target_term": "首选译文",
        "terminology.column.enabled": "启用",
        "terminology.column.source_language": "源",
        "terminology.column.target_language": "目标",
        "terminology.column.source": "源术语",
        "terminology.column.target": "首选译文",
        "terminology.error.required": "请选择源语言和目标语言，并填写两个术语。",
        "terminology.error.duplicate": "此语言对中已存在该源术语。",
        "terminology.confirm_delete_title": "删除术语条目？",
        "terminology.confirm_delete_body": "删除所选术语条目？",
        "terminology.import_success": "已导入 {added} 个新条目并更新 {updated} 个术语条目。",
        "terminology.export_success": "已导出 {count} 个术语条目。",
        "terminology.import_error": "无法导入术语：{error}",
        "terminology.export_error": "无法导出术语：{error}",
        "terminology.load_error_title": "术语不可用",
        "terminology.load_error_body": "无法读取术语文件：{error}\n\n请先修复或重新导入，以免翻译时静默忽略必需术语。",
        "tm_manager.title": "翻译记忆库",
        "tm_manager.search": "搜索源文本或译文",
        "tm_manager.refresh": "刷新",
        "tm_manager.delete": "删除所选",
        "tm_manager.close": "关闭",
        "tm_manager.column.source": "源文本",
        "tm_manager.column.translation": "译文",
        "tm_manager.column.languages": "语言",
        "tm_manager.column.uses": "使用次数",
        "tm_manager.empty": "未找到翻译记忆条目。",
        "tm_manager.showing": "显示 {shown}/{total} 条 · {size}",
        "report.column.details": "详情",
        "tm_manager.confirm_delete_title": "删除翻译记忆条目？",
        "tm_manager.confirm_delete_body": "删除所选已保存翻译？",
        "report.view": "查看报告",
        "report.title": "翻译报告",
        "report.no_report": "暂无已完成的翻译报告。",
        "report.summary.files": "文件：完成 {finished}，失败 {failed}，跳过 {skipped}，取消 {cancelled}",
        "report.summary.units": "文本单元：共 {units} · 已翻译 {translated} · 保留/跳过 {skipped}",
        "report.summary.cache": "翻译缓存：复用 {hits} · 保存 {writes} · 错误 {errors}",
        "report.summary.tm": "翻译记忆库：复用 {hits} · 新增 {writes} · 错误 {errors}",
        "report.summary.terminology": "术语：启用 {entries} 条 · 不匹配 {mismatches}",
        "report.summary.quality": "质量/布局：回退 {fallbacks} · 重试 {retries} · 布局警告 {warnings} · 失败 {failed}",
        "report.summary.selective": "选择性翻译：已是目标语言 {target} · 非语言内容 {nonling} · 非源语言 {other}",
        "report.summary.consistency": "文档一致性：在 {families} 个组中复用了 {reuses} 个展示形式变体",
        "report.terminology_conflicts": "自动源语言模式中忽略了有歧义的术语：{terms}",
        "report.completed_with_warnings": "已完成，存在 {count} 个警告 — 报告可用",
        "report.completed": "翻译完成 — 可查看报告",
    },
}

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
            template = _SUPPLEMENTAL_STRINGS.get(self.locale, {}).get(key)
        if template is None:
            template = _SUPPLEMENTAL_STRINGS.get(DEFAULT_UI_LOCALE, {}).get(key)
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