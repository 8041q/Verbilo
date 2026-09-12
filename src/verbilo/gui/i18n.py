from __future__ import annotations

from dataclasses import dataclass
import json
from importlib import resources
from typing import Any, Iterable


DEFAULT_UI_LOCALE = "en"

_SUPPLEMENTAL_STRINGS: dict[str, dict[str, str]] = {
    "en": {
        "sidebar.terminology": "Terminology",
        "app.exit_running_title": "Translation in progress",
        "app.exit_running_body": "Cancel the current translation and exit Verbilo?",
        "app.cancelling_exit": "Cancelling translation before exit…",
        "content.empty_queue": "No files added yet.\nUse Add Files or Select Folder to begin.",
        "content.remove_selected": "Remove selected",
        "content.clear_all": "Clear all",
        "tooltip.add_files": "Add files (Ctrl+O)",
        "tooltip.select_folder": "Select a folder (Ctrl+Shift+O)",
        "tooltip.clear_files": "Remove selected files, or clear the queue when nothing is selected",
        "tooltip.output_folder": "Choose output folder",
        "tooltip.start": "Start pending translations (Ctrl+Enter)",
        "tooltip.cancel": "Cancel the current translation",
        "tooltip.terminology": "Manage terminology (Ctrl+G)",
        "tooltip.settings": "Open Settings (Ctrl+,)",
        "tooltip.about": "About Verbilo (F1)",
        "tooltip.report": "Open the completed translation report",
        "tooltip.retry": "Retry failed or cancelled files only",
        "file_dialog.select_files": "Select files",
        "file_dialog.all_supported": "All supported files",
        "file_dialog.all_files": "All files",
        "content.drop_hint": "Drag and drop files here",
        "content.drop_hint_active": "Drop files to add them",
        "content.drop_added": "Added {count} file(s)",
        "content.file_list_locked": "The file list cannot be changed while translation is running.",
        "settings.ollama.runtime": "Ollama status",
        "settings.ollama.runtime.checking": "Checking Ollama…",
        "settings.ollama.runtime.connected": "Connected",
        "settings.ollama.runtime.stopped": "Installed, not running",
        "settings.ollama.runtime.not_installed": "Ollama is not installed",
        "settings.ollama.runtime.unavailable": "Ollama unavailable",
        "settings.ollama.model_heading": "Translation model",
        "settings.ollama.model.qwen": "Qwen 3.5 4B",
        "settings.ollama.model.hymt": "HY-MT 1.5 1.8B",
        "settings.ollama.model.translategemma": "TranslateGemma 4B",
        "settings.ollama.model.mistral": "Mistral 7B",
        "settings.ollama.model_desc.qwen": "Recommended general-purpose model. Good balance of translation quality and resource use.",
        "settings.ollama.model_desc.hymt": "Lightweight translation-focused model with lower memory requirements.",
        "settings.ollama.model_desc.translategemma": "Translation-focused model designed specifically for multilingual translation.",
        "settings.ollama.model_desc.mistral": "Larger general-purpose model. Uses more memory than the lighter options.",
        "settings.ollama.state.checking": "Checking selected model…",
        "settings.ollama.state.ready": "Ready to use",
        "settings.ollama.state.ready_detail": "This model is downloaded and available in Ollama.",
        "settings.ollama.state.not_downloaded": "Not downloaded",
        "settings.ollama.state.not_downloaded_detail": "Download this model before using it for local translation.",
        "settings.ollama.state.unknown": "Model status unavailable",
        "settings.ollama.state.unknown_detail": "Connect to Ollama to check whether this model is installed.",
        "settings.ollama.state.starting": "Starting Ollama…",
        "settings.ollama.state.downloading": "Downloading {model}",
        "settings.ollama.state.verifying": "Verifying download…",
        "settings.ollama.state.removing": "Removing {model}…",
        "settings.ollama.state.cancelled": "Download cancelled",
        "settings.ollama.state.error": "Action failed",
        "settings.ollama.progress.percent": "{stage} · {percent}%",
        "settings.ollama.progress.manifest": "Preparing model",
        "settings.ollama.progress.layers": "Downloading model data",
        "settings.ollama.progress.verifying": "Verifying model data",
        "settings.ollama.progress.writing_manifest": "Finalizing model",
        "settings.ollama.progress.cleanup": "Cleaning up",
        "settings.ollama.progress.downloading": "Downloading model",
        "settings.ollama.action.download": "Download model",
        "settings.ollama.action.start": "Start Ollama",
        "settings.ollama.action.get": "Get Ollama",
        "settings.ollama.action.retry": "Try again",
        "settings.ollama.action.cancel": "Cancel download",
        "settings.ollama.action.remove": "Remove…",
        "settings.ollama.action.refresh": "Refresh status",
        "settings.ollama.remove_confirm_title": "Remove downloaded model?",
        "settings.ollama.remove_confirm_body": "Remove {model} from Ollama? You can download it again later.",
        "settings.ollama.removed": "Model removed. You can download it again later.",
        "settings.ollama.cancel_requested": "Cancelling download…",
        "settings.ollama.running_note": "Model removal is disabled while a translation is running.",
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
        "terminology.entry_count": "{count} entries",
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
        "report.row.attempt": "attempt {attempt} · ",
        "report.row.details": "{attempt}translated {translated}, skipped {skipped}, TM {tm} hit(s), visual warnings {warnings}",
        "report.copy_summary": "Copy summary",
        "report.copied": "Copied",
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
        "report.summary.visual": "Output validation: {checks} integrity checks · {failures} integrity failures · {retry_candidates} layout candidates · {retry_accepted} improved · {autofit} PowerPoint autofit · {visual_warnings} visual warnings",
        "report.summary.selective": "Selective translation: {target} already-target · {nonling} non-linguistic · {other} non-source",
        "report.summary.consistency": "Document consistency: {reuses} presentation variants reused across {families} group(s)",
        "report.terminology_conflicts": "Ignored ambiguous auto-source terminology: {terms}",
        "report.completed_with_warnings": "Completed with {count} warning(s) — report available",
        "report.completed": "Translation complete — report available",
        "queue.retry_failed": "Retry failed",
        "queue.nothing_pending_title": "Nothing to translate",
        "queue.nothing_pending_body": "All files in the queue are already finished. Add more files to start another translation.",
        "queue.nothing_pending_retry": "There are no new pending files. Use Retry failed to run only failed or interrupted files again.",
        "table.status.skipped": "Skipped",
        "table.status.retrying": "Retrying",
    },
    "zh-Hans": {
        "sidebar.terminology": "术语",
        "app.exit_running_title": "翻译正在进行",
        "app.exit_running_body": "取消当前翻译并退出 Verbilo？",
        "app.cancelling_exit": "正在取消翻译并准备退出…",
        "content.empty_queue": "尚未添加文件。\n请使用“添加文件”或“选择文件夹”开始。",
        "content.remove_selected": "移除所选",
        "content.clear_all": "全部清除",
        "tooltip.add_files": "添加文件 (Ctrl+O)",
        "tooltip.select_folder": "选择文件夹 (Ctrl+Shift+O)",
        "tooltip.clear_files": "移除所选文件；未选择时清空队列",
        "tooltip.output_folder": "选择输出文件夹",
        "tooltip.start": "开始待处理翻译 (Ctrl+Enter)",
        "tooltip.cancel": "取消当前翻译",
        "tooltip.terminology": "管理术语 (Ctrl+G)",
        "tooltip.settings": "打开设置 (Ctrl+,)",
        "tooltip.about": "关于 Verbilo (F1)",
        "tooltip.report": "打开已完成的翻译报告",
        "tooltip.retry": "仅重试失败或已取消的文件",
        "file_dialog.select_files": "选择文件",
        "file_dialog.all_supported": "所有支持的文件",
        "file_dialog.all_files": "所有文件",
        "content.drop_hint": "将文件拖放到此处",
        "content.drop_hint_active": "松开以添加文件",
        "content.drop_added": "已添加 {count} 个文件",
        "content.file_list_locked": "翻译进行中时无法更改文件列表。",
        "settings.ollama.runtime": "Ollama 状态",
        "settings.ollama.runtime.checking": "正在检查 Ollama…",
        "settings.ollama.runtime.connected": "已连接",
        "settings.ollama.runtime.stopped": "已安装，但未运行",
        "settings.ollama.runtime.not_installed": "尚未安装 Ollama",
        "settings.ollama.runtime.unavailable": "Ollama 不可用",
        "settings.ollama.model_heading": "翻译模型",
        "settings.ollama.model.qwen": "Qwen 3.5 4B",
        "settings.ollama.model.hymt": "HY-MT 1.5 1.8B",
        "settings.ollama.model.translategemma": "TranslateGemma 4B",
        "settings.ollama.model.mistral": "Mistral 7B",
        "settings.ollama.model_desc.qwen": "推荐的通用模型，在翻译质量和资源占用之间取得良好平衡。",
        "settings.ollama.model_desc.hymt": "轻量级翻译专用模型，内存需求较低。",
        "settings.ollama.model_desc.translategemma": "专为多语言翻译设计的翻译专用模型。",
        "settings.ollama.model_desc.mistral": "较大的通用模型，比轻量级选项需要更多内存。",
        "settings.ollama.state.checking": "正在检查所选模型…",
        "settings.ollama.state.ready": "可以使用",
        "settings.ollama.state.ready_detail": "此模型已下载并可在 Ollama 中使用。",
        "settings.ollama.state.not_downloaded": "尚未下载",
        "settings.ollama.state.not_downloaded_detail": "使用本地翻译前，请先下载此模型。",
        "settings.ollama.state.unknown": "无法获取模型状态",
        "settings.ollama.state.unknown_detail": "连接 Ollama 后即可检查此模型是否已安装。",
        "settings.ollama.state.starting": "正在启动 Ollama…",
        "settings.ollama.state.downloading": "正在下载 {model}",
        "settings.ollama.state.verifying": "正在验证下载…",
        "settings.ollama.state.removing": "正在移除 {model}…",
        "settings.ollama.state.cancelled": "下载已取消",
        "settings.ollama.state.error": "操作失败",
        "settings.ollama.progress.percent": "{stage} · {percent}%",
        "settings.ollama.progress.manifest": "正在准备模型",
        "settings.ollama.progress.layers": "正在下载模型数据",
        "settings.ollama.progress.verifying": "正在验证模型数据",
        "settings.ollama.progress.writing_manifest": "正在完成模型安装",
        "settings.ollama.progress.cleanup": "正在清理",
        "settings.ollama.progress.downloading": "正在下载模型",
        "settings.ollama.action.download": "下载模型",
        "settings.ollama.action.start": "启动 Ollama",
        "settings.ollama.action.get": "获取 Ollama",
        "settings.ollama.action.retry": "重试",
        "settings.ollama.action.cancel": "取消下载",
        "settings.ollama.action.remove": "移除…",
        "settings.ollama.action.refresh": "刷新状态",
        "settings.ollama.remove_confirm_title": "移除已下载模型？",
        "settings.ollama.remove_confirm_body": "从 Ollama 中移除 {model}？之后仍可再次下载。",
        "settings.ollama.removed": "模型已移除。之后可以再次下载。",
        "settings.ollama.cancel_requested": "正在取消下载…",
        "settings.ollama.running_note": "翻译进行中时无法移除模型。",
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
        "terminology.entry_count": "{count} 条",
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
        "report.row.attempt": "第 {attempt} 次 · ",
        "report.row.details": "{attempt}已翻译 {translated}，跳过 {skipped}，翻译记忆命中 {tm}，视觉警告 {warnings}",
        "report.copy_summary": "复制摘要",
        "report.copied": "已复制",
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
        "report.summary.visual": "输出验证：完整性检查 {checks} · 完整性失败 {failures} · 布局候选 {retry_candidates} · 改善 {retry_accepted} · PowerPoint 自动适应 {autofit} · 视觉警告 {visual_warnings}",
        "report.summary.selective": "选择性翻译：已是目标语言 {target} · 非语言内容 {nonling} · 非源语言 {other}",
        "report.summary.consistency": "文档一致性：在 {families} 个组中复用了 {reuses} 个展示形式变体",
        "report.terminology_conflicts": "自动源语言模式中忽略了有歧义的术语：{terms}",
        "report.completed_with_warnings": "已完成，存在 {count} 个警告 — 报告可用",
        "report.completed": "翻译完成 — 可查看报告",
        "queue.retry_failed": "重试失败项",
        "queue.nothing_pending_title": "没有待翻译文件",
        "queue.nothing_pending_body": "队列中的文件均已完成。请添加新文件后再开始翻译。",
        "queue.nothing_pending_retry": "没有新的待处理文件。请使用“重试失败项”仅重新运行失败或中断的文件。",
        "table.status.skipped": "已跳过",
        "table.status.retrying": "重试中",
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