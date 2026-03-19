# Translator factory — selects the appropriate backend based on engine name.

from __future__ import annotations

import logging
from typing import Optional

from .base import Translator

logger = logging.getLogger(__name__)


class TranslatorFactory:
    @staticmethod
    def get(
        name: Optional[str] = None,
        source_lang: str = "auto",
        detector: str = "fasttext",
        *,
        engine: str = "google",
        proxies: Optional[dict] = None,
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
    ) -> Translator:
        engine = (engine or "google").strip().lower()

        # --- Local offline (OPUS-MT via CTranslate2) ---
        if engine == "local":
            from .local import OpusMTTranslator
            from pathlib import Path
            model_dir = local_model_dir or str(
                Path(__file__).resolve().parents[3] / "models" / "opus-mt"
            )
            return OpusMTTranslator(
                model_dir=model_dir, source_lang=source_lang, detector=detector,
            )

        from .google import (
            IdentityTranslator,
            DeepTranslatorWrapper,
            GoogleCloudTranslatorWrapper,
            GoogleCloudV3TranslatorWrapper,
        )

        # --- Microsoft Azure Translator ---
        if engine == "azure":
            if not azure_key or not azure_region:
                logger.error("Azure engine selected but api_key/region not provided")
                return IdentityTranslator()
            from .azure import AzureTranslatorWrapper
            return AzureTranslatorWrapper(
                api_key=azure_key, region=azure_region,
                source_lang=source_lang, detector=detector,
                proxies=proxies,
            )

        # --- DeepL Free / Pro ---
        if engine in ("deepl", "deepl-free", "deepl-pro"):
            if not deepl_api_key:
                logger.error("DeepL engine selected but api_key not provided")
                return IdentityTranslator()
            from .deepl import DeepLTranslatorWrapper
            return DeepLTranslatorWrapper(
                api_key=deepl_api_key,
                source_lang=source_lang, detector=detector,
                proxies=proxies,
                pro=(engine == "deepl-pro"),
            )

        # --- Baidu Translate (Standard or Premium tier) ---
        if engine == "baidu":
            if not baidu_appid or not baidu_appkey:
                logger.error("Baidu engine selected but appid/appkey not provided")
                return IdentityTranslator()
            from .baidu import BaiduTranslatorWrapper
            return BaiduTranslatorWrapper(
                appid=baidu_appid, appkey=baidu_appkey,
                source_lang=source_lang, detector=detector,
                proxies=proxies,
                tier=baidu_tier,
            )

        # --- Google Cloud Translation API v3 (Advanced) ---
        if engine == "google-cloud-v3":
            if not google_project_id:
                logger.error("Google Cloud v3 engine selected but project_id not provided")
                return IdentityTranslator()
            return GoogleCloudV3TranslatorWrapper(
                project_id=google_project_id,
                sa_json=google_sa_json,
                source_lang=source_lang,
                detector=detector,
            )

        # --- Google Cloud Translation API v2 (Basic) ---
        if engine in ("google-cloud", "google_cloud"):
            if not google_api_key:
                logger.warning("Google Cloud engine selected but no API key — falling back to free Google")
                engine = "google"
            else:
                return GoogleCloudTranslatorWrapper(
                    api_key=google_api_key,
                    source_lang=source_lang,
                    detector=detector,
                    proxies=proxies,
                )

        # --- Google Translate (free mobile scraper, default) ---
        if name and name.lower() == "identity":
            return IdentityTranslator()
        try:
            from deep_translator import GoogleTranslator  # type: ignore  # noqa: F401
            return DeepTranslatorWrapper(
                source_lang=source_lang, detector=detector,
                proxies=proxies,
            )
        except Exception:
            logger.warning(
                "deep_translator is not available — returning IdentityTranslator "
                "(text will NOT be translated). Install it with: pip install deep-translator"
            )
            return IdentityTranslator()
