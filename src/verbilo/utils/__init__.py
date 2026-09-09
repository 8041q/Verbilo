# raised when a cancel event fires mid-translation
class CancelledError(Exception):
    pass


class ConfigurationError(ValueError):
    """Raised when a selected translation backend cannot be configured."""


class TranslationFailedError(RuntimeError):
    """Raised when a backend cannot translate every required unit safely."""
