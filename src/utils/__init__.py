import threading


# raised when a cancel event fires mid-translation
class CancelledError(Exception):
    pass
