"""ASGI entrypoint for platforms that auto-detect `main.py`."""

from api.index import app

__all__ = ["app"]
