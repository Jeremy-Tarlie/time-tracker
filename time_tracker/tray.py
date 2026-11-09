"""Icône system tray (pystray) — actions dispatchées via callbacks Tk."""

from __future__ import annotations

import threading
from collections.abc import Callable
from pathlib import Path
from typing import Protocol

from PIL import Image, ImageDraw

from time_tracker.logging_setup import get_logger

logger = get_logger()


class _TrayIcon(Protocol):
    def run(self) -> None: ...
    def stop(self) -> None: ...


def _load_tray_image() -> Image.Image:
    assets = Path(__file__).resolve().parent.parent / "assets" / "tray.png"
    if assets.is_file():
        return Image.open(assets).convert("RGBA")
    image = Image.new("RGBA", (64, 64), (45, 106, 79, 255))
    draw = ImageDraw.Draw(image)
    draw.ellipse((12, 12, 52, 52), fill=(64, 145, 108, 255))
    draw.rectangle((30, 18, 34, 46), fill=(255, 255, 255, 255))
    draw.rectangle((22, 30, 42, 34), fill=(255, 255, 255, 255))
    return image


class TrayController:
    def __init__(
        self,
        *,
        on_show: Callable[[], None],
        on_quit: Callable[[], None],
        on_toggle_timer: Callable[[], None] | None = None,
    ) -> None:
        self._on_show = on_show
        self._on_quit = on_quit
        self._on_toggle_timer = on_toggle_timer
        self._icon: _TrayIcon | None = None
        self._thread: threading.Thread | None = None

    def start(self) -> None:
        try:
            import pystray
            from pystray import MenuItem
        except ImportError:
            logger.warning("pystray indisponible — tray désactivé")
            return

        menu_items = [
            MenuItem("Afficher", lambda _icon, _item: self._on_show(), default=True),
        ]
        if self._on_toggle_timer is not None:
            menu_items.append(
                MenuItem("Start / Stop", lambda _icon, _item: self._on_toggle_timer()),
            )
        menu_items.append(MenuItem("Quitter", lambda _icon, _item: self._request_quit()))
        image = _load_tray_image()
        icon: _TrayIcon = pystray.Icon(
            "TimeTracker",
            image,
            "Time Tracker",
            pystray.Menu(*menu_items),
        )
        self._icon = icon
        self._thread = threading.Thread(target=icon.run, daemon=True)
        self._thread.start()
        logger.info("Tray démarré")

    def _request_quit(self) -> None:
        self.stop()
        self._on_quit()

    def stop(self) -> None:
        icon = self._icon
        if icon is None:
            return
        try:
            icon.stop()
        except RuntimeError as exc:
            logger.warning("Arrêt tray : %s", exc)
        self._icon = None
