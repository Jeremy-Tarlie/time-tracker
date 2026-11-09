"""Raccourcis configurables + écoute globale (pynput)."""

from __future__ import annotations

import logging
import sys
import threading
from collections.abc import Callable

from time_tracker.constants import HotkeyAction, HotkeysConfig

logger = logging.getLogger(__name__)

DEFAULT_HOTKEYS: HotkeysConfig = {
    "toggle": "ctrl+shift+space",
    "stop": "ctrl+shift+enter",
    "settings": "ctrl+shift+comma",
    "show": "ctrl+shift+t",
    "today": "ctrl+shift+d",
    "period": "ctrl+shift+p",
    "manual": "ctrl+shift+m",
}

HOTKEY_LABELS: dict[HotkeyAction, str] = {
    "toggle": "Démarrer / Pause",
    "stop": "Arrêter",
    "settings": "Réglages",
    "show": "Afficher la fenêtre",
    "today": "Aujourd'hui",
    "period": "Période",
    "manual": "Saisie manuelle",
}

HOTKEY_ACTIONS: tuple[HotkeyAction, ...] = (
    "toggle",
    "stop",
    "settings",
    "show",
    "today",
    "period",
    "manual",
)

_MOD_ORDER = ("ctrl", "alt", "shift", "cmd")
_KEYSYM_ALIASES = {
    "return": "enter",
    "kp_enter": "enter",
    "escape": "esc",
    "comma": "comma",
    "period": "period",
    "space": "space",
}
_DISPLAY: dict[str, str] = {
    "ctrl": "Ctrl",
    "alt": "Alt",
    "shift": "Shift",
    "cmd": "Cmd",
    "enter": "Entrée",
    "space": "Espace",
    "comma": ",",
    "period": ".",
    "esc": "Échap",
}


def default_hotkeys() -> HotkeysConfig:
    return {
        "toggle": DEFAULT_HOTKEYS["toggle"],
        "stop": DEFAULT_HOTKEYS["stop"],
        "settings": DEFAULT_HOTKEYS["settings"],
        "show": DEFAULT_HOTKEYS["show"],
        "today": DEFAULT_HOTKEYS["today"],
        "period": DEFAULT_HOTKEYS["period"],
        "manual": DEFAULT_HOTKEYS["manual"],
    }


def normalize_hotkey(raw: str) -> str | None:
    """Normalise ``Ctrl+Shift+Space`` → ``ctrl+shift+space`` ; None si invalide."""
    text = raw.strip().lower().replace(" ", "")
    if not text:
        return None
    text = text.replace("<", "").replace(">", "")
    parts = [part for part in text.split("+") if part]
    if not parts:
        return None
    mods: list[str] = []
    key: str | None = None
    for part in parts:
        token = part.replace("control", "ctrl").replace("command", "cmd").replace("meta", "cmd")
        token = _KEYSYM_ALIASES.get(token, token)
        if token in _MOD_ORDER:
            if token not in mods:
                mods.append(token)
        else:
            if key is not None:
                return None
            key = token
    if key is None:
        return None
    ordered_mods = [mod for mod in _MOD_ORDER if mod in mods]
    # Interdit une touche seule (casse la frappe globale) sauf F1–F24.
    is_function_key = key.startswith("f") and key[1:].isdigit() and 1 <= int(key[1:]) <= 24
    if not ordered_mods and not is_function_key:
        return None
    return "+".join([*ordered_mods, key])


def parse_hotkeys(raw: object) -> HotkeysConfig:
    result = default_hotkeys()
    if not isinstance(raw, dict):
        return result
    for action in HOTKEY_ACTIONS:
        value = raw.get(action)
        if isinstance(value, str):
            normalized = normalize_hotkey(value)
            if normalized:
                result[action] = normalized
    return result


def format_hotkey_display(combo: str) -> str:
    parts: list[str] = []
    for part in combo.split("+"):
        if part in _DISPLAY:
            parts.append(_DISPLAY[part])
        elif len(part) == 1:
            parts.append(part.upper())
        else:
            parts.append(part.capitalize())
    return "+".join(parts)


def format_hotkeys_help(hotkeys: HotkeysConfig) -> str:
    lines = [
        f"{HOTKEY_LABELS[action]} : {format_hotkey_display(hotkeys[action])}"
        for action in HOTKEY_ACTIONS
    ]
    return "\n".join(lines)


def find_hotkey_conflicts(hotkeys: HotkeysConfig) -> list[str]:
    """Retourne les combos utilisées plus d'une fois."""
    seen: dict[str, HotkeyAction] = {}
    conflicts: list[str] = []
    for action in HOTKEY_ACTIONS:
        combo = hotkeys[action]
        if combo in seen:
            conflicts.append(combo)
        else:
            seen[combo] = action
    return conflicts


def hotkey_to_pynput(combo: str) -> str:
    """``ctrl+shift+space`` → ``<ctrl>+<shift>+<space>``."""
    parts = combo.split("+")
    mapped: list[str] = []
    for part in parts:
        if part == "enter":
            mapped.append("<enter>")
        elif part == "comma":
            mapped.append(",")
        elif part == "period":
            mapped.append(".")
        elif part in _MOD_ORDER:
            mapped.append(f"<{part}>")
        elif len(part) == 1:
            mapped.append(part)
        else:
            mapped.append(f"<{part}>")
    return "+".join(mapped)


def hotkey_to_tk(combo: str) -> str:
    """``ctrl+shift+space`` → ``<Control-Shift-space>``."""
    parts = combo.split("+")
    tokens: list[str] = []
    for part in parts:
        if part == "ctrl":
            tokens.append("Control")
        elif part == "alt":
            tokens.append("Alt")
        elif part == "shift":
            tokens.append("Shift")
        elif part == "cmd":
            tokens.append("Command")
        elif part == "enter":
            tokens.append("Return")
        elif part == "comma":
            tokens.append("comma")
        elif part == "esc":
            tokens.append("Escape")
        else:
            tokens.append(part)
    return "<" + "-".join(tokens) + ">"


def combo_from_tk_event(keysym: str, state: int) -> str | None:
    """Construit une combo depuis un événement Tk KeyPress."""
    mods: list[str] = []
    if state & 0x4:
        mods.append("ctrl")
    if state & 0x1:
        mods.append("shift")
    # Bits Tk : sous Windows, 0x8 = NumLock (pas Alt) ; Alt = 0x20000.
    if sys.platform == "win32":
        if state & 0x20000:
            mods.append("alt")
    else:
        if state & 0x8:
            mods.append("alt")
        if state & 0x40:
            mods.append("cmd")
    key = keysym.lower()
    if key in {
        "control_l",
        "control_r",
        "shift_l",
        "shift_r",
        "alt_l",
        "alt_r",
        "meta_l",
        "meta_r",
        "escape",
    }:
        return None
    key = _KEYSYM_ALIASES.get(key, key)
    return normalize_hotkey("+".join([*mods, key]))


class GlobalHotkeyController:
    """Écoute globale pynput ; callbacks exécutés via ``schedule`` (thread UI)."""

    def __init__(
        self,
        *,
        schedule: Callable[[Callable[[], None]], None],
        handlers: dict[HotkeyAction, Callable[[], None]],
    ) -> None:
        self._schedule = schedule
        self._handlers = handlers
        self._listener: object | None = None
        self._lock = threading.Lock()

    def start(self, hotkeys: HotkeysConfig) -> str | None:
        """Démarre l'écoute. Retourne un message d'erreur ou None."""
        self.stop()
        conflicts = find_hotkey_conflicts(hotkeys)
        if conflicts:
            return "Raccourcis en conflit : " + ", ".join(
                format_hotkey_display(combo) for combo in conflicts
            )
        mapping: dict[str, Callable[[], None]] = {}
        for action in HOTKEY_ACTIONS:
            combo = hotkeys[action]
            handler = self._handlers.get(action)
            if handler is None:
                continue
            pynput_combo = hotkey_to_pynput(combo)

            def _make(cb: Callable[[], None] = handler) -> Callable[[], None]:
                def fire() -> None:
                    self._schedule(cb)

                return fire

            mapping[pynput_combo] = _make()
        if not mapping:
            return "Aucun raccourci valide."
        try:
            from pynput import keyboard
        except ImportError:
            return "Module pynput indisponible."
        try:
            listener = keyboard.GlobalHotKeys(mapping)
            listener.start()
        except Exception as exc:  # noqa: BLE001 — droits OS / backends pynput variés
            logger.warning("Échec démarrage hotkeys globaux: %s", exc)
            return str(exc)
        # macOS : sans Accessibility / Input Monitoring, le listener démarre mais n'écoute pas.
        if sys.platform == "darwin":
            is_trusted = getattr(listener, "IS_TRUSTED", None)
            if is_trusted is None:
                is_trusted = getattr(keyboard.Listener, "IS_TRUSTED", None)
            if is_trusted is False:
                stop = getattr(listener, "stop", None)
                if callable(stop):
                    try:
                        stop()
                    except (OSError, RuntimeError):
                        logger.warning("Arrêt listener non trusted échoué", exc_info=True)
                return (
                    "Permissions macOS manquantes "
                    "(Confidentialité → Accessibilité et Surveillance des entrées)."
                )
        with self._lock:
            self._listener = listener
        return None

    def stop(self) -> None:
        with self._lock:
            listener = self._listener
            self._listener = None
        if listener is None:
            return
        stop = getattr(listener, "stop", None)
        if callable(stop):
            try:
                stop()
            except (OSError, RuntimeError):
                logger.warning("Arrêt hotkeys globaux en échec", exc_info=True)
        join = getattr(listener, "join", None)
        if callable(join):
            try:
                join(timeout=1.0)
            except (OSError, RuntimeError):
                logger.warning("Join hotkeys globaux en échec", exc_info=True)
