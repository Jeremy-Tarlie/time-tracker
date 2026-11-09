"""Thème light centralisé (Tkinter) — slate clair + accent teal."""

from __future__ import annotations

import tkinter as tk
from collections.abc import Callable
from dataclasses import dataclass
from tkinter import ttk


@dataclass(frozen=True)
class Theme:
    bg: str = "#F4F6F8"
    surface: str = "#FFFFFF"
    field: str = "#EEF1F5"
    text: str = "#1C2430"
    muted: str = "#5C6B7A"
    border: str = "#D5DCE5"
    accent: str = "#0D9488"
    accent_dim: str = "#0F766E"
    danger: str = "#E11D48"
    danger_dim: str = "#BE123C"
    warning: str = "#D97706"
    warning_dim: str = "#B45309"
    timer: str = "#0F172A"
    on_accent: str = "#FFFFFF"
    on_solid: str = "#FFFFFF"
    success: str = "#059669"
    pill_ok_bg: str = "#D1FAE5"
    pill_paused_bg: str = "#FEF3C7"
    pill_error_bg: str = "#FFE4E6"
    font_ui: str = "Segoe UI"
    font_mono: str = "Consolas"


THEME = Theme()


def relative_luminance(hex_color: str) -> float:
    raw = hex_color.lstrip("#")
    if len(raw) != 6:
        raise ValueError(f"Couleur invalide: {hex_color}")
    channels = [int(raw[i : i + 2], 16) / 255.0 for i in (0, 2, 4)]

    def linearize(channel: float) -> float:
        return channel / 12.92 if channel <= 0.04045 else ((channel + 0.055) / 1.055) ** 2.4

    r, g, b = (linearize(c) for c in channels)
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def contrast_ratio(foreground: str, background: str) -> float:
    lighter = max(relative_luminance(foreground), relative_luminance(background))
    darker = min(relative_luminance(foreground), relative_luminance(background))
    return (lighter + 0.05) / (darker + 0.05)


def font_ui(size: int, *, bold: bool = False) -> tuple[str, int] | tuple[str, int, str]:
    if bold:
        return (THEME.font_ui, size, "bold")
    return (THEME.font_ui, size)


def font_mono(size: int, *, bold: bool = False) -> tuple[str, int] | tuple[str, int, str]:
    if bold:
        return (THEME.font_mono, size, "bold")
    return (THEME.font_mono, size)


def apply_window(root: tk.Tk | tk.Toplevel | tk.Frame, theme: Theme = THEME) -> None:
    root.configure(bg=theme.bg)


def apply_ttk_theme(style: ttk.Style | None = None, theme: Theme = THEME) -> ttk.Style:
    ttk_style = style or ttk.Style()
    try:
        ttk_style.theme_use("clam")
    except tk.TclError:
        ttk_style.theme_use("default")
    ttk_style.configure(
        "App.TCombobox",
        fieldbackground=theme.field,
        background=theme.field,
        foreground=theme.text,
        selectbackground=theme.accent,
        selectforeground=theme.on_accent,
        arrowcolor=theme.muted,
        bordercolor=theme.border,
        lightcolor=theme.border,
        darkcolor=theme.border,
        padding=5,
    )
    ttk_style.map(
        "App.TCombobox",
        fieldbackground=[("readonly", theme.field), ("disabled", theme.surface)],
        foreground=[("disabled", theme.muted)],
        background=[("readonly", theme.field)],
    )
    return ttk_style


# Alias rétrocompat pour imports existants
apply_ttk_dark = apply_ttk_theme


def make_button(
    parent: tk.Misc,
    text: str,
    command: Callable[[], None],
    *,
    variant: str = "primary",
    theme: Theme = THEME,
    padx: int = 12,
    pady: int = 9,
    font_size: int = 10,
    bold: bool = True,
) -> tk.Button:
    colors = {
        "primary": (theme.accent, theme.on_accent, theme.accent_dim),
        "danger": (theme.danger, theme.on_solid, theme.danger_dim),
        "warning": (theme.warning, theme.on_solid, theme.warning_dim),
        "subtle": (theme.field, theme.text, theme.border),
        "ghost": (theme.bg, theme.muted, theme.field),
        "link": (theme.bg, theme.muted, theme.bg),
    }
    bg, fg, active = colors.get(variant, colors["primary"])
    if variant == "primary":
        active_fg = theme.on_accent
    elif variant in {"danger", "warning"}:
        active_fg = theme.on_solid
    else:
        active_fg = theme.accent
    button = tk.Button(
        parent,
        text=text,
        command=command,
        bg=bg,
        fg=fg,
        activebackground=active,
        activeforeground=active_fg if variant != "link" else theme.accent,
        bd=0,
        highlightthickness=0,
        relief="flat",
        cursor="hand2",
        padx=padx,
        pady=pady,
        font=font_ui(font_size, bold=bold),
        disabledforeground=theme.muted,
    )
    if variant == "link":
        button.configure(anchor="w", font=font_ui(font_size, bold=False))
    return button


def dialog_header(parent: tk.Misc, title: str, theme: Theme = THEME) -> tk.Frame:
    wrap = tk.Frame(parent, bg=theme.bg)
    wrap.pack(fill="x")
    hdr = tk.Frame(wrap, bg=theme.surface)
    hdr.pack(fill="x")
    accent_bar = tk.Frame(hdr, bg=theme.accent, width=4)
    accent_bar.pack(side="left", fill="y")
    tk.Label(
        hdr,
        text=title,
        bg=theme.surface,
        fg=theme.text,
        font=font_ui(12, bold=True),
    ).pack(side="left", padx=14, pady=12)
    tk.Frame(wrap, bg=theme.border, height=1).pack(fill="x")
    return hdr


def section_label(parent: tk.Misc, text: str, theme: Theme = THEME) -> tk.Label:
    label = tk.Label(
        parent,
        text=text.upper(),
        bg=theme.bg,
        fg=theme.muted,
        font=font_ui(8, bold=True),
    )
    return label


def hairline(parent: tk.Misc, theme: Theme = THEME) -> tk.Frame:
    line = tk.Frame(parent, bg=theme.border, height=1)
    line.pack(fill="x")
    return line


def style_entry(entry: tk.Entry, theme: Theme = THEME) -> None:
    entry.configure(
        bg=theme.field,
        fg=theme.text,
        insertbackground=theme.text,
        relief="flat",
        bd=0,
        highlightthickness=1,
        highlightbackground=theme.border,
        highlightcolor=theme.accent,
    )


def style_checkbutton(button: tk.Checkbutton, theme: Theme = THEME) -> None:
    button.configure(
        bg=theme.bg,
        fg=theme.text,
        activebackground=theme.bg,
        activeforeground=theme.text,
        selectcolor=theme.surface,
        highlightthickness=0,
        font=font_ui(9),
    )


def surface_card(
    parent: tk.Misc,
    theme: Theme = THEME,
    *,
    padx: int = 14,
    pady: int = 12,
) -> tuple[tk.Frame, tk.Frame]:
    """Retourne (outer, inner) : carte blanche bordée dans le fond page."""
    outer = tk.Frame(parent, bg=theme.bg)
    card = tk.Frame(
        outer,
        bg=theme.surface,
        highlightthickness=1,
        highlightbackground=theme.border,
    )
    card.pack(fill="both", expand=True)
    inner = tk.Frame(card, bg=theme.surface, padx=padx, pady=pady)
    inner.pack(fill="both", expand=True)
    return outer, inner


def style_session_listbox(listbox: tk.Listbox, theme: Theme = THEME) -> None:
    listbox.configure(
        font=font_ui(9),
        bg=theme.field,
        fg=theme.text,
        selectbackground=theme.accent,
        selectforeground=theme.on_accent,
        relief="flat",
        highlightthickness=0,
        activestyle="none",
        borderwidth=0,
        bd=0,
    )


def stat_pill(
    parent: tk.Misc,
    caption: str,
    value: str,
    *,
    theme: Theme = THEME,
    accent: bool = False,
) -> tk.Frame:
    bg = theme.pill_ok_bg if accent else theme.field
    fg_value = theme.success if accent else theme.text
    wrap = tk.Frame(parent, bg=bg, padx=12, pady=8)
    tk.Label(
        wrap,
        text=caption.upper(),
        bg=bg,
        fg=theme.muted,
        font=font_ui(7, bold=True),
        anchor="w",
    ).pack(fill="x")
    tk.Label(
        wrap,
        text=value,
        bg=bg,
        fg=fg_value,
        font=font_ui(11, bold=True),
        anchor="w",
    ).pack(fill="x", pady=(2, 0))
    return wrap
