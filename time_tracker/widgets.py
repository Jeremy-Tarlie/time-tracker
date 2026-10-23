from __future__ import annotations

import tkinter as tk


class SearchableDropdown(tk.Frame):
    """Select avec recherche integree dans le panneau deroulant."""

    POPUP_LIST_HEIGHT = 8

    def __init__(
        self,
        parent: tk.Misc,
        textvariable: tk.StringVar,
        *,
        bg: str,
        field_bg: str,
        border: str,
        muted: str,
        accent: str,
        highlight: str,
        text: str,
        empty_label: str = "Aucun projet",
        placeholder: str = "Choisir un projet…",
    ) -> None:
        super().__init__(parent, bg=bg)
        self.variable = textvariable
        self._bg = bg
        self._field_bg = field_bg
        self._border = border
        self._muted = muted
        self._accent = accent
        self._highlight = highlight
        self._text = text
        self._empty_label = empty_label
        self._placeholder = placeholder
        self._items: list[str] = []
        self._popup: tk.Toplevel | None = None
        self._listbox: tk.Listbox | None = None
        self._search_var = tk.StringVar()
        self._search_trace_id: str | None = None
        self._click_bind_id: str | None = None
        self._open = False
        self._updating_list = False
        self._outside_click_armed = False

        shell = tk.Frame(self, bg=border, padx=1, pady=1)
        shell.pack(fill="x")
        self.trigger = tk.Frame(shell, bg=field_bg, cursor="hand2")
        self.trigger.pack(fill="x")

        self.value_label = tk.Label(
            self.trigger, text="", bg=field_bg, fg=text, anchor="w",
            font=("Segoe UI", 10), padx=10, pady=8,
        )
        self.value_label.pack(side="left", fill="x", expand=True)
        caret = tk.Label(
            self.trigger, text="▼", bg=field_bg, fg=muted,
            font=("Segoe UI", 8),
        )
        caret.pack(side="right", padx=(0, 10))

        for widget in (self.trigger, self.value_label, caret):
            widget.bind("<Button-1>", self._toggle)

        self.variable.trace_add("write", lambda *_: self._update_display())
        self._update_display()

    def _update_display(self) -> None:
        value = self.variable.get().strip()
        if value:
            self.value_label.config(text=value, fg=self._text)
        else:
            self.value_label.config(text=self._placeholder, fg=self._muted)

    def set_options(self, options: list[str]) -> None:
        self._items = list(options)
        current = self.variable.get()
        if current and current not in self._items:
            if self._items:
                self.variable.set(self._items[0])
            else:
                self.variable.set("")
        elif not current and self._items:
            self.variable.set(self._items[0])
        if self._open:
            self._refresh_list()

    def close_if_open(self) -> None:
        if self._open:
            self._close()

    def _toggle(self, _event: tk.Event[tk.Misc] | None = None) -> None:
        if self._open:
            self._close()
        else:
            self._open_popup()

    def _open_popup(self) -> None:
        if not self._items or self._open:
            return

        self._close()
        self._popup = tk.Toplevel(self)
        self._popup.overrideredirect(True)
        self._popup.configure(bg=self._border)
        self._popup.attributes("-topmost", True)

        width = max(self.winfo_width(), 260)
        popup_height = 44 + self.POPUP_LIST_HEIGHT * 24 + 8
        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height() + 2
        self._popup.geometry(f"{width}x{popup_height}+{x}+{y}")

        container = tk.Frame(self._popup, bg=self._bg, padx=1, pady=1)
        container.pack(fill="both", expand=True)

        search_wrap = tk.Frame(container, bg=self._field_bg)
        search_wrap.pack(fill="x")
        self._search_var.set("")
        search_entry = tk.Entry(
            search_wrap, textvariable=self._search_var,
            font=("Segoe UI", 10), bg=self._field_bg, fg=self._text,
            relief="flat", bd=0, highlightthickness=1,
            highlightbackground=self._border, highlightcolor=self._highlight,
        )
        search_entry.pack(fill="x", padx=8, pady=8, ipady=3)
        search_entry.bind("<Escape>", lambda _event: self._close())
        search_entry.bind("<Return>", self._select_first)
        search_entry.bind("<Down>", lambda _event: self._focus_list())

        list_shell = tk.Frame(container, bg=self._border, padx=1)
        list_shell.pack(fill="both", expand=True, pady=(0, 1))
        list_frame = tk.Frame(list_shell, bg=self._bg)
        list_frame.pack(fill="both", expand=True)

        scrollbar = tk.Scrollbar(list_frame, orient="vertical")
        self._listbox = tk.Listbox(
            list_frame, yscrollcommand=scrollbar.set,
            font=("Segoe UI", 10), bg=self._bg, fg=self._text,
            selectbackground=self._highlight, selectforeground="white",
            activestyle="none", relief="flat", bd=0,
            highlightthickness=0, height=self.POPUP_LIST_HEIGHT,
        )
        scrollbar.config(command=self._listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self._listbox.pack(side="left", fill="both", expand=True)

        self._listbox.bind("<Double-Button-1>", self._on_list_select)
        self._listbox.bind("<Return>", self._on_list_select)
        self._listbox.bind("<Escape>", lambda _event: self._close())
        self._search_trace_id = self._search_var.trace_add(
            "write", lambda *_: self._refresh_list(),
        )

        self._open = True
        self._outside_click_armed = False
        self._refresh_list()
        search_entry.focus_set()

        root = self.winfo_toplevel()
        self.after(120, lambda: self._arm_outside_click(root))

    def _arm_outside_click(self, root: tk.Misc) -> None:
        if not self._open:
            return
        self._outside_click_armed = True
        self._click_bind_id = root.bind("<Button-1>", self._on_outside_click, add="+")
        if self._popup is not None and self._popup.winfo_exists():
            self._popup.bind("<Escape>", lambda _event: self._close())

    def _filtered_options(self) -> list[str]:
        query = self._search_var.get().casefold().strip()
        if not query:
            return self._items
        return [name for name in self._items if query in name.casefold()]

    def _listbox_alive(self) -> bool:
        return self._listbox is not None and self._listbox.winfo_exists()

    def _refresh_list(self) -> None:
        if not self._open or not self._listbox_alive() or self._listbox is None:
            return
        filtered = self._filtered_options()
        current = self.variable.get()
        self._updating_list = True
        self._listbox.delete(0, "end")
        if not filtered:
            self._listbox.insert("end", self._empty_label)
            self._listbox.config(state="disabled")
            self._updating_list = False
            return
        self._listbox.config(state="normal")
        for name in filtered:
            self._listbox.insert("end", name)
        for index, name in enumerate(filtered):
            if name == current:
                self._listbox.selection_set(index)
                self._listbox.see(index)
                break
        self._updating_list = False

    def _focus_list(self) -> None:
        if not self._listbox_alive() or self._listbox is None:
            return
        if self._listbox.size() > 0 and str(self._listbox.get(0)) != self._empty_label:
            self._listbox.focus_set()
            if not self._listbox.curselection():
                self._listbox.selection_set(0)

    def _select_first(self, _event: tk.Event[tk.Misc] | None = None) -> None:
        filtered = self._filtered_options()
        if filtered:
            self.variable.set(filtered[0])
        self._close()

    def _on_list_select(self, _event: tk.Event[tk.Misc] | None = None) -> None:
        if self._updating_list or not self._listbox_alive() or self._listbox is None:
            return
        if self._listbox.cget("state") == "disabled":
            return
        selection = self._listbox.curselection()
        if not selection:
            return
        value = self._listbox.get(selection[0])
        if value == self._empty_label:
            return
        self.variable.set(str(value))
        self._close()

    def _on_outside_click(self, event: tk.Event[tk.Misc]) -> None:
        if not self._open or not self._outside_click_armed or self._popup is None:
            return
        if self._contains(self.trigger, event.x_root, event.y_root):
            return
        if self._contains(self._popup, event.x_root, event.y_root):
            return
        self._close()

    def _contains(self, widget: tk.Misc, x_root: int, y_root: int) -> bool:
        if not widget.winfo_exists():
            return False
        x = widget.winfo_rootx()
        y = widget.winfo_rooty()
        return x <= x_root <= x + widget.winfo_width() and y <= y_root <= y + widget.winfo_height()

    def _close(self) -> None:
        if self._search_trace_id is not None:
            self._search_var.trace_remove("write", self._search_trace_id)
            self._search_trace_id = None
        if self._click_bind_id is not None:
            self.winfo_toplevel().unbind("<Button-1>", self._click_bind_id)
            self._click_bind_id = None
        if self._popup is not None and self._popup.winfo_exists():
            self._popup.destroy()
        self._popup = None
        self._listbox = None
        self._open = False
        self._outside_click_armed = False
