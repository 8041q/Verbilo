"""Reusable dropdown controls for the Verbilo desktop UI.

The application has two kinds of selects:

* :class:`SelectDropdown` is a read-only select control.
* :class:`SearchableDropdown` is an editable select whose entry text filters
  the available options but does not change the committed value until the user
  confirms a real option.

Both controls intentionally share one popup implementation.  This keeps focus,
outside-click handling, keyboard navigation, row hover, popup positioning and
state changes identical throughout the application.
"""

from __future__ import annotations

from collections.abc import Callable, Iterable, Sequence
import tkinter as tk
from tkinter import ttk

import customtkinter as ctk

from . import theme
from .icons import get_icon


_POPUP_ROWS = 8
_UNSET = object()


def filter_dropdown_values(values: Sequence[str], query: str) -> list[str]:
    """Return ``values`` containing *query* using Unicode-aware matching.

    Order is preserved.  ``casefold`` is used rather than ``lower`` so search
    behaves sensibly for a wider range of localized language/model names.
    """

    needle = (query or "").strip().casefold()
    if not needle:
        return list(values)
    return [value for value in values if needle in value.casefold()]


class _DropdownBase:
    """Shared popup, state and keyboard plumbing for application dropdowns."""

    _POPUP_ROWS = _POPUP_ROWS

    def __init__(
        self,
        parent,
        values: Iterable[str],
        variable,
        *,
        command: Callable[[str], None] | None = None,
    ) -> None:
        self._parent = parent
        self._values = list(values)
        self._variable = variable
        self._command = command
        initial = variable.get() if variable is not None else ""
        self._committed = initial if initial in self._values else (self._values[0] if self._values else "")

        self._popup: tk.Toplevel | None = None
        self._listbox: tk.Listbox | None = None
        self._scrollbar = None
        self._popup_items: list[str] = []
        self._hover_index: int | None = None
        self._popup_pointer_down = False
        self._disabled = False
        self._writing_external = False
        self._owner_click_binding: str | None = None
        self._external_trace: str | None = None

        p = theme.get()
        self._frame = ctk.CTkFrame(
            parent,
            fg_color=p.bg_input,
            corner_radius=theme.BUTTON_CORNER_RADIUS,
            border_width=0,
            border_color=p.border,
        )
        self._frame.grid_columnconfigure(0, weight=1)
        self._frame.bind("<Destroy>", self._on_frame_destroy, "+")

        self._build_control()
        self._install_owner_click_binding()
        self._install_external_variable_trace()
        self._display_committed_value()

    # -- hooks ---------------------------------------------------------

    def _build_control(self) -> None:  # pragma: no cover - abstract hook
        raise NotImplementedError

    def _display_committed_value(self) -> None:  # pragma: no cover - abstract hook
        raise NotImplementedError

    def _focus_target(self):  # pragma: no cover - abstract hook
        return self._frame

    def _outside_dismiss(self) -> None:
        self._close_popup()

    def _before_values_changed(self) -> None:
        pass

    # -- geometry passthrough -----------------------------------------

    def grid(self, **kwargs):
        return self._frame.grid(**kwargs)

    def pack(self, **kwargs):
        return self._frame.pack(**kwargs)

    def place(self, **kwargs):
        return self._frame.place(**kwargs)

    def grid_remove(self):
        return self._frame.grid_remove()

    def grid_forget(self):
        return self._frame.grid_forget()

    def focus_set(self):
        try:
            self._focus_target().focus_set()
        except Exception:
            pass

    def _restore_owner_activation(self, target=None) -> None:
        """Return OS/Tk keyboard ownership from a popup to its real owner.

        On Windows an ``overrideredirect`` Toplevel can become the active native
        window when its listbox is clicked.  Destroying that popup may leave the
        process with no active keyboard window: mouse clicks still work, but Entry
        widgets do not receive characters until another Toplevel is opened.  A
        normal ``focus_set()`` only changes Tk's logical focus and is not enough to
        repair that native activation state.

        This method is only used after an explicit popup keyboard/mouse action, so
        forcing the owner active on Win32 cannot steal focus from an unrelated
        application event.
        """
        try:
            owner = self._frame.winfo_toplevel()
        except Exception:
            return
        focus_target = target if target is not None else self._focus_target()

        def _restore() -> None:
            try:
                if not owner.winfo_exists():
                    return
            except Exception:
                return
            try:
                if self._frame.tk.call("tk", "windowingsystem") == "win32":
                    owner.focus_force()
            except Exception:
                pass
            try:
                focus_target.focus_set()
            except Exception:
                pass

        # Restore immediately, then once more after Tk has finished the popup's
        # ButtonRelease/Return transaction.  The second pass is important on
        # Windows where native activation can settle after the Tk callback returns.
        _restore()
        try:
            owner.after_idle(_restore)
        except Exception:
            pass

    # -- public state API ---------------------------------------------

    def get(self) -> str:
        return self._committed

    def set(self, value: str) -> None:
        """Programmatically commit *value* when it is a valid option."""

        if value not in self._values:
            return
        self._commit(value, notify=False)

    def update_values(self, values: Iterable[str]) -> None:
        """Replace available choices while keeping a valid selection."""

        self._before_values_changed()
        self._values = list(values)
        self._close_popup()
        if self._committed not in self._values:
            fallback = self._values[0] if self._values else ""
            self._commit(fallback, notify=False, allow_empty=True)
        else:
            self._display_committed_value()

    def configure(self, **kwargs) -> None:
        """Configure the wrapper using combobox-compatible options.

        Engine/provider changes in the application often refresh a combo with
        ``configure(values=..., variable=..., state=...)``.  Treating those
        options as no-ops leaves this wrapper logically detached from the UI,
        so handle the stateful options here instead of silently discarding them.
        """

        values = kwargs.pop("values", _UNSET)
        if values is not _UNSET:
            self.update_values(values)

        variable = kwargs.pop("variable", _UNSET)
        textvariable = kwargs.pop("textvariable", _UNSET)
        if variable is _UNSET and textvariable is not _UNSET:
            variable = textvariable
        if variable is not _UNSET:
            self._replace_external_variable(variable)

        command = kwargs.pop("command", _UNSET)
        if command is not _UNSET:
            self._command = command

        state = kwargs.pop("state", _UNSET)
        if state is not _UNSET:
            self._set_disabled(str(state).lower() == "disabled")

        # Layout/appearance options are intentionally not forwarded into CTk
        # internals.  The important compatibility contract is the control's
        # values, variable, command and enabled state.

    config = configure

    def destroy(self) -> None:
        self._cleanup()
        try:
            self._frame.destroy()
        except Exception:
            pass

    # -- committed/external value ------------------------------------

    @property
    def values(self) -> tuple[str, ...]:
        return tuple(self._values)

    def _commit(
        self,
        value: str,
        *,
        notify: bool,
        allow_empty: bool = False,
    ) -> None:
        if not allow_empty and value not in self._values:
            return
        if value and value not in self._values:
            return

        self._committed = value
        self._display_committed_value()
        if self._variable is not None:
            try:
                self._writing_external = True
                self._variable.set(value)
            finally:
                self._writing_external = False
        if notify and self._command is not None:
            self._command(value)

    def _install_external_variable_trace(self) -> None:
        if self._variable is None or not hasattr(self._variable, "trace_add"):
            return
        try:
            self._external_trace = self._variable.trace_add("write", self._on_external_variable_changed)
        except Exception:
            self._external_trace = None

    def _replace_external_variable(self, variable) -> None:
        """Rebind the external Tk variable without leaving stale traces behind."""

        if variable is self._variable:
            return
        try:
            if self._external_trace and self._variable is not None:
                self._variable.trace_remove("write", self._external_trace)
        except Exception:
            pass

        self._external_trace = None
        self._variable = variable
        self._install_external_variable_trace()

        incoming = ""
        if variable is not None:
            try:
                incoming = variable.get()
            except Exception:
                incoming = ""

        if incoming in self._values:
            self._committed = incoming
            self._display_committed_value()
            self._close_popup()
            return

        # Keep the replacement variable synchronized with the control's valid
        # committed value.  This is particularly important when an engine swap
        # replaces both ``values`` and ``variable`` in the same refresh pass.
        if self._committed not in self._values:
            self._committed = self._values[0] if self._values else ""
        self._display_committed_value()
        if variable is not None:
            try:
                self._writing_external = True
                variable.set(self._committed)
            except Exception:
                pass
            finally:
                self._writing_external = False
        self._close_popup()

    def _on_external_variable_changed(self, *_args) -> None:
        if self._writing_external or self._variable is None:
            return
        try:
            value = self._variable.get()
        except Exception:
            return
        if value in self._values and value != self._committed:
            self._committed = value
            self._display_committed_value()
            self._close_popup()

    # -- enabled/focus visuals ---------------------------------------

    def _set_disabled(self, disabled: bool) -> None:
        self._disabled = disabled
        if disabled:
            self._close_popup()
        for widget in self._state_widgets():
            try:
                widget.configure(state="disabled" if disabled else "normal")
            except Exception:
                pass

    def _state_widgets(self) -> tuple:
        return ()

    def _set_focus_visual(self, focused: bool) -> None:
        p = theme.get()
        try:
            self._frame.configure(
                border_width=1 if focused else 0,
                border_color=p.accent if focused else p.border,
            )
        except Exception:
            pass

    # -- popup ---------------------------------------------------------

    def _open_popup(self, items: Sequence[str] | None = None, *, focus_list: bool = False) -> None:
        if self._disabled:
            return
        display_items = list(self._values if items is None else items)
        if not display_items:
            self._close_popup()
            return

        if self._popup is None or not self._popup.winfo_exists():
            self._create_popup()
        self._populate_popup(display_items)
        self._position_popup()
        try:
            self._popup.deiconify()
            self._popup.lift()
        except Exception:
            pass
        if focus_list and self._listbox is not None:
            try:
                self._listbox.focus_set()
            except Exception:
                pass

    def _create_popup(self) -> None:
        p = theme.get()
        owner = self._frame.winfo_toplevel()
        popup = tk.Toplevel(owner)
        popup.withdraw()
        popup.wm_overrideredirect(True)
        try:
            popup.transient(owner)
        except Exception:
            pass
        popup.configure(bg=p.bg_popup)
        self._popup = popup

        outer = tk.Frame(popup, bg=p.bg_popup, bd=0, highlightthickness=0)
        outer.pack(fill="both", expand=True)
        outer.grid_rowconfigure(0, weight=1)
        outer.grid_columnconfigure(0, weight=1)

        self._configure_scrollbar_style(p)
        listbox = tk.Listbox(
            outer,
            height=self._POPUP_ROWS,
            font=(theme.FONT_FAMILY, theme.FONT_BODY[1]),
            activestyle="none",
            exportselection=False,
            selectbackground=p.accent,
            selectforeground=p.text_on_accent,
            bg=p.bg_popup,
            fg=p.text_secondary,
            relief="flat",
            borderwidth=0,
            highlightthickness=0,
            cursor="hand2",
        )
        scrollbar = ttk.Scrollbar(
            outer,
            orient="vertical",
            command=listbox.yview,
            style="Dropdown.Vertical.TScrollbar",
        )
        listbox.configure(yscrollcommand=scrollbar.set)
        listbox.grid(row=0, column=0, sticky="nsew", padx=(4, 0), pady=4)
        self._listbox = listbox
        self._scrollbar = scrollbar

        listbox.bind("<ButtonPress-1>", self._on_popup_press, "+")
        listbox.bind("<ButtonRelease-1>", self._on_popup_click)
        listbox.bind("<Return>", self._on_popup_confirm)
        listbox.bind("<Escape>", self._on_popup_escape)
        listbox.bind("<Motion>", self._on_popup_motion, "+")
        listbox.bind("<Leave>", self._on_popup_leave, "+")
        listbox.bind("<MouseWheel>", self._on_popup_mousewheel, "+")
        listbox.bind("<Button-4>", lambda _e: self._scroll_popup(-1), "+")
        listbox.bind("<Button-5>", lambda _e: self._scroll_popup(1), "+")
        listbox.bind("<FocusOut>", lambda _e: self._schedule_focus_departure(), "+")

    @staticmethod
    def _configure_scrollbar_style(p) -> None:
        style = ttk.Style()
        style.configure(
            "Dropdown.Vertical.TScrollbar",
            gripcount=0,
            background=p.bg_popup,
            darkcolor=p.bg_popup,
            lightcolor=p.bg_popup,
            troughcolor=p.bg_popup,
            bordercolor=p.bg_popup,
            arrowcolor=p.text_muted,
            relief="flat",
            borderwidth=0,
            arrowsize=10,
            width=9,
        )
        style.map(
            "Dropdown.Vertical.TScrollbar",
            background=[("active", p.border), ("!active", p.divider)],
        )

    def _populate_popup(self, items: Sequence[str]) -> None:
        if self._listbox is None:
            return
        self._popup_items = list(items)
        self._hover_index = None
        self._listbox.delete(0, tk.END)
        for item in self._popup_items:
            self._listbox.insert(tk.END, item)

        selected_index: int | None = None
        if self._committed in self._popup_items:
            selected_index = self._popup_items.index(self._committed)
        elif self._popup_items:
            selected_index = 0

        self._listbox.selection_clear(0, tk.END)
        if selected_index is not None:
            self._listbox.selection_set(selected_index)
            self._listbox.activate(selected_index)
            self._listbox.see(selected_index)

        if self._scrollbar is not None:
            try:
                if len(self._popup_items) > self._POPUP_ROWS:
                    self._scrollbar.grid(row=0, column=1, sticky="ns", padx=(0, 2), pady=4)
                else:
                    self._scrollbar.grid_remove()
            except Exception:
                pass

    def _position_popup(self) -> None:
        if self._popup is None or self._listbox is None:
            return
        self._frame.update_idletasks()
        x = self._frame.winfo_rootx()
        frame_y = self._frame.winfo_rooty()
        frame_h = max(1, self._frame.winfo_height())
        width = max(80, self._frame.winfo_width())
        rows = min(self._POPUP_ROWS, max(1, len(self._popup_items)))
        row_px = max(22, theme.scale(theme.FONT_BODY[1] + 10))
        height = rows * row_px + 8

        owner = self._frame.winfo_toplevel()
        try:
            left = owner.winfo_vrootx()
            top = owner.winfo_vrooty()
            right = left + owner.winfo_vrootwidth()
            bottom = top + owner.winfo_vrootheight()
        except Exception:
            left = top = 0
            right = owner.winfo_screenwidth()
            bottom = owner.winfo_screenheight()

        below_y = frame_y + frame_h + 2
        above_y = frame_y - height - 2
        if below_y + height <= bottom or above_y < top:
            y = min(below_y, max(top, bottom - height))
        else:
            y = max(top, above_y)
        x = max(left, min(x, max(left, right - width)))
        self._popup.geometry(f"{width}x{height}+{x}+{y}")

    def _close_popup(self) -> None:
        popup = self._popup
        self._popup = None
        self._listbox = None
        self._scrollbar = None
        self._popup_items = []
        self._hover_index = None
        self._popup_pointer_down = False
        if popup is not None:
            try:
                if popup.winfo_exists():
                    popup.destroy()
            except Exception:
                pass

    def _popup_is_open(self) -> bool:
        try:
            return bool(self._popup is not None and self._popup.winfo_exists())
        except Exception:
            return False

    # -- popup events --------------------------------------------------

    def _on_popup_press(self, _event=None):
        """Mark the popup as actively receiving a pointer click.

        On Windows an override-redirect listbox does not always own keyboard
        focus before the entry receives FocusOut.  Tracking the pointer press
        prevents a focus validator from tearing the popup down between mouse
        press and release.
        """

        self._popup_pointer_down = True
        # Pointer selection does not require keyboard focus.  In particular, do
        # not focus the Listbox here: on Windows that can activate the separate
        # overrideredirect Toplevel and leave the main application without native
        # keyboard ownership when the popup is destroyed.
        return None

    def _on_popup_click(self, _event=None):
        self._popup_pointer_down = False
        self._confirm_popup_selection()
        return "break"

    def _on_popup_confirm(self, _event=None):
        self._confirm_popup_selection()
        return "break"

    def _on_popup_escape(self, _event=None):
        self._outside_dismiss()
        self._restore_owner_activation()
        return "break"

    def _confirm_popup_selection(self) -> None:
        if self._listbox is None:
            return
        selection = self._listbox.curselection()
        if not selection:
            return
        value = self._listbox.get(selection[0])

        # Finish the popup interaction before any application callback runs.
        # Engine/detector callbacks can rebuild other dropdowns; running them while
        # this listbox still exists (and often owns focus on Windows) creates a
        # re-entrant focus transaction that can leave editable entries unable to
        # receive keyboard input afterwards.
        self._close_popup()
        self._commit(value, notify=False)
        self._after_user_commit()
        if self._command is not None:
            self._command(value)

    def _after_user_commit(self) -> None:
        pass

    def _on_popup_motion(self, event) -> None:
        if self._listbox is None or not self._popup_items:
            return
        try:
            index = self._listbox.nearest(event.y)
            bbox = self._listbox.bbox(index)
            if bbox is None or not (bbox[1] <= event.y < bbox[1] + bbox[3]):
                index = None
        except Exception:
            index = None
        if index == self._hover_index:
            return
        self._clear_hover_row()
        if index is None:
            return
        p = theme.get()
        try:
            self._listbox.itemconfigure(index, background=p.bg_heading, foreground=p.text_primary)
            self._hover_index = index
        except Exception:
            self._hover_index = None

    def _on_popup_leave(self, _event=None) -> None:
        self._clear_hover_row()

    def _clear_hover_row(self) -> None:
        if self._listbox is None or self._hover_index is None:
            self._hover_index = None
            return
        p = theme.get()
        try:
            self._listbox.itemconfigure(
                self._hover_index,
                background=p.bg_popup,
                foreground=p.text_secondary,
            )
        except Exception:
            pass
        self._hover_index = None

    def _on_popup_mousewheel(self, event):
        delta = getattr(event, "delta", 0)
        if not delta:
            return None
        self._scroll_popup(-1 if delta > 0 else 1)
        return "break"

    def _scroll_popup(self, units: int):
        if self._listbox is not None:
            try:
                self._listbox.yview_scroll(units, "units")
            except Exception:
                pass
        return "break"

    # -- focus/outside click ------------------------------------------

    def _install_owner_click_binding(self) -> None:
        try:
            owner = self._frame.winfo_toplevel()
            self._owner_click_binding = owner.bind("<Button-1>", self._on_owner_click, "+")
        except Exception:
            self._owner_click_binding = None

    def _on_owner_click(self, event) -> None:
        if not self._popup_is_open():
            return
        if self._widget_is_inside(event.widget, self._frame):
            return
        if self._popup is not None and self._widget_is_inside(event.widget, self._popup):
            return
        self._outside_dismiss()

    @staticmethod
    def _widget_is_inside(widget, container) -> bool:
        current = widget
        while current is not None:
            if current is container:
                return True
            current = getattr(current, "master", None)
        return False

    def _schedule_focus_departure(self) -> None:
        """Schedule a focus ownership check.

        Read-only selects can check on the next idle turn.  Searchable selects
        override this with a short debounce because their editable entry and
        popup are separate native Tk windows on Windows.
        """

        try:
            self._frame.after_idle(self._check_focus_departure)
        except Exception:
            self._check_focus_departure()

    def _check_focus_departure(self) -> None:
        if not self._popup_is_open():
            return
        if self._popup_pointer_down:
            # Do not close a popup between ButtonPress and ButtonRelease.
            # Re-check shortly if focus still has not settled.
            try:
                self._frame.after(80, self._check_focus_departure)
            except Exception:
                pass
            return
        try:
            focused = self._frame.focus_get()
        except Exception:
            focused = None
        if focused is not None:
            if self._widget_is_inside(focused, self._frame):
                return
            if self._popup is not None and self._widget_is_inside(focused, self._popup):
                return
        self._outside_dismiss()

    # -- cleanup -------------------------------------------------------

    def _on_frame_destroy(self, event) -> None:
        if event.widget is self._frame:
            self._cleanup()

    def _cleanup(self) -> None:
        self._close_popup()
        try:
            if self._external_trace and self._variable is not None:
                self._variable.trace_remove("write", self._external_trace)
        except Exception:
            pass
        self._external_trace = None
        try:
            owner = self._frame.winfo_toplevel()
            if self._owner_click_binding:
                owner.unbind("<Button-1>", self._owner_click_binding)
        except Exception:
            pass
        self._owner_click_binding = None


class SelectDropdown(_DropdownBase):
    """Read-only dropdown with standard pointer and keyboard interaction."""

    def _build_control(self) -> None:
        p = theme.get()
        font = ctk.CTkFont(family=theme.FONT_FAMILY, size=theme.FONT_BODY[1])
        self._label_var = tk.StringVar(value=self._committed)
        self._label = ctk.CTkLabel(
            self._frame,
            textvariable=self._label_var,
            font=font,
            text_color=p.text_secondary,
            anchor="w",
            fg_color="transparent",
            height=32,
            cursor="hand2",
        )
        self._label.grid(row=0, column=0, sticky="ew", padx=(10, 0))

        arrow = get_icon("chevron-down", size=14)
        button_kwargs = dict(
            master=self._frame,
            width=28,
            height=28,
            fg_color="transparent",
            hover_color=p.bg_heading,
            corner_radius=4,
            command=self._toggle_popup,
            border_width=0,
            cursor="hand2",
        )
        if arrow:
            self._button = ctk.CTkButton(text="", image=arrow, **button_kwargs)
        else:
            self._button = ctk.CTkButton(
                text="\u25bc",
                font=ctk.CTkFont(family=theme.FONT_FAMILY, size=10),
                text_color=p.text_muted,
                **button_kwargs,
            )
        self._button.grid(row=0, column=1, padx=(0, 2), pady=2)
        try:
            self._button.configure(takefocus=True)
        except Exception:
            pass

        self._label.bind("<Button-1>", lambda _e: self._toggle_popup())
        self._label.bind("<Enter>", lambda _e: self._set_control_hover(True), "+")
        self._label.bind("<Leave>", lambda _e: self._set_control_hover(False), "+")
        self._button.bind("<FocusIn>", lambda _e: self._set_focus_visual(True), "+")
        self._button.bind("<FocusOut>", lambda _e: self._set_focus_visual(False), "+")
        self._button.bind("<Down>", self._keyboard_open, "+")
        self._button.bind("<Up>", self._keyboard_open, "+")
        self._button.bind("<Escape>", self._keyboard_escape, "+")

    def _display_committed_value(self) -> None:
        if hasattr(self, "_label_var"):
            self._label_var.set(self._committed)

    def _focus_target(self):
        return self._button

    def _state_widgets(self) -> tuple:
        return (self._button,)

    def _set_control_hover(self, hovered: bool) -> None:
        if self._disabled:
            return
        p = theme.get()
        try:
            self._frame.configure(fg_color=p.bg_heading if hovered else p.bg_input)
        except Exception:
            pass

    def _toggle_popup(self) -> None:
        if self._disabled:
            return
        if self._popup_is_open():
            self._close_popup()
            return
        # Mouse-opened read-only dropdowns must not transfer keyboard focus into
        # the popup Toplevel.  Keyboard-opened dropdowns still focus the Listbox so
        # arrow/Return navigation continues to work.
        self._open_popup(focus_list=False)

    def _keyboard_open(self, _event=None):
        if not self._popup_is_open():
            self._open_popup(focus_list=True)
        return "break"

    def _keyboard_escape(self, _event=None):
        self._close_popup()
        self._restore_owner_activation(self._button)
        return "break"

    def _after_user_commit(self) -> None:
        self._restore_owner_activation(self._button)


class SearchableDropdown(_DropdownBase):
    """Editable dropdown whose popup never owns normal text-entry focus.

    The control uses an explicit interaction model rather than inferring user
    intent from FocusOut timing:

    * clicking the field focuses the real native ``tk.Entry`` and opens choices;
    * the first printable key replaces the committed label, then typing is normal;
    * changing the query only filters/repaints the popup and never refocuses it;
    * row selection/Enter commits, Escape or an outside click validates/reverts;
    * engine/provider refreshes may replace values/variables without rebuilding
      or disabling the editable entry.

    This avoids Windows focus races between ``CTkEntry`` and the separate
    override-redirect ``Toplevel`` used for the popup.
    """

    def __init__(self, parent, values, variable, command=None, **_kwargs) -> None:
        self._query_guard = False
        self._query_trace: str | None = None
        self._replace_on_type = True
        super().__init__(parent, values, variable, command=command)
        self._install_query_trace()

    def _build_control(self) -> None:
        p = theme.get()
        font = ctk.CTkFont(family=theme.FONT_FAMILY, size=theme.FONT_BODY[1])
        self._query_var = tk.StringVar(value=self._committed)
        self._entry = ctk.CTkEntry(
            self._frame,
            textvariable=self._query_var,
            fg_color="transparent",
            border_width=0,
            font=font,
            text_color=p.text_secondary,
            height=32,
        )
        self._entry.grid(row=0, column=0, sticky="ew", padx=(6, 0))

        arrow = get_icon("chevron-down", size=14)
        button_kwargs = dict(
            master=self._frame,
            width=28,
            height=28,
            fg_color="transparent",
            hover_color=p.bg_heading,
            corner_radius=4,
            command=self._toggle_popup,
            border_width=0,
            cursor="hand2",
        )
        if arrow:
            self._button = ctk.CTkButton(text="", image=arrow, **button_kwargs)
        else:
            self._button = ctk.CTkButton(
                text="\u25bc",
                font=ctk.CTkFont(family=theme.FONT_FAMILY, size=10),
                text_color=p.text_muted,
                **button_kwargs,
            )
        try:
            self._button.configure(takefocus=False)
        except Exception:
            pass
        self._button.grid(row=0, column=1, padx=(0, 2), pady=2)

        # CTkEntry is a wrapper.  All keyboard editing must target the native
        # tk.Entry, not the surrounding canvas/frame.
        self._tk_entry = self._entry
        try:
            native_entry = getattr(self._entry, "_entry", None)
            if isinstance(native_entry, tk.Entry):
                self._tk_entry = native_entry
            else:
                for child in self._entry.winfo_children():
                    if isinstance(child, tk.Entry):
                        self._tk_entry = child
                        break
        except Exception:
            pass

        self._tk_entry.bind("<FocusIn>", self._on_entry_focus_in, "+")
        self._tk_entry.bind("<FocusOut>", self._on_entry_focus_out, "+")
        self._tk_entry.bind("<Button-1>", self._on_entry_click, "+")
        self._entry.bind("<Button-1>", self._on_entry_click, "+")
        self._tk_entry.bind("<KeyPress>", self._on_entry_keypress, "+")
        self._tk_entry.bind("<<Paste>>", self._on_entry_paste, "+")
        self._tk_entry.bind("<Down>", self._focus_popup_list, "+")
        self._tk_entry.bind("<Up>", self._focus_popup_list, "+")
        self._tk_entry.bind("<Return>", self._on_entry_return, "+")
        self._tk_entry.bind("<Escape>", self._on_entry_escape, "+")

    def _install_query_trace(self) -> None:
        try:
            self._query_trace = self._query_var.trace_add("write", self._on_query_changed)
        except Exception:
            self._query_trace = None

    def _display_committed_value(self) -> None:
        if hasattr(self, "_query_var"):
            self._set_query_text(self._committed)
            self._replace_on_type = True

    def _before_values_changed(self) -> None:
        self._replace_on_type = True

    def update_values(self, values: Iterable[str]) -> None:
        """Refresh choices and leave the field ready for immediate editing."""
        super().update_values(values)
        self._replace_on_type = True
        self._ensure_entry_state()

    def configure(self, **kwargs) -> None:
        # Base handles combobox-compatible values/variable/command/state.
        super().configure(**kwargs)
        self._ensure_entry_state()

    config = configure

    def _set_disabled(self, disabled: bool) -> None:
        super()._set_disabled(disabled)
        self._ensure_entry_state()

    def _ensure_entry_state(self) -> None:
        """Keep CTkEntry and its native tk.Entry in exactly the same state."""
        state = "disabled" if self._disabled else "normal"
        try:
            self._entry.configure(state=state)
        except Exception:
            pass
        try:
            self._tk_entry.configure(state=state)
        except Exception:
            pass

    def _focus_target(self):
        return self._tk_entry

    def _state_widgets(self) -> tuple:
        return (self._entry, self._button)

    def _set_query_text(self, text: str) -> None:
        self._query_guard = True
        try:
            self._query_var.set(text)
        finally:
            self._query_guard = False

    def _select_all(self) -> None:
        try:
            self._tk_entry.select_range(0, tk.END)
            self._tk_entry.icursor(tk.END)
        except Exception:
            pass

    def _open_for_current_query(self) -> None:
        if self._disabled:
            return
        current = self._query_var.get()
        if current == self._committed:
            items = self._values
        else:
            matches = filter_dropdown_values(self._values, current)
            items = matches if matches else self._values
        self._open_popup(items, focus_list=False)

    # -- entry events --------------------------------------------------

    def _on_entry_focus_in(self, _event=None):
        self._set_focus_visual(True)

    def _on_entry_focus_out(self, _event=None):
        # Focus is only visual state.  Do not close/validate from FocusOut: the
        # popup is a separate native window and Windows may report transient
        # focus changes while the pointer merely moves over it.
        self._set_focus_visual(False)

    def _on_entry_click(self, _event=None):
        if self._disabled:
            return None
        self._ensure_entry_state()
        try:
            self._tk_entry.focus_set()
        except Exception:
            pass

        if self._query_var.get() == self._committed:
            self._replace_on_type = True
            # Selection is only a visual convenience.  First-key replacement is
            # enforced explicitly in _on_entry_keypress, so a later Tk class
            # binding cannot break the overwrite behavior.
            try:
                self._tk_entry.after_idle(self._select_all)
            except Exception:
                self._select_all()
        self._open_for_current_query()
        return None

    def _on_entry_keypress(self, event=None):
        if self._disabled or event is None:
            return None
        if not self._replace_on_type:
            return None

        keysym = getattr(event, "keysym", "") or ""
        if keysym in {"BackSpace", "Delete"}:
            self._replace_on_type = False
            self._query_var.set("")
            try:
                self._tk_entry.icursor(0)
            except Exception:
                pass
            return "break"

        char = getattr(event, "char", "") or ""
        if len(char) != 1 or not char.isprintable():
            return None

        self._replace_on_type = False
        self._query_var.set(char)
        try:
            self._tk_entry.icursor(tk.END)
            self._tk_entry.selection_clear()
        except Exception:
            pass
        return "break"

    def _on_entry_paste(self, _event=None):
        if self._disabled or not self._replace_on_type:
            return None
        try:
            pasted = self._tk_entry.clipboard_get()
        except Exception:
            return None
        self._replace_on_type = False
        self._query_var.set(pasted)
        try:
            self._tk_entry.icursor(tk.END)
            self._tk_entry.selection_clear()
        except Exception:
            pass
        return "break"

    def _on_query_changed(self, *_args) -> None:
        if self._query_guard or self._disabled:
            return
        query = self._query_var.get()
        if query != self._committed:
            self._replace_on_type = False
        matches = filter_dropdown_values(self._values, query)
        display = matches if matches else list(self._values)
        self._open_popup(display, focus_list=False)
        if self._listbox is not None and self._listbox.size():
            self._listbox.selection_clear(0, tk.END)
            self._listbox.selection_set(0)
            self._listbox.activate(0)
            self._listbox.see(0)

    def _focus_popup_list(self, _event=None):
        if not self._popup_is_open():
            self._open_for_current_query()
        if self._listbox is not None:
            try:
                self._listbox.focus_set()
                if not self._listbox.curselection() and self._listbox.size():
                    self._listbox.selection_set(0)
                    self._listbox.activate(0)
                    self._listbox.see(0)
            except Exception:
                pass
        return "break"

    def _on_entry_return(self, _event=None):
        if self._listbox is not None and self._popup_is_open():
            selection = self._listbox.curselection()
            if selection:
                self._confirm_popup_selection()
                return "break"
        self._validate_or_revert()
        return "break"

    def _on_entry_escape(self, _event=None):
        self._revert()
        return "break"

    def _on_popup_escape(self, _event=None):
        self._revert()
        self._restore_owner_activation(self._tk_entry)
        return "break"

    def _toggle_popup(self) -> None:
        if self._disabled:
            return
        self._ensure_entry_state()
        if self._popup_is_open():
            self._revert()
            return
        try:
            self._tk_entry.focus_set()
        except Exception:
            pass
        self._replace_on_type = True
        self._select_all()
        self._open_popup(self._values, focus_list=False)

    # -- selection / validation --------------------------------------

    def _exact_match(self, text: str) -> str | None:
        if text in self._values:
            return text
        folded = text.casefold()
        matches = [value for value in self._values if value.casefold() == folded]
        return matches[0] if len(matches) == 1 else None

    def _validate_or_revert(self) -> None:
        if self._popup_pointer_down:
            return
        typed = self._query_var.get().strip()
        exact = self._exact_match(typed)
        if exact is not None:
            changed = exact != self._committed
            self._commit(exact, notify=changed)
            self._close_popup()
        else:
            self._revert()

    def _outside_dismiss(self) -> None:
        # _DropdownBase._on_owner_click already proved the click is outside both
        # halves of this control; no secondary focus test is needed.
        self._validate_or_revert()

    def _revert(self) -> None:
        self._set_query_text(self._committed)
        self._replace_on_type = True
        self._close_popup()

    def _cancel_edit(self) -> None:
        self._revert()

    def _after_user_commit(self) -> None:
        self._replace_on_type = True
        self._restore_owner_activation(self._tk_entry)
        try:
            self._select_all()
        except Exception:
            pass

    def _cleanup(self) -> None:
        try:
            if self._query_trace:
                self._query_var.trace_remove("write", self._query_trace)
        except Exception:
            pass
        self._query_trace = None
        super()._cleanup()


# Backwards-compatible names used throughout app.py and by third-party callers.
SimpleComboBox = SelectDropdown
SearchableComboBox = SearchableDropdown


__all__ = [
    "SelectDropdown",
    "SearchableDropdown",
    "SimpleComboBox",
    "SearchableComboBox",
    "filter_dropdown_values",
]
