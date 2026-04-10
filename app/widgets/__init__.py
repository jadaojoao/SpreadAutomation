# app/widgets/__init__.py
# Design System widgets for spread_automation (CustomTkinter).
# All colours come from themes/tokens.py — never hardcode hex here.

from __future__ import annotations

import customtkinter as ctk
from tkinter import filedialog
from themes.tokens import DS, FONT, RADIUS, BORDER_WIDTH


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _font(role: str) -> ctk.CTkFont:
    spec = FONT[role]
    kwargs: dict = {"size": spec["size"], "weight": spec["weight"]}
    if "family" in spec:
        kwargs["family"] = spec["family"]
    return ctk.CTkFont(**kwargs)


def _pair(token: str) -> tuple[str, str]:
    return DS[token]


# ---------------------------------------------------------------------------
# SectionHeading — DS heading primitive (h2-level)
# ---------------------------------------------------------------------------

class SectionHeading(ctk.CTkFrame):
    """A labelled section divider following DS typography rules.

    Usage::

        heading = SectionHeading(parent, title="File Selection")
        heading.pack(fill="x", pady=(0, 8))
    """

    def __init__(self, master, title: str, **kwargs):
        kwargs.setdefault("fg_color", "transparent")
        kwargs.setdefault("corner_radius", 0)
        super().__init__(master, **kwargs)

        self._title = ctk.CTkLabel(
            self,
            text=title,
            font=_font("h3"),
            text_color=_pair("foreground"),
            anchor="w",
        )
        self._title.pack(side="left", padx=0, pady=0)

        sep = ctk.CTkFrame(
            self,
            height=1,
            corner_radius=0,
            fg_color=_pair("border"),
            border_width=0,
        )
        sep.pack(side="left", fill="x", expand=True, padx=(10, 0), pady=0)


# ---------------------------------------------------------------------------
# StyledCard — DS card surface (CTkFrame wrapper)
# ---------------------------------------------------------------------------

class StyledCard(ctk.CTkFrame):
    """A card-surface frame with DS bg, border, and radius.

    Usage::

        card = StyledCard(parent, padding=16)
        card.pack(fill="x")
    """

    def __init__(self, master, padding: int = 14, **kwargs):
        kwargs.setdefault("fg_color", _pair("card"))
        kwargs.setdefault("border_color", _pair("border"))
        kwargs.setdefault("border_width", BORDER_WIDTH)
        kwargs.setdefault("corner_radius", RADIUS["base"])
        super().__init__(master, **kwargs)
        self._padding = padding

    def inner_pack_kwargs(self) -> dict:
        """Return padx/pady to use when packing children inside this card."""
        p = self._padding
        return {"padx": p, "pady": p}


# ---------------------------------------------------------------------------
# FilePickerWidget — DS-styled file selector
# ---------------------------------------------------------------------------

class FilePickerWidget(ctk.CTkFrame):
    """Card-surface file picker with label, truncated path, and Browse button.

    Usage::

        picker = FilePickerWidget(parent, label="Source (CVM Excel):")
        picker.pack(fill="x")
        path = picker.get_path()
    """

    def __init__(self, master, label: str, filetypes: list | None = None, **kwargs):
        kwargs.setdefault("fg_color", _pair("card"))
        kwargs.setdefault("border_color", _pair("border"))
        kwargs.setdefault("border_width", BORDER_WIDTH)
        kwargs.setdefault("corner_radius", RADIUS["base"])
        super().__init__(master, **kwargs)

        self._filetypes = filetypes or [("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")]
        self.path_var = ctk.StringVar()

        self.columnconfigure(1, weight=1)

        self._label = ctk.CTkLabel(
            self,
            text=label,
            font=_font("label"),
            text_color=_pair("foreground"),
            anchor="w",
            width=160,
        )
        self._label.grid(row=0, column=0, padx=(14, 8), pady=12, sticky="w")

        self._entry = ctk.CTkEntry(
            self,
            textvariable=self.path_var,
            state="disabled",
            font=_font("body_sm"),
            fg_color=_pair("muted"),
            border_color=_pair("border"),
            text_color=_pair("muted_foreground"),
            placeholder_text="No file selected",
            corner_radius=RADIUS["md"],
        )
        self._entry.grid(row=0, column=1, padx=(0, 8), pady=12, sticky="ew")

        self._btn = ctk.CTkButton(
            self,
            text="Browse…",
            width=90,
            font=_font("body_sm"),
            fg_color=_pair("secondary"),
            hover_color=_pair("secondary_hover"),
            text_color=_pair("secondary_foreground"),
            border_color=_pair("border"),
            border_width=BORDER_WIDTH,
            corner_radius=RADIUS["md"],
            command=self._browse,
        )
        self._btn.grid(row=0, column=2, padx=(0, 14), pady=12)

    def _browse(self):
        path = filedialog.askopenfilename(filetypes=self._filetypes)
        if path:
            self.path_var.set(path)
            self._entry.configure(state="normal")
            self._entry.delete(0, "end")
            self._entry.insert(0, path)
            self._entry.configure(state="disabled")
            self._entry.configure(text_color=_pair("foreground"))

    def get_path(self) -> str:
        return self.path_var.get()

    def clear(self):
        self.path_var.set("")
        self._entry.configure(state="normal", text_color=_pair("muted_foreground"))
        self._entry.delete(0, "end")
        self._entry.configure(state="disabled")


# ---------------------------------------------------------------------------
# ProgressWidget — DS-styled progress bar + status label
# ---------------------------------------------------------------------------

class ProgressWidget(ctk.CTkFrame):
    """DS progress bar with status label above it.

    Usage::

        progress = ProgressWidget(parent)
        progress.pack(fill="x")
        progress.update(message="Running…", value=0.5)
        progress.reset()
    """

    def __init__(self, master, **kwargs):
        kwargs.setdefault("fg_color", "transparent")
        kwargs.setdefault("corner_radius", 0)
        super().__init__(master, **kwargs)

        self._status_var = ctk.StringVar(value="Ready")

        self._status_label = ctk.CTkLabel(
            self,
            textvariable=self._status_var,
            font=_font("body_sm"),
            text_color=_pair("muted_foreground"),
            anchor="w",
        )
        self._status_label.pack(fill="x", padx=0, pady=(0, 4))

        self._bar = ctk.CTkProgressBar(
            self,
            height=6,
            corner_radius=RADIUS["sm"],
            fg_color=_pair("muted"),
            progress_color=_pair("primary"),
        )
        self._bar.pack(fill="x", padx=0)
        self._bar.set(0)

    def update(self, message: str, value: float):
        self._status_var.set(message)
        self._bar.set(max(0.0, min(1.0, value)))
        self.update_idletasks()

    # Keep backward-compat with old ProgressWidget.update_progress() call site
    def update_progress(self, message: str, progress: float):
        self.update(message, progress)

    def reset(self):
        self.update("Ready", 0.0)


# ---------------------------------------------------------------------------
# LabeledField — DS field primitive (label + entry in a column)
# ---------------------------------------------------------------------------

class LabeledField(ctk.CTkFrame):
    """A DS-style field with a label above an entry.

    Usage::

        field = LabeledField(parent, label="Company", placeholder="e.g. Minerva")
        field.pack()
        value = field.get()
    """

    def __init__(
        self,
        master,
        label: str,
        placeholder: str = "",
        width: int = 140,
        textvariable: ctk.StringVar | None = None,
        **kwargs,
    ):
        kwargs.setdefault("fg_color", "transparent")
        kwargs.setdefault("corner_radius", 0)
        super().__init__(master, **kwargs)

        self._var = textvariable or ctk.StringVar()

        self._label = ctk.CTkLabel(
            self,
            text=label,
            font=_font("label"),
            text_color=_pair("foreground"),
            anchor="w",
        )
        self._label.pack(anchor="w", pady=(0, 4))

        self._entry = ctk.CTkEntry(
            self,
            textvariable=self._var,
            placeholder_text=placeholder,
            width=width,
            font=_font("body"),
            corner_radius=RADIUS["md"],
            fg_color=_pair("card"),
            border_color=_pair("border"),
            text_color=_pair("foreground"),
            placeholder_text_color=_pair("muted_foreground"),
        )
        self._entry.pack(fill="x")

    def get(self) -> str:
        return self._var.get()

    def set(self, value: str):
        self._var.set(value)

    @property
    def var(self) -> ctk.StringVar:
        return self._var
