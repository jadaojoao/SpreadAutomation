# app/screens/screen_1a.py
# Mode 1A screen — redesigned with Design System (2026-04-10).

from __future__ import annotations

import threading
import traceback

import customtkinter as ctk

from app.widgets import (
    FilePickerWidget,
    LabeledField,
    ProgressWidget,
    SectionHeading,
    StyledCard,
    _font,
    _pair,
)
from core.models import EntityType
from engine.workflow_1a import Mode1AWorkflow
from themes.tokens import DS, RADIUS, BORDER_WIDTH


class Screen1A(ctk.CTkFrame):
    """Mode 1A: fill an existing Spread Proxy from a CVM source file."""

    def __init__(self, master, **kwargs):
        kwargs.setdefault("fg_color", DS["background"])
        kwargs.setdefault("corner_radius", 0)
        super().__init__(master, **kwargs)

        self.columnconfigure(0, weight=1)
        self.rowconfigure(4, weight=1)   # log area expands

        # ── Page header ──────────────────────────────────────────────────
        header = ctk.CTkFrame(self, fg_color="transparent", corner_radius=0)
        header.grid(row=0, column=0, sticky="ew", padx=24, pady=(20, 4))
        header.columnconfigure(0, weight=1)

        ctk.CTkLabel(
            header,
            text="Mode 1A",
            font=_font("h1"),
            text_color=_pair("foreground"),
            anchor="w",
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkLabel(
            header,
            text="Fill existing Spread Proxy from a CVM source file",
            font=_font("body"),
            text_color=_pair("muted_foreground"),
            anchor="w",
        ).grid(row=1, column=0, sticky="w", pady=(2, 0))

        # ── File selection card ───────────────────────────────────────────
        files_card = StyledCard(self, padding=0)
        files_card.grid(row=1, column=0, sticky="ew", padx=24, pady=(16, 0))
        files_card.columnconfigure(0, weight=1)

        SectionHeading(files_card, title="Files").pack(
            fill="x", padx=14, pady=(14, 8)
        )

        self.source_picker = FilePickerWidget(files_card, label="Source (CVM Excel):")
        self.source_picker.pack(fill="x", padx=1, pady=(0, 1))

        sep = ctk.CTkFrame(files_card, height=1, fg_color=_pair("border"), corner_radius=0)
        sep.pack(fill="x", padx=14)

        self.spread_picker = FilePickerWidget(files_card, label="Spread Proxy:")
        self.spread_picker.pack(fill="x", padx=1, pady=(1, 0))

        # ── Settings card ─────────────────────────────────────────────────
        settings_card = StyledCard(self, padding=14)
        settings_card.grid(row=2, column=0, sticky="ew", padx=24, pady=(12, 0))
        settings_card.columnconfigure((0, 1, 2, 3), weight=1)

        SectionHeading(settings_card, title="Settings").grid(
            row=0, column=0, columnspan=4, sticky="ew", pady=(0, 12)
        )

        # Company + Period
        self.company_field = LabeledField(
            settings_card, label="Company", placeholder="e.g. Minerva"
        )
        self.company_field.grid(row=1, column=0, sticky="ew", padx=(0, 10), pady=(0, 10))

        self.period_field = LabeledField(
            settings_card, label="Period", placeholder="e.g. 4T24"
        )
        self.period_field.grid(row=1, column=1, sticky="ew", padx=(0, 10), pady=(0, 10))

        self.dest_col_field = LabeledField(
            settings_card, label="Dest Column", placeholder="C", width=80
        )
        self.dest_col_field.var.set("C")
        self.dest_col_field.grid(row=1, column=2, sticky="ew", padx=(0, 10), pady=(0, 10))

        self.prior_col_field = LabeledField(
            settings_card, label="Prior Col (opt.)", placeholder="D", width=80
        )
        self.prior_col_field.var.set("D")
        self.prior_col_field.grid(row=1, column=3, sticky="ew", pady=(0, 10))

        # Entity type
        entity_row = ctk.CTkFrame(settings_card, fg_color="transparent", corner_radius=0)
        entity_row.grid(row=2, column=0, columnspan=4, sticky="w")

        ctk.CTkLabel(
            entity_row,
            text="Entity type:",
            font=_font("label"),
            text_color=_pair("foreground"),
        ).pack(side="left", padx=(0, 12))

        self.entity_var = ctk.StringVar(value="consolidated")
        for label, value in (("Consolidated", "consolidated"), ("Individual", "individual")):
            ctk.CTkRadioButton(
                entity_row,
                text=label,
                variable=self.entity_var,
                value=value,
                font=_font("body"),
                text_color=_pair("foreground"),
                fg_color=_pair("primary"),
                border_color=_pair("border"),
                hover_color=_pair("primary_hover"),
            ).pack(side="left", padx=(0, 16))

        # ── Action row ────────────────────────────────────────────────────
        action_row = ctk.CTkFrame(self, fg_color="transparent", corner_radius=0)
        action_row.grid(row=3, column=0, sticky="ew", padx=24, pady=(16, 0))
        action_row.columnconfigure(1, weight=1)

        self.run_btn = ctk.CTkButton(
            action_row,
            text="Execute Mode 1A",
            font=_font("body"),
            width=180,
            height=38,
            fg_color=_pair("primary"),
            hover_color=_pair("primary_hover"),
            text_color=_pair("primary_foreground"),
            corner_radius=RADIUS["base"],
            command=self._run,
        )
        self.run_btn.grid(row=0, column=0, sticky="w")

        self.progress = ProgressWidget(action_row)
        self.progress.grid(row=0, column=1, sticky="ew", padx=(20, 0))

        # ── Log card ──────────────────────────────────────────────────────
        log_card = StyledCard(self, padding=0)
        log_card.grid(row=4, column=0, sticky="nsew", padx=24, pady=(12, 20))
        log_card.columnconfigure(0, weight=1)
        log_card.rowconfigure(1, weight=1)

        SectionHeading(log_card, title="Log").pack(
            fill="x", padx=14, pady=(14, 8)
        )

        self.log_box = ctk.CTkTextbox(
            log_card,
            font=_font("mono"),
            fg_color=_pair("card"),
            border_width=0,
            corner_radius=0,
            text_color=_pair("foreground"),
            scrollbar_button_color=_pair("scrollbar"),
            scrollbar_button_hover_color=_pair("scrollbar_hover"),
            wrap="word",
        )
        self.log_box.pack(fill="both", expand=True, padx=1, pady=(0, 1))

    # ── Log helper ───────────────────────────────────────────────────────

    def _log(self, text: str):
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")

    # backward-compat alias used in older call sites
    def log(self, text: str):
        self._log(text)

    # ── Workflow ──────────────────────────────────────────────────────────

    def _run(self):
        source = self.source_picker.get_path()
        spread = self.spread_picker.get_path()
        company = self.company_field.get().strip()
        period = self.period_field.get().strip()
        dest_col = self.dest_col_field.get().strip()
        prior_col = self.prior_col_field.get().strip()
        entity = self.entity_var.get()

        if not all([source, spread, company, period, dest_col]):
            self._log("[!] Fill in Source, Spread, Company, Period, and Dest Column before running.")
            return

        self.run_btn.configure(state="disabled")
        self.progress.update("Starting…", 0.05)
        self._log(f"[*] Mode 1A — {company} {period}")

        threading.Thread(
            target=self._execute,
            args=(source, spread, company, period, dest_col, prior_col, entity),
            daemon=True,
        ).start()

    def _execute(self, source, spread, company, period, dest_col, prior_col, entity):
        try:
            entity_enum = (
                EntityType.CONSOLIDATED if entity == "consolidated" else EntityType.INDIVIDUAL
            )
            workflow = Mode1AWorkflow()

            self.progress.update("Running workflow…", 0.5)

            result = workflow.execute(
                source_path=source,
                spread_path=spread,
                company=company,
                period=period,
                dest_col=dest_col,
                prior_col=prior_col if prior_col else None,
                entity_type=entity_enum,
            )

            self.progress.update("Done", 1.0)
            self._log("[+] Workflow completed successfully.")

            val = result.get("validation", {})
            self._log(f"    Mapped  : {result.get('mapped_count')} items")
            self._log(f"    Valid   : {val.get('is_valid')}  |  Missing: {len(val.get('missing', []))}")

        except Exception:
            self.progress.update("Failed", 0.0)
            self._log(f"[-] Workflow failed:\n{traceback.format_exc()}")
        finally:
            self.run_btn.configure(state="normal")
