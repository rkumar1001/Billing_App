from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING

import customtkinter as ctk

from .dialogs import ask_directory

if TYPE_CHECKING:
    from .app import BillingApp


class SettingsScreen(ctk.CTkFrame):
    def __init__(self, parent, app: "BillingApp") -> None:
        super().__init__(parent, fg_color="transparent")
        self.app = app
        self.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            self, text="Settings",
            font=ctk.CTkFont(size=24, weight="bold"),
        ).grid(row=0, column=0, sticky="w", padx=24, pady=(20, 16))

        card = ctk.CTkFrame(self, corner_radius=10)
        card.grid(row=1, column=0, sticky="ew", padx=24, pady=8)
        card.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(card, text="Default folder the picker opens in").grid(
            row=0, column=0, sticky="w", padx=14, pady=(14, 6),
        )
        self.out_var = ctk.StringVar(value=app.cfg.default_output_root)
        row = ctk.CTkFrame(card, fg_color="transparent")
        row.grid(row=0, column=1, sticky="ew", padx=14, pady=(14, 6))
        row.grid_columnconfigure(0, weight=1)
        ctk.CTkEntry(row, textvariable=self.out_var).grid(row=0, column=0, sticky="ew")
        ctk.CTkButton(row, text="Browse…", width=90, command=self._pick_folder).grid(
            row=0, column=1, padx=(6, 0)
        )

        ctk.CTkButton(card, text="Save", command=self._save).grid(
            row=1, column=0, columnspan=2, sticky="e", padx=14, pady=(10, 14),
        )

        ctk.CTkLabel(
            self,
            text="Field auto-detection is cached per source folder in a hidden `.billingapp.json` file. Delete that file to force re-detection.",
            font=ctk.CTkFont(size=12),
            text_color=("gray40", "gray70"),
            wraplength=600, justify="left",
        ).grid(row=2, column=0, sticky="w", padx=26, pady=(16, 6))

    def _pick_folder(self) -> None:
        start = (self.out_var.get() or "").strip()
        initialdir = start if Path(start).is_dir() else None
        parent = self.winfo_toplevel()
        folder = ask_directory(parent=parent, initialdir=initialdir, title="Choose default folder")
        if folder:
            self.out_var.set(folder)

    def _save(self) -> None:
        self.app.cfg.default_output_root = self.out_var.get()
        self.app.save_config()
