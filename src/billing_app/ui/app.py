from __future__ import annotations

import customtkinter as ctk

from ..resources import AppConfig, load_config, save_config
from .generator_screen import GeneratorScreen
from .settings import SettingsScreen


class BillingApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()

        self.config_ = load_config()

        ctk.set_appearance_mode(self.config_.theme)
        ctk.set_default_color_theme(self.config_.color_theme)

        self.title("Billing App — Monthly Invoice Duplicator")
        self.geometry(self.config_.window_geometry)
        self.minsize(1024, 640)

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._build_sidebar()
        self.content = ctk.CTkFrame(self, corner_radius=0, fg_color=("gray95", "gray15"))
        self.content.grid(row=0, column=1, sticky="nsew")
        self.content.grid_columnconfigure(0, weight=1)
        self.content.grid_rowconfigure(0, weight=1)

        self.current_screen: ctk.CTkFrame | None = None
        self.show_generator()

        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # --- layout ---------------------------------------------------------
    def _build_sidebar(self) -> None:
        sidebar = ctk.CTkFrame(self, width=200, corner_radius=0)
        sidebar.grid(row=0, column=0, sticky="nsw")
        sidebar.grid_rowconfigure(10, weight=1)

        ctk.CTkLabel(
            sidebar, text="Billing App",
            font=ctk.CTkFont(size=20, weight="bold"),
        ).grid(row=0, column=0, padx=20, pady=(22, 4), sticky="w")

        ctk.CTkLabel(
            sidebar, text="Invoice duplicator",
            font=ctk.CTkFont(size=12),
            text_color=("gray40", "gray70"),
        ).grid(row=1, column=0, padx=20, pady=(0, 18), sticky="w")

        def nav(text: str, command, row: int) -> None:
            ctk.CTkButton(
                sidebar, text=text, anchor="w", height=40,
                fg_color="transparent",
                text_color=("gray10", "gray90"),
                hover_color=("gray85", "gray25"),
                command=command,
            ).grid(row=row, column=0, sticky="ew", padx=10, pady=2)

        nav("  Generator", self.show_generator, 2)
        nav("  Settings", self.show_settings, 3)

        ctk.CTkLabel(sidebar, text="Theme", font=ctk.CTkFont(size=12)).grid(
            row=11, column=0, padx=20, pady=(10, 0), sticky="w",
        )
        menu = ctk.CTkOptionMenu(
            sidebar, values=["System", "Light", "Dark"], command=self._on_theme_change,
        )
        menu.set(self.config_.theme)
        menu.grid(row=12, column=0, padx=20, pady=(2, 20), sticky="ew")

    # --- navigation -----------------------------------------------------
    def _swap(self, screen: ctk.CTkFrame) -> None:
        if self.current_screen is not None:
            self.current_screen.destroy()
        screen.grid(row=0, column=0, sticky="nsew")
        self.current_screen = screen

    def show_generator(self) -> None:
        self._swap(GeneratorScreen(self.content, app=self))

    def show_settings(self) -> None:
        self._swap(SettingsScreen(self.content, app=self))

    # --- events ---------------------------------------------------------
    def _on_theme_change(self, value: str) -> None:
        ctk.set_appearance_mode(value)
        self.config_.theme = value
        save_config(self.config_)

    def _on_close(self) -> None:
        self.config_.window_geometry = self.geometry()
        save_config(self.config_)
        self.destroy()

    # --- helpers --------------------------------------------------------
    def save_config(self) -> None:
        save_config(self.config_)

    @property
    def cfg(self) -> AppConfig:
        return self.config_
