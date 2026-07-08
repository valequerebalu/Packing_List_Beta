from __future__ import annotations

import customtkinter as ctk

PRIMARY = "#0f6e3f"
PRIMARY_LIGHT = "#e8f5ee"
PRIMARY_HOVER = "#0b5832"
BG = "#f0f2f4"
BG_CONTENT = "#f7f8fa"
CARD = "#ffffff"
BORDER = "#e2e5e7"
BORDER_LIGHT = "#edf0f2"
TEXT = "#141617"
TEXT_MUTED = "#6b7278"
TEXT_HEADING = "#1b1e21"
WARN = "#f59e0b"
SUCCESS = "#10b981"
FONT_UI = "Segoe UI"


def configure_app() -> None:
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("green")
