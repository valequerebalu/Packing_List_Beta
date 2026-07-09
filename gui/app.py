from __future__ import annotations

import logging

import customtkinter as ctk

from gui.login_window import LoginWindow
from gui.main_window import MainWindow
from gui.theme import configure_app


class BetaPackingApp:
    def __init__(self):
        configure_app()
        self.root = ctk.CTk()
        self.root.report_callback_exception = self._report_callback_exception
        self.root.withdraw()
        self.main_window = None

    def _report_callback_exception(self, exc_type, exc_value, exc_traceback):
        logging.critical(
            "Excepción no controlada en Tkinter",
            exc_info=(exc_type, exc_value, exc_traceback),
        )

    def _on_logout(self):
        self.root.withdraw()
        LoginWindow(self.root, self._on_login)

    def _on_login(self, user: dict):
        self.root.deiconify()
        usuario = {
            "id": user.get("id"),
            "username": user.get("username", ""),
            "email": user.get("email") or user.get("username", "usuario@beta.com"),
            "nombre": user.get("full_name") or user.get("nombre", "Usuario Beta"),
            "is_admin": user.get("is_admin", False),
        }
        self.main_window = MainWindow(self.root, usuario, on_logout=self._on_logout)

    def run(self):
        LoginWindow(self.root, self._on_login)
        self.root.mainloop()
