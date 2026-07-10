from __future__ import annotations

import logging
import threading
import customtkinter as ctk
from tkinter import messagebox
from PIL import Image

from auth.db_auth import authenticate, DBAuthError
from config.settings import bundle_dir
from gui.theme import (
    BG,
    BORDER,
    CARD,
    FONT_UI,
    PRIMARY,
    PRIMARY_HOVER,
    TEXT,
    TEXT_MUTED,
)


class LoginWindow(ctk.CTkToplevel):
    WIDTH = 480
    HEIGHT = 480

    def __init__(self, parent, on_success):
        super().__init__(parent)
        self.parent = parent
        self.on_success = on_success
        self.status_var = ctk.StringVar(value="Ingresa tus credenciales.")
        self.title("Beta Packing List - Login")
        self.geometry(f"{self.WIDTH}x{self.HEIGHT}")
        self.minsize(self.WIDTH, self.HEIGHT)
        self.maxsize(self.WIDTH, self.HEIGHT)
        self.resizable(False, False)
        self.configure(fg_color=BG)
        self._build_ui()
        self.transient(parent)
        self.grab_set()
        self.lift()
        self.focus_force()
        self.after(0, self._center_window)
        self.after(200, self.entry_username.focus)

    def _center_window(self) -> None:
        self.update_idletasks()
        width = self.WIDTH
        height = self.HEIGHT
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = max((screen_width - width) // 2, 0)
        y = max((screen_height - height) // 2, 0)
        self.geometry(f"{width}x{height}+{x}+{y}")

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        outer = ctk.CTkFrame(
            self,
            corner_radius=24,
            fg_color=CARD,
            border_width=1,
            border_color=BORDER,
        )
        outer.grid(row=0, column=0, sticky="nsew", padx=22, pady=22)
        outer.grid_columnconfigure(0, weight=1)

        # ─── Header ──────────────────────────────────────────────────────
        head = ctk.CTkFrame(outer, fg_color="transparent")
        head.grid(row=0, column=0, sticky="ew", padx=24, pady=(24, 10))
        head.grid_columnconfigure(0, weight=1)

        ruta_imagen = str(bundle_dir() / "templates" / "logo_beta.png")

        imagen_badge = ctk.CTkImage(
            light_image=Image.open(ruta_imagen),
            dark_image=Image.open(ruta_imagen),
            size=(60, 44),
        )

        badge = ctk.CTkLabel(
            head,
            text="",
            image=imagen_badge,
            width=65,
            height=44,
        )
        badge.grid(row=0, column=0, sticky="w")

        text_box = ctk.CTkFrame(head, fg_color="transparent")
        text_box.grid(row=0, column=1, sticky="w", padx=(14, 0))
        ctk.CTkLabel(
            text_box,
            text="Beta Packing List",
            font=ctk.CTkFont(family=FONT_UI, size=20, weight="bold"),
            text_color=TEXT,
        ).pack(anchor="w")
        ctk.CTkLabel(
            text_box,
            text="Acceso corporativo",
            font=ctk.CTkFont(family=FONT_UI, size=12),
            text_color=TEXT_MUTED,
        ).pack(anchor="w", pady=(4, 0))

        # ─── Body ────────────────────────────────────────────────────────
        body = ctk.CTkFrame(outer, fg_color="#f8faf9", corner_radius=20)
        body.grid(row=1, column=0, sticky="nsew", padx=24, pady=12)
        body.grid_columnconfigure(0, weight=1)

        form = ctk.CTkFrame(body, fg_color="transparent")
        form.grid(row=0, column=0, sticky="ew", padx=18, pady=(18, 6))

        ctk.CTkLabel(
            form,
            text="Usuario",
            font=ctk.CTkFont(family=FONT_UI, size=12, weight="bold"),
            text_color=TEXT,
        ).pack(anchor="w", pady=(0, 4))

        self.entry_username = ctk.CTkEntry(
            form,
            height=38,
            font=ctk.CTkFont(family=FONT_UI, size=13),
            border_width=1,
            border_color=BORDER,
            fg_color="white",
        )
        self.entry_username.pack(fill="x", pady=(0, 12))
        self.entry_username.bind("<Return>", lambda e: self.entry_password.focus())

        ctk.CTkLabel(
            form,
            text="Contraseña",
            font=ctk.CTkFont(family=FONT_UI, size=12, weight="bold"),
            text_color=TEXT,
        ).pack(anchor="w", pady=(0, 4))

        self.entry_password = ctk.CTkEntry(
            form,
            height=38,
            font=ctk.CTkFont(family=FONT_UI, size=13),
            show="*",
            border_width=1,
            border_color=BORDER,
            fg_color="white",
        )
        self.entry_password.pack(fill="x", pady=(0, 6))
        self.entry_password.bind("<Return>", lambda e: self._login())

        # ─── Status ──────────────────────────────────────────────────────
        self.status_label = ctk.CTkLabel(
            form,
            textvariable=self.status_var,
            font=ctk.CTkFont(family=FONT_UI, size=11),
            text_color=TEXT_MUTED,
            wraplength=390,
            justify="left",
        )
        self.status_label.pack(anchor="w", pady=(4, 14))

        # ─── Login button ────────────────────────────────────────────────
        self.btn_login = ctk.CTkButton(
            form,
            text="Ingresar",
            command=self._login,
            height=42,
            corner_radius=14,
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            font=ctk.CTkFont(family=FONT_UI, size=13, weight="bold"),
        )
        self.btn_login.pack(fill="x", pady=(0, 10))

        # ─── Dev mode ────────────────────────────────────────────────────
        self.btn_mock = ctk.CTkButton(
            form,
            text="Usar modo desarrollo",
            command=self._mock_login,
            height=40,
            corner_radius=14,
            fg_color="#e9eeeb",
            hover_color="#d7e1da",
            text_color=TEXT,
            font=ctk.CTkFont(family=FONT_UI, size=12, weight="bold"),
        )
        self.btn_mock.pack(fill="x", pady=(0, 18))

    def _login(self) -> None:
        username = self.entry_username.get().strip()
        password = self.entry_password.get()

        if not username:
            self.status_var.set("Ingresa tu usuario.")
            self.entry_username.focus()
            return
        if not password:
            self.status_var.set("Ingresa tu contraseña.")
            self.entry_password.focus()
            return

        self.btn_login.configure(state="disabled", text="Validando...")
        self.btn_mock.configure(state="disabled")
        self.status_var.set("Conectando...")

        def worker():
            try:
                user = authenticate(username, password)

                def finish():
                    self.on_success(user)
                    self.destroy()

                self.after(0, finish)

            except DBAuthError as e:
                msg = str(e)
                logging.warning(f"Login fallido para '{username}': {msg}")

                def fail(mensaje=msg):
                    self.status_var.set(mensaje)
                    self.btn_login.configure(state="normal", text="Ingresar")
                    self.btn_mock.configure(state="normal")

                self.after(0, fail)

            except Exception as e:
                logging.exception("Error de conexión durante login")

                def error():
                    messagebox.showerror(
                        "Error de conexión",
                        "No se pudo conectar con el servidor. "
                        "Verifica tu conexión a internet e intenta nuevamente.",
                    )
                    self.status_var.set("Error de conexión.")
                    self.btn_login.configure(state="normal", text="Ingresar")
                    self.btn_mock.configure(state="normal")

                self.after(0, error)

        threading.Thread(target=worker, daemon=True).start()

    def _mock_login(self) -> None:
        user = {
            "id": 0,
            "username": "dev",
            "email": "dev@beta.com",
            "full_name": "Usuario Desarrollo",
            "is_active": True,
            "is_admin": True,
        }
        self.on_success(user)
        self.destroy()
