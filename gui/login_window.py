from __future__ import annotations

import logging
import threading
import webbrowser
import os
import customtkinter as ctk
from tkinter import messagebox
from PIL import Image

from auth.msal_auth import MicrosoftAuthenticator
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
    HEIGHT = 400  # Reducido: ya no se muestran código de verificación ni botones de navegador/copiar

    def __init__(self, parent, on_success):
        super().__init__(parent)
        self.parent = parent
        self.on_success = on_success
        self.authenticator = MicrosoftAuthenticator()
        self.status_var = ctk.StringVar(value="Listo para iniciar sesión.")
        self.code_var = ctk.StringVar(value="---")
        self.url_var = ctk.StringVar(value="")
        self.link_var = ctk.StringVar(value="")
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

        head = ctk.CTkFrame(outer, fg_color="transparent")
        head.grid(row=0, column=0, sticky="ew", padx=24, pady=(24, 10))
        head.grid_columnconfigure(0, weight=1)

        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

        ruta_imagen = os.path.join(base_dir, "templates", "logo_beta.png")

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
            text="Acceso corporativo con Microsoft 365",
            font=ctk.CTkFont(family=FONT_UI, size=12),
            text_color=TEXT_MUTED,
        ).pack(anchor="w", pady=(4, 0))

        body = ctk.CTkFrame(outer, fg_color="#f8faf9", corner_radius=20)
        body.grid(row=1, column=0, sticky="nsew", padx=24, pady=12)
        body.grid_columnconfigure(0, weight=1)

        info = ctk.CTkLabel(
            body,
            text="Inicia sesión para acceder a herramientas de Packing List.",
            font=ctk.CTkFont(family=FONT_UI, size=13),
            text_color=TEXT_MUTED,
            justify="left",
            wraplength=390,
        )
        info.grid(row=0, column=0, sticky="w", padx=18, pady=(18, 12))

        self.status_label = ctk.CTkLabel(
            body,
            textvariable=self.status_var,
            font=ctk.CTkFont(family=FONT_UI, size=11),
            text_color=TEXT_MUTED,
            wraplength=390,
            justify="left",
        )
        self.status_label.grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 14))

        # ------------------------------------------------------------------
        # BLOQUE DE CÓDIGO DE VERIFICACIÓN (device code flow) — DESACTIVADO
        # También se desactivan los botones "Abrir navegador" y "Copiar código"
        # que dependían de este flujo. Se conserva toda la lógica y los
        # widgets comentados por si se necesita reactivar más adelante.
        # No se elimina absolutamente nada, solo se deja de construir/gridear.
        # ------------------------------------------------------------------
        # code_card = ctk.CTkFrame(
        #     body,
        #     fg_color="#edf4ff",
        #     corner_radius=16,
        #     border_width=1,
        #     border_color="#bfd5ff",
        # )
        # code_card.grid(row=2, column=0, sticky="ew", padx=18, pady=(0, 10))
        # code_card.grid_columnconfigure(0, weight=1)
        #
        # ctk.CTkLabel(
        #     code_card,
        #     text="Código de verificación",
        #     font=ctk.CTkFont(family=FONT_UI, size=9, weight="bold"),
        #     text_color="#355a8a",
        # ).grid(row=0, column=0, sticky="w", padx=10, pady=(6, 1))
        #
        # self.code_label = ctk.CTkLabel(
        #     code_card,
        #     textvariable=self.code_var,
        #     font=ctk.CTkFont(family=FONT_UI, size=28, weight="bold"),
        #     text_color="#1d4ed8",
        # )
        # self.code_label.grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 10))
        #
        # self.link_hint = ctk.CTkLabel(
        #     body,
        #     text="To sign in, use a web browser to open the page below:",
        #     font=ctk.CTkFont(family=FONT_UI, size=11),
        #     text_color=TEXT_MUTED,
        #     justify="left",
        #     wraplength=390,
        # )
        # self.link_hint.grid(row=3, column=0, sticky="w", padx=18, pady=(0, 3))
        #
        # self.link_label = ctk.CTkLabel(
        #     body,
        #     textvariable=self.link_var,
        #     font=ctk.CTkFont(family=FONT_UI, size=11, underline=True),
        #     text_color="#1d4ed8",
        #     cursor="hand2",
        #     wraplength=390,
        #     justify="left",
        # )
        # self.link_label.grid(row=4, column=0, sticky="w", padx=18, pady=(0, 8))
        # self.link_label.bind("<Button-1>", lambda e: self._open_browser())
        #
        # actions = ctk.CTkFrame(body, fg_color="transparent")
        # actions.grid(row=5, column=0, sticky="ew", padx=18, pady=(0, 10))
        # actions.grid_columnconfigure((0, 1), weight=1)
        #
        # self.btn_open_browser = ctk.CTkButton(
        #     actions,
        #     text="Abrir navegador",
        #     command=self._open_browser,
        #     height=38,
        #     corner_radius=14,
        #     fg_color=PRIMARY,
        #     hover_color=PRIMARY_HOVER,
        #     font=ctk.CTkFont(family=FONT_UI, size=12, weight="bold"),
        #     state="disabled",
        # )
        # self.btn_open_browser.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        #
        # self.btn_copy_code = ctk.CTkButton(
        #     actions,
        #     text="Copiar código",
        #     command=self._copy_code,
        #     height=38,
        #     corner_radius=14,
        #     fg_color="#e9eeeb",
        #     hover_color="#d7e1da",
        #     text_color=TEXT,
        #     font=ctk.CTkFont(family=FONT_UI, size=12, weight="bold"),
        #     state="disabled",
        # )
        # self.btn_copy_code.grid(row=0, column=1, sticky="ew", padx=(8, 0))
        # ------------------------------------------------------------------
        # FIN BLOQUE DESACTIVADO
        # ------------------------------------------------------------------

        self.btn_login = ctk.CTkButton(
            body,
            text="Ingresar con Microsoft",
            command=self._login,
            height=42,
            corner_radius=14,
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            font=ctk.CTkFont(family=FONT_UI, size=13, weight="bold"),
        )
        self.btn_login.grid(row=6, column=0, sticky="ew", padx=18, pady=(4, 10))

        self.btn_mock = ctk.CTkButton(
            body,
            text="Usar modo desarrollo",
            command=self._mock_login,
            height=40,
            corner_radius=14,
            fg_color="#e9eeeb",
            hover_color="#d7e1da",
            text_color=TEXT,
            font=ctk.CTkFont(family=FONT_UI, size=12, weight="bold"),
        )
        self.btn_mock.grid(row=7, column=0, sticky="ew", padx=18, pady=(0, 18))

    def _set_status(self, payload) -> None:
        def update():
            if isinstance(payload, dict):
                code = payload.get("user_code")
                url = payload.get("verification_uri") or payload.get(
                    "verification_uri_complete"
                )
                message = payload.get("message") or payload.get("status")

                if code:
                    self.code_var.set(code)
                if url:
                    self.url_var.set(url)
                    self.link_var.set(url)
                if message:
                    self.status_var.set(message)

            self.status_var.set(str(payload))

        self.after(0, update)

    def _open_browser(self) -> None:
        url = self.url_var.get().strip()
        if not url:
            messagebox.showwarning(
                "URL no disponible", "Aún no se ha recibido la URL de verificación."
            )
            return
        webbrowser.open(url)

    def _copy_code(self) -> None:
        code = self.code_var.get().strip()
        if not code or code == "---":
            messagebox.showwarning(
                "Código no disponible",
                "Aún no se ha generado un código de verificación.",
            )
            return

        self.clipboard_clear()
        self.clipboard_append(code)
        self.update()
        messagebox.showinfo(
            "Código copiado", "El código de verificación fue copiado con éxito."
        )

    def _login(self) -> None:
        self.btn_login.configure(state="disabled", text="Conectando...")
        self.btn_mock.configure(state="disabled")

        def worker():
            try:
                result = self.authenticator.login_interactive(
                    progress_callback=self._set_status
                )

                def finish():
                    if result.success:
                        self.status_var.set("Login completado.")
                        self.on_success(result.account)
                        self.destroy()
                        return

                    messagebox.showerror(
                        "Login fallido", result.message or "No se pudo iniciar sesión."
                    )
                    self.status_var.set("No se pudo completar el inicio de sesión.")
                    self.btn_login.configure(
                        state="normal", text="Ingresar con Microsoft"
                    )
                    self.btn_mock.configure(state="normal")

                self.after(0, finish)
            except Exception as exc:
                logging.exception("Error durante el inicio de sesión")

                def fail():
                    messagebox.showerror("Login fallido", str(exc))
                    self.status_var.set("Error durante el inicio de sesión.")
                    self.btn_login.configure(
                        state="normal", text="Ingresar con Microsoft"
                    )
                    self.btn_mock.configure(state="normal")

                self.after(0, fail)

        threading.Thread(target=worker, daemon=True).start()

    def _mock_login(self) -> None:
        result = self.authenticator.login_mock()
        self.on_success(result.account)
        self.destroy()
