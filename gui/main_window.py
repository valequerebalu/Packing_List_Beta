import logging
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
from pathlib import Path
import threading
import customtkinter as ctk
from PIL import Image
from db.repository import obtener_template

from config.settings import bundle_dir
from core.excel_reader import (
    extraer_orden_beta_de_packing,
    leer_packing_beta,
    leer_termografo,
)
from core.excel_writer import generar_excel
from core.mapper import cruzar_datos, obtener_resumen_cruce
from core.sharepoint import (
    cargar_recursos_sharepoint,
    extraer_datos_sharepoint,
    extraer_filas_sharepoint,
    get_sharepoint_crop_dataframe,
)
from db.repository import obtener_clientes, obtener_cultivos

from gui.theme import (
    BG,
    BG_CONTENT,
    BORDER,
    BORDER_LIGHT,
    CARD,
    FONT_UI,
    PRIMARY,
    PRIMARY_HOVER,
    PRIMARY_LIGHT,
    TEXT,
    TEXT_HEADING,
    TEXT_MUTED,
    WARN,
)

W_INPUT = 260
H_INPUT = 38
RADIUS = 10


class MainWindow:
    def __init__(self, root: tk.Tk, usuario: dict, on_logout=None):
        self.root = root
        self.usuario = usuario
        self.on_logout = on_logout or (lambda: root.destroy())

        self.clientes_config = {}
        self.clientes, self.cultivos = self._cargar_listas()
        self.ruta_packing = tk.StringVar()
        self.ruta_termografo = tk.StringVar()
        self.cliente_var = tk.StringVar(value=self.clientes[0] if self.clientes else "")
        self.cultivo_var = tk.StringVar(value=self.cultivos[0] if self.cultivos else "")
        self.orden_beta_var = tk.StringVar()
        self.sp_estado = tk.StringVar(value="Conectando...")

        self.df_packing = None
        self.df_termografo = None

        # Referencias a labels de archivos para limpiar después
        self.lbl_packing_file = None
        self.lbl_termografo_file = None

        self._configurar_ventana()
        self._build_ui()
        self._simular_conexion_sharepoint()

    def _cargar_listas(self) -> tuple[list[str], list[str]]:
        try:
            clientes_data = obtener_clientes()
            cultivos = obtener_cultivos()
            if clientes_data and cultivos:
                # Guardar info de requires_thermograph por cliente
                self.clientes_config = {
                    c["name"]: c.get("requires_thermograph", True)
                    for c in clientes_data
                }
                clientes = [c["name"] for c in clientes_data]
                return clientes, cultivos

            logging.warning(
                "Las tablas de clientes o cultivos están vacías en Supabase."
            )
            return [], []
        except Exception as e:

            logging.error(f"Error de conexión al cargar tablas: {e}")

            messagebox.showerror(
                title="Error de Conexión",
                message="No se pudieron cargar los clientes y cultivos desde el servidor. Revisa tu conexión a internet.",
            )
        return [], []

    # ─── Ventana ──────────────────────────────────────────────────────────────

    def _configurar_ventana(self):
        self.root.title("Beta Packing — Generador de Packing List")
        self.root.geometry("1080x720")
        self.root.minsize(960, 660)
        self.root.configure(fg_color=BG)
        self.root.resizable(True, True)
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() - 1080) // 2
        y = (self.root.winfo_screenheight() - 720) // 2
        self.root.geometry(f"1080x720+{x}+{y}")

    # ─── UI ───────────────────────────────────────────────────────────────────

    def _build_ui(self):
        self._header()

        body = ctk.CTkFrame(self.root, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=32, pady=(0, 28))
        body.grid_columnconfigure(0, weight=7, minsize=420)
        body.grid_columnconfigure(1, weight=5, minsize=300)
        body.grid_rowconfigure(0, weight=1)

        left = ctk.CTkFrame(body, fg_color="transparent")
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 14))
        left.grid_columnconfigure(0, weight=1)

        right = ctk.CTkFrame(body, fg_color="transparent")
        right.grid(row=0, column=1, sticky="nsew", padx=(14, 0))
        right.grid_columnconfigure(0, weight=1)

        # Columna izquierda: Configuración (incluye Orden Beta) → Archivos → Generar Excel
        # El botón "Generar Excel" se ancla al fondo de la columna (side="bottom")
        # para que siempre quede visible sin necesidad de scroll.
        self._card_config(left)
        self._card_archivos(left)
        self._btn_limpiar(left)
        self._btn_generar(left)

        # Columna derecha: SharePoint → Registro de proceso (se expande para llenar el espacio)
        self._card_sharepoint(right)
        self._card_log(right)

    def _header(self):
        f = ctk.CTkFrame(
            self.root,
            fg_color=CARD,
            corner_radius=0,
            height=88,
            border_width=0,
        )
        f.pack(fill="x", pady=(0, 24))
        f.pack_propagate(False)

        inner = ctk.CTkFrame(f, fg_color="transparent")
        inner.pack(fill="both", padx=36, pady=14)

        ruta_imagen = str(bundle_dir() / "templates" / "logo_beta.png")

        imagen_badge = ctk.CTkImage(
            light_image=Image.open(ruta_imagen),
            dark_image=Image.open(ruta_imagen),
            size=(60, 44),
        )

        badge = ctk.CTkLabel(
            inner,
            text="",
            image=imagen_badge,
            width=60,
            height=44,
        )
        badge.pack(side="left", padx=5)  # Añadimos un pequeño margen si lo deseas

        col = ctk.CTkFrame(inner, fg_color="transparent")
        col.pack(side="left", padx=(16, 0))
        ctk.CTkLabel(
            col,
            text="Beta Packing List",
            font=ctk.CTkFont(family=FONT_UI, size=18, weight="bold"),
            text_color=TEXT_HEADING,
        ).pack(anchor="w")
        ctk.CTkLabel(
            col,
            text="Generador de packing list para exportaciones",
            font=ctk.CTkFont(family=FONT_UI, size=12),
            text_color=TEXT_MUTED,
        ).pack(anchor="w")

        right_block = ctk.CTkFrame(inner, fg_color="transparent")
        right_block.pack(side="right")

        self.user_badge = ctk.CTkLabel(
            right_block,
            text=self.usuario.get("email", "usuario@beta.com"),
            font=ctk.CTkFont(family=FONT_UI, size=12),
            text_color=TEXT_MUTED,
        )
        self.user_badge.pack(anchor="e")

        btn_logout = ctk.CTkButton(
            right_block,
            text="Salir",
            font=ctk.CTkFont(family=FONT_UI, size=11),
            fg_color="transparent",
            hover_color=PRIMARY_LIGHT,
            text_color=TEXT_MUTED,
            height=24,
            corner_radius=6,
            border_width=1,
            border_color=BORDER,
            command=self._on_logout,
        )
        btn_logout.pack(anchor="e", pady=(6, 0))

    # ─── Cards ────────────────────────────────────────────────────────────────

    def _card(self, parent, title, expand=False, **kwargs):
        card = ctk.CTkFrame(
            parent,
            fg_color=CARD,
            corner_radius=14,
            border_width=1,
            border_color=BORDER,
            **kwargs,
        )
        if expand:
            card.pack(fill="both", expand=True, pady=(0, 10))
        else:
            card.pack(fill="x", pady=(0, 10))
        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=18, pady=12)

        if title:
            ctk.CTkLabel(
                inner,
                text=title,
                font=ctk.CTkFont(family=FONT_UI, size=11, weight="bold"),
                text_color=TEXT_MUTED,
            ).pack(anchor="w")
            ctk.CTkFrame(
                inner,
                fg_color=BORDER_LIGHT,
                height=1,
            ).pack(fill="x", pady=(4, 6))

        return inner

    def _card_config(self, parent):
        # Card única: Cliente / Cultivo en la primera fila, Orden Beta en la segunda.
        # Se fusionó lo que antes era una card separada de "Número de orden"
        # para ahorrar el alto de un título + separador + padding extra.
        inner = self._card(parent, "CONFIGURACIÓN")

        grid = ctk.CTkFrame(inner, fg_color="transparent")
        grid.pack(fill="x")
        grid.grid_columnconfigure(0, weight=1, uniform="cfg")
        grid.grid_columnconfigure(1, weight=1, uniform="cfg")

        for col_idx, (label, values, var) in enumerate(
            [
                ("Cliente", self.clientes, self.cliente_var),
                ("Cultivo", self.cultivos, self.cultivo_var),
            ]
        ):
            col = ctk.CTkFrame(grid, fg_color="transparent")
            col.grid(
                row=0,
                column=col_idx,
                sticky="ew",
                padx=(0, 10) if col_idx == 0 else (10, 0),
            )
            ctk.CTkLabel(
                col,
                text=label,
                font=ctk.CTkFont(family=FONT_UI, size=11, weight="bold"),
                text_color=TEXT_HEADING,
            ).pack(anchor="w", pady=(0, 4))
            combo = ctk.CTkComboBox(
                col,
                values=values,
                variable=var,
                state="readonly",
                height=H_INPUT,
                font=ctk.CTkFont(family=FONT_UI, size=13),
                dropdown_font=ctk.CTkFont(family=FONT_UI, size=13),
                border_width=1,
                border_color=BORDER,
                button_color=PRIMARY,
                button_hover_color=PRIMARY_HOVER,
                fg_color=BG_CONTENT,
                command=self._on_cliente_cambio if label == "Cliente" else None,
            ).pack(fill="x")

        orden_block = ctk.CTkFrame(inner, fg_color="transparent")
        orden_block.pack(fill="x", pady=(14, 0))

        ctk.CTkLabel(
            orden_block,
            text="Orden Beta",
            font=ctk.CTkFont(family=FONT_UI, size=11, weight="bold"),
            text_color=TEXT_HEADING,
        ).pack(anchor="w", pady=(0, 4))

        entry_row = ctk.CTkFrame(orden_block, fg_color="transparent")
        entry_row.pack(fill="x")

        self.entry_orden = ctk.CTkEntry(
            entry_row,
            textvariable=self.orden_beta_var,
            placeholder_text="BP-493",
            width=160,
            height=H_INPUT,
            font=ctk.CTkFont(family=FONT_UI, size=14, weight="bold"),
            text_color=PRIMARY,
            border_width=1,
            border_color=BORDER,
            fg_color=BG_CONTENT,
            state="readonly",
        )
        self.entry_orden.pack(side="left")

        self.lbl_ordenes = ctk.CTkLabel(
            entry_row,
            text="Carga el packing para ver órdenes",
            font=ctk.CTkFont(family=FONT_UI, size=11),
            text_color=TEXT_MUTED,
            anchor="w",
        )
        self.lbl_ordenes.pack(side="left", padx=(16, 0))

    def _card_archivos(self, parent):
        # Los dos uploads van lado a lado (en vez de apilados) para reducir
        # el alto total de esta card y liberar espacio para el botón inferior.
        inner = self._card(parent, "ARCHIVOS DE ENTRADA")

        row = ctk.CTkFrame(inner, fg_color="transparent")
        row.pack(fill="x")
        row.grid_columnconfigure(0, weight=1, uniform="arch")
        row.grid_columnconfigure(1, weight=1, uniform="arch")

        self.lbl_packing_file = self._upload_row(
            row, 0, "Packing Beta", self.ruta_packing, self._seleccionar_packing
        )
        self.lbl_termografo_file = self._upload_row(
            row, 1, "Termógrafo", self.ruta_termografo, self._seleccionar_termografo
        )
        # Guardar referencia al label del título del termógrafo para actualizarlo
        self.lbl_termografo_title = self.lbl_termografo_file.master.winfo_children()[0]
        self._actualizar_label_termografo()

    def _upload_row(self, parent, col_idx, title, var, cmd):
        card = ctk.CTkFrame(
            parent,
            fg_color=BG_CONTENT,
            border_width=1,
            border_color=BORDER,
            corner_radius=10,
        )
        card.grid(
            row=0,
            column=col_idx,
            sticky="nsew",
            padx=(0, 7) if col_idx == 0 else (7, 0),
        )

        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=14, pady=10)

        ctk.CTkLabel(
            inner,
            text=title,
            font=ctk.CTkFont(family=FONT_UI, size=13, weight="bold"),
            text_color=TEXT_HEADING,
        ).pack(anchor="w")

        lbl_file = ctk.CTkLabel(
            inner,
            text="Ningún archivo seleccionado",
            font=ctk.CTkFont(family=FONT_UI, size=11),
            text_color=TEXT_MUTED,
            anchor="w",
            wraplength=200,
            justify="left",
        )
        lbl_file.pack(fill="x", pady=(6, 8), anchor="w")

        ctk.CTkButton(
            inner,
            text="Seleccionar archivo",
            font=ctk.CTkFont(family=FONT_UI, size=12, weight="bold"),
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            height=32,
            corner_radius=8,
            command=lambda: self._on_upload(var, cmd, lbl_file),
        ).pack(fill="x")

        return lbl_file

    def _btn_limpiar(self, parent):
        # Botón para limpiar todos los datos ingresados con estilo moderno minimalista
        self.btn_limpiar = ctk.CTkButton(
            parent,
            text="Limpiar datos",
            font=ctk.CTkFont(family=FONT_UI, size=12, weight="bold"),
            fg_color=("#E2E8F0", "#2D2D2D"),
            hover_color=("#CBD5E1", "#3D3D3D"),
            text_color=("#475569", "#E2E8F0"),
            height=36,
            corner_radius=10,
            border_width=0,
            command=self._on_limpiar,
        )
        self.btn_limpiar.pack(side="bottom", fill="x", pady=(15, 18))

    def _btn_generar(self, parent):
        # Anclado al fondo de la columna izquierda: siempre visible sin scroll.
        self.btn_generar = ctk.CTkButton(
            parent,
            text="Generar Excel",
            font=ctk.CTkFont(family=FONT_UI, size=14, weight="bold"),
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            height=44,
            corner_radius=12,
            command=self._on_generar,
        )
        self.btn_generar.pack(side="bottom", fill="x", pady=(18, 0))

    def _card_sharepoint(self, parent):
        inner = self._card(parent, "SHAREPOINT")
        row = ctk.CTkFrame(inner, fg_color="transparent")
        row.pack(fill="x", pady=(4, 0))

        col = ctk.CTkFrame(row, fg_color="transparent")
        col.pack(side="left", fill="x", expand=True)

        self.lbl_sp_archivo = ctk.CTkLabel(
            col,
            text="Sincronizando...",
            font=ctk.CTkFont(family=FONT_UI, size=13, weight="bold"),
            text_color=TEXT_HEADING,
            anchor="w",
        )
        self.lbl_sp_archivo.pack(anchor="w")

        self.lbl_sp_detalle = ctk.CTkLabel(
            col,
            text="Conectando con Microsoft 365...",
            font=ctk.CTkFont(family=FONT_UI, size=11),
            text_color=TEXT_MUTED,
            anchor="w",
        )
        self.lbl_sp_detalle.pack(anchor="w", pady=(4, 0))

        estado = ctk.CTkFrame(row, fg_color="transparent")
        estado.pack(side="right")
        self.sp_dot = ctk.CTkLabel(
            estado,
            text="●",
            font=ctk.CTkFont(family=FONT_UI, size=14),
            text_color=WARN,
        )
        self.sp_dot.pack(side="left", padx=(0, 4))
        self.lbl_estado = ctk.CTkLabel(
            estado,
            textvariable=self.sp_estado,
            font=ctk.CTkFont(family=FONT_UI, size=11, weight="bold"),
            text_color=WARN,
        )
        self.lbl_estado.pack(side="left")

    def _card_log(self, parent):
        inner = self._card(parent, "REGISTRO DE PROCESO", expand=True)
        frame = ctk.CTkFrame(
            inner,
            fg_color=BG_CONTENT,
            corner_radius=10,
            border_width=1,
            border_color=BORDER_LIGHT,
        )
        frame.pack(fill="both", expand=True)

        self.log_text = ctk.CTkTextbox(
            frame,
            height=160,
            font=ctk.CTkFont(family="Cascadia Code", size=10),
            fg_color=BG_CONTENT,
            text_color=TEXT,
            wrap="word",
            activate_scrollbars=False,
        )
        self.log_text.pack(fill="both", expand=True, padx=2, pady=2)
        self.log_text.tag_config("ok", foreground=PRIMARY)
        self.log_text.tag_config("error", foreground="#dc2626")
        self.log_text.tag_config("info", foreground="#2563eb")
        self.log_text.tag_config("warn", foreground="#d97706")
        self.log_text.tag_config("muted", foreground=TEXT_MUTED)

        self._log("Sistema iniciado. Esperando archivos...", "muted")

    # ─── Helpers ──────────────────────────────────────────────────────────────

    def _on_logout(self):
        self.on_logout()

    def _on_limpiar(self):
        """Limpia todos los datos: archivos, orden beta y listas desplegables."""
        # Limpiar rutas de archivos
        self.ruta_packing.set("")
        self.ruta_termografo.set("")

        # Limpiar orden beta
        self.orden_beta_var.set("")

        # Limpiar labels de archivos
        if self.lbl_packing_file:
            self.lbl_packing_file.configure(
                text="Ningún archivo seleccionado", text_color=TEXT_MUTED
            )
        if self.lbl_termografo_file:
            self.lbl_termografo_file.configure(
                text="Ningún archivo seleccionado", text_color=TEXT_MUTED
            )

        # Resetear listas desplegables a valores por defecto
        self.cliente_var.set(self.clientes[0] if self.clientes else "")
        self.cultivo_var.set(self.cultivos[0] if self.cultivos else "")

        # Resetear label de órdenes disponibles
        self.lbl_ordenes.configure(
            text="Carga el packing para ver órdenes", text_color=TEXT_MUTED
        )

        # Limpiar DataFrames en memoria
        self.df_packing = None
        self.df_termografo = None

        self._log("Datos limpiados. Sistema listo para nuevos archivos.", "ok")

    def _on_cliente_cambio(self, cliente_seleccionado):
        """Callback cuando cambia el cliente — actualiza label del termógrafo."""
        self._actualizar_label_termografo()

    def _actualizar_label_termografo(self):
        """Actualiza el label del termógrafo indicando si es opcional."""
        cliente = self.cliente_var.get().strip()
        requiere = self.clientes_config.get(cliente, True)
        if hasattr(self, 'lbl_termografo_title') and self.lbl_termografo_title:
            if requiere:
                self.lbl_termografo_title.configure(text="Termógrafo")
            else:
                self.lbl_termografo_title.configure(text="Termógrafo (Opcional)")

    def _log(self, mensaje: str, tipo: str = "info"):
        hora = datetime.now().strftime("%H:%M:%S")
        prefijos = {"ok": "✓", "error": "✗", "info": "→", "warn": "⚠", "muted": "#"}
        self.log_text.configure(state="normal")
        self.log_text.insert(
            "end", f"[{hora}]  {prefijos.get(tipo, '→')}  {mensaje}\n", tipo
        )
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    # ─── Eventos ──────────────────────────────────────────────────────────────

    def _on_upload(self, var, cmd, lbl_file):
        ruta = cmd()
        if not ruta:
            return
        var.set(ruta)
        nombre = Path(ruta).name
        lbl_file.configure(text=nombre, text_color=PRIMARY)
        self._log(f"Archivo cargado: {nombre}", "ok")
        if var is self.ruta_packing:
            self._leer_packing_en_hilo(ruta)

    def _leer_packing_en_hilo(self, ruta: str):
        def leer():
            try:
                from core.excel_reader import (
                    leer_packing_beta,
                    extraer_orden_beta_de_packing,
                )

                # Intentar obtener input_header_row e input_sheet_name del template
                input_header_row = None
                input_sheet_name = None
                cliente = self.cliente_var.get()
                cultivo = self.cultivo_var.get()
                if cliente and cultivo:
                    try:
                        template = obtener_template(cliente, cultivo)
                        if template:
                            mc = template["mapping_config"]
                            input_header_row = mc.get("input_header_row")
                            input_sheet_name = mc.get("input_sheet_name")
                    except Exception:
                        pass  # Sin template, auto-detección

                self._log("Analizando packing Beta...", "info")
                self.df_packing = leer_packing_beta(
                    ruta, header_row=input_header_row, sheet_name=input_sheet_name
                )
                ordenes = extraer_orden_beta_de_packing(self.df_packing)

                def ui():
                    if ordenes:
                        self.lbl_ordenes.configure(
                            text="  |  ".join(ordenes), text_color=PRIMARY
                        )
                        if len(ordenes) == 1:
                            self.orden_beta_var.set(ordenes[0])
                            self._log(f"Orden detectada: {ordenes[0]}", "ok")
                        else:
                            self._log(f"{len(ordenes)} órdenes detectadas", "ok")
                    else:
                        self.lbl_ordenes.configure(
                            text="No se detectaron órdenes", text_color="#dc2626"
                        )
                        cliente = self.cliente_var.get().strip()
                        cultivo = self.cultivo_var.get().strip()
                        self._log(
                            "No se detectaron órdenes en el packing. "
                            "Verifica que el cliente y cultivo sean correctos "
                            f"(seleccionado: {cliente} / {cultivo})", "warn"
                        )

                self.root.after(0, ui)
            except Exception as e:
                logging.exception("Error al leer packing Beta en segundo plano")
                error_msg = str(e)
                self.root.after(0, lambda: self._log(f"Error: {error_msg}", "error"))

        threading.Thread(target=leer, daemon=True).start()

    def _on_orden_change(self, event=None):
        self.orden_beta_var.set(self.orden_beta_var.get().upper().strip())

    def _seleccionar_packing(self):
        return filedialog.askopenfilename(
            title="Seleccionar Packing Beta",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")],
        )

    def _seleccionar_termografo(self):
        return filedialog.askopenfilename(
            title="Seleccionar archivo de Termógrafo",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")],
        )

    def _on_generar(self):
        if not self.ruta_packing.get():
            messagebox.showwarning("Faltante", "Selecciona el archivo de Packing Beta.")
            return

        cliente = self.cliente_var.get().strip()
        requiere_termografo = self.clientes_config.get(cliente, True)

        if requiere_termografo and not self.ruta_termografo.get():
            messagebox.showwarning("Faltante", "Selecciona el archivo de Termógrafo.")
            return

        orden = self.orden_beta_var.get().strip()
        if not orden:
            messagebox.showwarning("Faltante", "Ingresa el número de orden Beta.")
            self.entry_orden.focus()
            return

        if requiere_termografo:
            try:
                self.df_termografo = leer_termografo(self.ruta_termografo.get(), orden)
            except ValueError as e:
                messagebox.showerror("Error", str(e))
                return

            if self.df_termografo.empty:
                continuar = messagebox.askyesno(
                    "Termógrafo sin registros",
                    f"No se encontraron los códigos de termógrafo "
                    f"para la orden '{orden}'.\n\n"
                    f"¿Deseas continuar de todas formas?",
                )
                if not continuar:
                    self._log("Proceso cancelado — termógrafo sin códigos", "warn")
                    return
                self._log(
                    "Termógrafo sin códigos — se generará sin la última columna", "warn"
                )
        else:
            # Para clientes sin termógrafo (ej: Global Berry), usar DataFrame vacío
            import pandas as pd
            self.df_termografo = pd.DataFrame()
            self._log("Cliente sin termógrafo requerido — omitiendo archivo", "info")

        self.btn_generar.configure(state="disabled", text="Generando...")
        self._log(
            f"Iniciando → Cliente: {self.cliente_var.get()} | "
            f"Cultivo: {self.cultivo_var.get()} | "
            f"Orden: {orden}",
            "info",
        )
        threading.Thread(target=self._proceso_generacion, daemon=True).start()

    def _proceso_generacion(self):
        try:
            orden = self.orden_beta_var.get().strip()

            self._log("Cruzando información...", "info")
            # Obtener template y mapping_config desde Supabase

            template = obtener_template(self.cliente_var.get(), self.cultivo_var.get())

            if template is None:
                raise ValueError(
                    f"No hay plantilla configurada para "
                    f"'{self.cliente_var.get()}' — '{self.cultivo_var.get()}'"
                )

            self._log(f"Plantilla encontrada: {template['name']}", "ok")

            # Leer packing con input_header_row e input_sheet_name del mapping_config
            mc = template["mapping_config"]
            input_header_row = mc.get("input_header_row")
            input_sheet_name = mc.get("input_sheet_name")

            if self.df_packing is None:
                self._log("Leyendo packing Beta...", "info")
                self.df_packing = leer_packing_beta(
                    self.ruta_packing.get(),
                    header_row=input_header_row,
                    sheet_name=input_sheet_name,
                )
            else:
                # Re-leer si el header_row o sheet_name del template difieren
                if input_header_row is not None or input_sheet_name is not None:
                    self._log("Releyendo packing Beta con config del template...", "info")
                    self.df_packing = leer_packing_beta(
                        self.ruta_packing.get(),
                        header_row=input_header_row,
                        sheet_name=input_sheet_name,
                    )
            self._log(f"Packing: {len(self.df_packing)} filas totales", "ok")

            self._log(
                f"Termógrafo: {len(self.df_termografo)} registros",
                "ok" if not self.df_termografo.empty else "warn",
            )

            resumen = obtener_resumen_cruce(self.df_packing, self.df_termografo, orden)
            self._log(
                f"Cruce: {resumen['pallets_cruzados']}/{resumen['total_pallets']} pallets",
                "ok" if not resumen["advertencias"] else "warn",
            )
            for adv in resumen["advertencias"]:
                self._log(adv, "warn")

            self._log("Obteniendo datos de SharePoint...", "info")
            datos_sp = {}
            sharepoint_rows = []
            recursos_sp = {}
            try:
                resources = template.get("resources")
                if resources:
                    recursos_sp, advertencias = cargar_recursos_sharepoint(
                        resources,
                        self.cultivo_var.get(),
                        orden,
                        self.cliente_var.get(),
                    )
                    for advertencia in advertencias:
                        self._log(f"SharePoint: {advertencia}", "warn")
                    principal = recursos_sp.get("embarque")
                    if principal is None and recursos_sp:
                        principal = next(iter(recursos_sp.values()))
                    if principal:
                        datos_sp = principal["datos"]
                        sharepoint_rows = principal["filas"]
                else:
                    sp_header_row = template["mapping_config"].get(
                        "sharepoint_header_row", 3
                    )
                    df_sp = get_sharepoint_crop_dataframe(
                        self.cultivo_var.get(), header_row=sp_header_row
                    )
                    if df_sp is not None:
                        datos_sp = extraer_datos_sharepoint(
                            df_sp, orden, self.cliente_var.get()
                        )
                        sharepoint_rows = extraer_filas_sharepoint(
                            df_sp, orden, self.cliente_var.get()
                        )

                if sharepoint_rows:
                    self._log(
                        "SharePoint: datos sincronizados (%d filas)"
                        % len(sharepoint_rows),
                        "ok",
                    )
                elif resources:
                    self._log(
                        f"SharePoint: no se encontró la orden '{orden}' "
                        f"para el cliente '{self.cliente_var.get()}'.",
                        "warn",
                    )
                elif df_sp is not None:
                    self._log(
                        "SharePoint: no se encontraron filas para la orden.",
                        "warn",
                    )
            except FileNotFoundError as e:
                self._log(str(e), "warn")
                self._log("Continuando sin datos de SharePoint...", "warn")
            except ValueError as e:
                self._log(str(e), "warn")
                self._log("Continuando sin datos de SharePoint...", "warn")

            # Luego pasas el mapping_config a cruzar_datos
            filas = cruzar_datos(
                self.df_packing,
                self.df_termografo,
                datos_sp,
                self.cliente_var.get(),
                self.cultivo_var.get(),
                orden,
                template["mapping_config"],
                sharepoint_rows=sharepoint_rows,
                sharepoint_resources=recursos_sp,
            )
            self._log(f"Cruce completado: {len(filas)} filas", "ok")

            self._log("Generando Excel de salida...", "info")
            ruta_salida = generar_excel(
                filas,
                template,
                orden,
                datos_sp,
                self.df_termografo,
            )
            self._log(f"Excel generado: {ruta_salida}", "ok")

            self.root.after(
                0,
                lambda: messagebox.showinfo(
                    "Completado",
                    f"Packing list generado correctamente.\n\n"
                    f"Cliente: {self.cliente_var.get()}\n"
                    f"Orden: {orden}\n"
                    f"Filas: {len(filas)}\n"
                    f"Archivo: {ruta_salida}",
                ),
            )
        except Exception as e:
            logging.exception("Error durante el proceso de generación")
            self._log(f"Error: {str(e)}", "error")

            error_texto = str(e)
            self.root.after(0, lambda: messagebox.showerror("Error", error_texto))

        finally:
            self.root.after(
                0,
                lambda: self.btn_generar.configure(
                    state="normal", text="Generar Excel"
                ),
            )

    # ─── SharePoint ───────────────────────────────────────────────────────────

    def _simular_conexion_sharepoint(self):
        def conectar():
            import time

            time.sleep(1.5)
            self.root.after(0, self._actualizar_sp_conectado)

        threading.Thread(target=conectar, daemon=True).start()

    def _actualizar_sp_conectado(self):
        cultivo = self.cultivo_var.get().lower()
        self.lbl_sp_archivo.configure(text=f"data_{cultivo}_exportacion.xlsx")
        self.lbl_sp_detalle.configure(
            text=f"Última sincronización: hoy {datetime.now().strftime('%H:%M')} · v4.1"
        )
        self.sp_estado.set("Conectado")
        self.sp_dot.configure(text_color=PRIMARY)
        self.lbl_estado.configure(text_color=PRIMARY)
        self._log("SharePoint conectado — datos disponibles", "ok")
