from __future__ import annotations

import os
import re
import unicodedata
from pathlib import Path

import pandas as pd
from tkinter import messagebox

from config.settings import SHAREPOINT_LOCAL_DIR


def get_sharepoint_crop_dataframe(crop_name):

    try:
        file_name = f"Reporte Embarques de Exportaciónes - {crop_name} 2026.xlsx"

        roots = []
        if SHAREPOINT_LOCAL_DIR:
            roots.append(Path(SHAREPOINT_LOCAL_DIR))

        user_home = Path(os.path.expanduser("~"))
        roots.extend(
            [
                Path(
                    r"D:\OneDrive-vquerebalu\Complejo Agroindustrial Beta S.A\Intranet - Reporte Exportaciones"
                ),
                user_home
                / "Complejo Agroindustrial Beta S.A"
                / "Intranet - Reporte Exportaciones",
                user_home / "SharePoint",
            ]
        )

        full_path = None
        for root in roots:
            candidate = root / file_name
            if candidate.exists():
                full_path = candidate
                break

        if full_path is None:
            raise FileNotFoundError

        print(f"[Core] Intentando leer matriz local en: {full_path}")

        df = pd.read_excel(full_path, sheet_name="Base de datos", header=3)
        print(f"[Core] ¡Éxito! Datos de {crop_name} cargados correctamente en memoria.")
        return df

    except ValueError as exc:
        if "Worksheet named 'Base de datos' not found" not in str(exc):
            raise

        libro = pd.ExcelFile(full_path)
        error_msg = (
            f"El archivo '{file_name}' no tiene una pestaña llamada 'Base de datos'.\n\n"
            f"Pestañas disponibles: {libro.sheet_names}"
        )
        messagebox.showerror("Error de Archivo", error_msg)
        return None

    except FileNotFoundError:
        error_msg = (
            f"No se pudo encontrar el archivo 'Reporte Embarques de Exportaciónes - {crop_name} 2026.xlsx' en la ruta local.\n\n"
            "Verifica que:\n"
            f"1. El nombre en el código coincida con la carpeta sincronizada.\n"
            "2. El archivo se haya descargado por completo en tu explorador de Windows."
        )
        messagebox.showerror("Error de Archivo", error_msg)
        return None


def extraer_datos_sharepoint(df: pd.DataFrame, orden_beta: str | None = None) -> dict:
    """Normaliza una fila de SharePoint a las claves que espera el mapper."""
    if df is None or df.empty:
        return {}

    normalizado = df.copy()
    normalizado.columns = [_normalizar_columna(c) for c in normalizado.columns]

    if orden_beta and "orden_beta" in normalizado.columns:
        orden = _normalizar_orden(orden_beta)
        filtrado = normalizado[
            normalizado["orden_beta"].astype(str).map(_normalizar_orden) == orden
        ]
        if not filtrado.empty:
            normalizado = filtrado

    fila = normalizado.iloc[0].to_dict()

    return {
        "Cliente": _valor_alias(fila, "cliente", "customer", "customer_name"),
        "Nave": _valor_alias(fila, "nave", "vessel", "vessel_name"),
        "ETD Programada": _valor_alias(
            fila,
            "etd_programada",
            "etd",
            "etd_programado",
        ),
        "ETA Programado": _valor_alias(
            fila,
            "eta_programado",
            "eta",
            "eta_programada",
        ),
        "Puerto Descarga": _valor_alias(
            fila,
            "puerto_descarga",
            "destination_port",
            "puerto",
        ),
        "Orden Beta": _valor_alias(fila, "orden_beta", "orden", "order_beta"),
        " Contenedor 2.0": _valor_alias(
            fila,
            "contenedor_2_0",
            "container_number",
            "container_no",
        ),
        "AWB - Contenedor": _valor_alias(
            fila,
            "awb___contenedor",
            "awb_contenedor",
            "awb_container",
            "awb_no",
            "contenedor",
            "container_number",
            "container_no",
        ),
    }


def _normalizar_columna(nombre: object) -> str:
    texto = str(nombre).strip().lower()
    texto = (
        unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")
    )
    texto = "_".join(texto.split())
    texto = texto.replace("/", "_").replace("-", "_")
    return texto


def _valor(fila: dict, campo: str):
    valor = fila.get(campo)
    if valor is None:
        return None
    if pd.isna(valor):
        return None
    texto = str(valor).strip()
    if texto in ("", "nan", "None", "NaN"):
        return None
    return texto


def _valor_alias(fila: dict, *campos: str):
    for campo in campos:
        valor = _valor(fila, campo)
        if valor is not None:
            return valor
    return None


def _normalizar_orden(valor: str) -> str:
    texto = str(valor).strip().upper()
    partes = re.match(r"^([A-Z]+)-?(\d+)$", texto)
    if not partes:
        return re.sub(r"[^A-Z0-9]", "", texto)

    prefijo, numero = partes.groups()
    return f"{prefijo}{int(numero):04d}"
