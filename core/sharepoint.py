from __future__ import annotations

import logging
import os
import re
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd

from config.settings import SHAREPOINT_LOCAL_DIR

# Nombre de la carpeta del negocio en OneDrive
_NEGOCIO = "Complejo Agroindustrial Beta S.A"
_SUBCARPETA = "Intranet - Reporte Exportaciones"


def get_sharepoint_crop_dataframe(
    crop_name: str, header_row: int = 3
) -> pd.DataFrame | None:
    """
    Busca y carga el Excel de SharePoint sincronizado localmente para el cultivo dado.
    Busca automáticamente en todos los discos y rutas de OneDrive.

    Args:
        crop_name:  Nombre del cultivo (ej: 'Palta', 'Arandano')
        header_row: Fila del encabezado en la hoja 'Base de datos' (0-indexed).
                    Palta usa 3, Arándano usa 4.
    """
    anio = datetime.now().year
    file_name = f"Reporte Embarques de Exportaciónes - {crop_name} {anio}.xlsx"

    roots = _buscar_rutas_sharepoint()

    logging.debug("Buscando '%s' en %d rutas posibles", file_name, len(roots))

    full_path = None
    for root in roots:
        candidate = root / file_name
        if candidate.exists():
            full_path = candidate
            break

    if full_path is None:
        rutas = "\n".join(f"  - {r}" for r in roots)
        raise FileNotFoundError(
            f"No se encontró '{file_name}'.\n\n"
            f"Rutas buscadas:\n{rutas}"
        )

    logging.debug("SharePoint leyendo: %s", full_path)

    try:
        df = pd.read_excel(full_path, sheet_name="Base de datos", header=header_row)
        logging.debug("SharePoint %s cargado — %d filas", crop_name, len(df))
        return df

    except ValueError as exc:
        if "Worksheet named 'Base de datos' not found" in str(exc):
            libro = pd.ExcelFile(full_path)
            raise ValueError(
                f"El archivo '{file_name}' no tiene pestaña 'Base de datos'.\n"
                f"Pestañas disponibles: {libro.sheet_names}"
            )
        raise


def _buscar_rutas_sharepoint() -> list[Path]:
    """
    Escanea el sistema buscando la carpeta de SharePoint del negocio.
    Funciona en cualquier máquina sin configuración previa.
    """
    roots: list[Path] = []

    # 1. Override explícito vía env var (si se configuró)
    if SHAREPOINT_LOCAL_DIR and SHAREPOINT_LOCAL_DIR.exists():
        roots.append(SHAREPOINT_LOCAL_DIR)

    # 2. Escanear discos: OneDrive-* y cualquier carpeta con la subcarpeta del negocio
    for drive_letter in "CDEFGH":
        drive = Path(f"{drive_letter}:\\")
        if not drive.exists():
            continue
        try:
            for carpeta in drive.iterdir():
                if not carpeta.is_dir():
                    continue
                nombre = carpeta.name.lower()
                # Patrón OneDrive-*
                if nombre.startswith("onedrive"):
                    candidata = carpeta / _NEGOCIO / _SUBCARPETA
                    if candidata.exists():
                        roots.append(candidata)
                else:
                    # Búsqueda profunda: cualquier carpeta con el negocio adentro
                    candidata = carpeta / _NEGOCIO / _SUBCARPETA
                    if candidata.exists():
                        roots.append(candidata)
        except PermissionError:
            continue

    # 3. Fallbacks en home del usuario
    user_home = Path(os.path.expanduser("~"))
    fallbacks = [
        user_home / _NEGOCIO / _SUBCARPETA,
        user_home / "OneDrive" / _NEGOCIO / _SUBCARPETA,
        user_home / "SharePoint",
    ]
    for fb in fallbacks:
        if fb.exists():
            roots.append(fb)

    # 4. Deduplicar manteniendo orden
    seen: set[str] = set()
    unique: list[Path] = []
    for r in roots:
        resolved = str(r.resolve())
        if resolved not in seen:
            seen.add(resolved)
            unique.append(r)

    return unique


def extraer_datos_sharepoint(df: pd.DataFrame, orden_beta: str | None = None) -> dict:
    """
    Normaliza una fila de SharePoint y retorna TODOS los campos como dict.
    El mapper puede acceder a cualquier campo por nombre de columna normalizado.
    """
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

    fila = normalizado.iloc[0]

    # Retornar TODA la fila normalizada (valores limpios)
    resultado = {}
    for col, val in fila.items():
        if col.startswith("unnamed"):
            continue
        if pd.isna(val):
            resultado[col] = None
        else:
            texto = str(val).strip()
            resultado[col] = None if texto in ("", "nan", "None", "NaN") else texto

    # Aliases backward-compatible (nombres con mayúsculas que usa Tesco)
    resultado["Cliente"] = _valor_alias(resultado, "cliente", "customer", "customer_name")
    resultado["Nave"] = _valor_alias(resultado, "nave", "vessel", "vessel_name")
    resultado["ETD Programada"] = _valor_alias(resultado, "etd_programada", "etd", "etd_programado")
    resultado["ETA Programado"] = _valor_alias(resultado, "eta_programado", "eta", "eta_programada")
    resultado["Puerto Descarga"] = _valor_alias(resultado, "puerto_descarga", "destination_port", "puerto")
    resultado["Orden Beta"] = _valor_alias(resultado, "orden_beta", "orden", "order_beta")
    resultado["AWB - Contenedor"] = _valor_alias(
        resultado,
        "awb___contenedor",
        "awb_contenedor",
        "awb_container",
        "awb_no",
        "contenedor",
        "container_number",
        "container_no",
    )

    return resultado


# ─── Helpers privados ─────────────────────────────────────────────────────────


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
    if valor is None or pd.isna(valor):
        return None
    texto = str(valor).strip()
    return None if texto in ("", "nan", "None", "NaN") else texto


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
