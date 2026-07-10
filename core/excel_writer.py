"""
Generación del Excel final a partir de la plantilla del cliente.

Toma las filas ya cruzadas y mapeadas (provenientes de core/mapper.py)
y las escribe directamente sobre una copia de la plantilla .xlsx del cliente,
respetando el formato original (fechas, dropdowns, encabezados).

El orden de columnas y las columnas de fecha ahora se obtienen del
mapping_config almacenado en Supabase (tabla templates), NO están hardcodeados.
"""

from pathlib import Path
from datetime import datetime
import shutil

import openpyxl
import pandas as pd

from config.settings import TEMPLATES_DIR, OUTPUT_DIR


# ─── Función principal ────────────────────────────────────────────────────────


def generar_excel(
    filas: list[dict],
    template_data: dict,
    orden_beta: str,
) -> str:
    """
    Genera el archivo Excel final escribiendo las filas sobre una copia
    de la plantilla del cliente.

    Args:
        filas:         Lista de diccionarios (salida de core.mapper.cruzar_datos)
        template_data: Dict con info del template desde Supabase
                       (id, name, file_path, mapping_config)
        orden_beta:    Orden procesada, ej: 'BP-493'

    Returns:
        Ruta absoluta del archivo Excel generado
    """
    mapping_config = template_data["mapping_config"]

    # Obtener configuración desde mapping_config
    sheet_name = mapping_config.get("sheet_name", "Sheet1")
    data_start_row = mapping_config.get("data_start_row", 7)
    columns_config = mapping_config["columns"]

    # Orden de columnas = orden de keys del dict (Python 3.7+ preserva inserción)
    column_order = list(columns_config.keys())

    # Columnas de fecha: usar date_columns del mapping o detectar por "format"
    date_columns = set()
    if "date_columns" in mapping_config:
        date_columns = set(mapping_config["date_columns"])
    else:
        for col_name, col_cfg in columns_config.items():
            if col_cfg.get("format") == "date":
                date_columns.add(col_name)

    # Resolver ruta del template
    ruta_plantilla = TEMPLATES_DIR / template_data["file_path"]
    if not ruta_plantilla.exists():
        raise FileNotFoundError(
            f"No se encontró el archivo de plantilla: {ruta_plantilla}"
        )

    # Construir ruta de salida
    ruta_salida = _construir_ruta_salida(orden_beta)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # Copiar la plantilla para no tocar el original
    shutil.copy(ruta_plantilla, ruta_salida)

    wb = openpyxl.load_workbook(ruta_salida)
    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active

    _escribir_filas(ws, filas, column_order, date_columns, data_start_row)

    wb.save(ruta_salida)
    wb.close()

    return str(ruta_salida)


# ─── Helpers privados ─────────────────────────────────────────────────────────


def _construir_ruta_salida(orden_beta: str) -> Path:
    """Genera un nombre de archivo único para el output."""
    orden_limpia = str(orden_beta).strip().upper().replace(" ", "")
    prefijo = "".join(c for c in orden_limpia if c.isalpha())
    numeros = "".join(c for c in orden_limpia if c.isdigit())
    if not numeros:
        return OUTPUT_DIR / f"{prefijo}-PACKING_LIST.xlsx"
    return OUTPUT_DIR / f"{prefijo}{numeros.zfill(4)}-PACKING_LIST.xlsx"


def _escribir_filas(
    ws,
    filas: list[dict],
    column_order: list[str],
    date_columns: set[str],
    data_start_row: int,
):
    """
    Escribe cada fila del listado en la hoja, empezando en data_start_row.
    Respeta el orden exacto de columnas definido en mapping_config.
    Aplica formato de fecha a las columnas correspondientes.
    """
    fila_excel = data_start_row

    for fila_datos in filas:
        for idx, nombre_columna in enumerate(column_order, start=1):
            valor = fila_datos.get(nombre_columna)
            celda = ws.cell(row=fila_excel, column=idx)

            if nombre_columna in date_columns:
                celda.value = _convertir_fecha(valor)
                celda.number_format = "dd/mm/yyyy"
            else:
                celda.value = _convertir_valor(valor)

        fila_excel += 1


def _convertir_valor(valor):
    """
    Convierte el valor a su tipo más apropiado para Excel.
    Intenta convertir a número si es posible, mantiene texto si no.
    """
    if valor is None:
        return None

    valor_str = str(valor).strip()
    if not valor_str:
        return None

    try:
        if "." in valor_str:
            return float(valor_str)
        return int(valor_str)
    except ValueError:
        return valor_str


def _convertir_fecha(valor):
    """
    Convierte un valor de fecha (string en varios formatos posibles)
    a un objeto datetime que Excel pueda formatear correctamente.
    """
    if valor is None:
        return None

    valor_str = str(valor).strip()
    if not valor_str:
        return None

    formatos_posibles = [
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%dT%H:%M:%S",
        "%d/%m/%Y %H:%M:%S",
        "%d-%m-%Y %H:%M:%S",
        "%d/%m/%y",
        "%m/%d/%Y",
        "%Y/%m/%d",
    ]

    for formato in formatos_posibles:
        try:
            return datetime.strptime(valor_str, formato)
        except ValueError:
            continue

    if isinstance(valor, datetime):
        return valor

    try:
        fecha = pd.to_datetime(valor, errors="coerce", dayfirst=True)
        if pd.notna(fecha):
            return fecha.to_pydatetime()
    except Exception:
        pass

    return valor_str
