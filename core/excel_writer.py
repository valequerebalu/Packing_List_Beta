"""
Generación del Excel final a partir de la plantilla del cliente.

Toma las filas ya cruzadas y mapeadas (provenientes de core/mapper.py)
y las escribe directamente sobre una copia de la plantilla .xlsx del cliente,
respetando el formato original (fechas, dropdowns, encabezados).

El orden de columnas se obtiene leyendo los encabezados de la plantilla,
NO de un orden predefinido. Esto permite soportar plantillas con columnas
de fórmula que no se sobreescriben.
"""

from pathlib import Path
from datetime import datetime
import shutil
import re
import zipfile
import tempfile

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


    """
    mapping_config = template_data["mapping_config"]

    # Obtener configuración desde mapping_config
    sheet_name = mapping_config.get("sheet_name", "Sheet1")
    header_row = mapping_config.get("header_row", 6)
    data_start_row = mapping_config.get("data_start_row", 7)
    columns_config = mapping_config["columns"]

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

    # Extraer extLst del template ANTES de que openpyxl lo pierda
    extlst_block, worksheet_attrs = _extraer_extlst(ruta_plantilla)

    wb = openpyxl.load_workbook(ruta_salida)
    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active

    # Leer encabezados de la plantilla para mapear nombre → columna Excel
    header_map = _leer_encabezados(ws, header_row)

    _escribir_filas(ws, filas, columns_config, header_map, date_columns, data_start_row)

    wb.save(ruta_salida)
    wb.close()

    # Reinyectar extLst (data validations x14) que openpyxl descartó
    if extlst_block:
        _inyectar_extlst(ruta_salida, extlst_block, worksheet_attrs)

    return str(ruta_salida)


# ─── Preservación de extensiones x14 (data validations) ──────────────────────


def _extraer_extlst(xlsx_path: Path) -> tuple[str | None, str | None]:
    """
    Extrae el bloque <extLst> y las namespace declarations del sheet1.xml.
    openpyxl elimina los namespaces x14/xm/mc al reescribir el XML.

    Returns:
        (extlst_block, worksheet_root_attrs) o (None, None)
    """
    try:
        with zipfile.ZipFile(xlsx_path) as zf:
            xml = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
        match = re.search(r"<extLst>.*?</extLst>", xml, re.DOTALL)
        extlst = (
            match.group() if match and "x14:dataValidation" in match.group() else None
        )
        root_match = re.search(r"<worksheet([^>]*?)>", xml)
        root_attrs = root_match.group(1) if root_match else None
        return extlst, root_attrs
    except Exception:
        pass
    return None, None


def _inyectar_extlst(xlsx_path: Path, extlst_block: str, worksheet_attrs: str | None):
    """
    Reinyecta el bloque extLst y restaura las namespace declarations
    que openpyxl eliminó del tag <worksheet>.
    """
    tmp_dir = Path(tempfile.gettempdir())
    tmp_xlsx = tmp_dir / f"_extlst_patch_{id(extlst_block)}.xlsx"

    with zipfile.ZipFile(xlsx_path, "r") as zin, zipfile.ZipFile(
        tmp_xlsx, "w", zipfile.ZIP_DEFLATED
    ) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                xml = data.decode("utf-8")
                # Restaurar namespace declarations en <worksheet ...>
                if worksheet_attrs:
                    xml = re.sub(
                        r"<worksheet([^>]*?)>",
                        f"<worksheet{worksheet_attrs}>",
                        xml,
                    )
                # Inyectar extLst
                if "<extLst>" in xml:
                    xml = re.sub(
                        r"<extLst>.*?</extLst>",
                        extlst_block,
                        xml,
                        flags=re.DOTALL,
                    )
                else:
                    xml = xml.replace(
                        "</worksheet>",
                        extlst_block + "</worksheet>",
                    )
                data = xml.encode("utf-8")
            zout.writestr(item, data)

    shutil.move(str(tmp_xlsx), str(xlsx_path))


# ─── Helpers privados ─────────────────────────────────────────────────────────


def _leer_encabezados(ws, header_row: int) -> dict[str, int]:
    """
    Lee la fila de encabezados de la plantilla y retorna un dict
    que mapea nombre_de_columna → número_de_columna (1-indexed).
    Solo incluye columnas con valor no vacío.
    """
    header_map = {}
    for cell in ws[header_row]:
        if cell.value is not None:
            nombre = str(cell.value).strip()
            if nombre:
                header_map[nombre] = cell.column
    return header_map


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
    columns_config: dict,
    header_map: dict[str, int],
    date_columns: set[str],
    data_start_row: int,
):
    """
    Escribe cada fila del listado en la hoja, mapeando cada campo
    del mapping_config a la columna correcta de la plantilla (según header_map).
    Las columnas que no existen en la plantilla se ignoran.
    """
    fila_excel = data_start_row

    for fila_datos in filas:
        for nombre_columna, config in columns_config.items():
            # Determinar columna Excel: explícita en config o por nombre de encabezado
            col_idx = config.get("column") or header_map.get(nombre_columna)
            if col_idx is None:
                # Columna no existe en la plantilla — ignorar
                continue

            valor = fila_datos.get(nombre_columna)
            celda = ws.cell(row=fila_excel, column=col_idx)

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
