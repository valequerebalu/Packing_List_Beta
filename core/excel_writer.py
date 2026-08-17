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
    datos_sp: dict | None = None,
    df_termografo: pd.DataFrame = None,
) -> str:
    """
    Genera el archivo Excel final escribiendo las filas sobre una copia
    de la plantilla del cliente.
    """
    mapping_config = template_data["mapping_config"]

    # Obtener configuracion desde mapping_config
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
            f"No se encontro el archivo de plantilla: {ruta_plantilla}"
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

    # Escribir metadata en celdas especificas (si existe)
    metadata_config = mapping_config.get("metadata", {})
    if metadata_config:
        _escribir_metadata(ws, metadata_config, filas, columns_config, datos_sp or {},
                           df_termografo, orden_beta)

    # Leer encabezados de la plantilla para mapear nombre a columna Excel
    header_map = _leer_encabezados(ws, header_row)

    _escribir_filas(ws, filas, columns_config, header_map, date_columns, data_start_row)

    wb.save(ruta_salida)
    wb.close()

    # Reinyectar extLst (data validations x14) que openpyxl descarto
    if extlst_block:
        _inyectar_extlst(ruta_salida, extlst_block, worksheet_attrs)

    # Marcar pivot caches con refreshOnLoad para que Excel los actualice al abrir
    _marcar_pivot_refresh(ruta_salida)

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


def _marcar_pivot_refresh(xlsx_path: Path):
    """
    Agrega refreshOnLoad="1" a todos los pivotCacheDefinition*.xml
    para que Excel actualice la tabla dinámica al abrir el archivo.
    """
    tmp_dir = Path(tempfile.gettempdir())
    tmp_xlsx = tmp_dir / f"_pivot_refresh_{id(xlsx_path)}.xlsx"

    with zipfile.ZipFile(xlsx_path, "r") as zin, zipfile.ZipFile(
        tmp_xlsx, "w", zipfile.ZIP_DEFLATED
    ) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if "pivotCacheDefinition" in item.filename and item.filename.endswith(".xml"):
                xml = data.decode("utf-8")
                if "refreshOnLoad" not in xml:
                    xml = xml.replace(
                        'createdVersion="',
                        'refreshOnLoad="1" createdVersion="',
                    )
                    data = xml.encode("utf-8")
            zout.writestr(item, data)

    shutil.move(str(tmp_xlsx), str(xlsx_path))


# ─── Helpers privados ─────────────────────────────────────────────────────────


def _leer_encabezados(ws, header_row: int) -> dict[str, int]:
    """
    Lee la fila de encabezados de la plantilla y retorna un dict
    que mapea nombre_de_columna a numero_de_columna (1-indexed).
    Solo incluye columnas con valor no vacio.
    """
    header_map = {}
    for cell in ws[header_row]:
        if cell.value is not None:
            nombre = str(cell.value).strip()
            if nombre:
                header_map[nombre] = cell.column
    return header_map


_MESES_ES = {
    1: "ene", 2: "feb", 3: "mar", 4: "abr", 5: "may", 6: "jun",
    7: "jul", 8: "ago", 9: "sep", 10: "oct", 11: "nov", 12: "dic",
}


def _escribir_metadata(
    ws,
    metadata_config: dict,
    filas: list[dict],
    columns_config: dict,
    datos_sp: dict,
    df_termografo: pd.DataFrame = None,
    orden_beta: str = None,
):
    """
    Escribe valores en celdas especificas definidas en metadata del mapping_config.
    Soporta sources: fijo, sharepoint, sum_cajas, transform.
    """
    # Pre-calcular suma total de cajas para sum_cajas
    total_cajas = 0
    for fila in filas:
        for col_name, cfg in columns_config.items():
            if cfg["source"] == "packing" and cfg.get("field") == "total_cajas":
                val = fila.get(col_name)
                if val:
                    try:
                        total_cajas += int(float(str(val).replace(",", "")))
                    except (ValueError, TypeError):
                        pass
            elif cfg["source"] == "transform":
                transform = cfg.get("transform", {})
                if transform.get("type") == "sum_cajas":
                    val = fila.get(col_name)
                    if val:
                        try:
                            total_cajas = int(float(str(val).replace(",", "")))
                        except (ValueError, TypeError):
                            pass

    for celda_ref, config in metadata_config.items():
        fuente = config["source"]

        if fuente == "fijo":
            valor = config.get("value", "")

        elif fuente == "sharepoint":
            field = config.get("field", "")
            valor = datos_sp.get(field)
            # Formato de fecha personalizado (dd-mmm)
            if valor is not None and config.get("format") == "dd-mmm":
                valor = _formatear_fecha_dd_mmm(valor)

        elif fuente == "sum_cajas":
            valor = str(total_cajas) if total_cajas else ""

        elif fuente == "transform":
            from core.mapper import _aplicar_transform
            fila_mock = pd.Series({"orden_beta": orden_beta}) if orden_beta else pd.Series()
            valor = _aplicar_transform(
                config, fila_mock, datos_sp, {}, total_cajas,
                df_termografo, [],
            )

        else:
            valor = None

        if valor is not None:
            ws[celda_ref] = valor


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
            # Determinar columna Excel: explícita en config, por header_map, o por letra (A,B,C...)
            col_idx = config.get("column")
            if col_idx is None:
                # Buscar por nombre de encabezado en la plantilla (ANTES de interpretar como letra)
                col_idx = header_map.get(nombre_columna)
            if col_idx is None:
                # Intentar convertir letra de columna a número (A=1, B=2, ..., Z=26, AA=27, ...)
                col_idx = _columna_letra_a_numero(nombre_columna)
            if col_idx is None:
                # Columna no existe en la plantilla — ignorar
                continue

            valor = fila_datos.get(nombre_columna)
            celda = ws.cell(row=fila_excel, column=col_idx)

            if nombre_columna in date_columns:
                celda.value = _convertir_fecha(valor)
                celda.number_format = "dd/mm/yyyy"
            elif config.get("format") == "text":
                celda.value = str(valor) if valor is not None else None
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


def _formatear_fecha_dd_mmm(valor) -> str | None:
    """
    Convierte un valor de fecha a formato 'dd-mmm' en español (ej: '03-ago').
    """
    if valor is None:
        return None
    if isinstance(valor, datetime):
        fecha = valor
    else:
        valor_str = str(valor).strip()
        if not valor_str:
            return None
        fecha = _convertir_fecha(valor_str)
        if isinstance(fecha, str):
            return None
    mes = _MESES_ES.get(fecha.month)
    if mes is None:
        return None
    return f"{fecha.day:02d}-{mes}"


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


def _columna_letra_a_numero(letra: str) -> int | None:
    """
    Convierte una letra de columna de Excel a número (1-indexed).
    A=1, B=2, ..., Z=26, AA=27, AB=28, etc.
    Retorna None si no es una letra de columna válida.
    Máximo Excel: XFD = 16384 (3 caracteres).
    """
    if not isinstance(letra, str):
        return None
    letra = letra.strip().upper()
    if not letra or not letra.isalpha():
        return None
    if len(letra) > 3:  # Máx columnas Excel = XFD (3 chars)
        return None
    resultado = 0
    for c in letra:
        resultado = resultado * 26 + (ord(c) - ord('A') + 1)
    return resultado
