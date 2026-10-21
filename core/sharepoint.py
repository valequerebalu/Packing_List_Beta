from __future__ import annotations

import logging
import os
import re
import unicodedata
from datetime import datetime
from pathlib import Path
from urllib.parse import unquote, urlparse

import pandas as pd

from config.settings import SHAREPOINT_LOCAL_DIR

# Nombre de la carpeta del negocio en OneDrive
_NEGOCIO = "Complejo Agroindustrial Beta S.A"
_SUBCARPETA = "Intranet - Reporte Exportaciones"
_SUBCARPETA_PEDIDOS = "Intranet - Cuadros de Pedidos"


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

    file_name = (
        f"Reporte Embarque de Exportaciones - "
        f"{crop_name} {datetime.now().year}.xlsx"
    )
    return get_sharepoint_resource_dataframe(
        {
            "key": "embarque",
            "file_name": file_name,
            "sheet_name": "Base de datos",
            "header_row": header_row,
        },
        crop_name=crop_name,
    )


def get_sharepoint_resource_dataframe(
    resource: dict,
    crop_name: str,
) -> pd.DataFrame:
    """Carga un recurso configurado en templates desde SharePoint sincronizado."""
    file_name = _resolver_nombre_recurso(resource, crop_name)
    subfolder = resource.get("subfolder")
    resource_key = resource.get("key") or resource.get("name")
    if subfolder is None and str(resource_key).lower() == "pedido":
        subfolder = f"{_SUBCARPETA_PEDIDOS}\\{crop_name.upper()}"
    elif subfolder:
        subfolder = str(subfolder).format(
            crop=crop_name,
            year=datetime.now().year,
        )

    full_path = _buscar_archivo_sharepoint(file_name, subfolder)

    sheet_name = resource.get("sheet_name", "Base de datos")
    header_row = resource.get("header_row", 0)

    logging.debug("SharePoint leyendo: %s", full_path)

    try:
        df = pd.read_excel(full_path, sheet_name=sheet_name, header=header_row)
        logging.info(
            "Recurso SharePoint '%s': hoja=%s header_row=%s columnas=%s filas=%d",
            file_name,
            sheet_name,
            header_row,
            list(df.columns),
            len(df),
        )
        columnas_normalizadas = {_normalizar_columna(col) for col in df.columns}
        if "orden_beta" not in columnas_normalizadas:
            raise ValueError(
                f"El archivo '{file_name}' no tiene la columna de orden Beta "
                f"después de leer header_row={header_row}. "
                "Verifica que header_row use índice base cero."
            )
        return df

    except ValueError as exc:
        if f"Worksheet named '{sheet_name}' not found" in str(exc):
            libro = pd.ExcelFile(full_path)
            raise ValueError(
                f"El archivo '{file_name}' no tiene pestaña '{sheet_name}'.\n"
                f"Pestañas disponibles: {libro.sheet_names}"
            )
        raise


def cargar_recursos_sharepoint(
    resources: list[dict] | dict | None,
    crop_name: str,
    orden_beta: str,
    cliente_seleccionado: str | None = None,
) -> tuple[dict[str, dict], list[str]]:
    """
    Carga y filtra los recursos declarados por una plantilla.

    Retorna:
        recursos: key -> {"datos": dict, "filas": list[dict]}
        advertencias: recursos opcionales que no pudieron cargarse
    """
    if not resources:
        return {}, []
    if isinstance(resources, dict):
        resources = resources.get("items", resources.get("resources", []))
    if not isinstance(resources, list):
        raise ValueError("La configuración 'resources' debe ser una lista JSON.")

    cargados: dict[str, dict] = {}
    advertencias: list[str] = []
    for resource in resources:
        if not isinstance(resource, dict):
            raise ValueError("Cada recurso de SharePoint debe ser un objeto JSON.")
        if resource.get("enabled", True) is False:
            continue

        key = resource.get("key") or resource.get("name") or resource.get("id")
        if not key:
            raise ValueError("Cada recurso de SharePoint debe tener 'key' o 'name'.")
        key = str(key).strip().lower()

        try:
            df = get_sharepoint_resource_dataframe(resource, crop_name)
            datos = extraer_datos_sharepoint(df, orden_beta, cliente_seleccionado)
            filas = extraer_filas_sharepoint(df, orden_beta, cliente_seleccionado)
            logging.info(
                "Recurso SharePoint '%s': columnas=%s filas_orden=%d datos=%d",
                key,
                list(df.columns),
                len(filas),
                len(datos),
            )
            cargados[key] = {"datos": datos, "filas": filas}
        except (FileNotFoundError, ValueError) as exc:
            if resource.get("required", True):
                raise
            advertencias.append(f"{key}: {exc}")

    return cargados, advertencias


def _resolver_nombre_recurso(resource: dict, crop_name: str) -> str:
    """Resuelve file_name o file_pattern usando cultivo y año actual."""
    nombre = (
        resource.get("file_name")
        or resource.get("file_pattern")
        or Path(unquote(urlparse(str(resource.get("url", ""))).path)).name
    )
    if not nombre:
        raise ValueError("El recurso de SharePoint requiere 'file_name'.")
    return str(nombre).format(crop=crop_name, year=datetime.now().year)


def _buscar_archivo_sharepoint(
    file_name: str,
    subfolder: str | None = None,
) -> Path:
    roots = _buscar_rutas_sharepoint(subfolder or _SUBCARPETA)

    for root in roots:
        candidate = root / file_name
        if candidate.exists():
            return candidate

    rutas = "\n".join(f"  - {r}" for r in roots)
    raise FileNotFoundError(
        f"No se encontró '{file_name}'.\n\nRutas buscadas:\n{rutas}"
    )


def _buscar_rutas_sharepoint(subfolder: str = _SUBCARPETA) -> list[Path]:
    """
    Escanea el sistema buscando la carpeta de SharePoint del negocio.
    Funciona en cualquier máquina sin configuración previa.
    """
    roots: list[Path] = []

    # 1. Override explícito vía env var (si se configuró)
    if SHAREPOINT_LOCAL_DIR and SHAREPOINT_LOCAL_DIR.exists():
        if subfolder == _SUBCARPETA:
            roots.append(SHAREPOINT_LOCAL_DIR)
        else:
            roots.append(SHAREPOINT_LOCAL_DIR / subfolder)

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
                    candidata = carpeta / _NEGOCIO / subfolder
                    if candidata.exists():
                        roots.append(candidata)
                else:
                    # Búsqueda profunda: cualquier carpeta con el negocio adentro
                    candidata = carpeta / _NEGOCIO / subfolder
                    if candidata.exists():
                        roots.append(candidata)
        except PermissionError:
            continue

    # 3. Fallbacks en home del usuario
    user_home = Path(os.path.expanduser("~"))
    fallbacks = [
        user_home / _NEGOCIO / subfolder,
        user_home / "OneDrive" / _NEGOCIO / subfolder,
        user_home / "SharePoint" / subfolder,
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


def extraer_datos_sharepoint(
    df: pd.DataFrame,
    orden_beta: str | None = None,
    cliente_seleccionado: str | None = None,
) -> dict:
    """
    Normaliza una fila de SharePoint y retorna TODOS los campos como dict.
    El mapper puede acceder a cualquier campo por nombre de columna normalizado.

    Args:
        orden_beta: Orden a buscar (se strip sufijo letra para matching).
        cliente_seleccionado: Cliente seleccionado en el sistema para validar.
    """
    if df is None or df.empty:
        return {}

    normalizado = df.copy()
    normalizado.columns = [_normalizar_columna(c) for c in normalizado.columns]

    if orden_beta and "orden_beta" in normalizado.columns:
        orden_norm, _ = _normalizar_orden_sufijo(orden_beta)
        filtrado = normalizado[
            normalizado["orden_beta"]
            .astype(str)
            .map(lambda x: _normalizar_orden_sufijo(x)[0])
            == orden_norm
        ]
        if filtrado.empty:
            raise ValueError(
                f"No se encontró la orden '{orden_beta}' en el recurso SharePoint."
            )
        # Validar con cliente si se proporciona
        if cliente_seleccionado and "cliente" in filtrado.columns:
            filtrado_validado = filtrado[
                filtrado["cliente"].apply(
                    lambda c: _cliente_sharepoint_match(c, cliente_seleccionado)
                )
            ]
            if not filtrado_validado.empty:
                filtrado = filtrado_validado
        normalizado = filtrado
    elif orden_beta:
        raise ValueError(
            "El recurso SharePoint no contiene la columna normalizada "
            "'orden_beta'. Verifica el encabezado del archivo."
        )

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
    resultado["Cliente"] = _valor_alias(
        resultado, "cliente", "customer", "customer_name"
    )
    resultado["Nave"] = _valor_alias(resultado, "nave", "vessel", "vessel_name")
    resultado["ETD Programada"] = _valor_alias(
        resultado, "etd_programada", "etd", "etd_programado"
    )
    resultado["ETA Programado"] = _valor_alias(
        resultado, "eta_programado", "eta", "eta_programada"
    )
    resultado["Puerto Descarga"] = _valor_alias(
        resultado, "puerto_descarga", "destination_port", "puerto"
    )
    resultado["Puerto Carga"] = _valor_alias(
        resultado, "puerto_carga", "origin_port", "puerto"
    )
    resultado["Orden Beta"] = _valor_alias(
        resultado, "orden_beta", "orden", "order_beta"
    )
    aerolinea_naviera = _valor_alias(
        resultado,
        "aereolinea_naviera",
        "aerolinea___naviera",
        "naviera",
        "a_naviera",
    )
    resultado["aereolinea_naviera"] = aerolinea_naviera
    resultado["Aerolínea - Naviera"] = aerolinea_naviera

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
    resultado["Fecha de Llenado"] = _valor_alias(
        resultado, "fecha_de_llenado", "fecha_llenado"
    )
    resultado["# PO"] = _valor_alias(resultado, "po")
    resultado["Puerto Carga"] = _valor_alias(resultado, "puerto_carga", "origin_port")
    resultado["Tipo de Caja"] = _valor_alias(resultado, "tipo_caja")
    resultado["Etiqueta"] = _valor_alias(resultado, "etiqueta")

    return resultado


def extraer_filas_sharepoint(
    df: pd.DataFrame,
    orden_beta: str | None = None,
    cliente_seleccionado: str | None = None,
) -> list[dict]:
    """
    Similar a extraer_datos_sharepoint pero retorna TODAS las filas
    de la orden como lista de dicts.
    """
    if df is None or df.empty:
        return []

    normalizado = df.copy()
    normalizado.columns = [_normalizar_columna(c) for c in normalizado.columns]

    if orden_beta and "orden_beta" in normalizado.columns:
        orden_norm, _ = _normalizar_orden_sufijo(orden_beta)
        filtrado = normalizado[
            normalizado["orden_beta"]
            .astype(str)
            .map(lambda x: _normalizar_orden_sufijo(x)[0])
            == orden_norm
        ]
        if filtrado.empty:
            raise ValueError(
                f"No se encontró la orden '{orden_beta}' en el recurso SharePoint."
            )
        # Validar con cliente si se proporciona
        if cliente_seleccionado and "cliente" in filtrado.columns:
            filtrado_validado = filtrado[
                filtrado["cliente"].apply(
                    lambda c: _cliente_sharepoint_match(c, cliente_seleccionado)
                )
            ]
            if not filtrado_validado.empty:
                filtrado = filtrado_validado
        normalizado = filtrado
    elif orden_beta:
        raise ValueError(
            "El recurso SharePoint no contiene la columna normalizada "
            "'orden_beta'. Verifica el encabezado del archivo."
        )

    resultados = []
    for _, fila in normalizado.iterrows():
        resultado = {}
        for col, val in fila.items():
            if col.startswith("unnamed"):
                continue
            if pd.isna(val):
                resultado[col] = None
            else:
                texto = str(val).strip()
                resultado[col] = None if texto in ("", "nan", "None", "NaN") else texto
        resultados.append(resultado)

    return resultados


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


def _normalizar_orden_sufijo(valor: str) -> tuple[str, str]:
    """
    Normaliza una orden y extrae el sufijo de letra (ej: BAM0046A -> BAM0046, A).
    Retorna (orden_normalizada, sufijo).
    """
    texto = str(valor).strip().upper()
    partes = re.match(r"^([A-Z]+)-?(\d+)([A-Z])$", texto)
    if partes:
        prefijo, numero, sufijo = partes.groups()
        return f"{prefijo}{int(numero):04d}", sufijo
    return _normalizar_orden(valor), ""


# Aliases de cliente para validación en SharePoint
_CLIENTE_ALIASES = {
    "tesco": ["tesco", "westfalia fruit uk", "westfalia irlanda"],
    "united exports": ["united exports"],
    "walmart": ["walmart"],
    "global berry": ["global berry"],
}


def _cliente_sharepoint_match(
    cliente_sp: str | None, cliente_seleccionado: str
) -> bool:
    """
    Valida si el cliente de SharePoint coincide con el cliente seleccionado.
    Usa aliases para TESCO (acepta WESTFALIA FRUIT UK, WESTFALIA IRLANDA).
    Si no hay cliente en SharePoint, retorna True (sin validación).
    """
    if not cliente_sp:
        return True
    cliente_seleccionado_lower = cliente_seleccionado.strip().lower()
    aliases = _CLIENTE_ALIASES.get(
        cliente_seleccionado_lower, [cliente_seleccionado_lower]
    )
    cliente_sp_lower = cliente_sp.strip().lower()
    return any(alias in cliente_sp_lower for alias in aliases)
