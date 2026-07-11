"""
Fuentes:
- Packing Beta (Confirmación de planta): encabezado en fila 6, datos desde fila 7
- Termógrafo: encabezado en fila 2, datos desde fila 3
  → se filtra por ORDEN BET (ej: BP-493) para obtener solo el pedido actual

Campo de cruce entre ambos archivos: ORDEN BET (ej: BP-493)
"""

import pandas as pd
import re
from pathlib import Path

# ─── Mapeo de columnas ────────────────────────────────────────────────────────
# Nombre real en Excel → nombre interno normalizado

COLUMNAS_PACKING = {
    "CULTIVO": "cultivo",
    "MODALIDAD TRANSPORTE": "modalidad_transporte",
    "CENTRO": "centro",
    "CLIENTE": "cliente",
    "PAIS DESTINO": "pais_destino",
    "PO": "po",
    "ORDEN BETA": "orden_beta",
    "ID PALLET (HU)": "id_pallet",
    "FECHA COSECHA": "fecha_cosecha",
    "FECHA PROCESO": "fecha_proceso",
    "LOTE JULIANO": "lote_juliano",
    "LOTE FUNDO": "lote_fundo",
    "LOTE CLIENTE (OGL)": "lote_cliente_ogl",
    "CODIGO TRAZABILIDAD": "codigo_trazabilidad",
    "CODIGO LUGAR PRODUCCION": "codigo_lugar_produccion",
    "MARCA DISTINTIVA CAJA": "marca_distintiva_caja",
    "TIPO PALLET": "tipo_pallet",
    "TIPO ESQUINERO": "tipo_esquinero",
    "VARIEDAD": "variedad",
    "TIPO CAJA": "tipo_caja",
    "CATEGORIA": "categoria",
    "CALIBRE": "calibre",
    "PESO CAJA": "peso_caja",
    "PRESENTACION": "presentacion",
    "TOTAL CAJAS": "total_cajas",
    "TOTAL KILOS": "total_kilos",
}

COLUMNAS_TERMOGRAFO = {
    "N°": "numero",
    "FECHA DE EMBARQUE": "fecha_embarque",
    "SEMANA EMBARQUE": "semana_embarque",
    "MODALIDAD TRANSPORTE": "modalidad_transporte",
    "PLANTA": "planta",
    "ORDEN BETA": "orden_beta",
    "N° CONTENEDOR": "numero_contenedor",
    "BOOKING": "booking",
    "CLIENTE": "cliente",
    "DESTINO": "destino",
    "SET POINT (T°)": "set_point",
    "TIPO DE TERMOGRAFO": "tipo_termografo",
    "CODIGO DE TERMOGRAFO": "codigo_termografo",
    "PALLET": "pallet",
    "POSICION DE PALLET": "posicion_pallet",
    "VARIEDAD": "variedad",
    "PRESENTACIONES POR PALLET": "presentaciones_pallet",
    "TIPOS DE CAJA": "tipo_caja",
    "MARCA DE CAJA": "marca_caja",
    "TERMOGRAFO PRUEBA": "termografo_prueba",
}


# ─── Funciones principales ────────────────────────────────────────────────────


def leer_packing_beta(ruta: str) -> pd.DataFrame:
    """
    Lee el Excel de Confirmación de Planta Beta.
    Encabezado en fila 6 (index 5), datos desde fila 7.
    Elimina filas completamente vacías y normaliza columnas.
    """
    ruta = Path(ruta)
    _validar_archivo(ruta)

    df = pd.read_excel(
        ruta,
        header=5,  # fila 6 (0-indexado = 5)
        dtype=str,  # todo como texto para evitar conversiones automáticas
        engine="openpyxl",
    )

    df = _limpiar_dataframe(df)
    df = _renombrar_columnas(df, COLUMNAS_PACKING, fuente="Packing Beta")
    df = _normalizar_texto(df)

    return df


def leer_termografo(ruta: str, orden_beta: str) -> pd.DataFrame:
    """
    Lee el Excel de Termógrafo.
    Encabezado en fila 2 (index 1), datos desde fila 3.
    Filtra por ORDEN BET para obtener solo el pedido indicado.

    Args:
        ruta: Ruta al archivo de termógrafo
        orden_beta: Orden a filtrar, ej: 'BP-493'
    """
    ruta = Path(ruta)
    _validar_archivo(ruta)

    df = pd.read_excel(
        ruta, header=1, dtype=str, engine="openpyxl"  # fila 2 (0-indexado = 1)
    )

    df = _limpiar_dataframe(df)
    df = _renombrar_columnas(df, COLUMNAS_TERMOGRAFO, fuente="Termógrafo")
    df = _normalizar_texto(df)

    # Filtrar por orden Beta
    if "orden_beta" not in df.columns:
        raise ValueError(
            "No se encontró la columna 'ORDEN BET' en el archivo de termógrafo. "
            "Verifica que el archivo sea correcto."
        )

    orden_beta = _normalizar_orden(orden_beta)
    df_filtrado = df[
        df["orden_beta"].astype(str).map(_normalizar_orden) == orden_beta
    ].copy()

    if df_filtrado.empty:
        return pd.DataFrame(columns=list(COLUMNAS_TERMOGRAFO.values()))

    return df_filtrado.reset_index(drop=True)


def extraer_orden_beta_de_packing(df_packing: pd.DataFrame) -> list[str]:
    """
    Extrae los valores únicos de orden Beta presentes en el packing.
    Útil para mostrar en la GUI qué órdenes contiene el archivo cargado.
    """
    if "orden_beta" not in df_packing.columns:
        return []
    return sorted(df_packing["orden_beta"].dropna().unique().tolist())


def validar_cruce(df_packing: pd.DataFrame, df_termografo: pd.DataFrame) -> dict:
    """
    Verifica que ambos DataFrames tengan órdenes Beta coincidentes.
    Retorna un resumen del cruce para mostrar en el log de la GUI.
    """
    ordenes_packing = set(df_packing["orden_beta"].dropna().str.upper().str.strip())
    ordenes_termo = set(df_termografo["orden_beta"].dropna().str.upper().str.strip())

    return {
        "ordenes_packing": sorted(ordenes_packing),
        "ordenes_termo": sorted(ordenes_termo),
        "coincidencias": sorted(ordenes_packing & ordenes_termo),
        "faltantes_en_termo": sorted(ordenes_packing - ordenes_termo),
    }


# ─── Helpers privados ─────────────────────────────────────────────────────────


def _validar_archivo(ruta: Path):
    """Verifica que el archivo exista y sea un Excel válido."""
    if not ruta.exists():
        raise FileNotFoundError(f"No se encontró el archivo: {ruta}")
    if ruta.suffix.lower() not in (".xlsx", ".xls"):
        raise ValueError(f"El archivo debe ser .xlsx o .xls: {ruta.name}")


def _limpiar_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Elimina filas y columnas completamente vacías."""
    df = df.dropna(how="all")
    df = df.dropna(axis=1, how="all")
    df = df.reset_index(drop=True)
    return df


def _renombrar_columnas(df: pd.DataFrame, mapeo: dict, fuente: str) -> pd.DataFrame:
    """
    Renombra columnas según el mapeo definido.
    Ignora columnas del Excel que no estén en el mapeo (columnas extras).
    Lanza advertencia si faltan columnas esperadas.
    """
    # Normalizar nombres de columnas del Excel (quitar espacios extra)
    df.columns = [str(c).strip() for c in df.columns]

    columnas_presentes = set(df.columns)
    columnas_esperadas = set(mapeo.keys())
    columnas_faltantes = columnas_esperadas - columnas_presentes

    if columnas_faltantes:
        print(
            f"[WARN] {fuente}: columnas no encontradas en el archivo → "
            f"{sorted(columnas_faltantes)}"
        )

    # Solo renombrar las que existen
    mapeo_aplicable = {k: v for k, v in mapeo.items() if k in columnas_presentes}
    df = df.rename(columns=mapeo_aplicable)

    # Mantener solo columnas mapeadas
    columnas_internas = list(mapeo_aplicable.values())
    df = df[[c for c in columnas_internas if c in df.columns]]

    return df


def _normalizar_texto(df: pd.DataFrame) -> pd.DataFrame:
    """
    Limpia espacios en todas las columnas de texto.
    Reemplaza 'nan', 'None', 'NaN' por None real.
    """
    for col in df.columns:
        df[col] = (
            df[col]
            .astype(str)
            .str.strip()
            .replace({"nan": None, "None": None, "NaN": None, "": None})
        )
    return df


def _normalizar_orden(valor: str) -> str:
    texto = str(valor).strip().upper()
    partes = re.match(r"^([A-Z]+)-?(\d+)$", texto)
    if not partes:
        return re.sub(r"[^A-Z0-9]", "", texto)

    prefijo, numero = partes.groups()
    return f"{prefijo}{int(numero):04d}"
