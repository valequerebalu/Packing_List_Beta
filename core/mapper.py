"""
Cruce y mapeo de datos entre los tres orígenes de información:
- Packing Beta (Confirmación de planta)
- Termógrafo
- SharePoint (datos complementarios del cultivo)

El mapeo de columnas ya no está hardcodeado aquí — se recibe como
mapping_config (dict) proveniente de la columna mapping_config de
la tabla templates en Supabase.
"""

import pandas as pd
import json
import re
from datetime import datetime
from typing import Optional


def cruzar_datos(
    df_packing: pd.DataFrame,
    df_termografo: pd.DataFrame,
    datos_sp: Optional[dict],
    cliente: str,
    cultivo: str,
    orden_beta: str,
    mapping_config: dict | str,
) -> list[dict]:
    """
    Cruza los tres orígenes de datos y retorna una lista de filas
    listas para escribir en la plantilla del cliente.
    """
    if isinstance(mapping_config, str):
        mapping_config = json.loads(mapping_config)

    columnas = mapping_config["columns"]

    df_pack_filtrado = _filtrar_por_orden(df_packing, orden_beta)
    df_cruzado = _cruzar_por_pallet(df_pack_filtrado, df_termografo)

    # Pre-calcula sumas de cajas por pallet para transforms sum_per_pallet
    sumas_pallet = {}
    total_cajas_global = 0
    if "total_cajas" in df_cruzado.columns and "id_pallet" in df_cruzado.columns:
        sumas_pallet = (
            df_cruzado.groupby("id_pallet")["total_cajas"]
            .apply(lambda s: sum(_int_seguro(v) for v in s))
            .to_dict()
        )
        total_cajas_global = sum(_int_seguro(v) for v in df_cruzado["total_cajas"])

    filas = []
    pallet_ids = []

    for _, fila in df_cruzado.iterrows():
        fila_plantilla = _construir_fila(fila, columnas, datos_sp or {}, sumas_pallet, total_cajas_global)
        filas.append(fila_plantilla)
        pallet_ids.append(fila.get("id_pallet"))

    # Columnas con only_first_per_pallet
    cols_first_only = [
        col for col, config in columnas.items() if config.get("only_first_per_pallet")
    ]
    if cols_first_only:
        pallet_visto = set()
        for i, pallet_id in enumerate(pallet_ids):
            if pallet_id in pallet_visto:
                for col in cols_first_only:
                    filas[i][col] = None
            else:
                if pallet_id is not None and str(pallet_id).strip() not in ("", "nan"):
                    pallet_visto.add(pallet_id)

    if not filas:
        raise ValueError(
            f"No se generaron filas para la orden '{orden_beta}' "
            f"— cliente: {cliente}, cultivo: {cultivo}. "
            f"Verifica que los archivos correspondan a este pedido."
        )

    return filas


def obtener_resumen_cruce(
    df_packing: pd.DataFrame,
    df_termografo: pd.DataFrame,
    orden_beta: str,
) -> dict:
    """
    Genera un resumen del cruce para mostrar en el log de la GUI
    antes de ejecutar la generación completa.
    """
    df_pack = _filtrar_por_orden(df_packing, orden_beta)
    pallets_packing = set(df_pack["id_pallet"].dropna())
    pallets_termo = set(df_termografo["pallet"].dropna())

    pallets_cruzados = pallets_packing & pallets_termo
    solo_en_packing = pallets_packing - pallets_termo
    solo_en_termo = pallets_termo - pallets_packing

    advertencias = []
    if solo_en_packing:
        advertencias.append(
            f"Pallets en packing sin termógrafo: {sorted(solo_en_packing)}"
        )
    if solo_en_termo:
        advertencias.append(
            f"Pallets en termógrafo sin packing: {sorted(solo_en_termo)}"
        )

    return {
        "orden_beta": orden_beta,
        "total_pallets": len(pallets_packing),
        "pallets_cruzados": len(pallets_cruzados),
        "advertencias": advertencias,
        "filas_packing": len(df_pack),
        "filas_termografo": len(df_termografo),
    }


# ─── Helpers privados ─────────────────────────────────────────────────────────


def _construir_fila(
    fila: pd.Series, columnas: dict, datos_sp: dict, sumas_pallet: dict, total_cajas_global: int
) -> dict:
    """
    Construye un diccionario con los valores de la plantilla
    a partir de una fila del DataFrame cruzado y el mapping_config de Supabase.
    """
    fila_plantilla = {}

    for col_plantilla, config in columnas.items():
        fuente = config["source"]

        if fuente == "fijo":
            valor = config.get("value", "")

        elif fuente == "packing":
            valor = _obtener_valor(fila, config["field"])

        elif fuente == "termografo":
            valor = _obtener_valor(fila, config["field"])
            if valor is None:
                valor = _obtener_valor(fila, f"{config['field']}_termo")

        elif fuente == "sharepoint":
            valor = datos_sp.get(config["field"])

        elif fuente == "transform":
            valor = _aplicar_transform(config, fila, datos_sp, sumas_pallet, total_cajas_global)

        else:
            valor = None

        fila_plantilla[col_plantilla] = valor

    return fila_plantilla


def _aplicar_transform(
    config: dict, fila: pd.Series, datos_sp: dict, sumas_pallet: dict, total_cajas_global: int
) -> str | None:
    """
    Aplica un transform definido en el mapping_config.

    Tipos soportados:
    - prefix:         Concatena un prefijo al valor de un campo
    - max_date:       Retorna la fecha mas reciente de multiples campos
    - conditional:    Evalua condiciones (contains/equals) y retorna un valor
    - fixed:          Retorna un valor constante
    - sum_per_pallet: Suma un campo agrupado por id_pallet
    - sum_cajas:      Suma total de cajas de todo el embarque
    """
    transform = config["transform"]
    tipo = transform["type"]

    if tipo == "prefix":
        valor = _obtener_valor_fuente(
            transform["field"], transform["source"], fila, datos_sp
        )
        if valor is not None:
            return transform["prefix"] + valor
        return None

    elif tipo == "max_date":
        fechas = []
        for field in transform["fields"]:
            valor = _obtener_valor_fuente(
                field, transform["source"], fila, datos_sp
            )
            if valor:
                for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y"):
                    try:
                        fechas.append(datetime.strptime(valor, fmt))
                        break
                    except ValueError:
                        continue
        if fechas:
            return str(max(fechas))
        return None

    elif tipo == "conditional":
        valor = _obtener_valor_fuente(
            transform["field"], transform["source"], fila, datos_sp
        )
        if valor is None:
            return transform.get("default")
        valor_lower = valor.lower()
        for cond in transform.get("conditions", []):
            if "contains" in cond and cond["contains"].lower() in valor_lower:
                return cond["result"]
            if "equals" in cond and cond["equals"].lower() == valor_lower:
                return cond["result"]
        return transform.get("default")

    elif tipo == "fixed":
        return transform.get("value")

    elif tipo == "sum_per_pallet":
        pallet_id = fila.get("id_pallet")
        if pallet_id is not None:
            return str(sumas_pallet.get(pallet_id, 0))
        return None

    elif tipo == "sum_cajas":
        return str(total_cajas_global) if total_cajas_global else None

    return None


def _obtener_valor_fuente(
    campo: str, fuente: str, fila: pd.Series, datos_sp: dict
) -> str | None:
    """Obtiene un valor de cualquier fuente (packing, sharepoint, termografo)."""
    if fuente == "packing":
        return _obtener_valor(fila, campo)
    elif fuente == "sharepoint":
        return datos_sp.get(campo)
    elif fuente == "termografo":
        valor = _obtener_valor(fila, campo)
        if valor is None:
            valor = _obtener_valor(fila, f"{campo}_termo")
        return valor
    return None


def _int_seguro(valor) -> int:
    """Convierte un valor a int de forma segura. Retorna 0 si no es valido."""
    if valor is None or pd.isna(valor):
        return 0
    texto = str(valor).strip().replace(",", "")
    try:
        return int(float(texto))
    except (ValueError, TypeError):
        return 0


def _filtrar_por_orden(df: pd.DataFrame, orden_beta: str) -> pd.DataFrame:
    """Filtra el DataFrame de packing por orden Beta con normalización."""
    orden = _normalizar_orden(orden_beta)
    serie_ordenes = df["orden_beta"].fillna("").astype(str).map(_normalizar_orden)
    df_filtrado = df[serie_ordenes == orden].copy()

    if df_filtrado.empty:
        raise ValueError(
            f"No se encontraron registros para la orden '{orden_beta}' "
            f"en el archivo de packing Beta."
        )
    return df_filtrado.reset_index(drop=True)


def _cruzar_por_pallet(
    df_packing: pd.DataFrame,
    df_termografo: pd.DataFrame,
) -> pd.DataFrame:
    """LEFT JOIN entre packing y termógrafo por id_pallet."""
    return df_packing.merge(
        df_termografo.rename(columns={"pallet": "id_pallet"}),
        on="id_pallet",
        how="left",
        suffixes=("", "_termo"),
    )


def _obtener_valor(fila: pd.Series, campo: str):
    """Obtiene un valor de la fila de forma segura. Retorna None si no existe o está vacío."""
    if campo not in fila.index:
        return None
    valor = fila[campo]
    if pd.isna(valor) or str(valor).strip() in ("", "nan", "None", "NaN"):
        return None
    return str(valor).strip()


def _normalizar_orden(valor: str) -> str:
    """Convierte órdenes como BP-493 y BP0493 a la misma forma canónica."""
    texto = str(valor).strip().upper()
    partes = re.match(r"^([A-Z]+)-?(\d+)$", texto)
    if not partes:
        return re.sub(r"[^A-Z0-9]", "", texto)
    prefijo, numero = partes.groups()
    return f"{prefijo}{int(numero):04d}"
