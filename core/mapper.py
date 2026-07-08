"""
Cruce y mapeo de datos entre los tres orígenes de información:
- Packing Beta (Confirmación de planta)
- Termógrafo
- SharePoint (datos complementarios del cultivo)

El resultado es una lista de diccionarios, donde cada elemento
representa una fila de la plantilla del cliente lista para escribir.

Campo de cruce principal: orden_beta (ej: 'BP-493')
Agrupación: por id_pallet (cada pallet = una fila en la plantilla)
"""

import pandas as pd
import re
from typing import Optional

# ─── Mapeo hacia plantilla TESCO (palta) ───────────────────────────
# Cada clave es el nombre de columna en la plantilla del cliente
# El valor indica de dónde viene el dato: (fuente, campo_interno)
# Fuentes: "packing", "termografo", "sharepoint", "fijo"

MAPEO_TESCO_PALTA = {
    "Supplier name": ("fijo", "COMPLEJO AGROINDUSTRIAL BETA S.A."),
    "Vessel name": ("sharepoint", "Nave"),
    "Container number": ("sharepoint", "AWB - Contenedor"),
    "Customer name (Tesco/Westfalia)": ("sharepoint", "Cliente"),
    "ETD": ("sharepoint", "ETD Programada"),
    "ETA": ("sharepoint", "ETA Programado"),
    "Destination port": ("sharepoint", "Puerto Descarga"),
    "Supplier reference": ("sharepoint", "Orden Beta"),
    "Variety": ("packing", "variedad"),
    "Weight (kg)": ("packing", "peso_caja"),
    "Size": ("packing", "calibre"),
    "Class": ("packing", "categoria"),
    "Country of Origin": ("fijo", "PERU"),
    "Number of boxes": ("packing", "total_cajas"),
    "Pallet quantity": ("fijo", "1"),
    "Organic / Conventional": ("fijo", "CONVENTIONAL"),
    "Pallet ID": ("packing", "id_pallet"),
    "Global Gap Number": ("fijo", "4050373153151"),
    "Grower Code": ("packing", "codigo_lugar_produccion"),
    "Temp Tale Location": ("fijo", ""),
    "Ethylene Absorber Location": ("fijo", ""),
    "Supplier Packhouse": ("packing", "lote_fundo"),
    "Pack Date DD/MM/YYYY": ("packing", "fecha_proceso"),
    "": ("termografo", "codigo_termografo"),
}

# Registro de mapeos disponibles por cliente y cultivo
MAPEOS_DISPONIBLES = {
    ("Westfalia Fruit", "Palta"): MAPEO_TESCO_PALTA,
    ("Tesco", "Palta"): MAPEO_TESCO_PALTA,
    # Aquí se irán agregando los demás clientes y cultivos:
    # ("Tesco", "Palta"):     MAPEO_TESCO_PALTA,
    # ("Comexa", "Uva"):      MAPEO_COMEXA_UVA,
}


# ─── Función principal ────────────────────────────────────────────────────────


def cruzar_datos(
    df_packing: pd.DataFrame,
    df_termografo: pd.DataFrame,
    datos_sp: Optional[dict],
    cliente: str,
    cultivo: str,
    orden_beta: str,
) -> list[dict]:
    """
    Cruza los tres orígenes de datos y retorna una lista de filas
    listas para escribir en la plantilla del cliente.

    Args:
        df_packing:    DataFrame del packing Beta (ya leído y normalizado)
        df_termografo: DataFrame del termógrafo (ya filtrado por orden_beta)
        datos_sp:      Diccionario con datos complementarios de SharePoint
                       (puede ser None si SP no está disponible aún)
        cliente:       Nombre del cliente, ej: 'Westfalia Fruit'
        cultivo:       Cultivo del pedido, ej: 'Palta'
        orden_beta:    Orden a procesar, ej: 'BP-493'

    Returns:
        Lista de dicts, cada uno representa una fila de la plantilla
    """
    mapeo = _obtener_mapeo(cliente, cultivo)

    # Filtrar packing por orden Beta
    df_pack_filtrado = _filtrar_por_orden(df_packing, orden_beta)

    # Cruzar packing con termógrafo por id_pallet
    df_cruzado = _cruzar_por_pallet(df_pack_filtrado, df_termografo)

    # Construir filas para la plantilla
    filas = []
    pallet_ids = []  # Guardar también los pallet_ids para procesamiento posterior
    for _, fila in df_cruzado.iterrows():
        fila_plantilla = _construir_fila(fila, mapeo, datos_sp or {})
        filas.append(fila_plantilla)
        pallet_ids.append(fila.get("id_pallet"))

    # ─── Opción 1: Código de termógrafo solo en la primera fila de cada pallet ───
    # Mantener el código de termógrafo solo en la primera ocurrencia de cada pallet
    pallet_visto = set()
    for i, pallet_id in enumerate(pallet_ids):
        if pallet_id in pallet_visto:
            # Limpiar el código de termógrafo en filas duplicadas del mismo pallet
            filas[i][""] = None
        else:
            if pallet_id is not None and not pd.isna(pallet_id):
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

    Returns:
        dict con totales y advertencias
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


def _obtener_mapeo(cliente: str, cultivo: str) -> dict:
    """Busca el mapeo correspondiente al cliente y cultivo seleccionados."""
    clave = (cliente, cultivo)
    if clave not in MAPEOS_DISPONIBLES:
        claves_disponibles = list(MAPEOS_DISPONIBLES.keys())
        raise ValueError(
            f"No existe plantilla configurada para '{cliente}' — '{cultivo}'. "
            f"Plantillas disponibles: {claves_disponibles}"
        )
    return MAPEOS_DISPONIBLES[clave]


def _filtrar_por_orden(df: pd.DataFrame, orden_beta: str) -> pd.DataFrame:
    """Filtra el DataFrame de packing por orden Beta."""
    orden = _normalizar_orden(orden_beta)
    serie_ordenes = df["orden_beta"].fillna("").astype(str).map(_normalizar_orden)
    df_filtrado = df[serie_ordenes == orden].copy()

    if df_filtrado.empty:
        raise ValueError(
            f"No se encontraron registros para la orden '{orden}' "
            f"en el archivo de packing Beta."
        )
    return df_filtrado.reset_index(drop=True)


def _cruzar_por_pallet(
    df_packing: pd.DataFrame,
    df_termografo: pd.DataFrame,
) -> pd.DataFrame:
    """
    Hace un LEFT JOIN entre packing y termógrafo usando id_pallet.
    Se usa LEFT para no perder pallets del packing aunque no tengan termógrafo.
    Los campos del termógrafo quedan con sufijo _termo para evitar colisiones.
    """
    df_cruzado = df_packing.merge(
        df_termografo.rename(columns={"pallet": "id_pallet"}),
        on="id_pallet",
        how="left",
        suffixes=("", "_termo"),
    )
    return df_cruzado


def _construir_fila(
    fila: pd.Series,
    mapeo: dict,
    datos_sp: dict,
) -> dict:
    """
    Construye un diccionario con los valores de la plantilla
    a partir de una fila del DataFrame cruzado.

    Prioridad de fuentes:
    1. "fijo"       → valor constante definido en el mapeo
    2. "packing"    → columna del packing Beta
    3. "termografo" → columna del termógrafo (puede tener sufijo _termo si colisiona)
    4. "sharepoint" → clave en el dict datos_sp
    """
    fila_plantilla = {}

    for columna_plantilla, (fuente, campo) in mapeo.items():

        if fuente == "fijo":
            valor = campo

        elif fuente == "packing":
            valor = _obtener_valor(fila, campo)

        elif fuente == "termografo":
            # Intentar con nombre directo y con sufijo _termo
            valor = _obtener_valor(fila, campo)
            if valor is None:
                valor = _obtener_valor(fila, f"{campo}_termo")

        elif fuente == "sharepoint":
            valor = datos_sp.get(campo)

        else:
            valor = None

        fila_plantilla[columna_plantilla] = valor

    return fila_plantilla


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
