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
from datetime import date, datetime
from typing import Optional


def cruzar_datos(
    df_packing: pd.DataFrame,
    df_termografo: pd.DataFrame,
    datos_sp: Optional[dict],
    cliente: str,
    cultivo: str,
    orden_beta: str,
    mapping_config: dict | str,
    sharepoint_rows: Optional[list[dict]] = None,
    sharepoint_resources: Optional[dict[str, dict]] = None,
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
        fila_plantilla = _construir_fila(
            fila,
            columnas,
            datos_sp or {},
            sumas_pallet,
            total_cajas_global,
            df_termografo,
            sharepoint_rows or [],
            sharepoint_resources or {},
        )
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
    fila: pd.Series,
    columnas: dict,
    datos_sp: dict,
    sumas_pallet: dict,
    total_cajas_global: int,
    df_termografo: pd.DataFrame,
    sharepoint_rows: list[dict],
    sharepoint_resources: dict[str, dict],
) -> dict:
    """
    Construye un diccionario con los valores de la plantilla
    a partir de una fila del DataFrame cruzado y el mapping_config de Supabase.
    """
    fila_plantilla = {}

    for col_plantilla, config in columnas.items():
        fuente = config["source"]

        if fuente in ("fijo", "fixed"):
            valor = config.get("value", "")

        elif fuente == "packing":
            valor = _obtener_valor(fila, config["field"])

        elif fuente == "termografo":
            valor = _obtener_valor(fila, config["field"])
            if valor is None:
                valor = _obtener_valor(fila, f"{config['field']}_termo")

        elif fuente == "sharepoint":
            valor = datos_sp.get(config["field"])

        elif fuente.startswith("resource:"):
            recurso_key = fuente.split(":", 1)[1].strip().lower()
            recurso = sharepoint_resources.get(recurso_key, {})
            valor = recurso.get("datos", {}).get(config["field"])

        elif fuente == "transform":
            valor = _aplicar_transform(
                config,
                fila,
                datos_sp,
                sumas_pallet,
                total_cajas_global,
                df_termografo,
                sharepoint_rows,
                sharepoint_resources,
            )

        else:
            valor = None

        fila_plantilla[col_plantilla] = valor

    return fila_plantilla


def _aplicar_transform(
    config: dict,
    fila: pd.Series,
    datos_sp: dict,
    sumas_pallet: dict,
    total_cajas_global: int,
    df_termografo: pd.DataFrame = None,
    sharepoint_rows: list[dict] | None = None,
    sharepoint_resources: dict[str, dict] | None = None,
) -> str | None:
    """
    Aplica un transform definido en el mapping_config.

    Tipos soportados:
    - prefix:                     Concatena un prefijo al valor de un campo
    - max_date:                   Retorna la fecha mas reciente de multiples campos
    - conditional:                Evalua condiciones (contains/equals) y retorna un valor
    - fixed:                      Retorna un valor constante
    - sum_per_pallet:             Suma un campo agrupado por id_pallet
    - sum_cajas:                  Suma total de cajas de todo el embarque
    - value_map:                  Mapea valores de entrada a valores de salida
    - remove_spaces_hyphens:      Elimina espacios y guiones del valor de entrada
    - concat:                     Concatena partes fijas y valores condicionales
    - iso_week_day_code:          Retorna semana ISO y dia ISO como SS DD
    - multiply:                   Multiplica multiples campos
    - thermograph_per_pallet:     Busca termografo por pallet y retorna su valor
    - thermograph_aggregate:      Agrega valores de termografo para la orden
    - sharepoint_match_packing:   Busca en sharepoint por match con packing
    """
    transform = config["transform"]
    tipo = transform["type"]

    if tipo == "prefix":
        valor = _obtener_valor_fuente(
            transform["field"], transform["source"], fila, datos_sp,
            sharepoint_resources or {},
        )
        if valor is not None:
            return transform["prefix"] + valor
        return None

    elif tipo == "max_date":
        fechas = []
        for field in transform["fields"]:
            valor = _obtener_valor_fuente(
                field,
                transform["source"],
                fila,
                datos_sp,
                sharepoint_resources or {},
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
        return _evaluar_condicional(
            transform,
            fila,
            datos_sp,
            sharepoint_resources or {},
        )

    elif tipo == "fixed":
        return transform.get("value")

    elif tipo == "sum_per_pallet":
        pallet_id = fila.get("id_pallet")
        if pallet_id is not None:
            return str(sumas_pallet.get(pallet_id, 0))
        return None

    elif tipo == "sum_cajas":
        return str(total_cajas_global) if total_cajas_global else None

    elif tipo == "value_map":
        valor = _obtener_valor_fuente(
            transform["field"],
            transform.get("source", "packing"),
            fila,
            datos_sp,
            sharepoint_resources or {},
        )
        if valor is None:
            return transform.get("default")
        mapa = transform.get("map", {})
        valor_lower = valor.strip().lower()
        for clave, salida in mapa.items():
            if clave.lower() == valor_lower:
                return salida
        return transform.get("default")

    elif tipo == "remove_spaces_hyphens":
        valor = _obtener_valor_fuente(
            transform["field"],
            transform.get("source", "packing"),
            fila,
            datos_sp,
            sharepoint_resources or {},
        )
        if valor is None:
            return transform.get("default")
        return re.sub(r"[\s-]", "", valor)

    elif tipo == "concat":
        partes = []
        for parte in transform.get("parts", []):
            if "value" in parte:
                valor = parte["value"]
            elif "type" in parte:
                valor = _aplicar_transform(
                    {"transform": parte},
                    fila,
                    datos_sp,
                    sumas_pallet,
                    total_cajas_global,
                    df_termografo,
                    sharepoint_rows,
                    sharepoint_resources or {},
                )
            else:
                valor = _obtener_valor_fuente(
                    parte["field"],
                    parte["source"],
                    fila,
                    datos_sp,
                    sharepoint_resources or {},
                )

            if valor is not None and str(valor).strip():
                partes.append(str(valor).strip())

        return transform.get("separator", " ").join(partes) or transform.get(
            "default"
        )

    elif tipo == "iso_week_day_code":
        valor = _obtener_valor_fuente(
            transform["field"],
            transform.get("source", "packing"),
            fila,
            datos_sp,
            sharepoint_resources or {},
        )
        fecha = _parsear_fecha_transform(valor, transform.get("formats"))
        if fecha is None:
            return transform.get("default")

        calendario = fecha.isocalendar()
        return f"{calendario.week:02d}{calendario.weekday:02d}"

    elif tipo == "multiply":
        valores = []
        for field in transform["fields"]:
            v = _obtener_valor_fuente(
                field,
                transform.get("source", "packing"),
                fila,
                datos_sp,
                sharepoint_resources or {},
            )
            if v is None:
                return None
            try:
                valores.append(float(str(v).replace(",", "")))
            except (ValueError, TypeError):
                return None
        resultado = 1
        for v in valores:
            resultado *= v
        return str(resultado) if resultado else None

    elif tipo == "thermograph_per_pallet":
        pallet_id = _obtener_valor_fuente(
            transform["field"],
            transform.get("source", "packing"),
            fila,
            datos_sp,
            sharepoint_resources or {},
        )
        if pallet_id is None:
            return None
        thermograph_field = transform.get("thermograph_field", "codigo_termografo")
        if df_termografo is not None and not df_termografo.empty:
            match = df_termografo[df_termografo["pallet"] == pallet_id]
            if not match.empty:
                if "value" in transform:
                    return transform["value"]
                return _obtener_valor(match.iloc[0], thermograph_field)
        return None

    elif tipo == "thermograph_aggregate":
        thermograph_field = transform.get("field", "codigo_termografo")
        separator = transform.get("separator", "/")
        if df_termografo is None or df_termografo.empty:
            return None
        orden = fila.get("orden_beta")
        if orden is None:
            return None
        orden_norm = _normalizar_orden(orden)
        match = df_termografo[
            df_termografo["orden_beta"].fillna("").astype(str).map(_normalizar_orden)
            == orden_norm
        ]
        if match.empty:
            return None
        valores = []
        for _, termo_row in match.iterrows():
            val = _obtener_valor(termo_row, thermograph_field)
            if val and val not in valores:
                valores.append(val)
        return separator.join(valores) if valores else None

    elif tipo == "sharepoint_match_packing":
        packing_field = transform["packing_field"]
        sp_match_field = transform["sharepoint_match_field"]
        sp_return_field = transform["sharepoint_return_field"]
        match_patterns = transform.get("match_patterns", {})
        value_overrides = transform.get("value_overrides", {})

        packing_value = _obtener_valor(fila, packing_field)
        if packing_value is None:
            return None

        packing_lower = packing_value.strip().lower()

        # Determinar el valor de busqueda en sharepoint segun match_patterns
        search_value = None
        for pack_key, sp_value in match_patterns.items():
            if pack_key.lower() in packing_lower:
                search_value = sp_value.lower()
                break

        if search_value is None:
            # Sin patron de busqueda, usar el valor directo
            search_value = packing_lower

        # Buscar en sharepoint_rows la fila que haga match
        if sharepoint_rows:
            for sp_row in sharepoint_rows:
                sp_val = sp_row.get(sp_match_field)
                if sp_val is None:
                    continue
                sp_val_lower = str(sp_val).strip().lower()
                # Match: el valor de sharepoint contiene el search_value
                if search_value in sp_val_lower or sp_val_lower in search_value:
                    # Si hay value_overrides, usarlos
                    if value_overrides:
                        return value_overrides.get(
                            sp_val_lower, value_overrides.get(search_value)
                        )
                    return sp_row.get(sp_return_field)

        return None

    elif tipo == "description_uva":
        return ""

    return None


def _obtener_valor_fuente(
    campo: str,
    fuente: str,
    fila: pd.Series,
    datos_sp: dict,
    sharepoint_resources: dict[str, dict] | None = None,
) -> str | None:
    """Obtiene un valor de cualquier fuente (packing, sharepoint, termografo)."""
    if fuente == "packing":
        return _obtener_valor(fila, campo)
    elif fuente == "sharepoint":
        return datos_sp.get(campo)
    elif fuente.startswith("resource:"):
        recurso_key = fuente.split(":", 1)[1].strip().lower()
        recurso = (sharepoint_resources or {}).get(recurso_key, {})
        return recurso.get("datos", {}).get(campo)
    elif fuente == "termografo":
        valor = _obtener_valor(fila, campo)
        if valor is None:
            valor = _obtener_valor(fila, f"{campo}_termo")
        return valor
    return None


def _evaluar_condicional(
    config: dict,
    fila: pd.Series,
    datos_sp: dict,
    sharepoint_resources: dict[str, dict],
) -> str | None:
    """Evalúa una condición declarativa contra cualquiera de las fuentes."""
    valor = _obtener_valor_fuente(
        config["field"],
        config["source"],
        fila,
        datos_sp,
        sharepoint_resources,
    )
    if valor is None:
        return config.get("default")

    valor_lower = valor.strip().lower()
    for cond in config.get("conditions", []):
        if "contains" in cond and cond["contains"].lower() in valor_lower:
            return cond["result"]
        if "equals" in cond and cond["equals"].lower() == valor_lower:
            return cond["result"]
        if "starts_with" in cond and valor_lower.startswith(
            cond["starts_with"].lower()
        ):
            return cond["result"]

    else_config = config.get("else")
    if isinstance(else_config, dict):
        return _evaluar_condicional(
            else_config,
            fila,
            datos_sp,
            sharepoint_resources,
        )
    return config.get("default")


def _parsear_fecha_transform(
    valor: object,
    formatos: list[str] | None = None,
) -> date | None:
    """Convierte fechas de entrada del mapping a una fecha calendario."""
    if valor is None or pd.isna(valor):
        return None
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor

    texto = str(valor).strip()
    if not texto or texto.lower() in ("nan", "none", "nat"):
        return None

    formatos_intento = formatos or [
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M:%S",
    ]
    for formato in formatos_intento:
        try:
            return datetime.strptime(texto, formato).date()
        except ValueError:
            continue

    fecha = pd.to_datetime(texto, errors="coerce", dayfirst=True)
    return fecha.date() if pd.notna(fecha) else None


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
