"""
Generación del Excel final a partir de la plantilla del cliente.

Toma las filas ya cruzadas y mapeadas (provenientes de core/mapper.py)
y las escribe directamente sobre una copia de la plantilla .xlsx del cliente,
respetando el formato original (fechas, dropdowns, encabezados).

Estructura confirmada de la plantilla (ej. Tesco/Westfalia — Sheet1):
- Fila 2 (D2:G3, merged): título "PACKING LIST"
- Fila 6: encabezados (columnas A a W)
- Fila 7 en adelante: datos
- Columna J (Weight kg) y L (Class): dropdowns ligados a hoja LookUps,
  pero se escribe el valor tal cual viene de Beta (ya validado en origen)
- Columnas E, F, W: formato de fecha (mm-dd-yy)
"""

from pathlib import Path
from datetime import datetime
import shutil

import openpyxl
import pandas as pd

# ─── Configuración de la plantilla ────────────────────────────────────────────

FILA_ENCABEZADO = 6
FILA_INICIO_DATOS = 7

# Orden exacto de columnas A→W tal como están en la plantilla Tesco
COLUMNAS_PLANTILLA = [
    "Supplier name",
    "Vessel name",
    "Container number",
    "Customer name (Tesco/Westfalia)",
    "ETD",
    "ETA",
    "Destination port",
    "Supplier reference",
    "Variety",
    "Weight (kg)",
    "Size",
    "Class",
    "Country of Origin",
    "Number of boxes",
    "Pallet quantity",
    "Organic / Conventional",
    "Pallet ID",
    "Global Gap Number",
    "Grower Code",
    "Temp Tale Location",
    "Ethylene Absorber Location",
    "Supplier Packhouse",
    "Pack Date DD/MM/YYYY",
    "",
]

# Columnas que requieren formato de fecha en la plantilla
COLUMNAS_FECHA = {"ETD", "ETA", "Pack Date DD/MM/YYYY"}

# Ruta a las plantillas por cliente
CARPETA_TEMPLATES = Path(__file__).parent.parent / "templates"

# Mapeo cliente+cultivo → nombre de archivo de plantilla
PLANTILLAS_POR_CLIENTE = {
    ("Tesco", "Palta"): "PLANTILLA_TESCO.xlsx",
    (
        "Westfalia Fruit",
        "Palta",
    ): "PLANTILLA_TESCO.xlsx",  # misma estructura de columnas
}

CARPETA_SALIDA = Path.home() / "Desktop" / "Beta Packing - Generados"


# ─── Función principal ────────────────────────────────────────────────────────


def generar_excel(
    filas: list[dict],
    cliente: str,
    cultivo: str,
    orden_beta: str,
) -> str:
    """
    Genera el archivo Excel final escribiendo las filas sobre una copia
    de la plantilla del cliente.

    Args:
        filas:      Lista de diccionarios (salida de core.mapper.cruzar_datos)
        cliente:    Nombre del cliente, ej: 'Tesco'
        cultivo:    Cultivo del pedido, ej: 'Palta'
        orden_beta: Orden procesada, ej: 'BP-493' (para nombrar el archivo)

    Returns:
        Ruta absoluta del archivo Excel generado
    """
    ruta_plantilla = _obtener_ruta_plantilla(cliente, cultivo)
    ruta_salida = _construir_ruta_salida(cliente, orden_beta)

    # Copiar la plantilla para no tocar el original
    CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)
    shutil.copy(ruta_plantilla, ruta_salida)

    wb = openpyxl.load_workbook(ruta_salida)
    ws = wb["Sheet1"] if "Sheet1" in wb.sheetnames else wb.active

    _escribir_filas(ws, filas)

    wb.save(ruta_salida)
    wb.close()

    return str(ruta_salida)


def listar_plantillas_disponibles() -> list[tuple[str, str]]:
    """Retorna la lista de combinaciones (cliente, cultivo) que tienen plantilla configurada."""
    return list(PLANTILLAS_POR_CLIENTE.keys())


# ─── Helpers privados ─────────────────────────────────────────────────────────


def _obtener_ruta_plantilla(cliente: str, cultivo: str) -> Path:
    """Busca el archivo de plantilla correspondiente al cliente y cultivo."""
    clave = (cliente, cultivo)
    nombre_archivo = PLANTILLAS_POR_CLIENTE.get(clave)

    if nombre_archivo is None:
        disponibles = list(PLANTILLAS_POR_CLIENTE.keys())
        raise ValueError(
            f"No hay plantilla configurada para '{cliente}' — '{cultivo}'. "
            f"Plantillas disponibles: {disponibles}"
        )

    ruta = CARPETA_TEMPLATES / nombre_archivo
    if not ruta.exists():
        raise FileNotFoundError(f"No se encontró el archivo de plantilla: {ruta}")

    return ruta


def _construir_ruta_salida(cliente: str, orden_beta: str) -> Path:
    """Genera un nombre de archivo único para el output."""
    # cliente_limpio = cliente.replace(" ", "_").replace("/", "-")
    orden = orden_beta.replace(" ", "")
    orden_limpia = str(orden).strip().upper()
    prefijo = "".join([caracter for caracter in orden_limpia if caracter.isalpha()])
    numeros = "".join([caracter for caracter in orden_limpia if caracter.isdigit()])
    if not numeros:
        return prefijo
    # timestamp = datetime.now().strftime("%Y%m%d_%H%M")

    nombre = f"{prefijo}{numeros.zfill(4)}-PACKING_LIST.xlsx"
    return CARPETA_SALIDA / nombre


def _escribir_filas(ws, filas: list[dict]):
    """
    Escribe cada fila del listado en la hoja, empezando en FILA_INICIO_DATOS.
    Respeta el orden exacto de columnas de la plantilla.
    Aplica formato de fecha a las columnas correspondientes.
    """
    fila_excel = FILA_INICIO_DATOS

    for fila_datos in filas:
        for idx, nombre_columna in enumerate(COLUMNAS_PLANTILLA, start=1):
            valor = fila_datos.get(nombre_columna)
            celda = ws.cell(row=fila_excel, column=idx)

            if nombre_columna in COLUMNAS_FECHA:
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

    # Intentar convertir a número (entero o decimal)
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
    Si no se puede convertir, retorna el valor original como texto.
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

    # Si ya viene como datetime (por ejemplo desde pandas), retornarlo tal cual
    if isinstance(valor, datetime):
        return valor

    try:
        fecha = pd.to_datetime(valor, errors="coerce", dayfirst=True)
        if pd.notna(fecha):
            return fecha.to_pydatetime()
    except Exception:
        pass

    # No se pudo convertir — devolver como texto para no perder el dato
    return valor_str
