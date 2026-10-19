"""Consultas y operaciones de persistencia vía Supabase REST API."""

from __future__ import annotations

import json
import logging
from hashlib import sha256
from pathlib import Path
from typing import Any

from db.database import get_supabase


def obtener_clientes() -> list[dict]:
    """Retorna lista de clientes con su configuración."""
    supabase = get_supabase()
    try:
        result = (
            supabase.table("customers")
            .select("name, requires_thermograph")
            .eq("is_active", True)
            .order("name")
            .execute()
        )
        return result.data
    except Exception:
        # Fallback si la columna requires_thermograph no existe aún
        result = (
            supabase.table("customers")
            .select("name")
            .eq("is_active", True)
            .order("name")
            .execute()
        )
        return [
            {"name": row["name"], "requires_thermograph": True} for row in result.data
        ]


def obtener_nombres_clientes() -> list[str]:
    """Retorna solo los nombres de clientes (backward-compatible)."""
    supabase = get_supabase()
    result = (
        supabase.table("customers")
        .select("name")
        .eq("is_active", True)
        .order("name")
        .execute()
    )
    return [row["name"] for row in result.data]


def obtener_cultivos() -> list[str]:
    supabase = get_supabase()
    result = (
        supabase.table("crops")
        .select("name")
        .eq("is_active", True)
        .order("name")
        .execute()
    )
    return [row["name"] for row in result.data]


# ─── LÓGICA DE INSERCIÓN Y BÚSQUEDA (API REST) ───────────────────────────────


def get_or_create_customer(code: str, name: str, country: str | None = None) -> dict:
    """Busca un cliente por código; si no existe, lo crea."""
    supabase = get_supabase()

    result = supabase.table("customers").select("*").eq("code", code).execute()
    if result.data:
        return result.data[0]

    nuevo_cliente = {"code": code, "name": name, "country": country}
    result = supabase.table("customers").insert(nuevo_cliente).execute()
    return result.data[0]


def record_generation(
    user_id: int,
    customer_id: int,
    template_id: int,
    packing_file_name: str,
    thermograph_file_name: str,
    packing_file_hash: str | None = None,
    thermograph_file_hash: str | None = None,
    sharepoint_source_id: int | None = None,
    output_file_name: str | None = None,
    output_file_path: str | None = None,
) -> dict:
    """Registra el inicio de una generación de Excel en estado 'pending'."""
    supabase = get_supabase()

    datos_generacion = {
        "user_id": user_id,
        "customer_id": customer_id,
        "template_id": template_id,
        "sharepoint_source_id": sharepoint_source_id,
        "packing_file_name": packing_file_name,
        "packing_file_hash": packing_file_hash,
        "thermograph_file_name": thermograph_file_name,
        "thermograph_file_hash": thermograph_file_hash,
        "output_file_name": output_file_name,
        "output_file_path": output_file_path,
        "status": "pending",
    }

    result = supabase.table("generations").insert(datos_generacion).execute()
    return result.data[0]


def save_generation_success(
    generation_id: int,
    output_file_name: str,
    output_file_path: str,
    duration_seconds: int | None = None,
) -> dict:
    """Actualiza una generación a estado 'success'."""
    supabase = get_supabase()

    datos_actualizados = {
        "status": "success",
        "output_file_name": output_file_name,
        "output_file_path": output_file_path,
        "duration_seconds": duration_seconds,
    }

    result = (
        supabase.table("generations")
        .update(datos_actualizados)
        .eq("id", generation_id)
        .execute()
    )
    return result.data[0]


def save_generation_error(
    generation_id: int, error_detail: str, duration_seconds: int | None = None
) -> dict:
    """Actualiza una generación a estado 'error'."""
    supabase = get_supabase()

    datos_actualizados = {
        "status": "error",
        "error_detail": error_detail,
        "duration_seconds": duration_seconds,
    }

    result = (
        supabase.table("generations")
        .update(datos_actualizados)
        .eq("id", generation_id)
        .execute()
    )
    return result.data[0]


def obtener_template(cliente: str, cultivo: str) -> dict | None:

    supabase = get_supabase()

    result = (
        supabase.table("templates")
        .select(
            "id, name, file_path, mapping_config, resources, "
            "customers(name), crops(name)"
        )
        .eq("status", "active")
        .execute()
    )

    for row in result.data:
        nombre_cliente = (row.get("customers") or {}).get("name", "")
        nombre_cultivo = (row.get("crops") or {}).get("name", "")
        if nombre_cliente == cliente and nombre_cultivo == cultivo:
            mc = row["mapping_config"]
            if isinstance(mc, str):
                mc = json.loads(mc)
            row["mapping_config"] = mc
            logging.debug("Template encontrado: id=%s name=%s", row["id"], row["name"])
            return row

    logging.warning(
        "No hay template activo para cliente='%s' cultivo='%s'", cliente, cultivo
    )
    return None


def actualizar_mapping_config(template_id: int, mapping_config: dict) -> dict:
    """Actualiza el mapping_config de un template en Supabase."""
    supabase = get_supabase()
    result = (
        supabase.table("templates")
        .update({"mapping_config": mapping_config})
        .eq("id", template_id)
        .execute()
    )
    return result.data[0]


# ─── UTILIDADES ──────────────────────────────────────────────────────────────


def sha256_file(file_path: str | Path) -> str:
    """Genera el hash SHA-256 de un archivo físico."""
    digest = sha256()
    with Path(file_path).open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()
