from __future__ import annotations

from datetime import datetime, timezone
from typing import Any
import bcrypt
from db.database import get_supabase
from auth.session import update_session


class DBAuthError(Exception):
    pass


def authenticate(username: str, password: str) -> dict[str, Any]:
    supabase = get_supabase()

    result = supabase.table("users").select("*").eq("username", username).execute()

    if not result.data:
        raise DBAuthError("Usuario no encontrado.")

    user = result.data[0]

    if not user.get("is_active", True):
        raise DBAuthError("Usuario inactivo. Contacta al administrador.")

    if not bcrypt.checkpw(
        password.encode("utf-8"), user["password_hash"].encode("utf-8")
    ):
        raise DBAuthError("Contraseña incorrecta.")

    now = datetime.now(timezone.utc).isoformat()
    supabase.table("users").update({"last_login": now}).eq("id", user["id"]).execute()

    update_session(
        access_token=f"session_{user['id']}_{user['username']}",
        account=user,
        expires_at=None,
    )

    return user
