import bcrypt
from db.database import get_supabase


def crear_usuario(username, password, full_name, email, is_admin=False):
    supabase = get_supabase()
    hash_pw = bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
    result = (
        supabase.table("users")
        .insert(
            {
                "username": username,
                "password_hash": hash_pw,
                "full_name": full_name,
                "email": email,
                "is_admin": is_admin,
                "is_active": True,
            }
        )
        .execute()
    )
    print(f"✅ Usuario '{username}' creado con ID {result.data[0]['id']}")


if __name__ == "__main__":
    crear_usuario(
        "username",
        "password",
        "Full Name",
        "admin@example.com",
        is_admin=False,
    )
