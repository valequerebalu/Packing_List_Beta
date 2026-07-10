from supabase import create_client, Client
from config.settings import SUPABASE_URL, SUPABASE_KEY

_supabase_client: Client = None


def get_supabase() -> Client:
    global _supabase_client
    if _supabase_client is None:
        if not SUPABASE_URL or not SUPABASE_KEY:
            raise RuntimeError(
                "Faltan las variables SUPABASE_URL y SUPABASE_KEY en el archivo .env"
            )
        _supabase_client = create_client(SUPABASE_URL, SUPABASE_KEY)
    return _supabase_client
