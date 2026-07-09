from supabase import create_client, Client

SUPABASE_URL = "https://xythcnwjtbfwwcxxprbp.supabase.co"
SUPABASE_KEY = "sb_publishable_IHo7OuyqmYxFo3YQnj4sGw_rABAps5G"

_supabase_client: Client = None


def get_supabase() -> Client:

    global _supabase_client
    if _supabase_client is None:
        _supabase_client = create_client(SUPABASE_URL, SUPABASE_KEY)
    return _supabase_client
