"""Diagnostico de conexion a Supabase. Ejecutar en la maquina con problemas."""

import sys
import os
from pathlib import Path

# Si esta congelado, buscar .env en el directorio del exe
if getattr(sys, "frozen", False):
    env_path = Path(sys.executable).parent / ".env"
    meipass = Path(sys._MEIPASS)
    print(f"[INFO] Ejecutable congelado")
    print(f"[INFO] Directorio exe: {Path(sys.executable).parent}")
    print(f"[INFO] MEIPASS: {meipass}")
    print(f"[INFO] .env en exe dir: {env_path} (existe: {env_path.exists()})")
    env_path2 = meipass / ".env"
    print(f"[INFO] .env en MEIPASS: {env_path2} (existe: {env_path2.exists()})")
else:
    env_path = Path(__file__).resolve().parent / ".env"
    print(f"[INFO] Modo desarrollo")
    print(f"[INFO] .env: {env_path} (existe: {env_path.exists()})")

# Cargar .env manualmente
if env_path.exists():
    from dotenv import load_dotenv

    load_dotenv(env_path)
    print(f"[OK] .env cargado")
else:
    print(f"[WARN] .env NO encontrado en {env_path}")

url = os.getenv("SUPABASE_URL")
key = os.getenv("SUPABASE_KEY")
print(f"[INFO] SUPABASE_URL: {url}")
print(f"[INFO] SUPABASE_KEY: {key[:20]}..." if key else "[INFO] SUPABASE_KEY: None")

# Test 1: DNS
import socket

try:
    host = url.replace("https://", "").replace("http://", "").split("/")[0]
    print(f"\n--- Test DNS: {host} ---")
    ip = socket.gethostbyname(host)
    print(f"[OK] DNS resuelve a: {ip}")
except Exception as e:
    print(f"[FAIL] DNS fallo: {e}")

# Test 2: HTTP
try:
    import httpx

    print(f"\n--- Test HTTP ---")
    r = httpx.get(url + "/rest/v1/", timeout=10)
    print(f"[OK] HTTP status: {r.status_code}")
except Exception as e:
    print(f"[FAIL] HTTP fallo: {e}")

# Test 3: Supabase client
try:
    print(f"\n--- Test Supabase client ---")
    from supabase import create_client

    client = create_client(url, key)
    result = client.table("users").select("username").limit(1).execute()
    print(f"[OK] Supabase conecta. Usuarios: {result.data}")
except Exception as e:
    print(f"[FAIL] Supabase fallo: {e}")

print("\n--- Fin del diagnostico ---")
input("Presiona Enter para cerrar...")
