from __future__ import annotations

import logging
import sys
import threading

from config.settings import LOGS_DIR
from gui.app import BetaPackingApp


def _configurar_logging() -> None:
    LOGS_DIR.mkdir(exist_ok=True)
    log_file = LOGS_DIR / "app.log"

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(log_file, encoding="utf-8"),
        ],
    )


def _instalar_manejadores_globales() -> None:
    def _excepthook(exc_type, exc_value, exc_traceback):
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return

        logging.critical(
            "Excepción no controlada",
            exc_info=(exc_type, exc_value, exc_traceback),
        )

    sys.excepthook = _excepthook

    if hasattr(threading, "excepthook"):

        def _threading_excepthook(args):
            logging.critical(
                "Excepción no controlada en hilo %s",
                getattr(args.thread, "name", "sin-nombre"),
                exc_info=(args.exc_type, args.exc_value, args.exc_traceback),
            )

        threading.excepthook = _threading_excepthook


def main():
    _configurar_logging()
    _instalar_manejadores_globales()

    try:
        from db.database import get_supabase

        logging.info("Inicializando cliente API de Supabase...")
        supabase = get_supabase()

        # Hacemos un ping rápido e inofensivo a la tabla users para asegurar que responda
        supabase.table("users").select("id").limit(1).execute()
        logging.info("¡Conexión exitosa con Supabase en la nube!")

    except Exception:
        logging.exception("No se pudo conectar a la API de Supabase de forma correcta")

    app = BetaPackingApp()
    app.run()


if __name__ == "__main__":
    main()
