"""Configuración central del bot verificador de comprobantes."""
from __future__ import annotations

import os
from pathlib import Path
from zoneinfo import ZoneInfo


def _int(nombre: str, default: int) -> int:
    try:
        return int(os.environ.get(nombre, default))
    except (TypeError, ValueError):
        return default


def _float(nombre: str, default: float) -> float:
    try:
        return float(os.environ.get(nombre, default))
    except (TypeError, ValueError):
        return default


# --- Telegram (cuenta de usuario, Telethon) ---
API_ID = _int("TELEGRAM_API_ID", 0)
API_HASH = os.environ.get("TELEGRAM_API_HASH", "").strip()
PHONE = os.environ.get("TELEGRAM_PHONE", "").strip()
# String session generada con generar_session.py (NO subir al repo)
SESSION_STRING = os.environ.get("TELEGRAM_SESSION", "").strip()

# IDs de usuario autorizados a dar comandos, además de la propia cuenta.
ADMIN_IDS = {
    int(x) for x in os.environ.get("ADMIN_IDS", "").replace(" ", "").split(",") if x.strip("-").isdigit()
}

# --- Anthropic ---
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "").strip()
# claude-sonnet-5 = mejor lectura de comprobantes.
# claude-haiku-4-5-20251001 = más barato para volúmenes grandes.
MODELO = os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-5").strip()
MAX_TOKENS = _int("ANTHROPIC_MAX_TOKENS", 1200)

# --- Cloudinary (opcional; se puede compartir con otro bot) ---
CLOUDINARY_CLOUD_NAME = os.environ.get("CLOUDINARY_CLOUD_NAME", "").strip()
CLOUDINARY_API_KEY = os.environ.get("CLOUDINARY_API_KEY", "").strip()
CLOUDINARY_API_SECRET = os.environ.get("CLOUDINARY_API_SECRET", "").strip()
CLOUDINARY_FOLDER = os.environ.get("CLOUDINARY_FOLDER", "comprobantes").strip()

# --- Almacenamiento (Railway Volume montado en /data) ---
DATA_DIR = Path(os.environ.get("DATA_DIR", "/data"))
HISTORIAL_FILE = DATA_DIR / "historial.json"
VERIF_DIR = DATA_DIR / "verificaciones"
REPORTES_DIR = DATA_DIR / "reportes"
LOG_FILE = DATA_DIR / "bot.log"

# --- Parámetros de verificación ---
TZ = ZoneInfo(os.environ.get("TZ_LOCAL", "America/Argentina/Buenos_Aires"))
CONCURRENCIA = _int("MAX_CONCURRENCIA", 3)          # llamadas simultáneas a Claude
UMBRAL_NOMBRE = _float("UMBRAL_NOMBRE", 0.80)       # 0-1, similitud mínima de nombres
TOLERANCIA_MONTO = _float("TOLERANCIA_MONTO", 0.0)  # en pesos
CONFIANZA_MIN = _float("CONFIANZA_MIN", 0.50)       # confianza mínima del OCR
MAX_LADO_IMG = _int("MAX_LADO_IMG", 1568)           # px, redimensiona antes de enviar
INTERVALO_PROGRESO = _int("INTERVALO_PROGRESO", 15)  # seg entre updates de progreso
REINTENTOS_API = _int("REINTENTOS_API", 3)


def validar() -> None:
    """Falla temprano y claro si falta algo esencial."""
    faltan = []
    if not API_ID:
        faltan.append("TELEGRAM_API_ID")
    if not API_HASH:
        faltan.append("TELEGRAM_API_HASH")
    if not SESSION_STRING:
        faltan.append("TELEGRAM_SESSION")
    if not ANTHROPIC_API_KEY:
        faltan.append("ANTHROPIC_API_KEY")
    if faltan:
        raise SystemExit(
            "Faltan variables de entorno: " + ", ".join(faltan) +
            "\nGenerá la sesión con: python generar_session.py"
        )


def preparar_directorios() -> None:
    for d in (DATA_DIR, VERIF_DIR, REPORTES_DIR):
        d.mkdir(parents=True, exist_ok=True)
