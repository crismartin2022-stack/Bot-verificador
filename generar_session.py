"""Genera la TELEGRAM_SESSION (string session) para Railway.

⚠️ Correlo en TU COMPUTADORA, no en Railway: pide el código que Telegram
manda a la app / SMS.

    pip install telethon
    python generar_session.py

Si Telegram devuelve FloodWait: NO reintentes en loop. Esperá el tiempo que
indica (puede ser de horas) y volvé a correrlo una sola vez. Cada reintento
prematuro reinicia el contador.
"""
from __future__ import annotations

import os
import sys

from telethon import TelegramClient
from telethon.errors import (
    FloodWaitError,
    PhoneCodeInvalidError,
    SessionPasswordNeededError,
)
from telethon.sessions import StringSession

API_ID = int(os.environ.get("TELEGRAM_API_ID") or input("API ID: ").strip())
API_HASH = os.environ.get("TELEGRAM_API_HASH") or input("API HASH: ").strip()
PHONE = os.environ.get("TELEGRAM_PHONE") or input("Teléfono (+54911...): ").strip()


def main() -> None:
    with TelegramClient(StringSession(), API_ID, API_HASH) as cliente:
        try:
            cliente.start(phone=PHONE)
        except FloodWaitError as e:
            horas = e.seconds / 3600
            print(f"\n⛔ Telegram pide esperar {e.seconds} s (~{horas:.1f} h).")
            print("   No reintentes hasta que pase ese tiempo.")
            sys.exit(1)
        except (PhoneCodeInvalidError, SessionPasswordNeededError) as e:
            print(f"\n⛔ {type(e).__name__}: revisá el código o tu 2FA.")
            sys.exit(1)

        yo = cliente.get_me()
        print(f"\n✅ Sesión creada para {yo.first_name} (id {yo.id})")
        print("\nCopiá esto en Railway → Variables → TELEGRAM_SESSION:\n")
        print(cliente.session.save())
        print("\n⚠️ Tratalo como una contraseña: da acceso total a la cuenta.")


if __name__ == "__main__":
    main()
