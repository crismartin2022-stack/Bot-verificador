"""Subida de comprobantes a Cloudinary (opcional).

Si no hay credenciales configuradas, no hace nada y el bot sigue igual.
El public_id es determinístico (`<chat>_<msg>`), así una re-verificación
sobrescribe la misma imagen en vez de duplicarla, y se puede compartir la
cuenta con otro bot sin pisarse.
"""
from __future__ import annotations

import asyncio
import hashlib
import logging
import time

import httpx

import config

log = logging.getLogger("nube")


class Cloudinary:
    def __init__(self):
        self.cloud = config.CLOUDINARY_CLOUD_NAME
        self.key = config.CLOUDINARY_API_KEY
        self.secret = config.CLOUDINARY_API_SECRET
        self.carpeta = config.CLOUDINARY_FOLDER
        self.activo = bool(self.cloud and self.key and self.secret)
        self.sem = asyncio.Semaphore(config.CONCURRENCIA)
        self.subidas = 0
        self.fallos = 0
        self._cliente: httpx.AsyncClient | None = None

    def _firma(self, params: dict) -> str:
        base = "&".join(f"{k}={params[k]}" for k in sorted(params))
        return hashlib.sha1((base + self.secret).encode("utf-8")).hexdigest()

    async def _http(self) -> httpx.AsyncClient:
        if self._cliente is None:
            self._cliente = httpx.AsyncClient(timeout=60)
        return self._cliente

    async def subir(self, data: bytes, chat_id, msg_id, nombre_archivo: str = "") -> str:
        """Devuelve la URL segura, o '' si está desactivado o falló."""
        if not self.activo or not data:
            return ""
        public_id = f"{str(chat_id).replace('-', 'n')}_{msg_id}"
        params = {
            "folder": self.carpeta,
            "overwrite": "true",
            "public_id": public_id,
            "timestamp": str(int(time.time())),
        }
        datos = dict(params)
        datos["api_key"] = self.key
        datos["signature"] = self._firma(params)

        url = f"https://api.cloudinary.com/v1_1/{self.cloud}/image/upload"
        for intento in range(3):
            try:
                async with self.sem:
                    cli = await self._http()
                    r = await cli.post(
                        url, data=datos,
                        files={"file": (nombre_archivo or f"{public_id}.jpg", data)},
                    )
                if r.status_code == 200:
                    self.subidas += 1
                    return r.json().get("secure_url", "")
                log.warning("Cloudinary %s: %s", r.status_code, r.text[:200])
                if r.status_code < 500:
                    break
            except Exception as e:
                log.warning("Cloudinary error (%d): %s", intento + 1, e)
            await asyncio.sleep(2 * (intento + 1))
        self.fallos += 1
        return ""

    async def cerrar(self):
        if self._cliente is not None:
            await self._cliente.aclose()
            self._cliente = None
