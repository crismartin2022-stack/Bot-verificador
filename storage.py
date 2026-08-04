"""Persistencia en el Volume de Railway.

Dos capas separadas a propósito:

* historial.json  -> registro HISTÓRICO de todo comprobante ya procesado.
                     Nunca se borra con /reset. Es lo que permite detectar
                     que alguien reenvía un comprobante viejo.
* verificaciones/<chat_id>.json -> resultado de la verificación EN CURSO
                     de ese chat. Eso sí lo limpia /reset.
"""
from __future__ import annotations

import asyncio
import json
import logging
import os
import tempfile
from datetime import datetime, timezone
from pathlib import Path

log = logging.getLogger("storage")


def _ahora() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _leer_json(path: Path, default):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return default
    except (json.JSONDecodeError, OSError) as e:
        log.error("No se pudo leer %s (%s). Se usa valor por defecto.", path, e)
        try:
            path.replace(path.with_suffix(path.suffix + ".corrupto"))
        except OSError:
            pass
        return default


def _escribir_json(path: Path, data) -> None:
    """Escritura atómica: tmp + replace (sobrevive reinicios de Railway)."""
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=str(path.parent), suffix=".tmp")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=1)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, path)
    except Exception:
        try:
            os.unlink(tmp)
        except OSError:
            pass
        raise


class Storage:
    def __init__(self, historial_file: Path, verif_dir: Path):
        self.historial_file = Path(historial_file)
        self.verif_dir = Path(verif_dir)
        self._lock = asyncio.Lock()
        self.historial: dict = {"por_huella": {}, "por_archivo": {}, "actualizado": None}

    # ------------------------------------------------------------ historial
    async def cargar(self) -> None:
        data = await asyncio.to_thread(
            _leer_json, self.historial_file,
            {"por_huella": {}, "por_archivo": {}, "actualizado": None},
        )
        data.setdefault("por_huella", {})
        data.setdefault("por_archivo", {})
        self.historial = data
        log.info("Historial cargado: %d comprobantes únicos", len(data["por_huella"]))

    async def _flush(self) -> None:
        self.historial["actualizado"] = _ahora()
        await asyncio.to_thread(_escribir_json, self.historial_file, self.historial)

    def buscar_previo(self, huella: str | None, hash_img: str | None) -> dict | None:
        """Devuelve el registro original si este comprobante ya se procesó antes."""
        if hash_img and hash_img in self.historial["por_archivo"]:
            return self.historial["por_archivo"][hash_img]
        if huella and huella in self.historial["por_huella"]:
            return self.historial["por_huella"][huella]
        return None

    async def registrar(self, item: dict) -> None:
        """Guarda el comprobante en el historial global (si es nuevo)."""
        resumen = {
            "chat_id": item.get("chat_id"),
            "chat": item.get("chat"),
            "msg_id": item.get("msg_id"),
            "fecha_msg": item.get("fecha_msg"),
            "remitente": item.get("remitente"),
            "nombre": item.get("nombre_img") or item.get("nombre_pie"),
            "monto": item.get("monto_img") if item.get("monto_img") is not None else item.get("monto_pie"),
            "nro_operacion": item.get("nro_operacion"),
            "link": item.get("link"),
            "registrado": _ahora(),
        }
        async with self._lock:
            h, ha = item.get("huella"), item.get("hash_archivo")
            cambio = False
            if h and h not in self.historial["por_huella"]:
                self.historial["por_huella"][h] = resumen
                cambio = True
            if ha and ha not in self.historial["por_archivo"]:
                self.historial["por_archivo"][ha] = resumen
                cambio = True
            if cambio:
                await self._flush()

    def stats_historial(self) -> dict:
        return {
            "comprobantes": len(self.historial.get("por_huella", {})),
            "imagenes": len(self.historial.get("por_archivo", {})),
            "actualizado": self.historial.get("actualizado"),
        }

    # -------------------------------------------------------- verificaciones
    def _path_verif(self, chat_id: int) -> Path:
        return self.verif_dir / f"{chat_id}.json"

    async def guardar_verificacion(self, chat_id: int, data: dict) -> None:
        async with self._lock:
            await asyncio.to_thread(_escribir_json, self._path_verif(chat_id), data)

    async def cargar_verificacion(self, chat_id: int) -> dict | None:
        return await asyncio.to_thread(_leer_json, self._path_verif(chat_id), None)

    async def resetear_verificacion(self, chat_id: int) -> bool:
        """Borra el resultado en curso. El historial global NO se toca."""
        p = self._path_verif(chat_id)
        async with self._lock:
            if p.exists():
                archivo = p.with_name(f"{chat_id}.{datetime.now().strftime('%Y%m%d%H%M%S')}.bak")
                await asyncio.to_thread(p.replace, archivo)
                return True
        return False

    async def ultimas_verificaciones(self, limite: int = 10) -> list[dict]:
        archivos = sorted(
            self.verif_dir.glob("*.json"),
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )[:limite]
        salida = []
        for p in archivos:
            d = await asyncio.to_thread(_leer_json, p, None)
            if d:
                salida.append({
                    "chat_id": d.get("chat_id"),
                    "chat": d.get("chat"),
                    "fin": d.get("fin"),
                    "total": len(d.get("items", [])),
                })
        return salida
