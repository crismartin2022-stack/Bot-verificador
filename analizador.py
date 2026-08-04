"""Lectura de comprobantes con Claude (visión) y contraste contra el pie de foto."""
from __future__ import annotations

import asyncio
import base64
import io
import json
import logging
import re

from anthropic import AsyncAnthropic

import config
import utils

log = logging.getLogger("analizador")

SYSTEM = """Sos un extractor de datos de comprobantes bancarios ARGENTINOS
(Mercado Pago, Banco Nación, Galicia, Santander, BBVA, Brubank, Ualá, Naranja X,
Personal Pay, Cuenta DNI, Prex, Lemon, Belo, etc.).

Reglas:
- Respondé ÚNICAMENTE con un objeto JSON válido. Sin texto previo, sin ```.
- Los montos argentinos usan punto de miles y coma decimal: "1.234.567,89".
  Devolvé el monto SIEMPRE normalizado con punto decimal: "1234567.89".
- "monto" es el importe transferido/pagado. NUNCA el saldo de la cuenta,
  ni el disponible, ni comisiones, ni CVU/CBU.
- "nombre_destino" es quien RECIBE el dinero (Para / Destinatario / Destino).
  "nombre_origen" es quien envía (De / Origen / Titular de la cuenta).
- Si un dato no está visible, poné null. No inventes NADA.
- "confianza" refleja qué tan legible está la imagen (0 a 1).
- Si la imagen no es un comprobante de pago/transferencia, es_comprobante = false.

Formato exacto:
{
  "es_comprobante": true,
  "banco": "Mercado Pago",
  "tipo": "transferencia|pago|deposito|otro",
  "monto": "1234567.89",
  "moneda": "ARS",
  "fecha": "2026-08-04",
  "hora": "14:35",
  "nombre_origen": null,
  "nombre_destino": "JUAN PEREZ",
  "cuit_destino": null,
  "cvu_destino": null,
  "nro_operacion": null,
  "estado": "aprobado|pendiente|rechazado|desconocido",
  "confianza": 0.95,
  "observaciones": ""
}"""

USER_PROMPT = (
    "Extraé los datos de este comprobante. Respondé solo el JSON, sin explicaciones."
)

CAMPOS = (
    "es_comprobante banco tipo monto moneda fecha hora nombre_origen "
    "nombre_destino cuit_destino cvu_destino nro_operacion estado confianza observaciones"
).split()


def _preparar_imagen(data: bytes, mime: str) -> tuple[bytes, str]:
    """Redimensiona/convierte si hace falta (límites de la API de visión)."""
    try:
        from PIL import Image
    except ImportError:
        return data, mime if mime in ("image/jpeg", "image/png", "image/webp", "image/gif") else "image/jpeg"

    try:
        img = Image.open(io.BytesIO(data))
        img.load()
        cambio = False
        if img.mode not in ("RGB", "L"):
            img = img.convert("RGB")
            cambio = True
        lado = max(img.size)
        if lado > config.MAX_LADO_IMG:
            factor = config.MAX_LADO_IMG / lado
            img = img.resize((max(1, int(img.width * factor)), max(1, int(img.height * factor))), Image.LANCZOS)
            cambio = True
        if not cambio and len(data) < 3_500_000 and mime in ("image/jpeg", "image/png", "image/webp"):
            return data, mime
        buf = io.BytesIO()
        img.convert("RGB").save(buf, format="JPEG", quality=88, optimize=True)
        return buf.getvalue(), "image/jpeg"
    except Exception as e:  # imagen rara: la mandamos como vino
        log.warning("No se pudo preprocesar la imagen (%s)", e)
        return data, "image/jpeg"


def _extraer_json(texto: str) -> dict | None:
    t = texto.strip()
    t = re.sub(r"^```(?:json)?|```$", "", t, flags=re.MULTILINE).strip()
    ini, fin = t.find("{"), t.rfind("}")
    if ini == -1 or fin == -1:
        return None
    try:
        return json.loads(t[ini:fin + 1])
    except json.JSONDecodeError:
        return None


class Analizador:
    def __init__(self, api_key: str | None = None, modelo: str | None = None):
        self.cliente = AsyncAnthropic(api_key=api_key or config.ANTHROPIC_API_KEY)
        self.modelo = modelo or config.MODELO
        self.sem = asyncio.Semaphore(config.CONCURRENCIA)
        self.llamadas = 0

    async def leer_comprobante(self, data: bytes, mime: str = "image/jpeg") -> dict:
        """Devuelve el dict de campos. Si falla, {'es_comprobante': False, 'error': ...}."""
        img, media_type = _preparar_imagen(data, mime)
        b64 = base64.standard_b64encode(img).decode("ascii")

        ultimo_error = None
        for intento in range(config.REINTENTOS_API):
            try:
                async with self.sem:
                    resp = await self.cliente.messages.create(
                        model=self.modelo,
                        max_tokens=config.MAX_TOKENS,
                        system=SYSTEM,
                        messages=[{
                            "role": "user",
                            "content": [
                                {"type": "image", "source": {
                                    "type": "base64",
                                    "media_type": media_type,
                                    "data": b64,
                                }},
                                {"type": "text", "text": USER_PROMPT},
                            ],
                        }],
                    )
                self.llamadas += 1
                if getattr(resp, "stop_reason", None) == "refusal":
                    return {"es_comprobante": False, "error": "respuesta rechazada por el modelo", "confianza": 0}
                texto = "".join(b.text for b in resp.content if getattr(b, "type", "") == "text")
                datos = _extraer_json(texto)
                if datos is None:
                    ultimo_error = "respuesta no-JSON"
                    continue
                for c in CAMPOS:
                    datos.setdefault(c, None)
                datos["monto_num"] = utils.parse_monto(datos.get("monto"))
                try:
                    datos["confianza"] = float(datos.get("confianza") or 0)
                except (TypeError, ValueError):
                    datos["confianza"] = 0.0
                return datos
            except Exception as e:  # rate limit, red, 5xx...
                ultimo_error = f"{type(e).__name__}: {e}"
                espera = 2 ** intento * 2
                log.warning("Claude falló (intento %d): %s. Reintento en %ss", intento + 1, ultimo_error, espera)
                await asyncio.sleep(espera)

        return {"es_comprobante": False, "error": ultimo_error or "desconocido", "confianza": 0.0}


# ------------------------------------------------------------------ veredicto
ESTADOS = {
    "ok": "✅",
    "duplicado": "🔁",
    "monto": "❌",
    "nombre": "❌",
    "ambos": "❌",
    "ilegible": "❓",
    "sin_pie": "⚠️",
    "no_comprobante": "🚫",
}


def comparar(datos_img: dict, nombre_pie: str, monto_pie: float | None) -> dict:
    """Contrasta lo leído en la imagen contra lo declarado en el pie."""
    if not datos_img.get("es_comprobante"):
        return {
            "estado": "no_comprobante",
            "detalle": datos_img.get("error") or "la imagen no parece un comprobante",
            "sim_nombre": 0.0,
        }

    conf = float(datos_img.get("confianza") or 0)
    monto_img = datos_img.get("monto_num")
    nombre_img = datos_img.get("nombre_destino") or datos_img.get("nombre_origen") or ""

    if conf < config.CONFIANZA_MIN or monto_img is None:
        return {
            "estado": "ilegible",
            "detalle": f"confianza {conf:.2f}, monto leído: {utils.fmt_monto(monto_img)}",
            "sim_nombre": utils.similitud_nombres(nombre_img, nombre_pie),
        }

    if not nombre_pie and monto_pie is None:
        return {"estado": "sin_pie", "detalle": "el mensaje no trae nombre ni monto", "sim_nombre": 0.0}

    sim = utils.similitud_nombres(nombre_img, nombre_pie)
    ok_nombre = sim >= config.UMBRAL_NOMBRE if nombre_pie else True
    ok_monto = utils.montos_iguales(monto_img, monto_pie, config.TOLERANCIA_MONTO) if monto_pie is not None else True

    if ok_nombre and ok_monto:
        estado, detalle = "ok", "nombre y monto coinciden"
    elif not ok_nombre and not ok_monto:
        estado = "ambos"
        detalle = (f"imagen: {nombre_img} / {utils.fmt_monto(monto_img)} · "
                   f"pie: {nombre_pie} / {utils.fmt_monto(monto_pie)}")
    elif not ok_monto:
        estado = "monto"
        detalle = f"imagen {utils.fmt_monto(monto_img)} ≠ pie {utils.fmt_monto(monto_pie)}"
    else:
        estado = "nombre"
        detalle = f"imagen «{nombre_img}» ≠ pie «{nombre_pie}» (similitud {sim:.2f})"

    return {"estado": estado, "detalle": detalle, "sim_nombre": sim}
