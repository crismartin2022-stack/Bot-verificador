import os
import json
import asyncio
import base64
import logging
from datetime import datetime, timezone, timedelta

import anthropic
from telethon import TelegramClient, events
from telethon.tl.types import MessageMediaPhoto, MessageMediaDocument

logging.basicConfig(format="%(asctime)s | %(levelname)s | %(message)s", level=logging.INFO)
log = logging.getLogger(__name__)

# ── Config ────────────────────────────────────────────────────────────────────
API_ID       = int(os.environ["TELEGRAM_API_ID"])
API_HASH     = os.environ["TELEGRAM_API_HASH"]
PHONE        = os.environ["TELEGRAM_PHONE"]      # +54911...
ADMIN_ID     = 531707598
ANTHROPIC_KEY = os.environ["ANTHROPIC_API_KEY"]
ARG_TZ       = timezone(timedelta(hours=-3))

claude = anthropic.Anthropic(api_key=ANTHROPIC_KEY)

# Grupos a monitorear (chat_id negativos de Telegram)
GRUPOS_MONITOREAR = json.loads(os.environ.get("GRUPOS_VERIFICAR", "[]"))

# Session string (para no pedir código cada vez)
SESSION_STRING = os.environ.get("TELETHON_SESSION", "")

# ── Prompt de verificación ────────────────────────────────────────────────────
SYSTEM_PROMPT_VERIFICAR = """Eres un verificador de comprobantes bancarios argentinos.
Te doy una imagen de comprobante y el texto del pie que escribió el usuario.
Analizá el comprobante y respondé ÚNICAMENTE con JSON válido:
{
  "nombre_imagen": "nombre extraído de la imagen",
  "nombre_pie": "nombre del pie de texto",
  "monto_imagen": número o null,
  "coincide_nombre": true/false,
  "coincide_monto": true/false,
  "observaciones": "descripción de diferencias si las hay o vacío",
  "confianza": "alta/media/baja"
}
Sé flexible con variaciones menores de nombres (mayúsculas, orden de palabras, abreviaturas).
"""

async def analizar_comprobante(image_bytes: bytes, mime: str, pie: str) -> dict:
    """Analiza imagen y verifica contra el pie de texto."""
    b64 = base64.standard_b64encode(image_bytes).decode("utf-8")
    try:
        resp = await asyncio.get_event_loop().run_in_executor(
            None,
            lambda: claude.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=500,
                system=[{"type": "text", "text": SYSTEM_PROMPT_VERIFICAR, "cache_control": {"type": "ephemeral"}}],
                messages=[{"role": "user", "content": [
                    {"type": "image", "source": {"type": "base64", "media_type": mime, "data": b64}},
                    {"type": "text", "text": f"Pie del usuario: {pie}\n\nVerificá este comprobante."}
                ]}]
            )
        )
        text = resp.content[0].text
        return json.loads(text.replace("```json", "").replace("```", "").strip())
    except Exception as e:
        log.error(f"Error analizando: {e}")
        return {"error": str(e)}

async def verificar_historial(client, chat_id: int, limit: int = 1000):
    """Lee el historial de un grupo y verifica comprobantes."""
    log.info(f"Verificando historial de {chat_id} — últimos {limit} mensajes")
    
    resultados = []
    pie_pendiente = None
    foto_pendiente = None
    
    mensajes = []
    async for msg in client.iter_messages(chat_id, limit=limit):
        mensajes.append(msg)
    
    # Procesar en orden cronológico
    mensajes.reverse()
    
    for msg in mensajes:
        # Capturar texto (pie)
        if msg.text and not msg.from_id:
            continue
            
        # Foto con caption
        if msg.photo or (msg.document and msg.document.mime_type and msg.document.mime_type.startswith("image/")):
            caption = msg.message or ""
            if caption:
                # Descargar imagen
                try:
                    image_bytes = await client.download_media(msg, bytes)
                    mime = "image/jpeg"
                    resultado = await analizar_comprobante(image_bytes, mime, caption)
                    resultado["msg_id"] = msg.id
                    resultado["fecha"] = msg.date.strftime("%d/%m/%Y %H:%M")
                    resultado["pie"] = caption[:50]
                    resultados.append(resultado)
                    
                    if not resultado.get("coincide_nombre") or not resultado.get("coincide_monto"):
                        log.warning(f"⚠️ Discrepancia en msg #{msg.id}: {resultado.get('observaciones')}")
                    
                    await asyncio.sleep(1)  # Rate limit
                except Exception as e:
                    log.error(f"Error procesando msg #{msg.id}: {e}")
    
    return resultados

async def generar_reporte(resultados: list, chat_id: int) -> str:
    """Genera reporte de verificación."""
    total = len(resultados)
    errores = [r for r in resultados if not r.get("coincide_nombre") or not r.get("coincide_monto")]
    ok = total - len(errores)
    
    texto = f"📋 *Reporte de verificación*\n"
    texto += f"─────────────────────\n"
    texto += f"✅ Correctos: {ok}/{total}\n"
    texto += f"⚠️ Con discrepancias: {len(errores)}\n\n"
    
    if errores:
        texto += f"*Discrepancias encontradas:*\n"
        for r in errores[:20]:
            texto += f"\n• msg #{r.get('msg_id')} — {r.get('fecha','')}\n"
            texto += f"  Imagen: _{r.get('nombre_imagen','—')}_\n"
            texto += f"  Pie: _{r.get('nombre_pie','—')}_\n"
            if r.get('observaciones'):
                texto += f"  ⚠️ {r.get('observaciones')}\n"
        if len(errores) > 20:
            texto += f"\n... y {len(errores)-20} más"
    
    return texto

async def main():
    # Crear cliente
    if SESSION_STRING:
        client = TelegramClient(StringSession(SESSION_STRING), API_ID, API_HASH)
    else:
        client = TelegramClient("verificador_session", API_ID, API_HASH)
    
    await client.start(phone=PHONE)
    log.info("✅ Cliente Telethon conectado")
    
    # Comando para verificar: enviar mensaje al bot en privado
    @client.on(events.NewMessage(from_users=ADMIN_ID, pattern=r'/verificar(?:\s+(-?\d+))?(?:\s+(\d+))?'))
    async def cmd_verificar(event):
        args = event.pattern_match
        chat_id = int(args.group(1)) if args.group(1) else None
        limit   = int(args.group(2)) if args.group(2) else 500
        
        if not chat_id:
            # Listar grupos disponibles
            texto = "📋 *Grupos disponibles:*\n"
            async for dialog in client.iter_dialogs():
                if dialog.is_group:
                    texto += f"`{dialog.id}` — {dialog.name}\n"
            await event.reply(texto, parse_mode="Markdown")
            return
        
        await event.reply(f"🔍 Verificando últimos {limit} mensajes... Esto puede tardar varios minutos.")
        
        try:
            resultados = await verificar_historial(client, chat_id, limit)
            reporte = await generar_reporte(resultados, chat_id)
            await event.reply(reporte, parse_mode="Markdown")
            
            # Guardar resultados detallados en JSON
            with open(f"/data/verificacion_{chat_id}_{datetime.now().strftime('%Y%m%d_%H%M')}.json", "w") as f:
                json.dump(resultados, f, ensure_ascii=False, indent=2)
                
        except Exception as e:
            await event.reply(f"❌ Error: {e}")
    
    log.info("🤖 Verificador listo. Enviá /verificar al cliente para empezar.")
    await client.run_until_disconnected()

if __name__ == "__main__":
    from telethon.sessions import StringSession
    asyncio.run(main())
