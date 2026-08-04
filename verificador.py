import os
import json
import asyncio
import base64
import logging
from datetime import datetime, timezone, timedelta

import anthropic
from telethon import TelegramClient, events
from telethon.sessions import StringSession

logging.basicConfig(format="%(asctime)s | %(levelname)s | %(message)s", level=logging.INFO)
log = logging.getLogger(__name__)

# ── Config ────────────────────────────────────────────────────────────────────
API_ID        = int(os.environ["TELEGRAM_API_ID"])
API_HASH      = os.environ["TELEGRAM_API_HASH"]
ADMIN_ID      = 531707598
ANTHROPIC_KEY = os.environ["ANTHROPIC_API_KEY"]
SESSION_STRING = os.environ["TELETHON_SESSION"]
ARG_TZ        = timezone(timedelta(hours=-3))

claude = anthropic.Anthropic(api_key=ANTHROPIC_KEY)

# ── Prompt de verificación ────────────────────────────────────────────────────
SYSTEM_PROMPT = """Eres un verificador de comprobantes bancarios argentinos.
Te doy una imagen de comprobante y el texto del pie que escribió el usuario.
Analizá el comprobante y respondé ÚNICAMENTE con JSON válido sin backticks:
{
  "nombre_imagen": "nombre extraído de la imagen o vacío",
  "nombre_pie": "nombre del pie de texto",
  "monto_imagen": número o null,
  "coincide_nombre": true o false,
  "coincide_monto": true o false,
  "observaciones": "descripción de diferencias o vacío",
  "confianza": "alta/media/baja"
}
Sé flexible con variaciones menores (mayúsculas, orden, abreviaturas).
"""

async def analizar_comprobante(image_bytes: bytes, pie: str) -> dict:
    b64 = base64.standard_b64encode(image_bytes).decode("utf-8")
    try:
        resp = await asyncio.get_event_loop().run_in_executor(
            None,
            lambda: claude.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=500,
                system=[{"type": "text", "text": SYSTEM_PROMPT, "cache_control": {"type": "ephemeral"}}],
                messages=[{"role": "user", "content": [
                    {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": b64}},
                    {"type": "text", "text": f"Pie del usuario: {pie}\n\nVerificá este comprobante."}
                ]}]
            )
        )
        text = resp.content[0].text
        return json.loads(text.replace("```json", "").replace("```", "").strip())
    except Exception as e:
        log.error(f"Error analizando: {e}")
        return {"error": str(e)}

async def verificar_historial(client, chat_id: int, limit: int = 500):
    log.info(f"Verificando historial de {chat_id} — últimos {limit} mensajes")
    resultados = []
    mensajes = []
    async for msg in client.iter_messages(chat_id, limit=limit):
        mensajes.append(msg)
    mensajes.reverse()

    for msg in mensajes:
        if not msg.photo:
            continue
        caption = (msg.message or "").strip()
        if not caption:
            continue
        try:
            image_bytes = await client.download_media(msg, bytes)
            if not image_bytes:
                continue
            resultado = await analizar_comprobante(image_bytes, caption)
            resultado["msg_id"] = msg.id
            resultado["fecha"] = msg.date.strftime("%d/%m/%Y %H:%M")
            resultado["pie"] = caption[:60]
            resultados.append(resultado)
            log.info(f"msg #{msg.id} — coincide_nombre:{resultado.get('coincide_nombre')} coincide_monto:{resultado.get('coincide_monto')}")
            await asyncio.sleep(1)
        except Exception as e:
            log.error(f"Error procesando msg #{msg.id}: {e}")

    return resultados

async def generar_reporte(resultados: list) -> str:
    total = len(resultados)
    if total == 0:
        return "📭 No se encontraron comprobantes con pie de texto."
    
    errores = [r for r in resultados if not r.get("coincide_nombre") or not r.get("coincide_monto")]
    ok = total - len(errores)

    texto = f"📋 *Reporte de verificación*\n─────────────────────\n"
    texto += f"✅ Correctos: {ok}/{total}\n"
    texto += f"⚠️ Con discrepancias: {len(errores)}\n"

    if errores:
        texto += f"\n*Discrepancias:*\n"
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
    # Usar session string — no pide código
    client = TelegramClient(StringSession(SESSION_STRING), API_ID, API_HASH)
    await client.connect()
    
    if not await client.is_user_authorized():
        log.error("❌ Session string inválido o expirado")
        return

    me = await client.get_me()
    log.info(f"✅ Conectado como {me.first_name} (@{me.username})")

    @client.on(events.NewMessage(from_users=ADMIN_ID))
    async def handler(event):
        text = event.text or ""
        
        if text == "/grupos":
            resp = "📋 *Grupos disponibles:*\n"
            async for dialog in client.iter_dialogs():
                if dialog.is_group or dialog.is_channel:
                    resp += f"`{dialog.id}` — {dialog.name}\n"
            await event.reply(resp, parse_mode="Markdown")
            return

        if text.startswith("/verificar"):
            parts = text.split()
            if len(parts) < 2:
                await event.reply("Uso: `/verificar CHAT_ID [LIMITE]`\nEjemplo: `/verificar -1001234567890 500`")
                return
            
            try:
                chat_id = int(parts[1])
                limit = int(parts[2]) if len(parts) > 2 else 500
            except ValueError:
                await event.reply("❌ CHAT_ID inválido")
                return

            await event.reply(f"🔍 Verificando últimos {limit} mensajes... puede tardar varios minutos.")
            try:
                resultados = await verificar_historial(client, chat_id, limit)
                reporte = await generar_reporte(resultados)
                await event.reply(reporte, parse_mode="Markdown")
                # Guardar JSON
                fname = f"/tmp/verificacion_{abs(chat_id)}_{datetime.now().strftime('%Y%m%d_%H%M')}.json"
                with open(fname, "w") as f:
                    json.dump(resultados, f, ensure_ascii=False, indent=2)
                log.info(f"Resultados guardados en {fname}")
            except Exception as e:
                await event.reply(f"❌ Error: {e}")

    log.info("🤖 Verificador listo — enviame /grupos o /verificar CHAT_ID")
    await client.run_until_disconnected()

if __name__ == "__main__":
    asyncio.run(main())
