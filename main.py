"""Bot verificador de comprobantes — cuenta de usuario (Telethon).

Comandos (desde tu propia cuenta o desde un ADMIN_ID):
  /grupos [texto]                        lista los chats accesibles con su ID
  /verificar CHAT_ID [desde] [hasta]     verifica todo el historial (fechas dd/mm/aaaa)
  /reset CHAT_ID                         limpia la verificación (conserva el historial)
  /estado                                estado del bot y del historial
  /cancelar [CHAT_ID]                    corta una verificación en curso
  /ayuda                                 esta ayuda

Cruce con Excel: mandá el .xlsx al privado. En el pie del archivo podés poner:
  (nada)                 cruza contra TODOS los grupos verificados, todas las hojas
  /excel 30-7            todos los grupos, solo la hoja «30-7»
  /excel -100123 30-7    solo ese grupo, solo esa hoja
"""
from __future__ import annotations

import asyncio
import logging
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

from telethon import TelegramClient, events
from telethon.errors import FloodWaitError, RPCError
from telethon.sessions import StringSession
from telethon.tl.types import Channel, User

import analizador as an
import config
import excel as xl
import utils
from storage import Storage

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(name)s | %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logging.getLogger("telethon").setLevel(logging.WARNING)
logging.getLogger("httpx").setLevel(logging.WARNING)
log = logging.getLogger("bot")

RE_CMD = re.compile(
    r"(?i)^/(grupos|verificar|reset|estado|cancelar|ayuda|start|excel)(?:@\w+)?\s*(.*)$",
    re.DOTALL,
)


def es_imagen(msg) -> bool:
    if getattr(msg, "photo", None):
        return True
    doc = getattr(msg, "document", None)
    if doc and getattr(doc, "mime_type", "") and doc.mime_type.startswith("image/"):
        return True
    return False


def link_mensaje(chat_id: int, msg_id: int, username: str | None) -> str:
    if username:
        return f"https://t.me/{username}/{msg_id}"
    interno = str(chat_id)
    if interno.startswith("-100"):
        return f"https://t.me/c/{interno[4:]}/{msg_id}"
    return f"(msg {msg_id})"


def nombre_entidad(e) -> str:
    if isinstance(e, User):
        return " ".join(x for x in [e.first_name, e.last_name] if x) or (e.username or str(e.id))
    return getattr(e, "title", None) or str(getattr(e, "id", "?"))


class Bot:
    def __init__(self):
        self.client = TelegramClient(
            StringSession(config.SESSION_STRING), config.API_ID, config.API_HASH,
            connection_retries=None, retry_delay=5, request_retries=5,
        )
        self.storage = Storage(config.HISTORIAL_FILE, config.VERIF_DIR)
        self.analizador = an.Analizador()
        self.sem_descarga = asyncio.Semaphore(config.CONCURRENCIA)
        self.lock_dup = asyncio.Lock()
        self.tareas: dict[int, asyncio.Task] = {}
        self.progreso: dict[int, dict] = {}
        self.ultima_verif: int | None = None
        self.yo_id: int | None = None
        self.inicio = datetime.now(timezone.utc)

    # -------------------------------------------------------------- permisos
    def es_admin(self, event) -> bool:
        if event.out:
            return True
        return event.sender_id in config.ADMIN_IDS or event.sender_id == self.yo_id

    # ----------------------------------------------------------------- arranque
    async def start(self):
        config.validar()
        config.preparar_directorios()
        await self.storage.cargar()
        await self.client.connect()
        if not await self.client.is_user_authorized():
            raise SystemExit(
                "La sesión de Telegram no es válida o expiró.\n"
                "Regenerala localmente con `python generar_session.py` y actualizá "
                "la variable TELEGRAM_SESSION en Railway."
            )
        yo = await self.client.get_me()
        self.yo_id = yo.id
        log.info("Conectado como %s (id %s)", nombre_entidad(yo), yo.id)

        self.client.add_event_handler(self.on_comando, events.NewMessage(pattern=RE_CMD))
        self.client.add_event_handler(self.on_documento, events.NewMessage())

        vs = await self.storage.ultimas_verificaciones(1)
        if vs:
            self.ultima_verif = vs[0]["chat_id"]
        log.info("Bot listo. Historial: %s", self.storage.stats_historial())
        await self.client.run_until_disconnected()

    # ---------------------------------------------------------------- comandos
    async def on_comando(self, event):
        if not self.es_admin(event):
            return
        m = RE_CMD.match(event.raw_text or "")
        if not m:
            return
        cmd, resto = m.group(1).lower(), (m.group(2) or "").strip()
        try:
            if cmd in ("ayuda", "start"):
                await event.respond(__doc__)
            elif cmd == "grupos":
                await self.cmd_grupos(event, resto)
            elif cmd == "verificar":
                await self.cmd_verificar(event, resto)
            elif cmd == "reset":
                await self.cmd_reset(event, resto)
            elif cmd == "estado":
                await self.cmd_estado(event)
            elif cmd == "cancelar":
                await self.cmd_cancelar(event, resto)
            elif cmd == "excel":
                await event.respond(
                    "Adjuntá el .xlsx a este mensaje.\n"
                    "Sin pie: cruza contra todos los grupos verificados.\n"
                    "`/excel 30-7` limita a esa hoja · `/excel -100123 30-7` a un grupo y hoja.")
        except FloodWaitError as e:
            log.warning("FloodWait %ss", e.seconds)
            await asyncio.sleep(min(e.seconds, 60))
        except Exception as e:
            log.exception("Error en /%s", cmd)
            await event.respond(f"⚠️ Error en /{cmd}: `{type(e).__name__}: {e}`")

    async def cmd_grupos(self, event, filtro: str):
        filtro_n = utils.normalizar(filtro)
        lineas, n = [], 0
        async for d in self.client.iter_dialogs():
            if not (d.is_group or d.is_channel):
                continue
            titulo = d.name or "(sin título)"
            if filtro_n and filtro_n not in utils.normalizar(titulo):
                continue
            tipo = "canal" if isinstance(d.entity, Channel) and d.entity.broadcast else "grupo"
            n += 1
            lineas.append(f"`{d.id}` — {titulo} ({tipo})")
        if not lineas:
            await event.respond("No encontré grupos con ese filtro.")
            return
        cabecera = f"**Grupos accesibles ({n})**\nUsá: `/verificar CHAT_ID`\n\n"
        await self.responder_largo(event, cabecera + "\n".join(lineas))

    async def cmd_reset(self, event, resto: str):
        partes = resto.split()
        if not partes:
            await event.respond("Uso: `/reset CHAT_ID`")
            return
        try:
            chat_id = int(partes[0])
        except ValueError:
            await event.respond("El CHAT_ID tiene que ser numérico (mirá `/grupos`).")
            return
        habia = await self.storage.resetear_verificacion(chat_id)
        st = self.storage.stats_historial()
        await event.respond(
            (f"🧹 Verificación de `{chat_id}` reiniciada.\n" if habia
             else f"No había verificación guardada para `{chat_id}`.\n") +
            f"📚 El historial se mantiene intacto: **{st['comprobantes']}** comprobantes registrados.\n"
            f"Los reenvíos de comprobantes viejos se van a seguir detectando."
        )

    async def cmd_estado(self, event):
        st = self.storage.stats_historial()
        up = datetime.now(timezone.utc) - self.inicio
        horas, resto = divmod(int(up.total_seconds()), 3600)
        activas = [
            f"  · `{cid}`: {p['procesados']}/{p['detectados']} comprobantes"
            for cid, p in self.progreso.items()
            if cid in self.tareas and not self.tareas[cid].done()
        ]
        ultimas = await self.storage.ultimas_verificaciones(5)
        txt = [
            "**📊 Estado del bot**",
            f"Cuenta: `{self.yo_id}` · modelo: `{self.analizador.modelo}`",
            f"Uptime: {horas}h {resto // 60}m · llamadas a Claude: {self.analizador.llamadas}",
            f"Volume: `{config.DATA_DIR}`",
            "",
            f"📚 Historial: **{st['comprobantes']}** comprobantes · {st['imagenes']} imágenes",
            f"Última actualización: {st['actualizado'] or '—'}",
            "",
            "▶️ Verificaciones en curso: " + ("\n" + "\n".join(activas) if activas else "ninguna"),
        ]
        if ultimas:
            txt.append("\n🗂 Últimas verificaciones guardadas:")
            for u in ultimas:
                txt.append(f"  · `{u['chat_id']}` {u['chat'] or ''} — {u['total']} comprobantes ({u['fin'] or '—'})")
        await event.respond("\n".join(txt))

    async def cmd_cancelar(self, event, resto: str):
        ids = [int(x) for x in resto.split() if x.lstrip("-").isdigit()] or list(self.tareas)
        cancelados = 0
        for cid in ids:
            t = self.tareas.get(cid)
            if t and not t.done():
                t.cancel()
                cancelados += 1
        await event.respond(f"⏹ Cancelé {cancelados} verificación(es)." if cancelados
                            else "No hay verificaciones en curso.")

    async def cmd_verificar(self, event, resto: str):
        partes = resto.split()
        if not partes:
            await event.respond("Uso: `/verificar CHAT_ID [dd/mm/aaaa] [dd/mm/aaaa]`")
            return
        try:
            chat_id = int(partes[0])
        except ValueError:
            await event.respond("CHAT_ID inválido. Mirá `/grupos`.")
            return
        if chat_id in self.tareas and not self.tareas[chat_id].done():
            await event.respond("Ya hay una verificación en curso para ese chat. `/cancelar` para cortarla.")
            return

        desde = utils.parse_fecha(partes[1], config.TZ) if len(partes) > 1 else None
        hasta = utils.parse_fecha(partes[2], config.TZ, fin_del_dia=True) if len(partes) > 2 else None
        if len(partes) > 1 and desde is None:
            await event.respond("Fecha 'desde' inválida. Formato: `dd/mm/aaaa`.")
            return
        if len(partes) > 2 and hasta is None:
            await event.respond("Fecha 'hasta' inválida. Formato: `dd/mm/aaaa`.")
            return

        tarea = asyncio.create_task(self.verificar(event, chat_id, desde, hasta))
        self.tareas[chat_id] = tarea

    # ----------------------------------------------------------- verificación
    async def resolver_entidad(self, chat_id: int):
        try:
            return await self.client.get_entity(chat_id)
        except (ValueError, RPCError):
            async for d in self.client.iter_dialogs():
                if d.id == chat_id:
                    return d.entity
            raise ValueError(f"No tengo acceso al chat {chat_id}. Verificá el ID con /grupos.")

    async def recolectar(self, entidad, chat_id, desde, hasta, estado):
        """Va emitiendo (mensaje, pie). El pie puede venir del caption o de un
        mensaje de texto del mismo remitente pegado en el tiempo."""
        pendiente = None
        textos: dict[int, tuple[str, datetime]] = {}
        caption_album: dict[int, str] = {}
        kwargs = {"reverse": True}
        if desde:
            kwargs["offset_date"] = desde

        async for msg in self.client.iter_messages(entidad, **kwargs):
            estado["mensajes"] += 1
            if hasta and msg.date > hasta:
                break
            if desde and msg.date < desde:
                continue

            if es_imagen(msg):
                pie = (msg.message or "").strip()
                gid = getattr(msg, "grouped_id", None)
                if not pie and gid and gid in caption_album:
                    pie = caption_album[gid]
                if pie and gid:
                    caption_album[gid] = pie
                if pendiente is not None:
                    yield pendiente, ""
                    pendiente = None
                if pie:
                    yield msg, pie
                else:
                    prev = textos.pop(msg.sender_id, None)
                    if prev and utils.cerca_en_tiempo(msg.date, prev[1]):
                        yield msg, prev[0]
                    else:
                        pendiente = msg
            elif (msg.message or "").strip():
                txt = msg.message.strip()
                if txt.startswith("/"):
                    continue
                if pendiente is not None and pendiente.sender_id == msg.sender_id \
                        and utils.cerca_en_tiempo(msg.date, pendiente.date):
                    yield pendiente, txt
                    pendiente = None
                else:
                    if pendiente is not None:
                        yield pendiente, ""
                        pendiente = None
                    textos[msg.sender_id] = (txt, msg.date)
        if pendiente is not None:
            yield pendiente, ""

    async def descargar(self, msg) -> bytes | None:
        for intento in range(3):
            try:
                async with self.sem_descarga:
                    return await msg.download_media(file=bytes)
            except FloodWaitError as e:
                espera = min(e.seconds, 900)
                log.warning("FloodWait al descargar (%ss). Esperando…", espera)
                await asyncio.sleep(espera + 1)
            except Exception as e:
                log.warning("No pude descargar msg %s: %s", msg.id, e)
                await asyncio.sleep(2 * (intento + 1))
        return None

    async def procesar(self, msg, pie: str, ctx: dict, estado: dict) -> dict:
        nombre_pie, monto_pie = utils.parse_pie(pie)
        remitente = ""
        try:
            s = await msg.get_sender()
            remitente = nombre_entidad(s) if s else ""
        except Exception:
            pass

        item = {
            "chat_id": ctx["chat_id"], "chat": ctx["chat"], "msg_id": msg.id,
            "fecha_msg": utils.a_local(msg.date, config.TZ).strftime("%d/%m/%Y %H:%M"),
            "remitente": remitente, "remitente_id": msg.sender_id,
            "pie": pie, "nombre_pie": nombre_pie, "monto_pie": monto_pie,
            "link": link_mensaje(ctx["chat_id"], msg.id, ctx.get("username")),
        }

        data = await self.descargar(msg)
        if not data:
            item.update(estado="ilegible", detalle="no se pudo descargar la imagen")
            estado["procesados"] += 1
            return item

        hash_img = utils.hash_archivo(data)
        item["hash_archivo"] = hash_img

        # ¿Imagen idéntica ya procesada antes? No hace falta gastar una llamada.
        async with self.lock_dup:
            previo = self.storage.buscar_previo(None, hash_img)
        if previo:
            item.update(
                estado="duplicado",
                detalle=(f"misma imagen ya registrada el {previo.get('fecha_msg', '?')} "
                         f"en {previo.get('chat') or previo.get('chat_id')}"),
                dup_detalle=previo.get("link", ""),
                nombre_img=previo.get("nombre"), monto_img=previo.get("monto"),
                huella=None,
            )
            estado["procesados"] += 1
            return item

        datos = await self.analizador.leer_comprobante(data, "image/jpeg")
        item.update(
            nombre_img=datos.get("nombre_destino") or datos.get("nombre_origen") or "",
            monto_img=datos.get("monto_num"),
            banco=datos.get("banco") or "",
            nro_operacion=datos.get("nro_operacion") or "",
            fecha_comp=datos.get("fecha") or "",
            confianza=datos.get("confianza"),
        )
        huella = utils.huella_comprobante(datos) if datos.get("es_comprobante") else None
        item["huella"] = huella

        async with self.lock_dup:
            previo = self.storage.buscar_previo(huella, hash_img)
            if previo:
                item.update(
                    estado="duplicado",
                    detalle=(f"comprobante ya registrado el {previo.get('fecha_msg', '?')} "
                             f"({utils.fmt_monto(previo.get('monto'))} — {previo.get('nombre') or '?'})"),
                    dup_detalle=previo.get("link", ""),
                )
                estado["procesados"] += 1
                return item
            await self.storage.registrar(item)

        veredicto = an.comparar(datos, nombre_pie, monto_pie)
        item.update(estado=veredicto["estado"], detalle=veredicto["detalle"],
                    sim_nombre=veredicto["sim_nombre"])
        estado["procesados"] += 1
        return item

    async def verificar(self, event, chat_id: int, desde, hasta):
        estado = {"mensajes": 0, "detectados": 0, "procesados": 0}
        self.progreso[chat_id] = estado
        aviso = None
        tareas: list[asyncio.Task] = []
        try:
            entidad = await self.resolver_entidad(chat_id)
            ctx = {
                "chat_id": chat_id,
                "chat": nombre_entidad(entidad),
                "username": getattr(entidad, "username", None),
            }
            rango = (f"{desde.strftime('%d/%m/%Y')} → {hasta.strftime('%d/%m/%Y') if hasta else 'hoy'}"
                     if desde else "todo el historial")
            aviso = await event.respond(f"🔎 Verificando **{ctx['chat']}**\nRango: {rango}\nLeyendo mensajes…")

            reporter = asyncio.create_task(self.reportar_progreso(aviso, estado, ctx["chat"]))
            try:
                async for msg, pie in self.recolectar(entidad, chat_id, desde, hasta, estado):
                    estado["detectados"] += 1
                    tareas.append(asyncio.create_task(self.procesar(msg, pie, ctx, estado)))
                items = list(await asyncio.gather(*tareas))
            finally:
                reporter.cancel()

            items.sort(key=lambda x: x.get("msg_id", 0))
            resultado = {
                "chat_id": chat_id, "chat": ctx["chat"],
                "desde": desde.isoformat() if desde else None,
                "hasta": hasta.isoformat() if hasta else None,
                "mensajes": estado["mensajes"],
                "fin": datetime.now(config.TZ).strftime("%d/%m/%Y %H:%M"),
                "items": items,
            }
            await self.storage.guardar_verificacion(chat_id, resultado)
            self.ultima_verif = chat_id
            await self.enviar_resumen(event, aviso, resultado, rango)

        except asyncio.CancelledError:
            for t in tareas:
                t.cancel()
            if aviso:
                try:
                    await aviso.edit("⏹ Verificación cancelada.")
                except Exception:
                    pass
            raise
        except Exception as e:
            log.exception("Verificación fallida")
            await event.respond(f"⚠️ No pude terminar la verificación: `{type(e).__name__}: {e}`")
        finally:
            self.progreso.pop(chat_id, None)

    async def reportar_progreso(self, aviso, estado, titulo):
        anterior = None
        while True:
            await asyncio.sleep(config.INTERVALO_PROGRESO)
            actual = (estado["mensajes"], estado["procesados"])
            if actual == anterior:
                continue
            anterior = actual
            try:
                await aviso.edit(
                    f"🔎 Verificando **{titulo}**\n"
                    f"Mensajes leídos: {estado['mensajes']:,}\n"
                    f"Comprobantes: {estado['procesados']}/{estado['detectados']}"
                )
            except Exception:
                pass

    async def enviar_resumen(self, event, aviso, resultado, rango):
        items = resultado["items"]
        c = {}
        for it in items:
            c[it["estado"]] = c.get(it["estado"], 0) + 1
        total_ok = sum((it.get("monto_img") or 0) for it in items if it["estado"] == "ok")

        txt = [
            f"📋 **{resultado['chat']}** — verificación terminada",
            f"Rango: {rango} · mensajes leídos: {resultado['mensajes']:,}",
            f"Comprobantes detectados: **{len(items)}**",
            "",
            f"✅ Correctos: **{c.get('ok', 0)}**",
            f"🔁 Duplicados / reenviados: **{c.get('duplicado', 0)}**",
            f"❌ Monto no coincide: **{c.get('monto', 0)}**",
            f"❌ Nombre no coincide: **{c.get('nombre', 0)}**",
            f"❌ Nombre y monto: **{c.get('ambos', 0)}**",
            f"❓ Ilegibles: **{c.get('ilegible', 0)}**",
            f"⚠️ Sin pie de texto: **{c.get('sin_pie', 0)}**",
            f"🚫 No son comprobantes: **{c.get('no_comprobante', 0)}**",
            "",
            f"💰 Total válido: **{utils.fmt_monto(total_ok)}**",
        ]
        problemas = [it for it in items if it["estado"] not in ("ok",)]
        if problemas:
            txt.append(f"\n**Observaciones ({len(problemas)}):**")
            for it in problemas[:40]:
                icono = an.ESTADOS.get(it["estado"], "•")
                quien = it.get("remitente") or "?"
                txt.append(f"{icono} {it['fecha_msg']} · {quien} — {it.get('detalle', '')}\n{it.get('link', '')}")
            if len(problemas) > 40:
                txt.append(f"… y {len(problemas) - 40} más (mirá la planilla adjunta).")

        try:
            if aviso:
                await aviso.delete()
        except Exception:
            pass
        await self.responder_largo(event, "\n".join(txt))

        ruta = config.REPORTES_DIR / f"verificacion_{resultado['chat_id']}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        try:
            await asyncio.to_thread(
                xl.generar_reporte_verificacion, items,
                {"chat": resultado["chat"], "chat_id": resultado["chat_id"],
                 "desde": resultado["desde"], "hasta": resultado["hasta"],
                 "mensajes": resultado["mensajes"]},
                ruta,
            )
            await event.respond(file=str(ruta),
                                message="📎 Detalle completo. Mandame el Excel de referencia para cruzarlo.")
        except Exception as e:
            log.warning("No pude generar/enviar el reporte: %s", e)

    # ------------------------------------------------------------------ Excel
    async def on_documento(self, event):
        if not (event.is_private and event.document and self.es_admin(event)):
            return
        nombre = ""
        for attr in getattr(event.document, "attributes", []):
            nombre = getattr(attr, "file_name", "") or nombre
        if not nombre.lower().endswith((".xlsx", ".xlsm")):
            return

        # Sin IDs en el pie: cruza contra TODOS los grupos verificados.
        # Con IDs: solo contra esos. Cualquier otra palabra = nombre de hoja.
        pie = (event.raw_text or "").strip()
        ids = [int(x) for x in re.findall(r"-?\d{6,}", pie)]
        resto = re.sub(r"-?\d{6,}", " ", pie)
        resto = re.sub(r"(?i)^\s*/excel\b", " ", resto)
        hojas = [t for t in re.split(r"[\s,]+", resto) if t and not t.startswith("/")]
        if not ids:
            ids = [v["chat_id"] for v in await self.storage.ultimas_verificaciones(100)]
        if not ids:
            await event.respond("No tengo ninguna verificación guardada todavía. "
                                "Corré `/verificar CHAT_ID` primero.")
            return

        verifs = []
        for cid in ids:
            v = await self.storage.cargar_verificacion(cid)
            if v:
                verifs.append(v)
        if not verifs:
            await event.respond("No encontré verificaciones guardadas para esos IDs.")
            return

        items, grupos = [], []
        for v in verifs:
            etiqueta = v.get("chat") or str(v.get("chat_id"))
            grupos.append(etiqueta)
            for it in v.get("items", []):
                it.setdefault("chat", etiqueta)
                it.setdefault("chat_id", v.get("chat_id"))
                items.append(it)

        chat_id = verifs[0].get("chat_id")
        aviso = await event.respond(
            f"📥 Leyendo `{nombre}` y cruzando contra **{len(verifs)}** grupo(s): "
            + ", ".join(grupos) + "…")
        tmp = Path("/tmp") / f"admin_{event.id}.xlsx"
        try:
            await event.download_media(file=str(tmp))
            filas, desc = await asyncio.to_thread(xl.leer_excel, tmp, hojas)
            if not filas:
                await aviso.edit(f"No pude leer filas del Excel.\n_{desc}_\n\n"
                                 "Necesito una columna de nombre (Nombre/Titular/Cliente) "
                                 "y una de monto (Monto/Importe/Pago).")
                return
            res = await asyncio.to_thread(xl.cruzar, items, filas)
            ruta = config.REPORTES_DIR / f"cruce_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            await asyncio.to_thread(
                xl.generar_reporte, res,
                {"chat": ", ".join(grupos), "chat_id": chat_id,
                 "archivo": nombre, "grupos": grupos}, ruta,
            )

            lineas = [
                f"📊 **Cruce con `{nombre}`**",
                f"_{desc}_",
                f"Grupos: **{len(grupos)}** — " + ", ".join(grupos),
                "",
                f"✅ Coinciden: **{len(res['coinciden'])}**",
                f"🔁 Duplicados / reenviados: **{len(res['duplicados'])}**",
                f"⚠️ En el chat pero NO en el Excel: **{len(res['solo_chat'])}**",
                f"❌ En el Excel pero NO en el chat: **{len(res['solo_excel'])}**",
                "",
                f"Filas del Excel: {res['total_excel']} · comprobantes del chat: {res['total_chat']}",
            ]
            if res["solo_excel"]:
                lineas.append("\n**❌ Faltan en el chat:**")
                for f in res["solo_excel"][:25]:
                    lineas.append(f"• {f['nombre']} — {utils.fmt_monto(f['monto'])} (fila {f['fila']})")
                if len(res["solo_excel"]) > 25:
                    lineas.append(f"… y {len(res['solo_excel']) - 25} más.")
            if res["solo_chat"]:
                lineas.append("\n**⚠️ En el chat sin respaldo en el Excel:**")
                for it in res["solo_chat"][:25]:
                    n, mo = xl._datos_item(it)
                    lineas.append(f"• [{it.get('chat', '')}] {n or '?'} — {utils.fmt_monto(mo)} · "
                                  f"{it.get('fecha_msg', '')} {it.get('link', '')}")
                if len(res["solo_chat"]) > 25:
                    lineas.append(f"… y {len(res['solo_chat']) - 25} más.")

            await aviso.delete()
            await self.responder_largo(event, "\n".join(lineas))
            await event.respond(file=str(ruta), message="📎 Reporte completo del cruce.")
        except Exception as e:
            log.exception("Error procesando Excel")
            await event.respond(f"⚠️ No pude procesar el Excel: `{type(e).__name__}: {e}`")
        finally:
            tmp.unlink(missing_ok=True)

    # ----------------------------------------------------------------- helpers
    async def responder_largo(self, event, texto: str, limite: int = 3800):
        bloque = ""
        for linea in texto.split("\n"):
            if len(bloque) + len(linea) + 1 > limite:
                await event.respond(bloque)
                await asyncio.sleep(0.6)
                bloque = ""
            bloque += linea + "\n"
        if bloque.strip():
            await event.respond(bloque)


async def main():
    while True:
        bot = Bot()
        try:
            await bot.start()
        except FloodWaitError as e:
            espera = min(e.seconds + 5, 3600)
            log.error("FloodWait de Telegram: espero %ss antes de reconectar", espera)
            await asyncio.sleep(espera)
        except (KeyboardInterrupt, SystemExit):
            raise
        except Exception:
            log.exception("Caída inesperada. Reintento en 30s")
            await asyncio.sleep(30)


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        log.info("Chau.")
