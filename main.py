"""Bot verificador de comprobantes — cuenta de usuario (Telethon).

Comandos (desde tu propia cuenta o desde un ADMIN_ID):
  /grupos [texto]                        lista los chats accesibles con su ID
  /verificar CHAT_ID [desde] [hasta]     verifica todo el historial (fechas dd/mm/aaaa)
  /reset CHAT_ID                         limpia la verificación (conserva el historial)
  /reset CHAT_ID total                   además olvida ese chat del historial
  /verificar CHAT_ID ... forzar          vuelve a leer con Claude, ignora lo cacheado
  /estado                                estado del bot y del historial
  /cancelar [CHAT_ID]                    corta una verificación en curso
  /ayuda                                 esta ayuda

Cruce con Excel: mandá el .xlsx al privado. En el pie del archivo podés poner:
  (nada)                 cruza contra TODOS los grupos verificados, todas las hojas
  /excel 30-7            todos los grupos, solo la hoja «30-7»
  /excel -100123 30-7    solo ese grupo, solo esa hoja

Varias semanas en archivos separados:
  /sumar                 (pie del archivo) acumula ese Excel a la referencia
  /faltantes             consolidado: qué hay en los chats y en NINGÚN Excel
  /sumar limpiar         vacía la referencia acumulada
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
    r"(?i)^/(grupos|verificar|reset|estado|cancelar|ayuda|start|excel|importar|sumar|faltantes)(?:@\w+)?\s*(.*)$",
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
            elif cmd == "faltantes":
                await self.cmd_faltantes(event, resto)
            elif cmd == "sumar":
                if resto.strip().lower().startswith("limpiar"):
                    n = await self.storage.limpiar_referencia()
                    await event.respond(f"🧹 Referencia vaciada ({n} filas).")
                else:
                    ref = await self.storage.cargar_referencia()
                    if not ref["archivos"]:
                        await event.respond(
                            "Mandame cada Excel semanal con el pie `/sumar`.\n"
                            "Cuando estén todos, pedí `/faltantes`.")
                    else:
                        await event.respond(
                            f"📚 Referencia acumulada: **{len(ref['filas'])}** filas de "
                            f"{len(ref['archivos'])} archivo(s)\n"
                            + "\n".join(f"  · {a['nombre']} — {a['filas']} filas"
                                        for a in ref["archivos"])
                            + "\n\nPedí `/faltantes` para el consolidado, o "
                              "`/sumar limpiar` para empezar de nuevo.")
            elif cmd == "importar":
                await event.respond(
                    "Adjuntá una planilla `verificacion_*.xlsx` que te haya mandado el bot.\n"
                    "La reimporta al historial sin volver a gastar API.")
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
        total = any(p.lower() in ("total", "todo", "historial") for p in partes[1:])
        habia = await self.storage.resetear_verificacion(chat_id)
        borrados = await self.storage.olvidar_chat(chat_id) if total else 0
        st = self.storage.stats_historial()
        await event.respond(
            (f"🧹 Verificación de `{chat_id}` reiniciada.\n" if habia
             else f"No había verificación guardada para `{chat_id}`.\n") +
            (f"🗑 Borradas **{borrados}** entradas del historial de ese chat: "
             f"se van a releer con la API (tiene costo).\n" if total else
             f"📚 El historial se mantiene intacto: **{st['comprobantes']}** comprobantes.\n"
             f"Los reenvíos de comprobantes viejos se van a seguir detectando.\n"
             f"Para releer todo desde cero: `/verificar {chat_id} forzar`")
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
        forzar = bool(re.search(r"(?i)\b(forzar|releer|completo)\b", resto))
        resto = re.sub(r"(?i)\b(forzar|releer|completo)\b", " ", resto)
        partes = resto.split()
        if not partes:
            await event.respond(
                "Uso: `/verificar CHAT_ID [dd/mm/aaaa] [dd/mm/aaaa]`\n"
                "Agregá `forzar` al final para releer con Claude lo ya procesado.")
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

        tarea = asyncio.create_task(self.verificar(event, chat_id, desde, hasta, forzar))
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

    async def procesar(self, msg, pie: str, ctx: dict, estado: dict,
                       forzar: bool = False) -> dict:
        nombre_pie, monto_pie, doc_pie = utils.parse_pie(pie)
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
            "doc_pie": doc_pie,
            "link": link_mensaje(ctx["chat_id"], msg.id, ctx.get("username")),
        }

        # ¿Este mensaje exacto ya se leyó alguna vez? Se rehace el veredicto
        # con lo guardado: sin descarga y sin llamada a Claude.
        previo_msg = None if forzar else self.storage.buscar_mensaje(ctx["chat_id"], msg.id)
        if previo_msg and previo_msg.get("monto") is not None:
            datos = {
                "es_comprobante": True,
                "monto_num": previo_msg.get("monto"),
                "nombre_origen": previo_msg.get("nombre_origen") or previo_msg.get("nombre") or "",
                "nombre_destino": previo_msg.get("nombre_destino") or previo_msg.get("nombre") or "",
                "cuit_origen": previo_msg.get("cuit_origen") or "",
                "cuit_destino": previo_msg.get("cuit_destino") or "",
                "banco": previo_msg.get("banco") or "",
                "nro_operacion": previo_msg.get("nro_operacion") or "",
                "fecha": previo_msg.get("fecha_comp") or "",
                "hora": previo_msg.get("hora_comp") or "",
                "confianza": previo_msg.get("confianza") if previo_msg.get("confianza") is not None else 1.0,
            }
            item.update(
                nombre_img=datos["nombre_origen"] or datos["nombre_destino"],
                nombre_origen=datos["nombre_origen"], nombre_destino=datos["nombre_destino"],
                cuit_origen=datos.get("cuit_origen", ""), cuit_destino=datos.get("cuit_destino", ""),
                monto_img=datos["monto_num"],
                banco=datos["banco"], nro_operacion=datos["nro_operacion"],
                fecha_comp=datos["fecha"], hora_comp=datos.get("hora", ""),
                confianza=datos["confianza"],
                huella=utils.huella_comprobante(datos), reusado=True,
            )
            veredicto = an.comparar(datos, nombre_pie, monto_pie, doc_pie)
            item.update(estado=veredicto["estado"], detalle=veredicto["detalle"],
                        sim_nombre=veredicto["sim_nombre"])
            estado["reusados"] = estado.get("reusados", 0) + 1
            estado["procesados"] += 1
            return item

        data = await self.descargar(msg)
        if not data:
            item.update(estado="ilegible", detalle="no se pudo descargar la imagen")
            estado["procesados"] += 1
            return item

        hash_img = utils.hash_archivo(data)
        item["hash_archivo"] = hash_img

        # ¿Ya procesado antes? No hace falta gastar otra llamada a Claude.
        async with self.lock_dup:
            previo = self.storage.buscar_previo(None, hash_img)
        if previo and not (forzar and previo.get("chat_id") == ctx["chat_id"]
                           and previo.get("msg_id") == msg.id):
            mismo_mensaje = (previo.get("chat_id") == ctx["chat_id"]
                             and previo.get("msg_id") == msg.id)
            if mismo_mensaje and not forzar:
                # Es LA MISMA publicación releída (p. ej. una corrida cortada):
                # se rehace el veredicto con lo ya leído, sin pagar de nuevo.
                datos = {
                    "es_comprobante": True,
                    "monto_num": previo.get("monto"),
                    "nombre_destino": previo.get("nombre") or "",
                    "banco": previo.get("banco") or "",
                    "nro_operacion": previo.get("nro_operacion") or "",
                    "fecha": previo.get("fecha_comp") or "",
                "hora": previo.get("hora_comp") or "",
                    "confianza": previo.get("confianza") if previo.get("confianza") is not None else 1.0,
                }
                item.update(
                    nombre_img=datos["nombre_origen"] or datos["nombre_destino"],
                nombre_origen=datos["nombre_origen"], nombre_destino=datos["nombre_destino"],
                cuit_origen=datos.get("cuit_origen", ""), cuit_destino=datos.get("cuit_destino", ""),
                monto_img=datos["monto_num"],
                    banco=datos["banco"], nro_operacion=datos["nro_operacion"],
                    fecha_comp=datos["fecha"], hora_comp=datos.get("hora", ""),
                confianza=datos["confianza"],
                    huella=utils.huella_comprobante(datos), reusado=True,
                )
                veredicto = an.comparar(datos, nombre_pie, monto_pie, doc_pie)
                item.update(estado=veredicto["estado"], detalle=veredicto["detalle"],
                            sim_nombre=veredicto["sim_nombre"])
                estado["reusados"] = estado.get("reusados", 0) + 1
                estado["procesados"] += 1
                return item

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
            nombre_img=datos.get("nombre_origen") or datos.get("nombre_destino") or "",
            nombre_origen=datos.get("nombre_origen") or "",
            nombre_destino=datos.get("nombre_destino") or "",
            cuit_origen=datos.get("cuit_origen") or "",
            cuit_destino=datos.get("cuit_destino") or "",
            cvu_destino=datos.get("cvu_destino") or "",
            monto_img=datos.get("monto_num"),
            banco=datos.get("banco") or "",
            nro_operacion=datos.get("nro_operacion") or "",
            fecha_comp=datos.get("fecha") or "",
            hora_comp=datos.get("hora") or "",
            confianza=datos.get("confianza"),
        )
        huella = utils.huella_comprobante(datos) if datos.get("es_comprobante") else None
        item["huella"] = huella

        async with self.lock_dup:
            previo = None if ctx.get("forzar") else self.storage.buscar_previo(huella, hash_img)
            if previo:
                item.update(
                    estado="duplicado",
                    detalle=(f"comprobante ya registrado el {previo.get('fecha_msg', '?')} "
                             f"({utils.fmt_monto(previo.get('monto'))} — {previo.get('nombre') or '?'})"),
                    dup_detalle=previo.get("link", ""),
                )
                estado["procesados"] += 1
                return item
            await self.storage.registrar(item, actualizar=bool(ctx.get("forzar")))

        veredicto = an.comparar(datos, nombre_pie, monto_pie, doc_pie)
        item.update(estado=veredicto["estado"], detalle=veredicto["detalle"],
                    sim_nombre=veredicto["sim_nombre"])
        estado["procesados"] += 1
        return item

    async def verificar(self, event, chat_id: int, desde, hasta, forzar: bool = False):
        estado = {"mensajes": 0, "detectados": 0, "procesados": 0, "reusados": 0}
        self.progreso[chat_id] = estado
        aviso = None
        tareas: list[asyncio.Task] = []
        try:
            entidad = await self.resolver_entidad(chat_id)
            ctx = {
                "chat_id": chat_id,
                "chat": nombre_entidad(entidad),
                "username": getattr(entidad, "username", None),
                "forzar": forzar,
            }
            rango = (f"{desde.strftime('%d/%m/%Y')} → {hasta.strftime('%d/%m/%Y') if hasta else 'hoy'}"
                     if desde else "todo el historial")
            aviso = await event.respond(
                f"🔎 Verificando **{ctx['chat']}**\nRango: {rango}\n"
                + ("♻️ Relectura forzada: se vuelve a consultar Claude (tiene costo)\n" if forzar else "")
                + "Leyendo mensajes…")

            reporter = asyncio.create_task(self.reportar_progreso(aviso, estado, ctx["chat"]))
            try:
                async for msg, pie in self.recolectar(entidad, chat_id, desde, hasta, estado):
                    estado["detectados"] += 1
                    tareas.append(asyncio.create_task(
                        self.procesar(msg, pie, ctx, estado, forzar)))
                items = list(await asyncio.gather(*tareas))
            finally:
                reporter.cancel()

            items.sort(key=lambda x: x.get("msg_id", 0))
            resultado = {
                "chat_id": chat_id, "chat": ctx["chat"],
                "desde": desde.isoformat() if desde else None,
                "hasta": hasta.isoformat() if hasta else None,
                "mensajes": estado["mensajes"],
                "reusados": estado.get("reusados", 0),
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
            (f"♻️ Reusados del historial (sin costo de API): **{resultado.get('reusados', 0)}**"
             if resultado.get("reusados") else ""),
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
        tmp = Path("/tmp") / f"admin_{event.id}.xlsx"
        aviso = None
        try:
            await event.download_media(file=str(tmp))

            # ¿Es una planilla generada por este bot? Entonces se IMPORTA
            # (recupera la verificación sin repetir el gasto de API).
            forzar_cruce = bool(re.match(r"(?i)\s*/excel\b", pie))
            es_reporte = await asyncio.to_thread(xl.es_reporte_verificacion, tmp)
            if not forzar_cruce and (es_reporte or re.match(r"(?i)\s*/importar\b", pie)):
                await self.importar_planilla(event, tmp, nombre)
                return

            # Modo acumulativo: /sumar guarda las filas para el consolidado
            if re.match(r"(?i)\s*/sumar\b", pie):
                filas, desc = await asyncio.to_thread(xl.leer_excel, tmp, hojas)
                if not filas:
                    await event.respond(f"No pude leer filas.\n_{desc}_")
                    return
                ref = await self.storage.sumar_referencia(nombre, filas, desc)
                await event.respond(
                    f"➕ Sumado `{nombre}`: {len(filas)} filas\n"
                    f"📚 Referencia total: **{len(ref['filas'])}** filas de "
                    f"{len(ref['archivos'])} archivo(s)\n\n"
                    "Seguí mandando semanas, o pedí `/faltantes`.")
                return

            # A partir de acá: Excel de referencia para cruzar.
            if not ids:
                ids = [v["chat_id"] for v in await self.storage.ultimas_verificaciones(100)]
            verifs = []
            for cid in ids:
                v = await self.storage.cargar_verificacion(cid)
                if v:
                    verifs.append(v)
            if not verifs:
                await event.respond(
                    "No tengo verificaciones guardadas para cruzar.\n"
                    "Corré `/verificar CHAT_ID`, o mandame una planilla "
                    "`verificacion_*.xlsx` para importarla.")
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

    async def cmd_faltantes(self, event, resto: str):
        """Cruza TODOS los grupos contra TODOS los Excel acumulados."""
        ref = await self.storage.cargar_referencia()
        if not ref["filas"]:
            await event.respond(
                "No hay Excel acumulados todavía.\n"
                "Mandame cada semana con el pie `/sumar` y después pedí `/faltantes`.")
            return

        ids = [int(x) for x in re.findall(r"-?\d{6,}", resto)]
        semana = re.sub(r"-?\d{6,}", " ", resto).strip()
        if not ids:
            ids = [v["chat_id"] for v in await self.storage.ultimas_verificaciones(100)]
        items, grupos = [], []
        for cid in ids:
            v = await self.storage.cargar_verificacion(cid)
            if not v:
                continue
            etiqueta = v.get("chat") or str(v.get("chat_id"))
            grupos.append(etiqueta)
            for it in v.get("items", []):
                it.setdefault("chat", etiqueta)
                it.setdefault("chat_id", v.get("chat_id"))
                items.append(it)
        if not items:
            await event.respond("No hay verificaciones guardadas para cruzar.")
            return

        aviso = await event.respond(
            f"🔎 Cruzando **{len(items)}** comprobantes de {len(grupos)} grupo(s) "
            f"contra **{len(ref['filas'])}** filas de {len(ref['archivos'])} archivo(s)…")
        try:
            res = await asyncio.to_thread(xl.cruzar, items, ref["filas"])
            ruta = config.REPORTES_DIR / f"faltantes_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            await asyncio.to_thread(
                xl.generar_reporte, res,
                {"chat": ", ".join(grupos), "chat_id": ids[0] if ids else "",
                 "archivo": " + ".join(a["nombre"] for a in ref["archivos"]),
                 "grupos": grupos, "semana": semana}, ruta,
            )
            lineas = [
                "📊 **Consolidado de faltantes**",
                f"Archivos: {len(ref['archivos'])} · filas de referencia: {len(ref['filas'])}",
                f"Grupos: {len(grupos)} · comprobantes: {res['total_chat']}",
                "",
                f"✅ Coinciden: **{len(res['coinciden'])}**",
                f"🔁 Duplicados / reenviados: **{len(res['duplicados'])}**",
                f"⚠️ En el chat y NO en ningún Excel: **{len(res['solo_chat'])}**",
                f"❌ En el Excel y sin comprobante: **{len(res['solo_excel'])}**",
            ]
            if res["solo_chat"]:
                total = sum((xl._datos_item(s)[1] or 0) for s in res["solo_chat"])
                lineas.append(f"\n💰 Monto sin cargar: **{utils.fmt_monto(total)}**")
                lineas.append("\n**⚠️ Sin respaldo en ningún Excel:**")
                for it in res["solo_chat"][:30]:
                    n, mo = xl._datos_item(it)
                    lineas.append(f"• [{it.get('chat', '')}] {n or '?'} — {utils.fmt_monto(mo)} · "
                                  f"{it.get('fecha_msg', '')} {it.get('link', '')}")
                if len(res["solo_chat"]) > 30:
                    lineas.append(f"… y {len(res['solo_chat']) - 30} más en la planilla.")
            await aviso.delete()
            await self.responder_largo(event, "\n".join(lineas))
            await event.respond(
                file=str(ruta),
                message="📎 Consolidado. La hoja **📋 Para cargar** trae los faltantes "
                        "con las mismas columnas que tu planilla, listos para pegar.")
        except Exception as e:
            log.exception("Error en /faltantes")
            await event.respond(f"⚠️ No pude armar el consolidado: `{type(e).__name__}: {e}`")

    async def importar_planilla(self, event, tmp, nombre: str):
        """Reconstruye una verificación desde una planilla del propio bot."""
        aviso = await event.respond(f"♻️ Importando `{nombre}`…")
        try:
            meta, items = await asyncio.to_thread(xl.importar_verificacion, tmp)
            if not items or not meta.get("chat_id"):
                await aviso.edit("No pude leer la planilla: le falta la hoja «Comprobantes» "
                                 "o el Chat ID en el «Resumen».")
                return

            nuevos = 0
            for it in items:
                if it.get("msg_id") and it.get("monto_img") is not None:
                    antes = len(self.storage.historial.get("por_mensaje", {}))
                    await self.storage.registrar(it)
                    nuevos += len(self.storage.historial.get("por_mensaje", {})) - antes

            chat_id = meta["chat_id"]
            previa = await self.storage.cargar_verificacion(chat_id)
            if previa:
                por_msg = {i.get("msg_id"): i for i in previa.get("items", []) if i.get("msg_id")}
                for it in items:
                    por_msg.setdefault(it.get("msg_id"), it)
                items = sorted(por_msg.values(), key=lambda x: x.get("msg_id") or 0)

            await self.storage.guardar_verificacion(chat_id, {
                "chat_id": chat_id,
                "chat": meta.get("chat") or str(chat_id),
                "desde": meta.get("desde"),
                "hasta": meta.get("hasta"),
                "mensajes": meta.get("mensajes") or 0,
                "fin": datetime.now(config.TZ).strftime("%d/%m/%Y %H:%M"),
                "items": items,
                "importado_de": nombre,
            })
            self.ultima_verif = chat_id

            conteo: dict[str, int] = {}
            for it in items:
                conteo[it.get("estado", "?")] = conteo.get(it.get("estado", "?"), 0) + 1
            st = self.storage.stats_historial()

            await aviso.edit(
                f"♻️ **Importado**: {len(items)} comprobantes de "
                f"**{meta.get('chat') or chat_id}** (`{chat_id}`)\n"
                f"Nuevos en el historial: {nuevos}\n"
                f"Estados: " + " · ".join(f"{k}: {v}" for k, v in sorted(conteo.items())) + "\n\n"
                f"📚 Historial: {st['comprobantes']} comprobantes · {st.get('mensajes', 0)} mensajes\n"
                "Ya podés mandarme el Excel de referencia para cruzarlo."
            )
        except Exception as e:
            log.exception("Error importando planilla")
            await event.respond(f"⚠️ No pude importar la planilla: `{type(e).__name__}: {e}`")

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
