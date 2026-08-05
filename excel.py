"""Lectura del Excel del admin y cruce contra lo verificado en el chat."""
from __future__ import annotations

import re
import logging
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import config
import utils

log = logging.getLogger("excel")

CLAVES_NOMBRE = ("nombre", "apellido", "titular", "cliente", "beneficiario",
                 "destinatario", "socio", "alumno", "persona", "paciente")
CLAVES_MONTO = ("monto", "importe", "valor", "pago", "total", "abono", "deposito", "depósito")
CLAVES_FECHA = ("fecha", "dia", "día")
ENCABEZADOS_RUIDO = {"NOMBRE", "MONTO", "IMPORTE", "TITULAR", "CLIENTE", "FECHA", "TOTAL",
                     "APELLIDO", "NOMBRE Y APELLIDO", "BENEFICIARIO", "DESTINATARIO"}


def _detectar_columnas(hoja, max_filas: int = 15) -> tuple[int, dict]:
    """Devuelve (fila_encabezado, {'nombre': idx, 'monto': idx, 'fecha': idx})."""
    for fila in range(1, min(max_filas, hoja.max_row or 1) + 1):
        mapa = {}
        for col in range(1, min(hoja.max_column or 1, 40) + 1):
            val = hoja.cell(row=fila, column=col).value
            if not isinstance(val, str):
                continue
            t = utils.normalizar(val).lower()
            if not t:
                continue
            if "nombre" not in mapa and any(k in t for k in CLAVES_NOMBRE):
                mapa["nombre"] = col
            elif "monto" not in mapa and any(k in t for k in CLAVES_MONTO):
                mapa["monto"] = col
            elif "fecha" not in mapa and any(k in t for k in CLAVES_FECHA):
                mapa["fecha"] = col
        if "nombre" in mapa and "monto" in mapa:
            return fila, mapa
    # Sin encabezados reconocibles: asumimos A=nombre, B=monto
    return 0, {"nombre": 1, "monto": 2}


def leer_excel(path: str | Path, hojas: list[str] | None = None) -> tuple[list[dict], str]:
    """Lee las hojas del libro (todas, o solo las de `hojas`).

    Devuelve (filas, cómo se interpretó).
    """
    wb = load_workbook(filename=str(path), data_only=True, read_only=True)
    filas: list[dict] = []
    detalles: list[str] = []
    ignoradas: list[str] = []

    pedidas = [utils.normalizar(h) for h in (hojas or []) if h]
    seleccionadas = wb.worksheets
    if pedidas:
        seleccionadas = [h for h in wb.worksheets if utils.normalizar(h.title) in pedidas]
        if not seleccionadas:
            disponibles = ", ".join(h.title for h in wb.worksheets)
            wb.close()
            return [], f"no encontré esa(s) hoja(s). Disponibles: {disponibles}"

    for hoja in seleccionadas:
        fila_enc, cols = _detectar_columnas(hoja)
        if not fila_enc and hoja.max_row and hoja.max_row > 2:
            # Sin encabezado reconocible: solo se asume A/B si es la única hoja
            if len(seleccionadas) > 1:
                ignoradas.append(hoja.title)
                continue

        n_antes = len(filas)
        inicio = fila_enc + 1 if fila_enc else 1

        for i, fila in enumerate(hoja.iter_rows(min_row=inicio, values_only=True), start=inicio):
            def celda(clave):
                idx = cols.get(clave)
                if not idx or idx > len(fila):
                    return None
                return fila[idx - 1]

            nombre_raw = celda("nombre")
            monto_raw = celda("monto")
            if nombre_raw is None and monto_raw is None:
                continue
            nombre = str(nombre_raw).strip() if nombre_raw is not None else ""
            if utils.normalizar(nombre) in ENCABEZADOS_RUIDO:
                continue  # encabezado repetido en medio de la hoja
            monto = utils.parse_monto(monto_raw)
            if not utils.tokens_nombre(nombre) and monto is None:
                continue
            fecha = celda("fecha")
            if isinstance(fecha, datetime):
                fecha = fecha.date().isoformat()
            filas.append({
                "hoja": hoja.title,
                "fila": i,
                "nombre": nombre,
                "monto": monto,
                "fecha": str(fecha) if fecha else "",
            })

        if len(filas) > n_antes:
            detalles.append(f"«{hoja.title}» {len(filas) - n_antes} filas "
                            f"({get_column_letter(cols['nombre'])}/{get_column_letter(cols['monto'])})")
    wb.close()

    desc = f"{len(filas)} filas de {len(detalles)} hoja(s): " + "; ".join(detalles)
    if ignoradas:
        desc += " · sin encabezado reconocible: " + ", ".join(ignoradas)
    return filas, desc


# ------------------------------------------------------------------- cruce
def _datos_item(it: dict) -> tuple[str, float | None]:
    nombre = it.get("nombre_img") or it.get("nombre_pie") or ""
    monto = it.get("monto_img")
    if monto is None:
        monto = it.get("monto_pie")
    return nombre, monto


def cruzar(items_chat: list[dict], filas_excel: list[dict]) -> dict:
    """Clasifica en coinciden / duplicados / solo_chat / solo_excel."""
    candidatos = [
        it for it in items_chat
        if it.get("estado") not in ("no_comprobante", "sin_pie")
    ]

    # Duplicados: marcados durante la verificación, o repetidos dentro del lote
    # (incluye el mismo comprobante mandado a DOS grupos distintos).
    duplicados = [it for it in candidatos if it.get("estado") == "duplicado"]
    ids_dup = {id(it) for it in duplicados}
    vistos: dict[str, dict] = {}
    for it in candidatos:
        if it.get("estado") == "duplicado":
            continue
        claves = [k for k in (it.get("huella"), it.get("hash_archivo")) if k]
        if not claves:
            continue
        previo = next((vistos[k] for k in claves if k in vistos), None)
        if previo is not None:
            copia = dict(it)
            copia["dup_de"] = previo.get("msg_id")
            copia["dup_detalle"] = (f"ya contado en {previo.get('chat') or previo.get('chat_id')} "
                                    f"({previo.get('fecha_msg', '')})")
            duplicados.append(copia)
            ids_dup.add(id(it))       # el original queda fuera de los faltantes
        else:
            for k in claves:
                vistos[k] = it

    disponibles = [it for it in candidatos if id(it) not in ids_dup]
    usados: set[int] = set()

    coinciden, solo_excel = [], []
    for fila in filas_excel:
        mejor, mejor_sim = None, -1.0
        for j, it in enumerate(disponibles):
            if j in usados:
                continue
            nombre_it, monto_it = _datos_item(it)
            if fila["monto"] is not None and monto_it is not None:
                if not utils.montos_iguales(monto_it, fila["monto"], config.TOLERANCIA_MONTO):
                    continue
            sim = utils.similitud_nombres(nombre_it, fila["nombre"])
            if sim >= config.UMBRAL_NOMBRE and sim > mejor_sim:
                mejor, mejor_sim = j, sim
        if mejor is not None:
            usados.add(mejor)
            coinciden.append({"excel": fila, "chat": disponibles[mejor], "similitud": mejor_sim})
        else:
            solo_excel.append(fila)

    solo_chat = [it for j, it in enumerate(disponibles) if j not in usados]

    return {
        "coinciden": coinciden,
        "duplicados": duplicados,
        "solo_chat": solo_chat,
        "solo_excel": solo_excel,
        "total_excel": len(filas_excel),
        "total_chat": len(candidatos),
    }


# ------------------------------------------------------------------ reporte
_CAB = Font(bold=True, color="FFFFFF")
_FILL = {
    "ok": PatternFill("solid", fgColor="1E7B34"),
    "dup": PatternFill("solid", fgColor="B58900"),
    "chat": PatternFill("solid", fgColor="C25E00"),
    "excel": PatternFill("solid", fgColor="A32020"),
    "res": PatternFill("solid", fgColor="333333"),
}


def _hoja(wb, titulo, cabeceras, filas, color):
    ws = wb.create_sheet(titulo[:31])
    ws.append(cabeceras)
    for c in range(1, len(cabeceras) + 1):
        celda = ws.cell(row=1, column=c)
        celda.font = _CAB
        celda.fill = _FILL[color]
        celda.alignment = Alignment(horizontal="center")
    for f in filas:
        ws.append(f)
    for c in range(1, len(cabeceras) + 1):
        largo = max([len(str(cabeceras[c - 1]))] + [len(str(f[c - 1])) for f in filas[:200] if len(f) >= c] or [10])
        ws.column_dimensions[get_column_letter(c)].width = min(max(largo + 2, 12), 45)
    ws.freeze_panes = "A2"
    return ws


def generar_reporte(res: dict, meta: dict, destino: Path) -> Path:
    wb = Workbook()
    wb.remove(wb.active)

    resumen = wb.create_sheet("Resumen")
    filas_res = [
        ("Chat", meta.get("chat", "")),
        ("Chat ID", meta.get("chat_id", "")),
        ("Excel", meta.get("archivo", "")),
        ("Generado", datetime.now(config.TZ).strftime("%d/%m/%Y %H:%M")),
        ("", ""),
        ("✅ Coinciden", len(res["coinciden"])),
        ("🔁 Duplicados / reenviados", len(res["duplicados"])),
        ("⚠️ En el chat y NO en el Excel", len(res["solo_chat"])),
        ("❌ En el Excel y NO en el chat", len(res["solo_excel"])),
        ("", ""),
        ("Filas del Excel", res["total_excel"]),
        ("Comprobantes del chat", res["total_chat"]),
        ("Monto total coincidente", utils.fmt_monto(
            sum((c["excel"]["monto"] or 0) for c in res["coinciden"]))),
    ]
    for k, v in filas_res:
        resumen.append([k, v])

    if meta.get("grupos"):
        resumen.append(["", ""])
        resumen.append(["Grupos incluidos", ""])
        por_grupo: dict[str, int] = {}
        for s in res["solo_chat"]:
            g = s.get("chat") or str(s.get("chat_id", ""))
            por_grupo[g] = por_grupo.get(g, 0) + 1
        for g in meta["grupos"]:
            resumen.append([f"  {g}", f"{por_grupo.get(g, 0)} sin respaldo en el Excel"])

    for r in range(1, resumen.max_row + 1):
        resumen.cell(row=r, column=1).font = Font(bold=True)
    resumen.column_dimensions["A"].width = 34
    resumen.column_dimensions["B"].width = 40

    _hoja(wb, "✅ Coinciden",
          ["Nombre (Excel)", "Monto (Excel)", "Fila", "Grupo", "Nombre (imagen)",
           "Monto (imagen)", "Similitud", "Fecha msg", "Mensaje"],
          [[c["excel"]["nombre"], c["excel"]["monto"], c["excel"]["fila"],
            c["chat"].get("chat", ""), _datos_item(c["chat"])[0], _datos_item(c["chat"])[1],
            round(c["similitud"], 2), c["chat"].get("fecha_msg", ""), c["chat"].get("link", "")]
           for c in res["coinciden"]], "ok")

    _hoja(wb, "🔁 Duplicados",
          ["Grupo", "Nombre", "Monto", "Fecha msg", "Remitente", "Original", "Mensaje"],
          [[d.get("chat", ""), _datos_item(d)[0], _datos_item(d)[1], d.get("fecha_msg", ""),
            d.get("remitente", ""), d.get("dup_detalle") or d.get("dup_de", ""), d.get("link", "")]
           for d in res["duplicados"]], "dup")

    _hoja(wb, "⚠️ Solo en chat",
          ["Grupo", "Nombre", "Monto", "Estado", "Detalle", "Fecha msg", "Remitente", "Mensaje"],
          [[s.get("chat", ""), _datos_item(s)[0], _datos_item(s)[1], s.get("estado", ""),
            s.get("detalle", ""), s.get("fecha_msg", ""), s.get("remitente", ""), s.get("link", "")]
           for s in res["solo_chat"]], "chat")

    _hoja(wb, "❌ Solo en Excel",
          ["Nombre", "Monto", "Fecha", "Fila del Excel"],
          [[s["nombre"], s["monto"], s.get("fecha", ""), s["fila"]] for s in res["solo_excel"]],
          "excel")

    destino.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(destino))
    return destino


def generar_reporte_verificacion(items: list[dict], meta: dict, destino: Path) -> Path:
    """Planilla con TODOS los comprobantes leídos en el chat (sin Excel de por medio)."""
    wb = Workbook()
    wb.remove(wb.active)

    resumen = wb.create_sheet("Resumen")
    conteo: dict[str, int] = {}
    for it in items:
        conteo[it.get("estado", "?")] = conteo.get(it.get("estado", "?"), 0) + 1
    for k, v in [
        ("Chat", meta.get("chat", "")),
        ("Chat ID", meta.get("chat_id", "")),
        ("Desde", meta.get("desde", "todo el historial")),
        ("Hasta", meta.get("hasta", "hoy")),
        ("Mensajes leídos", meta.get("mensajes", "")),
        ("Comprobantes", len(items)),
        ("", ""),
        *[(k2, v2) for k2, v2 in conteo.items()],
    ]:
        resumen.append([k, v])
    resumen.column_dimensions["A"].width = 30
    resumen.column_dimensions["B"].width = 40

    _hoja(wb, "Comprobantes",
          ["Estado", "Nombre (imagen)", "Monto (imagen)", "Nombre (pie)", "Monto (pie)",
           "Detalle", "Banco", "N° operación", "Fecha comprobante", "Fecha msg",
           "Remitente", "Mensaje"],
          [[it.get("estado", ""), it.get("nombre_img", ""), it.get("monto_img"),
            it.get("nombre_pie", ""), it.get("monto_pie"), it.get("detalle", ""),
            it.get("banco", ""), it.get("nro_operacion", ""), it.get("fecha_comp", ""),
            it.get("fecha_msg", ""), it.get("remitente", ""), it.get("link", "")]
           for it in items], "res")

    destino.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(destino))
    return destino


# ------------------------------------------------- importar reporte propio
CABECERAS_VERIF = ("Estado", "Nombre (imagen)", "Monto (imagen)", "Nombre (pie)",
                   "Monto (pie)", "Detalle", "Fecha msg", "Mensaje")


def es_reporte_verificacion(path: str | Path) -> bool:
    """True si el .xlsx es una planilla generada por este bot (no un Excel de cruce)."""
    try:
        wb = load_workbook(filename=str(path), data_only=True, read_only=True)
    except Exception:
        return False
    try:
        if "Comprobantes" not in wb.sheetnames:
            return False
        ws = wb["Comprobantes"]
        fila = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
        titulos = {str(c).strip() for c in fila if c}
        return len(titulos & set(CABECERAS_VERIF)) >= 5
    finally:
        wb.close()


def importar_verificacion(path: str | Path) -> tuple[dict, list[dict]]:
    """Reconstruye (meta, items) desde una planilla de verificación del bot."""
    wb = load_workbook(filename=str(path), data_only=True, read_only=True)
    try:
        meta: dict = {}
        if "Resumen" in wb.sheetnames:
            for fila in wb["Resumen"].iter_rows(values_only=True):
                if fila and fila[0]:
                    meta[str(fila[0]).strip()] = fila[1] if len(fila) > 1 else None

        ws = wb["Comprobantes"]
        filas = ws.iter_rows(values_only=True)
        cab = [str(c).strip() if c is not None else "" for c in next(filas, ())]
        idx = {n: i for i, n in enumerate(cab)}

        def val(fila, nombre):
            i = idx.get(nombre)
            if i is None or i >= len(fila):
                return None
            return fila[i]

        try:
            chat_id = int(str(meta.get("Chat ID") or 0))
        except (TypeError, ValueError):
            chat_id = 0
        chat = str(meta.get("Chat") or "")

        items: list[dict] = []
        for fila in filas:
            if not fila or not any(fila):
                continue
            link = str(val(fila, "Mensaje") or "")
            m = re.search(r"/(\d+)\s*$", link)
            msg_id = int(m.group(1)) if m else None
            if not chat_id:
                m2 = re.search(r"/c/(\d+)/", link)
                if m2:
                    chat_id = int(f"-100{m2.group(1)}")

            datos = {
                "monto": val(fila, "Monto (imagen)"),
                "fecha": val(fila, "Fecha comprobante"),
                "nombre_destino": val(fila, "Nombre (imagen)"),
                "nro_operacion": val(fila, "N° operación"),
                "banco": val(fila, "Banco"),
            }
            items.append({
                "chat_id": chat_id,
                "chat": chat,
                "msg_id": msg_id,
                "fecha_msg": str(val(fila, "Fecha msg") or ""),
                "remitente": str(val(fila, "Remitente") or ""),
                "nombre_img": str(val(fila, "Nombre (imagen)") or ""),
                "monto_img": utils.parse_monto(val(fila, "Monto (imagen)")),
                "nombre_pie": str(val(fila, "Nombre (pie)") or ""),
                "monto_pie": utils.parse_monto(val(fila, "Monto (pie)")),
                "banco": str(val(fila, "Banco") or ""),
                "nro_operacion": str(val(fila, "N° operación") or ""),
                "fecha_comp": str(val(fila, "Fecha comprobante") or ""),
                "estado": str(val(fila, "Estado") or "").strip(),
                "detalle": str(val(fila, "Detalle") or ""),
                "link": link,
                "huella": utils.huella_comprobante(datos),
                "importado": True,
            })

        meta_out = {
            "chat_id": chat_id,
            "chat": chat,
            "mensajes": meta.get("Mensajes leídos") or 0,
            "desde": meta.get("Desde"),
            "hasta": meta.get("Hasta"),
        }
        return meta_out, items
    finally:
        wb.close()
