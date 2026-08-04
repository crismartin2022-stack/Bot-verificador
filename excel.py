"""Lectura del Excel del admin y cruce contra lo verificado en el chat."""
from __future__ import annotations

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


def leer_excel(path: str | Path) -> tuple[list[dict], str]:
    """Devuelve (filas, descripción de cómo se interpretó el archivo)."""
    wb = load_workbook(filename=str(path), data_only=True, read_only=True)
    hoja = wb.active
    fila_enc, cols = _detectar_columnas(hoja)
    filas: list[dict] = []
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
        monto = utils.parse_monto(monto_raw)
        if not utils.tokens_nombre(nombre) and monto is None:
            continue
        fecha = celda("fecha")
        if isinstance(fecha, datetime):
            fecha = fecha.date().isoformat()
        filas.append({
            "fila": i,
            "nombre": nombre,
            "monto": monto,
            "fecha": str(fecha) if fecha else "",
        })
    wb.close()

    desc = (f"hoja «{hoja.title}», columnas: "
            f"nombre={get_column_letter(cols['nombre'])}, monto={get_column_letter(cols['monto'])}"
            + (f", fecha={get_column_letter(cols['fecha'])}" if cols.get("fecha") else "")
            + (f" (encabezado en fila {fila_enc})" if fila_enc else " (sin encabezado detectado)"))
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

    # Duplicados: marcados durante la verificación o repetidos dentro del mismo lote
    duplicados = [it for it in candidatos if it.get("estado") == "duplicado"]
    vistos: dict[str, dict] = {}
    for it in candidatos:
        h = it.get("huella")
        if not h or it.get("estado") == "duplicado":
            continue
        if h in vistos:
            it = dict(it)
            it["dup_de"] = vistos[h].get("msg_id")
            duplicados.append(it)
        else:
            vistos[h] = it

    ids_dup = {id(x) for x in duplicados}
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
    for r in range(1, resumen.max_row + 1):
        resumen.cell(row=r, column=1).font = Font(bold=True)
    resumen.column_dimensions["A"].width = 34
    resumen.column_dimensions["B"].width = 40

    _hoja(wb, "✅ Coinciden",
          ["Nombre (Excel)", "Monto (Excel)", "Fila", "Nombre (imagen)", "Monto (imagen)",
           "Similitud", "Fecha msg", "Mensaje"],
          [[c["excel"]["nombre"], c["excel"]["monto"], c["excel"]["fila"],
            _datos_item(c["chat"])[0], _datos_item(c["chat"])[1],
            round(c["similitud"], 2), c["chat"].get("fecha_msg", ""), c["chat"].get("link", "")]
           for c in res["coinciden"]], "ok")

    _hoja(wb, "🔁 Duplicados",
          ["Nombre", "Monto", "Fecha msg", "Remitente", "Original", "Mensaje"],
          [[_datos_item(d)[0], _datos_item(d)[1], d.get("fecha_msg", ""), d.get("remitente", ""),
            d.get("dup_detalle") or d.get("dup_de", ""), d.get("link", "")]
           for d in res["duplicados"]], "dup")

    _hoja(wb, "⚠️ Solo en chat",
          ["Nombre", "Monto", "Estado", "Detalle", "Fecha msg", "Remitente", "Mensaje"],
          [[_datos_item(s)[0], _datos_item(s)[1], s.get("estado", ""), s.get("detalle", ""),
            s.get("fecha_msg", ""), s.get("remitente", ""), s.get("link", "")]
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
