"""Normalización de nombres/montos/fechas y huellas de comprobantes."""
from __future__ import annotations

import hashlib
import re
import unicodedata
from datetime import datetime
from difflib import SequenceMatcher

# Palabras que no aportan a la identidad de una persona
STOPWORDS = {
    "de", "del", "la", "las", "los", "el", "y", "da", "do", "dos", "sr", "sra",
    "srta", "don", "dona", "titular", "destinatario", "para", "a", "cuenta",
    "alias", "cbu", "cvu", "cuit", "cuil", "dni", "transferencia", "pago",
    "comprobante", "envio", "enviado", "recibido", "sa", "srl",
    "cuota", "cuotas", "mes", "mensual", "semana", "efectivo", "deposito",
    "abono", "saldo", "resto", "total", "gracias", "ok", "listo", "hola",
}

_RE_MONTO = re.compile(
    r"(?<![\w,.])"
    r"(\d{1,3}(?:\.\d{3})+(?:,\d{1,2})?"      # 1.234.567,89
    r"|\d{1,3}(?:,\d{3})+(?:\.\d{1,2})?"      # 1,234,567.89
    r"|\d+(?:[.,]\d{1,2})?)"                  # 12345 / 12345,50
    r"(?![\w])"
)


# ----------------------------------------------------------------- nombres
def normalizar(texto: str) -> str:
    """Mayúsculas, sin acentos, sin puntuación."""
    if not texto:
        return ""
    t = unicodedata.normalize("NFKD", str(texto))
    t = "".join(c for c in t if not unicodedata.combining(c))
    t = t.upper().replace("Ñ", "N")
    t = re.sub(r"[^A-Z0-9 ]+", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def tokens_nombre(texto: str) -> list[str]:
    return [
        t for t in normalizar(texto).split()
        if len(t) > 1 and t.lower() not in STOPWORDS and not t.isdigit()
    ]


def similitud_nombres(a: str, b: str) -> float:
    """0-1. Tolera orden invertido (apellido/nombre) y nombres parciales."""
    ta, tb = tokens_nombre(a), tokens_nombre(b)
    if not ta or not tb:
        return 0.0

    sa, sb = set(ta), set(tb)
    chico, grande = (sa, sb) if len(sa) <= len(sb) else (sb, sa)

    # Coincidencia token a token (el más parecido de la otra lista)
    puntajes = []
    for t in chico:
        mejor = max(
            (SequenceMatcher(None, t, o).ratio() for o in grande), default=0.0
        )
        # Iniciales: "J" contra "JUAN" no suma casi nada; ya filtramos len<=1
        puntajes.append(mejor)
    cobertura = sum(puntajes) / len(puntajes)

    # Comparación global con tokens ordenados alfabéticamente
    global_ratio = SequenceMatcher(
        None, " ".join(sorted(sa)), " ".join(sorted(sb))
    ).ratio()

    # Si el set chico está totalmente contenido, es casi seguro la misma persona
    contenido = 1.0 if chico <= grande else 0.0

    resultado = max(cobertura, global_ratio, contenido)

    # Un solo token en común (típico "Gomez 15000") es evidencia más débil:
    # se acepta, pero queda registrado con menor similitud.
    if len(chico) == 1 and len(grande) > 1:
        resultado *= 0.9

    return round(resultado, 4)


# ----------------------------------------------------------------- montos
def parse_monto(valor) -> float | None:
    """Interpreta '$ 1.234.567,89', '1234567.89', 15000 -> float."""
    if valor is None:
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    txt = str(valor).strip()
    if not txt:
        return None
    txt = re.sub(r"(?i)(ars|\$|pesos|usd)", " ", txt)
    m = _RE_MONTO.search(txt.replace(" ", ""))
    if not m:
        return None
    num = m.group(1)
    if "." in num and "," in num:
        # el separador decimal es el último que aparece
        if num.rfind(",") > num.rfind("."):
            num = num.replace(".", "").replace(",", ".")
        else:
            num = num.replace(",", "")
    elif "," in num:
        ent, _, dec = num.rpartition(",")
        num = f"{ent.replace('.', '')}.{dec}" if len(dec) <= 2 else num.replace(",", "")
    elif num.count(".") == 1:
        ent, _, dec = num.rpartition(".")
        if len(dec) == 3 and len(ent) <= 3:
            num = ent + dec           # 15.000 -> 15000
    else:
        num = num.replace(".", "")
    try:
        return float(num)
    except ValueError:
        return None


def fmt_monto(v: float | None) -> str:
    if v is None:
        return "—"
    s = f"{v:,.2f}".replace(",", "@").replace(".", ",").replace("@", ".")
    return f"${s}"


def montos_iguales(a: float | None, b: float | None, tolerancia: float = 0.0) -> bool:
    if a is None or b is None:
        return False
    return abs(a - b) <= max(tolerancia, 0.009)


# ------------------------------------------------------------------ documentos
_RE_DOC_ETIQUETA = re.compile(
    r"(?i)\b(?:cuit|cuil|dni|doc(?:umento)?|nro\.?\s*doc)\b[\s:./-]*"
    r"(\d{2}[-.\s]?\d{7,8}[-.\s]?\d|\d{1,3}(?:[.\s]\d{3}){2,3}|\d{7,11})"
)
_RE_CUIT = re.compile(r"\b(\d{2}[-.\s]\d{7,8}[-.\s]\d)\b")


def clave_doc(valor) -> str:
    """Normaliza CUIT/CUIL/DNI/CVU a una clave comparable.

    '20-44141951-3', '20441419513' y '44141951' -> la misma persona.
    Los números largos (CVU/CBU/ID) se recortan por precisión de Excel.
    """
    if valor is None:
        return ""
    txt = str(valor)
    if "e+" in txt.lower():                     # Excel guardó 8.57e+16
        try:
            txt = f"{float(txt):.0f}"
        except ValueError:
            pass
    d = re.sub(r"\D", "", txt)
    if len(d) < 7 or set(d) == {"0"}:
        return ""
    if len(d) == 11:
        return d[2:10]
    if len(d) in (7, 8):
        return d.zfill(8)
    return d


def extraer_doc(texto: str) -> str:
    """Saca el DNI/CUIT de un texto libre. '' si no hay."""
    if not texto:
        return ""
    m = _RE_DOC_ETIQUETA.search(texto)
    if m:
        return clave_doc(m.group(1))
    m = _RE_CUIT.search(texto)
    if m:
        return clave_doc(m.group(1))
    return ""


def docs_iguales(a, b) -> bool:
    ka, kb = clave_doc(a), clave_doc(b)
    if not ka or not kb:
        return False
    if ka == kb:
        return True
    # CVU/CBU/IDs largos: Excel pierde precisión en los últimos dígitos
    if len(ka) >= 12 and len(kb) >= 12:
        n = min(len(ka), len(kb), 14)
        return ka[:n] == kb[:n]
    return False


# ------------------------------------------------------------------ pie
def parse_pie(texto: str) -> tuple[str, float | None, str]:
    """Del pie de foto saca (nombre, monto, documento).

    Formatos tolerados:
      'Juan Perez 15000' · 'JUAN PEREZ - $15.000'
      'Brian David Munoz CUIT CUIL 44141951' -> doc, NO monto
      'Ana Gomez DNI 30.111.222 $25.500'
    """
    if not texto:
        return "", None, ""
    limpio = " ".join(str(texto).split())

    # 1) El documento sale primero y se saca del texto, para que no se
    #    confunda con un importe (un DNI de 8 dígitos parece $44.141.951).
    doc = ""
    m = _RE_DOC_ETIQUETA.search(limpio) or _RE_CUIT.search(limpio)
    if m:
        doc = clave_doc(m.group(1))
        limpio = limpio[: m.start()] + " " + limpio[m.end():]

    # 2) Monto: primero lo que tenga $; si no, un número con decimales.
    monto = None
    con_signo = re.search(r"\$\s*([\d.,]+)", limpio)
    if con_signo:
        monto = parse_monto(con_signo.group(1))
        limpio = limpio.replace(con_signo.group(0), " ")
    else:
        candidatos = []
        for mm in _RE_MONTO.finditer(limpio):
            crudo = mm.group(1)
            solo_digitos = re.sub(r"\D", "", crudo)
            tiene_decimales = bool(re.search(r"[.,]\d{1,2}$", crudo))
            # Un entero de 7-8 dígitos sin decimales suele ser un DNI suelto
            if not tiene_decimales and len(solo_digitos) in (7, 8, 11):
                if not doc:
                    doc = clave_doc(crudo)
                    limpio = limpio[: mm.start()] + " " + limpio[mm.end():]
                continue
            v = parse_monto(crudo)
            if v is not None and v >= 100:
                candidatos.append((mm, v))
        if candidatos:
            mm, monto = candidatos[-1]
            limpio = limpio[: mm.start()] + " " + limpio[mm.end():]

    # 3) Lo que queda, limpio de etiquetas, es el nombre
    nombre = re.sub(r"(?i)\b(monto|importe|nombre|titular|total|pago|transferencia|"
                    r"cuit|cuil|dni|documento|doc)\b[:\-]?", " ", limpio)
    nombre = re.sub(r"[\d$]+", " ", nombre)
    nombre = re.sub(r"[^\w\sÁÉÍÓÚÜÑáéíóúüñ]+", " ", nombre)
    nombre = " ".join(nombre.split())
    return nombre.strip(), monto, doc


# ----------------------------------------------------------------- fechas
def parse_fecha(texto: str, tz, fin_del_dia: bool = False) -> datetime | None:
    """Acepta dd/mm/aaaa, dd-mm-aaaa, aaaa-mm-dd, dd/mm."""
    if not texto:
        return None
    t = texto.strip()
    formatos = ["%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y", "%d/%m"]
    for f in formatos:
        try:
            d = datetime.strptime(t, f)
            if f == "%d/%m":
                d = d.replace(year=datetime.now(tz).year)
            if fin_del_dia:
                d = d.replace(hour=23, minute=59, second=59)
            return d.replace(tzinfo=tz)
        except ValueError:
            continue
    return None


def a_local(dt: datetime, tz) -> datetime:
    return dt.astimezone(tz)


# ----------------------------------------------------------------- huellas
def huella_comprobante(datos: dict) -> str:
    """Identidad lógica del comprobante, para detectar reenvíos.

    Prioriza el número de operación; si no hay, usa monto+fecha+destino.
    """
    nro = normalizar(datos.get("nro_operacion") or "")
    banco = normalizar(datos.get("banco") or "")
    if nro and len(nro) >= 6:
        base = f"OP|{nro}|{banco[:12]}"
    else:
        monto = parse_monto(datos.get("monto"))
        base = "|".join([
            "MIX",
            f"{monto:.2f}" if monto is not None else "?",
            str(datos.get("fecha") or "?"),
            " ".join(sorted(tokens_nombre(datos.get("nombre_destino") or ""))),
            normalizar(datos.get("cvu_destino") or "")[-8:],
        ])
    return hashlib.sha1(base.encode("utf-8")).hexdigest()[:20]


def hash_archivo(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()[:24]


def dentro_de(dt: datetime, desde: datetime | None, hasta: datetime | None) -> bool:
    if desde and dt < desde:
        return False
    if hasta and dt > hasta:
        return False
    return True


def cerca_en_tiempo(a: datetime, b: datetime, segundos: int = 90) -> bool:
    return abs((a - b).total_seconds()) <= segundos


__all__ = [n for n in dir() if not n.startswith("_")]
