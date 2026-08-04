# Bot Verificador de Comprobantes (Telegram + Claude)

Userbot de Telethon que lee el historial de un grupo, analiza cada foto de
comprobante bancario argentino con Claude y la contrasta contra el pie de texto
(nombre + monto). Guarda un historial persistente para detectar reenvíos y cruza
todo contra el Excel del admin.

---

## 1. Archivos

| Archivo | Qué hace |
|---|---|
| `main.py` | cliente Telethon, comandos, pipeline de verificación |
| `analizador.py` | lectura del comprobante con Claude + veredicto |
| `excel.py` | lee el Excel del admin, cruza y arma los reportes `.xlsx` |
| `storage.py` | persistencia en el Volume (historial global + verificaciones) |
| `utils.py` | montos argentinos, nombres, fechas, huellas |
| `config.py` | variables de entorno |
| `generar_session.py` | genera `TELEGRAM_SESSION` (correr **local**) |

## 2. Regenerar la sesión (pendiente por el flood wait)

En tu máquina, no en Railway:

```bash
pip install telethon
python generar_session.py
```

Te pide el código que llega a Telegram y escupe la string session.
Copiala en Railway → Variables → `TELEGRAM_SESSION`.

**Si aparece `FloodWaitError`:** esperá el tiempo exacto que indica antes de
volver a intentar. Cada reintento prematuro reinicia el contador, y ese es
justamente el motivo por el que quedó bloqueada. Un solo intento cuando venza.

## 3. Railway

1. Reanudá el servicio `Bot-verificador` (proyecto `faithful-adventure`).
2. **Volume**: montalo en `/data` (Settings → Volumes → Mount path `/data`).
3. Variables (ver `.env.example`):
   `TELEGRAM_API_ID`, `TELEGRAM_API_HASH`, `TELEGRAM_PHONE`, `TELEGRAM_SESSION`,
   `ANTHROPIC_API_KEY`, `DATA_DIR=/data`.
4. Start command: `python main.py` (ya está en `railway.json`).

El servicio es un *worker*: no expone puerto y no necesita healthcheck HTTP.

## 4. Uso

Mandate los comandos a vos mismo (Mensajes Guardados) o desde un `ADMIN_IDS`.

```
/grupos                              lista los chats con su ID
/grupos cobranza                     filtra por nombre
/verificar -1001234567890            TODO el historial, sin límite
/verificar -1001234567890 01/08/2026            desde esa fecha
/verificar -1001234567890 01/08/2026 31/08/2026 rango cerrado
/reset -1001234567890                limpia la verificación, NO el historial
/estado                              qué hay corriendo y cuánto hay guardado
/cancelar [CHAT_ID]                  corta una verificación en curso
```

Después de verificar, mandá el **`.xlsx` al privado** (opcionalmente con el pie
`/excel -1001234567890`) y devuelve el cruce:

| | |
|---|---|
| ✅ | coinciden nombre + monto |
| 🔁 | duplicados o comprobantes viejos reenviados |
| ⚠️ | están en el chat pero no en el Excel |
| ❌ | están en el Excel pero no en el chat |

El Excel puede tener el encabezado en cualquier fila: se detectan columnas con
títulos tipo *Nombre / Apellido / Titular* e *Importe / Monto / Pago / Total*.
Si no encuentra encabezados, asume A = nombre, B = monto.

## 5. Cómo decide cada comprobante

1. Descarga la imagen y calcula su hash. Si esa imagen **ya se procesó alguna
   vez** (aunque haya sido en otro chat o hace meses) → 🔁 sin gastar API.
2. Claude extrae banco, monto, fecha, destinatario, N° de operación.
3. Se arma una **huella lógica**: N° de operación si existe; si no,
   monto + fecha + destinatario + CVU. Si esa huella ya está en el historial → 🔁.
4. Se compara contra el pie:
   - montos con tolerancia `TOLERANCIA_MONTO` (0 por defecto),
   - nombres normalizados sin acentos, tolerando orden invertido y nombres
     parciales, con umbral `UMBRAL_NOMBRE` (0.80).
5. Todo se guarda en `/data/historial.json` (permanente) y en
   `/data/verificaciones/<chat_id>.json` (lo que borra `/reset`).

El pie puede venir como caption de la foto o como mensaje de texto del mismo
remitente dentro de los 90 segundos (antes o después). Los álbumes heredan el
caption del primer elemento.

## 6. Costos y velocidad

Una llamada a Claude por comprobante nuevo. `claude-sonnet-5` lee bien
comprobantes borrosos o con capturas recortadas; si el volumen es alto, probá
`ANTHROPIC_MODEL=claude-haiku-4-5-20251001`. Las imágenes se redimensionan a
1568 px antes de enviarse. `MAX_CONCURRENCIA=3` es conservador para no comerte
un flood wait de Telegram en las descargas; podés subirlo a 5-6.

## 7. Seguridad

- `TELEGRAM_SESSION` da **acceso total** a la cuenta Fixcal: solo en variables
  de Railway, nunca en el repo (`.gitignore` ya cubre `.env` y `*.session`).
- Si el `api_hash` estuvo expuesto, creá una app nueva en
  https://my.telegram.org y reemplazá `TELEGRAM_API_ID` / `TELEGRAM_API_HASH`.
- Cualquiera en `ADMIN_IDS` puede leer historiales completos. Dejalo vacío si
  solo vas a operar desde tu propia cuenta.
