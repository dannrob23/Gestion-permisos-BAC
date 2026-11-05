#!/usr/bin/env python3
"""Núcleo del Asistente de Correos BAC.

Procesa un correo .eml de Colsof Automate (Solicitud Ingreso), extrae los datos,
determina el aprobador por SBAN, reconstruye el correo con el formato EXACTO
(tabla azul #1f497d + foto) y lo deja listo para pegar en Outlook Web.

Uso:
    python asistente.py "ruta/correo.eml"            # procesa y muestra
    python asistente.py "ruta/correo.eml" --copy     # además copia al portapapeles
    python asistente.py "ruta/correo.eml" --html salida.html
"""
import argparse, base64, email, html as htmlmod, json, os, re, sqlite3, subprocess, sys, tempfile
from email import policy
from datetime import datetime
from zoneinfo import ZoneInfo
try:
    import fitz  # PyMuPDF para procesamiento de PDFs
except ImportError:
    fitz = None

BASE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(BASE, "datos")
DB = os.path.join(BASE, "base.db")

# ---------- Cargar mapeo y aprobadores ----------
def _cargar_json(nombre):
    ruta = os.path.join(DATA, nombre)
    if not os.path.exists(ruta):
        return {}
    try:
        with open(ruta, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

APROBADORES = _cargar_json("aprobadores.json")
MAPEO_SBAN = _cargar_json("mapeo_sban.json")

# Reemplazo temporal por regional (ej: vacaciones/permiso).
def _reemplazos_regional():
    datos = _cargar_json("reemplazos_aprobadores.json")
    out = {}
    for k, v in datos.items():
        if not isinstance(v, dict) or k.startswith("_"):
            continue
        out[k] = {"ingeniero": v.get("ingeniero", ""), "correo": v.get("correo", "")}
    return out

def init_novedades_db(con=None):
    """Crea la tabla de novedades de aprobadores y la puebla si está vacía."""
    cerrar = False
    if con is None:
        con = sqlite3.connect(DB)
        cerrar = True
    con.execute("""CREATE TABLE IF NOT EXISTS novedades_aprobadores (
        regional TEXT PRIMARY KEY,
        ingeniero_titular TEXT,
        correo_titular TEXT,
        ingeniero_reemplazo TEXT,
        correo_reemplazo TEXT,
        motivo TEXT DEFAULT 'Vacaciones',
        activo INTEGER DEFAULT 0,
        fecha_inicio TEXT DEFAULT '',
        fecha_fin TEXT DEFAULT '',
        actualizado_en TEXT DEFAULT ''
    )""")
    count = con.execute("SELECT COUNT(*) FROM novedades_aprobadores").fetchone()[0]
    if count == 0 and APROBADORES:
        reemps = _reemplazos_regional()
        ahora = datetime.now().strftime("%Y-%m-%d %H:%M")
        for reg, inf in APROBADORES.items():
            r = reemps.get(reg, {})
            activo = 1 if (r.get("ingeniero") and r.get("correo")) else 0
            con.execute("""INSERT OR REPLACE INTO novedades_aprobadores
                (regional, ingeniero_titular, correo_titular, ingeniero_reemplazo, correo_reemplazo, motivo, activo, actualizado_en)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                (reg, inf.get("ingeniero", ""), inf.get("correo", ""),
                 r.get("ingeniero", ""), r.get("correo", ""),
                 "Vacaciones" if activo else "", activo, ahora))
        con.commit()
    if cerrar:
        con.close()

def _sincronizar_json_aprobadores():
    """Mantiene sincronizados datos/reemplazos_aprobadores.json y datos/aprobadores.json."""
    try:
        con = sqlite3.connect(DB)
        con.row_factory = sqlite3.Row
        filas = con.execute("SELECT * FROM novedades_aprobadores").fetchall()
        con.close()
        
        aprob_dict = {}
        reemp_dict = {"_instrucciones": "Sincronizado automáticamente desde el Administrador de Aprobadores."}
        for f in filas:
            reg = f["regional"]
            aprob_dict[reg] = {"ingeniero": f["ingeniero_titular"], "correo": f["correo_titular"]}
            if f["activo"] and f["ingeniero_reemplazo"]:
                reemp_dict[reg] = {"ingeniero": f["ingeniero_reemplazo"], "correo": f["correo_reemplazo"], "motivo": f["motivo"]}
        
        ruta_aprob = os.path.join(DATA, "aprobadores.json")
        ruta_reemp = os.path.join(DATA, "reemplazos_aprobadores.json")
        with open(ruta_aprob, "w", encoding="utf-8") as fa:
            json.dump(aprob_dict, fa, indent=2, ensure_ascii=False)
        with open(ruta_reemp, "w", encoding="utf-8") as fr:
            json.dump(reemp_dict, fr, indent=2, ensure_ascii=False)
    except Exception:
        pass

def obtener_novedades_aprobadores():
    """Retorna lista de diccionarios con el estado de aprobadores por regional."""
    init_novedades_db()
    con = sqlite3.connect(DB)
    con.row_factory = sqlite3.Row
    filas = con.execute("SELECT * FROM novedades_aprobadores ORDER BY regional").fetchall()
    res = [dict(f) for f in filas]
    con.close()
    return res

def guardar_novedad_aprobador(regional, titular, correo_titular, reemplazo="", correo_reemplazo="", motivo="Vacaciones", activo=0, fecha_inicio="", fecha_fin=""):
    """Guarda o actualiza la novedad de un aprobador en SQLite y sincroniza el JSON."""
    init_novedades_db()
    con = sqlite3.connect(DB)
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M")
    con.execute("""INSERT OR REPLACE INTO novedades_aprobadores
        (regional, ingeniero_titular, correo_titular, ingeniero_reemplazo, correo_reemplazo, motivo, activo, fecha_inicio, fecha_fin, actualizado_en)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
        (regional, titular.strip(), correo_titular.strip(),
         reemplazo.strip(), correo_reemplazo.strip(),
         motivo.strip() or "Vacaciones", 1 if activo else 0,
         fecha_inicio.strip(), fecha_fin.strip(), ahora))
    con.commit()
    con.close()
    _sincronizar_json_aprobadores()

def toggle_novedad_aprobador(regional, activo):
    """Activa o desactiva el reemplazo temporal de una regional."""
    init_novedades_db()
    con = sqlite3.connect(DB)
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M")
    con.execute("UPDATE novedades_aprobadores SET activo=?, actualizado_en=? WHERE regional=?",
                (1 if activo else 0, ahora, regional))
    con.commit()
    con.close()
    _sincronizar_json_aprobadores()

def actualizar_titular_definitivo(regional, nuevo_titular, nuevo_correo):
    """Actualiza permanentemente el titular de una regional (por retiros/cambios)."""
    init_novedades_db()
    con = sqlite3.connect(DB)
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M")
    con.execute("UPDATE novedades_aprobadores SET ingeniero_titular=?, correo_titular=?, activo=0, actualizado_en=? WHERE regional=?",
                (nuevo_titular.strip(), nuevo_correo.strip(), ahora, regional))
    con.commit()
    con.close()
    _sincronizar_json_aprobadores()

def _aprobador_regional(regional):
    """Devuelve el aprobador de la regional aplicando novedades/reemplazos temporales desde SQLite."""
    try:
        init_novedades_db()
        con = sqlite3.connect(DB)
        con.row_factory = sqlite3.Row
        fila = con.execute("SELECT * FROM novedades_aprobadores WHERE regional=?", (regional,)).fetchone()
        con.close()
        if fila:
            if fila["activo"] and fila["ingeniero_reemplazo"] and fila["correo_reemplazo"]:
                return {
                    "ingeniero": fila["ingeniero_reemplazo"],
                    "correo": fila["correo_reemplazo"],
                    "regional": regional,
                    "en_reemplazo": True,
                    "motivo": fila["motivo"] or "Vacaciones",
                    "titular": fila["ingeniero_titular"],
                    "correo_titular": fila["correo_titular"]
                }
            elif fila["ingeniero_titular"] and fila["correo_titular"]:
                return {
                    "ingeniero": fila["ingeniero_titular"],
                    "correo": fila["correo_titular"],
                    "regional": regional,
                    "en_reemplazo": False,
                    "motivo": "",
                    "titular": fila["ingeniero_titular"],
                    "correo_titular": fila["correo_titular"]
                }
    except Exception:
        pass
    # Fallback legacy si la BD no está lista
    a = _reemplazos_regional().get(regional) or APROBADORES.get(regional)
    if not a:
        return None
    return {"ingeniero": a["ingeniero"], "correo": a["correo"], "regional": regional, "en_reemplazo": False}

FIRMA = os.path.join(BASE, "firma.html")

# Reglas especiales para SBAN sin oficina comercial (ej: 9000 = Dirección Nacional / DG).
# Los datos del aprobador se toman de Funcionarios_BAC (archivo), NO inventados.
REGLAS_ESPECIALES = {
    "9000": {
        "oficina": "DIRECCION NACIONAL / DG",
        "ingeniero": "Laureano Amado",
        "correo": "laureano.amado@bancoagrario.gov.co",
        "regional": "Dirección Nacional",
    },
    "3030": {
        "oficina": "DIRECCION GENERAL / DG",
        "ingeniero": "Freddy Andres Ortiz Rodriguez",
        "correo": "freddy.ortiz@bancoagrario.gov.co",
        "regional": "Dirección Nacional",
    },
}
def get_firma():
    if os.path.exists(FIRMA):
        with open(FIRMA, encoding="utf-8") as f:
            f_text = f.read().strip()
        if f_text:
            return f_text
    return '<p style="color:#555;font-size:0.9em">[AQUÍ VA TU FIRMA]</p>'


# ---------- 1. Leer .eml ----------
def leer_eml(ruta):
    with open(ruta, "rb") as f:
        msg = email.message_from_binary_file(f, policy=policy.default)
    datos = {
        "asunto": msg.get("Subject", ""),
        "from": msg.get("From", ""),
        "to": msg.get("To", ""),
        "cc": msg.get("Cc", ""),
        "date": msg.get("Date", ""),
        "recibido": (msg.get_all("Received") or [None])[0],   # hora de recepción (Outlook)
        "cuerpo_plano": "",
        "cuerpo_html": "",
        "foto": None,          # bytes
        "foto_nombre": None,
        "adjuntos": [],
    }
    body = msg.get_body(preferencelist=("html", "plain"))
    if body:
        if body.get_content_type() == "text/html":
            datos["cuerpo_html"] = body.get_content()
        else:
            datos["cuerpo_plano"] = body.get_content()
    # adjuntos (foto del funcionario)
    for part in msg.walk():
        fn = part.get_filename()
        if fn and part.get_content_type().startswith("image/"):
            datos["foto"] = part.get_content()
            datos["foto_nombre"] = fn
            datos["adjuntos"].append((fn, part.get_content_type(), part.get_content()))
    return datos


# ---------- 2. Extraer datos del cuerpo ----------
def limpiar(s):
    return re.sub(r"\s+", " ", s).strip()

def extraer_datos(datos):
    asunto = datos["asunto"]
    cuerpo = datos["cuerpo_html"] or datos["cuerpo_plano"]

    # SBAN del asunto: "4519-LA MACARENA" (primer número)
    m = re.search(r"-\s*(\d{2,5})\s*-", asunto)
    sban = m.group(1) if m else None
    if not sban:
        m = re.search(r"(\d{3,5})\s*[-–]", asunto)
        sban = m.group(1) if m else None

    # Caso: IMxxxx o N/A
    m = re.search(r"(\bIM\w+)", asunto) or re.search(r"Caso:\s*(\w+(?:/\w+)?)", cuerpo)
    caso = m.group(1) if m else ""

    # Datos desde el HTML (celdas con estilo de datos)
    empresa = funcionario = doc = eps = arl = arl_vence = celular = ""
    foto_html_cell = ""
    cells = re.findall(r'<td style="padding:0\.1em 1em">(.*?)</td>', cuerpo, re.S)
    if not cells:
        # Formato Outlook/manual: celdas con padding distinto; extraer todos los <td>
        # y tomar la fila de datos (la que contiene una cédula).
        # Formato Outlook/manual: extraer la fila <tr> que contiene la cédula
        rows = re.findall(r"<tr[^>]*>(.*?)</tr>", cuerpo, re.S)
        for row in rows:
            tds = re.findall(r"<td[^>]*>(.*?)</td>", row, re.S)
            texts = [limpiar(re.sub(r"<[^>]+>", "", t)) for t in tds]
            if any(re.search(r"\d{2}\.?\d{3}\.?\d{2,4}", t) for t in texts):
                cells = texts[:7]
                break
        if not cells:
            all_tds = [limpiar(re.sub(r"<[^>]+>", "", c)) for c in
                       re.findall(r'<td[^>]*>(.*?)</td>', cuerpo, re.S)]
            for i, txt in enumerate(all_tds):
                if re.search(r"\d{2}\.?\d{3}\.?\d{2,4}", txt):
                    row_start = (i // 7) * 7
                    cells = all_tds[row_start:row_start + 7]
                    break
    if len(cells) >= 6:
        empresa = limpiar(re.sub(r"<[^>]+>", "", cells[0]))
        funcionario = limpiar(re.sub(r"<[^>]+>", "", cells[1]))
        doc = limpiar(re.sub(r"<[^>]+>", "", cells[2]))
        eps = limpiar(re.sub(r"<[^>]+>", "", cells[3]))
        arl_cell = cells[4]
        arl_m = re.search(r"([^<>]+?)(?:\s*<br|\s*</td)", arl_cell, re.S)
        arl = limpiar(arl_m.group(1)) if arl_m else limpiar(re.sub(r"<[^>]+>", "", arl_cell))
        vence_m = re.search(r"Vence:\s*([\d/]+)", arl_cell)
        arl_vence = vence_m.group(1) if vence_m else ""
        celular = limpiar(re.sub(r"<[^>]+>", "", cells[5]))
        foto_html_cell = cells[6] if len(cells) > 6 else ""

    # Fecha de ingreso
    m = re.search(r"Fecha de Ingreso\s*([\d\- ]+)", cuerpo)
    fecha_ingreso = limpiar(m.group(1)) if m else ""

    # Oficina / Caso / Descripción
    m = re.search(r"Oficina:\s*(?:</[^>]+>\s*)*<[^>]*>([^<]+)", cuerpo)
    if not m:
        m = re.search(r"Oficina:\s*([^<\n]+)", cuerpo)
    oficina = limpiar(m.group(1)) if m else ""
    if not sban and oficina:
        m2 = re.search(r"(\d{2,5})", oficina)
        sban = m2.group(1) if m2 else sban
    m = re.search(r"Descripci[oó]n:\s*([^<\n]+)", cuerpo)
    descripcion = limpiar(m.group(1)) if m else ""

    # Regional desde paréntesis del asunto: "... (BOYACÁ)"
    m = re.search(r"\(([^)]+)\)", asunto)
    regional_asunto = limpiar(m.group(1)) if m else ""

    # Asunto limpio: "Solicitud Ingreso - SBAN-NOMBRE OFICINA" (sin [EXTERNAL])
    asunto_limpio = "Solicitud Ingreso"
    if sban:
        m2 = re.search(r"(\d{2,5})\s*[-–]\s*(.+)", oficina) if oficina else None
        nombre_oficina = limpiar(m2.group(2)) if m2 else (oficina or "")
        asunto_limpio = f"Solicitud Ingreso - {sban}-{nombre_oficina}"

    # Solicitante: el FUNCIONARIO (la persona para quien es el permiso)
    solicitante = funcionario or ""

    # Solicitudes que mencionan UPS: el solicitante es siempre Alejandro Bonilla
    if re.search(r"\bUPS\b", f"{asunto} {cuerpo}", re.I):
        solicitante = "Alejandro Bonilla"

    # Fecha/hora en que se RECIBIÓ el correo: usa la cabecera Received (la misma
    # que muestra Outlook), en hora de Bogotá. Si no hay Received, cae al Date.
    fecha_solicitud = ""
    fecha_dt = None
    try:
        from email.utils import parsedate_to_datetime
        rec_raw = datos.get("recibido")
        if rec_raw:
            m = re.search(r";\s*((?:\w{3},\s*)?\d{1,2}\s+\w{3}\s+\d{4}\s+\d{1,2}:\d{2}(?::\d{2})?\s+[+-]\d{4})", rec_raw)
            if m:
                fecha_dt = parsedate_to_datetime(m.group(1))
        if fecha_dt is None and (datos.get("date") or datos.get("Date", "")):
            fecha_dt = parsedate_to_datetime(datos["date"] or datos.get("Date", ""))
        if fecha_dt is not None:
            if fecha_dt.tzinfo is not None:
                try:
                    from zoneinfo import ZoneInfo
                    fecha_dt = fecha_dt.astimezone(ZoneInfo("America/Bogota"))
                except Exception:
                    pass
            fecha_solicitud = fecha_dt.strftime("%d/%m/%Y %H:%M")
    except Exception:
        fecha_solicitud = limpiar(datos.get("date") or "")[:16]

    return {
        "sban": sban, "caso": caso, "empresa": empresa, "funcionario": funcionario,
        "documento": doc, "eps": eps, "arl": arl, "arl_vence": arl_vence,
        "celular": celular, "fecha_ingreso": fecha_ingreso, "oficina": oficina,
        "descripcion": descripcion, "asunto": asunto_limpio, "asunto_original": asunto,
        "foto_html_cell": foto_html_cell, "solicitante": solicitante,
        "fecha_solicitud": fecha_solicitud,
    }


# ---------- 2b. Extraer datos desde TEXTO PLANO ----------
def extraer_datos_texto(texto):
    """Parsea los datos planos pegados del correo o extraídos de PDF.
    Soporta formato con tabs, múltiples espacios y formato secuencial multilínea.
    Devuelve dict compatible con extraer_datos()."""
    lineas = [l.strip() for l in texto.splitlines() if l.strip()]
    d = {k: "" for k in ("sban", "caso", "empresa", "funcionario", "documento",
                          "eps", "arl", "arl_vence", "celular", "fecha_ingreso",
                          "oficina", "descripcion", "asunto", "foto_html_cell",
                          "solicitante", "fecha_solicitud")}

    # Fecha de ingreso (una línea o multilínea)
    m = re.search(r"Fecha de Ingreso:?\s*([\d\-/ ]+)", texto, re.I)
    if not m:
        m = re.search(r"Fecha de Ingreso:?\s*\n\s*([\d\-/ ]+)", texto, re.I)
    if m:
        d["fecha_ingreso"] = limpiar(m.group(1))

    # Oficina (SBAN-NOMBRE), caso, descripción
    m = re.search(r"Oficina:\s*([^\n]+)", texto, re.I)
    if not m:
        m = re.search(r"Oficina:?\s*\n\s*([^\n]+)", texto, re.I)
    if m:
        d["oficina"] = limpiar(m.group(1))
        m2 = re.search(r"(\d{2,5})\s*[-–]\s*(.+)", d["oficina"])
        if m2:
            d["sban"] = m2.group(1)
            d["oficina_nombre"] = limpiar(m2.group(2))
        else:
            d["oficina_nombre"] = d["oficina"]
            m_sban = re.search(r"(\d{2,5})", d["oficina"])
            if m_sban:
                d["sban"] = m_sban.group(1)

    m = re.search(r"Caso:\s*([^\n]+)", texto, re.I)
    if not m:
        m = re.search(r"Caso:?\s*\n\s*([^\n]+)", texto, re.I)
    if m:
        c_val = limpiar(m.group(1))
        m_c = re.search(r"(\bIM\w+|\bN/A\b|\bNA\b|\w+(?:/\w+)?)", c_val, re.I)
        d["caso"] = m_c.group(1) if m_c else c_val

    m = re.search(r"Descripci[oó]n:\s*([^\n]+)", texto, re.I)
    if not m:
        m = re.search(r"Descripci[oó]n:?\s*\n\s*([^\n]+)", texto, re.I)
    if m:
        d["descripcion"] = limpiar(m.group(1))

    # ARL vence + celular (en cualquier parte del texto)
    m = re.search(r"Vence:\s*([\d/\-]+)", texto, re.I)
    if m:
        d["arl_vence"] = m.group(1)
    m = re.search(r"(\b3\d{9}\b)", texto)
    if m:
        d["celular"] = m.group(1)

    # Buscar fila o bloque de datos del contratista
    # 1. Buscar la línea con la cédula (con puntos 1.053... o sin puntos 7-10 dígitos)
    fila_idx = None
    for i, ln in enumerate(lineas):
        if re.search(r"\b\d{1,3}(?:\.\d{3}){2,3}\b", ln):
            fila_idx = i
            break
        elif re.search(r"\b\d{7,10}\b", ln):
            val = re.search(r"\b\d{7,10}\b", ln).group(0)
            if not (len(val) == 10 and val.startswith("3")):  # descartar número de celular
                fila_idx = i
                break

    if fila_idx is not None:
        ln_ced = lineas[fila_idx]
        campos = [c.strip() for c in re.split(r"\t+|\s{2,}", ln_ced) if c.strip()]
        if len(campos) >= 3:
            # Fila completa en una sola línea (con tabs o espacios dobles)
            d["empresa"] = campos[0]
            d["funcionario"] = campos[1]
            d["documento"] = campos[2]
            if len(campos) > 3: d["eps"] = campos[3]
            if len(campos) > 4:
                d["arl"] = campos[4]
                m_v = re.search(r"Vence:\s*([\d/\-]+)", d["arl"], re.I)
                if m_v:
                    d["arl_vence"] = m_v.group(1)
                    d["arl"] = re.sub(r"Vence:\s*[\d/\-]+", "", d["arl"], flags=re.I).strip()
            if len(campos) > 5 and not d["celular"]:
                d["celular"] = campos[5]
        else:
            # Formato secuencial multilínea (típico de exportación PDF de Outlook)
            m_doc = re.search(r"\b\d{1,3}(?:\.\d{3}){2,3}\b|\b\d{7,10}\b", ln_ced)
            d["documento"] = m_doc.group(0) if m_doc else ln_ced

            # Líneas previas a la cédula (funcionario y empresa)
            prev_lines = [lineas[j] for j in range(max(0, fila_idx - 6), fila_idx) if lineas[j]]
            header_terms = ["EMPRESA", "FUNCIONARIO", "DOCUMENTO", "IDENTIDAD", "INGRESO", "SOLICITUD", "DILIGENCIE", "AUTORIZACION", "POR FAVOR"]
            prev_clean = [l for l in prev_lines if not any(h in l.upper() for h in header_terms)]
            if len(prev_clean) >= 2:
                d["empresa"] = prev_clean[-2]
                d["funcionario"] = prev_clean[-1]
            elif len(prev_clean) == 1:
                d["funcionario"] = prev_clean[0]

            # Líneas posteriores a la cédula (EPS, ARL, Celular)
            next_lines = [lineas[j] for j in range(fila_idx + 1, min(len(lineas), fila_idx + 9)) if lineas[j]]
            stop_terms = ["OFICINA:", "CASO:", "DESCRIPCIÓN:", "DESCRIPCION:", "SE HACE LA SOLICITUD"]
            next_clean = []
            for l in next_lines:
                if any(l.upper().startswith(st) for st in stop_terms):
                    break
                next_clean.append(l)

            for l in next_clean:
                if re.search(r"\b3\d{9}\b", l):
                    if not d["celular"]:
                        d["celular"] = re.search(r"\b3\d{9}\b", l).group(0)
                elif "vence:" in l.lower():
                    m_v = re.search(r"vence:\s*([\d/\-]+)", l, re.I)
                    if m_v: d["arl_vence"] = m_v.group(1)
                    arl_part = re.sub(r"vence:\s*[\d/\-]+", "", l, flags=re.I).strip()
                    if arl_part and not d["arl"]:
                        d["arl"] = arl_part
                elif not d["eps"]:
                    d["eps"] = l
                elif not d["arl"]:
                    d["arl"] = l

    # Si falta oficina pero está en el texto como "Solicitud Ingreso - ..."
    if not d.get("oficina"):
        m = re.search(r"(?:Solicitud Ingreso|ingreso)\s*[-–]\s*([0-9A-Za-zÁÉÍÓÚÑáéíóúñü -]+)", texto, re.I)
        if m:
            d["oficina"] = limpiar(m.group(1))
            m2 = re.search(r"(\d{2,5})", d["oficina"])
            if m2 and not d.get("sban"):
                d["sban"] = m2.group(1)

    # Asunto automático: "Solicitud Ingreso - SBAN-NOMBRE OFICINA"
    if d.get("sban") and d.get("oficina_nombre"):
        d["asunto"] = f"Solicitud Ingreso - {d['sban']}-{d['oficina_nombre']}"
    elif d.get("oficina"):
        d["asunto"] = f"Solicitud Ingreso - {d['oficina']}"
    else:
        d["asunto"] = "Solicitud Ingreso"

    # Solicitante = el funcionario (persona para quien es el permiso)
    d["solicitante"] = d.get("funcionario", "") or ""
    if re.search(r"\bUPS\b", f"{d['asunto']} {texto}", re.I):
        d["solicitante"] = "Alejandro Bonilla"

    d["fecha_solicitud"] = datetime.now(ZoneInfo("America/Bogota")).strftime("%d/%m/%Y %H:%M")
    return d


# ---------- 3. Resolver aprobador ----------
def resolver_aprobador(sban):
    if not sban:
        return None
    # 1) Reglas especiales (SBAN 9000 = DG, etc.): datos desde Funcionarios_BAC
    if sban in REGLAS_ESPECIALES:
        r = REGLAS_ESPECIALES[sban]
        if r.get("correo"):
            return {
                "oficina": r["oficina"],
                "regional": r.get("regional", "Especial"),
                "ingeniero": r.get("ingeniero", r["correo"].split("@")[0]),
                "correo": r["correo"],
                "en_reemplazo": False,
                "titular": r.get("ingeniero", ""),
                "motivo": "",
            }
        aprob = _aprobador_regional(r["regional"])
        if not aprob:
            return None
        return {
            "oficina": r["oficina"],
            "regional": r["regional"],
            "ingeniero": aprob["ingeniero"],
            "correo": aprob["correo"],
            "en_reemplazo": aprob.get("en_reemplazo", False),
            "titular": aprob.get("titular", ""),
            "motivo": aprob.get("motivo", ""),
        }
    # 2) Mapeo normal de oficinas
    info = MAPEO_SBAN.get(sban)
    if not info:
        return None
    aprob = _aprobador_regional(info["regional"])
    if not aprob:
        return None
    return {
        "oficina": info["oficina"],
        "regional": info["regional"],
        "ingeniero": aprob["ingeniero"],
        "correo": aprob["correo"],
        "en_reemplazo": aprob.get("en_reemplazo", False),
        "titular": aprob.get("titular", ""),
        "motivo": aprob.get("motivo", ""),
    }


# ---------- 4. Generar HTML (réplica exacta) ----------
def generar_html(d, aprobador, foto_b64=None, firma=None):
    if firma is None:
        firma = get_firma()
    foto_cell = ""
    if foto_b64:
        foto_cell = (f'<img src="data:image/jpeg;base64,{foto_b64}" height="150" '
                     'style="object-fit: cover;" alt="Foto del funcionario">')
    elif d.get("foto_html_cell"):
        foto_cell = d["foto_html_cell"]

    arl_cell = htmlmod.escape(d["arl"])
    if d.get("arl_vence"):
        arl_cell += f'<br><span style="font-size:0.85em;color:#555">Vence: {htmlmod.escape(d["arl_vence"])}</span>'

    # Nota de regional: usa la regional del aprobador resuelto
    regional_nota = aprobador["regional"] if aprobador else "la regional correspondiente"
    if aprobador:
        nota_extra = ""
        if aprobador.get("en_reemplazo") and aprobador.get("titular"):
            mot = (aprobador.get("motivo") or "vacaciones").lower()
            nota_extra = f' (en reemplazo por {mot} de {htmlmod.escape(aprobador["titular"])})'
        nota_correo = (f'<p><strong>Nota:</strong> Este permiso es para <strong>'
                       f'{htmlmod.escape(aprobador["ingeniero"])}</strong> '
                       f'({htmlmod.escape(aprobador["correo"])}){nota_extra} — regional '
                       f'<strong>{htmlmod.escape(regional_nota)}</strong>.</p>')
    else:
        nota_correo = (f'<p><strong>Nota:</strong> Este permiso es para '
                       f'<strong>la regional correspondiente</strong> según la regional asignada.</p>')

    html_doc = f"""<html lang="es"><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style>
    table.datos {{ border-collapse: collapse; border: 1px solid black; }}
    table.datos td {{ border: 1px solid black; }}
</style>
</head>
<body>
    <p>Buen día,</p>
    <p>Solicito su colaboración y gestión con la siguiente autorización. </p>

    <table border="1" class="datos">
        <tr style="background-color:#1f497d">
            <td colspan="7">
                <p style="text-align:center;color:white">POR FAVOR DILIGENCIE LA SIGUIENTE INFORMACIÓN PARA SOLICITAR
                    LA AUTORIZACIÓN DE INGRESO
                </p>
                <p style="text-align:center;color:white">(Para solicitar autorización de ingreso de Proveedores,
                    Contratistas)</p>
            </td>
        </tr>
        <tr>
            <td colspan="7">Fecha de Ingreso {htmlmod.escape(d['fecha_ingreso'])}
            </td>
        </tr>
        <tr style="background-color:#1f497d;color:white">
            <td style="text-align:center;padding:1em">NOMBRE DE LA EMPRESA A LA QUE PERTENECE</td>
            <td style="text-align:center;padding:1em">NOMBRE DEL FUNCIONARIO</td>
            <td style="padding:1em">No. DOCUMENTO DE IDENTIDAD</td>
            <td style="text-align:center;padding:1em">EPS</td>
            <td style="text-align:center;padding:1em">ARL</td>
            <td style="text-align:center;padding:1em">CELULAR</td>
            <td style="text-align:center;padding:1em">FOTO</td>
        </tr>
        <tr>
            <td style="padding:0.1em 1em">{htmlmod.escape(d['empresa'])}</td>
            <td style="padding:0.1em 1em">{htmlmod.escape(d['funcionario'])}</td>
            <td style="padding:0.1em 1em">{htmlmod.escape(d['documento'])}</td>
            <td style="padding:0.1em 1em">{htmlmod.escape(d['eps'])}</td>
            <td style="padding:0.1em 1em">{arl_cell}</td>
            <td style="padding:0.1em 1em">{htmlmod.escape(d['celular'])}</td>
            <td style="padding:0.1em 1em;text-align:center">{foto_cell}</td>
        </tr>
        <tr>
            <td colspan="7">
                <p style="margin:0!important;padding:0!important">Se hace la solicitud de ingreso a la siguiente
                    Oficina: {htmlmod.escape(d['oficina'])}</p>
                <p style="margin:0!important;padding:0!important">Caso: {htmlmod.escape(d['caso'])}</p>
                <p style="margin:0!important;padding:0!important">Descripción: {htmlmod.escape(d['descripcion'])}
                </p>
            </td>
        </tr>
    </table>
    {nota_correo}
    <p>Cordial saludo,</p>
    {firma}
</body>
</html>"""
    return html_doc


# ---------- 5. Portapapeles (HTML formateado vía PowerShell) ----------
def copiar_portapapeles(html_doc):
    tmp = os.path.join(tempfile.gettempdir(), "asistente_correo_tmp.html")
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(html_doc)
    ps = (
        "$html = Get-Content -Raw -Encoding UTF8 '" + tmp.replace("'", "''") + "'; "
        "Set-Clipboard -AsHtml -Value $html"
    )
    r = subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                       capture_output=True, text=True)
    if r.returncode != 0:
        raise RuntimeError(f"PowerShell error: {r.stderr[:300]}")
    return tmp


# ---------- 5b. Envío de correo (SMTP configurado localmente) ----------
def enviar_correo(html_doc, destino, asunto):
    """Envía el correo formateado usando SMTP configurado en config_smtp.json.
    NOTA: Se eliminó el uso de Gmail (claves expuestas). Configura SMTP Colsof
    con `configurar_smtp.py` para usar el envío directo."""
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText

    cfg_path = os.path.join(BASE, "config_smtp.json")
    if not os.path.exists(cfg_path):
        raise RuntimeError("No se encontró config_smtp.json. Configura SMTP con configurar_smtp.py")

    with open(cfg_path, encoding="utf-8") as f:
        cfg = json.load(f)

    if not cfg.get("password"):
        raise RuntimeError("SMTP no configurado (password vacía). Ejecuta configurar_smtp.py")

    remitente = f"{cfg['remitente_nombre']} <{cfg['remitente_correo']}>"
    host, port = cfg["smtp_host"], cfg["smtp_port"]
    login, pw = cfg["usuario"], cfg["password"]

    msg = MIMEMultipart("alternative")
    msg["From"] = remitente
    msg["To"] = destino
    msg["Subject"] = asunto
    msg.attach(MIMEText("Mensaje con formato - ver versión HTML.", "plain", "utf-8"))
    msg.attach(MIMEText(html_doc, "html", "utf-8"))

    with smtplib.SMTP(host, port, timeout=30) as s:
        s.ehlo()
        s.starttls()
        s.ehlo()
        s.login(login, pw)
        s.sendmail(login, [destino], msg.as_string())
    return destino


# ---------- 5c. Procesar TEXTO PLANO (flujo completo) ----------
def procesar_texto(texto, copiar=False):
    d = extraer_datos_texto(texto)
    aprobador = resolver_aprobador(d["sban"])
    html_doc = generar_html(d, aprobador)
    estado = "PENDIENTE APROBACIÓN"
    notas = ""
    if not aprobador:
        notas = f"SBAN {d['sban'] or 'desconocido'} sin aprobador en el mapeo"
    con = init_db()
    reg_id = registrar(con, d, aprobador, estado, notas)
    con.close()
    tmp_html = copiar_portapapeles(html_doc) if copiar else None
    return {
        "id": reg_id, "datos": d, "aprobador": aprobador,
        "html": html_doc, "portapapeles": tmp_html, "estado": estado, "notas": notas,
        "asunto": d["asunto"],
    }


# ---------- 5d. OCR de imágenes (capturas de pantalla) ----------
_ocr_engine = None

def ocr_imagen(ruta_imagen):
    """Extrae texto de una imagen (captura de pantalla) usando RapidOCR local."""
    global _ocr_engine
    if _ocr_engine is None:
        from rapidocr_onnxruntime import RapidOCR
        _ocr_engine = RapidOCR()
    result, _ = _ocr_engine(ruta_imagen)
    if not result:
        return ""
    lineas = [line[1] for line in result]
    return "\n".join(lineas)


def procesar_imagen(ruta_imagen, copiar=False):
    """Procesa una captura de pantalla: OCR -> datos -> aprobador -> HTML -> registro."""
    texto = ocr_imagen(ruta_imagen)
    if not texto.strip():
        raise ValueError("No se detectó texto en la imagen. Sube una captura más nítida o usa el .eml.")
    return procesar_texto(texto, copiar=copiar)


# ---------- 5e. Procesar PDF (exportado de Outlook Móvil o Web) ----------
def extraer_datos_pdf(ruta_o_bytes):
    """Extrae texto, campos y foto de un PDF de solicitud (Outlook Móvil/Web).
    Soporta ruta a archivo o bytes en memoria."""
    if fitz is None:
        raise ImportError("PyMuPDF (fitz) no está disponible. Instálalo con: pip install pymupdf")

    if isinstance(ruta_o_bytes, (bytes, bytearray)):
        doc = fitz.open(stream=ruta_o_bytes, filetype="pdf")
    else:
        doc = fitz.open(ruta_o_bytes)

    # 1. Extraer texto completo
    texto_completo = ""
    for page in doc:
        texto_completo += page.get_text("text") + "\n"

    # Si el texto es mínimo o escaneado, intentar OCR sobre la primera página
    if len(texto_completo.strip()) < 30 and len(doc) > 0:
        try:
            pix = doc[0].get_pixmap(dpi=150)
            img_bytes = pix.tobytes("png")
            tmp_img = os.path.join(tempfile.gettempdir(), "pdf_ocr_tmp.png")
            with open(tmp_img, "wb") as f_tmp:
                f_tmp.write(img_bytes)
            texto_completo = ocr_imagen(tmp_img)
            if os.path.exists(tmp_img):
                os.remove(tmp_img)
        except Exception:
            pass

    # 2. Extraer foto del contratista
    foto_b64 = None
    mejor_img = None
    mayor_area = 0
    for page in doc:
        for img_info in page.get_images(full=True):
            xref = img_info[0]
            try:
                base_img = doc.extract_image(xref)
                if not base_img:
                    continue
                w = base_img.get("width", 0)
                h = base_img.get("height", 0)
                img_data = base_img.get("image")
                area = w * h
                # Descartar iconos, firmas o píxeles pequeños (< 50x50 o < 500 bytes)
                if w >= 50 and h >= 50 and len(img_data) > 500:
                    if area > mayor_area:
                        mayor_area = area
                        mejor_img = img_data
            except Exception:
                continue
    doc.close()

    if mejor_img:
        foto_b64 = base64.b64encode(mejor_img).decode()

    # 3. Parsear texto
    d = extraer_datos_texto(texto_completo)
    return d, foto_b64, texto_completo


def procesar_pdf(ruta_o_bytes, copiar=False):
    """Procesa un archivo PDF: extrae datos + foto -> resuelve aprobador -> HTML -> registra."""
    d, foto_b64, texto_crudo = extraer_datos_pdf(ruta_o_bytes)
    aprobador = resolver_aprobador(d["sban"])
    html_doc = generar_html(d, aprobador, foto_b64=foto_b64)

    estado = "PENDIENTE APROBACIÓN"
    notas = ""
    if not aprobador:
        notas = f"SBAN {d['sban'] or 'desconocido'} sin aprobador en el mapeo"

    con = init_db()
    reg_id = registrar(con, d, aprobador, estado, notas)
    con.close()

    tmp_html = copiar_portapapeles(html_doc) if copiar else None

    return {
        "id": reg_id,
        "datos": d,
        "aprobador": aprobador,
        "html": html_doc,
        "portapapeles": tmp_html,
        "estado": estado,
        "notas": notas,
        "foto_b64": foto_b64,
        "asunto_original": d.get("asunto", ""),
        "asunto": d["asunto"],
        "texto_crudo": texto_crudo,
    }


# ---------- 6. Base de datos ----------
def init_db():
    con = sqlite3.connect(DB)
    con.execute("""CREATE TABLE IF NOT EXISTS historial (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fecha_proceso TEXT, asunto TEXT, sban TEXT, caso TEXT, oficina TEXT,
        regional TEXT, aprobador TEXT, correo_aprobador TEXT,
        estado TEXT DEFAULT 'PENDIENTE', notas TEXT
    )""")
    # migración: columna solicitante (si no existe)
    cols = [r[1] for r in con.execute("PRAGMA table_info(historial)")]
    if "solicitante" not in cols:
        con.execute("ALTER TABLE historial ADD COLUMN solicitante TEXT DEFAULT ''")
    if "fecha_solicitud" not in cols:
        con.execute("ALTER TABLE historial ADD COLUMN fecha_solicitud TEXT DEFAULT ''")
    if "fecha_respuesta" not in cols:
        con.execute("ALTER TABLE historial ADD COLUMN fecha_respuesta TEXT DEFAULT ''")
    init_novedades_db(con)
    con.commit()
    return con

def registrar(con, d, aprobador, estado="PENDIENTE", notas=""):
    ahora = datetime.now().isoformat(timespec="seconds")
    caso = (d.get("caso") or "").strip()
    # Si la misma solicitud (caso + oficina) ya existe, se actualiza en vez de
    # crear un duplicado (evita que el historial replique información).
    if caso:
        fila = con.execute(
            "SELECT id FROM historial WHERE TRIM(caso)=? AND oficina=? ORDER BY id DESC LIMIT 1",
            (caso, d["oficina"])).fetchone()
        if fila:
            reg_id = fila[0]
            # Conserva el solicitante ya corregido manualmente si el correo no lo trae
            solicitante = d.get("solicitante", "").strip()
            if not solicitante:
                exist = con.execute("SELECT solicitante FROM historial WHERE id=?", (reg_id,)).fetchone()
                solicitante = (exist[0] or "") if exist else ""
            con.execute("""UPDATE historial SET
                fecha_proceso=?, asunto=?, sban=?, regional=?, aprobador=?,
                correo_aprobador=?, estado=?, notas=?, solicitante=?, fecha_solicitud=?
                WHERE id=?""",
                (ahora, d["asunto"], d["sban"],
                 aprobador["regional"] if aprobador else "SIN REGIONAL",
                 aprobador["ingeniero"] if aprobador else "SIN APROBADOR",
                 aprobador["correo"] if aprobador else "",
                 estado, notas, solicitante, d.get("fecha_solicitud", ""),
                 reg_id))
            con.commit()
            return reg_id
    con.execute("""INSERT INTO historial
        (fecha_proceso, asunto, sban, caso, oficina, regional, aprobador, correo_aprobador, estado, notas, solicitante, fecha_solicitud)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
        (ahora, d["asunto"], d["sban"], d["caso"],
         d["oficina"], aprobador["regional"] if aprobador else "SIN REGIONAL",
         aprobador["ingeniero"] if aprobador else "SIN APROBADOR",
         aprobador["correo"] if aprobador else "",
         estado, notas, d.get("solicitante", ""), d.get("fecha_solicitud", "")))
    con.commit()
    return con.execute("SELECT last_insert_rowid()").fetchone()[0]


# ---------- 7. Flujo principal ----------
def procesar(ruta_eml, copiar=False):
    datos = leer_eml(ruta_eml)
    d = extraer_datos(datos)
    aprobador = resolver_aprobador(d["sban"])

    foto_b64 = None
    if datos["foto"]:
        foto_b64 = base64.b64encode(datos["foto"]).decode()

    html_doc = generar_html(d, aprobador, foto_b64)

    estado = "PENDIENTE APROBACIÓN"
    notas = ""
    if not aprobador:
        notas = f"SBAN {d['sban'] or 'desconocido'} sin aprobador en el mapeo"

    con = init_db()
    reg_id = registrar(con, d, aprobador, estado, notas)
    con.close()

    tmp_html = copiar_portapapeles(html_doc) if copiar else None

    return {
        "id": reg_id, "datos": d, "aprobador": aprobador,
        "html": html_doc, "portapapeles": tmp_html, "estado": estado, "notas": notas,
        "foto_b64": foto_b64, "asunto_original": datos["asunto"],
        "asunto": d["asunto"],
    }


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Asistente de Correos BAC")
    ap.add_argument("eml", help="Ruta al archivo .eml")
    ap.add_argument("--copy", action="store_true", help="Copiar correo al portapapeles")
    ap.add_argument("--html", help="Guardar HTML en archivo")
    args = ap.parse_args()

    res = procesar(args.eml, copiar=args.copy)
    d, apb = res["datos"], res["aprobador"]
    print("=" * 60)
    print(f"Registro #{res['id']} | {res['estado']}")
    print(f"  Asunto  : {d['asunto']}")
    print(f"  SBAN    : {d['sban']} | Oficina: {d['oficina']}")
    print(f"  Caso    : {d['caso']}")
    print(f"  Fecha   : {d['fecha_ingreso']}")
    print(f"  Empresa : {d['empresa']}")
    print(f"  Persona : {d['funcionario']} | Doc: {d['documento']}")
    print(f"  EPS/ARL : {d['eps']} | {d['arl']} ({d['arl_vence']})")
    print(f"  Celular : {d['celular']}")
    print(f"  Desc    : {d['descripcion'][:80]}...")
    print("-" * 60)
    if apb:
        print(f"  ▶ ENVIAR A: {apb['ingeniero']} <{apb['correo']}>")
        print(f"    Regional: {apb['regional']} | Oficina mapeo: {apb['oficina']}")
    else:
        print(f"  ⚠ SIN APROBADOR ({res['notas']})")
    if args.html:
        with open(args.html, "w", encoding="utf-8") as f:
            f.write(res["html"])
        print(f"  HTML guardado: {args.html}")
    if res["portapapeles"]:
        print("  ✅ Copiado al portapapeles (pega en Outlook Web)")
