#!/usr/bin/env python3
"""Dashboard del Asistente de Correos BAC (Streamlit).

Flujo:
1. Subir el .eml descargado de Outlook Web (o escribir la ruta).
2. Ver la vista previa del correo reconstruido + aprobador asignado.
3. Botón "Copiar al portapapeles" → pegar en Outlook Web.
4. Marcar estado (Recibido / Pendiente aprobación / Enviado / Aprobado / Rechazado).
5. Historial completo con autor, fecha, caso, oficina, regional, estado.

Ejecutar:  streamlit run dashboard.py
"""
import base64, glob, io, json, os, re, sqlite3, subprocess, tempfile, time, sys, urllib.parse
from datetime import datetime
from zoneinfo import ZoneInfo

import streamlit as st
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import asistente as A
import sincronizar_publico as SINC

st.set_page_config(page_title="Asistente Correos BAC", page_icon="📧", layout="wide")

# ---------- Estilos (tema claro) ----------
st.markdown("""
<style>
    .stApp { background-color: #ffffff; color: #111827; }
    .card { background:#ffffff; border:1px solid #d5dbe5; border-radius:12px; padding:18px;
            margin-bottom:14px; color:#111827; }
    .badge { display:inline-block; padding:3px 12px; border-radius:20px; font-size:0.85rem;
             font-weight:600; }
    .b-pend { background:#fff7e6; color:#b45309; border:1px solid #f59e0b88; }
    .b-env  { background:#eff6ff; color:#1d4ed8; border:1px solid #60a5fa99; }
    .b-not  { background:#f0fdf4; color:#15803d; border:1px solid #4ade8088; }
    .b-rev  { background:#fef2f2; color:#b91c1c; border:1px solid #ef444488; }
</style>
""", unsafe_allow_html=True)

st.title("📧 Asistente de Correos BAC — Solicitudes de Ingreso")
st.caption("Sube el correo (.eml) de Colsof Automate → genera el reenvío con formato oficial → copia y pega en Outlook Web.")

# ---------- Helper ----------
def render_copiar_cliente(html_code, text_plain=""):
    """Botón de copiar multiplataforma (Móvil/PC) usando iframe aislado.
    Se renderiza dentro de su propio iframe para no interferir con el DOM de React/Streamlit."""
    html_json = json.dumps(html_code)
    text_clean = text_plain or re.sub(r'<[^>]+>', ' ', html_code)
    text_clean = " ".join(text_clean.split())
    text_json = json.dumps(text_clean)

    # Todo el HTML está dentro de un documento iframe completo y aislado.
    # Nunca toca document.body del padre — evita el error removeChild de React.
    comp = f"""<!DOCTYPE html>
<html>
<head><meta charset="utf-8">
<style>
  body {{ margin:0; padding:4px 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; }}
  #copy-btn {{
    width: 100%; background-color: #1f497d; color: #fff;
    font-size: 1rem; font-weight: 600; padding: 11px 16px;
    border: none; border-radius: 8px; cursor: pointer;
    box-shadow: 0 2px 4px rgba(0,0,0,0.15);
    display: flex; align-items: center; justify-content: center; gap: 8px;
  }}
  #copy-btn:hover {{ background-color: #163657; }}
  #copy-status {{ font-size: 0.84rem; color: #15803d; text-align: center;
    margin-top: 6px; display: none; font-weight: 600; }}
</style>
</head>
<body>
  <button id="copy-btn" onclick="copiar()">
    <span>📋</span> <span id="btn-text">Copiar correo con formato (Móvil / PC)</span>
  </button>
  <div id="copy-status">✅ ¡Copiado! Abre Outlook y pega con Ctrl+V o manteniendo presionado</div>
  <script>
    const HTML_SRC = {html_json};
    const TEXT_SRC = {text_json};

    async function copiar() {{
      const btn = document.getElementById('copy-btn');
      const lbl = document.getElementById('btn-text');
      const ok  = document.getElementById('copy-status');
      try {{
        if (typeof ClipboardItem !== 'undefined' && navigator.clipboard && navigator.clipboard.write) {{
          const bH = new Blob([HTML_SRC], {{type:'text/html'}});
          const bT = new Blob([TEXT_SRC], {{type:'text/plain'}});
          await navigator.clipboard.write([new ClipboardItem({{'text/html':bH,'text/plain':bT}})]);
        }} else {{
          await navigator.clipboard.writeText(TEXT_SRC);
        }}
        btn.style.backgroundColor = '#15803d';
        lbl.innerText = '¡Copiado con éxito!';
        ok.style.display = 'block';
        setTimeout(() => {{
          btn.style.backgroundColor = '#1f497d';
          lbl.innerText = 'Copiar correo con formato (Móvil / PC)';
        }}, 4000);
      }} catch(e) {{
        // Fallback: área de texto temporal confinada en el iframe
        const ta = document.createElement('textarea');
        ta.value = TEXT_SRC;
        ta.style.cssText = 'position:absolute;left:-9999px;top:0;opacity:0;';
        document.body.appendChild(ta);
        ta.focus(); ta.select();
        try {{ document.execCommand('copy'); }} catch(_) {{}}
        document.body.removeChild(ta);
        btn.style.backgroundColor = '#15803d';
        lbl.innerText = '¡Copiado (texto)!';
        ok.style.display = 'block';
        setTimeout(() => {{
          btn.style.backgroundColor = '#1f497d';
          lbl.innerText = 'Copiar correo con formato (Móvil / PC)';
        }}, 4000);
      }}
    }}
  </script>
</body>
</html>"""
    st.components.v1.html(comp, height=78)

def copiar_html_portapapeles(html_doc):
    tmp = os.path.join(tempfile.gettempdir(), "asistente_dash_tmp.html")
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(html_doc)
    try:
        ps = ("$h = Get-Content -Raw -Encoding UTF8 '" + tmp.replace("'", "''") + "'; "
              "Set-Clipboard -AsHtml -Value $h")
        r = subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                           capture_output=True, text=True)
        if r.returncode == 0:
            return True, tmp
    except Exception:
        pass
    return False, tmp

def marcar_estado(reg_id, estado):
    con = A.init_db()
    if estado in ("APROBADO", "RECHAZADO"):
        ahora = datetime.now(ZoneInfo("America/Bogota")).strftime("%d/%m/%Y %H:%M")
        con.execute(
            "UPDATE historial SET estado=?, fecha_respuesta="
            "CASE WHEN fecha_respuesta IS NULL OR fecha_respuesta='' THEN ? "
            "ELSE fecha_respuesta END WHERE id=?",
            (estado, ahora, reg_id))
    else:
        con.execute("UPDATE historial SET estado=? WHERE id=?", (estado, reg_id))
    con.commit()
    con.close()

def _parse_fecha_resp(estado, fecha_resp):
    """Devuelve datetime de fecha_respuesta o None si no aplica."""
    if estado != "APROBADO" or not fecha_resp:
        return None
    try:
        return datetime.strptime(str(fecha_resp).strip()[:16], "%d/%m/%Y %H:%M")
    except Exception:
        return None

def leer_historial(filtrar_aprobados=True):
    """Lee el historial. Si filtrar_aprobados=True, oculta APROBADO con más de 2 días
    de antigüedad (accesible siempre vía histórico descendente o descarga)."""
    con = A.init_db()
    df = pd.read_sql_query("SELECT * FROM historial ORDER BY id DESC", con)
    con.close()
    if filtrar_aprobados and not df.empty:
        if "fecha_respuesta" in df.columns:
            df["_es_visible"] = True
            mask = (df["estado"] == "APROBADO") & df["fecha_respuesta"].notna() & (df["fecha_respuesta"] != "")
            if mask.any():
                visibles = ~(mask & (df["fecha_respuesta"].map(
                    lambda x: _dias_aprobado(x) > 2)))
                df = df[visibles].copy()
        df = df.drop(columns=["_es_visible"], errors="ignore")
    return df

def _dias_aprobado(fecha_resp_str):
    """Días desde la fecha de respuesta. >2 días significa histórico."""
    try:
        dt = datetime.strptime(str(fecha_resp_str).strip()[:16], "%d/%m/%Y %H:%M")
        return (datetime.now() - dt).days
    except Exception:
        return 0

ESTADOS = ["RECIBIDO", "PENDIENTE APROBACIÓN", "ENVIADO", "APROBADO",
           "RECHAZADO"]

# Semáforo: color por estado (hex) + texto legible sobre fondo claro
COLOR_ESTADO = {
    "RECIBIDO": "#94a3b8",              # gris: aún no procesado
    "PENDIENTE APROBACIÓN": "#fbbf24",  # amarillo: en espera de aprobación
    "ENVIADO": "#60a5fa",               # azul: enviado al aprobador
    "APROBADO": "#4ade80",              # verde: autorizado
    "RECHAZADO": "#f87171",             # rojo: rechazado
}
COLOR_ESTADO_TEXTO = {
    "RECIBIDO": "#475569",
    "PENDIENTE APROBACIÓN": "#b45309",
    "ENVIADO": "#1d4ed8",
    "APROBADO": "#15803d",
    "RECHAZADO": "#b91c1c",
}

def _color_de_estado(estado):
    c = COLOR_ESTADO.get(str(estado).strip().upper(), COLOR_ESTADO.get(str(estado), "#94a3b8"))
    t = COLOR_ESTADO_TEXTO.get(str(estado).strip().upper(),
                               COLOR_ESTADO_TEXTO.get(str(estado), "#334155"))
    return c, t

def chip_estado(estado):
    """Badge HTML con el color del semáforo para un estado (legible sobre blanco)."""
    c, t = _color_de_estado(estado)
    return (f'<span style="display:inline-block;padding:2px 10px;margin-right:6px;'
            f'border-radius:20px;font-size:0.85rem;font-weight:600;'
            f'background:{c}26;color:{t};border:1px solid {c}88">{estado}</span>')

def semaforo_leyenda():
    """Fila de badges (semáforo) para todos los estados."""
    return "".join(chip_estado(e) for e in ESTADOS)

def tarjeta_estado(estado, n):
    """Tarjeta HTML de un estado con su conteo (semáforo, legible sobre blanco)."""
    c, t = _color_de_estado(estado)
    return (f'<div style="display:inline-block;min-width:132px;margin:0 8px 8px 0;padding:8px 10px;'
            f'border-radius:10px;background:{c}1f;border:1px solid {c}66;text-align:center;'
            f'vertical-align:top">'
            f'<div style="font-size:1.7rem;font-weight:700;line-height:1.1;color:{t}">{n}</div>'
            f'<div style="font-size:0.78rem;color:#374151;margin-top:2px">{estado}</div></div>')

def reporte_semaforo(df):
    """HTML con el conteo por cada estado del flujo + total (suma = filas del historial)."""
    tarjetas = "".join(tarjeta_estado(e, int((df["estado"] == e).sum())) for e in ESTADOS)
    total = ('<div style="display:inline-block;min-width:132px;margin:0 8px 8px 0;padding:8px 10px;'
             'border-radius:10px;background:#eef2f7;border:1px solid #cbd5e1;text-align:center;'
             'vertical-align:top"><div style="font-size:1.7rem;font-weight:700;line-height:1.1;'
             f'color:#1e293b">{len(df)}</div>'
             '<div style="font-size:0.78rem;color:#334155;margin-top:2px">Total</div></div>')
    return total + tarjetas

def fondo_estado(val):
    """Color de fondo semitransparente para la columna Estado en la tabla."""
    c, _ = _color_de_estado(val)
    return f"background-color:{c}33"

def migrar_estados_legacy():
    """Migra estados antiguos ya retirados del flujo:
    REVISAR → PENDIENTE APROBACIÓN, NOTIFICADO → APROBADO, PENDIENTE → RECIBIDO."""
    con = A.init_db()
    con.execute("UPDATE historial SET estado='PENDIENTE APROBACIÓN' WHERE estado IN ('REVISAR','PENDIENTE')")
    con.execute("UPDATE historial SET estado='APROBADO' WHERE estado='NOTIFICADO'")
    con.commit()
    con.close()

def auto_sync_publico():
    """Publica el reporte en la web automáticamente.
    Funciona tanto en local como en Streamlit Cloud.
    En Cloud, usa el token de GitHub de st.secrets."""
    import shutil
    directorio = os.path.dirname(os.path.abspath(__file__))
    marca = os.path.join(directorio, ".ultimo_sync_publico")
    repo_local = os.path.join(directorio, "publico")
    try:
        ultimo = 0.0
        if os.path.exists(marca):
            try:
                ultimo = float(open(marca, encoding="utf-8").read().strip() or 0)
            except Exception:
                ultimo = 0.0
        if time.time() - ultimo < 20:
            return

        # Obtener token de GitHub (local o cloud)
        token = None
        try:
            token = st.secrets.get("GITHUB_TOKEN")
        except Exception:
            pass
        if not token:
            token = os.environ.get("GITHUB_TOKEN")

        # Si no existe publico/, clonar desde GitHub
        if not os.path.exists(repo_local):
            if not token:
                return
            url = f"https://x-access-token:{token}@github.com/dannrob23/Solicitud-de-permisos-BAC.git"
            subprocess.run(["git", "clone", url, repo_local],
                          capture_output=True, text=True, env=dict(os.environ, GIT_TERMINAL_PROMPT="0"))

        SINC.exportar(verbose=False) if os.path.exists(os.path.join(directorio, "base.db")) else None

        # Asegurar que base.db esté en publico/
        base_src = os.path.join(directorio, "base.db")
        base_dst = os.path.join(repo_local, "base.db")
        if os.path.exists(base_src):
            shutil.copy2(base_src, base_dst)

        env = dict(os.environ, GIT_TERMINAL_PROMPT="0")
        if token:
            env["GIT_ASKPASS"] = "echo"
        r = subprocess.run(["git", "-C", repo_local, "status", "--porcelain"],
                           capture_output=True, text=True, env=env)
        if r.returncode != 0 or not r.stdout.strip():
            return
        subprocess.run(["git", "-C", repo_local, "add", "-A"], capture_output=True, text=True, env=env)
        subprocess.run(["git", "-C", repo_local, "commit", "-m", "Actualizacion automatica del reporte"],
                       capture_output=True, text=True, env=env)
        # Configurar remote con token si existe
        if token:
            subprocess.run(["git", "-C", repo_local, "remote", "set-url", "origin",
                            f"https://x-access-token:{token}@github.com/dannrob23/Solicitud-de-permisos-BAC.git"],
                           capture_output=True, text=True, env=env)
        p = subprocess.run(["git", "-C", repo_local, "push", "origin", "main"],
                           capture_output=True, text=True, env=env)
        if p.returncode != 0:
            subprocess.run(["git", "-C", repo_local, "pull", "--rebase", "origin", "main"],
                           capture_output=True, text=True, env=env)
            p = subprocess.run(["git", "-C", repo_local, "push", "origin", "main"],
                               capture_output=True, text=True, env=env)
        with open(marca, "w", encoding="utf-8") as f:
            f.write(str(time.time()))
        if p.returncode == 0:
            st.toast("Reporte web actualizándose automáticamente...")
        else:
            st.toast(f"Sync error: {p.stderr[:100]}")
    except Exception:
        pass

def consolidar_duplicados():
    """Borra registros duplicados de una misma solicitud (mismo caso + oficina),
    conservando el más reciente. Devuelve cuántos se eliminaron."""
    con = A.init_db()
    df = pd.read_sql_query("SELECT * FROM historial ORDER BY id", con)
    if df.empty:
        con.close()
        return 0
    def clave(r):
        caso = str(r.get("caso") or "").strip()
        if caso:
            return f"CASO|{caso}|{r.get('oficina')}"
        return f"MAN|{r.get('oficina')}|{r.get('solicitante')}|{r.get('sban')}|{r.get('fecha_solicitud')}"
    df["_clave"] = [clave(r) for _, r in df.iterrows()]
    conservar = df.groupby("_clave")["id"].max()
    borrar_ids = df.loc[~df["id"].isin(conservar), "id"].tolist()
    if borrar_ids:
        con.executemany("DELETE FROM historial WHERE id=?", [(i,) for i in borrar_ids])
        con.commit()
    con.close()
    return len(borrar_ids)

def badge(estado):
    cls = {"RECIBIDO": "b-pend", "PENDIENTE APROBACIÓN": "b-pend",
           "ENVIADO": "b-env", "APROBADO": "b-not", "RECHAZADO": "b-rev",
           "NOTIFICADO": "b-not", "REVISAR": "b-rev"}.get(estado, "b-pend")
    return f'<span class="badge {cls}">{estado}</span>'

# ==================== Bandeja de Descargas ====================
CARPETA_DEFAULT = os.path.join(os.path.expanduser("~"), "Downloads")

_TERM_INCLUIR = ["solicitud", "ingreso", "autorizaci", "aprobaci", "permiso", "re_"]
_TERM_EXCLUIR = ["presentaci", "entrevista", "hoja de vida", "hdv", "retiro",
                 "vacaciones", "incapacidad", "descanso"]

def es_correo_ingreso(nombre):
    """True si el nombre del .eml corresponde a una solicitud de ingreso o su respuesta."""
    n = (nombre or "").lower()
    if not n.endswith(".eml"):
        return False
    if any(t in n for t in _TERM_EXCLUIR):
        return False
    return any(t in n for t in _TERM_INCLUIR)

def _norm_caso(c):
    c = re.sub(r"\s+", "", str(c or "")).upper().replace("_", "")
    return "" if c in ("", "NA", "N/A", "NONE") else c

def _norm_oficina(o):
    return re.sub(r"\s+", " ", str(o or "")).strip()

def _normalizar_texto(t):
    return re.sub(r"\s+", " ", str(t or "")).strip()

def _buscar_registro(rows, caso, oficina):
    """Busca en `rows` (tuplas id,caso,oficina,estado) la solicitud que corresponde
    al correo. Con caso real solo fusiona por caso (nunca por oficina); sin caso
    (N/A o vacío) fusiona únicamente con solicitudes sin caso de la misma oficina."""
    if caso:
        match = [r for r in rows if _norm_caso(r[1]) == caso]
        if not match:
            return None
        exact = [r for r in match if _norm_oficina(r[2]).upper() == (oficina or "").upper()]
        return (exact or match)[-1]
    if oficina:
        match = [r for r in rows
                 if not _norm_caso(r[1]) and _norm_oficina(r[2]).upper() == oficina.upper()]
        return match[-1] if match else None
    return None

def _marcas_procesadas():
    con = A.init_db()
    con.execute("""CREATE TABLE IF NOT EXISTS descargas (
        archivo TEXT NOT NULL, mtime REAL NOT NULL, id_historial INTEGER,
        PRIMARY KEY (archivo, mtime))""")
    con.commit()
    marcas = {}
    for archivo, mtime, rid in con.execute("SELECT archivo, mtime, id_historial FROM descargas"):
        marcas[(archivo, mtime)] = rid
    con.close()
    return marcas

def _guardar_marca(archivo, mtime, rid):
    con = A.init_db()
    con.execute("""INSERT OR REPLACE INTO descargas (archivo, mtime, id_historial)
                   VALUES (?,?,?)""", (archivo, mtime, rid))
    con.commit()
    con.close()

def _registrar_archivo(d):
    """Registra la solicitud del correo en el historial (como al procesarla)."""
    aprobador = A.resolver_aprobador(d.get("sban"))
    estado = "PENDIENTE APROBACIÓN"
    notas = ""
    if not aprobador:
        notas = f"SBAN {d['sban'] or 'desconocido'} sin aprobador en el mapeo"
    con = A.init_db()
    rid = A.registrar(con, d, aprobador, estado, notas)
    con.close()
    return rid

def escanear_descargas(carpeta, forzar=False):
    """Lista los .eml de solicitudes de ingreso de una carpeta y los compara con el
    historial. Los que no estén registrados se registran automáticamente.
    El resultado se cachea y solo se recalcula cuando cambian los archivos."""
    lista = sorted(glob.glob(os.path.join(carpeta, "*.eml")),
                   key=os.path.getmtime, reverse=True)
    sig = tuple((os.path.basename(p), round(os.path.getmtime(p), 1))
                for p in lista if es_correo_ingreso(os.path.basename(p)))
    if not forzar and st.session_state.get("desc_sig") == sig and st.session_state.get("desc_filas") is not None:
        return st.session_state["desc_filas"]

    con = A.init_db()
    hist = con.execute("SELECT id, caso, oficina, estado FROM historial").fetchall()
    con.close()
    marcas = _marcas_procesadas()

    filas, nuevos = [], 0
    for p in lista:
        nombre = os.path.basename(p)
        if not es_correo_ingreso(nombre):
            continue
        mtime = os.path.getmtime(p)
        f = {"archivo": nombre, "llegada": mtime, "caso": "", "oficina": "",
             "id": None, "estado": None, "nuevo": False, "nota": ""}
        try:
            datos = A.leer_eml(p)
            d = A.extraer_datos(datos)
            caso = _norm_caso(d.get("caso"))
            oficina = _norm_oficina(d.get("oficina") or "")
            if not oficina and d.get("asunto_original"):
                m = re.search(r"(?:Solicitud Ingreso|ingreso)\s*[-–]\s*([0-9A-Za-zÁÉÍÓÚÑáéíóúñü -]+)", d["asunto_original"], re.I)
                if m:
                    oficina = _norm_oficina(m.group(1))
            f["caso"] = (d.get("caso") or "").strip()
            f["oficina"] = oficina
            reg = _buscar_registro(hist, caso, oficina)
            if reg is not None:
                f["id"], f["estado"] = reg[0], reg[3]
            elif (nombre, mtime) in marcas:
                f["id"] = marcas[(nombre, mtime)]
                for r in hist:
                    if r[0] == f["id"]:
                        f["estado"] = r[3]
                        break
            elif caso or oficina:
                rid = _registrar_archivo(d)
                _guardar_marca(nombre, mtime, rid)
                hist.append((rid, d.get("caso") or "", oficina, "PENDIENTE APROBACIÓN"))
                f["id"], f["estado"], f["nuevo"] = rid, "PENDIENTE APROBACIÓN", True
                nuevos += 1
            else:
                f["nota"] = "Sin caso ni oficina detectados: revísalo manualmente"
        except Exception as e:
            f["nota"] = f"Error al leer: {str(e)[:100]}"
        filas.append(f)

    st.session_state["desc_sig"] = sig
    st.session_state["desc_filas"] = filas
    st.session_state["desc_nuevos"] = nuevos
    return filas

# ---------- Sidebar ----------
# Publica automáticamente la base en la web (solo si cambió algo)
auto_sync_publico()

with st.sidebar:
    st.header("⚙️ Configuración")
    st.caption("Firma que se agrega al final del correo. Déjala vacía para usar la de Outlook.")
    firma_actual = A.get_firma()
    firma_es_marcador = "[AQUÍ VA TU FIRMA]" in firma_actual
    nueva_firma = st.text_area("Firma (HTML o texto)",
                               value="" if firma_es_marcador else firma_actual,
                               height=100,
                               help="Ej: Daniel Robles | Gestor Documental | tel | correo")
    if st.button("💾 Guardar firma"):
        with open(A.FIRMA, "w", encoding="utf-8") as f:
            f.write(nueva_firma.strip())
        st.success("Firma guardada")
    if st.button("🔄 Regenerar mapeo (oficinas)"):
        import subprocess as sp
        r = sp.run([sys.executable, os.path.join(A.BASE, "generar_mapeo.py")],
                   capture_output=True, text=True)
        st.info(r.stdout[-400:] if r.stdout else r.stderr[-400:])

# ---------- Carga del correo ----------
tab1, tab_aprob, tab_desc, tab_hist = st.tabs([
    "📥 Procesar solicitud (PDF / EML)",
    "👥 Aprobadores y Vacaciones",
    "📂 Descargas",
    "📋 Historial"
])

with tab1:
    st.markdown("#### 📥 Elige cómo ingresar la solicitud")
    tipo_entrada = st.radio("Fuente de la solicitud:", [
        "📄 Archivo PDF o EML (Recomendado)",
        "📝 Texto plano (pegar)",
        "📸 Captura de pantalla (imagen)"
    ], horizontal=True)

    ruta = None
    texto_plano = None
    bytes_archivo = None
    es_pdf = False

    if tipo_entrada.startswith("📄"):
        c1, c2 = st.columns([2, 1])
        with c1:
            archivo_subido = st.file_uploader(
                "Sube el PDF (impreso desde Outlook Móvil) o archivo .EML (Outlook Web):",
                type=["pdf", "eml"]
            )
        with c2:
            ruta_manual = st.text_input("O escribe la ruta local:", placeholder="C:\\...\\solicitud.pdf")
        if archivo_subido is not None:
            bytes_archivo = archivo_subido.getvalue()
            es_pdf = archivo_subido.name.lower().endswith(".pdf")
            if not es_pdf:
                tmp = os.path.join(tempfile.gettempdir(), "asistente_upload.eml")
                with open(tmp, "wb") as f:
                    f.write(bytes_archivo)
                ruta = tmp
        elif ruta_manual.strip():
            ruta = ruta_manual.strip()
            es_pdf = ruta.lower().endswith(".pdf")

    elif tipo_entrada.startswith("📝"):
        texto_plano = st.text_area("Pega aquí el contenido del correo (los datos de la solicitud):",
                                   height=280,
                                   placeholder="Buen día,\n\nSolicito su colaboración y gestión...\n\nFecha de Ingreso...\nCOLSOF\tFuncionario\tDocumento...\n\nSe hace la solicitud de ingreso a la siguiente Oficina: 4519-LA MACARENA\nCaso: IM3236508\nDescripción: ...")

    else:
        img = st.file_uploader("Arrastra aquí la captura de pantalla del correo (PNG/JPG):",
                               type=["png", "jpg", "jpeg"])
        if img is not None:
            tmp = os.path.join(tempfile.gettempdir(), "asistente_captura.png")
            with open(tmp, "wb") as f:
                f.write(img.getbuffer())
            ruta = tmp

    if bytes_archivo is not None or (ruta and os.path.exists(ruta)) or (texto_plano and texto_plano.strip()):
        try:
            if bytes_archivo is not None and es_pdf:
                res = A.procesar_pdf(bytes_archivo)
            elif ruta and es_pdf:
                res = A.procesar_pdf(ruta)
            elif texto_plano and texto_plano.strip():
                res = A.procesar_texto(texto_plano)
            elif ruta and ruta.lower().endswith((".png", ".jpg", ".jpeg")):
                res = A.procesar_imagen(ruta)
            else:
                res = A.procesar(ruta)
            d, apb = res["datos"], res["aprobador"]

            st.markdown("### 📬 Resultado del análisis")
            if apb:
                aviso_reemp = ""
                if apb.get("en_reemplazo"):
                    mot = (apb.get("motivo") or "vacaciones").upper()
                    aviso_reemp = (f"<div style='margin-top:8px;padding:6px 10px;border-radius:6px;"
                                   f"background:#fef3c7;color:#92400e;font-size:0.88rem;font-weight:600;'>"
                                   f"🏖️ En Reemplazo Temporal por {mot} de: {apb.get('titular')}</div>")
                st.markdown(f"""
                <div class="card">
                <b>▶ ENVIAR A:</b> {apb['ingeniero']} &nbsp; <code>{apb['correo']}</code><br>
                <b>Regional:</b> {apb['regional']} &nbsp;|&nbsp; <b>Oficina:</b> {apb['oficina']} (SBAN {d['sban']})
                {aviso_reemp}
                </div>""", unsafe_allow_html=True)
            else:
                st.error(f"⚠️ **SIN APROBADOR:** {res['notas']}. Revisa manualmente.")

            cols = st.columns(4)
            with cols[0]: st.metric("Caso", d["caso"])
            with cols[1]: st.metric("Fecha ingreso", d["fecha_ingreso"])
            with cols[2]: st.metric("Funcionario", d["funcionario"] or "—")
            with cols[3]: st.metric("Empresa", d["empresa"] or "—")

            # Solicitante: visible y editable
            sol_visible = st.text_input("👤 Solicitado por (quién pidió el permiso):",
                                        value=d.get("solicitante", ""))
            if sol_visible.strip() and sol_visible.strip() != d.get("solicitante", ""):
                d["solicitante"] = sol_visible.strip()
                con_upd = A.init_db()
                con_upd.execute("UPDATE historial SET solicitante=? WHERE id=?", (d["solicitante"], res["id"]))
                con_upd.commit(); con_upd.close()
                res["datos"]["solicitante"] = d["solicitante"]

            with st.expander("✏️ Revisar / corregir datos (útil si algún campo necesita ajuste)", expanded=False):
                st.caption("Corrige cualquier campo y pulsa el botón para regenerar el correo.")
                f_empresa = st.text_input("Empresa", value=d["empresa"])
                f_funcionario = st.text_input("Funcionario", value=d["funcionario"])
                f_doc = st.text_input("Documento", value=d["documento"])
                f_eps = st.text_input("EPS", value=d["eps"])
                f_arl = st.text_input("ARL", value=d["arl"])
                f_vence = st.text_input("Vence", value=d["arl_vence"])
                f_cel = st.text_input("Celular", value=d["celular"])
                f_fecha = st.text_input("Fecha de ingreso", value=d["fecha_ingreso"])
                f_oficina = st.text_input("Oficina", value=d["oficina"])
                f_caso = st.text_input("Caso", value=d["caso"])
                f_solicitante = st.text_input("Solicitado por (quién pidió el permiso)", value=d.get("solicitante", ""))
                f_desc = st.text_area("Descripción", value=d["descripcion"], height=70)
                if st.button("🔄 Regenerar correo con datos corregidos"):
                    d2 = dict(d)
                    d2.update({"empresa": f_empresa, "funcionario": f_funcionario,
                               "documento": f_doc, "eps": f_eps, "arl": f_arl,
                               "arl_vence": f_vence, "celular": f_cel,
                               "fecha_ingreso": f_fecha, "oficina": f_oficina,
                               "caso": f_caso, "solicitante": f_solicitante,
                               "descripcion": f_desc})
                    res["datos"] = d2
                    res["html"] = A.generar_html(d2, res["aprobador"], foto_b64=res.get("foto_b64"))
                    st.success("Correo regenerado con los datos corregidos ⬇️")

            st.markdown("#### 📋 Acciones del correo")
            render_copiar_cliente(res["html"])

            cA, cB = st.columns(2)
            with cA:
                if apb and apb.get("correo"):
                    mailto_url = f"mailto:{apb['correo']}?subject={urllib.parse.quote(res['asunto'])}"
                    st.markdown(f'<a href="{mailto_url}" target="_blank" style="display:block;text-align:center;background:#f8fafc;color:#1e293b;border:1px solid #cbd5e1;padding:10px;border-radius:8px;font-weight:600;text-decoration:none;font-size:0.95rem;">✉️ Abrir correo nuevo en Outlook App</a>', unsafe_allow_html=True)
                st.download_button("📥 Descargar correo HTML", res["html"], file_name="correo_solicitud.html", mime="text/html", use_container_width=True)
            with cB:
                estado_actual = res["estado"]
                idx = ESTADOS.index(estado_actual) if estado_actual in ESTADOS else 1
                nuevo_estado = st.selectbox("Estado", ESTADOS, index=idx)
                if st.button("💾 Actualizar estado", use_container_width=True):
                    marcar_estado(res["id"], nuevo_estado)
                    st.success(f"Estado → {nuevo_estado}")
                    st.rerun()

            st.markdown("### 📤 Envío directo (SMTP opcional)")
            st.caption("Si tienes SMTP configurado, puedes despacharlo directamente.")
            c_dest = st.columns(2)
            with c_dest[0]:
                destino_def = apb["correo"] if apb else "danielcamilo.roblespinto@colsof.com.co"
                destino_envio = st.text_input("Enviar a:", value=destino_def)
            with c_dest[1]:
                st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
                if st.button("🚀 Enviar correo ahora (SMTP)", use_container_width=True):
                    try:
                        A.enviar_correo(res["html"], destino_envio.strip(), res["asunto"])
                        st.success(f"✅ Correo enviado a {destino_envio.strip()}")
                        marcar_estado(res["id"], "ENVIADO")
                    except Exception as e:
                        st.error(f"❌ No se pudo enviar: {str(e)[:250]}")
                        st.info("Configura SMTP con `configurar_smtp.py` o copia el correo arriba.")

            st.markdown("#### 👁️ Vista previa del correo generado")
            st.components.v1.html(res["html"], height=900, scrolling=True)

            if res.get("foto_b64"):
                st.caption("📸 Foto del funcionario extraída:")
                st.image(base64.b64decode(res["foto_b64"]), width=120)

        except Exception as e:
            st.error(f"Error procesando el archivo: {e}")
    else:
        st.info("👈 Sube un archivo PDF o .eml para comenzar.")

# ==================== Pestaña: Aprobadores y Vacaciones ====================
with tab_aprob:
    st.markdown("### 👥 Directorio de Aprobadores y Novedades")
    st.caption("Gestiona los ingenieros aprobadores por Regional. Si alguien sale a vacaciones o licencia, activa su suplente con un toque sin tener que editar archivos ni markdown.")

    novedades = A.obtener_novedades_aprobadores()
    totales = len(novedades)
    en_vacaciones = sum(1 for n in novedades if n["activo"] == 1)

    col_m1, col_m2, col_m3 = st.columns(3)
    with col_m1:
        st.metric("Regionales", totales)
    with col_m2:
        st.metric("Titulares activos", totales - en_vacaciones)
    with col_m3:
        st.metric("En Reemplazo / Vacaciones", en_vacaciones)

    st.markdown("---")
    st.markdown("#### 📍 Estado actual por Regional")

    for n in novedades:
        reg = n["regional"]
        activo = (n["activo"] == 1)
        with st.container():
            c_info, c_action = st.columns([3, 1])
            with c_info:
                if activo:
                    estado_badge = "<span style='background:#fef3c7;color:#92400e;padding:3px 10px;border-radius:12px;font-weight:600;font-size:0.8rem;'>🟡 EN REEMPLAZO</span>"
                    st.markdown(f"**Regional {reg}** {estado_badge}", unsafe_allow_html=True)
                    st.markdown(f"**Aprobador Activo:** {n['ingeniero_reemplazo']} (`{n['correo_reemplazo']}`)")
                    motivo_txt = n['motivo'] or 'Vacaciones'
                    fechas_txt = f"({n['fecha_inicio']} a {n['fecha_fin']})" if (n['fecha_inicio'] or n['fecha_fin']) else ""
                    st.caption(f"Motivo: {motivo_txt} {fechas_txt} | Titular habitual: {n['ingeniero_titular']}")
                else:
                    estado_badge = "<span style='background:#dcfce7;color:#166534;padding:3px 10px;border-radius:12px;font-weight:600;font-size:0.8rem;'>🟢 TITULAR ACTIVO</span>"
                    st.markdown(f"**Regional {reg}** {estado_badge}", unsafe_allow_html=True)
                    st.markdown(f"**Ingeniero Titular:** {n['ingeniero_titular']} (`{n['correo_titular']}`)")
            with c_action:
                if activo:
                    if st.button("🔄 Reactivar Titular", key=f"btn_react_{reg}", use_container_width=True):
                        A.toggle_novedad_aprobador(reg, activo=False)
                        st.toast(f"✅ Titular de {reg} reactivado")
                        st.rerun()
                else:
                    st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)
                    st.caption("Normal")
            st.markdown("<hr style='margin:6px 0;border:0;border-top:1px solid #f1f5f9;'>", unsafe_allow_html=True)

    with st.expander("⚙️ Programar Vacaciones, Reemplazo o Cambiar Titular", expanded=False):
        regiones_nombres = [n["regional"] for n in novedades]
        reg_sel = st.selectbox("Selecciona la Regional:", regiones_nombres)
        info_reg = next((n for n in novedades if n["regional"] == reg_sel), None)

        tipo_cambio = st.radio("Tipo de cambio:", [
            "🏖️ Reemplazo Temporal (Vacaciones / Licencia / Incapacidad)",
            "👔 Cambio de Titular Definitivo (Retiro / Nuevo Ingreso)"
        ])

        if tipo_cambio.startswith("🏖️"):
            st.info(f"Titular actual: **{info_reg['ingeniero_titular']}** ({info_reg['correo_titular']})")
            c_nom, c_cor = st.columns(2)
            with c_nom:
                nom_suplente = st.text_input("Nombre del Ingeniero Suplente:", value=info_reg.get("ingeniero_reemplazo", "") or "")
            with c_cor:
                cor_suplente = st.text_input("Correo del Ingeniero Suplente:", value=info_reg.get("correo_reemplazo", "") or "")

            c_mot, c_f1, c_f2 = st.columns(3)
            with c_mot:
                motivos_lista = ["Vacaciones", "Incapacidad", "Licencia", "Comisión", "Otro"]
                motivo_idx = motivos_lista.index(info_reg["motivo"]) if info_reg.get("motivo") in motivos_lista else 0
                motivo_sel = st.selectbox("Motivo:", motivos_lista, index=motivo_idx)
            with c_f1:
                f_ini = st.text_input("Fecha Inicio (opcional):", value=info_reg.get("fecha_inicio", "") or "", placeholder="YYYY-MM-DD")
            with c_f2:
                f_fin = st.text_input("Fecha Fin (opcional):", value=info_reg.get("fecha_fin", "") or "", placeholder="YYYY-MM-DD")

            activar_ahora = st.checkbox("Activar reemplazo inmediatamente", value=True)

            if st.button("💾 Guardar Novedad Temporal", type="primary", use_container_width=True):
                if not nom_suplente.strip() or not cor_suplente.strip():
                    st.error("Por favor completa el nombre y correo del suplente.")
                else:
                    A.guardar_novedad_aprobador(
                        regional=reg_sel,
                        titular=info_reg["ingeniero_titular"],
                        correo_titular=info_reg["correo_titular"],
                        reemplazo=nom_suplente.strip(),
                        correo_reemplazo=cor_suplente.strip(),
                        motivo=motivo_sel,
                        activo=1 if activar_ahora else 0,
                        fecha_inicio=f_ini.strip(),
                        fecha_fin=f_fin.strip()
                    )
                    st.success(f"Novedad guardada para regional {reg_sel}")
                    st.rerun()
        else:
            st.warning("⚠️ Esta acción actualizará permanentemente al titular de la regional.")
            nuevo_titular = st.text_input("Nuevo Nombre del Titular:", value=info_reg["ingeniero_titular"])
            nuevo_correo = st.text_input("Nuevo Correo del Titular:", value=info_reg["correo_titular"])
            if st.button("💾 Actualizar Titular Permanente", type="primary", use_container_width=True):
                if not nuevo_titular.strip() or not nuevo_correo.strip():
                    st.error("Completa el nombre y correo del nuevo titular.")
                else:
                    A.actualizar_titular_definitivo(reg_sel, nuevo_titular.strip(), nuevo_correo.strip())
                    st.success(f"Titular de {reg_sel} actualizado correctamente.")
                    st.rerun()

# ==================== Bandeja de Descargas ====================
with tab_desc:
    migrar_estados_legacy()
    st.markdown("### 📂 Bandeja de Descargas")
    st.caption("Lee los correos .eml de la carpeta. Los que aún no están en el historial se **registran automáticamente** (Pendiente aprobación) y puedes cambiarles el estado con el desplegable de la tabla. "
               "Nota: varios correos (solicitud + respuesta RE) pueden referirse a la MISMA solicitud; el Historial cuenta solicitudes, no correos.")

    carpeta_desc = st.text_input("Carpeta con los correos .eml:", value=CARPETA_DEFAULT)
    revisar = st.button("🔄 Revisar carpeta de nuevo", use_container_width=False)

    if not os.path.isdir(carpeta_desc):
        st.error(f"No existe la carpeta: {carpeta_desc}")
    else:
        filas = escanear_descargas(carpeta_desc, forzar=bool(revisar))
        nuevos = st.session_state.get("desc_nuevos", 0)

        if not filas:
            st.info("No hay correos .eml de solicitud de ingreso en esa carpeta.")
        else:
            # Refresca el estado real de cada registro desde la base
            con = A.init_db()
            estado_real = dict(con.execute("SELECT id, estado FROM historial"))
            con.close()
            for f in filas:
                if f["id"]:
                    f["estado"] = estado_real.get(f["id"]) or f["estado"]
                    if f["estado"] not in ESTADOS:
                        f["estado"] = "RECIBIDO"

            sin_reg = [f for f in filas if not f["id"]]
            pend = sum(1 for f in filas if f["id"] and f["estado"] in ("RECIBIDO", "PENDIENTE APROBACIÓN"))

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("📄 Correos en la carpeta", len(filas))
            m2.metric("🗂️ En el historial", sum(1 for f in filas if f["id"]))
            m3.metric("🆕 Registrados automáticamente", nuevos)
            m4.metric("⏳ Por gestionar", pend)
            if nuevos:
                st.success(f"Se registraron automáticamente {nuevos} correo(s) nuevo(s). Ya aparecen en el Historial.")
            if sin_reg:
                with st.expander("⚠️ Correos para revisión manual (sin caso u oficina legibles)", expanded=False):
                    for f in sin_reg:
                        nombre = f["archivo"].replace("[EXTERNAL]", "")
                        if f["nota"]:
                            st.markdown(f"- **{nombre}** — {f['nota']}")
                        else:
                            st.markdown(f"- **{nombre}** — arrástralo en la pestaña **📥 Procesar correo** para revisar sus datos.")

            st.markdown("**Semáforo de estados:** " + semaforo_leyenda(), unsafe_allow_html=True)

            st.markdown("#### ⚙️ Filtros rápidos")
            fd1, fd2 = st.columns([2, 3])
            with fd1:
                f_desc_estado = st.multiselect("Estado", ESTADOS, key="filtro_estado_descargas")
            with fd2:
                f_desc_buscar = st.text_input("Buscar por oficina, caso o nombre del correo:")

            filas_utiles = [f for f in filas if f["id"]]
            if not filas_utiles:
                st.info("Ningún correo de la carpeta está registrado para poder cambiar su estado.")
            else:
                base_desc = pd.DataFrame([{
                    "Correo": f["archivo"].replace("[EXTERNAL]", ""),
                    "Llegó": datetime.fromtimestamp(f["llegada"]).strftime("%d/%m/%Y %H:%M"),
                    "Caso": f["caso"] or "—",
                    "Oficina": f["oficina"] or "—",
                    "Estado": f["estado"],
                } for f in filas_utiles])
                ids_desc = [f["id"] for f in filas_utiles]

                mask_desc = pd.Series(True, index=base_desc.index)
                if f_desc_estado:
                    mask_desc &= base_desc["Estado"].isin(f_desc_estado)
                if f_desc_buscar.strip():
                    qd = f_desc_buscar.strip()
                    mask_desc &= (base_desc["Oficina"].astype(str).str.contains(qd, case=False, na=False) |
                                  base_desc["Caso"].astype(str).str.contains(qd, case=False, na=False) |
                                  base_desc["Correo"].astype(str).str.contains(qd, case=False, na=False))

                vis_desc = base_desc[mask_desc].reset_index(drop=True)
                if vis_desc.empty:
                    st.info("Ningún correo coincide con los filtros.")
                else:
                    ids_vis = [ids_desc[i] for i in base_desc.index[mask_desc]]
                    st.caption("Cambia el estado en la celda y pulsa **💾 Guardar cambios**.")
                    st.caption("Un mismo correo original y su respuesta (RE) comparten la misma solicitud: al cambiar uno cambia la solicitud.")
                    editor_desc_key = f"editor_desc_{len(vis_desc)}_{'-'.join(f_desc_estado)}_{f_desc_buscar}_{st.session_state.get('desc_rev', 0)}"
                    edit_desc = st.data_editor(
                        vis_desc,
                        key=editor_desc_key,
                        use_container_width=True,
                        height=420,
                        hide_index=True,
                        disabled=["Correo", "Llegó", "Caso", "Oficina"],
                        column_config={
                            "Correo": st.column_config.TextColumn("Correo", width="large"),
                            "Llegó": st.column_config.TextColumn("Llegó", width="small"),
                            "Caso": st.column_config.TextColumn("Caso", width="small"),
                            "Oficina": st.column_config.TextColumn("Oficina", width="medium"),
                            "Estado": st.column_config.SelectboxColumn("Estado", options=ESTADOS,
                                                                       required=True, width="medium"),
                        },
                    )
                    if st.button("💾 Guardar cambios de estado", type="primary", use_container_width=True, key="btn_guardar_descargas"):
                        camb = edit_desc[edit_desc["Estado"] != vis_desc["Estado"]]
                        if camb.empty:
                            st.info("No hay cambios de estado por guardar.")
                        else:
                            for pos in camb.index:
                                marcar_estado(ids_vis[pos], str(camb.at[pos, "Estado"]))
                            st.session_state["desc_rev"] = st.session_state.get("desc_rev", 0) + 1
                            st.success(f"✔ Estado actualizado en {len(camb)} solicitud(es).")
                            st.rerun()

with tab_hist:
    migrar_estados_legacy()
    st.markdown("### 📋 Historial de solicitudes")

    # Limpieza de duplicados: si una solicitud se procesó varias veces,
    # se conserva el registro más reciente (con su último estado).
    consolidados = consolidar_duplicados()

    # --- Toggle: Simplificar vista ---
    col_toggle, _ = st.columns([2, 1])
    with col_toggle:
        simplificado = st.toggle("🔍 Simplificar vista (oculta aprobados >2 días)", value=True,
                                 help="Muestra solo solicitudes del día y aprobados recientes. "
                                      "Los aprobados con más de 2 días se ocultan pero siguen "
                                      "disponibles en el histórico descendente o en la descarga.")
    st.markdown("---")

    # --- Contadores por estado (semáforo) ---
    df = leer_historial(filtrar_aprobados=simplificado)
    if df.empty:
        st.info("Aún no hay registros.")
    else:
        if consolidados:
            st.info(f"🧹 Se consolidaron {consolidados} registro(s) duplicado(s) de la misma solicitud.")
        # Normaliza estados legados ("PENDIENTE" antiguo → RECIBIDO)
        df["estado"] = df["estado"].apply(lambda e: e if e in ESTADOS else "RECIBIDO")

        st.markdown("**Estado del flujo**")
        st.markdown(reporte_semaforo(df), unsafe_allow_html=True)
        st.caption("Cada solicitud tiene UN solo estado y la suma de tarjetas siempre es igual al Total de filas del historial. "
                   "Ver 0 en Recibidas/Enviadas/Rechazadas significa que ninguna solicitud está marcada así todavía.")

        # --- Exportar reporte ---
        exp = df[["id", "caso", "oficina", "solicitante", "fecha_solicitud",
                  "regional", "aprobador", "estado", "fecha_proceso"]].copy()
        exp.columns = ["ID", "Caso", "Solicitud", "Solicitado por", "Solicitado el",
                       "Regional", "Aprobador", "Estado", "Registrado el"]
        csv_bytes = exp.to_csv(index=False).encode("utf-8-sig")
        excel_buf = io.BytesIO()
        with pd.ExcelWriter(excel_buf, engine="openpyxl") as xw:
            exp.to_excel(xw, index=False, sheet_name="Historial")
        xe1, xe2 = st.columns(2)
        with xe1:
            st.download_button("⬇️ Descargar reporte (CSV)",
                               data=csv_bytes,
                               file_name="historial_solicitudes.csv",
                               mime="text/csv", key="dl_hist_csv", use_container_width=True)
        with xe2:
            st.download_button("⬇️ Descargar reporte (Excel)",
                               data=excel_buf.getvalue(),
                               file_name="historial_solicitudes.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               key="dl_hist_xlsx", use_container_width=True)


        st.markdown("#### 🔍 Filtros")
        f1, f2 = st.columns([2, 3])
        with f1:
            f_estado = st.multiselect("Estado", ESTADOS, key="filtro_estado_historial")
        with f2:
            f_buscar = st.text_input("Buscar por oficina, caso, solicitante o aprobador:")

        st.markdown("#### 📄 Tabla de solicitudes")
        st.markdown("**Semáforo de estados:** " + semaforo_leyenda(), unsafe_allow_html=True)

        modo_tabla = st.radio("Ver la tabla como:", ["🎨 Vista con colores", "✏️ Editar estados"],
                              horizontal=True, key="modo_tabla")

        base = df[["id", "caso", "oficina", "solicitante", "fecha_solicitud",
                   "regional", "aprobador", "estado"]].copy().reset_index(drop=True)
        for col in ["caso", "oficina", "solicitante", "fecha_solicitud", "regional", "aprobador"]:
            base[col] = base[col].fillna("").astype(str)
        base["estado"] = base["estado"].where(base["estado"].isin(ESTADOS), "RECIBIDO")

        display = pd.DataFrame({
            "Solicitud": base["oficina"],
            "Caso": base["caso"],
            "Solicitado por": base["solicitante"],
            "Solicitado el": base["fecha_solicitud"],
            "Regional": base["regional"],
            "Aprobador": base["aprobador"],
            "Estado": base["estado"],
        })

        # Aplica los filtros a la vista de la tabla
        mask = pd.Series(True, index=display.index)
        if f_estado:
            mask &= display["Estado"].isin(f_estado)
        if f_buscar.strip():
            q = f_buscar.strip()
            mask &= (display["Solicitud"].astype(str).str.contains(q, case=False, na=False) |
                     display["Caso"].astype(str).str.contains(q, case=False, na=False) |
                     display["Solicitado por"].astype(str).str.contains(q, case=False, na=False) |
                     display["Aprobador"].astype(str).str.contains(q, case=False, na=False))

        idx_vis = display.index[mask].tolist()
        if not idx_vis:
            st.info("Ninguna solicitud coincide con los filtros.")
        else:
            vis = display.loc[idx_vis].reset_index(drop=True)

            if modo_tabla == "🎨 Vista con colores":
                # Semáforo: cada fila se colorea según el estado
                styled = vis.style.map(fondo_estado, subset=["Estado"])
                st.dataframe(styled, use_container_width=True, height=420)
                st.caption("Gris = Recibido · Amarillo = Pendiente aprobación · Azul = Enviado · Verde = Aprobado · Rojo = Rechazado")
            else:
                st.caption("Cambia el valor de **Estado** en la celda y pulsa **💾 Guardar cambios de estado**.")
                pos_a_idx = {pos: idx_vis[pos] for pos in range(len(idx_vis))}
                # key dependiente del contenido/filtros: evita que queden ediciones viejas en memoria
                editor_key = f"editor_historial_{len(vis)}_{'-'.join(f_estado)}_{f_buscar}_{st.session_state.get('hist_rev', 0)}"
                edited = st.data_editor(
                    vis,
                    key=editor_key,
                    use_container_width=True,
                    height=420,
                    hide_index=True,
                    disabled=["Solicitud", "Caso", "Solicitado por", "Solicitado el",
                              "Regional", "Aprobador"],
                    column_config={
                        "Solicitud": st.column_config.TextColumn("Solicitud", width="medium"),
                        "Caso": st.column_config.TextColumn("Caso", width="small"),
                        "Solicitado el": st.column_config.TextColumn("Solicitado el", width="small"),
                        "Aprobador": st.column_config.TextColumn("Aprobador", width="medium"),
                        "Estado": st.column_config.SelectboxColumn("Estado", options=ESTADOS,
                                                                   required=True, width="medium"),
                    },
                )
                if st.button("💾 Guardar cambios de estado", type="primary", use_container_width=True, key="btn_guardar_historial"):
                    cambiadas = edited[edited["Estado"] != vis["Estado"]]
                    if cambiadas.empty:
                        st.info("No hay cambios de estado por guardar.")
                    else:
                        for pos in cambiadas.index:
                            rid = int(base.at[pos_a_idx[pos], "id"])
                            marcar_estado(rid, str(cambiadas.at[pos, "Estado"]))
                        st.session_state["hist_rev"] = st.session_state.get("hist_rev", 0) + 1
                        st.success(f"✔ Estado actualizado en {len(cambiadas)} solicitud(es).")
                        st.rerun()


