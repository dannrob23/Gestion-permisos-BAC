#!/usr/bin/env python3
"""Reporte público (solo lectura) del Historial de Solicitudes de Ingreso.

Vista de consulta para compañeros: sin edición, sin subida de correos.
Ejecutar:  streamlit run reporte.py --server.headless true --server.port 8502
"""
import os, sqlite3, io
from datetime import datetime
from zoneinfo import ZoneInfo

import streamlit as st
import pandas as pd

BASE = os.path.dirname(os.path.abspath(__file__))
DB = os.path.join(BASE, "base.db")
ZONA_COLOMBIA = ZoneInfo("America/Bogota")

def ahora_colombia():
    """Fecha y hora actual en Colombia (UTC-5), sin importar el huso del servidor."""
    return datetime.now(ZONA_COLOMBIA).strftime("%d/%m/%Y %H:%M:%S")
st.set_page_config(page_title="Consulta Solicitudes de Ingreso", page_icon="📋", layout="wide")

st.markdown("""
<style>
    .stApp { background-color: #ffffff; color: #111827; }
</style>
""", unsafe_allow_html=True)

# ---------- Acceso opcional con contraseña ----------
# Si existe el secreto [consulta].password (en .streamlit/secrets.toml o en
# Streamlit Cloud -> Settings -> Secrets), el reporte pide esa contraseña.
def _acceso_permitido():
    try:
        pw = str(st.secrets["consulta"]["password"])
    except Exception:
        return True  # sin contraseña configurada -> acceso libre
    if st.session_state.get("consulta_ok"):
        return True
    st.markdown("### 🔐 Consulta protegida")
    st.text_input("Contraseña de consulta", type="password", key="pw_consulta")
    if st.button("Entrar"):
        if st.session_state["pw_consulta"] == pw:
            st.session_state["consulta_ok"] = True
            st.rerun()
        else:
            st.error("Contraseña incorrecta.")
    return False

if not _acceso_permitido():
    st.stop()

st.title("📋 Solicitudes de Ingreso — Consulta")
st.caption("Reporte de consulta (solo lectura).")

ESTADOS = ["PENDIENTE APROBACIÓN", "APROBADO", "RECHAZADO"]

COLOR_ESTADO = {
    "RECIBIDO": "#94a3b8",
    "PENDIENTE APROBACIÓN": "#fbbf24",
    "ENVIADO": "#60a5fa",
    "APROBADO": "#4ade80",
    "RECHAZADO": "#f87171",
}
COLOR_ESTADO_TEXTO = {
    "RECIBIDO": "#475569",
    "PENDIENTE APROBACIÓN": "#b45309",
    "ENVIADO": "#1d4ed8",
    "APROBADO": "#15803d",
    "RECHAZADO": "#b91c1c",
}

LEGACY = {"REVISAR": "PENDIENTE APROBACIÓN", "NOTIFICADO": "APROBADO",
          "PENDIENTE": "PENDIENTE APROBACIÓN", "RECIBIDO": "PENDIENTE APROBACIÓN",
          "ENVIADO": "PENDIENTE APROBACIÓN"}

def estado_norm(e):
    e = (e or "").strip().upper()
    return e if e in ESTADOS else LEGACY.get(e, "RECIBIDO")

def _colores(estado):
    c = COLOR_ESTADO.get(str(estado), "#94a3b8")
    t = COLOR_ESTADO_TEXTO.get(str(estado), "#334155")
    return c, t

def tarjeta(estado, n):
    c, t = _colores(estado)
    return (f'<div style="display:inline-block;min-width:132px;margin:0 8px 8px 0;padding:8px 10px;'
            f'border-radius:10px;background:{c}1f;border:1px solid {c}66;text-align:center;'
            f'vertical-align:top">'
            f'<div style="font-size:1.7rem;font-weight:700;line-height:1.1;color:{t}">{n}</div>'
            f'<div style="font-size:0.78rem;color:#374151;margin-top:2px">{estado}</div></div>')

def reporte(df):
    tarjetas = "".join(tarjeta(e, int((df["Estado"] == e).sum())) for e in ESTADOS)
    total = ('<div style="display:inline-block;min-width:132px;margin:0 8px 8px 0;padding:8px 10px;'
             'border-radius:10px;background:#eef2f7;border:1px solid #cbd5e1;text-align:center;'
             'vertical-align:top"><div style="font-size:1.7rem;font-weight:700;line-height:1.1;'
             f'color:#1e293b">{len(df)}</div>'
             '<div style="font-size:0.78rem;color:#334155;margin-top:2px">Total</div></div>')
    return total + tarjetas

def fondo_estado(val):
    c, _ = _colores(str(val))
    return f"background-color:{c}33"

@st.cache_data(ttl=30)
def cargar_datos(simplificar=True):
    if not os.path.exists(DB):
        return pd.DataFrame()
    con = sqlite3.connect(DB)
    df = pd.read_sql_query(
        """SELECT id, caso, oficina, solicitante, fecha_solicitud, regional,
                  aprobador, estado, fecha_proceso, fecha_respuesta
           FROM historial ORDER BY id DESC""", con)
    con.close()
    if simplificar and not df.empty:
        df = _aplicar_filtro_aprobados(df)
    return df

def _dias_aprobado(fecha_resp_str):
    """Días desde la fecha de respuesta. >2 días significa histórico."""
    try:
        dt = datetime.strptime(str(fecha_resp_str).strip()[:16], "%d/%m/%Y %H:%M")
        return (datetime.now() - dt).days
    except Exception:
        return 0

def _aplicar_filtro_aprobados(df):
    """Oculta APROBADO con más de 2 días de antigüedad."""
    if df.empty or "fecha_respuesta" not in df.columns:
        return df
    mask = (df["estado"] == "APROBADO") & df["fecha_respuesta"].notna() & (df["fecha_respuesta"] != "")
    if mask.any():
        visibles = ~(mask & (df["fecha_respuesta"].map(lambda x: _dias_aprobado(x) > 2)))
        df = df[visibles].copy()
    return df

def _parsedt(txt):
    """Convierte 'dd/mm/yyyy HH:MM' a datetime o None."""
    try:
        return datetime.strptime(str(txt).strip(), "%d/%m/%Y %H:%M")
    except Exception:
        return None

df = cargar_datos()

st.caption(f"Actualizado: {ahora_colombia()} (hora de Colombia)")

# --- Toggle: Simplificar vista (debe estar antes de usarlo) ---
col_t, _ = st.columns([2, 1])
with col_t:
    simplificado = st.toggle("🔍 Simplificar vista (oculta aprobados >2 días)", value=True,
                             help="Muestra solo pendientes y aprobados recientes (<2 días). "
                                  "Los aprobados antiguos se ocultan pero siguen disponibles "
                                  "en el histórico descendente o en la descarga.")
st.markdown("---")

# Recargar con el toggle aplicado
df = cargar_datos(simplificar=simplificado)

if st.button("🔄 Actualizar ahora"):
    cargar_datos.clear()
    st.rerun()

if df.empty:
    st.info("Aún no hay solicitudes registradas.")
    st.stop()

df["Estado"] = df["estado"].apply(estado_norm)
df = df.drop(columns=["estado", "id"])

st.markdown("**Estado del flujo**")
st.markdown(reporte(df), unsafe_allow_html=True)

# ---------- Gráficos ----------
st.markdown("#### 📊 Análisis")
g1, g2 = st.columns(2)

with g1:
    st.markdown("**Permisos por regional**")
    conteo = df["regional"].replace("", "Sin regional").fillna("Sin regional").value_counts()
    if not conteo.empty:
        st.bar_chart(conteo, horizontal=True, color="#0f766e", height=260, sort=False)
        st.caption("Total de solicitudes por regional (la más alta = la de mayor volumen).")
    else:
        st.info("Sin datos para el gráfico.")

with g2:
    st.markdown("**Tiempo de respuesta por regional (días)**")
    fin = df[df["Estado"].isin(["APROBADO", "RECHAZADO"])].copy()
    fin["sol_dt"] = fin["fecha_solicitud"].map(_parsedt)
    fin["res_dt"] = fin["fecha_respuesta"].map(_parsedt)
    fin = fin.dropna(subset=["sol_dt", "res_dt"])
    if fin.empty:
        st.info("Aún no hay respuestas con fecha.\nSe llena automáticamente cuando se marca "
                "APROBADO o RECHAZADO desde la gestión.")
    else:
        fin["dias"] = ((fin["res_dt"] - fin["sol_dt"]).dt.total_seconds() / 86400.0).round(1)
        prom = fin.groupby(fin["regional"].replace("", "Sin regional").fillna("Sin regional"))["dias"].mean()
        if prom.empty:
            st.info("Sin datos para el gráfico.")
        else:
            prom = prom.sort_values(ascending=False)
            st.bar_chart(prom, horizontal=True, color="#0d9488", height=260, sort=False)
            st.caption(f"Promedio por regional sobre {len(fin)} solicitudes respondidas.")

# ---------- Contadores por regional ----------
st.markdown("#### Tabla por regional")
tab_reg = df.copy()
tab_reg["Regional"] = tab_reg["regional"].replace("", "Sin regional").fillna("Sin regional")
cruce = pd.crosstab(tab_reg["Regional"], tab_reg["Estado"])
cruce = cruce.reindex(columns=["PENDIENTE APROBACIÓN", "APROBADO", "RECHAZADO"], fill_value=0)
cruce.columns = ["Pendiente aprobación", "Aprobados", "Rechazados"]
cruce["Total"] = cruce.sum(axis=1)
cruce = cruce.sort_values("Total", ascending=False)
st.dataframe(cruce, use_container_width=True,
             height=min(60 + 35 * len(cruce), 520))
st.caption("Contadores por regional: pendientes por aprobación y aprobados (los que usas para medir la carga de cada regional).")

st.markdown("#### 🔍 Filtros")
f1, f2 = st.columns([2, 3])
with f1:
    f_estado = st.multiselect("Estado", ESTADOS, key="reporte_estado")
with f2:
    f_buscar = st.text_input("Buscar por solicitud (oficina), caso, solicitante o aprobador:",
                             key="reporte_buscar")

cols = ["oficina", "caso", "solicitante", "fecha_solicitud", "regional", "aprobador", "Estado"]
muestras = df[cols].copy()
muestras.columns = ["Solicitud", "Caso", "Solicitado por", "Solicitado el",
                    "Regional", "Aprobador", "Estado"]
muestras = muestras[["Solicitud", "Caso", "Solicitado por", "Solicitado el",
                     "Regional", "Aprobador", "Estado"]]

mask = pd.Series(True, index=muestras.index)
if f_estado:
    mask &= muestras["Estado"].isin(f_estado)
if f_buscar.strip():
    q = f_buscar.strip()
    mask &= (muestras["Solicitud"].astype(str).str.contains(q, case=False, na=False) |
             muestras["Caso"].astype(str).str.contains(q, case=False, na=False) |
             muestras["Solicitado por"].astype(str).str.contains(q, case=False, na=False) |
             muestras["Aprobador"].astype(str).str.contains(q, case=False, na=False))

vis = muestras[mask]
if vis.empty:
    st.info("Ninguna solicitud coincide con los filtros.")
else:
    st.dataframe(vis.style.map(fondo_estado, subset=["Estado"]),
                 use_container_width=True, height=420,
                 hide_index=True)

    exp = muestras.copy()
    csv_bytes = exp.to_csv(index=False).encode("utf-8-sig")
    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as xw:
        exp.to_excel(xw, index=False, sheet_name="Historial")
    c1, c2 = st.columns(2)
    with c1:
        st.download_button("⬇️ Descargar reporte (CSV)", data=csv_bytes,
                           file_name="historial_solicitudes.csv", mime="text/csv")
    with c2:
        st.download_button("⬇️ Descargar reporte (Excel)", data=excel_buf.getvalue(),
                           file_name="historial_solicitudes.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
