#!/usr/bin/env python3
"""Exporta una copia SANITIZADA del historial para la consulta pública.

Solo se copian los campos que muestra el reporte (sin correos internos,
asuntos, notas ni datos personales sensibles) y los estados se llevan
al flujo actual (RECIBIDO / PENDIENTE APROBACIÓN / ENVIADO / APROBADO / RECHAZADO).

Uso:
    python sincronizar_publico.py
"""
import os, sqlite3

BASE = os.path.dirname(os.path.abspath(__file__))
ORIGEN = os.path.join(BASE, "base.db")
DESTINO = os.path.join(BASE, "publico", "base.db")

COLUMNAS = ["id", "caso", "oficina", "solicitante", "fecha_solicitud",
            "regional", "aprobador", "estado", "fecha_proceso", "fecha_respuesta"]

def estado_limpio(e):
    e = (e or "").strip().upper()
    if e in ("RECIBIDO", "PENDIENTE APROBACIÓN", "ENVIADO", "APROBADO", "RECHAZADO"):
        return e
    if e in ("REVISAR", "PENDIENTE"):
        return "PENDIENTE APROBACIÓN"
    if e == "NOTIFICADO":
        return "APROBADO"
    return "RECIBIDO"

def exportar(verbose=True):
    if not os.path.exists(ORIGEN):
        raise SystemExit("No se encontró la base local: " + ORIGEN)
    con_o = sqlite3.connect(ORIGEN)
    filas = con_o.execute(
        "SELECT " + ", ".join(COLUMNAS) + " FROM historial ORDER BY id DESC").fetchall()
    con_o.close()

    if os.path.exists(DESTINO):
        os.remove(DESTINO)
    con_d = sqlite3.connect(DESTINO)
    con_d.execute("""CREATE TABLE historial (
        id INTEGER PRIMARY KEY, caso TEXT, oficina TEXT, solicitante TEXT,
        fecha_solicitud TEXT, regional TEXT, aprobador TEXT, estado TEXT,
        fecha_proceso TEXT, fecha_respuesta TEXT)""")
    con_d.executemany(
        "INSERT INTO historial VALUES (?,?,?,?,?,?,?,?,?,?)",
        [(*r[:7], estado_limpio(r[7]), r[8], r[9]) for r in filas])
    con_d.commit()
    con_d.close()
    if verbose:
        print(f"OK: {len(filas)} solicitudes exportadas a {DESTINO}")

if __name__ == "__main__":
    exportar()
