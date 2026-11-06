#!/usr/bin/env python3
"""Copia el historial local (base.db) hacia la tabla de Supabase.

Uso:
    python migrar_a_supabase.py            # copia todas las solicitudes
    python migrar_a_supabase.py --probar   # solo verifica la conexión

Requiere SUPABASE_URL y SUPABASE_SERVICE_KEY (variables de entorno o
.streamlit/secrets.toml). Es idempotente: si una solicitud ya existe, la
actualiza en lugar de duplicarla.
"""
from __future__ import annotations

import os
import sqlite3
import sys

try:
    import tomllib
except ModuleNotFoundError:  # Python < 3.11
    tomllib = None

import supabase_datos as SB

BASE = os.path.dirname(os.path.abspath(__file__))
ORIGEN = os.path.join(BASE, "base.db")
SECRETOS = os.path.join(BASE, ".streamlit", "secrets.toml")

LEGACY = {
    "REVISAR": "PENDIENTE APROBACIÓN",
    "PENDIENTE": "PENDIENTE APROBACIÓN",
    "NOTIFICADO": "APROBADO",
}


def _estado_valido(valor: str) -> str:
    texto = str(valor or "").strip().upper()
    if texto in SB.ESTADOS_VALIDOS:
        return texto
    return LEGACY.get(texto, "RECIBIDO")


def _leer_secretos() -> dict:
    datos: dict = {}
    if os.path.exists(SECRETOS) and tomllib is not None:
        try:
            with open(SECRETOS, "rb") as archivo:
                datos = tomllib.load(archivo)
        except Exception:
            datos = {}
    for clave in ("SUPABASE_URL", "SUPABASE_KEY", "SUPABASE_SERVICE_KEY"):
        if not datos.get(clave) and os.environ.get(clave):
            datos[clave] = os.environ[clave]
    return datos


def _filas_locales() -> list[dict]:
    if not os.path.exists(ORIGEN):
        raise SystemExit(f"No se encontró la base local: {ORIGEN}")
    con = sqlite3.connect(ORIGEN)
    con.row_factory = sqlite3.Row
    registros = con.execute(
        "SELECT id, caso, oficina, solicitante, fecha_solicitud, regional, "
        "aprobador, estado, fecha_proceso, fecha_respuesta "
        "FROM historial ORDER BY id"
    ).fetchall()
    con.close()

    filas = []
    for r in registros:
        filas.append({
            "id": int(r["id"]),
            "caso": r["caso"] or "",
            "oficina": r["oficina"] or "",
            "solicitante": r["solicitante"] or "",
            "fecha_solicitud": r["fecha_solicitud"] or "",
            "regional": r["regional"] or "",
            "aprobador": r["aprobador"] or "",
            "estado": _estado_valido(r["estado"]),
            "fecha_proceso": r["fecha_proceso"] or "",
            "fecha_respuesta": r["fecha_respuesta"] or "",
        })
    return filas


def main() -> int:
    secretos = _leer_secretos()
    url = str(secretos.get("SUPABASE_URL") or "").strip()
    servicio = str(secretos.get("SUPABASE_SERVICE_KEY") or "").strip()
    lectura = str(secretos.get("SUPABASE_KEY") or servicio).strip()

    if not url or not lectura:
        print("Falta SUPABASE_URL o SUPABASE_KEY.")
        print(f"Configúralos en {SECRETOS} o como variables de entorno.")
        return 1

    ok, mensaje = SB.probar(url, lectura)
    print(mensaje)
    if not ok:
        return 1
    if "--probar" in sys.argv:
        return 0

    if not servicio:
        print("Falta SUPABASE_SERVICE_KEY: sin ella no se pueden escribir datos.")
        return 1

    filas = _filas_locales()
    print(f"Registros locales a migrar: {len(filas)}")
    total = SB.insertar_solicitudes(url, servicio, filas)
    print(f"Migrados correctamente: {total}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
