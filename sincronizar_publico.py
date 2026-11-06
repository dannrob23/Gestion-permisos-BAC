#!/usr/bin/env python3
"""Sincroniza el historial privado con la base de datos pública.

La base local (`base.db`) puede contener campos internos. Por eso el reporte
público se publica en una copia sanitizada: `publico/base.db`.

El módulo de gestión llama a :func:`sincronizar` después de guardar cambios.
La función exporta la copia y, si cambió, crea un commit y lo publica en
GitHub. El token se recibe desde `st.secrets` o `GITHUB_TOKEN`; nunca se
escribe en el remoto ni en archivos del repositorio.
"""
from __future__ import annotations

import os
import sqlite3
import subprocess
from typing import Optional

BASE = os.path.dirname(os.path.abspath(__file__))
ORIGEN = os.path.join(BASE, "base.db")
DESTINO = os.path.join(BASE, "publico", "base.db")
RUTA_PUBLICA_RELATIVA = "publico/base.db"
RAMA = "main"

COLUMNAS = [
    "id", "caso", "oficina", "solicitante", "fecha_solicitud",
    "regional", "aprobador", "estado", "fecha_proceso", "fecha_respuesta",
]


def estado_limpio(estado: object) -> str:
    """Normaliza estados antiguos al flujo visible en la consulta pública."""
    valor = str(estado or "").strip().upper()
    if valor in ("RECIBIDO", "PENDIENTE APROBACIÓN", "ENVIADO", "APROBADO", "RECHAZADO"):
        return valor
    if valor in ("REVISAR", "PENDIENTE"):
        return "PENDIENTE APROBACIÓN"
    if valor == "NOTIFICADO":
        return "APROBADO"
    return "RECIBIDO"


def _filas_publicas(filas: list[tuple]) -> list[tuple]:
    return [(*fila[:7], estado_limpio(fila[7]), fila[8], fila[9]) for fila in filas]


def _copia_equivalente(filas: list[tuple]) -> bool:
    """Evita reescribir SQLite si el contenido público no cambió."""
    if not os.path.exists(DESTINO):
        return False
    try:
        con = sqlite3.connect(DESTINO)
        schema = [fila[1] for fila in con.execute("PRAGMA table_info(historial)").fetchall()]
        actual = con.execute(
            "SELECT " + ", ".join(COLUMNAS) + " FROM historial ORDER BY id DESC"
        ).fetchall()
        con.close()
        return schema == COLUMNAS and actual == _filas_publicas(filas)
    except (sqlite3.Error, OSError):
        return False


def _crear_copia(filas: list[tuple]) -> int:
    """Escribe la copia de forma atómica y devuelve el número de filas."""
    if _copia_equivalente(filas):
        return len(filas)
    os.makedirs(os.path.dirname(DESTINO), exist_ok=True)
    temporal = DESTINO + ".tmp"
    if os.path.exists(temporal):
        os.remove(temporal)

    con_destino = sqlite3.connect(temporal)
    try:
        con_destino.execute("""CREATE TABLE historial (
            id INTEGER PRIMARY KEY, caso TEXT, oficina TEXT, solicitante TEXT,
            fecha_solicitud TEXT, regional TEXT, aprobador TEXT, estado TEXT,
            fecha_proceso TEXT, fecha_respuesta TEXT)""")
        con_destino.executemany(
            "INSERT INTO historial VALUES (?,?,?,?,?,?,?,?,?,?)",
            _filas_publicas(filas),
        )
        con_destino.commit()
    finally:
        con_destino.close()

    os.replace(temporal, DESTINO)
    return len(filas)


def exportar(verbose: bool = True) -> int:
    """Genera `publico/base.db` con únicamente los campos publicados."""
    if not os.path.exists(ORIGEN):
        raise FileNotFoundError(f"No se encontró la base local: {ORIGEN}")

    con_origen = sqlite3.connect(ORIGEN)
    try:
        filas = con_origen.execute(
            "SELECT " + ", ".join(COLUMNAS) + " FROM historial ORDER BY id DESC"
        ).fetchall()
    finally:
        con_origen.close()

    cantidad = _crear_copia(filas)
    if verbose:
        print(f"OK: {cantidad} solicitudes exportadas a {DESTINO}")
    return cantidad


def _git(args: list[str], token: Optional[str] = None) -> subprocess.CompletedProcess[str]:
    """Ejecuta git sin guardar el token en la URL del remoto."""
    entorno = dict(os.environ)
    entorno["GIT_TERMINAL_PROMPT"] = "0"
    if token:
        # Git consume la cabecera desde variables temporales del proceso.
        # No se modifica origin y el secreto no queda en .git/config.
        entorno["GIT_CONFIG_COUNT"] = "1"
        entorno["GIT_CONFIG_KEY_0"] = "http.extraHeader"
        entorno["GIT_CONFIG_VALUE_0"] = "AUTHORIZATION: bearer " + token
    return subprocess.run(
        ["git", "-C", BASE, *args],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=entorno,
    )


def sincronizar(token: Optional[str] = None, commit_message: str = "Sincronizar historial público") -> dict:
    """Exporta y publica la copia si cambió.

    Devuelve un diccionario verificable con `cambio`, `publicado`, `filas` y
    `mensaje`. Si no hay token, la copia queda actualizada localmente, pero se
    informa que falta publicar en GitHub.
    """
    filas = exportar(verbose=False)
    estado = _git(["status", "--porcelain", "--", RUTA_PUBLICA_RELATIVA], token)
    if estado.returncode != 0:
        return {"cambio": False, "publicado": False, "filas": filas,
                "mensaje": f"Git no disponible: {estado.stderr.strip()[:250]}"}

    if not estado.stdout.strip():
        return {"cambio": False, "publicado": True, "filas": filas,
                "mensaje": "La copia pública ya estaba actualizada."}

    if not token:
        return {"cambio": True, "publicado": False, "filas": filas,
                "mensaje": "Copia local actualizada; falta GITHUB_TOKEN para publicarla."}

    agregado = _git(["add", "--", RUTA_PUBLICA_RELATIVA], token)
    if agregado.returncode != 0:
        return {"cambio": True, "publicado": False, "filas": filas,
                "mensaje": f"No se pudo preparar la copia: {agregado.stderr.strip()[:250]}"}

    commit = _git(["commit", "-m", commit_message], token)
    if commit.returncode != 0:
        return {"cambio": True, "publicado": False, "filas": filas,
                "mensaje": f"No se pudo crear el commit: {commit.stderr.strip()[:250]}"}

    push = _git(["push", "origin", RAMA], token)
    if push.returncode != 0:
        # Recupera cambios de otras sesiones y reintenta una sola vez.
        pull = _git(["pull", "--rebase", "--autostash", "origin", RAMA], token)
        if pull.returncode == 0:
            push = _git(["push", "origin", RAMA], token)

    if push.returncode == 0:
        return {"cambio": True, "publicado": True, "filas": filas,
                "mensaje": "Historial público actualizado en GitHub."}
    return {"cambio": True, "publicado": False, "filas": filas,
            "mensaje": f"La copia quedó en un commit local, pero falló el push: {push.stderr.strip()[:300]}"}


if __name__ == "__main__":
    resultado = sincronizar(token=os.environ.get("GITHUB_TOKEN"))
    print(resultado["mensaje"])
