#!/usr/bin/env python3
"""Acceso a Supabase para el historial de solicitudes de ingreso.

Usa la API REST (PostgREST) con la librería estándar: no requiere dependencias
nuevas. La aplicación pública lee con la clave pública (anon); las escrituras
requieren la clave de servicio, que nunca debe publicarse en el repositorio.

Variables esperadas:
    SUPABASE_URL           https://<proyecto>.supabase.co
    SUPABASE_KEY           clave anon (solo lectura)
    SUPABASE_SERVICE_KEY   clave service_role (lectura y escritura)
"""
from __future__ import annotations

import json
import os
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timezone
from typing import Optional

TABLA = "solicitudes"
AUDITORIA = "historial_estados"
COLUMNAS = [
    "id", "caso", "oficina", "solicitante", "fecha_solicitud",
    "regional", "aprobador", "estado", "fecha_proceso", "fecha_respuesta",
]
ESTADOS_VALIDOS = {
    "RECIBIDO", "PENDIENTE APROBACIÓN", "ENVIADO", "APROBADO", "RECHAZADO",
}


class ErrorSupabase(RuntimeError):
    """Falla de comunicación o de permisos contra Supabase."""


def configurado(url: Optional[str] = None, key: Optional[str] = None) -> bool:
    url = url if url is not None else os.environ.get("SUPABASE_URL", "")
    key = key if key is not None else os.environ.get("SUPABASE_KEY", "")
    return bool(url.strip() and key.strip())


def _pedir(
    url: str,
    key: str,
    ruta: str,
    metodo: str = "GET",
    cuerpo: object = None,
    cabeceras_extra: Optional[dict] = None,
) -> list:
    """Ejecuta una petición REST contra Supabase y devuelve las filas."""
    destino = url.rstrip("/") + "/rest/v1/" + ruta
    datos = None
    cabeceras = {
        "apikey": key,
        "Authorization": "Bearer " + key,
        "Accept": "application/json",
        "Content-Type": "application/json",
    }
    if cabeceras_extra:
        cabeceras.update(cabeceras_extra)
    if cuerpo is not None:
        datos = json.dumps(cuerpo).encode("utf-8")

    peticion = urllib.request.Request(destino, data=datos, headers=cabeceras, method=metodo)
    try:
        with urllib.request.urlopen(peticion, timeout=30) as respuesta:
            texto = respuesta.read().decode("utf-8", "replace")
    except urllib.error.HTTPError as exc:
        detalle = exc.read().decode("utf-8", "replace")[:300]
        raise ErrorSupabase(f"Supabase respondió {exc.code}: {detalle}") from exc
    except urllib.error.URLError as exc:
        raise ErrorSupabase(f"No se pudo conectar con Supabase: {exc.reason}") from exc

    if not texto.strip():
        return []
    try:
        resultado = json.loads(texto)
    except json.JSONDecodeError as exc:
        raise ErrorSupabase("Respuesta no válida de Supabase") from exc
    return resultado if isinstance(resultado, list) else [resultado]


def cargar_solicitudes(url: str, key: str) -> list[dict]:
    """Devuelve todas las solicitudes ordenadas de la más reciente a la más antigua."""
    columnas = ",".join(COLUMNAS + ["actualizado_en"])
    ruta = f"{TABLA}?select={columnas}&order=id.desc"
    return _pedir(url, key, ruta)


def ultima_actualizacion(url: str, key: str) -> str:
    """Fecha del cambio más reciente registrado en la tabla."""
    filas = _pedir(url, key, f"{TABLA}?select=actualizado_en&order=actualizado_en.desc&limit=1")
    if not filas:
        return ""
    return str(filas[0].get("actualizado_en") or "")


def actualizar_estado(
    url: str,
    clave_servicio: str,
    solicitud_id: int,
    nuevo_estado: str,
    usuario: str = "",
    observacion: str = "",
    estado_anterior: str = "",
) -> dict:
    """Cambia el estado de una solicitud y deja registro de auditoría."""
    estado = str(nuevo_estado or "").strip().upper()
    if estado not in ESTADOS_VALIDOS:
        raise ErrorSupabase(f"Estado no permitido: {nuevo_estado!r}")

    marca = datetime.now(timezone.utc).isoformat(timespec="seconds")
    cuerpo = {
        "estado": estado,
        "actualizado_por": (usuario or "").strip()[:120],
        "actualizado_en": marca,
    }
    if estado in ("APROBADO", "RECHAZADO") and not estado_anterior:
        cuerpo["fecha_respuesta"] = datetime.now().strftime("%d/%m/%Y %H:%M")

    filas = _pedir(
        url, clave_servicio,
        f"{TABLA}?id=eq.{int(solicitud_id)}",
        metodo="PATCH",
        cuerpo=cuerpo,
        cabeceras_extra={"Prefer": "return=representation"},
    )

    try:
        _pedir(
            url, clave_servicio, AUDITORIA, metodo="POST",
            cuerpo={
                "solicitud_id": int(solicitud_id),
                "estado_anterior": estado_anterior or None,
                "estado_nuevo": estado,
                "usuario": (usuario or "sin identificar")[:120],
                "observacion": (observacion or "")[:500],
            },
            cabeceras_extra={"Prefer": "return=minimal"},
        )
    except ErrorSupabase:
        # El cambio ya quedó aplicado; la auditoría no debe bloquear la operación.
        pass

    return filas[0] if filas else {"id": int(solicitud_id), "estado": estado}


def insertar_solicitudes(url: str, clave_servicio: str, filas: list[dict]) -> int:
    """Inserta o actualiza solicitudes en lote (usado por la migración)."""
    if not filas:
        return 0
    total = 0
    for inicio in range(0, len(filas), 200):
        lote = filas[inicio:inicio + 200]
        _pedir(
            url, clave_servicio, TABLA, metodo="POST", cuerpo=lote,
            cabeceras_extra={
                "Prefer": "resolution=merge-duplicates,return=minimal"
            },
        )
        total += len(lote)
    return total


def probar(url: str, key: str) -> tuple[bool, str]:
    """Comprueba que la tabla sea accesible. No revela las claves."""
    try:
        filas = _pedir(url, key, f"{TABLA}?select=id&limit=1")
    except ErrorSupabase as exc:
        return False, str(exc)
    return True, f"Conexión correcta. Registros visibles: {len(filas)}."
