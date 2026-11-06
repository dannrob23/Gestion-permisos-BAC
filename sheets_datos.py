#!/usr/bin/env python3
"""Acceso a la hoja de Google que almacena las solicitudes de ingreso.

La hoja es la fuente única de datos: la consulta pública lee de ella y las
actualizaciones de estado se escriben en ella. Requiere una cuenta de servicio
con permiso de edición sobre la hoja (ver cargar_a_sheets.py).
"""
from __future__ import annotations

import datetime
import email.utils
import os
import time
import urllib.request
from typing import Optional

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
PESTANA = "Solicitudes"

ENCABEZADO = [
    "id", "caso", "oficina", "solicitante", "fecha_solicitud",
    "regional", "aprobador", "estado", "fecha_proceso", "fecha_respuesta",
]

ESTADOS_VALIDOS = {
    "RECIBIDO", "PENDIENTE APROBACIÓN", "ENVIADO", "APROBADO", "RECHAZADO",
}
LEGACY = {
    "REVISAR": "PENDIENTE APROBACIÓN",
    "PENDIENTE": "PENDIENTE APROBACIÓN",
    "NOTIFICADO": "APROBADO",
}


class ErrorSheets(RuntimeError):
    """Falla de comunicación o de permisos con Google Sheets."""


_RUTA_CREDENCIAL = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "credenciales_google.json")


def configurar(ruta_credencial: Optional[str] = None) -> None:
    """Permite indicar otra ubicación para el archivo de credenciales."""
    global _RUTA_CREDENCIAL
    if ruta_credencial:
        _RUTA_CREDENCIAL = ruta_credencial


def disponible(ruta: Optional[str] = None) -> bool:
    return os.path.exists(ruta or _RUTA_CREDENCIAL)


def _corregir_reloj() -> None:
    """Compensa el desfase del reloj local frente a Google.

    Google rechaza tokens firmados con hora futura; si el equipo está
    adelantado se ajusta la marca de tiempo usada al firmar.
    """
    # Evita reajustar si ya se hizo en esta ejecución.
    if getattr(_corregir_reloj, "_aplicado", False):
        return

    desfase = _desfase_cache()
    if desfase is None:
        try:
            respuesta = urllib.request.urlopen("https://www.google.com", timeout=8)
            real = email.utils.parsedate_to_datetime(respuesta.headers["Date"]).timestamp()
            ahora = datetime.datetime.now(datetime.timezone.utc).timestamp()
            desfase = ahora - real
            if abs(desfase) > 86400:  # más de un día: medida no confiable
                desfase = 0.0
            _guardar_desfase(desfase)
        except Exception:
            desfase = 0.0

    if abs(desfase) < 60:
        _corregir_reloj._aplicado = True
        return

    import google.auth._helpers as helpers
    helpers.utcnow = (
        lambda: datetime.datetime.utcnow()
        - datetime.timedelta(seconds=desfase)
    )
    _corregir_reloj._aplicado = True


_ARCHIVO_DESFASE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), ".desfase_reloj")


def _desfase_cache() -> Optional[float]:
    """Lee el desfase medido previamente, si existe y es reciente."""
    try:
        if not os.path.exists(_ARCHIVO_DESFASE):
            return None
        if time.time() - os.path.getmtime(_ARCHIVO_DESFASE) > 86400:
            return None
        with open(_ARCHIVO_DESFASE, encoding="utf-8") as archivo:
            return float(archivo.read().strip())
    except Exception:
        return None


def _guardar_desfase(valor: float) -> None:
    try:
        with open(_ARCHIVO_DESFASE, "w", encoding="utf-8", newline="\n") as archivo:
            archivo.write(str(int(valor)))
    except Exception:
        pass


def servicio(hoja_id: str, ruta_credencial: Optional[str] = None):
    """Devuelve el cliente de la API de Sheets."""
    ruta = ruta_credencial or _RUTA_CREDENCIAL
    if not os.path.exists(ruta):
        raise ErrorSheets("No se encontró el archivo de credenciales de Google.")
    _corregir_reloj()
    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        cred = service_account.Credentials.from_service_account_file(
            ruta, scopes=SCOPES)
        return build("sheets", "v4", credentials=cred, cache_discovery=False)
    except Exception as exc:
        raise ErrorSheets(f"No se pudo autenticar con Google: {exc}") from exc


def estado_valido(valor: object) -> str:
    texto = str(valor or "").strip().upper()
    if texto in ESTADOS_VALIDOS:
        return texto
    return LEGACY.get(texto, "RECIBIDO")


def leer_filas(hoja_id: str, ruta_credencial: Optional[str] = None) -> list[dict]:
    """Devuelve todas las filas como diccionarios con el encabezado aplicado."""
    api = servicio(hoja_id, ruta_credencial)
    try:
        resultado = api.spreadsheets().values().get(
            spreadsheetId=hoja_id, range=f"{PESTANA}!A1:Z100000").execute()
    except Exception as exc:
        raise ErrorSheets(f"No se pudo leer la hoja: {str(exc)[:200]}") from exc

    valores = resultado.get("values", [])
    if not valores:
        return []
    cabecera = [str(c).strip() for c in valores[0]]
    filas = []
    for fila in valores[1:]:
        if not any(str(c).strip() for c in fila):
            continue
        datos = dict(zip(cabecera, fila))
        # Rellena columnas faltantes para que el DataFrame tenga forma estable.
        for columna in ENCABEZADO:
            datos.setdefault(columna, "")
        datos["estado"] = estado_valido(datos.get("estado"))
        filas.append(datos)
    return filas


def _columna_estado(api, hoja_id: str) -> int:
    """Número de columna (1-based) que contiene el estado."""
    resultado = api.spreadsheets().values().get(
        spreadsheetId=hoja_id, range=f"{PESTANA}!A1:Z1").execute()
    cabecera = [str(c).strip() for c in resultado.get("values", [[]])[0]]
    try:
        return cabecera.index("estado") + 1
    except ValueError as exc:
        raise ErrorSheets("La hoja no tiene la columna 'estado'.") from exc


def actualizar_estado(
    hoja_id: str,
    solicitud_id: int,
    nuevo_estado: str,
    ruta_credencial: Optional[str] = None,
    fecha_respuesta: str = "",
) -> dict:
    """Cambia el estado de una solicitud identificada por su id."""
    estado = str(nuevo_estado or "").strip().upper()
    if estado not in ESTADOS_VALIDOS:
        raise ErrorSheets(f"Estado no permitido: {nuevo_estado!r}")

    api = servicio(hoja_id, ruta_credencial)
    resultado = api.spreadsheets().values().get(
        spreadsheetId=hoja_id, range=f"{PESTANA}!A1:Z100000").execute()
    valores = resultado.get("values", [])
    if not valores:
        raise ErrorSheets("La hoja está vacía.")

    cabecera = [str(c).strip() for c in valores[0]]
    if "id" not in cabecera or "estado" not in cabecera:
        raise ErrorSheets("La hoja no tiene las columnas 'id' y 'estado'.")

    col_id = cabecera.index("id")
    col_estado = cabecera.index("estado")
    col_respuesta = cabecera.index("fecha_respuesta") if "fecha_respuesta" in cabecera else None

    for numero, fila in enumerate(valores[1:], start=2):
        if len(fila) <= col_id:
            continue
        if str(fila[col_id]).strip() != str(int(solicitud_id)):
            continue

        letra = chr(ord("A") + col_estado)
        api.spreadsheets().values().update(
            spreadsheetId=hoja_id,
            range=f"{PESTANA}!{letra}{numero}",
            valueInputOption="RAW",
            body={"values": [[estado]]},
        ).execute()

        if fecha_respuesta and col_respuesta is not None:
            letra_r = chr(ord("A") + col_respuesta)
            api.spreadsheets().values().update(
                spreadsheetId=hoja_id,
                range=f"{PESTANA}!{letra_r}{numero}",
                valueInputOption="RAW",
                body={"values": [[fecha_respuesta]]},
            ).execute()

        return {"id": int(solicitud_id), "estado": estado, "fila": numero}

    raise ErrorSheets(f"No se encontró la solicitud {solicitud_id} en la hoja.")


def ultima_actualizacion(hoja_id: str, ruta_credencial: Optional[str] = None) -> str:
    """Fecha de la última modificación de la hoja, en formato local."""
    api = servicio(hoja_id, ruta_credencial)
    try:
        meta = api.spreadsheets().get(
            spreadsheetId=hoja_id,
            fields="properties.title",
            includeGridData=False,
        ).execute()
        return meta.get("properties", {}).get("title", "")
    except Exception:
        return ""
