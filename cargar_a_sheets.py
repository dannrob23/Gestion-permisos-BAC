#!/usr/bin/env python3
"""Carga el historial local (base.db) en la hoja de cálculo de Google.

Uso:
    python cargar_a_sheets.py            # carga o actualiza todas las filas
    python cargar_a_sheets.py --probar   # solo verifica el acceso

Requiere credenciales_google.json y la hoja compartida con la cuenta de servicio.
Es idempotente: reescribe la pestaña respetando el encabezado.
"""
from __future__ import annotations

import datetime
import email.utils
import os
import sqlite3
import sys
import urllib.request

BASE = os.path.dirname(os.path.abspath(__file__))
ORIGEN = os.path.join(BASE, "base.db")
CREDENCIAL = os.path.join(BASE, "credenciales_google.json")
HOJA_ID = "1u3U5aCG22tqxVN6_NVL96LZlBJ0kvfCiFYHuRBwo0sA"
PESTANA = "Solicitudes"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

ENCABEZADO = [
    "id", "caso", "oficina", "solicitante", "fecha_solicitud",
    "regional", "aprobador", "estado", "fecha_proceso", "fecha_respuesta",
]

LEGACY = {
    "REVISAR": "PENDIENTE APROBACIÓN",
    "PENDIENTE": "PENDIENTE APROBACIÓN",
    "NOTIFICADO": "APROBADO",
}
VALIDOS = {"RECIBIDO", "PENDIENTE APROBACIÓN", "ENVIADO", "APROBADO", "RECHAZADO"}


def _corregir_reloj() -> None:
    """Compensa el desfase del reloj local frente a Google.

    Google rechaza los tokens firmados con hora futura. Si el equipo tiene el
    reloj adelantado, se ajusta la marca de tiempo usada para firmar.
    """
    try:
        respuesta = urllib.request.urlopen("https://www.google.com", timeout=20)
        real = email.utils.parsedate_to_datetime(respuesta.headers["Date"]).timestamp()
        ahora = datetime.datetime.now(datetime.timezone.utc).timestamp()
        desfase = ahora - real
        if abs(desfase) < 60:
            return
        import google.auth._helpers as helpers
        helpers.utcnow = (
            lambda: datetime.datetime.utcnow()
            - datetime.timedelta(seconds=desfase)
        )
        print(f"Aviso: reloj local desfasado {int(desfase)} s; compensado.")
    except Exception:
        pass


def _servicio():
    if not os.path.exists(CREDENCIAL):
        raise SystemExit(f"No se encontró {CREDENCIAL}")
    _corregir_reloj()
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    cred = service_account.Credentials.from_service_account_file(
        CREDENCIAL, scopes=SCOPES)
    return build("sheets", "v4", credentials=cred, cache_discovery=False)


def _estado(valor: str) -> str:
    texto = str(valor or "").strip().upper()
    if texto in VALIDOS:
        return texto
    return LEGACY.get(texto, "RECIBIDO")


def _filas_locales() -> list[list[str]]:
    if not os.path.exists(ORIGEN):
        raise SystemExit(f"No se encontró la base local: {ORIGEN}")
    con = sqlite3.connect(ORIGEN)
    registros = con.execute(
        "SELECT id, caso, oficina, solicitante, fecha_solicitud, regional, "
        "aprobador, estado, fecha_proceso, fecha_respuesta "
        "FROM historial ORDER BY id DESC"
    ).fetchall()
    con.close()

    filas = []
    for r in registros:
        fila = ["" if v is None else str(v) for v in r]
        fila[7] = _estado(fila[7])
        filas.append(fila)
    return filas


def _asegurar_pestana(api) -> None:
    meta = api.spreadsheets().get(
        spreadsheetId=HOJA_ID, fields="sheets.properties.title").execute()
    titulos = [s["properties"]["title"] for s in meta.get("sheets", [])]
    if PESTANA in titulos:
        return
    api.spreadsheets().batchUpdate(
        spreadsheetId=HOJA_ID,
        body={"requests": [{"addSheet": {"properties": {"title": PESTANA}}}]},
    ).execute()
    print(f"Pestaña '{PESTANA}' creada.")


def main() -> int:
    api = _servicio()
    try:
        meta = api.spreadsheets().get(
            spreadsheetId=HOJA_ID, fields="properties.title").execute()
        print("Hoja:", meta["properties"]["title"])
    except Exception as exc:
        print("Sin acceso a la hoja:", str(exc)[:250])
        return 1

    if "--probar" in sys.argv:
        print("Acceso correcto.")
        return 0

    _asegurar_pestana(api)
    filas = _filas_locales()
    print(f"Registros a escribir: {len(filas)}")

    api.spreadsheets().values().clear(
        spreadsheetId=HOJA_ID, range=f"{PESTANA}!A1:Z10000").execute()
    api.spreadsheets().values().update(
        spreadsheetId=HOJA_ID,
        range=f"{PESTANA}!A1",
        valueInputOption="RAW",
        body={"values": [ENCABEZADO] + filas},
    ).execute()

    verificacion = api.spreadsheets().values().get(
        spreadsheetId=HOJA_ID, range=f"{PESTANA}!A1:A10000").execute()
    total = max(0, len(verificacion.get("values", [])) - 1)
    print(f"Verificado en la hoja: {total} solicitudes")
    return 0 if total == len(filas) else 1


if __name__ == "__main__":
    raise SystemExit(main())
