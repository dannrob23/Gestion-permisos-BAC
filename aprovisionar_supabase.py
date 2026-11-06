#!/usr/bin/env python3
"""Crea y configura el proyecto de Supabase para la consulta de permisos.

Requiere un token personal de Supabase (PAT) guardado en el archivo
`.supabase_pat` en la raíz del repositorio. Ese archivo está excluido de Git.

Pasos que ejecuta automáticamente:
  1. Crea el proyecto (si no existe).
  2. Espera a que quede disponible.
  3. Obtiene las claves anon y service_role.
  4. Crea las tablas con supabase/esquema.sql.
  5. Escribe .streamlit/secrets.toml.
  6. Migra el historial local a Supabase.

Uso:
    python aprovisionar_supabase.py
    python aprovisionar_supabase.py --estado    # solo informa el avance
"""
from __future__ import annotations

import json
import os
import secrets
import string
import sys
import time
import urllib.error
import urllib.request

BASE = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_PAT = os.path.join(BASE, ".supabase_pat")
ARCHIVO_CLAVE_DB = os.path.join(BASE, ".supabase_db_pass")
SECRETOS = os.path.join(BASE, ".streamlit", "secrets.toml")
ESQUEMA = os.path.join(BASE, "supabase", "esquema.sql")

ORG = "xosnjptymldqntghgjzy"
NOMBRE = "gestion-permisos-bac"
REGION = "sa-east-1"
API = "https://api.supabase.com/v1"


def _pat() -> str:
    if not os.path.exists(ARCHIVO_PAT):
        raise SystemExit(
            "Falta el archivo .supabase_pat con el token personal de Supabase.\n"
            "Créalo en https://supabase.com/dashboard/account/tokens y guárdalo ahí."
        )
    token = ""
    with open(ARCHIVO_PAT, encoding="utf-8") as archivo:
        for linea in archivo:
            limpia = linea.strip()
            if not limpia or limpia.startswith("#"):
                continue
            token = limpia.strip().strip('"').strip("'")
            break
    if not token.startswith("sbp_"):
        raise SystemExit(
            "El archivo .supabase_pat todavía no contiene un token válido.\n"
            "Debe haber una línea que empiece con sbp_ (sin comillas).\n"
            "Genéralo en https://supabase.com/dashboard/account/tokens"
        )
    return token


def _api(metodo: str, ruta: str, token: str, cuerpo: object = None):
    destino = API + ruta
    datos = json.dumps(cuerpo).encode("utf-8") if cuerpo is not None else None
    peticion = urllib.request.Request(
        destino,
        data=datos,
        headers={
            "Authorization": "Bearer " + token,
            "Content-Type": "application/json",
            "Accept": "application/json",
        },
        method=metodo,
    )
    try:
        with urllib.request.urlopen(peticion, timeout=120) as respuesta:
            texto = respuesta.read().decode("utf-8", "replace")
    except urllib.error.HTTPError as exc:
        detalle = exc.read().decode("utf-8", "replace")[:400]
        raise SystemExit(f"Supabase respondió {exc.code}: {detalle}") from exc
    except urllib.error.URLError as exc:
        raise SystemExit(f"Sin conexión con la API de Supabase: {exc.reason}") from exc
    return json.loads(texto) if texto.strip() else {}


def _clave_db() -> str:
    if os.path.exists(ARCHIVO_CLAVE_DB):
        with open(ARCHIVO_CLAVE_DB, encoding="utf-8") as archivo:
            guardada = archivo.read().strip()
        if guardada:
            return guardada
    alfabeto = string.ascii_letters + string.digits
    nueva = "".join(secrets.choice(alfabeto) for _ in range(28)) + "Aa9"
    with open(ARCHIVO_CLAVE_DB, "w", encoding="utf-8", newline="\n") as archivo:
        archivo.write(nueva + "\n")
    return nueva


def _proyectos(token: str) -> list[dict]:
    datos = _api("GET", f"/organizations/{ORG}/projects", token)
    return datos if isinstance(datos, list) else []


def _esperar(ref: str, token: str, limite: int = 600) -> dict:
    inicio = time.time()
    while time.time() - inicio < limite:
        estado = _api("GET", f"/projects/{ref}", token)
        situacion = str(estado.get("status") or "")
        print(f"  estado: {situacion or 'desconocido'}")
        if situacion in ("ACTIVE_HEALTHY", "ACTIVE_UNHEALTHY"):
            return estado
        if situacion in ("INIT_FAILED", "PAUSED", "REMOVED"):
            raise SystemExit(f"El proyecto quedó en estado {situacion}.")
        time.sleep(15)
    raise SystemExit("El proyecto no quedó disponible en el tiempo esperado.")


def _claves(ref: str, token: str) -> tuple[str, str]:
    datos = _api("GET", f"/projects/{ref}/api-keys", token)
    anon = servicio = ""
    for item in datos if isinstance(datos, list) else []:
        nombre = str(item.get("name") or "")
        valor = str(item.get("api_key") or "")
        if nombre == "anon":
            anon = valor
        elif nombre == "service_role":
            servicio = valor
    if not anon or not servicio:
        raise SystemExit("No se pudieron obtener las claves anon y service_role.")
    return anon, servicio


def _sql(ref: str, token: str, consulta: str) -> None:
    _api("POST", f"/projects/{ref}/database/query", token, {"query": consulta})


def _escribir_secretos(url: str, anon: str, servicio: str) -> None:
    os.makedirs(os.path.dirname(SECRETOS), exist_ok=True)
    lineas = []
    if os.path.exists(SECRETOS):
        with open(SECRETOS, encoding="utf-8") as archivo:
            for linea in archivo:
                if not linea.strip().startswith(("SUPABASE_URL", "SUPABASE_KEY", "SUPABASE_SERVICE_KEY")):
                    lineas.append(linea)
    lineas.append(f'SUPABASE_URL = "{url}"\n')
    lineas.append(f'SUPABASE_KEY = "{anon}"\n')
    lineas.append(f'SUPABASE_SERVICE_KEY = "{servicio}"\n')
    with open(SECRETOS, "w", encoding="utf-8", newline="\n") as archivo:
        archivo.writelines(lineas)
    try:
        os.chmod(SECRETOS, 0o600)
    except OSError:
        pass


def main() -> int:
    token = _pat()
    print("Organización:", ORG)

    existentes = _proyectos(token)
    print(f"Proyectos actuales en la organización: {len(existentes)}")
    for p in existentes:
        print(f"  - {p.get('name')} ({p.get('id')}) estado={p.get('status')}")

    objetivo = next((p for p in existentes if p.get("name") == NOMBRE), None)

    if objetivo is None:
        print(f"\nCreando proyecto '{NOMBRE}' en {REGION}...")
        creado = _api("POST", "/projects", token, {
            "organization_id": ORG,
            "name": NOMBRE,
            "db_pass": _clave_db(),
            "region": REGION,
            "plan": "free",
        })
        ref = str(creado.get("id") or "")
        if not ref:
            raise SystemExit(f"No se recibió el identificador del proyecto: {creado}")
    else:
        ref = str(objetivo.get("id"))
        print(f"\nEl proyecto ya existe: {ref}")

    print("Esperando disponibilidad del proyecto...")
    datos = _esperar(ref, token)
    url = "https://" + ref + ".supabase.co"
    print("Proyecto disponible:", url)

    print("Obteniendo claves...")
    anon, servicio = _claves(ref, token)

    print("Creando tablas y políticas...")
    with open(ESQUEMA, encoding="utf-8") as archivo:
        _sql(ref, token, archivo.read())

    print("Guardando secretos locales...")
    _escribir_secretos(url, anon, servicio)

    print("\n--- RESULTADO ---")
    print("SUPABASE_URL =", url)
    print("SUPABASE_KEY (anon) =", anon[:12] + "...")
    print("SUPABASE_SERVICE_KEY =", servicio[:12] + "...")
    print("Guardadas en:", SECRETOS)
    print("Contraseña de la base en:", ARCHIVO_CLAVE_DB)
    print("\nSiguiente paso: python migrar_a_supabase.py")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
