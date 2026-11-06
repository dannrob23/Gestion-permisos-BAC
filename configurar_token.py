#!/usr/bin/env python3
"""Configura el token de GitHub que publica el historial en la consulta pública.

Uso:
    python configurar_token.py            # pide el token y lo verifica
    python configurar_token.py --estado   # solo informa si ya está configurado

El token se guarda en .streamlit/secrets.toml, archivo excluido de Git.
Nunca se imprime el valor del token en pantalla.
"""
from __future__ import annotations

import json
import os
import sys
from getpass import getpass
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

BASE = os.path.dirname(os.path.abspath(__file__))
DESTINO = os.path.join(BASE, ".streamlit", "secrets.toml")
REPO = "dannrob23/Gestion-permisos-BAC"


def _api(url: str, token: str) -> dict:
    peticion = Request(
        url,
        headers={
            "Authorization": "token " + token,
            "User-Agent": "configurar-token-bac",
            "Accept": "application/vnd.github+json",
        },
    )
    with urlopen(peticion, timeout=25) as respuesta:
        return json.load(respuesta)


def verificar(token: str) -> tuple[bool, str]:
    """Comprueba identidad y permiso de escritura sobre el repositorio."""
    try:
        usuario = _api("https://api.github.com/user", token)
    except HTTPError as exc:
        return False, f"GitHub rechazó el token (HTTP {exc.code})."
    except URLError as exc:
        return False, f"Sin conexión con GitHub: {exc.reason}"
    except Exception as exc:
        return False, f"Error consultando GitHub: {type(exc).__name__}"

    try:
        repo = _api(f"https://api.github.com/repos/{REPO}", token)
    except HTTPError as exc:
        return False, f"El token no accede al repositorio (HTTP {exc.code})."

    permisos = repo.get("permissions") or {}
    if not (permisos.get("push") or permisos.get("admin") or permisos.get("maintain")):
        return False, "El token no tiene permiso de escritura (push) sobre el repositorio."

    return True, f"Token válido para {usuario.get('login')} con permiso de escritura."


def escribir(token: str) -> None:
    os.makedirs(os.path.dirname(DESTINO), exist_ok=True)
    with open(DESTINO, "w", encoding="utf-8", newline="\n") as archivo:
        archivo.write('GITHUB_TOKEN = "' + token + '"\n')
    try:
        os.chmod(DESTINO, 0o600)
    except OSError:
        pass


def ya_configurado() -> bool:
    if not os.path.exists(DESTINO):
        return False
    with open(DESTINO, encoding="utf-8") as archivo:
        return "GITHUB_TOKEN" in archivo.read()


def main() -> int:
    if "--estado" in sys.argv:
        print("Token configurado: " + ("SI" if ya_configurado() else "NO"))
        print("Archivo: " + DESTINO)
        return 0

    print("Configuración del token de GitHub para publicar el historial público.")
    print("Créalo en https://github.com/settings/tokens con permiso sobre el repositorio.")
    print("El texto que escribas no se mostrará en pantalla.\n")

    token = getpass("Pega el token: ").strip()
    if not token:
        print("No se recibió ningún token. No se modificó nada.")
        return 1
    if len(token) < 20:
        print("El valor parece incompleto. No se modificó nada.")
        return 1

    correcto, mensaje = verificar(token)
    print(mensaje)
    if not correcto:
        print("No se guardó el token.")
        return 1

    escribir(token)
    print("Token guardado en .streamlit/secrets.toml (excluido de Git).")
    print("Ya puedes ejecutar el dashboard y los estados se publicarán automáticamente.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
