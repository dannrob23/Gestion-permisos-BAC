#!/usr/bin/env python3
"""Convierte las credenciales de Google al formato TOML de Streamlit Cloud.

Lee el archivo JSON de la cuenta de servicio en el equipo local, genera el
bloque de secretos y lo copia al portapapeles de Windows. El contenido nunca
se imprime en pantalla ni se envía a ningún servicio.

Uso:
    python preparar_secretos.py

Luego, en Streamlit Cloud: Manage app -> Settings -> Secrets -> pegar y guardar.
"""
from __future__ import annotations

import json
import os
import subprocess
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
CREDENCIAL = os.path.join(BASE, "credenciales_google.json")
SALIDA = os.path.join(BASE, ".streamlit", "secrets_cloud.toml")
HOJA_ID = "1u3U5aCG22tqxVN6_NVL96LZlBJ0kvfCiFYHuRBwo0sA"

CAMPOS = [
    "type", "project_id", "private_key_id", "private_key", "client_email",
    "client_id", "auth_uri", "token_uri", "auth_provider_x509_cert_url",
    "client_x509_cert_url",
]


def _escapar(valor: str) -> str:
    """Escapa un valor para usarlo entre comillas dobles en TOML."""
    texto = str(valor)
    texto = texto.replace("\\", "\\\\")
    texto = texto.replace('"', '\\"')
    texto = texto.replace("\r\n", "\\n").replace("\n", "\\n")
    texto = texto.replace("\r", "\\n")
    texto = texto.replace("\t", "\\t")
    return texto


def _copiar(texto: str) -> bool:
    """Copia al portapapeles de Windows sin mostrar el contenido."""
    try:
        proceso = subprocess.run(
            ["powershell.exe", "-NoProfile", "-Command", "Set-Clipboard -Value ([Console]::In.ReadToEnd())"],
            input=texto.encode("utf-8"),
            capture_output=True,
            timeout=60,
        )
        return proceso.returncode == 0
    except Exception:
        return False


def main() -> int:
    if not os.path.exists(CREDENCIAL):
        print(f"No se encontró el archivo: {CREDENCIAL}")
        print("Descárgalo de Google Cloud -> IAM -> Cuentas de servicio -> Claves.")
        return 1

    with open(CREDENCIAL, encoding="utf-8") as archivo:
        datos = json.load(archivo)

    faltantes = [c for c in CAMPOS if c not in datos]
    if faltantes:
        print("El archivo no parece ser una cuenta de servicio. Faltan:",
              ", ".join(faltantes))
        return 1

    lineas = [
        f'GOOGLE_SHEET_ID = "{HOJA_ID}"',
        "",
        "[gcp_service_account]",
    ]
    for campo in CAMPOS:
        lineas.append(f'{campo} = "{_escapar(datos[campo])}"')

    lineas.append("")
    lineas.append("# Clave del modo administrador de la consulta.")
    lineas.append("# Cambiala por una propia y no la compartas.")
    lineas.append("[admin]")
    lineas.append('password = "CAMBIA_ESTA_CLAVE"')
    lineas.append("")

    contenido = "\n".join(lineas)

    os.makedirs(os.path.dirname(SALIDA), exist_ok=True)
    with open(SALIDA, "w", encoding="utf-8", newline="\n") as archivo:
        archivo.write(contenido)
    try:
        os.chmod(SALIDA, 0o600)
    except OSError:
        pass

    copiado = _copiar(contenido)

    print("Bloque de secretos generado correctamente.")
    print(f"Archivo local: {SALIDA}")
    print(f"Cuenta de servicio: {datos['client_email']}")
    print(f"Proyecto: {datos['project_id']}")
    print(f"Campos incluidos: {len(CAMPOS)}")
    print()
    if copiado:
        print("YA ESTA COPIADO AL PORTAPAPELES.")
        print("Abre Streamlit Cloud -> Manage app -> Settings -> Secrets")
        print("Selecciona todo (Ctrl+A), pega (Ctrl+V) y guarda.")
    else:
        print("No se pudo copiar automaticamente. Abre el archivo indicado")
        print("y copia su contenido manualmente.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
