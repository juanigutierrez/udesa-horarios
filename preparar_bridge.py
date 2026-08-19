from __future__ import annotations

import json
import re
import secrets
from pathlib import Path

ROOT = Path(__file__).resolve().parent
PRIVATE_JSON = ROOT / "admin_fuentes_privadas.json"
PRIVATE_GS = ROOT / "apps_script" / "CONFIGURACION_PRIVADA.gs"
SECRET_TEMPLATE = ROOT / ".streamlit" / "secrets.local.template.toml"


def extract_id(value: str) -> str:
    value = value.strip()
    for p in (r"/d/([A-Za-z0-9_-]+)", r"[?&]id=([A-Za-z0-9_-]+)"):
        m = re.search(p, value)
        if m:
            return m.group(1)
    if re.fullmatch(r"[A-Za-z0-9_-]{15,}", value):
        return value
    raise ValueError("No pude identificar el ID. Pegá el link completo de Drive o el ID del archivo.")


def ask(label: str) -> str:
    while True:
        value = input(f"{label}\nLink o ID: ").strip()
        try:
            return extract_id(value)
        except ValueError as exc:
            print("  ERROR:", exc)


def generate_private_files(private: dict) -> None:
    token = private["token"]
    sources = private["sources"]
    code = [
        "function configurarBridge() {",
        f"  const TOKEN = '{token}';",
        "  const SOURCES = {",
    ]
    items = list(sources.items())
    for i, (key, src) in enumerate(items):
        comma = "," if i < len(items)-1 else ""
        display = str(src.get("displayName", key)).replace("'", "\\'")
        code.append(f"    {key}: {{fileId: '{src['fileId']}', displayName: '{display}'}}{comma}")
    code += [
        "  };",
        "  PropertiesService.getScriptProperties().setProperties({",
        "    UDESA_TOKEN: TOKEN,",
        "    UDESA_SOURCES_JSON: JSON.stringify(SOURCES)",
        "  });",
        "  Logger.log('Bridge configurado con ' + Object.keys(SOURCES).length + ' fuentes.');",
        "}",
    ]
    PRIVATE_GS.parent.mkdir(exist_ok=True)
    PRIVATE_GS.write_text("\n".join(code) + "\n", encoding="utf-8")
    SECRET_TEMPLATE.parent.mkdir(exist_ok=True)
    SECRET_TEMPLATE.write_text(
        "# Después de desplegar Apps Script, pegá la URL /exec abajo.\n"
        "[source_bridge]\n"
        "url = \"PEGAR_URL_DEL_WEB_APP\"\n"
        f"token = \"{token}\"\n",
        encoding="utf-8",
    )


def main() -> int:
    print("=" * 72)
    print("UdeSA Horarios — Preparar Bridge de fuentes")
    print("Autor: Juan Ignacio Gutiérrez Julián")
    print("=" * 72)
    print("Pegá el link de cada archivo PRIVADO de Drive. No cambia permisos ni publica nada.\n")
    sources = {
        "aulas_2026_primavera": {"fileId": ask("1/6 AULAS- PRIMAVERA 2026- FINAL.xlsx"), "displayName": "AULAS- PRIMAVERA 2026- FINAL.xlsx"},
        "cursos_2026_primavera": {"fileId": ask("2/6 CURSOS PRIMAVERA 26.xlsx"), "displayName": "CURSOS PRIMAVERA 26.xlsx"},
        "eventos_2026": {"fileId": ask("3/6 Registro de actividades 2026"), "displayName": "Registro de actividades 2026.xlsx"},
        "plan_master": {"fileId": ask("4/6 udesa_plan_academico_master.xlsx"), "displayName": "udesa_plan_academico_master.xlsx"},
        "catalogo_espacios": {"fileId": ask("5/6 aulas_udesA.xlsx"), "displayName": "aulas_udesA.xlsx"},
        "catalogo_legacy": {"fileId": ask("6/6 Area de Charlas, Camada y codigos.xlsx"), "displayName": "Area de Charlas, Camada y codigos.xlsx"},
    }
    private = {"token": secrets.token_urlsafe(48), "sources": sources}
    PRIVATE_JSON.write_text(json.dumps(private, ensure_ascii=False, indent=2), encoding="utf-8")
    generate_private_files(private)
    print("\nListo. Se generaron archivos PRIVADOS ignorados por Git:")
    print(" -", PRIVATE_JSON.name)
    print(" - apps_script/" + PRIVATE_GS.name)
    print(" - .streamlit/" + SECRET_TEMPLATE.name)
    print("\nSiguiente: copiá CONFIGURACION_PRIVADA.gs al proyecto Apps Script y ejecutá configurarBridge() una vez.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
