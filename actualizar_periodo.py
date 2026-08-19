from __future__ import annotations

import json
from pathlib import Path
from preparar_bridge import PRIVATE_JSON, extract_id, generate_private_files

ROOT = Path(__file__).resolve().parent
CONFIG = ROOT / "fuentes_udesahorarios.json"


def ask_id(label: str) -> str:
    while True:
        v = input(label + "\nLink o ID: ").strip()
        try:
            return extract_id(v)
        except ValueError as e:
            print("ERROR:", e)


def main() -> int:
    if not PRIVATE_JSON.is_file():
        print("Primero ejecutá preparar_bridge.py una vez en esta computadora.")
        return 1
    private = json.loads(PRIVATE_JSON.read_text(encoding="utf-8"))
    cfg = json.loads(CONFIG.read_text(encoding="utf-8"))
    year = int(input("Año del nuevo período (ej. 2027): ").strip())
    term_raw = input("Semestre [O=Otoño / P=Primavera]: ").strip().upper()
    term = "OTONO" if term_raw.startswith("O") else "PRIMAVERA"
    term_label = "OTOÑO" if term == "OTONO" else "PRIMAVERA"
    pid = f"{year}_{term}"
    cohort = int(input("Camada que ingresa ese año (ej. 39): ").strip())
    aulas_id = ask_id(f"AULAS {term_label.title()} {year}")
    cursos_id = ask_id(f"CURSOS {term_label.title()} {year}")
    event_key = f"eventos_{year}"
    if event_key not in private["sources"]:
        eventos_id = ask_id(f"Registro de actividades {year}")
        private["sources"][event_key] = {"fileId": eventos_id, "displayName": f"Registro de actividades {year}.xlsx"}
    aulas_key = f"aulas_{year}_{term.lower()}"
    cursos_key = f"cursos_{year}_{term.lower()}"
    private["sources"][aulas_key] = {"fileId": aulas_id, "displayName": f"AULAS {term_label} {year}.xlsx"}
    private["sources"][cursos_key] = {"fileId": cursos_id, "displayName": f"CURSOS {term_label} {year}.xlsx"}

    cfg["periodo_activo"] = pid
    cfg["cohorte_ancla"] = {"anio": year, "camada_ingreso": cohort}
    cfg.setdefault("fuentes_anuales", {}).setdefault("eventos", {})[str(year)] = f"data/Registro de actividades {year}.xlsx"
    cfg.setdefault("periodos", {})[pid] = {
        "anio": year,
        "semestre": term,
        "fecha_inicio": None,
        "fecha_fin": None,
        "fuentes": {
            "cursos": f"data/CURSOS {term_label} {year}.xlsx",
            "aulas": f"data/AULAS {term_label} {year}.xlsx",
        },
    }
    sync_sources = cfg.setdefault("sincronizacion", {}).setdefault("fuentes", {})
    sync_sources[f"aulas_{pid.lower()}"] = {"bridge_key": aulas_key, "destino": f"data/AULAS {term_label} {year}.xlsx", "label": "AULAS", "periodos": [pid]}
    sync_sources[f"cursos_{pid.lower()}"] = {"bridge_key": cursos_key, "destino": f"data/CURSOS {term_label} {year}.xlsx", "label": "CURSOS", "periodos": [pid]}
    sync_sources[f"eventos_{year}"] = {"bridge_key": event_key, "destino": f"data/Registro de actividades {year}.xlsx", "label": "Eventos", "periodos": [pid]}
    # Las fuentes transversales deben aplicar al nuevo período también.
    for k in ("plan_master", "catalogo_espacios", "catalogo_legacy"):
        if k in sync_sources:
            periods = sync_sources[k].setdefault("periodos", [])
            if pid not in periods:
                periods.append(pid)

    PRIVATE_JSON.write_text(json.dumps(private, ensure_ascii=False, indent=2), encoding="utf-8")
    CONFIG.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    generate_private_files(private)
    print("\nPeríodo actualizado:", pid)
    print("1. En Apps Script, reemplazá/pegá CONFIGURACION_PRIVADA.gs y ejecutá configurarBridge().")
    print("2. Subí a GitHub solamente el cambio de fuentes_udesahorarios.json (no los archivos privados).")
    print("3. Streamlit redeploya automáticamente desde GitHub.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
