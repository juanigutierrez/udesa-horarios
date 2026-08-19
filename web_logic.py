from __future__ import annotations

import datetime as dt
import re
from typing import Any, Dict, Iterable, List, Sequence, Tuple

WEEKDAY_LABELS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]


def clamp_default_window(period, today: dt.date | None = None, days: int = 14) -> Tuple[dt.date, dt.date]:
    today = today or dt.date.today()
    if today < period.start or today > period.end:
        start = period.start
    else:
        start = today
    end = min(period.end, start + dt.timedelta(days=max(days - 1, 0)))
    return start, end


def normalize_date_range(value: Any) -> Tuple[dt.date, dt.date]:
    if isinstance(value, dt.date):
        return value, value
    if isinstance(value, (list, tuple)):
        if not value:
            raise ValueError("Elegí una fecha o un rango de fechas.")
        if len(value) == 1:
            return value[0], value[0]
        return value[0], value[1]
    raise ValueError("No se pudo interpretar el rango de fechas.")


def dates_between(start: dt.date, end: dt.date, weekdays_only: bool = True) -> List[dt.date]:
    if end < start:
        start, end = end, start
    rows: List[dt.date] = []
    current = start
    while current <= end:
        if not weekdays_only or current.weekday() < 5:
            rows.append(current)
        current += dt.timedelta(days=1)
    return rows


def parse_extra_codes(text: str | None) -> List[str]:
    if not text:
        return []
    tokens = re.split(r"[\s,;|]+", text.strip())
    out: List[str] = []
    for token in tokens:
        code = token.strip().upper()
        if code and code not in out:
            out.append(code)
    return out


def slot_label(slot: str, slot_map: Dict[str, Sequence[str]]) -> str:
    start, end = slot_map[str(slot)]
    return f"Slot {slot} · {start}–{end}"


def slot_options(slot_map: Dict[str, Sequence[str]]) -> List[Tuple[str, str]]:
    keys = sorted(slot_map, key=lambda x: int(x) if str(x).isdigit() else 999)
    return [(slot_label(str(k), slot_map), str(k)) for k in keys]


def physical_sites(engine) -> List[str]:
    values = sorted({str(row.get("Sede", "")).strip() for row in engine.catalog_rows if str(row.get("Sede", "")).strip()})
    return values


def room_choices(engine, site: str | None = None, include_non_classrooms: bool = True) -> List[Tuple[str, str]]:
    choices: List[Tuple[str, str]] = []
    for row in engine.catalog_rows:
        room = str(row.get("Espacio_ID", "")).strip()
        if not room:
            continue
        if site and str(row.get("Sede", "")).strip() != site:
            continue
        if not include_non_classrooms and str(row.get("Tipo", "")).strip() != "Aula":
            continue
        name = str(row.get("Nombre", room)).strip() or room
        cap = row.get("Capacidad", "")
        suffix = f" · {name}" if name and name.upper() != room.upper() else ""
        if cap not in (None, ""):
            suffix += f" · cap. {cap}"
        choices.append((f"{room}{suffix}", room))
    return sorted(choices, key=lambda x: x[0].upper())


def population_payload(
    career: str,
    cohorts: Sequence[int] | None = None,
    academic_site: str | None = None,
    variant: str | None = None,
    include_optatives: bool = False,
    extra_codes: Sequence[str] | None = None,
) -> Dict[str, Any]:
    return {
        "Carrera": career,
        "Camadas": list(cohorts) if cohorts else "TODAS",
        "Sede": academic_site,
        "Variante": variant,
        "Incluir_optativas": include_optatives,
        "Códigos_extra": list(extra_codes or []),
    }


def validate_populations(service, populations: Sequence[Dict[str, Any]]) -> Tuple[bool, List[Dict[str, Any]]]:
    resolved: List[Dict[str, Any]] = []
    ok = True
    for pop in populations:
        result = service.resolver_poblacion_simple(
            pop.get("Carrera"),
            camadas=pop.get("Camadas", "TODAS"),
            sede=pop.get("Sede"),
            variante=pop.get("Variante"),
            incluir_optativas=pop.get("Incluir_optativas"),
            codigos_extra=pop.get("Códigos_extra", []),
        )
        resolved.append(result)
        if result.get("Estado") != "OK":
            ok = False
    return ok, resolved


def best_weekly_rows(rows: Sequence[Dict[str, Any]], limit: int = 10) -> List[Dict[str, Any]]:
    day_order = {name: i for i, name in enumerate(WEEKDAY_LABELS)}
    return sorted(
        rows,
        key=lambda r: (
            int(r.get("Alumnos ocupados", 0) or 0),
            int(r.get("Materias afectadas", 0) or 0),
            day_order.get(str(r.get("Día", "")), 99),
            int(r.get("Slot", 999) or 999),
        ),
    )[:limit]


def weekly_grid(rows: Sequence[Dict[str, Any]]) -> List[Dict[str, Any]]:
    by_day: Dict[str, Dict[str, Any]] = {}
    for row in rows:
        day = str(row.get("Día", ""))
        record = by_day.setdefault(day, {"Día": day})
        slot = str(row.get("Slot", ""))
        record[f"Slot {slot}"] = int(row.get("Alumnos ocupados", 0) or 0)
    return [by_day[d] for d in WEEKDAY_LABELS[:5] if d in by_day]


def status_icon(status: str) -> str:
    return {
        "LIBRE": "🟢",
        "OCUPADA": "🔴",
        "REVISAR": "🟠",
        "CONFLICTO": "🟣",
    }.get(str(status).upper(), "⚪")


def period_label_map(registry) -> Dict[str, str]:
    result: Dict[str, str] = {}
    for pid in registry.list_period_ids():
        raw = registry.period_defs[pid]
        term = str(raw.get("semestre", "")).strip().capitalize()
        result[f"{term} {raw.get('anio')}"] = pid
    return result
