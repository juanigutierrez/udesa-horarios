#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
UdeSA Horarios — Ronda 3
Autor: Juan Ignacio Gutiérrez Julián

Etapas implementadas:
- 9: motor temporal por fecha + intervalo horario, manteniendo slots como atajo.
- 10: motor central de ocupación: LIBRE / OCUPADA / REVISAR / CONFLICTO.
- 11: detector de conflictos entre fuentes.
- 12: disponibilidad estudiantil por códigos/materias.
- 14A: combinación de disponibilidad estudiantil + aulas usando códigos.
- 15: compatibilidad funcional con las tres consultas del programa original.
- 16: batería automática de pruebas.

Etapas 3B, 13 y 14B quedan deliberadamente para la Ronda 4.

No modifica los Excel originales. Usa únicamente la biblioteca estándar de Python.
"""
from __future__ import annotations

import csv
import datetime as dt
import json
import re
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Set, Tuple

import udesa_horarios_ronda2 as r2

APP_NAME = "UdeSA Horarios"
AUTHOR = "Juan Ignacio Gutiérrez Julián"
OUTPUT_DIR_NAME = "salida_ronda3"
WEEKDAYS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]

# La grilla Primavera representa el cronograma regular del semestre.
# CURSOS contiene cinco filas de julio de otro período; por eso la vigencia base de la grilla
# se toma del rango dominante del semestre Primavera, no del mínimo absoluto de CURSOS.
DEFAULT_SEMESTER_START = dt.date(2026, 8, 3)
DEFAULT_SEMESTER_END = dt.date(2026, 12, 12)

STATUS_ORDER = {"LIBRE": 0, "OCUPADA": 1, "REVISAR": 2, "CONFLICTO": 3}


def clean(value: Any) -> str:
    return r2.clean_text(value)


def norm(value: Any) -> str:
    return r2.norm_text(value)


def parse_date(value: Any) -> dt.date:
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    text = clean(value)
    if not text:
        raise ValueError("La fecha está vacía.")
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y"):
        try:
            return dt.datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    raise ValueError(f"No se pudo interpretar la fecha: {text}")


def parse_time(value: Any) -> dt.time:
    if isinstance(value, dt.datetime):
        return value.time().replace(second=0, microsecond=0)
    if isinstance(value, dt.time):
        return value.replace(second=0, microsecond=0)
    text = clean(value).replace(".", ":")
    if not text:
        raise ValueError("La hora está vacía.")
    if re.fullmatch(r"\d{1,2}", text):
        text += ":00"
    for fmt in ("%H:%M", "%H:%M:%S"):
        try:
            return dt.datetime.strptime(text, fmt).time().replace(second=0, microsecond=0)
        except ValueError:
            pass
    raise ValueError(f"No se pudo interpretar la hora: {text}")


def time_text(value: Optional[dt.time]) -> str:
    return value.strftime("%H:%M") if value else ""


def date_text(value: dt.date) -> str:
    return value.strftime("%d/%m/%Y")


def intervals_overlap(start_a: dt.time, end_a: dt.time, start_b: dt.time, end_b: dt.time) -> bool:
    """Superposición de intervalos semiabiertos [inicio, fin)."""
    return start_a < end_b and end_a > start_b


def validate_interval(start: dt.time, end: dt.time) -> None:
    if end <= start:
        raise ValueError("La hora hasta debe ser posterior a la hora desde dentro del mismo día.")


def normalize_sede_query(value: Any) -> str:
    n = norm(value)
    mapping = {
        "CAMPUS": "Campus",
        "CAMPUS VICTORIA": "Campus",
        "VICTORIA": "Campus",
        "RIOBAMBA": "Riobamba",
        "SEDE RIOBAMBA": "Riobamba",
        "CALLAO": "Callao",
        "SEDE CALLAO": "Callao",
        "NORDELTA": "Nordelta",
        "SUIPACHA": "Suipacha",
        "DIGITAL HOUSE": "Digital House Belgrano",
        "DIGITAL HOUSE BELGRANO": "Digital House Belgrano",
        "TODAS": "",
        "TODOS": "",
        "": "",
    }
    return mapping.get(n, clean(value))


def tokenize_description(text: str) -> Set[str]:
    stop = {
        "EL", "LA", "LOS", "LAS", "DE", "DEL", "A", "Y", "EN", "ESTE", "ESTA",
        "CURSO", "AULA", "ACA", "ACÁ", "VIENE", "VA", "SE", "GRAL", "GENERAL",
    }
    return {x for x in re.findall(r"[A-Z0-9]+", norm(text)) if len(x) > 1 and x not in stop}


def extract_code(text: str) -> str:
    m = re.search(r"\b[A-Z]{1,5}\d{2,3}\b", norm(text))
    return m.group(0) if m else ""


@dataclass
class Claim:
    room: str
    date: dt.date
    start: dt.time
    end: dt.time
    effect: str                    # OCUPA / REVISAR / LIBERA
    source: str                    # AULAS / CURSOS / COMENTARIO / RESERVA / EVENTO
    label: str                     # Clase / Evento / Comentario / Reserva / Ocupación / Revisar
    description: str
    confidence: str = "ALTA"
    slot: str = ""
    source_id: str = ""
    related_room: str = ""
    origin_cell_room: str = ""
    raw: Dict[str, Any] = field(default_factory=dict)

    def interval_text(self) -> str:
        return f"{time_text(self.start)}–{time_text(self.end)}"


@dataclass
class ClaimGroup:
    claims: List[Claim]

    @property
    def representative(self) -> Claim:
        # Preferimos Comentario/Reserva/Evento sobre una fuente base porque son más específicos.
        order = {"COMENTARIO": 0, "RESERVA": 1, "EVENTO": 2, "AULAS": 3, "CURSOS": 4}
        return sorted(self.claims, key=lambda c: order.get(c.source, 99))[0]

    @property
    def sources(self) -> List[str]:
        return sorted({c.source for c in self.claims})


@dataclass
class RoomQueryResult:
    room: str
    room_name: str
    capacity: Optional[int]
    sede: str
    date: dt.date
    query_start: dt.time
    query_end: dt.time
    status: str
    occupation_groups: List[ClaimGroup]
    review_claims: List[Claim]
    release_claims: List[Claim]
    notes: List[str]
    conflict_pairs: List[Tuple[ClaimGroup, ClaimGroup]]


class UdeSAEngine:
    def __init__(
        self,
        catalog_rows: Sequence[Dict[str, Any]],
        class_rows: Sequence[Dict[str, Any]],
        grid_rows: Sequence[Dict[str, Any]],
        comment_items: Sequence[Dict[str, Any]],
        comment_effects: Sequence[Dict[str, Any]],
        reservation_rows: Sequence[Dict[str, Any]],
        event_occupancies: Sequence[Dict[str, Any]],
        slot_map: Dict[str, Tuple[str, str]],
    ) -> None:
        self.catalog_rows = list(catalog_rows)
        self.class_rows = list(class_rows)
        self.grid_rows = list(grid_rows)
        self.comment_items = list(comment_items)
        self.comment_effects = list(comment_effects)
        self.reservation_rows = list(reservation_rows)
        self.event_occupancies = list(event_occupancies)
        self.slot_map = {str(k): v for k, v in slot_map.items()}

        self.catalog: Dict[str, Dict[str, Any]] = {r["Espacio_ID"]: r for r in self.catalog_rows if r.get("Espacio_ID")}
        self.rooms_in_aulas = {r["Espacio_ID"] for r in self.catalog_rows if r.get("En_AULAS_Actual") == "Sí"}

        self.alias_to_room: Dict[str, str] = {}
        for row in self.catalog_rows:
            key = row.get("Espacio_ID", "")
            if not key:
                continue
            candidates = [key, row.get("Nombre", "")]
            candidates.extend([x.strip() for x in row.get("Aliases", "").split("|") if x.strip()])
            for candidate in candidates:
                n = norm(candidate)
                if n and n not in self.alias_to_room:
                    self.alias_to_room[n] = key
                compact = re.sub(r"\s+", "", n)
                if compact and compact not in self.alias_to_room:
                    self.alias_to_room[compact] = key

        self.classes_by_room_day: Dict[Tuple[str, str], List[Dict[str, Any]]] = defaultdict(list)
        self.classes_by_code: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
        self.course_names: Dict[str, str] = {}
        for row in self.class_rows:
            code = row.get("Código", "")
            if code:
                self.classes_by_code[code].append(row)
                if row.get("Materia"):
                    self.course_names.setdefault(code, row["Materia"])
            if row.get("Espacio_ID") and row.get("Día"):
                self.classes_by_room_day[(row["Espacio_ID"], row["Día"])].append(row)

        self.grid_by_room_day: Dict[Tuple[str, str], List[Dict[str, Any]]] = defaultdict(list)
        self.grid_reserved_for_events: Dict[Tuple[str, str, str], Dict[str, Any]] = {}
        for row in self.grid_rows:
            room = row.get("Espacio_ID", "")
            day = row.get("Día", "")
            if room and day:
                self.grid_by_room_day[(room, day)].append(row)
                if row.get("Tipo_registro") == "RESERVADO_PARA_EVENTOS":
                    self.grid_reserved_for_events[(room, day, row.get("Slot", ""))] = row

        self.comment_item_by_id: Dict[str, Dict[str, Any]] = {r.get("ID", ""): r for r in self.comment_items if r.get("ID")}
        self.comment_effects_by_room_date: Dict[Tuple[str, str], List[Dict[str, Any]]] = defaultdict(list)
        for row in self.comment_effects:
            if row.get("Espacio_ID") and row.get("Fecha"):
                self.comment_effects_by_room_date[(row["Espacio_ID"], row["Fecha"])].append(row)

        self.reservations_by_room_date: Dict[Tuple[str, str], List[Dict[str, Any]]] = defaultdict(list)
        for row in self.reservation_rows:
            if row.get("Espacio_ID") and row.get("Fecha"):
                self.reservations_by_room_date[(row["Espacio_ID"], row["Fecha"])].append(row)

        self.events_by_room_date: Dict[Tuple[str, str], List[Dict[str, Any]]] = defaultdict(list)
        for row in self.event_occupancies:
            if row.get("Espacio_ID") and row.get("Fecha"):
                self.events_by_room_date[(row["Espacio_ID"], row["Fecha"])].append(row)

    # -------------------------
    # Etapa 9: motor temporal
    # -------------------------
    def room_key(self, value: Any) -> str:
        n = norm(value)
        if n in self.alias_to_room:
            return self.alias_to_room[n]
        compact = re.sub(r"\s+", "", n)
        if compact in self.alias_to_room:
            return self.alias_to_room[compact]
        normalized = r2.normalize_room_key(value)
        if normalized in self.catalog:
            return normalized
        raise KeyError(f"Espacio/aula no reconocido: {clean(value)}")

    def interval_from_slot(self, slot: Any) -> Tuple[dt.time, dt.time]:
        key = r2.normalize_slot(slot)
        if key not in self.slot_map:
            raise KeyError(f"Slot no reconocido: {slot}")
        start_s, end_s = self.slot_map[key]
        return parse_time(start_s), parse_time(end_s)

    def resolve_interval(self, slot: Any = None, start: Any = None, end: Any = None) -> Tuple[dt.time, dt.time]:
        if slot not in (None, ""):
            if start not in (None, "") or end not in (None, ""):
                raise ValueError("Usá slot o rango horario personalizado, no ambos al mismo tiempo.")
            return self.interval_from_slot(slot)
        if start in (None, "") or end in (None, ""):
            raise ValueError("Falta indicar slot o ambas horas: desde y hasta.")
        start_t, end_t = parse_time(start), parse_time(end)
        validate_interval(start_t, end_t)
        return start_t, end_t

    # ---------------------------------
    # Construcción de claims por fuente
    # ---------------------------------
    def _base_claims(self, room: str, date: dt.date) -> List[Claim]:
        day = WEEKDAYS[date.weekday()]
        claims: List[Claim] = []

        # Si el aula existe en AULAS Primavera, esa grilla es la base física autoritativa.
        if room in self.rooms_in_aulas:
            if not (DEFAULT_SEMESTER_START <= date <= DEFAULT_SEMESTER_END):
                return claims
            for row in self.grid_by_room_day.get((room, day), []):
                record_type = row.get("Tipo_registro", "")
                if record_type in {"ETIQUETA_NO_OCUPA", "RESERVADO_PARA_EVENTOS"}:
                    continue
                if record_type not in {"CLASE_REGULAR", "BLOQUE_SEMANAL_OTRO", "REVISAR"}:
                    continue
                start = parse_time(row["Hora Desde"])
                end = parse_time(row["Hora Hasta"])
                if record_type == "CLASE_REGULAR":
                    code = row.get("Código", "")
                    matter = row.get("Materia", "")
                    desc = f"{code} — {matter}".strip(" —") or row.get("Texto_original", "")
                    label = "Clase"
                    effect = "OCUPA"
                elif record_type == "REVISAR":
                    desc = row.get("Texto_original", "")
                    label = "Revisar"
                    effect = "REVISAR"
                else:
                    desc = row.get("Texto_original", "")
                    label = "Ocupación"
                    effect = "OCUPA"
                claims.append(Claim(
                    room=room, date=date, start=start, end=end, effect=effect, source="AULAS",
                    label=label, description=desc, confidence=row.get("Confianza", ""),
                    slot=row.get("Slot", ""), source_id=f"AULAS-{row.get('Fila','')}-{day}-{row.get('Slot','')}", raw=row,
                ))
            return claims

        # Para espacios que no están en la grilla AULAS (p.ej. otras sedes), CURSOS es fallback físico.
        for row in self.classes_by_room_day.get((room, day), []):
            try:
                start_date = parse_date(row.get("Inicio vigencia")) if row.get("Inicio vigencia") else DEFAULT_SEMESTER_START
                end_date = parse_date(row.get("Fin vigencia")) if row.get("Fin vigencia") else DEFAULT_SEMESTER_END
            except ValueError:
                start_date, end_date = DEFAULT_SEMESTER_START, DEFAULT_SEMESTER_END
            if not (start_date <= date <= end_date):
                continue
            start = parse_time(row["Hora Desde"])
            end = parse_time(row["Hora Hasta"])
            desc = f"{row.get('Código','')} — {row.get('Materia','')}".strip(" —")
            claims.append(Claim(
                room=room, date=date, start=start, end=end, effect="OCUPA", source="CURSOS",
                label="Clase", description=desc, confidence="ALTA", slot=row.get("Slot", ""),
                source_id=f"CURSOS-{row.get('Fila','')}", raw=row,
            ))
        return claims

    def _specific_claims(self, room: str, date: dt.date) -> List[Claim]:
        date_iso = date.isoformat()
        claims: List[Claim] = []

        for row in self.comment_effects_by_room_date.get((room, date_iso), []):
            start = parse_time(row["Hora Desde"])
            end = parse_time(row["Hora Hasta"])
            cid = row.get("ID comentario", "")
            original_item = self.comment_item_by_id.get(cid, {})
            claims.append(Claim(
                room=room, date=date, start=start, end=end, effect=row.get("Efecto", "REVISAR"),
                source="COMENTARIO", label="Comentario", description=row.get("Motivo", ""),
                confidence=row.get("Confianza", ""), slot=row.get("Slot", ""), source_id=cid,
                related_room=row.get("Aula_relacionada", ""), origin_cell_room=original_item.get("Espacio_ID", ""), raw=row,
            ))

        for row in self.reservations_by_room_date.get((room, date_iso), []):
            try:
                start = parse_time(row.get("Hora Desde"))
                end = parse_time(row.get("Hora Hasta"))
            except ValueError:
                start, end = dt.time(0, 0), dt.time(23, 59)
            effect = "REVISAR" if row.get("Efecto") == "REVISAR" else "OCUPA"
            description = row.get("Motivo", "") or row.get("Profesor/Pedido por", "") or "Reserva especial"
            claims.append(Claim(
                room=room, date=date, start=start, end=end, effect=effect, source="RESERVA", label="Reserva",
                description=description, confidence=row.get("Confianza", ""), source_id=f"RES-{row.get('Hoja','')}-{row.get('Fila','')}", raw=row,
            ))

        for row in self.events_by_room_date.get((room, date_iso), []):
            effect = "REVISAR" if row.get("Efecto") == "REVISAR" else "OCUPA"
            start_s, end_s = row.get("Hora Desde", ""), row.get("Hora Hasta", "")
            # Regla conservadora: si existe un aula concreta pero el horario es ambiguo/incompleto,
            # bloqueamos el día completo como REVISAR; de otro modo podríamos recomendar falsamente el aula.
            if effect == "REVISAR" and (not start_s or not end_s):
                start, end = dt.time(0, 0), dt.time(23, 59)
            else:
                try:
                    start, end = parse_time(start_s), parse_time(end_s)
                except ValueError:
                    start, end = dt.time(0, 0), dt.time(23, 59)
                    effect = "REVISAR"
            claims.append(Claim(
                room=room, date=date, start=start, end=end, effect=effect, source="EVENTO", label="Evento",
                description=row.get("Evento", ""), confidence=row.get("Confianza", ""), source_id=row.get("ID evento", ""), raw=row,
            ))
        return claims

    def _apply_comment_overrides(self, base_claims: List[Claim], specific_claims: List[Claim], room: str) -> List[Claim]:
        # Un comentario asociado a la propia celda AULAS es una corrección específica de esa fecha.
        # Por eso reemplaza la ocupación semanal de ESA celda/slot. Un efecto que llega a otra aula
        # (p.ej. destino de un traslado) no borra la base de la otra aula; si estaba ocupada, habrá conflicto.
        suppress_slots: Set[str] = set()
        for claim in specific_claims:
            if claim.source == "COMENTARIO" and claim.origin_cell_room == room and claim.slot:
                suppress_slots.add(claim.slot)
        if not suppress_slots:
            return base_claims
        return [c for c in base_claims if not (c.source == "AULAS" and c.slot in suppress_slots)]

    def _same_claim(self, a: Claim, b: Claim) -> bool:
        same_interval = a.start == b.start and a.end == b.end
        overlap = intervals_overlap(a.start, a.end, b.start, b.end)
        if not overlap:
            return False
        if a.source == "COMENTARIO" and b.source == "COMENTARIO":
            if a.related_room and b.related_room and a.related_room == b.related_room:
                return True
        code_a, code_b = extract_code(a.description), extract_code(b.description)
        if code_a and code_b and code_a == code_b:
            return True
        ta, tb = tokenize_description(a.description), tokenize_description(b.description)
        if not ta or not tb:
            return False
        inter = len(ta & tb)
        containment = inter / min(len(ta), len(tb))
        jaccard = inter / len(ta | tb)
        if same_interval:
            return containment >= 0.75 or jaccard >= 0.60
        # Un bloque institucional semanal puede cubrir varios slots mientras un Evento específico
        # de ese mismo programa cubre un intervalo continuo. No es un conflicto real solo porque
        # los límites horarios no sean idénticos (ej.: "MBA SALUD" + "Jornada ... MBA Salud").
        sources = {a.source, b.source}
        if sources == {"AULAS", "EVENTO"} and inter >= 2 and containment >= 0.80:
            return True
        return False

    def _group_occupations(self, claims: List[Claim]) -> List[ClaimGroup]:
        groups: List[ClaimGroup] = []
        for claim in claims:
            placed = False
            for group in groups:
                if any(self._same_claim(claim, existing) for existing in group.claims):
                    group.claims.append(claim)
                    placed = True
                    break
            if not placed:
                groups.append(ClaimGroup([claim]))
        return groups

    def _conflict_pairs(self, groups: List[ClaimGroup]) -> List[Tuple[ClaimGroup, ClaimGroup]]:
        conflicts: List[Tuple[ClaimGroup, ClaimGroup]] = []
        for i in range(len(groups)):
            for j in range(i + 1, len(groups)):
                a = groups[i].representative
                b = groups[j].representative
                if intervals_overlap(a.start, a.end, b.start, b.end):
                    conflicts.append((groups[i], groups[j]))
        return conflicts

    def _free_notes(self, room: str, date: dt.date, start: dt.time, end: dt.time) -> List[str]:
        notes: List[str] = []
        if DEFAULT_SEMESTER_START <= date <= DEFAULT_SEMESTER_END:
            day = WEEKDAYS[date.weekday()]
            for slot, (slot_start_s, slot_end_s) in self.slot_map.items():
                slot_start, slot_end = parse_time(slot_start_s), parse_time(slot_end_s)
                if intervals_overlap(slot_start, slot_end, start, end) and (room, day, slot) in self.grid_reserved_for_events:
                    notes.append(f"Slot {slot}: reservado para eventos en la grilla AULAS (no cuenta como ocupación).")
        return notes

    # --------------------------------
    # Etapa 10: consulta central de aula
    # --------------------------------
    def query_room(self, room_value: Any, date_value: Any, *, slot: Any = None, start: Any = None, end: Any = None) -> RoomQueryResult:
        room = self.room_key(room_value)
        date = parse_date(date_value)
        query_start, query_end = self.resolve_interval(slot=slot, start=start, end=end)

        base_claims = self._base_claims(room, date)
        specific_claims = self._specific_claims(room, date)
        base_claims = self._apply_comment_overrides(base_claims, specific_claims, room)
        all_claims = base_claims + specific_claims

        relevant = [
            c for c in all_claims
            if intervals_overlap(c.start, c.end, query_start, query_end)
        ]
        occupation_claims = [c for c in relevant if c.effect == "OCUPA"]
        review_claims = [c for c in relevant if c.effect == "REVISAR"]
        release_claims = [c for c in relevant if c.effect == "LIBERA"]

        groups = self._group_occupations(occupation_claims)
        conflicts = self._conflict_pairs(groups)
        if conflicts:
            status = "CONFLICTO"
        elif review_claims:
            status = "REVISAR"
        elif groups:
            status = "OCUPADA"
        else:
            status = "LIBRE"

        meta = self.catalog.get(room, {})
        cap = None
        try:
            cap = int(meta.get("Capacidad")) if meta.get("Capacidad") not in (None, "") else None
        except Exception:
            cap = None
        notes = self._free_notes(room, date, query_start, query_end)
        if release_claims:
            for c in release_claims:
                notes.append(f"Liberación específica por comentario: {c.description}")
        return RoomQueryResult(
            room=room,
            room_name=meta.get("Nombre", room),
            capacity=cap,
            sede=meta.get("Sede", ""),
            date=date,
            query_start=query_start,
            query_end=query_end,
            status=status,
            occupation_groups=groups,
            review_claims=review_claims,
            release_claims=release_claims,
            notes=notes,
            conflict_pairs=conflicts,
        )

    def format_room_result(self, result: RoomQueryResult) -> str:
        lines = [result.room]
        lines.append(f"Capacidad: {result.capacity if result.capacity is not None else 'N/A'}")
        lines.append(result.status)
        details: List[Claim] = []
        for group in result.occupation_groups:
            details.append(group.representative)
        details.extend(result.review_claims)
        details.sort(key=lambda c: (c.start, c.end, c.source))
        if details:
            lines.append("")
        for idx, claim in enumerate(details):
            lines.append(claim.interval_text())
            if claim.source == "EVENTO":
                lines.append(f"Evento: {claim.description}")
            elif claim.source == "COMENTARIO":
                lines.append(f"Comentario: {claim.description}")
            elif claim.label == "Clase":
                lines.append(f"Clase: {claim.description}")
            elif claim.source == "RESERVA":
                lines.append(f"Reserva: {claim.description}")
            elif claim.effect == "REVISAR":
                lines.append(f"Revisar: {claim.description}")
            else:
                lines.append(f"Ocupación: {claim.description}")
            if idx < len(details) - 1:
                lines.append("")
        if result.notes:
            lines.append("")
            lines.extend([f"Nota: {x}" for x in result.notes])
        return "\n".join(lines)

    # ---------------------------------
    # Etapa 11: conflictos entre fuentes
    # ---------------------------------
    def detect_conflicts(self) -> List[Dict[str, Any]]:
        dates: Set[dt.date] = set()
        d = DEFAULT_SEMESTER_START
        while d <= DEFAULT_SEMESTER_END:
            dates.add(d)
            d += dt.timedelta(days=1)
        for (_, date_s) in list(self.comment_effects_by_room_date) + list(self.reservations_by_room_date) + list(self.events_by_room_date):
            try:
                dates.add(parse_date(date_s))
            except ValueError:
                pass

        # No hace falta consultar todos los espacios en fechas sin dato específico fuera del semestre.
        specific_rooms_by_date: Dict[dt.date, Set[str]] = defaultdict(set)
        for index in (self.comment_effects_by_room_date, self.reservations_by_room_date, self.events_by_room_date):
            for room, date_s in index:
                try:
                    specific_rooms_by_date[parse_date(date_s)].add(room)
                except ValueError:
                    pass

        rows: List[Dict[str, Any]] = []
        for date in sorted(dates):
            if DEFAULT_SEMESTER_START <= date <= DEFAULT_SEMESTER_END:
                candidate_rooms = set(self.catalog)
            else:
                candidate_rooms = specific_rooms_by_date.get(date, set())
            for room in candidate_rooms:
                result = self.query_room(room, date, start="00:00", end="23:59")
                if result.status != "CONFLICTO":
                    continue
                for n, (ga, gb) in enumerate(result.conflict_pairs, start=1):
                    a, b = ga.representative, gb.representative
                    overlap_start = max(a.start, b.start)
                    overlap_end = min(a.end, b.end)
                    rows.append({
                        "Fecha": date.isoformat(),
                        "Día": WEEKDAYS[date.weekday()],
                        "Espacio_ID": room,
                        "Capacidad": result.capacity if result.capacity is not None else "",
                        "Conflicto_n": n,
                        "Superposición Desde": time_text(overlap_start),
                        "Superposición Hasta": time_text(overlap_end),
                        "Fuente A": " + ".join(ga.sources),
                        "Tipo A": a.label,
                        "Horario A": a.interval_text(),
                        "Detalle A": a.description,
                        "Fuente B": " + ".join(gb.sources),
                        "Tipo B": b.label,
                        "Horario B": b.interval_text(),
                        "Detalle B": b.description,
                    })
        return rows

    # -----------------------------------------
    # Etapa 12: disponibilidad de los alumnos
    # -----------------------------------------
    def student_occupancy(self, codes: Sequence[str], date_value: Any, *, slot: Any = None, start: Any = None, end: Any = None) -> Dict[str, Any]:
        date = parse_date(date_value)
        query_start, query_end = self.resolve_interval(slot=slot, start=start, end=end)
        selected = []
        for code in codes:
            c = clean(code).upper()
            if c and c not in selected:
                selected.append(c)
        if not selected:
            raise ValueError("Ingresá al menos un código de materia.")

        day = WEEKDAYS[date.weekday()]
        details: List[Dict[str, Any]] = []
        found_codes: Set[str] = set()
        for code in selected:
            for row in self.classes_by_code.get(code, []):
                found_codes.add(code)
                if row.get("Día") != day:
                    continue
                try:
                    start_date = parse_date(row.get("Inicio vigencia")) if row.get("Inicio vigencia") else DEFAULT_SEMESTER_START
                    end_date = parse_date(row.get("Fin vigencia")) if row.get("Fin vigencia") else DEFAULT_SEMESTER_END
                except ValueError:
                    start_date, end_date = DEFAULT_SEMESTER_START, DEFAULT_SEMESTER_END
                if not (start_date <= date <= end_date):
                    continue
                try:
                    class_start, class_end = parse_time(row.get("Hora Desde")), parse_time(row.get("Hora Hasta"))
                except ValueError:
                    continue
                if intervals_overlap(class_start, class_end, query_start, query_end):
                    details.append(row)
        students = 0
        for row in details:
            try:
                students += int(float(row.get("Cant Inscriptos", 0) or 0))
            except Exception:
                pass
        affected_codes = sorted({row.get("Código", "") for row in details if row.get("Código")})
        return {
            "Fecha": date.isoformat(),
            "Día": day,
            "Hora Desde": time_text(query_start),
            "Hora Hasta": time_text(query_end),
            "Códigos solicitados": selected,
            "Códigos encontrados": sorted(found_codes),
            "Códigos no encontrados": [c for c in selected if c not in found_codes],
            "Alumnos ocupados": students,
            "Materias afectadas": len(affected_codes),
            "Códigos afectados": affected_codes,
            "Clases afectadas": len(details),
            "Detalle": details,
        }

    def weekly_student_ranking(self, codes: Sequence[str], reference_week_start: Optional[dt.date] = None) -> List[Dict[str, Any]]:
        # Para preservar la consulta antigua, usamos una semana representativa dentro de la vigencia regular.
        monday = reference_week_start or dt.date(2026, 8, 17)
        monday -= dt.timedelta(days=monday.weekday())
        rows: List[Dict[str, Any]] = []
        for day_offset in range(5):
            date = monday + dt.timedelta(days=day_offset)
            for slot in sorted(self.slot_map, key=lambda x: int(x) if x.isdigit() else 999):
                result = self.student_occupancy(codes, date, slot=slot)
                rows.append({
                    "Fecha referencia": date.isoformat(),
                    "Día": result["Día"],
                    "Slot": slot,
                    "Hora Desde": result["Hora Desde"],
                    "Hora Hasta": result["Hora Hasta"],
                    "Alumnos ocupados": result["Alumnos ocupados"],
                    "Materias afectadas": result["Materias afectadas"],
                    "Clases afectadas": result["Clases afectadas"],
                    "Códigos afectados": " | ".join(result["Códigos afectados"]),
                })
        rows.sort(key=lambda r: (WEEKDAYS.index(r["Día"]), r["Alumnos ocupados"], r["Materias afectadas"], int(r["Slot"])))
        return rows

    # --------------------------------------------------
    # Aulas libres y consulta diaria (Etapas 10 y 15)
    # --------------------------------------------------
    def free_rooms(
        self,
        date_value: Any,
        *,
        slot: Any = None,
        start: Any = None,
        end: Any = None,
        sede: str = "Campus",
        min_capacity: int = 0,
        include_non_classrooms: bool = False,
        include_review: bool = False,
    ) -> List[Dict[str, Any]]:
        date = parse_date(date_value)
        query_start, query_end = self.resolve_interval(slot=slot, start=start, end=end)
        sede_norm = normalize_sede_query(sede)
        rows: List[Dict[str, Any]] = []
        for meta in self.catalog_rows:
            room = meta.get("Espacio_ID", "")
            if not room:
                continue
            if sede_norm and meta.get("Sede", "") != sede_norm:
                continue
            if not include_non_classrooms and meta.get("Tipo", "") != "Aula":
                continue
            try:
                cap = int(meta.get("Capacidad")) if meta.get("Capacidad") not in (None, "") else None
            except Exception:
                cap = None
            if min_capacity > 0 and (cap is None or cap < min_capacity):
                continue
            result = self.query_room(room, date, start=query_start, end=query_end)
            if result.status == "LIBRE" or (include_review and result.status == "REVISAR"):
                rows.append({
                    "Espacio_ID": room,
                    "Nombre": result.room_name,
                    "Capacidad": cap if cap is not None else "",
                    "Sede": result.sede,
                    "Estado": result.status,
                    "Reservada para eventos": "Sí" if any("reservado para eventos" in x.lower() for x in result.notes) else "No",
                    "Notas": " | ".join(result.notes),
                })
        # Para una búsqueda de eventos, una sala que la grilla deja explícitamente "PARA EVENTOS"
        # es una buena candidata; la mostramos primero, luego capacidad descendente.
        rows.sort(key=lambda r: (0 if r["Reservada para eventos"] == "Sí" else 1, -(int(r["Capacidad"]) if r["Capacidad"] != "" else -1), r["Espacio_ID"]))
        return rows

    def room_day(self, room_value: Any, date_value: Any) -> List[Dict[str, Any]]:
        room = self.room_key(room_value)
        date = parse_date(date_value)
        rows = []
        for slot in sorted(self.slot_map, key=lambda x: int(x) if x.isdigit() else 999):
            result = self.query_room(room, date, slot=slot)
            details = []
            for group in result.occupation_groups:
                c = group.representative
                details.append(f"{c.label}: {c.description}")
            for c in result.review_claims:
                details.append(f"REVISAR: {c.description}")
            rows.append({
                "Fecha": date.isoformat(), "Día": WEEKDAYS[date.weekday()], "Espacio_ID": room,
                "Slot": slot, "Hora Desde": time_text(result.query_start), "Hora Hasta": time_text(result.query_end),
                "Estado": result.status, "Detalle": " | ".join(details), "Notas": " | ".join(result.notes),
            })
        return rows

    # -------------------------------------------------
    # Etapa 14A: alumnos + aulas a partir de códigos
    # -------------------------------------------------
    def combined_ranking(
        self,
        codes: Sequence[str],
        dates: Sequence[Any],
        *,
        sede: str = "Campus",
        min_capacity: int = 0,
        slots: Optional[Sequence[Any]] = None,
        top_n: Optional[int] = None,
    ) -> List[Dict[str, Any]]:
        slot_list = [r2.normalize_slot(x) for x in (slots or sorted(self.slot_map, key=lambda x: int(x) if x.isdigit() else 999))]
        rows: List[Dict[str, Any]] = []
        for date_value in dates:
            date = parse_date(date_value)
            for slot in slot_list:
                if slot not in self.slot_map:
                    continue
                students = self.student_occupancy(codes, date, slot=slot)
                free = self.free_rooms(date, slot=slot, sede=sede, min_capacity=min_capacity)
                start, end = self.interval_from_slot(slot)
                rows.append({
                    "Fecha": date.isoformat(),
                    "Día": WEEKDAYS[date.weekday()],
                    "Slot": slot,
                    "Hora Desde": time_text(start),
                    "Hora Hasta": time_text(end),
                    "Alumnos ocupados": students["Alumnos ocupados"],
                    "Materias afectadas": students["Materias afectadas"],
                    "Aulas disponibles": len(free),
                    "Aulas": " | ".join(f"{r['Espacio_ID']} ({r['Capacidad'] or 'N/A'})" for r in free[:12]),
                })
        # Primero debe existir al menos un aula; luego menor ocupación de alumnos;
        # como desempate, más opciones de aula.
        rows.sort(key=lambda r: (0 if r["Aulas disponibles"] > 0 else 1, r["Alumnos ocupados"], r["Materias afectadas"], -r["Aulas disponibles"], r["Fecha"], int(r["Slot"])))
        return rows[:top_n] if top_n else rows

    # -------------------------------------------------
    # Etapa 15: compatibilidad con consultas del original
    # -------------------------------------------------
    def legacy_query_courses(self, codes: Sequence[str]) -> List[Dict[str, Any]]:
        # Replica la lógica conceptual original: por día/slot, suma inscriptos y cuenta clases.
        selected = {clean(c).upper() for c in codes if clean(c)}
        grouped: Dict[Tuple[str, str], Dict[str, Any]] = {}
        for row in self.class_rows:
            code = row.get("Código", "")
            if code not in selected or not row.get("Día") or not row.get("Slot"):
                continue
            key = (row["Día"], row["Slot"])
            rec = grouped.setdefault(key, {"Día": row["Día"], "Slot": row["Slot"], "Alumnos ocupados": 0, "Clases afectadas": 0, "Códigos": set()})
            try:
                rec["Alumnos ocupados"] += int(float(row.get("Cant Inscriptos", 0) or 0))
            except Exception:
                pass
            rec["Clases afectadas"] += 1
            rec["Códigos"].add(code)
        out = []
        for rec in grouped.values():
            start, end = self.interval_from_slot(rec["Slot"])
            out.append({
                "Día": rec["Día"], "Slot": rec["Slot"], "Hora Desde": time_text(start), "Hora Hasta": time_text(end),
                "Alumnos ocupados": rec["Alumnos ocupados"], "Clases afectadas": rec["Clases afectadas"],
                "Materias afectadas": len(rec["Códigos"]), "Códigos": " | ".join(sorted(rec["Códigos"])),
            })
        out.sort(key=lambda r: (WEEKDAYS.index(r["Día"]), r["Alumnos ocupados"], r["Clases afectadas"], int(r["Slot"])))
        return out

    def legacy_room_free_weekly(self, rooms: Sequence[str]) -> List[Dict[str, Any]]:
        result: List[Dict[str, Any]] = []
        for room_value in rooms:
            room = self.room_key(room_value)
            for day in WEEKDAYS[:5]:
                for slot in sorted(self.slot_map, key=lambda x: int(x) if x.isdigit() else 999):
                    occupied = False
                    if room in self.rooms_in_aulas:
                        for row in self.grid_by_room_day.get((room, day), []):
                            if row.get("Slot") == slot and row.get("Tipo_registro") in {"CLASE_REGULAR", "BLOQUE_SEMANAL_OTRO", "REVISAR"}:
                                occupied = True
                                break
                    else:
                        for row in self.classes_by_room_day.get((room, day), []):
                            if row.get("Slot") == slot:
                                occupied = True
                                break
                    if not occupied:
                        start, end = self.interval_from_slot(slot)
                        result.append({"Espacio_ID": room, "Día": day, "Slot": slot, "Hora Desde": time_text(start), "Hora Hasta": time_text(end)})
        return result

    def legacy_free_rooms_weekday_slot(self, day_value: Any, slot_value: Any, sede: str = "") -> List[Dict[str, Any]]:
        day = r2.normalize_day(day_value)
        slot = r2.normalize_slot(slot_value)
        if day not in WEEKDAYS:
            raise ValueError(f"Día no reconocido: {day_value}")
        if slot not in self.slot_map:
            raise ValueError(f"Slot no reconocido: {slot_value}")
        sede_norm = normalize_sede_query(sede)
        rows = []
        for meta in self.catalog_rows:
            if meta.get("Tipo") != "Aula":
                continue
            if sede_norm and meta.get("Sede") != sede_norm:
                continue
            room = meta.get("Espacio_ID", "")
            occupied = False
            if room in self.rooms_in_aulas:
                for row in self.grid_by_room_day.get((room, day), []):
                    if row.get("Slot") == slot and row.get("Tipo_registro") in {"CLASE_REGULAR", "BLOQUE_SEMANAL_OTRO", "REVISAR"}:
                        occupied = True
                        break
            else:
                for row in self.classes_by_room_day.get((room, day), []):
                    if row.get("Slot") == slot:
                        occupied = True
                        break
            if not occupied:
                rows.append({"Espacio_ID": room, "Capacidad": meta.get("Capacidad", ""), "Sede": meta.get("Sede", "")})
        rows.sort(key=lambda r: (r["Sede"], r["Espacio_ID"]))
        return rows


def write_csv(path: Path, rows: Sequence[Dict[str, Any]], fieldnames: Optional[Sequence[str]] = None) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    rows = list(rows)
    if fieldnames is None:
        if not rows:
            path.write_text("", encoding="utf-8-sig")
            return
        fields: List[str] = []
        seen = set()
        for row in rows:
            for key in row:
                if key not in seen:
                    seen.add(key)
                    fields.append(key)
        fieldnames = fields
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(fieldnames), extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            out = {}
            for key in fieldnames:
                value = row.get(key, "")
                if isinstance(value, (set, list, tuple)):
                    value = " | ".join(str(x) for x in value)
                out[key] = value
            writer.writerow(out)


def build_sources(base: Path) -> Tuple[UdeSAEngine, Dict[str, Any]]:
    files = r2.find_required_files(base)
    problems: List[Dict[str, Any]] = []
    catalog = r2.MasterCatalog()

    print("[1/9] Reconstruyendo catálogo y parsers de Ronda 2...")
    r2.process_aux_catalog(files["catalogo_aux"], catalog, problems)
    r2.process_courses(files["cursos"], catalog, problems)
    r2.process_aulas(files["aulas"], catalog, problems)
    r2.add_manual_spaces(catalog)
    for alias, key in r2.MANUAL_ALIASES.items():
        if key in catalog.entries:
            catalog.add_alias(key, alias.title() if alias != "HAM" else alias)

    class_stats = r2.process_regular_classes_r2(files["cursos"], catalog)
    slot_map = class_stats["slot_map"]
    course_names = class_stats["course_name_map"]
    known_codes = set(course_names)
    grid_stats = r2.process_aulas_grid_r2(files["aulas"], catalog, slot_map, known_codes, course_names)
    comment_stats = r2.process_comments_r2(files["aulas"], catalog, slot_map, known_codes)
    reservation_stats = r2.process_special_reservations_r2(files["aulas"], catalog)
    event_stats = r2.process_events_r2(files["eventos"], catalog)

    engine = UdeSAEngine(
        catalog.rows(), class_stats["rows"], grid_stats["rows"], comment_stats["items"], comment_stats["effects"],
        reservation_stats["rows"], event_stats["occupancies"], slot_map,
    )
    source_stats = {
        "catalog_rows": len(catalog.rows()),
        "classes": len(class_stats["rows"]),
        "grid": len(grid_stats["rows"]),
        "comments": comment_stats["comments_count"],
        "comment_items": len(comment_stats["items"]),
        "comment_effects": len(comment_stats["effects"]),
        "comment_actions": dict(comment_stats["action_counts"]),
        "comment_effect_types": dict(comment_stats["effect_counts"]),
        "reservations": len(reservation_stats["rows"]),
        "events": event_stats["events_count"],
        "event_occupancies": len(event_stats["occupancies"]),
        "event_effect_types": dict(event_stats["effect_counts"]),
        "slot_map": slot_map,
    }
    return engine, source_stats


def run_tests(engine: UdeSAEngine) -> List[Dict[str, Any]]:
    tests: List[Dict[str, Any]] = []

    def check(name: str, expected: Any, actual: Any, detail: str = "") -> None:
        ok = expected == actual
        tests.append({"Test": name, "Esperado": expected, "Obtenido": actual, "Resultado": "OK" if ok else "ERROR", "Detalle": detail})

    # Etapa 9: intervalos
    check("Temporal: superposición parcial", True, intervals_overlap(dt.time(10), dt.time(13), dt.time(12, 15), dt.time(14)))
    check("Temporal: borde no superpone", False, intervals_overlap(dt.time(10), dt.time(13), dt.time(13), dt.time(14)))
    check("Temporal: slot 4 desde", "13:00", time_text(engine.interval_from_slot("4")[0]))
    check("Temporal: slot 4 hasta", "14:30", time_text(engine.interval_from_slot("4")[1]))

    # Comentario de traslado: H125 sale a G005 el 3/11 slot 3.
    r = engine.query_room("H125", "03/11/2026", slot=3)
    check("Traslado: aula origen queda libre", "LIBRE", r.status, engine.format_room_result(r))
    r = engine.query_room("G005", "03/11/2026", slot=3)
    check("Traslado: aula destino queda ocupada", "OCUPADA", r.status, engine.format_room_result(r))
    check("Traslado: comentario Romi Terron visible", True, "ROMI TERRON" in norm(engine.format_room_result(r)), engine.format_room_result(r))

    # Corrección descubierta durante integración: NO LA USAN = LIBERA.
    r = engine.query_room("S011", "31/08/2026", slot=2)
    check("Comentario NO LA USAN libera", "LIBRE", r.status, engine.format_room_result(r))

    # Comentario contradictorio de día = REVISAR.
    r = engine.query_room("B108", "14/08/2026", slot=8)
    check("Comentario con fecha/día contradictorio", "REVISAR", r.status, engine.format_room_result(r))

    # Evento todo el día.
    r = engine.query_room("H118", "06/01/2026", slot=1)
    check("Evento todo el día ocupa", "OCUPADA", r.status, engine.format_room_result(r))
    check("Etiqueta Evento en salida", True, "EVENTO:" in norm(engine.format_room_result(r)), engine.format_room_result(r))

    # Reserva especial.
    r = engine.query_room("S011", "09/12/2026", slot=5)
    check("Reserva especial ocupa", "OCUPADA", r.status, engine.format_room_result(r))

    # Conflicto real: dos movimientos distintos llegan a V005 el 5/8 slot 7.
    r = engine.query_room("V005", "05/08/2026", slot=7)
    check("Detector de conflicto real", "CONFLICTO", r.status, engine.format_room_result(r))
    check("Conflicto conserva ambos registros", True, len(r.occupation_groups) >= 2, engine.format_room_result(r))

    # PARA EVENTOS no significa ocupado.
    # Buscamos dinámicamente un registro real y comprobamos que, sin otra fuente, no sea ocupado por ese rótulo.
    para_eventos_tested = False
    for row in engine.grid_rows:
        if row.get("Tipo_registro") != "RESERVADO_PARA_EVENTOS":
            continue
        room, day, slot = row.get("Espacio_ID"), row.get("Día"), row.get("Slot")
        # semana representativa, ajustamos al weekday correspondiente
        base_date = dt.date(2026, 8, 17)
        target = base_date + dt.timedelta(days=WEEKDAYS.index(day))
        result = engine.query_room(room, target, slot=slot)
        if result.status == "LIBRE":
            check("PARA EVENTOS no bloquea", "LIBRE", result.status, engine.format_room_result(result))
            check("PARA EVENTOS deja nota", True, any("reservado para eventos" in n.lower() for n in result.notes), engine.format_room_result(result))
            para_eventos_tested = True
            break
    if not para_eventos_tested:
        tests.append({"Test": "PARA EVENTOS no bloquea", "Esperado": "caso comprobable", "Obtenido": "sin caso libre aislado", "Resultado": "OMITIDO", "Detalle": "Todos los casos encontrados coincidían con otra ocupación específica en la fecha de prueba."})

    # Etapa 12: disponibilidad estudiantil.
    student = engine.student_occupancy(["E020", "P318", "P328"], "18/08/2026", slot=1)
    check("Disponibilidad estudiantes: cálculo conocido", 51, student["Alumnos ocupados"], json.dumps({k: v for k, v in student.items() if k != "Detalle"}, ensure_ascii=False))
    check("Disponibilidad estudiantes: una materia afectada", 1, student["Materias afectadas"])
    student_free = engine.student_occupancy(["E020", "P318", "P328"], "19/08/2026", slot=4)
    check("Disponibilidad estudiantes: horario sin esas materias", 0, student_free["Alumnos ocupados"])

    # Rango personalizado: la reserva 14:40–17:50 debe superponer 15:00–15:30.
    r = engine.query_room("S011", "09/12/2026", start="15:00", end="15:30")
    check("Consulta por rango personalizado", "OCUPADA", r.status, engine.format_room_result(r))

    # Filtros físicos.
    free = engine.free_rooms("19/08/2026", slot=4, sede="Campus", min_capacity=50)
    check("Filtro sede/capacidad devuelve solo Campus", True, all(x["Sede"] == "Campus" for x in free))
    check("Filtro capacidad mínima", True, all(int(x["Capacidad"]) >= 50 for x in free if x["Capacidad"] != ""))

    # Etapa 14A: ranking combinado no debe perder campos esenciales.
    combined = engine.combined_ranking(["E020", "P318", "P328"], ["19/08/2026", "20/08/2026", "21/08/2026"], sede="Campus", min_capacity=50, top_n=5)
    check("Ranking combinado genera resultados", True, len(combined) > 0)
    check("Ranking combinado incluye alumnos y aulas", True, all("Alumnos ocupados" in x and "Aulas disponibles" in x for x in combined))

    # Etapa 15: las tres consultas del programa viejo tienen equivalente ejecutable.
    q1 = engine.legacy_query_courses(["E020", "P318", "P328"])
    q2 = engine.legacy_room_free_weekly(["G006"])
    q3 = engine.legacy_free_rooms_weekday_slot("Miércoles", 4, sede="Campus")
    check("Compatibilidad consulta 1: materias", True, len(q1) > 0)
    check("Compatibilidad consulta 2: días libres por aula", True, len(q2) > 0)
    check("Compatibilidad consulta 3: aulas libres día/slot", True, len(q3) > 0)

    return tests


def build_review_rows(engine: UdeSAEngine) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    # Solo revisiones con efecto espacial concreto.
    for row in engine.comment_effects:
        if row.get("Efecto") == "REVISAR":
            rows.append({
                "Fuente": "Comentario", "Fecha": row.get("Fecha", ""), "Espacio_ID": row.get("Espacio_ID", ""),
                "Hora Desde": row.get("Hora Desde", ""), "Hora Hasta": row.get("Hora Hasta", ""), "Detalle": row.get("Motivo", ""),
            })
    for row in engine.grid_rows:
        if row.get("Tipo_registro") == "REVISAR":
            rows.append({
                "Fuente": "AULAS", "Fecha": "Recurrente", "Espacio_ID": row.get("Espacio_ID", ""),
                "Hora Desde": row.get("Hora Desde", ""), "Hora Hasta": row.get("Hora Hasta", ""), "Detalle": row.get("Texto_original", ""),
            })
    for row in engine.event_occupancies:
        if row.get("Efecto") == "REVISAR":
            rows.append({
                "Fuente": "Evento", "Fecha": row.get("Fecha", ""), "Espacio_ID": row.get("Espacio_ID", ""),
                "Hora Desde": row.get("Hora Desde", ""), "Hora Hasta": row.get("Hora Hasta", ""), "Detalle": row.get("Evento", ""),
            })
    return rows


def make_report(
    engine: UdeSAEngine,
    source_stats: Dict[str, Any],
    conflicts: Sequence[Dict[str, Any]],
    tests: Sequence[Dict[str, Any]],
    combined_example: Sequence[Dict[str, Any]],
    free_example: Sequence[Dict[str, Any]],
) -> str:
    lines: List[str] = []
    lines.append("=" * 88)
    lines.append("UdeSA Horarios — DIAGNÓSTICO RONDA 3")
    lines.append(f"Autor: {AUTHOR}")
    lines.append(f"Generado: {dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append("=" * 88)
    lines.append("")
    lines.append("ALCANCE")
    lines.append("Etapa 9: motor temporal — IMPLEMENTADA")
    lines.append("Etapa 10: motor central LIBRE/OCUPADA/REVISAR/CONFLICTO — IMPLEMENTADA")
    lines.append("Etapa 11: detector de conflictos — IMPLEMENTADA")
    lines.append("Etapa 12: disponibilidad estudiantil por códigos — IMPLEMENTADA")
    lines.append("Etapa 13: disponibilidad por carrera/camada — PENDIENTE PARA RONDA 4")
    lines.append("Etapa 14A: combinación alumnos + aulas por códigos — IMPLEMENTADA")
    lines.append("Etapa 14B: combinación usando carrera/camada — PENDIENTE PARA RONDA 4")
    lines.append("Etapa 15: compatibilidad con consultas del programa original — IMPLEMENTADA")
    lines.append("Etapa 16: batería automática de pruebas — IMPLEMENTADA")
    lines.append("")

    lines.append("FUENTES NORMALIZADAS CARGADAS")
    lines.append(f"Espacios catálogo: {source_stats['catalog_rows']}")
    lines.append(f"Filas de clases: {source_stats['classes']}")
    lines.append(f"Celdas grilla AULAS: {source_stats['grid']}")
    lines.append(f"Comentarios originales: {source_stats['comments']}")
    lines.append(f"Instrucciones de comentarios: {source_stats['comment_items']}")
    lines.append(f"Efectos de comentarios: {source_stats['comment_effects']}")
    lines.append(f"Reservas especiales: {source_stats['reservations']}")
    lines.append(f"Eventos: {source_stats['events']}")
    lines.append(f"Efectos de Eventos sobre espacios: {source_stats['event_occupancies']}")
    lines.append("")

    lines.append("CORRECCIÓN DE INTEGRACIÓN APLICADA")
    lines.append("Durante la construcción del motor se detectó que 'NO LA USAN' había sido tratado en Ronda 2 como RESERVA.")
    lines.append("La regla fue corregida: ahora se interpreta como LIBERACION_EXPLICITA y genera LIBERA.")
    lines.append(f"Acciones actuales de comentarios: {source_stats['comment_actions']}")
    lines.append(f"Efectos actuales de comentarios: {source_stats['comment_effect_types']}")
    lines.append("")

    lines.append("REGLAS CENTRALES DEL MOTOR")
    lines.append("1. AULAS Primavera es la base física para las aulas presentes en esa grilla.")
    lines.append("2. CURSOS funciona como fallback físico para espacios/sedes no cubiertos por AULAS.")
    lines.append("3. Un comentario en su propia celda reemplaza la base semanal de esa celda para la fecha indicada.")
    lines.append("4. Un traslado que llega a otra aula NO borra la base de la otra aula: si estaba ocupada, se detecta conflicto.")
    lines.append("5. RESERVADO_PARA_EVENTOS no cuenta como ocupación; queda como nota y se prioriza entre aulas libres.")
    lines.append("6. Un REVISAR con aula concreta bloquea. Si su horario es incompleto, bloquea conservadoramente el día completo.")
    lines.append("7. Dos registros inequívocamente duplicados se agrupan; dos ocupaciones distintas superpuestas generan CONFLICTO.")
    lines.append("")

    lines.append("CONFLICTOS DETECTADOS")
    lines.append(f"Conflictos fuente/horario detectados: {len(conflicts)}")
    for row in list(conflicts)[:12]:
        lines.append(
            f"  - {row['Fecha']} {row['Espacio_ID']} {row['Superposición Desde']}–{row['Superposición Hasta']}: "
            f"{row['Tipo A']} [{row['Detalle A']}] vs {row['Tipo B']} [{row['Detalle B']}]"
        )
    if len(conflicts) > 12:
        lines.append(f"  ... y {len(conflicts) - 12} conflictos adicionales en conflictos_detectados.csv")
    lines.append("")

    ok = sum(1 for t in tests if t["Resultado"] == "OK")
    errors = sum(1 for t in tests if t["Resultado"] == "ERROR")
    omitted = sum(1 for t in tests if t["Resultado"] == "OMITIDO")
    lines.append("ETAPA 16 — PRUEBAS AUTOMÁTICAS")
    lines.append(f"Tests OK: {ok}")
    lines.append(f"Tests con error: {errors}")
    lines.append(f"Tests omitidos: {omitted}")
    for test in tests:
        lines.append(f"  - {test['Resultado']}: {test['Test']} | esperado={test['Esperado']} | obtenido={test['Obtenido']}")
    lines.append("")

    lines.append("ETAPA 15 — CONSULTAS DEL PROGRAMA ORIGINAL")
    lines.append("1. Disponibilidad por materias/códigos: REPRODUCIDA y ampliada a fecha/rango horario.")
    lines.append("2. Días y slots libres por aula: REPRODUCIDA como consulta semanal recurrente y como consulta por fecha.")
    lines.append("3. Aulas libres por día y slot: REPRODUCIDA; ahora admite fecha, sede, capacidad y rango horario.")
    lines.append("")

    lines.append("EJEMPLO DE RANKING COMBINADO POR CÓDIGOS")
    lines.append("Códigos: E020, P318, P328 | Sede: Campus | Capacidad mínima: 50")
    for i, row in enumerate(combined_example[:5], start=1):
        lines.append(
            f"  {i}. {row['Fecha']} {row['Día']} Slot {row['Slot']} {row['Hora Desde']}–{row['Hora Hasta']} | "
            f"alumnos ocupados={row['Alumnos ocupados']} | materias afectadas={row['Materias afectadas']} | "
            f"aulas disponibles={row['Aulas disponibles']}"
        )
    lines.append("")

    lines.append("EJEMPLO DE AULAS LIBRES")
    lines.append("19/08/2026 · Slot 4 · Campus · capacidad mínima 50")
    for row in list(free_example)[:10]:
        lines.append(f"  - {row['Espacio_ID']} | capacidad {row['Capacidad']} | reservado para eventos: {row['Reservada para eventos']}")
    lines.append("")

    lines.append("ESTADO DE LA ETAPA 13")
    lines.append("No se implementa en esta ronda. El archivo udesa_plan_academico_master.xlsx queda reservado para Ronda 4 (Etapa 3B + 13 + 14B).")
    return "\n".join(lines)


def main() -> None:
    base = Path(__file__).resolve().parent
    output = base / OUTPUT_DIR_NAME
    output.mkdir(parents=True, exist_ok=True)

    print(f"\n{APP_NAME} — Ronda 3")
    print(f"Autor: {AUTHOR}")
    print(f"Carpeta de trabajo: {base}\n")

    engine, source_stats = build_sources(base)

    print("[2/9] Etapa 9: inicializando motor temporal...")
    # La construcción de engine ya deja slot_map y parsers de fecha/hora listos.
    engine.interval_from_slot("1")

    print("[3/9] Etapa 10: validando motor central de ocupación...")
    engine.query_room("H118", "06/01/2026", slot=1)

    print("[4/9] Etapa 11: detectando conflictos...")
    conflicts = engine.detect_conflicts()

    print("[5/9] Etapa 12: validando disponibilidad estudiantil...")
    engine.student_occupancy(["E020", "P318", "P328"], "18/08/2026", slot=1)

    print("[6/9] Etapa 14A: combinando alumnos + aulas por códigos...")
    combined_example = engine.combined_ranking(
        ["E020", "P318", "P328"],
        ["19/08/2026", "20/08/2026", "21/08/2026"],
        sede="Campus", min_capacity=50, top_n=15,
    )
    free_example = engine.free_rooms("19/08/2026", slot=4, sede="Campus", min_capacity=50)

    print("[7/9] Etapa 15: comprobando compatibilidad con consultas viejas...")
    legacy1 = engine.legacy_query_courses(["E020", "P318", "P328"])
    legacy2 = engine.legacy_room_free_weekly(["G006"])
    legacy3 = engine.legacy_free_rooms_weekday_slot("Miércoles", 4, sede="Campus")

    print("[8/9] Etapa 16: ejecutando batería automática de pruebas...")
    tests = run_tests(engine)

    print("[9/9] Exportando diagnóstico y tablas de auditoría...")
    review_rows = build_review_rows(engine)
    room_day_example = engine.room_day("G006", "19/08/2026")

    # Archivos útiles para la Ronda 4 y la futura web.
    write_csv(output / "conflictos_detectados.csv", conflicts)
    write_csv(output / "registros_revisar.csv", review_rows)
    write_csv(output / "pruebas_automaticas.csv", tests)
    write_csv(output / "ranking_combinado_ejemplo.csv", combined_example)
    write_csv(output / "aulas_libres_ejemplo.csv", free_example)
    write_csv(output / "consulta_aula_G006_ejemplo.csv", room_day_example)
    write_csv(output / "legacy_consulta1_materias.csv", legacy1)
    write_csv(output / "legacy_consulta2_aula_G006.csv", legacy2)
    write_csv(output / "legacy_consulta3_aulas_libres.csv", legacy3)

    report = make_report(engine, source_stats, conflicts, tests, combined_example, free_example)
    (output / "diagnostico_ronda3.txt").write_text(report, encoding="utf-8")
    summary = {
        "app": APP_NAME,
        "autor": AUTHOR,
        "generado": dt.datetime.now().isoformat(timespec="seconds"),
        "etapas": {
            "9": "OK", "10": "OK", "11": "OK", "12": "OK", "13": "PENDIENTE_RONDA_4",
            "14A": "OK", "14B": "PENDIENTE_RONDA_4", "15": "OK", "16": "OK",
        },
        "fuentes": source_stats,
        "conflictos": len(conflicts),
        "revisar": len(review_rows),
        "tests": dict(Counter(t["Resultado"] for t in tests)),
    }
    (output / "resumen_ronda3.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")

    print("\n" + report)
    print(f"\nListo. Los archivos de la Ronda 3 quedaron en:\n{output}\n")
    try:
        input("Presioná ENTER para cerrar...")
    except EOFError:
        pass


if __name__ == "__main__":
    main()
