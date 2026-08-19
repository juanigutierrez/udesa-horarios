#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
UdeSA Horarios — Ronda 4.5
Autor: Juan Ignacio Gutiérrez Julián

Objetivos:
- Generalizar el sistema para cualquier período académico (Otoño / Primavera).
- Separar configuración de fuentes del código.
- Mantener CURSOS y AULAS como fuentes semestrales, Eventos como fuente anual,
  y el maestro académico como fuente transversal.
- Preparar un contrato simple para la futura web: defaults sencillos y filtros avanzados opcionales.
- Distinguir validez curricular de oferta en el período.
- Mantener regresión completa de las Rondas 3 y 4.

No modifica los Excel originales. Usa únicamente la biblioteca estándar de Python.
"""
from __future__ import annotations

import csv
import datetime as dt
import json
import os
import re
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Set, Tuple
from zipfile import ZipFile, BadZipFile

import udesa_horarios_ronda2 as r2
import udesa_horarios_ronda3 as r3

APP_NAME = "UdeSA Horarios"
AUTHOR = "Juan Ignacio Gutiérrez Julián"
OUTPUT_DIR_NAME = "salida_ronda45"
CONFIG_FILENAME = "fuentes_udesahorarios.json"

OPTIONAL_TYPES = {"OPTATIVA", "ELECTIVA"}
USABLE_PERIOD_STATES = {
    "OFERTADO_CODIGO_PLAN",
    "RESUELTO_OFERTA_UNICA_EN_PERIODO",
    "RESUELTO_NOMBRE_EXACTO_EN_PERIODO",
}
OPERATIONAL_COVERAGE_STATES = {"OPERATIVO", "OPERATIVO_CON_REVISIONES"}

CAREER_ALIASES = {
    "RRII": "Relaciones Internacionales",
    "RIII": "Relaciones Internacionales",
    "RI": "Relaciones Internacionales",
    "RELACIONES INTERNACIONALES": "Relaciones Internacionales",
    "CP": "Ciencia Política y Gobierno",
    "CIENCIA POLITICA": "Ciencia Política y Gobierno",
    "CIENCIA POLITICA Y GOBIERNO": "Ciencia Política y Gobierno",
    "ADMIN": "Administración de Empresas",
    "ADMINISTRACION": "Administración de Empresas",
    "ADMINISTRACION DE EMPRESAS": "Administración de Empresas",
    "ECO": "Economía",
    "ECONOMIA": "Economía",
    "FIN": "Finanzas",
    "FINANZAS": "Finanzas",
    "ABOGACIA": "Abogacía",
    "NEGOCIOS DIGITALES": "Negocios Digitales",
}


def clean(value: Any) -> str:
    return r2.clean_text(value)


def norm(value: Any) -> str:
    return r2.norm_text(value)


def to_int(value: Any) -> Optional[int]:
    if value in (None, ""):
        return None
    try:
        return int(float(value))
    except Exception:
        return None


def normalize_variant(value: Any) -> str:
    text = clean(value)
    if not text or norm(text) in {"NONE", "NULL", "N/A", "NA", "REGULAR"}:
        return ""
    return text


def variant_label(value: Any) -> str:
    v = normalize_variant(value)
    return v if v else "Regular"


def normalize_term(value: Any) -> str:
    n = norm(value)
    mapping = {
        "OTONO": "OTONO",
        "OTOÑO": "OTONO",
        "1": "OTONO",
        "PRIMER SEMESTRE": "OTONO",
        "SEMESTRE 1": "OTONO",
        "PRIMAVERA": "PRIMAVERA",
        "2": "PRIMAVERA",
        "SEGUNDO SEMESTRE": "PRIMAVERA",
        "SEMESTRE 2": "PRIMAVERA",
    }
    if n not in mapping:
        raise ValueError(f"Semestre académico no reconocido: {value}")
    return mapping[n]


def term_label(value: Any) -> str:
    return "Otoño" if normalize_term(value) == "OTONO" else "Primavera"


def normalize_course_name(value: Any) -> str:
    text = clean(value)
    if not text:
        return ""
    text = unicodedata.normalize("NFKD", text)
    text = "".join(c for c in text if not unicodedata.combining(c)).upper()
    text = re.sub(r"\([^)]*\)", " ", text)
    text = re.sub(r"\b[A-Z]{1,5}\d{2,3}\b", " ", text)
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    return " ".join(text.split())


def extract_codes(value: Any) -> List[str]:
    return sorted(set(re.findall(r"\b[A-Z]{1,5}\d{2,3}\b", clean(value).upper())))


def canonical_career(value: Any, careers: Sequence[str]) -> str:
    text = clean(value)
    if not text:
        return ""
    aliased = CAREER_ALIASES.get(norm(text), text)
    for c in careers:
        if norm(c) == norm(aliased):
            return c
    return aliased


def parse_iso_date(value: Any) -> Optional[dt.date]:
    text = clean(value)
    if not text:
        return None
    try:
        return dt.date.fromisoformat(text[:10])
    except Exception:
        return None


@dataclass(frozen=True)
class AcademicPeriod:
    year: int
    term: str
    start: dt.date
    end: dt.date

    @property
    def term_number(self) -> int:
        return 1 if self.term == "OTONO" else 2

    @property
    def id(self) -> str:
        return f"{self.year}_{self.term}"

    @property
    def label(self) -> str:
        return f"{term_label(self.term)} {self.year}"

    def contains(self, date_value: Any) -> bool:
        date = r3.parse_date(date_value)
        return self.start <= date <= self.end


@dataclass(frozen=True)
class PlanKey:
    carrera: str
    camada: Optional[int]
    sede: str
    variante: str

    def label(self) -> str:
        return f"{self.carrera} · {self.camada if self.camada is not None else 'N/A'} · {self.sede} · {variant_label(self.variante)}"


class SourceRegistry:
    """Lee la configuración de fuentes. Ningún nombre de archivo de semestre queda en el motor."""

    def __init__(self, base: Path, config_path: Optional[Path] = None) -> None:
        self.base = Path(base)
        self.config_path = Path(config_path or self.base / CONFIG_FILENAME)
        if not self.config_path.is_file():
            raise FileNotFoundError(f"No se encontró {self.config_path.name} en {self.config_path.parent}")
        self.config = json.loads(self.config_path.read_text(encoding="utf-8"))
        self.period_defs: Dict[str, Dict[str, Any]] = dict(self.config.get("periodos", {}))
        self.active_period_id = clean(self.config.get("periodo_activo"))
        self.defaults = dict(self.config.get("defaults", {}))
        self.global_sources = dict(self.config.get("fuentes_globales", {}))
        self.annual_sources = dict(self.config.get("fuentes_anuales", {}))
        self.manual_adjustments = dict(self.config.get("ajustes_manuales", {}))
        self.cohort_anchor = dict(self.config.get("cohorte_ancla", {}))
        if "anio" not in self.cohort_anchor or "camada_ingreso" not in self.cohort_anchor:
            raise ValueError("La configuración debe incluir cohorte_ancla con anio y camada_ingreso.")
        if self.active_period_id not in self.period_defs:
            raise ValueError(f"El período activo '{self.active_period_id}' no existe en la configuración.")

    def list_period_ids(self) -> List[str]:
        return sorted(self.period_defs, key=lambda pid: (int(self.period_defs[pid].get("anio", 0)), 1 if normalize_term(self.period_defs[pid].get("semestre")) == "OTONO" else 2))

    @staticmethod
    def _is_readable_source(path: Path) -> bool:
        """Devuelve True solo si la fuente existe y puede abrirse realmente.

        OneDrive/Excel puede dejar un archivo visible para ``is_file()`` pero bloquear su
        apertura temporalmente. En ese caso no debemos romper toda la app: la resolución
        de fuentes continúa con el siguiente candidato (por ejemplo, el snapshot de data/).
        Para archivos Office basados en ZIP también comprobamos que el contenedor pueda
        abrirse, evitando seleccionar una copia incompleta o en sincronización.
        """
        try:
            if not path.is_file():
                return False
            with path.open("rb") as fh:
                fh.read(4)
            if path.suffix.lower() in {".xlsx", ".xlsm", ".xltx", ".xltm"}:
                with ZipFile(path) as zf:
                    # Leer el índice alcanza para verificar que el XLSX es accesible y válido.
                    zf.namelist()[:1]
            return True
        except (PermissionError, OSError, BadZipFile):
            return False

    def _search_file(self, value: str, env_name: Optional[str] = None) -> Path:
        if env_name:
            env = os.environ.get(env_name, "").strip()
            if env and self._is_readable_source(Path(env)):
                return Path(env)
        raw = Path(value)
        if raw.is_absolute() and self._is_readable_source(raw):
            return raw
        filename = raw.name
        candidates: List[Path] = []
        # Para fuentes vivas, una copia central en una carpeta Excels cercana debe ganar
        # sobre el snapshot incluido en data/. Esto permite actualizar archivos sin tocar código.
        current = self.base
        central_candidates: List[Path] = []
        nearby_candidates: List[Path] = []
        for _ in range(8):
            central_candidates.append(current / "Excels" / filename)
            nearby_candidates.append(current / filename)
            if current.parent == current:
                break
            current = current.parent
        candidates.extend(central_candidates)
        # Después usamos la ruta explícita del proyecto como fallback reproducible.
        candidates.append(self.base / raw)
        candidates.extend(nearby_candidates)
        seen: Set[str] = set()
        for candidate in candidates:
            try:
                key = str(candidate.resolve())
            except Exception:
                key = str(candidate)
            if key in seen:
                continue
            seen.add(key)
            if self._is_readable_source(candidate):
                return candidate
        attempted = "\n  - ".join(str(x) for x in candidates[:20])
        raise FileNotFoundError(f"No se encontró la fuente '{value}'. Se buscó en:\n  - {attempted}")

    def _raw_period(self, period_id: Optional[str]) -> Tuple[str, Dict[str, Any]]:
        pid = self.active_period_id if period_id in (None, "", "ACTUAL") else clean(period_id).upper()
        if pid not in self.period_defs:
            raise KeyError(f"Período no configurado: {pid}")
        return pid, self.period_defs[pid]

    def source_paths(self, period_id: Optional[str] = None) -> Dict[str, Path]:
        pid, raw = self._raw_period(period_id)
        year = int(raw["anio"])
        semester_sources = dict(raw.get("fuentes", {}))
        event_map = dict(self.annual_sources.get("eventos", {}))
        if str(year) not in event_map:
            raise KeyError(f"No hay fuente anual de Eventos configurada para {year}.")
        paths = {
            "cursos": self._search_file(clean(semester_sources.get("cursos"))),
            "aulas": self._search_file(clean(semester_sources.get("aulas"))),
            "eventos": self._search_file(clean(event_map[str(year)])),
            "catalogo_aux": self._search_file(clean(self.global_sources.get("catalogo_espacios"))),
            "plan_master": self._search_file(clean(self.global_sources.get("plan_master")), "UDESA_PLAN_MASTER"),
        }
        legacy = clean(self.global_sources.get("catalogo_academico_legacy"))
        if legacy:
            paths["catalogo_academico_legacy"] = self._search_file(legacy)
        return paths

    def provisional_period(self, period_id: Optional[str] = None) -> AcademicPeriod:
        pid, raw = self._raw_period(period_id)
        year = int(raw["anio"])
        term = normalize_term(raw["semestre"])
        start = parse_iso_date(raw.get("fecha_inicio"))
        end = parse_iso_date(raw.get("fecha_fin"))
        # Si no se ingresaron fechas, usamos un rango provisorio amplio solo para poder parsear;
        # después se reemplaza por el rango dominante de CURSOS.
        if start is None:
            start = dt.date(year, 1 if term == "OTONO" else 7, 1)
        if end is None:
            end = dt.date(year, 7 if term == "OTONO" else 12, 31)
        return AcademicPeriod(year, term, start, end)

    def build_period_from_classes(self, period_id: Optional[str], class_rows: Sequence[Dict[str, Any]]) -> AcademicPeriod:
        _, raw = self._raw_period(period_id)
        provisional = self.provisional_period(period_id)
        explicit_start = parse_iso_date(raw.get("fecha_inicio"))
        explicit_end = parse_iso_date(raw.get("fecha_fin"))
        if explicit_start and explicit_end:
            return AcademicPeriod(provisional.year, provisional.term, explicit_start, explicit_end)

        ranges: Counter[Tuple[str, str]] = Counter()
        for row in class_rows:
            s, e = clean(row.get("Inicio vigencia")), clean(row.get("Fin vigencia"))
            if s and e:
                ranges[(s, e)] += 1
        if ranges:
            (start_s, end_s), _ = ranges.most_common(1)[0]
            start = parse_iso_date(start_s)
            end = parse_iso_date(end_s)
            if start and end:
                return AcademicPeriod(provisional.year, provisional.term, explicit_start or start, explicit_end or end)
        return provisional

    def active_period(self) -> AcademicPeriod:
        # Sin clases cargadas todavía, devuelve definición/provisorio.
        return self.provisional_period(self.active_period_id)

    def source_status_rows(self, period_id: Optional[str] = None) -> List[Dict[str, Any]]:
        period = self.provisional_period(period_id)
        paths = self.source_paths(period.id)
        rows = []
        mapping = {
            "cursos": ("CURSOS", "Semestral"),
            "aulas": ("AULAS", "Semestral + actualización continua"),
            "eventos": ("Eventos", "Anual + actualización continua"),
            "catalogo_aux": ("Catálogo de espacios", "Transversal"),
            "plan_master": ("Planes académicos", "Transversal"),
            "catalogo_academico_legacy": ("Catálogo académico auxiliar", "Transversal / respaldo"),
        }
        for key, path in paths.items():
            stat = path.stat()
            rows.append({
                "Fuente": mapping[key][0],
                "Frecuencia": mapping[key][1],
                "Período": period.label if key in {"cursos", "aulas"} else (str(period.year) if key == "eventos" else "Global"),
                "Archivo": path.name,
                "Ruta": str(path),
                "Última modificación local": dt.datetime.fromtimestamp(stat.st_mtime).isoformat(timespec="seconds"),
                "Estado": "OK",
            })
        return rows


class AcademicResolver45:
    def __init__(self, master_path: Path, engine: r3.UdeSAEngine, period: AcademicPeriod, registry: SourceRegistry, legacy_path: Optional[Path] = None) -> None:
        self.master_path = Path(master_path)
        self.legacy_path = Path(legacy_path) if legacy_path else None
        self.engine = engine
        self.period = period
        self.registry = registry
        self.plan_rows = self._load("PLAN_MASTER")
        self.coverage_rows = self._load("COBERTURA")
        self.catalog_code_rows = self._load("CATALOGO_CODIGOS")
        self.legacy_rows = self._load_legacy()
        self.known_curricular_codes: Set[str] = {clean(r.get("Código")).upper() for r in self.catalog_code_rows if clean(r.get("Código"))}

        self.current_codes: Set[str] = set(engine.classes_by_code.keys())
        self.current_course_names: Dict[str, str] = dict(engine.course_names)
        self.name_to_codes: Dict[str, Set[str]] = defaultdict(set)
        for code, name in self.current_course_names.items():
            n = normalize_course_name(name)
            if n:
                self.name_to_codes[n].add(code)

        self.plan_by_key: Dict[PlanKey, List[Dict[str, Any]]] = defaultdict(list)
        for row in self.plan_rows:
            self.plan_by_key[self.key_from_row(row)].append(row)
        self.coverage_by_key: Dict[PlanKey, Dict[str, Any]] = {}
        for row in self.coverage_rows:
            self.coverage_by_key[self.key_from_row(row)] = row
        self.careers = sorted({k.carrera for k in self.coverage_by_key if k.carrera})

        self.operational_rows: List[Dict[str, Any]] = []
        self.operational_by_key: Dict[PlanKey, List[Dict[str, Any]]] = defaultdict(list)
        self.coverage_operational: List[Dict[str, Any]] = []
        self._build_operational_plan()

    def _load(self, sheet_name: str) -> List[Dict[str, Any]]:
        with r2.XLSXReader(self.master_path) as book:
            if sheet_name not in book.sheet_names():
                raise RuntimeError(f"El maestro académico no contiene la hoja {sheet_name}.")
            return [dict(row) for _, row in book.dict_rows(sheet_name, 1)]

    def _load_legacy(self) -> List[Dict[str, Any]]:
        if not self.legacy_path or not self.legacy_path.is_file():
            return []
        try:
            with r2.XLSXReader(self.legacy_path) as book:
                if "MODELO" not in book.sheet_names():
                    return []
                return [dict(row) for _, row in book.dict_rows("MODELO", 1)]
        except Exception:
            return []

    @staticmethod
    def key_from_row(row: Dict[str, Any]) -> PlanKey:
        return PlanKey(
            carrera=clean(row.get("Carrera")),
            camada=to_int(row.get("Camada")),
            sede=clean(row.get("Sede")) or "No especificada",
            variante=normalize_variant(row.get("Variante")),
        )

    def _cohort_academic_year(self, cohort: Optional[int]) -> Optional[int]:
        if cohort is None:
            return None
        anchor_year = int(self.registry.cohort_anchor["anio"])
        anchor_cohort = int(self.registry.cohort_anchor["camada_ingreso"])
        first_year_cohort = anchor_cohort + (self.period.year - anchor_year)
        academic_year = first_year_cohort - cohort + 1
        return academic_year if 1 <= academic_year <= 8 else None

    def _period_rows(self, key: PlanKey, rows: Sequence[Dict[str, Any]]) -> Tuple[List[Dict[str, Any]], str, str]:
        # Prioridad 1: calendario explícito + semestre dentro del año explícito.
        exact_year = [r for r in rows if to_int(r.get("Año_calendario_fuente")) == self.period.year]
        if exact_year:
            exact_term = [r for r in exact_year if to_int(r.get("Semestre_en_año")) == self.period.term_number]
            if exact_term:
                return exact_term, "EXACTO_CALENDARIO_Y_SEMESTRE", f"Año_calendario_fuente={self.period.year}; Semestre_en_año={self.period.term_number}"
            # Compatibilidad con fuentes donde Semestre_en_año no está informado: mínimo=Otoño, máximo=Primavera.
            semesters = sorted({to_int(r.get("Semestre_fuente")) for r in exact_year if to_int(r.get("Semestre_fuente")) is not None})
            if semesters:
                selected_sem = min(semesters) if self.period.term == "OTONO" else max(semesters)
                selected = [r for r in exact_year if to_int(r.get("Semestre_fuente")) == selected_sem]
                if selected:
                    return selected, "EXACTO_CALENDARIO_FALLBACK_SEMESTRE", f"Año calendario explícito; Semestre_fuente seleccionado={selected_sem}"

        # Prioridad 2: inferencia transparente por camada, solo si no existe calendario explícito del período.
        academic_year = self._cohort_academic_year(key.camada)
        if academic_year is not None:
            inferred = [
                r for r in rows
                if to_int(r.get("Año_académico")) == academic_year
                and to_int(r.get("Semestre_en_año")) == self.period.term_number
            ]
            if inferred:
                return inferred, "INFERIDO_CAMADA", f"Año_académico inferido={academic_year}; Semestre_en_año={self.period.term_number}; inferencia marcada explícitamente"

        return [], "SIN_PERIODO_RESOLUBLE", f"No hay tramo resoluble para {self.period.label} sin inferencia adicional."

    def _resolve_plan_row(self, row: Dict[str, Any]) -> Tuple[str, str, str, str]:
        """Devuelve código operativo, estado período, oferta, detalle. Nunca invalida un código curricular por no ofertarse."""
        curricular_state = clean(row.get("Estado"))
        code = clean(row.get("Código")).upper()

        if curricular_state == "VALIDADO_DESDE_PLAN":
            if not code:
                return "", "REVISAR_VALIDADO_SIN_CODIGO", "DESCONOCIDA", "La fila curricular figura validada pero Código está vacío."
            if code in self.current_codes:
                return code, "OFERTADO_CODIGO_PLAN", "OFERTADO", f"Código validado por el plan y ofertado en {self.period.label}."
            return "", "SIN_OFERTA_EN_PERIODO", "NO_OFERTADO", f"Código curricular válido ({code}), pero no aparece en CURSOS de {self.period.label}. No se considera inválido."

        if curricular_state == "REVISAR_MULTICODIGO":
            candidates = extract_codes(row.get("Códigos_detectados"))
            active = [c for c in candidates if c in self.current_codes]
            if len(active) == 1:
                return active[0], "RESUELTO_OFERTA_UNICA_EN_PERIODO", "OFERTADO", f"Candidatos curriculares: {', '.join(candidates)}. Solo {active[0]} está ofertado en {self.period.label}. Resolución válida únicamente para este período."
            if len(active) > 1:
                return "", "REVISAR_MULTICODIGO_ACTIVO", "AMBIGUA", f"Más de un candidato está ofertado en {self.period.label}: {', '.join(active)}."
            return "", "SIN_OFERTA_EN_PERIODO_MULTICODIGO", "NO_OFERTADO", f"Ninguno de los candidatos está ofertado en {self.period.label}: {', '.join(candidates) or '[sin candidatos]'}. La ambigüedad curricular se conserva."

        if curricular_state == "REVISAR_SIN_CODIGO":
            normalized_name = normalize_course_name(row.get("Materia"))
            matching = sorted(self.name_to_codes.get(normalized_name, set())) if normalized_name else []
            if len(matching) == 1:
                return matching[0], "RESUELTO_NOMBRE_EXACTO_EN_PERIODO", "OFERTADO", f"Nombre normalizado coincide inequívocamente con {matching[0]} en CURSOS de {self.period.label}."
            if len(matching) > 1:
                return "", "REVISAR_NOMBRE_MULTIPLE", "AMBIGUA", f"El nombre coincide con varios códigos ofertados: {', '.join(matching)}."
            return "", "REVISAR_SIN_COINCIDENCIA", "DESCONOCIDA", f"No hubo coincidencia inequívoca por nombre en CURSOS de {self.period.label}."

        if curricular_state == "REQUISITO_GENERICO":
            return "", "REQUISITO_GENERICO", "NO_DETERMINABLE", "Optativa/electiva genérica: no se asigna un curso concreto automáticamente."

        return "", "REVISAR_ESTADO_DESCONOCIDO", "DESCONOCIDA", f"Estado curricular no reconocido: {curricular_state or '[vacío]'}."

    def _build_operational_plan(self) -> None:
        for key, coverage in self.coverage_by_key.items():
            selected, method, method_detail = self._period_rows(key, self.plan_by_key.get(key, []))
            state_counts: Counter[str] = Counter()
            mandatory_usable: Set[str] = set()
            optional_usable: Set[str] = set()

            for row in selected:
                operational_code, period_state, offering_state, detail = self._resolve_plan_row(row)
                state_counts[period_state] += 1
                tipo = clean(row.get("Tipo")).upper()
                if operational_code and period_state in USABLE_PERIOD_STATES:
                    if tipo in OPTIONAL_TYPES:
                        optional_usable.add(operational_code)
                    else:
                        mandatory_usable.add(operational_code)
                out = dict(row)
                out.update({
                    "Variante_normalizada": key.variante,
                    "Variante_mostrar": variant_label(key.variante),
                    "Periodo_ID": self.period.id,
                    "Periodo": self.period.label,
                    "Año_periodo": self.period.year,
                    "Semestre_periodo": term_label(self.period.term),
                    "Metodo_periodo": method,
                    "Detalle_periodo": method_detail,
                    "Código_operativo": operational_code,
                    "Estado_curricular": clean(row.get("Estado")),
                    "Estado_periodo": period_state,
                    "Oferta_en_periodo": offering_state,
                    "Detalle_resolución": detail,
                    "Es_optativa_electiva": "Sí" if tipo in OPTIONAL_TYPES else "No",
                })
                self.operational_rows.append(out)
                self.operational_by_key[key].append(out)

            scope = clean(coverage.get("Alcance_fuente"))
            usable_all = mandatory_usable | optional_usable
            if scope != "MALLA_EXTRAIDA":
                operational_status = scope or "FUENTE_NO_OPERATIVA"
            elif not selected:
                operational_status = "SIN_PERIODO_RESOLUBLE"
            elif not usable_all:
                operational_status = "SIN_CODIGOS_USABLES"
            elif any(k.startswith("REVISAR") or k == "REQUISITO_GENERICO" for k in state_counts):
                operational_status = "OPERATIVO_CON_REVISIONES"
            else:
                operational_status = "OPERATIVO"

            self.coverage_operational.append({
                "Periodo_ID": self.period.id,
                "Periodo": self.period.label,
                "Carrera": key.carrera,
                "Camada": key.camada if key.camada is not None else "",
                "Sede": key.sede,
                "Variante": variant_label(key.variante),
                "Variante_raw": key.variante,
                "Plan_version": coverage.get("Plan_version", ""),
                "Alcance_fuente": scope,
                "Metodo_periodo": method,
                "Detalle_periodo": method_detail,
                "Estado_operativo": operational_status,
                "Filas_periodo": len(selected),
                "Códigos_obligatorios_utilizables": len(mandatory_usable),
                "Códigos_optativos_identificados": len(optional_usable),
                "Códigos_obligatorios": " | ".join(sorted(mandatory_usable)),
                "Códigos_optativos": " | ".join(sorted(optional_usable)),
                "Revisiones": sum(v for k, v in state_counts.items() if k.startswith("REVISAR")),
                "Requisitos_genéricos": state_counts.get("REQUISITO_GENERICO", 0),
                "Sin_oferta_en_periodo": sum(v for k, v in state_counts.items() if k.startswith("SIN_OFERTA_EN_PERIODO")),
                "Archivo": coverage.get("Archivo", ""),
                "Fuente_URL": coverage.get("Fuente_URL", ""),
                "Drive_ID": coverage.get("Drive_ID", ""),
            })

    def selection_options(self, carrera: Any, camada: Any = None) -> List[Dict[str, Any]]:
        career = canonical_career(carrera, self.careers)
        cohort = to_int(camada)
        rows = [r for r in self.coverage_operational if norm(r["Carrera"]) == norm(career)]
        if cohort is not None:
            rows = [r for r in rows if to_int(r["Camada"]) == cohort]
        return sorted(rows, key=lambda r: (to_int(r["Camada"]) or 999, r["Sede"], r["Variante"]))

    def _matching_keys(self, carrera: Any, camada: Any = None, sede: Optional[str] = None, variante: Optional[str] = None) -> List[PlanKey]:
        career = canonical_career(carrera, self.careers)
        cohort = to_int(camada)
        keys = [k for k in self.coverage_by_key if norm(k.carrera) == norm(career)]
        if cohort is not None:
            keys = [k for k in keys if k.camada == cohort]
        if sede is not None and clean(sede):
            keys = [k for k in keys if norm(k.sede) == norm(sede)]
        if variante is not None:
            requested = normalize_variant(variante)
            keys = [k for k in keys if k.variante == requested]
        return keys

    def resolve_selection(self, carrera: Any, camada: Any, *, sede: Optional[str] = None, variante: Optional[str] = None, incluir_optativas: bool = False) -> Dict[str, Any]:
        career = canonical_career(carrera, self.careers)
        cohort = to_int(camada)
        candidates = self._matching_keys(career, cohort, sede=sede, variante=variante)
        if not candidates:
            return {
                "Estado": "NO_ENCONTRADO", "Carrera": career, "Camada": cohort, "Códigos": [], "Códigos_obligatorios": [], "Códigos_optativos": [],
                "Revisiones": [], "Mensaje": "No existe una combinación que coincida con esos filtros.", "Opciones": self.selection_options(career, cohort),
            }
        if len(candidates) != 1:
            return {
                "Estado": "AMBIGUO", "Carrera": career, "Camada": cohort, "Códigos": [], "Códigos_obligatorios": [], "Códigos_optativos": [],
                "Revisiones": [], "Mensaje": "Hay más de un plan posible. Usá sede/variante en Más filtros; no se unirán silenciosamente.",
                "Opciones": [
                    {"Carrera": k.carrera, "Camada": k.camada, "Sede": k.sede, "Variante": variant_label(k.variante)}
                    for k in sorted(candidates, key=lambda x: (x.sede, x.variante))
                ],
            }

        key = candidates[0]
        coverage = self.coverage_by_key[key]
        rows = self.operational_by_key.get(key, [])
        mandatory_codes = sorted({r["Código_operativo"] for r in rows if r.get("Código_operativo") and r.get("Estado_periodo") in USABLE_PERIOD_STATES and clean(r.get("Tipo")).upper() not in OPTIONAL_TYPES})
        optional_codes = sorted({r["Código_operativo"] for r in rows if r.get("Código_operativo") and r.get("Estado_periodo") in USABLE_PERIOD_STATES and clean(r.get("Tipo")).upper() in OPTIONAL_TYPES})
        selected_codes = sorted(set(mandatory_codes) | (set(optional_codes) if incluir_optativas else set()))
        reviews = [
            r for r in rows
            if (r.get("Estado_periodo", "").startswith("REVISAR") or r.get("Estado_periodo") == "REQUISITO_GENERICO")
            and (incluir_optativas or clean(r.get("Tipo")).upper() not in OPTIONAL_TYPES)
        ]
        cov = next((x for x in self.coverage_operational if x["Carrera"] == key.carrera and to_int(x["Camada"]) == key.camada and x["Sede"] == key.sede and normalize_variant(x["Variante_raw"]) == key.variante), None)
        scope = clean(coverage.get("Alcance_fuente"))
        if scope != "MALLA_EXTRAIDA":
            status = "FUENTE_INCOMPLETA"
        elif not rows:
            status = "SIN_PERIODO_RESOLUBLE"
        elif selected_codes:
            status = "OK_CON_REVISIONES" if reviews else "OK"
        else:
            status = "SIN_CODIGOS_USABLES"
        return {
            "Estado": status,
            "Periodo": self.period.label,
            "Carrera": key.carrera,
            "Camada": key.camada,
            "Sede": key.sede,
            "Variante": variant_label(key.variante),
            "Variante_raw": key.variante,
            "Plan_version": coverage.get("Plan_version", ""),
            "Alcance_fuente": scope,
            "Metodo_periodo": cov.get("Metodo_periodo", "") if cov else "",
            "Códigos_obligatorios": mandatory_codes,
            "Códigos_optativos": optional_codes,
            "Códigos": selected_codes,
            "Incluir_optativas": incluir_optativas,
            "Revisiones": reviews,
            "Mensaje": "Plan resuelto." if status.startswith("OK") else "El plan no puede utilizarse automáticamente sin revisión adicional.",
            "Opciones": [],
        }

    def program_duration(self, carrera: Any) -> Optional[int]:
        career = canonical_career(carrera, self.careers)
        years = [
            to_int(r.get("Año_académico")) for r in self.plan_rows
            if norm(r.get("Carrera")) == norm(career) and to_int(r.get("Año_académico")) is not None
        ]
        return max(years) if years else None

    def current_cohort_window(self, carrera: Any) -> List[int]:
        duration = self.program_duration(carrera)
        if not duration:
            return sorted({k.camada for k in self.coverage_by_key if norm(k.carrera) == norm(canonical_career(carrera, self.careers)) and k.camada is not None})
        anchor_year = int(self.registry.cohort_anchor["anio"])
        anchor_cohort = int(self.registry.cohort_anchor["camada_ingreso"])
        first_year_cohort = anchor_cohort + (self.period.year - anchor_year)
        return list(range(first_year_cohort - duration + 1, first_year_cohort + 1))

    def _normalize_legacy_code(self, value: Any) -> str:
        code = re.sub(r"\s+", "", clean(value).upper())
        if not code:
            return ""
        if code in self.known_curricular_codes or code in self.current_codes:
            return code
        # Error histórico conocido del archivo auxiliar: letra O usada en lugar de cero.
        # Solo se corrige si la alternativa existe en el maestro o en CURSOS.
        candidates = []
        for i, ch in enumerate(code):
            if ch == "O":
                candidates.append(code[:i] + "0" + code[i + 1:])
        valid = [c for c in candidates if c in self.known_curricular_codes or c in self.current_codes]
        code = valid[0] if len(valid) == 1 else code
        return code if re.fullmatch(r"[A-Z]{1,5}\d{2,3}", code) else ""

    def _legacy_rows_for(self, carrera: Any, cohort: int) -> List[Dict[str, Any]]:
        career = canonical_career(carrera, self.careers)
        out: List[Dict[str, Any]] = []
        for row in self.legacy_rows:
            legacy_career = canonical_career(row.get("CARRERA"), self.careers)
            if norm(legacy_career) != norm(career):
                continue
            if to_int(row.get("CAMADA")) != cohort:
                continue
            out.append(row)
        return out

    @staticmethod
    def _explicit_optional(row: Dict[str, Any]) -> bool:
        text = norm(f"{row.get('Materia', '')} {row.get('Materia_raw', '')} {row.get('Tipo', '')}")
        return "OPTATIVA" in text or "ELECTIVA" in text or clean(row.get("Estado")) == "REQUISITO_GENERICO"

    def _availability_keys(
        self,
        carrera: Any,
        cohort: int,
        *,
        sede: Optional[str] = None,
        variante: Optional[str] = None,
    ) -> List[PlanKey]:
        """Selecciona rutas académicas sin exigir completitud perfecta.

        Reglas de búsqueda simple:
        - si el usuario indicó sede/variante, se respetan;
        - si no indicó sede, Campus tiene prioridad; si no existe, se prefiere
          "No especificada"; si tampoco existe, se conservan las rutas disponibles;
        - si no indicó variante, NO se fuerza Regular: se aprovechan todas las
          variantes de la sede elegida. La disponibilidad se calcula sobre la unión
          de códigos, sin duplicar códigos repetidos.
        """
        career = canonical_career(carrera, self.careers)
        keys = self._matching_keys(career, cohort)
        if not keys:
            return []

        if sede is not None:
            wanted = norm(sede)
            keys = [k for k in keys if norm(k.sede) == wanted]
            if not keys:
                return []
        else:
            campus = [k for k in keys if norm(k.sede) == "CAMPUS"]
            if campus:
                keys = campus
            else:
                unspecified = [k for k in keys if norm(k.sede) in {"NO ESPECIFICADA", "NO ESPECIFICADO", ""}]
                if unspecified:
                    keys = unspecified

        if variante is not None:
            wanted_variant = normalize_variant(variante)
            keys = [k for k in keys if k.variante == wanted_variant]

        # Dedupe defensivo manteniendo orden estable.
        seen: Set[PlanKey] = set()
        out: List[PlanKey] = []
        for key in keys:
            if key not in seen:
                seen.add(key)
                out.append(key)
        return out

    def resolve_availability_selection(
        self,
        carrera: Any,
        camada: Any,
        *,
        sede: Optional[str] = None,
        variante: Optional[str] = None,
        incluir_optativas: bool = False,
    ) -> Dict[str, Any]:
        """Resuelve códigos para disponibilidad estudiantil con criterio amplio.

        Para cada camada toma el conjunto curricular del año académico completo
        (Otoño + Primavera). CURSOS del período decide después cuáles de esos códigos
        tienen oferta real. Una fuente parcial o una variante incompleta nunca hace
        desaparecer la camada: se usa todo código confiable disponible y se advierte
        lo que no pudo resolverse.
        """
        career = canonical_career(carrera, self.careers)
        cohort = to_int(camada)
        if cohort is None:
            return {
                "Estado": "CAMADA_INVALIDA", "Carrera": career, "Camada": camada,
                "Códigos": [], "Códigos_ofertados": [], "Códigos_sin_oferta": [],
                "Revisiones": [], "Fuentes_académicas": [], "Mensaje": "No se pudo interpretar la camada.", "Opciones": [],
            }

        keys = self._availability_keys(career, cohort, sede=sede, variante=variante)
        academic_year = self._cohort_academic_year(cohort)
        master_rows: List[Dict[str, Any]] = []
        for key in keys:
            master_rows.extend([
                r for r in self.plan_by_key.get(key, [])
                if academic_year is not None and to_int(r.get("Año_académico")) == academic_year
            ])

        used_codes: Set[str] = set()
        candidate_codes: Set[str] = set()
        reviews: List[Dict[str, Any]] = []
        sources: Set[str] = set()

        for row in master_rows:
            status = clean(row.get("Estado"))
            code = clean(row.get("Código")).upper()
            is_optional = self._explicit_optional(row)
            if is_optional and not incluir_optativas:
                continue
            if code:
                candidate_codes.add(code)
                used_codes.add(code)
                sources.add("PLAN_MASTER")
                continue

            candidates = [c for c in extract_codes(row.get("Códigos_detectados")) if c in self.known_curricular_codes or c in self.current_codes]
            candidate_codes.update(candidates)
            active = sorted(set(candidates) & self.current_codes)
            if status == "REVISAR_MULTICODIGO":
                if len(active) == 1:
                    used_codes.add(active[0])
                    sources.add("PLAN_MASTER")
                elif len(active) > 1:
                    reviews.append({"Materia": row.get("Materia", ""), "Estado": "REVISAR_MULTICODIGO_ACTIVO", "Candidatos": active})
                continue

            if status == "REVISAR_SIN_CODIGO":
                name_codes = sorted(self.name_to_codes.get(normalize_course_name(row.get("Materia")), set()))
                if len(name_codes) == 1:
                    used_codes.add(name_codes[0])
                    candidate_codes.add(name_codes[0])
                    sources.add("PLAN_MASTER")
                elif len(name_codes) > 1:
                    reviews.append({"Materia": row.get("Materia", ""), "Estado": "REVISAR_NOMBRE_MULTIPLE", "Candidatos": name_codes})
                else:
                    reviews.append({"Materia": row.get("Materia", ""), "Estado": "REVISAR_SIN_CODIGO", "Candidatos": []})
                continue

            if status == "REQUISITO_GENERICO" and incluir_optativas:
                reviews.append({"Materia": row.get("Materia", ""), "Estado": "REQUISITO_GENERICO", "Candidatos": []})

        # El catálogo auxiliar reproduce el conjunto amplio de códigos usado por el
        # sistema original. Se usa como suplemento incluso cuando el PLAN_MASTER es
        # parcial, carece de malla o no permite resolver una variante concreta.
        legacy_rows = self._legacy_rows_for(career, cohort)
        for row in legacy_rows:
            code = self._normalize_legacy_code(row.get("CODIGO"))
            if not code:
                continue
            candidate_codes.add(code)
            used_codes.add(code)
            sources.add("CATALOGO_ACADEMICO_AUXILIAR")

        # Evitamos repetir la misma advertencia si el mismo requisito aparece en
        # varias versiones/semestres del plan.
        deduped_reviews: List[Dict[str, Any]] = []
        seen_reviews: Set[str] = set()
        for review in reviews:
            tokens = []
            for token in normalize_course_name(review.get("Materia")).split():
                if token in {"DE", "DEL", "LA", "EL", "LOS", "LAS"}:
                    continue
                if len(token) > 4 and token.endswith("S"):
                    token = token[:-1]
                tokens.append(token)
            concept = " ".join(tokens) or clean(review.get("Estado"))
            if concept in seen_reviews:
                continue
            seen_reviews.add(concept)
            deduped_reviews.append(review)
        reviews = deduped_reviews

        offered = sorted(used_codes & self.current_codes)
        not_offered = sorted(used_codes - self.current_codes)
        selected_sites = sorted({k.sede for k in keys})
        selected_variants = sorted({variant_label(k.variante) for k in keys})

        if used_codes:
            state = "OK" if not reviews else "OK_CON_REVISIONES"
            message = "Se usó todo código confiable disponible para la camada; CURSOS del período decide cuáles tienen oferta."
        elif master_rows or keys:
            state = "PARCIAL_SIN_CODIGOS"
            message = "La camada está presente pero no se identificaron códigos concretos suficientes para calcular disponibilidad. Se conserva y se advierte."
        else:
            state = "SIN_FUENTE_CODIFICADA"
            message = "La camada se mantiene en la ventana vigente, pero no hay una fuente codificada suficiente para calcular disponibilidad."

        return {
            "Estado": state,
            "Carrera": career,
            "Camada": cohort,
            "Sede": " | ".join(selected_sites) if selected_sites else (clean(sede) if sede else "No especificada"),
            "Variante": " | ".join(selected_variants) if selected_variants else (variant_label(variante) if variante else "Automática"),
            "Año_académico": academic_year,
            "Códigos": sorted(used_codes),
            "Códigos_ofertados": offered,
            "Códigos_sin_oferta": not_offered,
            "Códigos_candidatos": sorted(candidate_codes),
            "Incluir_optativas": incluir_optativas,
            "Revisiones": reviews,
            "Fuentes_académicas": sorted(sources),
            "Mensaje": message,
            "Opciones": [
                {"Sede": k.sede, "Variante": variant_label(k.variante)} for k in keys
            ],
        }

    def active_cohorts(self, carrera: Any, *, sede: Optional[str] = None, variante: Optional[str] = None, incluir_optativas: bool = False) -> List[int]:
        """Camadas vigentes que aportan al menos un código. Se conserva por compatibilidad."""
        career = canonical_career(carrera, self.careers)
        active: List[int] = []
        for cohort in self.current_cohort_window(career):
            result = self.resolve_availability_selection(career, cohort, sede=sede, variante=variante, incluir_optativas=incluir_optativas)
            if result.get("Códigos"):
                active.append(cohort)
        return active

    def visible_cohorts(self, carrera: Any) -> List[int]:
        """Todas las camadas de la ventana vigente, aunque la fuente sea parcial."""
        return self.current_cohort_window(carrera)

    def inactive_or_problem_cohorts(self, carrera: Any, *, sede: Optional[str] = None, variante: Optional[str] = None) -> List[Dict[str, Any]]:
        career = canonical_career(carrera, self.careers)
        rows = []
        for cohort in self.current_cohort_window(career):
            result = self.resolve_availability_selection(career, cohort, sede=sede, variante=variante)
            if not result.get("Códigos"):
                rows.append({"Camada": cohort, "Estado": result.get("Estado", ""), "Mensaje": result.get("Mensaje", "")})
        return rows



class UdeSAService:
    """Contrato que consumirá la futura web. Los filtros avanzados son opcionales."""

    def __init__(self, registry: SourceRegistry, period: AcademicPeriod, engine: r3.UdeSAEngine, resolver: AcademicResolver45) -> None:
        self.registry = registry
        self.period = period
        self.engine = engine
        self.resolver = resolver
        self.default_sede = clean(registry.defaults.get("sede")) or "Campus"
        self.default_variant = clean(registry.defaults.get("variante")) or "Regular"
        self.default_include_optatives = bool(registry.defaults.get("incluir_optativas", False))

    def listar_periodos(self) -> List[Dict[str, Any]]:
        out = []
        for pid in self.registry.list_period_ids():
            raw = self.registry.period_defs[pid]
            out.append({
                "ID": pid,
                "Año": int(raw["anio"]),
                "Semestre": term_label(raw["semestre"]),
                "Activo": pid == self.registry.active_period_id,
            })
        return out

    def listar_carreras(self) -> List[str]:
        return list(self.resolver.careers)

    def listar_camadas(self, carrera: Any, *, sede: Optional[str] = None, variante: Optional[str] = None, solo_activas: bool = True) -> List[int]:
        if solo_activas:
            # "Activas" en interfaz significa vigentes por cohorte, aunque alguna fuente sea parcial.
            return self.resolver.visible_cohorts(carrera)
        career = canonical_career(carrera, self.resolver.careers)
        return sorted({k.camada for k in self.resolver.coverage_by_key if norm(k.carrera) == norm(career) and k.camada is not None})

    def listar_sedes(self, carrera: Any, camada: Any = None) -> List[str]:
        keys = self.resolver._matching_keys(carrera, camada)
        return sorted({k.sede for k in keys})

    def listar_variantes(self, carrera: Any, camada: Any = None, *, sede: Optional[str] = None) -> List[str]:
        keys = self.resolver._matching_keys(carrera, camada, sede=sede)
        return sorted({variant_label(k.variante) for k in keys})

    def _academic_path_options(self, carrera: Any) -> List[Dict[str, Any]]:
        career = canonical_career(carrera, self.resolver.careers)
        groups: Dict[Tuple[str, str], Set[int]] = defaultdict(set)
        for cohort in self.resolver.current_cohort_window(career):
            for key in self.resolver._matching_keys(career, cohort):
                result = self.resolver.resolve_selection(career, cohort, sede=key.sede, variante=variant_label(key.variante), incluir_optativas=False)
                if result.get("Estado", "").startswith("OK") and result.get("Códigos"):
                    groups[(key.sede, variant_label(key.variante))].add(cohort)
        return [
            {"Sede": sede, "Variante": variante, "Camadas_operativas": sorted(cohorts)}
            for (sede, variante), cohorts in sorted(groups.items())
        ]

    def _resolve_academic_path(self, carrera: Any, sede: Optional[str], variante: Optional[str]) -> Dict[str, Any]:
        options = self._academic_path_options(carrera)
        if sede is not None or variante is not None:
            requested_sede = clean(sede) if sede is not None else None
            requested_variant = variant_label(variante) if variante is not None else None
            filtered = [o for o in options if (requested_sede is None or norm(o["Sede"]) == norm(requested_sede)) and (requested_variant is None or norm(o["Variante"]) == norm(requested_variant))]
            if len(filtered) == 1:
                return {"Estado": "OK", **filtered[0], "Opciones": options}
            if len(filtered) > 1:
                # Si falta solo la variante, Regular es la preferida cuando existe; si falta solo sede, Campus es la preferida.
                if requested_variant is None:
                    regular = [o for o in filtered if o["Variante"] == "Regular"]
                    if len(regular) == 1:
                        return {"Estado": "OK", **regular[0], "Opciones": options}
                if requested_sede is None:
                    campus = [o for o in filtered if norm(o["Sede"]) == "CAMPUS"]
                    if len(campus) == 1:
                        return {"Estado": "OK", **campus[0], "Opciones": options}
                return {"Estado": "NECESITA_FILTRO", "Opciones": filtered}
            return {"Estado": "NO_ENCONTRADO", "Opciones": options}

        # Sin filtros avanzados: Campus + Regular primero.
        preferred = [o for o in options if norm(o["Sede"]) == "CAMPUS" and o["Variante"] == "Regular"]
        if len(preferred) == 1:
            return {"Estado": "OK", **preferred[0], "Opciones": options}
        # Si no existe esa combinación, una única ruta académica puede usarse sin molestar al usuario.
        if len(options) == 1:
            return {"Estado": "OK", **options[0], "Opciones": options}
        return {"Estado": "NECESITA_FILTRO" if options else "SIN_RUTA_OPERATIVA", "Opciones": options}

    def resolver_poblacion_simple(
        self,
        carrera: Any,
        *,
        camadas: Any = "TODAS",
        sede: Optional[str] = None,
        variante: Optional[str] = None,
        incluir_optativas: Optional[bool] = None,
        codigos_extra: Optional[Sequence[str]] = None,
    ) -> Dict[str, Any]:
        include_opt = self.default_include_optatives if incluir_optativas is None else bool(incluir_optativas)
        career = canonical_career(carrera, self.resolver.careers)
        # En la búsqueda simple sede y variante quedan automáticas. La resolución se hace
        # por camada: Campus se prioriza si existe y las variantes se complementan entre sí.
        # Solo se restringen cuando el usuario eligió explícitamente un filtro avanzado.
        if isinstance(camadas, str) and norm(camadas) in {"TODAS", "TODOS", "ALL", ""}:
            cohort_list = self.resolver.visible_cohorts(career)
        elif isinstance(camadas, (list, tuple, set)):
            cohort_list = [x for x in (to_int(v) for v in camadas) if x is not None]
        else:
            c = to_int(camadas)
            cohort_list = [c] if c is not None else []

        resolutions: List[Dict[str, Any]] = []
        codes: Set[str] = set()
        offered_codes: Set[str] = set()
        not_offered_codes: Set[str] = set()
        academic_sources: Set[str] = set()
        warnings: List[str] = []
        if not cohort_list:
            warnings.append(f"{career}: no hay camadas operativas para {sede} / {variant_label(variante)} en {self.period.label} con los filtros predeterminados.")
        for cohort in sorted(set(cohort_list)):
            result = self.resolver.resolve_availability_selection(career, cohort, sede=sede, variante=variante, incluir_optativas=include_opt)
            resolutions.append(result)
            if result.get("Estado", "").startswith("OK"):
                codes.update(result.get("Códigos", []))
                offered_codes.update(result.get("Códigos_ofertados", []))
                not_offered_codes.update(result.get("Códigos_sin_oferta", []))
                academic_sources.update(result.get("Fuentes_académicas", []))
                if result.get("Revisiones"):
                    warnings.append(f"{career} {cohort}: {len(result['Revisiones'])} requisito(s) no automatizados.")
            else:
                warnings.append(f"{career} {cohort}: {result.get('Estado')} — {result.get('Mensaje')} La camada permanece incluida en la selección.")

        extras = sorted({clean(c).upper() for c in (codigos_extra or []) if clean(c)})
        codes.update(extras)
        offered_codes.update(c for c in extras if c in self.resolver.current_codes)
        not_offered_codes.update(c for c in extras if c not in self.resolver.current_codes)
        resolved_sites = sorted({part.strip() for r in resolutions for part in clean(r.get("Sede")).split("|") if part.strip()})
        resolved_variants = sorted({part.strip() for r in resolutions for part in clean(r.get("Variante")).split("|") if part.strip() and part.strip() != "Automática"})
        site_display = clean(sede) if sede is not None else (resolved_sites[0] if len(resolved_sites) == 1 else ("Automática (" + " + ".join(resolved_sites) + ")" if resolved_sites else "Automática"))
        variant_display = variant_label(variante) if variante is not None else ("Automática (" + " + ".join(resolved_variants) + ")" if resolved_variants else "Automática")
        return {
            "Estado": "OK" if codes else "SIN_CODIGOS",
            "Periodo": self.period.label,
            "Carrera": career,
            "Camadas": sorted(set(cohort_list)),
            "Sede": site_display,
            "Variante": variant_display,
            "Incluir_optativas": include_opt,
            "Códigos": sorted(codes),
            "Códigos_ofertados": sorted(offered_codes),
            "Códigos_sin_oferta": sorted(not_offered_codes),
            "Códigos_extra": extras,
            "Fuentes_académicas": sorted(academic_sources),
            "Resoluciones": resolutions,
            "Advertencias": warnings,
            "Opciones": [opt for r in resolutions for opt in r.get("Opciones", [])],
        }

    def resolver_poblaciones(self, poblaciones: Sequence[Dict[str, Any]]) -> Dict[str, Any]:
        codes: Set[str] = set()
        resolutions = []
        warnings: List[str] = []
        for pop in poblaciones:
            r = self.resolver_poblacion_simple(
                pop.get("Carrera"),
                camadas=pop.get("Camadas", "TODAS"),
                sede=pop.get("Sede"),
                variante=pop.get("Variante"),
                incluir_optativas=pop.get("Incluir_optativas"),
                codigos_extra=pop.get("Códigos_extra", []),
            )
            resolutions.append(r)
            codes.update(r.get("Códigos", []))
            warnings.extend(r.get("Advertencias", []))
        return {"Códigos": sorted(codes), "Poblaciones": resolutions, "Advertencias": warnings}

    def disponibilidad_estudiantil(self, carrera: Any, date_value: Any, *, slot: Any = None, start: Any = None, end: Any = None, **advanced: Any) -> Dict[str, Any]:
        population = self.resolver_poblacion_simple(carrera, **advanced)
        if not population["Códigos"]:
            return {"Estado": "SIN_CODIGOS", **population}
        result = self.engine.student_occupancy(population["Códigos"], date_value, slot=slot, start=start, end=end)
        result.update({
            "Estado": "OK" if not population["Advertencias"] else "OK_CON_ADVERTENCIAS",
            "Población": population,
            "Advertencias": population["Advertencias"],
        })
        return result

    def buscar_mejor_horario(
        self,
        poblaciones: Sequence[Dict[str, Any]],
        dates: Sequence[Any],
        *,
        capacidad_minima: int = 0,
        sede_aulas: Optional[str] = None,
        slots: Optional[Sequence[Any]] = None,
        top_n: Optional[int] = None,
    ) -> Dict[str, Any]:
        resolved = self.resolver_poblaciones(poblaciones)
        if not resolved["Códigos"]:
            return {"Estado": "SIN_CODIGOS", "Ranking": [], **resolved}
        room_sede = sede_aulas or self.default_sede
        ranking = self.engine.combined_ranking(
            resolved["Códigos"], dates, sede=room_sede, min_capacity=capacidad_minima, slots=slots, top_n=top_n
        )
        for row in ranking:
            row["Período"] = self.period.label
            row["Códigos derivados"] = " | ".join(resolved["Códigos"])
            row["Advertencias"] = " | ".join(resolved["Advertencias"])
        return {
            "Estado": "OK" if not resolved["Advertencias"] else "OK_CON_ADVERTENCIAS",
            "Periodo": self.period.label,
            "Ranking": ranking,
            **resolved,
        }



def write_csv(path: Path, rows: Sequence[Dict[str, Any]], fieldnames: Optional[Sequence[str]] = None) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    rows = list(rows)
    if fieldnames is None:
        fields: List[str] = []
        seen: Set[str] = set()
        for row in rows:
            for key in row:
                if key not in seen:
                    seen.add(key)
                    fields.append(key)
        fieldnames = fields or ["Sin datos"]
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(fieldnames), extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            out: Dict[str, Any] = {}
            for key in fieldnames:
                value = row.get(key, "")
                if isinstance(value, (list, tuple, set)):
                    value = " | ".join(str(x) for x in value)
                elif isinstance(value, dict):
                    value = json.dumps(value, ensure_ascii=False)
                out[key] = value
            writer.writerow(out)


def build_runtime(base: Path, config_path: Optional[Path] = None, period_id: Optional[str] = None) -> Tuple[SourceRegistry, AcademicPeriod, r3.UdeSAEngine, AcademicResolver45, UdeSAService, Dict[str, Any]]:
    registry = SourceRegistry(base, config_path=config_path)
    pid = registry.active_period_id if period_id in (None, "", "ACTUAL") else clean(period_id).upper()
    paths = registry.source_paths(pid)
    provisional = registry.provisional_period(pid)

    # El año de los comentarios sin año se define por período, no por una constante permanente.
    r2.YEAR = provisional.year

    problems: List[Dict[str, Any]] = []
    catalog = r2.MasterCatalog()
    r2.process_aux_catalog(paths["catalogo_aux"], catalog, problems)
    r2.process_courses(paths["cursos"], catalog, problems)
    r2.process_aulas(paths["aulas"], catalog, problems)
    r2.add_manual_spaces(catalog)
    for alias, key in r2.MANUAL_ALIASES.items():
        if key in catalog.entries:
            catalog.add_alias(key, alias.title() if alias != "HAM" else alias)

    # Correcciones manuales explícitas del proyecto tienen prioridad máxima y quedan auditables.
    for room, capacity in dict(registry.manual_adjustments.get("capacidades", {})).items():
        try:
            cap = int(capacity)
        except Exception:
            continue
        key = r2.normalize_room_key(room)
        if key:
            catalog.set_capacity(key, cap, "Corrección manual validada", priority=1000)

    class_stats = r2.process_regular_classes_r2(paths["cursos"], catalog)
    period = registry.build_period_from_classes(pid, class_stats["rows"])

    # El motor heredado usa estos límites como contexto de vigencia; se actualizan al período cargado.
    r3.DEFAULT_SEMESTER_START = period.start
    r3.DEFAULT_SEMESTER_END = period.end

    slot_map = class_stats["slot_map"]
    course_names = class_stats["course_name_map"]
    known_codes = set(course_names)
    grid_stats = r2.process_aulas_grid_r2(paths["aulas"], catalog, slot_map, known_codes, course_names)
    comment_stats = r2.process_comments_r2(paths["aulas"], catalog, slot_map, known_codes)
    reservation_stats = r2.process_special_reservations_r2(paths["aulas"], catalog)
    event_stats = r2.process_events_r2(paths["eventos"], catalog)

    engine = r3.UdeSAEngine(
        catalog.rows(), class_stats["rows"], grid_stats["rows"], comment_stats["items"], comment_stats["effects"],
        reservation_stats["rows"], event_stats["occupancies"], slot_map,
    )
    resolver = AcademicResolver45(paths["plan_master"], engine, period, registry, paths.get("catalogo_academico_legacy"))
    service = UdeSAService(registry, period, engine, resolver)
    stats = {
        "paths": {k: str(v) for k, v in paths.items()},
        "catalog_rows": len(catalog.rows()),
        "classes": len(class_stats["rows"]),
        "grid": len(grid_stats["rows"]),
        "comments": comment_stats["comments_count"],
        "comment_items": len(comment_stats["items"]),
        "comment_effects": len(comment_stats["effects"]),
        "reservations": len(reservation_stats["rows"]),
        "events": len(event_stats["events"]),
        "event_occupancies": len(event_stats["occupancies"]),
    }
    return registry, period, engine, resolver, service, stats


def flatten_simple_population(pop: Dict[str, Any]) -> Dict[str, Any]:
    return {
        "Estado": pop.get("Estado", ""),
        "Período": pop.get("Periodo", ""),
        "Carrera": pop.get("Carrera", ""),
        "Camadas": " | ".join(str(x) for x in pop.get("Camadas", [])),
        "Sede": pop.get("Sede", ""),
        "Variante": pop.get("Variante", ""),
        "Incluir optativas": "Sí" if pop.get("Incluir_optativas") else "No",
        "Códigos": " | ".join(pop.get("Códigos", [])),
        "Cantidad códigos": len(pop.get("Códigos", [])),
        "Advertencias": " | ".join(pop.get("Advertencias", [])),
    }


def run_tests(registry: SourceRegistry, period: AcademicPeriod, engine: r3.UdeSAEngine, resolver: AcademicResolver45, service: UdeSAService) -> List[Dict[str, Any]]:
    tests: List[Dict[str, Any]] = []

    def check(name: str, expected: Any, actual: Any, detail: str = "") -> None:
        tests.append({"Test": name, "Esperado": expected, "Obtenido": actual, "Resultado": "OK" if expected == actual else "ERROR", "Detalle": detail})

    # Períodos y configuración.
    check("Período activo: ID", "2026_PRIMAVERA", registry.active_period_id)
    check("Período activo: semestre", "PRIMAVERA", period.term)
    check("Período activo: número semestre", 2, period.term_number)
    check("Período activo: rango inferido desde CURSOS", "2026-08-03|2026-12-12", f"{period.start.isoformat()}|{period.end.isoformat()}")
    synthetic_autumn = AcademicPeriod(2027, "OTONO", dt.date(2027, 3, 1), dt.date(2027, 7, 31))
    check("Generalización: Otoño es semestre 1", 1, synthetic_autumn.term_number)
    check("Generalización: ID futuro no contiene 2026", "2027_OTONO", synthetic_autumn.id)
    # Cohorte: 2026->38 primer año; 2027->39 primer año.
    original_period = resolver.period
    resolver.period = synthetic_autumn
    check("Generalización cohorte: camada 39 es primer año en 2027", 1, resolver._cohort_academic_year(39))
    resolver.period = original_period

    # Semántica de oferta: código curricular válido sin oferta NO se invalida.
    sin_oferta = [r for r in resolver.operational_rows if r.get("Estado_periodo") == "SIN_OFERTA_EN_PERIODO"]
    check("Código curricular sin oferta queda diferenciado", True, len(sin_oferta) > 0)
    check("Código sin oferta conserva Estado curricular", True, all(r.get("Estado_curricular") == "VALIDADO_DESDE_PLAN" for r in sin_oferta))

    # Defaults UX: RRII sola -> todas las camadas activas, Campus, Regular, obligatorias.
    ri_simple = service.resolver_poblacion_simple("RRII")
    check("Simple: RRII usa Campus por defecto", "Campus", ri_simple["Sede"], json.dumps(flatten_simple_population(ri_simple), ensure_ascii=False))
    check("Simple: RRII usa variante automática", True, ri_simple["Variante"].startswith("Automática"))
    check("Simple: optativas apagadas", False, ri_simple["Incluir_optativas"])
    check("Simple: RRII incluye todas las camadas vigentes", [35, 36, 37, 38], ri_simple["Camadas"], json.dumps(flatten_simple_population(ri_simple), ensure_ascii=False))
    check("Simple: RRII 38 incompleta se conserva", True, 38 in ri_simple["Camadas"] and "H250" in ri_simple["Códigos"])
    check("Simple: no molesta con camadas históricas", False, any(any(f" {c} " in f" {w} " for c in range(30,35)) for w in ri_simple["Advertencias"]))
    check("Simple: optativas genéricas apagadas no generan advertencia", False, any("requisito(s) no automatizados" in w for w in ri_simple["Advertencias"]))

    # Avanzado: camada específica.
    ri36 = service.resolver_poblacion_simple("RRII", camadas=[36])
    check("Avanzado: filtra solo camada 36", [36], ri36["Camadas"])
    check("Avanzado: RI36 contiene P318", True, "P318" in ri36["Códigos"])

    # Optativas: Diseño 37 tiene optativas concretas identificadas en el período.
    diseno37_base = service.resolver_poblacion_simple("Diseño", camadas=[37], incluir_optativas=False)
    diseno37_opt = service.resolver_poblacion_simple("Diseño", camadas=[37], incluir_optativas=True)
    check("Optativas: apagadas no incluyen CC101", False, "CC101" in diseno37_base["Códigos"])
    check("Optativas: activadas incluyen CC101", True, "CC101" in diseno37_opt["Códigos"])
    check("Optativas: activar nunca reduce códigos", True, len(diseno37_opt["Códigos"]) >= len(diseno37_base["Códigos"]))
    cp36_base = service.resolver_poblacion_simple("CP", camadas=[36], incluir_optativas=False)
    cp36_opt = service.resolver_poblacion_simple("CP", camadas=[36], incluir_optativas=True)
    check("Optativas: requisito genérico no molesta cuando están apagadas", False, any("requisito(s) no automatizados" in w for w in cp36_base["Advertencias"]))
    check("Optativas: requisito genérico se advierte cuando están activadas", True, any("requisito(s) no automatizados" in w for w in cp36_opt["Advertencias"]))

    # Filtros de opciones para futura UI.
    careers = service.listar_carreras()
    check("Contrato web: lista carreras", True, "Relaciones Internacionales" in careers and len(careers) >= 19)
    check("Contrato web: lista camadas vigentes", [35, 36, 37, 38], service.listar_camadas("RRII"))
    check("Contrato web: lista sedes", True, "Campus" in service.listar_sedes("RRII", 37))
    check("Contrato web: lista variantes", True, "Regular" in service.listar_variantes("RRII", 37, sede="Campus"))
    humanidades = service.resolver_poblacion_simple("Humanidades")
    check("Simple: ruta académica única se elige sin mostrar filtro", "No especificada", humanidades.get("Sede"), json.dumps(flatten_simple_population(humanidades), ensure_ascii=False))
    comportamiento = service.resolver_poblacion_simple("Ciencias del Comportamiento")
    check("Simple: rutas académicas múltiples no bloquean por defecto", "OK", comportamiento.get("Estado"))
    check("Simple: ambigüedad ofrece opciones", True, len(comportamiento.get("Opciones", [])) >= 2)

    # Regresión académica de Ronda 4.
    cp36_reg = resolver.resolve_selection("CP", 36, sede="Campus", variante="Regular", incluir_optativas=False)
    check("Regresión R4: CP36 resuelve", True, cp36_reg.get("Estado", "").startswith("OK"))
    check("Regresión R4: CP36 incluye E020/P318/P328", True, {"E020","P318","P328"}.issubset(set(cp36_reg.get("Códigos", []))))
    ri37_reg = resolver.resolve_selection("RRII", 37, sede="Campus", variante="Regular")
    check("Regresión R4: RI37 resuelve H021", True, "H021" in ri37_reg.get("Códigos", []))
    admin38 = resolver.resolve_selection("Administración de Empresas", 38)
    check("Regresión R4: Admin38 sin filtros sigue ambiguo", "AMBIGUO", admin38.get("Estado"))
    ri38_regular = resolver.resolve_selection("RRII", 38, sede="Campus", variante="Regular")
    check("Regresión R4: RI38 regular conserva fuente incompleta", "FUENTE_INCOMPLETA", ri38_regular.get("Estado"))
    ri38_aug = resolver.resolve_selection("RRII", 38, sede="Campus", variante="Ingreso agosto")
    check("Regresión R4: RI38 ingreso agosto contiene P050", True, "P050" in ri38_aug.get("Códigos", []))

    # Multi-población y ranking final.
    combined = service.buscar_mejor_horario(
        [{"Carrera": "RRII"}, {"Carrera": "CP", "Camadas": [36]}],
        ["19/08/2026", "20/08/2026", "21/08/2026"],
        capacidad_minima=50, top_n=5,
    )
    check("Simple multi-carrera: genera ranking", True, len(combined.get("Ranking", [])) > 0)
    check("Simple multi-carrera: conserva alumnos y aulas", True, all("Alumnos ocupados" in r and "Aulas disponibles" in r for r in combined.get("Ranking", [])))

    # Códigos manuales avanzados.
    extra = service.resolver_poblacion_simple("RRII", camadas=[36], codigos_extra=["P328"])
    check("Avanzado: acepta código manual extra", True, "P328" in extra["Códigos"])

    # Estado de fuentes y frecuencia.
    status_rows = registry.source_status_rows(period.id)
    check("Fuentes: cinco fuentes registradas", 5, len(status_rows))
    check("Fuentes: Eventos figura anual", True, any(r["Fuente"] == "Eventos" and r["Frecuencia"].startswith("Anual") for r in status_rows))
    check("Fuentes: AULAS figura semestral y continua", True, any(r["Fuente"] == "AULAS" and "continua" in r["Frecuencia"] for r in status_rows))

    # Regresión del motor físico anterior.
    old_tests = r3.run_tests(engine)
    check("Regresión Ronda 3: 27 tests siguen OK", 27, sum(1 for t in old_tests if t.get("Resultado") == "OK"))
    check("Regresión Ronda 3: sin errores", 0, sum(1 for t in old_tests if t.get("Resultado") == "ERROR"))

    return tests


def make_report(registry: SourceRegistry, period: AcademicPeriod, resolver: AcademicResolver45, service: UdeSAService, source_stats: Dict[str, Any], tests: Sequence[Dict[str, Any]], simple_examples: Sequence[Dict[str, Any]], combined: Dict[str, Any]) -> str:
    row_counts = Counter(r.get("Estado_periodo", "") for r in resolver.operational_rows)
    coverage_counts = Counter(r.get("Estado_operativo", "") for r in resolver.coverage_operational)
    method_counts = Counter(r.get("Metodo_periodo", "") for r in resolver.coverage_operational)
    lines: List[str] = []
    lines.append("=" * 96)
    lines.append("UdeSA Horarios — DIAGNÓSTICO RONDA 4.5")
    lines.append(f"Autor: {AUTHOR}")
    lines.append(f"Generado: {dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append("=" * 96)
    lines.append("")
    lines.append("ALCANCE")
    lines.append("Generalización de períodos académicos — IMPLEMENTADA")
    lines.append("Registro de fuentes reemplazables — IMPLEMENTADO")
    lines.append("Semántica Otoño=1 / Primavera=2 — IMPLEMENTADA")
    lines.append("Separación validez curricular vs oferta en período — IMPLEMENTADA")
    lines.append("Defaults simples + filtros avanzados opcionales — IMPLEMENTADOS")
    lines.append("Contrato de funciones para la futura web — IMPLEMENTADO")
    lines.append("Automatización remota de Drive/Sheets — RESERVADA PARA RONDA 6")
    lines.append("")

    lines.append("PERÍODO ACTIVO")
    lines.append(f"ID: {period.id}")
    lines.append(f"Período: {period.label}")
    lines.append(f"Semestre numérico: {period.term_number}")
    lines.append(f"Vigencia física inferida/configurada: {period.start.isoformat()} a {period.end.isoformat()}")
    lines.append("El motor ya no contiene Primavera 2026 como regla fija; el período proviene de fuentes_udesahorarios.json.")
    lines.append("")

    lines.append("FUENTES")
    for row in registry.source_status_rows(period.id):
        lines.append(f"  - {row['Fuente']}: {row['Frecuencia']} | {row['Período']} | {row['Archivo']} | {row['Estado']}")
    lines.append("")
    lines.append("REGLA DE ACTUALIZACIÓN")
    lines.append("CURSOS y AULAS pertenecen a un período semestral; Eventos pertenece al año; planes y catálogo de espacios son transversales.")
    lines.append("Para un semestre futuro se agrega/cambia configuración de fuentes; no se modifica el motor Python.")
    lines.append("")

    lines.append("CAPA ACADÉMICA GENERALIZADA")
    lines.append(f"Filas PLAN_MASTER: {len(resolver.plan_rows)}")
    lines.append(f"Combinaciones COBERTURA: {len(resolver.coverage_rows)}")
    lines.append(f"Filas del plan operativo para {period.label}: {len(resolver.operational_rows)}")
    lines.append("Métodos de selección de período:")
    for k, v in method_counts.most_common():
        lines.append(f"  - {k}: {v}")
    lines.append("Estados por período:")
    for k, v in row_counts.most_common():
        lines.append(f"  - {k}: {v}")
    lines.append("")
    lines.append("IMPORTANTE SOBRE CÓDIGOS")
    lines.append("Un código validado por el plan que no aparece en CURSOS del período queda como SIN_OFERTA_EN_PERIODO.")
    lines.append("No se considera inválido ni se elimina del maestro. Un multicódigo resuelto por oferta única se resuelve solo para ese período.")
    lines.append("")

    lines.append("DEFAULTS DE LA FUTURA WEB")
    lines.append(f"Período: ACTUAL ({period.label})")
    lines.append("Ruta académica: Campus se prioriza cuando existe; las variantes se aprovechan automáticamente y solo restringen si el usuario las elige")
    lines.append(f"Sede de AULAS para búsquedas físicas: {service.default_sede}")
    lines.append("Camadas: TODAS LAS VIGENTES DEL PERÍODO, incluso si alguna fuente es parcial")
    lines.append("Materias: obligatorias/estructurales; optativas y electivas apagadas por defecto")
    lines.append("Los filtros de camada, sede, variante, período, optativas y códigos manuales quedan para '+ Más filtros'.")
    lines.append("")

    lines.append("EJEMPLOS DE BÚSQUEDA SIMPLE / AVANZADA")
    for example in simple_examples:
        lines.append(
            f"  - {example['Carrera']} | camadas={example['Camadas']} | sede={example['Sede']} | variante={example['Variante']} | "
            f"optativas={example['Incluir optativas']} | códigos={example['Cantidad códigos']}"
        )
        if example.get("Advertencias"):
            lines.append(f"    Advertencias: {example['Advertencias']}")
    lines.append("")

    lines.append("EJEMPLO DE RANKING SIMPLE")
    lines.append("Poblaciones: RRII (todas las camadas operativas) + CP 36 | Campus | obligatorias | capacidad mínima 50")
    for i, row in enumerate(combined.get("Ranking", [])[:5], 1):
        lines.append(
            f"  {i}. {row['Fecha']} {row['Día']} Slot {row['Slot']} {row['Hora Desde']}–{row['Hora Hasta']} | "
            f"alumnos ocupados={row['Alumnos ocupados']} | materias afectadas={row['Materias afectadas']} | aulas disponibles={row['Aulas disponibles']}"
        )
    lines.append("")

    lines.append("COBERTURA OPERATIVA")
    for k, v in coverage_counts.most_common():
        lines.append(f"  - {k}: {v}")
    lines.append("")

    ok = sum(1 for t in tests if t["Resultado"] == "OK")
    errors = sum(1 for t in tests if t["Resultado"] == "ERROR")
    lines.append("PRUEBAS AUTOMÁTICAS RONDA 4.5")
    lines.append(f"Tests OK: {ok}")
    lines.append(f"Tests con error: {errors}")
    for t in tests:
        lines.append(f"  - {t['Resultado']}: {t['Test']} | esperado={t['Esperado']} | obtenido={t['Obtenido']}")
    lines.append("")

    lines.append("CONTRATO PARA LA WEB")
    lines.append("Funciones disponibles: listar_periodos, listar_carreras, listar_camadas, listar_sedes, listar_variantes,")
    lines.append("resolver_poblacion_simple, disponibilidad_estudiantil y buscar_mejor_horario.")
    lines.append("La interfaz podrá mostrar solo Carrera + Buscar y desplegar '+ Más filtros' sin cambiar el motor.")
    return "\n".join(lines)


def main() -> None:
    base = Path(__file__).resolve().parent
    output = base / OUTPUT_DIR_NAME
    output.mkdir(parents=True, exist_ok=True)

    print(f"\n{APP_NAME} — Ronda 4.5")
    print(f"Autor: {AUTHOR}")
    print(f"Carpeta de trabajo: {base}\n")

    print("[1/8] Leyendo período activo y registro de fuentes...")
    registry = SourceRegistry(base)
    print("[2/8] Reconstruyendo motor físico con fuentes configuradas...")
    registry, period, engine, resolver, service, source_stats = build_runtime(base)
    print(f"[3/8] Generalizando capa académica para {period.label}...")
    print("[4/8] Construyendo defaults simples y filtros avanzados...")

    ri_simple = service.resolver_poblacion_simple("RRII")
    ri36 = service.resolver_poblacion_simple("RRII", camadas=[36])
    diseno37_opt = service.resolver_poblacion_simple("Diseño", camadas=[37], incluir_optativas=True)
    examples = [flatten_simple_population(x) for x in [ri_simple, ri36, diseno37_opt]]

    print("[5/8] Probando búsqueda combinada preparada para la web...")
    combined = service.buscar_mejor_horario(
        [{"Carrera": "RRII"}, {"Carrera": "CP", "Camadas": [36]}],
        ["19/08/2026", "20/08/2026", "21/08/2026"],
        capacidad_minima=50, top_n=20,
    )

    print("[6/8] Ejecutando batería de pruebas y regresión Ronda 3...")
    tests = run_tests(registry, period, engine, resolver, service)

    print("[7/8] Exportando configuración, tablas y contrato de interfaz...")
    source_rows = registry.source_status_rows(period.id)
    options_rows = []
    for r in resolver.coverage_operational:
        options_rows.append({
            "Periodo": r["Periodo"], "Carrera": r["Carrera"], "Camada": r["Camada"], "Sede": r["Sede"], "Variante": r["Variante"],
            "Estado_operativo": r["Estado_operativo"], "Códigos_obligatorios": r["Códigos_obligatorios"], "Códigos_optativos": r["Códigos_optativos"],
        })
    review_rows = [r for r in resolver.operational_rows if r.get("Estado_periodo", "").startswith("REVISAR") or r.get("Estado_periodo") == "REQUISITO_GENERICO"]

    write_csv(output / f"plan_operativo_{period.id.lower()}.csv", resolver.operational_rows)
    write_csv(output / f"cobertura_operativa_{period.id.lower()}.csv", resolver.coverage_operational)
    write_csv(output / "estado_fuentes.csv", source_rows)
    write_csv(output / "opciones_interfaz.csv", options_rows)
    write_csv(output / "revisiones_academicas.csv", review_rows)
    write_csv(output / "ejemplos_busqueda_simple.csv", examples)
    write_csv(output / "ranking_simple_ejemplo.csv", combined.get("Ranking", []))
    write_csv(output / "pruebas_automaticas_ronda45.csv", tests)

    contract = {
        "defaults": {
            "periodo": "ACTUAL",
            "ruta_academica": "AUTO_SEGURO",
            "sede_aulas": service.default_sede,
            "camadas": "TODAS_OPERATIVAS",
            "incluir_optativas": service.default_include_optatives,
        },
        "busqueda_simple": ["Carrera"],
        "mas_filtros": ["Camadas", "Sede", "Variante", "Período", "Incluir optativas identificadas", "Códigos extra", "Slots/horario"],
        "funciones": ["listar_periodos", "listar_carreras", "listar_camadas", "listar_sedes", "listar_variantes", "resolver_poblacion_simple", "disponibilidad_estudiantil", "buscar_mejor_horario"],
        "reglas": [
            "Todas las camadas = todas las camadas operativas del período para los filtros seleccionados.",
            "Optativas apagadas por defecto.",
            "Optativas genéricas nunca se inventan.",
            "Variantes paralelas nunca se unen silenciosamente.",
            "Un código curricular no ofertado en el período no se considera inválido.",
        ],
    }
    (output / "contrato_web.json").write_text(json.dumps(contract, ensure_ascii=False, indent=2), encoding="utf-8")

    print("[8/8] Generando diagnóstico final...")
    report = make_report(registry, period, resolver, service, source_stats, tests, examples, combined)
    (output / "diagnostico_ronda45.txt").write_text(report, encoding="utf-8")
    summary = {
        "app": APP_NAME,
        "autor": AUTHOR,
        "generado": dt.datetime.now().isoformat(timespec="seconds"),
        "periodo_activo": period.id,
        "periodo_label": period.label,
        "vigencia": {"inicio": period.start.isoformat(), "fin": period.end.isoformat()},
        "fuentes": source_rows,
        "plan_master_filas": len(resolver.plan_rows),
        "cobertura_combinaciones": len(resolver.coverage_rows),
        "plan_operativo_filas": len(resolver.operational_rows),
        "estados_periodo": dict(Counter(r.get("Estado_periodo", "") for r in resolver.operational_rows)),
        "cobertura_operativa": dict(Counter(r.get("Estado_operativo", "") for r in resolver.coverage_operational)),
        "tests": dict(Counter(t.get("Resultado", "") for t in tests)),
        "defaults": contract["defaults"],
    }
    (output / "resumen_ronda45.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")

    print("\n" + report)
    print(f"\nListo. Los archivos de la Ronda 4.5 quedaron en:\n{output}\n")
    try:
        input("Presioná ENTER para cerrar...")
    except EOFError:
        pass


if __name__ == "__main__":
    main()
