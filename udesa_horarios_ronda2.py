#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
UdeSA Horarios — Ronda 1
Autor: Juan Ignacio Gutiérrez Julián

Objetivo:
- Reconstruir y comparar las fuentes actuales.
- Crear un catálogo maestro de aulas/espacios.
- Detectar comentarios, reservas, eventos y problemas de calidad.

No modifica ninguno de los Excel de entrada.
No necesita pandas ni openpyxl: usa únicamente la biblioteca estándar de Python.
"""

from __future__ import annotations

import csv
import datetime as dt
import json
import math
import os
import posixpath
import re
import sys
import unicodedata
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Set, Tuple
from zipfile import ZipFile

APP_NAME = "UdeSA Horarios"
AUTHOR = "Juan Ignacio Gutiérrez Julián"
YEAR = 2026

REQUIRED_FILES = {
    "cursos": "CURSOS PRIMAVERA 26.xlsx",
    "aulas": "AULAS- PRIMAVERA 2026- FINAL.xlsx",
    "eventos": "Registro de actividades 2026 (1).xlsx",
    "carreras": "Area de Charlas, Camada y codigos.xlsx",
    "catalogo_aux": "aulas_udesA.xlsx",
}

OUTPUT_DIR_NAME = "salida_ronda1"

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

WEEKDAYS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]

# Espacios que no son aulas codificadas, pero deben poder formar parte del catálogo.
# CAMPUS se incluye como espacio canónico, pero la palabra "Campus" sola NO se resuelve
# automáticamente desde Registro de Eventos porque también puede referirse a la sede.
MANUAL_SPACES = {
    "CUBO": {
        "nombre": "Cubo",
        "tipo": "Espacio común",
        "sede": "Campus",
        "aliases": ["Cubo"],
    },
    "COMEDOR": {
        "nombre": "Comedor",
        "tipo": "Espacio común",
        "sede": "Campus",
        "aliases": ["Comedor", "Comerdor"],
    },
    "COMEDOR VIP": {
        "nombre": "Comedor VIP",
        "tipo": "Espacio común",
        "sede": "Campus",
        "aliases": ["Comedor VIP", "VIP Comedor", "VIP del Comedor"],
    },
    "COFFEE QUAD": {
        "nombre": "Coffee Quad",
        "tipo": "Espacio común",
        "sede": "Campus",
        "aliases": ["Coffee Quad", "Coffee QUAD"],
    },
    "GALERIA DISENO": {
        "nombre": "Galería Diseño",
        "tipo": "Espacio común",
        "sede": "Campus",
        "aliases": ["Galería Diseño", "Galeria Diseño", "Galería de Diseño", "Galeria de Diseño"],
    },
    "GALERIA LAGUNA": {
        "nombre": "Galería Laguna",
        "tipo": "Espacio común",
        "sede": "Campus",
        "aliases": ["Galería Laguna", "Galeria Laguna", "Galería de la Laguna", "Galeria de la Laguna", "Galería Laguna sector Foodtruck", "Galeria Laguna sector Foodtruck", "Sector Foodtruck"],
    },
    "SALA RECTORADO": {
        "nombre": "Sala Rectorado",
        "tipo": "Sala",
        "sede": "Campus",
        "aliases": ["Sala Rectorado", "Sala de Rectorado"],
    },
    "SALA 4 BIBLIOTECA": {
        "nombre": "Sala 4 de Biblioteca",
        "tipo": "Sala",
        "sede": "Campus",
        "aliases": ["Sala 4 de Biblioteca", "Sala 4 Biblioteca", "Biblioteca Sala 4"],
    },
    "HALL SULLAIR": {
        "nombre": "Hall Sullair",
        "tipo": "Hall",
        "sede": "Campus",
        "aliases": ["Hall Sullair", "Hall del Sullair"],
    },
    "CAMPUS": {
        "nombre": "Campus",
        "tipo": "Espacio común",
        "sede": "Campus",
        "aliases": ["Espacio Campus"],
    },
}

# Alias manuales que sabemos que son inequívocos.
MANUAL_ALIASES = {
    "AGARDY": "HAM",
    "AGÁRDY": "HAM",
    "AULA MAGNA": "HAM",
    "AULA MAGNA AGARDY": "HAM",
    "AULA MAGNA DR JENO AGARDY": "HAM",
}

# Prefijos de aulas que normalmente usan tres dígitos.
PAD3_PREFIXES = {"B", "C", "G", "H", "M", "R", "RS", "S", "SS", "V"}


def strip_accents(text: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", text) if unicodedata.category(c) != "Mn")


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    text = str(value).replace("\u00a0", " ").strip()
    text = re.sub(r"\s+", " ", text)
    return text


def norm_text(value: Any) -> str:
    text = strip_accents(clean_text(value)).upper()
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def norm_header(value: Any) -> str:
    text = norm_text(value)
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def normalize_day(value: Any) -> str:
    n = norm_text(value)
    mapping = {
        "LUNES": "Lunes",
        "MARTES": "Martes",
        "MIERCOLES": "Miércoles",
        "JUEVES": "Jueves",
        "VIERNES": "Viernes",
        "SABADO": "Sábado",
        "DOMINGO": "Domingo",
    }
    return mapping.get(n, clean_text(value))


def normalize_sede(value: Any) -> str:
    n = norm_text(value)
    if not n or n in {"N/A", "NA", "NONE"}:
        return ""
    if n in {"VICTORIA", "CAMPUS", "CAMPUS VICTORIA", "SEDE VICTORIA"}:
        return "Campus"
    if "RIOBAMBA" in n:
        return "Riobamba"
    if "CALLAO" in n:
        return "Callao"
    if "SUIPACHA" in n:
        return "Suipacha"
    if "DIGITAL HOUSE" in n:
        return "Digital House Belgrano"
    if "AREA BETA" in n:
        return "Area Beta"
    return clean_text(value)


def normalize_floor(value: Any) -> str:
    if value is None or clean_text(value) in {"", "N/A"}:
        return ""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        num = int(value)
        mapping = {
            -2: "Segundo Subsuelo",
            -1: "Primer Subsuelo",
            0: "Planta Baja",
            1: "Primer Piso",
            2: "Segundo Piso",
            3: "Tercer Piso",
            4: "Cuarto Piso",
            5: "Quinto Piso",
        }
        return mapping.get(num, str(num))
    t = clean_text(value)
    if norm_text(t) in {"BAJA", "PB", "PLANTA BAJA"}:
        return "Planta Baja"
    return t


def normalize_room_key(value: Any) -> str:
    """Normaliza códigos sin confundir nombres de espacios especiales."""
    raw = clean_text(value)
    if not raw:
        return ""
    n = norm_text(raw)
    n = re.sub(r"^(AULA|SALON)\s+", "", n).strip()
    n = n.replace(".", "")
    n = re.sub(r"\s*[-–—]\s*", "-", n)

    # Nombres como Sala 1 son identificadores reales en Callao y no deben convertirse en SALA1.
    m = re.fullmatch(r"SALA\s+(\d+)", n)
    if m:
        return f"SALA {int(m.group(1))}"

    # Código estándar: G 006 -> G006, S 8 -> S008, RS 2 -> RS002.
    m = re.fullmatch(r"([A-Z]{1,4})\s*[-]?\s*(\d{1,3})", n)
    if m:
        prefix, digits = m.groups()
        if prefix in PAD3_PREFIXES:
            return f"{prefix}{int(digits):03d}"
        return f"{prefix}{int(digits)}"

    # "DH 1" se conserva como DH1 para comparación.
    m = re.fullmatch(r"DH\s*(\d+)", n)
    if m:
        return f"DH{int(m.group(1))}"

    return n


def display_room_from_key(key: str, preferred: str = "") -> str:
    if preferred:
        return clean_text(preferred)
    if key == "GALERIA DISENO":
        return "Galería Diseño"
    if key.startswith("SALA "):
        return key.title()
    return key


def canonical_career(value: Any) -> str:
    n = norm_text(value).replace(".", "")
    mapping = {
        "CP": "CP",
        "CIENCIA POLITICA": "CP",
        "RRII": "RRII",
        "RIII": "RRII",  # variante/typo observada en la fuente
        "RELACIONES INTERNACIONALES": "RRII",
        "ABOGACIA": "Abogacía",
        "ADMIN": "Administración",
        "ADMINISTRACION": "Administración",
        "ECO": "Economía",
        "ECONOMIA": "Economía",
    }
    return mapping.get(n, clean_text(value).strip())


def parse_numeric_capacity(value: Any) -> Optional[int]:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        try:
            if math.isnan(float(value)):
                return None
        except Exception:
            pass
        return int(value)
    t = clean_text(value)
    if not t or norm_text(t) in {"N/A", "NA"}:
        return None
    m = re.fullmatch(r"\d+(?:\.0+)?", t)
    return int(float(t)) if m else None


def excel_datetime(serial: float) -> dt.datetime:
    return dt.datetime(1899, 12, 30) + dt.timedelta(days=serial)


def format_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, dt.datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(value, dt.date):
        return value.isoformat()
    if isinstance(value, dt.time):
        return value.strftime("%H:%M:%S")
    return clean_text(value)


def parse_time_token(value: Any) -> Optional[dt.time]:
    if isinstance(value, dt.datetime):
        return value.time().replace(microsecond=0)
    if isinstance(value, dt.time):
        return value.replace(microsecond=0)
    if isinstance(value, (int, float)) and 0 <= float(value) < 1:
        seconds = round(float(value) * 86400) % 86400
        return dt.time(seconds // 3600, (seconds % 3600) // 60, seconds % 60)
    text = clean_text(value)
    if not text:
        return None
    # Corrige O mayúscula/minúscula usada como cero dentro de horarios (14:OO).
    text = re.sub(r"(?<=\d)[:.]([Oo]{1,2})\b", lambda m: ":" + "0" * len(m.group(1)), text)
    text = text.replace(".", ":")
    m = re.search(r"(?<!\d)(\d{1,2})(?::(\d{1,2}))?(?::(\d{1,2}))?(?!\d)", text)
    if not m:
        return None
    h = int(m.group(1))
    mi = int(m.group(2) or 0)
    sec = int(m.group(3) or 0)
    if h > 23 or mi > 59 or sec > 59:
        return None
    return dt.time(h, mi, sec)


def parse_time_range(value: Any) -> Tuple[Optional[dt.time], Optional[dt.time], str]:
    """Devuelve inicio, fin, estado: OK | INCOMPLETO | INVALIDO | VACIO."""
    if value is None or clean_text(value) == "":
        return None, None, "VACIO"
    if isinstance(value, (dt.time, dt.datetime)) or (isinstance(value, (int, float)) and 0 <= float(value) < 1):
        start = parse_time_token(value)
        return start, None, "INCOMPLETO" if start else "INVALIDO"

    text = clean_text(value)
    text_fixed = re.sub(r"(?<=\d)[:.]([Oo]{1,2})\b", lambda m: ":" + "0" * len(m.group(1)), text)
    text_fixed = text_fixed.replace(".", ":")
    tokens = re.findall(r"(?<!\d)(\d{1,2}(?::\d{1,2})?(?::\d{1,2})?)(?!\d)", text_fixed)
    times = []
    for tok in tokens:
        t = parse_time_token(tok)
        if t is not None:
            times.append(t)
    if len(times) >= 2:
        return times[0], times[1], "OK"
    if len(times) == 1:
        return times[0], None, "INCOMPLETO"
    return None, None, "INVALIDO"


def normalize_slot(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if float(value).is_integer():
            return str(int(value))
        return str(value)
    text = clean_text(value)
    if not text:
        return ""
    m = re.match(r"^\s*(\d+)", text)
    return m.group(1) if m else text


class XLSXReader:
    """Lector liviano de XLSX basado en XML/ZIP (solo lectura)."""

    def __init__(self, path: Path):
        self.path = Path(path)
        self.z = ZipFile(self.path)
        self.shared_strings = self._load_shared_strings()
        self.date_styles = self._load_date_styles()
        self.sheets = self._load_sheets()

    def close(self) -> None:
        self.z.close()

    def __enter__(self) -> "XLSXReader":
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()

    @staticmethod
    def _resolve(base: str, target: str) -> str:
        if target.startswith("/"):
            return target.lstrip("/")
        if target.startswith("xl/"):
            return target
        return posixpath.normpath(posixpath.join(posixpath.dirname(base), target))

    def _load_shared_strings(self) -> List[str]:
        name = "xl/sharedStrings.xml"
        if name not in self.z.namelist():
            return []
        root = ET.fromstring(self.z.read(name))
        result = []
        for si in root.findall(f"{{{MAIN_NS}}}si"):
            result.append("".join(t.text or "" for t in si.iter(f"{{{MAIN_NS}}}t")))
        return result

    def _load_date_styles(self) -> Set[int]:
        name = "xl/styles.xml"
        if name not in self.z.namelist():
            return set()
        root = ET.fromstring(self.z.read(name))
        custom_formats: Dict[int, str] = {}
        num_fmts = root.find(f"{{{MAIN_NS}}}numFmts")
        if num_fmts is not None:
            for nf in num_fmts:
                try:
                    custom_formats[int(nf.attrib["numFmtId"])] = nf.attrib.get("formatCode", "")
                except Exception:
                    pass

        builtin_date = {14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 30, 36, 45, 46, 47, 50, 57}
        date_styles: Set[int] = set()
        xfs = root.find(f"{{{MAIN_NS}}}cellXfs")
        if xfs is None:
            return date_styles
        for idx, xf in enumerate(xfs):
            try:
                fmt_id = int(xf.attrib.get("numFmtId", "0"))
            except Exception:
                fmt_id = 0
            fmt = custom_formats.get(fmt_id, "").lower()
            clean_fmt = re.sub(r'"[^"]*"|\\.|\[[^\]]*\]', "", fmt)
            if fmt_id in builtin_date or any(token in clean_fmt for token in ("yy", "dd", "hh", "ss")):
                date_styles.add(idx)
        return date_styles

    def _load_sheets(self) -> List[Dict[str, str]]:
        wb = ET.fromstring(self.z.read("xl/workbook.xml"))
        rel_root = ET.fromstring(self.z.read("xl/_rels/workbook.xml.rels"))
        relmap = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rel_root}
        sheets: List[Dict[str, str]] = []
        sheet_node = wb.find(f"{{{MAIN_NS}}}sheets")
        if sheet_node is None:
            return sheets
        for sheet in sheet_node:
            rid = sheet.attrib.get(f"{{{OFFICE_REL_NS}}}id")
            if not rid or rid not in relmap:
                continue
            sheets.append({
                "name": sheet.attrib.get("name", ""),
                "path": self._resolve("xl/workbook.xml", relmap[rid]),
                "state": sheet.attrib.get("state", "visible"),
            })
        return sheets

    @staticmethod
    def _col_letters(ref: str) -> str:
        m = re.match(r"([A-Z]+)", ref)
        return m.group(1) if m else ""

    @staticmethod
    def _row_number(ref: str) -> int:
        m = re.search(r"(\d+)$", ref)
        return int(m.group(1)) if m else 0

    def _cell_value(self, cell: ET.Element) -> Any:
        cell_type = cell.attrib.get("t", "")
        style_idx = int(cell.attrib.get("s", "0") or 0)

        if cell_type == "inlineStr":
            return "".join(t.text or "" for t in cell.iter(f"{{{MAIN_NS}}}t"))

        value_node = cell.find(f"{{{MAIN_NS}}}v")
        if value_node is None:
            return ""
        raw = value_node.text or ""

        if cell_type == "s":
            try:
                return self.shared_strings[int(raw)]
            except Exception:
                return raw
        if cell_type == "b":
            return raw == "1"
        if cell_type in {"str", "e"}:
            return raw

        try:
            num = float(raw)
            if style_idx in self.date_styles:
                converted = excel_datetime(num)
                if 0 <= num < 1:
                    return converted.time().replace(microsecond=0)
                return converted
            if num.is_integer():
                return int(num)
            return num
        except Exception:
            return raw

    def sheet_names(self, include_hidden: bool = True) -> List[str]:
        if include_hidden:
            return [s["name"] for s in self.sheets]
        return [s["name"] for s in self.sheets if s["state"] == "visible"]

    def sheet_info(self, sheet_name: str) -> Dict[str, str]:
        for s in self.sheets:
            if s["name"] == sheet_name:
                return s
        raise KeyError(f"No existe la hoja: {sheet_name}")

    def rows(self, sheet_name: str) -> List[Tuple[int, Dict[str, Any]]]:
        info = self.sheet_info(sheet_name)
        root = ET.fromstring(self.z.read(info["path"]))
        out: List[Tuple[int, Dict[str, Any]]] = []
        for row in root.findall(f".//{{{MAIN_NS}}}sheetData/{{{MAIN_NS}}}row"):
            row_num = int(row.attrib.get("r", "0") or 0)
            data: Dict[str, Any] = {}
            for cell in row.findall(f"{{{MAIN_NS}}}c"):
                ref = cell.attrib.get("r", "")
                col = self._col_letters(ref)
                if col:
                    data[col] = self._cell_value(cell)
            out.append((row_num, data))
        return out

    def find_header_row(self, sheet_name: str, required_headers: Sequence[str], max_rows: int = 30) -> int:
        required = {norm_header(h) for h in required_headers}
        for row_num, row in self.rows(sheet_name)[:max_rows]:
            present = {norm_header(v) for v in row.values() if clean_text(v)}
            if required.issubset(present):
                return row_num
        return 0

    def dict_rows(self, sheet_name: str, header_row: int) -> List[Tuple[int, Dict[str, Any]]]:
        rows = self.rows(sheet_name)
        header: Dict[str, str] = {}
        for row_num, row in rows:
            if row_num == header_row:
                for col, value in row.items():
                    header[col] = clean_text(value)
                break
        if not header:
            return []

        result = []
        for row_num, row in rows:
            if row_num <= header_row:
                continue
            item: Dict[str, Any] = {}
            for col, value in row.items():
                key = header.get(col, col)
                if key:
                    item[key] = value
            if any(clean_text(v) for v in item.values()):
                result.append((row_num, item))
        return result

    def comments(self, sheet_name: str) -> List[Dict[str, Any]]:
        info = self.sheet_info(sheet_name)
        sheet_path = info["path"]
        rel_path = posixpath.join(posixpath.dirname(sheet_path), "_rels", posixpath.basename(sheet_path) + ".rels")
        if rel_path not in self.z.namelist():
            return []
        rel_root = ET.fromstring(self.z.read(rel_path))
        comments_path = ""
        for rel in rel_root:
            if rel.attrib.get("Type", "").endswith("/comments"):
                comments_path = self._resolve(sheet_path, rel.attrib.get("Target", ""))
                break
        if not comments_path or comments_path not in self.z.namelist():
            return []

        root = ET.fromstring(self.z.read(comments_path))
        authors_node = root.find(f"{{{MAIN_NS}}}authors")
        authors = []
        if authors_node is not None:
            authors = [a.text or "" for a in authors_node.findall(f"{{{MAIN_NS}}}author")]

        result = []
        comment_list = root.find(f"{{{MAIN_NS}}}commentList")
        if comment_list is None:
            return result
        for comment in comment_list.findall(f"{{{MAIN_NS}}}comment"):
            ref = comment.attrib.get("ref", "")
            author_id = int(comment.attrib.get("authorId", "0") or 0)
            text_node = comment.find(f"{{{MAIN_NS}}}text")
            text = ""
            if text_node is not None:
                text = "".join(t.text or "" for t in text_node.iter(f"{{{MAIN_NS}}}t"))
            result.append({
                "celda": ref,
                "fila": self._row_number(ref),
                "columna": self._col_letters(ref),
                "autor": authors[author_id] if 0 <= author_id < len(authors) else "",
                "comentario": text.strip(),
            })
        return result


@dataclass
class CatalogEntry:
    key: str
    nombre: str
    tipo: str = "Aula"
    capacidad: Optional[int] = None
    sede: str = ""
    edificio: str = ""
    piso: str = ""
    aliases: Set[str] = None
    fuente_capacidad: str = ""
    fuente_sede: str = ""
    fuente_edificio: str = ""
    fuente_piso: str = ""
    en_aulas_actual: bool = False
    en_cursos: bool = False
    en_catalogo_auxiliar: bool = False
    observaciones: List[str] = None

    def __post_init__(self):
        if self.aliases is None:
            self.aliases = set()
        if self.observaciones is None:
            self.observaciones = []


class MasterCatalog:
    def __init__(self):
        self.entries: Dict[str, CatalogEntry] = {}
        self.capacity_conflicts: List[Dict[str, Any]] = []
        self.alias_conflicts: List[Dict[str, Any]] = []

    def get_or_create(self, key: str, preferred_name: str = "", tipo: str = "Aula") -> CatalogEntry:
        if key not in self.entries:
            self.entries[key] = CatalogEntry(key=key, nombre=display_room_from_key(key, preferred_name), tipo=tipo)
        elif preferred_name and self.entries[key].nombre == self.entries[key].key:
            self.entries[key].nombre = clean_text(preferred_name)
        return self.entries[key]

    def set_capacity(self, key: str, capacity: Optional[int], source: str, priority: int, preferred_name: str = "") -> None:
        if capacity is None or not key:
            return
        e = self.get_or_create(key, preferred_name)
        current_priority = getattr(e, "_capacity_priority", -1)
        if e.capacidad is not None and e.capacidad != capacity:
            self.capacity_conflicts.append({
                "Espacio": e.nombre,
                "Espacio_ID": key,
                "Capacidad_elegida_antes": e.capacidad,
                "Fuente_anterior": e.fuente_capacidad,
                "Capacidad_nueva": capacity,
                "Fuente_nueva": source,
                "Se_elige_nueva": priority > current_priority,
            })
        if e.capacidad is None or priority > current_priority:
            e.capacidad = capacity
            e.fuente_capacidad = source
            setattr(e, "_capacity_priority", priority)

    def set_attr(self, key: str, attr: str, value: str, source: str, priority: int, preferred_name: str = "") -> None:
        if not key or not clean_text(value):
            return
        e = self.get_or_create(key, preferred_name)
        priority_attr = f"_{attr}_priority"
        current_priority = getattr(e, priority_attr, -1)
        if not getattr(e, attr) or priority > current_priority:
            setattr(e, attr, clean_text(value))
            setattr(e, f"fuente_{attr}", source)
            setattr(e, priority_attr, priority)

    def add_alias(self, key: str, alias: str) -> None:
        alias = clean_text(alias)
        if not key or not alias:
            return
        if norm_text(alias) in {"N/A", "NA", "AULA", "SALA"}:
            return
        e = self.get_or_create(key)
        e.aliases.add(alias)

    def alias_index(self) -> Dict[str, str]:
        candidates: Dict[str, Set[str]] = defaultdict(set)
        for key, entry in self.entries.items():
            candidates[norm_text(entry.nombre)].add(key)
            candidates[norm_text(key)].add(key)
            for alias in entry.aliases:
                candidates[norm_text(alias)].add(key)
        for alias, key in MANUAL_ALIASES.items():
            if key in self.entries:
                candidates[norm_text(alias)].add(key)

        index: Dict[str, str] = {}
        self.alias_conflicts = []
        for alias_norm, keys in candidates.items():
            if not alias_norm:
                continue
            if len(keys) == 1:
                index[alias_norm] = next(iter(keys))
            else:
                self.alias_conflicts.append({"Alias": alias_norm, "Espacios": " | ".join(sorted(keys))})
        return index

    def rows(self) -> List[Dict[str, Any]]:
        result = []
        for key, e in sorted(self.entries.items(), key=lambda kv: (kv[1].sede or "ZZZ", kv[1].nombre)):
            result.append({
                "Espacio_ID": key,
                "Nombre": e.nombre,
                "Tipo": e.tipo,
                "Capacidad": e.capacidad if e.capacidad is not None else "",
                "Sede": e.sede,
                "Edificio": e.edificio,
                "Piso": e.piso,
                "Aliases": " | ".join(sorted(e.aliases, key=lambda x: norm_text(x))),
                "Fuente_Capacidad": e.fuente_capacidad,
                "Fuente_Sede": e.fuente_sede,
                "Fuente_Edificio": e.fuente_edificio,
                "Fuente_Piso": e.fuente_piso,
                "En_AULAS_Actual": "Sí" if e.en_aulas_actual else "No",
                "En_CURSOS": "Sí" if e.en_cursos else "No",
                "En_Catalogo_Auxiliar": "Sí" if e.en_catalogo_auxiliar else "No",
                "Observaciones": " | ".join(e.observaciones),
            })
        return result


def header_map(row: Dict[str, Any]) -> Dict[str, Any]:
    return {norm_header(k): v for k, v in row.items()}


def val(row: Dict[str, Any], *headers: str) -> Any:
    hm = header_map(row)
    for h in headers:
        key = norm_header(h)
        if key in hm:
            return hm[key]
    return ""


def safe_date(value: Any) -> Optional[dt.date]:
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    text = clean_text(value)
    if not text:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y"):
        try:
            return dt.datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def write_csv(path: Path, rows: List[Dict[str, Any]], fieldnames: Optional[List[str]] = None) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if not rows and not fieldnames:
        fieldnames = ["Sin datos"]
    if fieldnames is None:
        fieldnames = list(rows[0].keys()) if rows else ["Sin datos"]
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow({k: format_value(row.get(k, "")) for k in fieldnames})


def find_required_files(base: Path) -> Dict[str, Path]:
    found = {}
    missing = []
    for key, filename in REQUIRED_FILES.items():
        path = base / filename
        if path.exists():
            found[key] = path
        else:
            missing.append(filename)
    if missing:
        print("\nERROR: faltan archivos necesarios en la misma carpeta que este programa:\n")
        for filename in missing:
            print(f"  - {filename}")
        print("\nCopialos a esta carpeta y volvé a ejecutar el programa.")
        raise SystemExit(1)
    return found


def process_aux_catalog(path: Path, catalog: MasterCatalog, problems: List[Dict[str, Any]]) -> Dict[str, Any]:
    with XLSXReader(path) as x:
        sheet = x.sheet_names(include_hidden=False)[0]
        header = x.find_header_row(sheet, ["Código", "Capacidad", "Sede"])
        rows = x.dict_rows(sheet, header)

    # Primero recolectamos los nombres viejos para no convertir alias ambiguos (p.ej. "Sala") en equivalencias falsas.
    alias_to_keys: Dict[str, Set[str]] = defaultdict(set)
    parsed_rows = []
    for row_num, row in rows:
        code_raw = val(row, "Código")
        key = normalize_room_key(code_raw)
        if not key:
            continue
        old_name = clean_text(val(row, "Nombre Viejo"))
        parsed_rows.append((row_num, row, key, old_name))
        if old_name and norm_text(old_name) not in {"N/A", "NA", "AULA", "SALA"}:
            alias_to_keys[norm_text(old_name)].add(key)

    # Genera alias abreviados seguros a partir de nombres propios únicos.
    # Ej.: "Juan José Vergez" -> "Vergez" y "Aula Vergez".
    short_alias_candidates: Dict[str, Set[str]] = defaultdict(set)
    for _, _, key, old_name in parsed_rows:
        if not old_name or norm_text(old_name) in {"N/A", "NA", "AULA", "SALA"}:
            continue
        words = re.findall(r"[A-Za-zÁÉÍÓÚÜÑáéíóúüñ]+", old_name)
        if words:
            last = words[-1]
            last_norm = norm_text(last)
            if len(last_norm) >= 4 and last_norm not in {"AULA", "SALA", "LABORATORIO", "FAMILIA", "EDIFICIO"}:
                short_alias_candidates[last_norm].add(key)

    for row_num, row, key, old_name in parsed_rows:
        preferred = clean_text(val(row, "Código")) or key
        e = catalog.get_or_create(key, preferred_name=preferred)
        e.en_catalogo_auxiliar = True
        catalog.set_capacity(key, parse_numeric_capacity(val(row, "Capacidad")), "Catálogo auxiliar", priority=50, preferred_name=preferred)
        catalog.set_attr(key, "sede", normalize_sede(val(row, "Sede")), "Catálogo auxiliar", priority=30, preferred_name=preferred)
        building = clean_text(val(row, "Edificio"))
        if norm_text(building) not in {"", "N/A", "NA"}:
            catalog.set_attr(key, "edificio", building, "Catálogo auxiliar", priority=30, preferred_name=preferred)
        floor = normalize_floor(val(row, "Planta"))
        catalog.set_attr(key, "piso", floor, "Catálogo auxiliar", priority=30, preferred_name=preferred)
        if old_name and alias_to_keys.get(norm_text(old_name)) == {key}:
            catalog.add_alias(key, old_name)
            words = re.findall(r"[A-Za-zÁÉÍÓÚÜÑáéíóúüñ]+", old_name)
            if words:
                last = words[-1]
                if short_alias_candidates.get(norm_text(last)) == {key}:
                    catalog.add_alias(key, last)
                    catalog.add_alias(key, f"Aula {last}")

    return {"filas": len(parsed_rows)}


def process_courses(path: Path, catalog: MasterCatalog, problems: List[Dict[str, Any]]) -> Dict[str, Any]:
    with XLSXReader(path) as x:
        sheet = x.sheet_names(include_hidden=False)[0]
        header = x.find_header_row(sheet, ["Materia", "Dia", "Espacio", "Cod Materia", "Cant Inscriptos"])
        rows = x.dict_rows(sheet, header)

    course_codes: Set[str] = set()
    valid_assignments = 0
    missing_code = 0
    missing_room = 0
    missing_sede = 0
    invalid_times = 0
    course_rows_out = []

    for row_num, row in rows:
        code = norm_text(val(row, "Cod Materia"))
        course_name = clean_text(val(row, "Materia"))
        room_raw = val(row, "Espacio")
        room = normalize_room_key(room_raw)
        sede = normalize_sede(val(row, "Sede"))
        building = clean_text(val(row, "Edificio"))
        floor = normalize_floor(val(row, "Piso"))
        day = normalize_day(val(row, "Dia"))
        slot = normalize_slot(val(row, "N Slot", "Slot"))
        start = parse_time_token(val(row, "Hora Desde"))
        end = parse_time_token(val(row, "Hora Hasta"))
        students = val(row, "Cant Inscriptos")

        if code:
            course_codes.add(code)
        else:
            missing_code += 1
            problems.append({"Tipo": "Curso sin código", "Fuente": path.name, "Hoja": sheet, "Fila": row_num, "Valor": course_name, "Detalle": "Cod Materia vacío"})

        if not room:
            missing_room += 1
            problems.append({"Tipo": "Curso sin aula", "Fuente": path.name, "Hoja": sheet, "Fila": row_num, "Valor": course_name or code, "Detalle": "Espacio vacío"})
        if not sede:
            missing_sede += 1
            problems.append({"Tipo": "Curso sin sede", "Fuente": path.name, "Hoja": sheet, "Fila": row_num, "Valor": course_name or code, "Detalle": "Sede vacía"})
        if (clean_text(val(row, "Hora Desde")) and not start) or (clean_text(val(row, "Hora Hasta")) and not end):
            invalid_times += 1
            problems.append({"Tipo": "Horario inválido", "Fuente": path.name, "Hoja": sheet, "Fila": row_num, "Valor": f"{format_value(val(row, 'Hora Desde'))} - {format_value(val(row, 'Hora Hasta'))}", "Detalle": course_name or code})

        if room:
            e = catalog.get_or_create(room, preferred_name=clean_text(room_raw) or room)
            e.en_cursos = True
            catalog.set_attr(room, "sede", sede, "CURSOS Primavera 26", priority=20, preferred_name=clean_text(room_raw))
            catalog.set_attr(room, "edificio", building, "CURSOS Primavera 26", priority=20, preferred_name=clean_text(room_raw))
            catalog.set_attr(room, "piso", floor, "CURSOS Primavera 26", priority=20, preferred_name=clean_text(room_raw))

        if code and day and room and start and end:
            valid_assignments += 1

        course_rows_out.append({
            "Fila": row_num,
            "Código": code,
            "Materia": course_name,
            "Día": day,
            "Hora Desde": start.strftime("%H:%M") if start else "",
            "Hora Hasta": end.strftime("%H:%M") if end else "",
            "Slot": slot,
            "Aula_original": clean_text(room_raw),
            "Espacio_ID": room,
            "Edificio": building,
            "Piso": floor,
            "Sede": sede,
            "Cant Inscriptos": students,
        })

    return {
        "filas": len(rows),
        "cursos_unicos": len(course_codes),
        "asignaciones_validas": valid_assignments,
        "sin_codigo": missing_code,
        "sin_aula": missing_room,
        "sin_sede": missing_sede,
        "horarios_invalidos": invalid_times,
        "rows_out": course_rows_out,
    }


def process_aulas(path: Path, catalog: MasterCatalog, problems: List[Dict[str, Any]]) -> Dict[str, Any]:
    with XLSXReader(path) as x:
        visible = x.sheet_names(include_hidden=False)
        grid_sheet = next((s for s in visible if norm_text(s).startswith("AULAS PRIMAVERA")), visible[0])
        header = x.find_header_row(grid_sheet, ["Aula", "Capacidad", "Slot", "Lunes"])
        grid_rows = x.dict_rows(grid_sheet, header)

        grid_row_lookup: Dict[int, Dict[str, Any]] = {rn: row for rn, row in grid_rows}
        raw_rows = {rn: row for rn, row in x.rows(grid_sheet)}

        occupancy_cells = 0
        rooms: Set[str] = set()
        capacity_values: Dict[str, Set[int]] = defaultdict(set)

        for row_num, row in grid_rows:
            room_raw = val(row, "Aula")
            room = normalize_room_key(room_raw)
            if not room:
                continue
            rooms.add(room)
            e = catalog.get_or_create(room, preferred_name=clean_text(room_raw) or room)
            e.en_aulas_actual = True
            # Esta grilla es del Campus Victoria.
            catalog.set_attr(room, "sede", "Campus", "AULAS Primavera 26", priority=25, preferred_name=clean_text(room_raw))
            cap = parse_numeric_capacity(val(row, "Capacidad"))
            if cap is not None:
                capacity_values[room].add(cap)
                # Prioridad máxima para capacidad vigente.
                catalog.set_capacity(room, cap, "AULAS Primavera 26", priority=40, preferred_name=clean_text(room_raw))
            else:
                problems.append({"Tipo": "Aula sin capacidad en grilla", "Fuente": path.name, "Hoja": grid_sheet, "Fila": row_num, "Valor": clean_text(room_raw), "Detalle": "Capacidad vacía/no numérica"})

            for day in ["Lunes", "Martes", "Miercoles", "Jueves", "Viernes"]:
                cell = clean_text(val(row, day))
                if cell:
                    occupancy_cells += 1
                    # Alias observable en HAM - AGARDY.
                    if room == "HAM" and "AGARDY" in norm_text(cell):
                        catalog.add_alias("HAM", "Agardy")

        for room, caps in capacity_values.items():
            if len(caps) > 1:
                problems.append({"Tipo": "Capacidad inconsistente dentro de AULAS", "Fuente": path.name, "Hoja": grid_sheet, "Fila": "", "Valor": room, "Detalle": " | ".join(map(str, sorted(caps)))})

        # Comentarios con contexto de aula, slot y día.
        comment_rows = []
        # Mapeo de columna a encabezado a partir de la fila de header.
        header_raw = raw_rows.get(header, {})
        col_to_header = {col: clean_text(value) for col, value in header_raw.items()}
        for c in x.comments(grid_sheet):
            context = grid_row_lookup.get(c["fila"], {})
            room_raw = val(context, "Aula")
            room = normalize_room_key(room_raw)
            slot = normalize_slot(val(context, "Slot"))
            day = normalize_day(col_to_header.get(c["columna"], ""))
            comment_rows.append({
                "Hoja": grid_sheet,
                "Celda": c["celda"],
                "Aula_original": clean_text(room_raw),
                "Espacio_ID": room,
                "Slot": slot,
                "Día": day,
                "Autor": c["autor"],
                "Comentario": c["comentario"],
            })

        # Reservas especiales: cualquier hoja visible que tenga FECHA + AULA.
        reservations = []
        reservation_sheets = []
        for sheet in visible:
            if sheet == grid_sheet:
                continue
            h = x.find_header_row(sheet, ["FECHA", "AULA"], max_rows=20)
            if not h:
                continue
            reservation_sheets.append(sheet)
            for row_num, row in x.dict_rows(sheet, h):
                date = safe_date(val(row, "FECHA"))
                room_raw = val(row, "AULA")
                room = normalize_room_key(room_raw)
                if not date and not room:
                    continue
                start = parse_time_token(val(row, "HORA DESDE"))
                end = parse_time_token(val(row, "HORA HASTA"))
                if room:
                    e = catalog.get_or_create(room, preferred_name=clean_text(room_raw) or room)
                    # Las reservas de esta planilla son del campus salvo evidencia contraria.
                    catalog.set_attr(room, "sede", "Campus", f"{path.name} / {sheet}", priority=15, preferred_name=clean_text(room_raw))
                if clean_text(val(row, "HORA DESDE")) and not start:
                    problems.append({"Tipo": "Horario inválido", "Fuente": path.name, "Hoja": sheet, "Fila": row_num, "Valor": format_value(val(row, "HORA DESDE")), "Detalle": "Reserva especial - hora desde"})
                if clean_text(val(row, "HORA HASTA")) and not end:
                    problems.append({"Tipo": "Horario inválido", "Fuente": path.name, "Hoja": sheet, "Fila": row_num, "Valor": format_value(val(row, "HORA HASTA")), "Detalle": "Reserva especial - hora hasta"})
                reservations.append({
                    "Hoja": sheet,
                    "Fila": row_num,
                    "Profesor/Pedido por": clean_text(val(row, "PROFESOR", "PROFESOR o PEDIDO POR")),
                    "Motivo": clean_text(val(row, "MOTIVO")),
                    "Fecha": date.isoformat() if date else "",
                    "Hora Desde": start.strftime("%H:%M") if start else "",
                    "Hora Hasta": end.strftime("%H:%M") if end else "",
                    "Aula_original": clean_text(room_raw),
                    "Espacio_ID": room,
                    "Antes": clean_text(val(row, "ANTES")),
                    "Observaciones": clean_text(val(row, "OBS")),
                })

    return {
        "grid_sheet": grid_sheet,
        "aulas_grid": len(rooms),
        "ocupaciones_grid": occupancy_cells,
        "comentarios": len(comment_rows),
        "comments_out": comment_rows,
        "reservas": len([r for r in reservations if r["Fecha"] and r["Espacio_ID"]]),
        "reservations_out": reservations,
        "reservation_sheets": reservation_sheets,
    }


def build_location_parser(catalog: MasterCatalog):
    alias_index = catalog.alias_index()

    # Frases de alias con suficiente longitud para evitar falsos positivos.
    named_aliases: List[Tuple[str, str]] = []
    for alias_norm, key in alias_index.items():
        # Los códigos se capturan aparte; acá usamos nombres textuales.
        if re.fullmatch(r"[A-Z]{1,4}\d{1,3}", alias_norm.replace(" ", "")):
            continue
        if alias_norm in {"CAMPUS", "AULA", "SALA"} or re.fullmatch(r"SALA \d+", alias_norm):
            continue
        if len(alias_norm) >= 4:
            named_aliases.append((alias_norm, key))
    named_aliases.sort(key=lambda x: len(x[0]), reverse=True)

    known_keys = set(catalog.entries.keys())

    def parse_location(where: Any) -> Dict[str, Any]:
        raw = clean_text(where)
        n = norm_text(raw)
        rooms: List[str] = []

        if not raw:
            return {"estado": "VACIA", "espacios": [], "sede": "", "detalle": ""}

        # Sede inferida por texto.
        sede = ""
        if n.startswith("CAMPUS"):
            sede = "Campus"
        elif "RIOBAMBA" in n:
            sede = "Riobamba"
        elif "CALLAO" in n:
            sede = "Callao"
        elif "SUIPACHA" in n:
            sede = "Suipacha"

        # 1) códigos de aula. Se toleran espacios: G 006, S 011, RS 002.
        for match in re.finditer(r"\b([A-Z]{1,3})\s*(\d{1,3})\b", n):
            candidate = normalize_room_key(match.group(0))
            if candidate in known_keys and candidate not in rooms:
                rooms.append(candidate)

        # 2) HAM, que no tiene dígitos.
        if re.search(r"\bHAM\b", n) and "HAM" in known_keys and "HAM" not in rooms:
            rooms.append("HAM")

        # 3) nombres especiales/aliases. Evitamos COMEDOR si ya matcheó COMEDOR VIP.
        matched_spans: List[Tuple[int, int]] = []
        for alias_norm, key in named_aliases:
            start = 0
            while True:
                pos = n.find(alias_norm, start)
                if pos < 0:
                    break
                span = (pos, pos + len(alias_norm))
                overlap = any(not (span[1] <= a or span[0] >= b) for a, b in matched_spans)
                if not overlap:
                    if key not in rooms:
                        rooms.append(key)
                    matched_spans.append(span)
                start = pos + len(alias_norm)

        virtual_terms = ["ZOOM", "VIRTUAL", "TEAMS", "GOOGLE MEET", "MEET"]

        # Si la ubicación está expresamente "a confirmar", aunque mencione un espacio conocido,
        # la conservamos como ambigua para no convertirla en una ocupación confirmada.
        if "A CONFIRMAR" in n or "POR CONFIRMAR" in n:
            return {
                "estado": "AMBIGUA_SIN_ESPACIO",
                "espacios": rooms,
                "sede": sede,
                "detalle": "La ubicación menciona un posible espacio, pero figura a confirmar",
            }

        # Clasificación.
        if rooms:
            return {"estado": "RECONOCIDA", "espacios": rooms, "sede": sede, "detalle": ""}

        # Zoom/Meet/etc. se tratan como modalidad virtual. Si además aparece una sede física
        # pero no un espacio concreto, conocemos la sede pero la ubicación física sigue ambigua.
        if any(term in n for term in virtual_terms):
            if sede:
                return {"estado": "AMBIGUA_SIN_ESPACIO", "espacios": [], "sede": sede, "detalle": "Sede identificada, sin espacio físico concreto; además incluye modalidad virtual"}
            return {"estado": "NO_APLICA_VIRTUAL", "espacios": [], "sede": sede, "detalle": ""}

        ambiguous_exact = {
            "CAMPUS", "CAMPUS.", "CAMPUS VICTORIA", "SEDE CALLAO", "SEDE RIOBAMBA", "SEDE SUIPACHA",
            "A CONFIRMAR", "A CONFIRMAR.", "POR CONFIRMAR", "PRESENCIAL", "VIRTUAL",
        }
        if n in ambiguous_exact:
            return {"estado": "AMBIGUA_SIN_ESPACIO", "espacios": [], "sede": sede, "detalle": "No identifica un aula/espacio concreto"}

        # Sullair es un edificio. Una referencia al edificio o a un piso, sin aula/hall concreto,
        # no se convierte en un espacio reservable.
        if sede == "Campus" and "SULLAIR" in n:
            return {"estado": "AMBIGUA_SIN_ESPACIO", "espacios": [], "sede": sede, "detalle": "Referencia al edificio Sullair/piso, sin espacio concreto"}

        if n.startswith("CAMPUS") or "SEDE CALLAO" in n or "SEDE RIOBAMBA" in n or "SEDE SUIPACHA" in n:
            return {"estado": "NO_RECONOCIDA_UDESA", "espacios": [], "sede": sede, "detalle": "Parece una ubicación UdeSA pero no está mapeada"}

        # Lugares externos: se conservan pero no son un problema del catálogo de aulas UdeSA.
        return {"estado": "EXTERNA_O_NO_UDESA", "espacios": [], "sede": sede, "detalle": ""}

    return parse_location


def process_events(path: Path, catalog: MasterCatalog, problems: List[Dict[str, Any]]) -> Dict[str, Any]:
    parse_location = build_location_parser(catalog)
    events = []
    location_counts = Counter()
    location_status_counts = Counter()
    unrecognized_unique = Counter()
    ambiguous_unique = Counter()
    invalid_times = 0
    incomplete_times = 0
    missing_dates = 0

    with XLSXReader(path) as x:
        for sheet in x.sheet_names(include_hidden=False):
            header = x.find_header_row(sheet, ["Fecha", "Actividad", "¿Dónde?"], max_rows=10)
            if not header:
                continue
            for row_num, row in x.dict_rows(sheet, header):
                activity = clean_text(val(row, "Actividad"))
                if not activity:
                    continue
                date = safe_date(val(row, "Fecha"))
                schedule_raw = val(row, "Horario")
                start, end, time_status = parse_time_range(schedule_raw)
                where = clean_text(val(row, "¿Dónde?"))
                parsed = parse_location(where)
                location_counts[where] += 1 if where else 0
                location_status_counts[parsed["estado"]] += 1

                if date is None:
                    missing_dates += 1
                    problems.append({"Tipo": "Evento sin fecha", "Fuente": path.name, "Hoja": sheet, "Fila": row_num, "Valor": activity, "Detalle": where})
                if time_status == "INVALIDO":
                    invalid_times += 1
                    problems.append({"Tipo": "Horario inválido", "Fuente": path.name, "Hoja": sheet, "Fila": row_num, "Valor": format_value(schedule_raw), "Detalle": activity})
                elif time_status == "INCOMPLETO":
                    incomplete_times += 1
                    problems.append({"Tipo": "Horario incompleto", "Fuente": path.name, "Hoja": sheet, "Fila": row_num, "Valor": format_value(schedule_raw), "Detalle": activity})

                if parsed["estado"] == "NO_RECONOCIDA_UDESA":
                    unrecognized_unique[where] += 1
                    problems.append({"Tipo": "Ubicación UdeSA no reconocida", "Fuente": path.name, "Hoja": sheet, "Fila": row_num, "Valor": where, "Detalle": activity})
                elif parsed["estado"] == "AMBIGUA_SIN_ESPACIO":
                    ambiguous_unique[where] += 1
                    problems.append({"Tipo": "Ubicación ambigua", "Fuente": path.name, "Hoja": sheet, "Fila": row_num, "Valor": where, "Detalle": activity})

                events.append({
                    "Mes/Hoja": sheet.strip(),
                    "Fila": row_num,
                    "Fecha": date.isoformat() if date else "",
                    "Horario_original": format_value(schedule_raw),
                    "Hora Desde": start.strftime("%H:%M") if start else "",
                    "Hora Hasta": end.strftime("%H:%M") if end else "",
                    "Estado_horario": time_status,
                    "Evento": activity,
                    "Sector": clean_text(val(row, "Sector a cargo")),
                    "Modalidad": clean_text(val(row, "Virtual / Presencial / Híbrida", "Virtual Presencial Híbrida")),
                    "Ubicación_original": where,
                    "Estado_ubicación": parsed["estado"],
                    "Sede_inferida": parsed["sede"],
                    "Espacios_ID": " | ".join(parsed["espacios"]),
                    "Responsable": clean_text(val(row, "Responsable")),
                })

    return {
        "eventos": len(events),
        "eventos_con_fecha": len(events) - missing_dates,
        "sin_fecha": missing_dates,
        "horarios_invalidos": invalid_times,
        "horarios_incompletos": incomplete_times,
        "ubicaciones_unicas": len([k for k in location_counts if k]),
        "ubicaciones_no_reconocidas_unicas": len(unrecognized_unique),
        "ubicaciones_ambiguas_unicas": len(ambiguous_unique),
        "status_counts": dict(location_status_counts),
        "unrecognized": unrecognized_unique,
        "ambiguous": ambiguous_unique,
        "events_out": events,
    }


def process_careers(path: Path, problems: List[Dict[str, Any]]) -> Dict[str, Any]:
    with XLSXReader(path) as x:
        candidates = x.sheet_names(include_hidden=True)
        model_sheet = next((s for s in candidates if norm_text(s) == "MODELO"), candidates[-1])
        header = x.find_header_row(model_sheet, ["CARRERA", "CAMADA", "SEMESTRE", "CODIGO"])
        rows = x.dict_rows(model_sheet, header)

    normalized = Counter()
    raw = Counter()
    missing_code = 0
    out = []
    for row_num, row in rows:
        career_raw = clean_text(val(row, "CARRERA"))
        if not career_raw:
            continue
        career = canonical_career(career_raw)
        raw[career_raw] += 1
        normalized[career] += 1
        code = norm_text(val(row, "CODIGO"))
        if not code:
            missing_code += 1
            problems.append({"Tipo": "Materia de plan sin código", "Fuente": path.name, "Hoja": model_sheet, "Fila": row_num, "Valor": clean_text(val(row, "NOMBRE")), "Detalle": f"{career} - camada {format_value(val(row, 'CAMADA'))}"})
        out.append({
            "Fila": row_num,
            "Carrera_original": career_raw,
            "Carrera_normalizada": career,
            "Camada": val(row, "CAMADA"),
            "Semestre": clean_text(val(row, "SEMESTRE")),
            "Código": code,
            "Nombre": clean_text(val(row, "NOMBRE")),
        })

    return {
        "carreras_normalizadas": len(normalized),
        "carreras_raw": len(raw),
        "filas": len(out),
        "sin_codigo": missing_code,
        "normalized_counts": dict(normalized),
        "raw_counts": dict(raw),
        "careers_out": out,
    }


def add_manual_spaces(catalog: MasterCatalog) -> None:
    for key, data in MANUAL_SPACES.items():
        e = catalog.get_or_create(key, preferred_name=data["nombre"], tipo=data["tipo"])
        e.tipo = data["tipo"]
        catalog.set_attr(key, "sede", data["sede"], "Definición manual", priority=50, preferred_name=data["nombre"])
        for alias in data.get("aliases", []):
            catalog.add_alias(key, alias)


def finalize_catalog_problems(catalog: MasterCatalog, problems: List[Dict[str, Any]]) -> None:
    for row in catalog.rows():
        if row["Capacidad"] == "" and norm_text(row.get("Tipo", "")) == "AULA":
            problems.append({"Tipo": "Aula sin capacidad", "Fuente": "Catálogo maestro", "Hoja": "", "Fila": "", "Valor": row["Nombre"], "Detalle": f"ID={row['Espacio_ID']} | Sede={row['Sede'] or '[sin sede]'}"})
        if not row["Sede"]:
            problems.append({"Tipo": "Espacio sin sede", "Fuente": "Catálogo maestro", "Hoja": "", "Fila": "", "Valor": row["Nombre"], "Detalle": f"ID={row['Espacio_ID']}"})
    for conflict in catalog.capacity_conflicts:
        problems.append({
            "Tipo": "Conflicto de capacidad entre fuentes",
            "Fuente": "Catálogo maestro",
            "Hoja": "",
            "Fila": "",
            "Valor": conflict["Espacio_ID"],
            "Detalle": f"{conflict['Fuente_anterior']}: {conflict['Capacidad_elegida_antes']} vs {conflict['Fuente_nueva']}: {conflict['Capacidad_nueva']}",
        })
    # Fuerza la construcción del índice y registra aliases ambiguos.
    catalog.alias_index()
    for conflict in catalog.alias_conflicts:
        problems.append({"Tipo": "Alias ambiguo", "Fuente": "Catálogo maestro", "Hoja": "", "Fila": "", "Valor": conflict["Alias"], "Detalle": conflict["Espacios"]})


def make_report(
    files: Dict[str, Path],
    catalog: MasterCatalog,
    aux_stats: Dict[str, Any],
    course_stats: Dict[str, Any],
    aulas_stats: Dict[str, Any],
    event_stats: Dict[str, Any],
    career_stats: Dict[str, Any],
    problems: List[Dict[str, Any]],
) -> str:
    catalog_rows = catalog.rows()
    spaces_with_capacity = sum(1 for r in catalog_rows if r["Capacidad"] != "")
    rooms_without_capacity = [r for r in catalog_rows if r["Capacidad"] == "" and norm_text(r.get("Tipo", "")) == "AULA"]
    nonroom_spaces_without_capacity = [r for r in catalog_rows if r["Capacidad"] == "" and norm_text(r.get("Tipo", "")) != "AULA"]
    spaces_without_sede = [r for r in catalog_rows if not r["Sede"]]

    problem_counts = Counter(p["Tipo"] for p in problems)
    lines = []
    lines.append("=" * 72)
    lines.append(f"{APP_NAME} — DIAGNÓSTICO RONDA 1")
    lines.append(f"Autor: {AUTHOR}")
    lines.append(f"Generado: {dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append("=" * 72)
    lines.append("")
    lines.append("RESUMEN PRINCIPAL")
    lines.append(f"Aulas/espacios detectados: {len(catalog_rows)}")
    lines.append(f"Capacidades detectadas: {spaces_with_capacity}")
    lines.append(f"Cursos detectados: {course_stats['cursos_unicos']}")
    lines.append(f"Asignaciones regulares: {course_stats['asignaciones_validas']}")
    lines.append(f"Ocupaciones semanales en grilla AULAS: {aulas_stats['ocupaciones_grid']}")
    lines.append(f"Reservas especiales: {aulas_stats['reservas']}")
    lines.append(f"Eventos detectados: {event_stats['eventos']}")
    lines.append(f"Comentarios detectados: {aulas_stats['comentarios']}")
    lines.append(f"Carreras detectadas (normalizadas): {career_stats['carreras_normalizadas']}")
    lines.append("")
    lines.append("CALIDAD / PROBLEMAS")
    lines.append(f"Aulas sin capacidad: {len(rooms_without_capacity)}")
    lines.append(f"Otros espacios sin capacidad registrada (informativo): {len(nonroom_spaces_without_capacity)}")
    lines.append(f"Aulas/espacios sin sede: {len(spaces_without_sede)}")
    lines.append(f"Cursos sin código: {course_stats['sin_codigo']}")
    lines.append(f"Cursos sin aula: {course_stats['sin_aula']}")
    lines.append(f"Cursos sin sede: {course_stats['sin_sede']}")
    lines.append(f"Materias de plan/camada sin código: {career_stats['sin_codigo']}")
    lines.append(f"Horarios inválidos (CURSOS): {course_stats['horarios_invalidos']}")
    lines.append(f"Horarios inválidos (Eventos): {event_stats['horarios_invalidos']}")
    lines.append(f"Horarios incompletos (Eventos): {event_stats['horarios_incompletos']}")
    lines.append(f"Eventos sin fecha: {event_stats['sin_fecha']}")
    lines.append(f"Ubicaciones UdeSA no reconocidas (formas únicas): {event_stats['ubicaciones_no_reconocidas_unicas']}")
    lines.append(f"Ubicaciones ambiguas (formas únicas): {event_stats['ubicaciones_ambiguas_unicas']}")
    lines.append(f"Conflictos de capacidad entre fuentes: {len(catalog.capacity_conflicts)}")
    lines.append("")

    lines.append("CARRERAS EN LA FUENTE ACTUAL")
    for career, count in sorted(career_stats["normalized_counts"].items()):
        lines.append(f"  - {career}: {count} registros")
    lines.append("")

    if catalog.capacity_conflicts:
        lines.append("CONFLICTOS DE CAPACIDAD")
        for c in catalog.capacity_conflicts:
            chosen = c["Capacidad_nueva"] if c["Se_elige_nueva"] else c["Capacidad_elegida_antes"]
            lines.append(
                f"  - {c['Espacio_ID']}: {c['Fuente_anterior']}={c['Capacidad_elegida_antes']} | "
                f"{c['Fuente_nueva']}={c['Capacidad_nueva']} | usada={chosen}"
            )
        lines.append("")

    if rooms_without_capacity:
        lines.append("AULAS SIN CAPACIDAD (primeras 40)")
        for r in rooms_without_capacity[:40]:
            lines.append(f"  - {r['Nombre']} [{r['Espacio_ID']}] — {r['Sede'] or 'sin sede'}")
        if len(rooms_without_capacity) > 40:
            lines.append(f"  ... y {len(rooms_without_capacity) - 40} más. Ver catalogo_maestro_espacios.csv")
        lines.append("")

    if nonroom_spaces_without_capacity:
        lines.append("OTROS ESPACIOS SIN CAPACIDAD REGISTRADA (informativo; primeros 40)")
        for r in nonroom_spaces_without_capacity[:40]:
            lines.append(f"  - {r['Nombre']} [{r['Espacio_ID']}] — {r['Sede'] or 'sin sede'}")
        if len(nonroom_spaces_without_capacity) > 40:
            lines.append(f"  ... y {len(nonroom_spaces_without_capacity) - 40} más. Ver catalogo_maestro_espacios.csv")
        lines.append("")

    if spaces_without_sede:
        lines.append("ESPACIOS SIN SEDE")
        for r in spaces_without_sede[:50]:
            lines.append(f"  - {r['Nombre']} [{r['Espacio_ID']}]")
        lines.append("")

    if event_stats["unrecognized"]:
        lines.append("UBICACIONES UDESA NO RECONOCIDAS MÁS FRECUENTES")
        for where, count in event_stats["unrecognized"].most_common(30):
            lines.append(f"  - {count}x — {where}")
        lines.append("")

    if event_stats["ambiguous"]:
        lines.append("UBICACIONES AMBIGUAS MÁS FRECUENTES")
        for where, count in event_stats["ambiguous"].most_common(20):
            lines.append(f"  - {count}x — {where}")
        lines.append("")

    lines.append("ARCHIVOS DE ENTRADA")
    for key, path in files.items():
        lines.append(f"  - {path.name}")
    lines.append("")
    lines.append("ARCHIVOS GENERADOS")
    lines.append("  - diagnostico_ronda1.txt")
    lines.append("  - catalogo_maestro_espacios.csv")
    lines.append("  - cursos_normalizados.csv")
    lines.append("  - comentarios_detectados.csv")
    lines.append("  - reservas_especiales.csv")
    lines.append("  - eventos_detectados.csv")
    lines.append("  - carreras_detectadas.csv")
    lines.append("  - problemas_detectados.csv")
    lines.append("  - resumen_ronda1.json")
    lines.append("")
    lines.append("NOTA SOBRE 'CAMPUS'")
    lines.append("El catálogo contiene un espacio canónico llamado Campus, pero una ubicación del Registro de Eventos que diga solamente 'Campus' se marca como AMBIGUA, porque también puede significar la sede y no necesariamente ese espacio físico.")
    lines.append("")
    return "\n".join(lines)


# ============================================================================
# RONDA 2 — Etapas 4 a 8
# ============================================================================

OUTPUT_DIR_NAME_R2 = "salida_ronda2"


def clean_course_title(raw_name: Any, code: str = "") -> str:
    """Extrae un nombre de materia más legible sin destruir aclaraciones útiles."""
    text = clean_text(raw_name)
    if not text:
        return ""
    if code:
        text = re.sub(rf"^\s*{re.escape(code)}\s+", "", text, flags=re.I)
    # Quita la cola de grupo/semestre que viene del buscador de cursos.
    text = re.sub(r"\s*\(\s*Grupo\s*:\s*[^)]*\)\s*-\s*\d+\s*Semestre\s*$", "", text, flags=re.I)
    text = re.sub(r"\s*\(\s*Grupo\s*:\s*[^)]*\)\s*$", "", text, flags=re.I)
    return clean_text(text)


def safe_int(value: Any) -> Optional[int]:
    if value is None or clean_text(value) == "":
        return None
    try:
        return int(float(value))
    except Exception:
        return None


def safe_iso_date(value: Any) -> str:
    d = safe_date(value)
    return d.isoformat() if d else ""


def build_slot_map(course_rows: Sequence[Dict[str, Any]]) -> Dict[str, Tuple[str, str]]:
    pairs: Dict[str, Counter] = defaultdict(Counter)
    for row in course_rows:
        slot = clean_text(row.get("Slot"))
        start = clean_text(row.get("Hora Desde"))
        end = clean_text(row.get("Hora Hasta"))
        if slot and start and end:
            pairs[slot][(start, end)] += 1
    result: Dict[str, Tuple[str, str]] = {}
    for slot, counter in pairs.items():
        if counter:
            result[slot] = counter.most_common(1)[0][0]
    return result


def course_name_map_from_rows(course_rows: Sequence[Dict[str, Any]]) -> Dict[str, str]:
    names: Dict[str, Counter] = defaultdict(Counter)
    for row in course_rows:
        code = norm_text(row.get("Código"))
        if not code:
            continue
        title = clean_course_title(row.get("Materia"), code)
        if title:
            names[code][title] += 1
    return {code: counter.most_common(1)[0][0] for code, counter in names.items() if counter}


def process_regular_classes_r2(path: Path, catalog: MasterCatalog) -> Dict[str, Any]:
    with XLSXReader(path) as x:
        sheet = x.sheet_names(include_hidden=False)[0]
        header = x.find_header_row(sheet, ["Materia", "Dia", "Espacio", "Cod Materia", "Cant Inscriptos"])
        rows = x.dict_rows(sheet, header)

    out: List[Dict[str, Any]] = []
    codes: Set[str] = set()
    valid_for_students = 0
    valid_for_rooms = 0
    for row_num, row in rows:
        code = norm_text(val(row, "Cod Materia"))
        raw_name = clean_text(val(row, "Materia"))
        title = clean_course_title(raw_name, code)
        day = normalize_day(val(row, "Dia"))
        start = parse_time_token(val(row, "Hora Desde"))
        end = parse_time_token(val(row, "Hora Hasta"))
        slot = normalize_slot(val(row, "N Slot", "Slot"))
        room_raw = clean_text(val(row, "Espacio"))
        room = normalize_room_key(room_raw)
        sede = normalize_sede(val(row, "Sede"))
        inicio = safe_date(val(row, "Inicio"))
        fin = safe_date(val(row, "Fin"))
        if code:
            codes.add(code)
        if code and day and start and end:
            valid_for_students += 1
            if room:
                valid_for_rooms += 1
        capacity = ""
        if room and room in catalog.entries and catalog.entries[room].capacidad is not None:
            capacity = catalog.entries[room].capacidad
        out.append({
            "Fuente": path.name,
            "Hoja": sheet,
            "Fila": row_num,
            "Código": code,
            "Materia": title,
            "Materia_original": raw_name,
            "Tipo Curso": clean_text(val(row, "Tipo Curso")),
            "Tipo Clase": clean_text(val(row, "Tipo Clase")),
            "Docentes": clean_text(val(row, "Docentes")),
            "Inicio vigencia": inicio.isoformat() if inicio else "",
            "Fin vigencia": fin.isoformat() if fin else "",
            "Día": day,
            "Hora Desde": start.strftime("%H:%M") if start else "",
            "Hora Hasta": end.strftime("%H:%M") if end else "",
            "Slot": slot,
            "Aula_original": room_raw,
            "Espacio_ID": room,
            "Capacidad": capacity,
            "Edificio": clean_text(val(row, "Edificio")),
            "Piso": normalize_floor(val(row, "Piso")),
            "Sede": sede,
            "Cant Inscriptos": safe_int(val(row, "Cant Inscriptos")) if safe_int(val(row, "Cant Inscriptos")) is not None else "",
            "Programa": clean_text(val(row, "Programa")),
            "Departamento": clean_text(val(row, "Departamento")),
            "Id Moodle": clean_text(val(row, "Id Moodle")),
        })
    slot_map = build_slot_map(out)
    return {
        "rows": out,
        "cursos_unicos": len(codes),
        "filas": len(out),
        "validas_estudiantes": valid_for_students,
        "validas_aulas": valid_for_rooms,
        "slot_map": slot_map,
        "course_name_map": course_name_map_from_rows(out),
    }


def extract_any_course_code(text: Any, known_codes: Set[str]) -> str:
    n = norm_text(text)
    # Primero prioriza códigos realmente observados en CURSOS.
    for token in re.findall(r"\b[A-Z]{1,5}\d{2,3}\b", n):
        if token in known_codes:
            return token
    # Si no existe en CURSOS, conserva un código plausible en vez de perderlo.
    m = re.search(r"\b([A-Z]{1,5}\d{2,3})\b", n)
    return m.group(1) if m else ""


def parse_grid_class_details(text: str, known_codes: Set[str], course_names: Dict[str, str]) -> Dict[str, Any]:
    raw = clean_text(text)
    n = norm_text(raw)
    code = extract_any_course_code(raw, known_codes)
    result = {
        "Tipo_registro": "",
        "Código": code,
        "Materia": course_names.get(code, "") if code else "",
        "Grupo": "",
        "Tipo_clase_grilla": "",
        "Docente_grilla": "",
        "Inscriptos_grilla": "",
        "Confianza": "ALTA",
        "Detalle_parser": "",
    }

    if n == "HAM - AGARDY":
        result.update({"Tipo_registro": "ETIQUETA_NO_OCUPA", "Confianza": "ALTA", "Detalle_parser": "Etiqueta/nombre del aula, no una ocupación"})
        return result

    if n in {"PARA EVENTOS", "EVENTOS"}:
        result.update({"Tipo_registro": "RESERVADO_PARA_EVENTOS", "Confianza": "ALTA", "Detalle_parser": "Bloque identificado por Alumnos para uso de Eventos"})
        return result

    # Formato típico: C094 G:1 Teó.1 (Apellido)-45
    if code and n.startswith(code):
        result["Tipo_registro"] = "CLASE_REGULAR"
        m_group = re.search(r"\bG\s*:\s*([^\s]+)", raw, flags=re.I)
        if m_group:
            result["Grupo"] = clean_text(m_group.group(1))
        # Segmento entre grupo y docente.
        m_detail = re.search(r"G\s*:\s*[^\s]+\s+(.+?)\s*\((.*?)\)\s*-\s*(\d+)\s*$", raw, flags=re.I)
        if m_detail:
            result["Tipo_clase_grilla"] = clean_text(m_detail.group(1))
            result["Docente_grilla"] = clean_text(m_detail.group(2))
            result["Inscriptos_grilla"] = int(m_detail.group(3))
        else:
            result["Confianza"] = "MEDIA"
            result["Detalle_parser"] = "Código reconocido, pero el resto de la celda no sigue el formato habitual"
        if code not in known_codes:
            result["Confianza"] = "MEDIA"
            result["Detalle_parser"] = "Código plausible que no aparece en CURSOS Primavera 26"
        return result

    # Las siguientes leyendas representan ocupaciones/bloqueos semanales aunque no sean una clase de grado.
    block_keywords = [
        "EDUCACION EJECUTIVA", "EMBA", "MBA", "POSGRADO", "MAESTRIA", "DOCTORADO", "DBA",
        "ADMISION", "INGRESO", "ALPHA", "CONSEJO SUPERIOR", "QUIMICA", "INGENIERIA",
        "CONSULTA", "RECUP", "TALLER", "JESSUP", "TORNEO", "AJEDREZ", "DEPORTES",
        "EXTENSION", "BIODRAMA", "SEM INTERD", "SOCIALES", "CLINICA JURIDICA",
    ]
    if any(k in n for k in block_keywords):
        result["Tipo_registro"] = "BLOQUE_SEMANAL_OTRO"
        result["Confianza"] = "ALTA" if any(k in n for k in ["EDUCACION EJECUTIVA", "EMBA", "MBA", "POSGRADO", "MAESTRIA", "DOCTORADO", "DBA", "ADMISION"]) else "MEDIA"
        if code and not result["Materia"]:
            result["Detalle_parser"] = "Contiene un código de materia no resuelto como clase regular"
        return result

    # Instrucciones, números sueltos o textos que no podemos interpretar con seguridad.
    if re.fullmatch(r"\d+", n) or any(k in n for k in ["CORRER AULA", "LUEGO DEL", "SOLO EN", "A CONFIRMAR"]):
        result.update({"Tipo_registro": "REVISAR", "Confianza": "BAJA", "Detalle_parser": "Texto operativo/ambiguo en la grilla"})
        return result

    # Conservador: una celda no vacía desconocida se conserva como posible bloqueo.
    result.update({"Tipo_registro": "BLOQUE_SEMANAL_OTRO", "Confianza": "BAJA", "Detalle_parser": "Texto no clasificado; se conserva como posible ocupación"})
    return result


def process_aulas_grid_r2(path: Path, catalog: MasterCatalog, slot_map: Dict[str, Tuple[str, str]], known_codes: Set[str], course_names: Dict[str, str]) -> Dict[str, Any]:
    with XLSXReader(path) as x:
        visible = x.sheet_names(include_hidden=False)
        grid_sheet = next((s for s in visible if norm_text(s).startswith("AULAS PRIMAVERA")), visible[0])
        header = x.find_header_row(grid_sheet, ["Aula", "Capacidad", "Slot", "Lunes"])
        rows = x.dict_rows(grid_sheet, header)

    out: List[Dict[str, Any]] = []
    counts = Counter()
    weekdays = [("Lunes", "Lunes"), ("Martes", "Martes"), ("Miercoles", "Miércoles"), ("Jueves", "Jueves"), ("Viernes", "Viernes")]
    for row_num, row in rows:
        room_raw = clean_text(val(row, "Aula"))
        room = normalize_room_key(room_raw)
        if not room:
            continue
        slot = normalize_slot(val(row, "Slot"))
        start, end = slot_map.get(slot, ("", ""))
        cap = parse_numeric_capacity(val(row, "Capacidad"))
        for source_day, day in weekdays:
            cell = clean_text(val(row, source_day))
            if not cell:
                continue
            parsed = parse_grid_class_details(cell, known_codes, course_names)
            counts[parsed["Tipo_registro"]] += 1
            out.append({
                "Fuente": path.name,
                "Hoja": grid_sheet,
                "Fila": row_num,
                "Aula_original": room_raw,
                "Espacio_ID": room,
                "Capacidad": cap if cap is not None else "",
                "Sede": catalog.entries.get(room).sede if room in catalog.entries else "Campus",
                "Día": day,
                "Slot": slot,
                "Hora Desde": start,
                "Hora Hasta": end,
                "Texto_original": cell,
                **parsed,
            })
    return {"rows": out, "counts": counts, "total": len(out), "grid_sheet": grid_sheet}


COMMENT_META_RE = re.compile(r"======\s*ID#[^\s]+\s+(.+?)\s*\((\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})\)\s*", flags=re.S)
DATE_ITEM_RE = re.compile(r"(?<!\d)(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?(?!\d)")


def split_comment_threads(raw: str) -> List[Dict[str, str]]:
    text = clean_text(raw)
    matches = list(COMMENT_META_RE.finditer(text))
    if not matches:
        return [{"autor": "", "timestamp": "", "texto": text}]
    out = []
    for i, m in enumerate(matches):
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        body = clean_text(text[start:end])
        out.append({"autor": clean_text(m.group(1)), "timestamp": clean_text(m.group(2)), "texto": body})
    return out


def split_dated_items(text: str) -> List[Dict[str, Any]]:
    matches = list(DATE_ITEM_RE.finditer(text))
    if not matches:
        return [{"fecha": None, "texto": clean_text(text)}] if clean_text(text) else []
    out = []
    prefix = clean_text(text[:matches[0].start()])
    if prefix:
        out.append({"fecha": None, "texto": prefix})
    for i, m in enumerate(matches):
        day, month = int(m.group(1)), int(m.group(2))
        year_text = m.group(3)
        year = YEAR if not year_text else int(year_text) + (2000 if len(year_text) == 2 else 0)
        try:
            date = dt.date(year, month, day)
        except Exception:
            date = None
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        body = clean_text(text[start:end]).strip(" -;,")
        out.append({"fecha": date, "texto": body})
    return out


def detect_room_codes_in_text(text: str) -> List[str]:
    n = norm_text(text)
    rooms = []
    for m in re.finditer(r"\b([A-Z]{1,4})\s*(\d{1,3})\b", n):
        room = normalize_room_key(m.group(0))
        if room and room not in rooms:
            rooms.append(room)
    if re.search(r"\bHAM\b", n) and "HAM" not in rooms:
        rooms.append("HAM")
    return rooms


def resolve_room_reference(token: str, known_rooms: Set[str]) -> str:
    """Resuelve una referencia de aula; tolera un dígito repetido accidental (H1113 -> H113)."""
    raw = norm_text(token).replace(" ", "")
    direct = normalize_room_key(raw)
    if direct in known_rooms:
        return direct
    m = re.fullmatch(r"([A-Z]{1,4})(\d{4})", raw)
    if m:
        prefix, digits = m.groups()
        candidates = set()
        for i in range(len(digits)):
            reduced = digits[:i] + digits[i+1:]
            cand = normalize_room_key(prefix + reduced)
            if cand in known_rooms:
                candidates.add(cand)
        if len(candidates) == 1:
            return next(iter(candidates))
    return direct if direct in known_rooms else ""


def interpret_comment_item(text: str, current_room: str, current_slot: str, date: Optional[dt.date], expected_day: str, slot_map: Dict[str, Tuple[str, str]], known_rooms: Set[str], known_course_codes: Set[str]) -> Dict[str, Any]:
    raw = clean_text(text)
    n = norm_text(raw)
    actual_day = WEEKDAYS[date.weekday()] if date else ""
    day_ok = (not date) or (not expected_day) or actual_day == expected_day
    effects: List[Dict[str, Any]] = []
    action = "RESERVA"
    origin = ""
    destination = ""
    confidence = "ALTA"
    detail = ""
    mentioned_code = ""
    for token in re.findall(r"\b[A-Z]{1,5}\d{2,3}\b", n):
        if token in known_course_codes:
            mentioned_code = token
            break

    if not date:
        return {
            "accion": "REVISAR",
            "origen": "",
            "destino": "",
            "confianza": "BAJA",
            "detalle": "Comentario/instrucción sin fecha explícita",
            "weekday_real": "",
            "weekday_coincide": "",
            "effects": [{"room": current_room, "slot": current_slot, "effect": "REVISAR", "related": "", "reason": raw}],
            "codigo_mencionado": mentioned_code,
        }

    if not day_ok:
        return {
            "accion": "REVISAR",
            "origen": "",
            "destino": "",
            "confianza": "BAJA",
            "detalle": f"La fecha cae {actual_day}, pero el comentario está en la columna {expected_day}",
            "weekday_real": actual_day,
            "weekday_coincide": "NO",
            "effects": [{"room": current_room, "slot": current_slot, "effect": "REVISAR", "related": "", "reason": raw}],
            "codigo_mencionado": mentioned_code,
        }

    # Liberación explícita del aula actual. Esta regla se evalúa antes del fallback de RESERVA:
    # comentarios como "NO LA USAN" significan que la ocupación semanal de esa celda no aplica
    # en la fecha indicada, no que exista una nueva reserva.
    if re.search(r"\b(?:NO\s+LA\s+USAN|NO\s+LO\s+USAN|NO\s+SE\s+USA)\b", n):
        return {
            "accion": "LIBERACION_EXPLICITA",
            "origen": current_room,
            "destino": "",
            "confianza": "ALTA",
            "detalle": "El comentario indica explícitamente que el aula no se utiliza en esa fecha",
            "weekday_real": actual_day,
            "weekday_coincide": "SÍ",
            "effects": [{"room": current_room, "slot": current_slot, "effect": "LIBERA", "related": "", "reason": raw}],
            "codigo_mencionado": mentioned_code,
        }

    # Salida de la clase/aula actual hacia otra aula.
    move_out = re.search(r"(?:ESTE\s+(?:CURSO|AULA|GRUPO)\s+(?:SE\s+)?VA\s+A)\s+([A-Z]{1,4}\s*\d{1,4})", n)
    if move_out:
        destination = resolve_room_reference(move_out.group(1), known_rooms) or normalize_room_key(move_out.group(1))
        action = "TRASLADO_SALE"
        effects.append({"room": current_room, "slot": current_slot, "effect": "LIBERA", "related": destination, "reason": raw})
        effects.append({"room": destination, "slot": current_slot, "effect": "OCUPA", "related": current_room, "reason": raw})

        # Caso compuesto: se va una clase, pero entra otra en el aula actual.
        if re.search(r"\b(?:Y\s+)?(?:ACA|ACÁ)\s+VIENE\b", n):
            effects.append({"room": current_room, "slot": current_slot, "effect": "OCUPA", "related": "", "reason": "Entra otra actividad/clase según el mismo comentario: " + raw})
            detail = "Traslado de salida y reemplazo/entrada en el aula original"
        return {
            "accion": action, "origen": current_room, "destino": destination, "confianza": confidence,
            "detalle": detail, "weekday_real": actual_day, "weekday_coincide": "SÍ", "effects": effects,
            "codigo_mencionado": mentioned_code,
        }

    # Entrada al aula actual desde otra aula identificada.
    # "viene el curso E490" NO significa que E490 sea un aula; solo liberamos origen si el texto
    # identifica explícitamente una sala ("de la V015", "viene el aula H006", etc.).
    incoming_patterns = [
        r"VIENE\s+EL\s+CURSO.*?\s+DE\s+LA\s+([A-Z]{1,4}\s*\d{1,4})",
        r"VIENE\s+EL\s+AULA\s+([A-Z]{1,4}\s*\d{1,4})",
        r"(?:ACA|ACÁ)\s+VIENE\s+(?:EL\s+)?(?:AULA\s+)?([A-Z]{1,4}\s*\d{1,4})",
    ]
    for pat in incoming_patterns:
        m = re.search(pat, n)
        if m:
            origin = resolve_room_reference(m.group(1), known_rooms)
            if not origin:
                # La referencia parece aula, pero no pudimos resolverla de forma segura.
                action = "ENTRA_SIN_ORIGEN"
                confidence = "MEDIA"
                effects.append({"room": current_room, "slot": current_slot, "effect": "OCUPA", "related": "", "reason": raw})
                detail = f"Se entiende que el aula actual se ocupa; referencia de origen no resuelta: {clean_text(m.group(1))}"
                return {
                    "accion": action, "origen": "", "destino": current_room, "confianza": confidence,
                    "detalle": detail, "weekday_real": actual_day, "weekday_coincide": "SÍ", "effects": effects,
                    "codigo_mencionado": mentioned_code,
                }
            action = "TRASLADO_ENTRA"
            effects.append({"room": current_room, "slot": current_slot, "effect": "OCUPA", "related": origin, "reason": raw})
            if origin and origin != current_room:
                effects.append({"room": origin, "slot": current_slot, "effect": "LIBERA", "related": current_room, "reason": raw})
            return {
                "accion": action, "origen": origin, "destino": current_room, "confianza": confidence,
                "detalle": detail, "weekday_real": actual_day, "weekday_coincide": "SÍ", "effects": effects,
                "codigo_mencionado": mentioned_code,
            }

    # Entrada mencionada, pero sin aula origen identificable.
    if re.search(r"\b(?:VIENE\s+EL\s+(?:CURSO|AULA)|(?:ACA|ACÁ)\s+VIENE)\b", n):
        action = "ENTRA_SIN_ORIGEN"
        confidence = "MEDIA"
        effects.append({"room": current_room, "slot": current_slot, "effect": "OCUPA", "related": "", "reason": raw})
        detail = "Se entiende que el aula actual se ocupa, pero no se pudo identificar el aula de origen"
    else:
        # Por defecto, una instrucción fechada en la celda implica reserva/ocupación del aula actual.
        effects.append({"room": current_room, "slot": current_slot, "effect": "OCUPA", "related": "", "reason": raw})

    # Liberaciones explícitas adicionales: "libera slot 5 H006 y 6 H007".
    if "LIBERA" in n:
        # Primera forma: SLOT 5 H006
        found_release = False
        for m in re.finditer(r"(?:SLOT\s*)?(\d)\s+([A-Z]{1,4}\s*\d{1,3})", n):
            rel_slot = m.group(1)
            rel_room = normalize_room_key(m.group(2))
            # Solo se interpreta como liberación si aparece después de la palabra LIBERA en el texto.
            if m.start() >= n.find("LIBERA"):
                effects.append({"room": rel_room, "slot": rel_slot, "effect": "LIBERA", "related": current_room, "reason": raw})
                found_release = True
        if found_release:
            action = "RESERVA_CON_LIBERACION"

    return {
        "accion": action, "origen": origin, "destino": destination, "confianza": confidence,
        "detalle": detail, "weekday_real": actual_day, "weekday_coincide": "SÍ", "effects": effects,
        "codigo_mencionado": mentioned_code,
    }


def process_comments_r2(path: Path, catalog: MasterCatalog, slot_map: Dict[str, Tuple[str, str]], known_course_codes: Set[str]) -> Dict[str, Any]:
    with XLSXReader(path) as x:
        visible = x.sheet_names(include_hidden=False)
        grid_sheet = next((s for s in visible if norm_text(s).startswith("AULAS PRIMAVERA")), visible[0])
        header = x.find_header_row(grid_sheet, ["Aula", "Capacidad", "Slot", "Lunes"])
        grid_rows = x.dict_rows(grid_sheet, header)
        grid_lookup = {rn: row for rn, row in grid_rows}
        raw_rows = {rn: row for rn, row in x.rows(grid_sheet)}
        header_raw = raw_rows.get(header, {})
        col_to_header = {col: clean_text(value) for col, value in header_raw.items()}
        comments = x.comments(grid_sheet)

    interpreted: List[Dict[str, Any]] = []
    effects_out: List[Dict[str, Any]] = []
    action_counts = Counter()
    effect_counts = Counter()
    known_rooms = set(catalog.entries.keys())
    sequence = 0
    for c in comments:
        context = grid_lookup.get(c["fila"], {})
        room_raw = clean_text(val(context, "Aula"))
        current_room = normalize_room_key(room_raw)
        slot = normalize_slot(val(context, "Slot"))
        expected_day = normalize_day(col_to_header.get(c["columna"], ""))
        start, end = slot_map.get(slot, ("", ""))
        raw_comment = c["comentario"]
        for thread_index, thread in enumerate(split_comment_threads(raw_comment), start=1):
            items = split_dated_items(thread["texto"])
            for item_index, item in enumerate(items, start=1):
                sequence += 1
                date = item["fecha"]
                item_text = item["texto"]
                parsed = interpret_comment_item(item_text, current_room, slot, date, expected_day, slot_map, known_rooms, known_course_codes)
                action_counts[parsed["accion"]] += 1
                item_id = f"COM-{sequence:04d}"
                interpreted.append({
                    "ID": item_id,
                    "Fuente": path.name,
                    "Hoja": grid_sheet,
                    "Celda": c["celda"],
                    "Aula_original": room_raw,
                    "Espacio_ID": current_room,
                    "Día_columna": expected_day,
                    "Slot": slot,
                    "Hora Desde": start,
                    "Hora Hasta": end,
                    "Autor": thread["autor"] or c.get("autor", ""),
                    "Timestamp comentario": thread["timestamp"],
                    "Fecha": date.isoformat() if date else "",
                    "Día_fecha": parsed["weekday_real"],
                    "Coincide día": parsed["weekday_coincide"],
                    "Texto_instrucción": item_text,
                    "Acción": parsed["accion"],
                    "Aula_origen": parsed["origen"],
                    "Aula_destino": parsed["destino"],
                    "Código_mencionado": parsed.get("codigo_mencionado", ""),
                    "Confianza": parsed["confianza"],
                    "Detalle_parser": parsed["detalle"],
                    "Comentario_original_completo": raw_comment,
                })
                for effect_index, effect in enumerate(parsed["effects"], start=1):
                    effect_slot = clean_text(effect.get("slot")) or slot
                    e_start, e_end = slot_map.get(effect_slot, ("", ""))
                    effect_counts[effect["effect"]] += 1
                    effects_out.append({
                        "ID comentario": item_id,
                        "Efecto_n": effect_index,
                        "Fecha": date.isoformat() if date else "",
                        "Día": WEEKDAYS[date.weekday()] if date else expected_day,
                        "Espacio_ID": effect["room"],
                        "Slot": effect_slot,
                        "Hora Desde": e_start,
                        "Hora Hasta": e_end,
                        "Efecto": effect["effect"],
                        "Aula_relacionada": effect.get("related", ""),
                        "Motivo": effect.get("reason", item_text),
                        "Confianza": parsed["confianza"],
                        "Celda_origen": c["celda"],
                    })
    return {
        "comments_count": len(comments),
        "items": interpreted,
        "effects": effects_out,
        "action_counts": action_counts,
        "effect_counts": effect_counts,
    }


def parse_reservation_time(value: Any) -> Tuple[str, str]:
    t = parse_time_token(value)
    return (t.strftime("%H:%M"), "OK") if t else ("", "INVALIDO" if clean_text(value) else "VACIO")


def process_special_reservations_r2(path: Path, catalog: MasterCatalog) -> Dict[str, Any]:
    out: List[Dict[str, Any]] = []
    with XLSXReader(path) as x:
        visible = x.sheet_names(include_hidden=False)
        grid_sheet = next((s for s in visible if norm_text(s).startswith("AULAS PRIMAVERA")), visible[0])
        for sheet in visible:
            if sheet == grid_sheet:
                continue
            h = x.find_header_row(sheet, ["FECHA", "AULA"], max_rows=20)
            if not h:
                continue
            for row_num, row in x.dict_rows(sheet, h):
                date = safe_date(val(row, "FECHA"))
                room_raw = clean_text(val(row, "AULA"))
                room = normalize_room_key(room_raw)
                if not date and not room and not clean_text(val(row, "MOTIVO")):
                    continue
                start = parse_time_token(val(row, "HORA DESDE"))
                end = parse_time_token(val(row, "HORA HASTA"))
                if date and room and start and end:
                    state = "OCUPA"
                    confidence = "ALTA"
                elif room:
                    state = "REVISAR"
                    confidence = "BAJA"
                else:
                    state = "SIN_AULA"
                    confidence = "BAJA"
                cap = ""
                if room in catalog.entries and catalog.entries[room].capacidad is not None:
                    cap = catalog.entries[room].capacidad
                out.append({
                    "Fuente": path.name,
                    "Hoja": sheet,
                    "Fila": row_num,
                    "Fecha": date.isoformat() if date else "",
                    "Hora Desde": start.strftime("%H:%M") if start else "",
                    "Hora Hasta": end.strftime("%H:%M") if end else "",
                    "Espacio_ID": room,
                    "Aula_original": room_raw,
                    "Capacidad": cap,
                    "Profesor/Pedido por": clean_text(val(row, "PROFESOR", "PROFESOR o PEDIDO POR")),
                    "Motivo": clean_text(val(row, "MOTIVO")),
                    "Antes": clean_text(val(row, "ANTES")),
                    "Observaciones": clean_text(val(row, "OBS")),
                    "Efecto": state,
                    "Confianza": confidence,
                })
    return {"rows": out, "valid": sum(1 for r in out if r["Efecto"] == "OCUPA"), "review": sum(1 for r in out if r["Efecto"] == "REVISAR")}


def parse_event_schedule_r2(value: Any) -> Dict[str, Any]:
    raw = clean_text(value)
    n = norm_text(raw)
    if not raw:
        return {"start": "", "end": "", "status": "VACIO", "effect": "REVISAR"}
    if n in {"TODO EL DIA", "DURANTE EL DIA"}:
        return {"start": "00:00", "end": "23:59", "status": "TODO_EL_DIA", "effect": "OCUPA"}
    if n == "FERIADO":
        return {"start": "", "end": "", "status": "FERIADO", "effect": "NO_APLICA"}
    if "A CONFIRMAR" in n or n in {"TBC", "TBD"}:
        return {"start": "", "end": "", "status": "A_CONFIRMAR", "effect": "REVISAR"}
    start, end, status = parse_time_range(value)
    if status == "OK" and start and end:
        return {"start": start.strftime("%H:%M"), "end": end.strftime("%H:%M"), "status": "OK", "effect": "OCUPA"}
    if status == "INCOMPLETO" and start:
        return {"start": start.strftime("%H:%M"), "end": "", "status": "INCOMPLETO", "effect": "REVISAR"}
    return {"start": "", "end": "", "status": "INVALIDO", "effect": "REVISAR"}


def process_events_r2(path: Path, catalog: MasterCatalog) -> Dict[str, Any]:
    parse_location = build_location_parser(catalog)
    events: List[Dict[str, Any]] = []
    occupancies: List[Dict[str, Any]] = []
    status_counts = Counter()
    effect_counts = Counter()
    seq = 0
    with XLSXReader(path) as x:
        for sheet in x.sheet_names(include_hidden=False):
            header = x.find_header_row(sheet, ["Fecha", "Actividad", "¿Dónde?"], max_rows=10)
            if not header:
                continue
            for row_num, row in x.dict_rows(sheet, header):
                activity = clean_text(val(row, "Actividad"))
                if not activity:
                    continue
                seq += 1
                eid = f"EV-{seq:04d}"
                date = safe_date(val(row, "Fecha"))
                where = clean_text(val(row, "¿Dónde?"))
                loc = parse_location(where)
                sched = parse_event_schedule_r2(val(row, "Horario"))
                status_counts[(loc["estado"], sched["status"])] += 1
                event_row = {
                    "ID": eid,
                    "Fuente": path.name,
                    "Mes/Hoja": sheet.strip(),
                    "Fila": row_num,
                    "Fecha": date.isoformat() if date else "",
                    "Horario_original": format_value(val(row, "Horario")),
                    "Hora Desde": sched["start"],
                    "Hora Hasta": sched["end"],
                    "Estado_horario": sched["status"],
                    "Evento": activity,
                    "Sector": clean_text(val(row, "Sector a cargo")),
                    "Modalidad": clean_text(val(row, "Virtual / Presencial / Híbrida", "Virtual Presencial Híbrida")),
                    "Ubicación_original": where,
                    "Estado_ubicación": loc["estado"],
                    "Sede_inferida": loc["sede"],
                    "Espacios_ID": " | ".join(loc["espacios"]),
                    "Responsable": clean_text(val(row, "Responsable")),
                }
                events.append(event_row)

                # Una ocupación/revisión por espacio concreto reconocido o sugerido.
                if loc["espacios"]:
                    for room in loc["espacios"]:
                        if loc["estado"] == "RECONOCIDA" and date and sched["effect"] == "OCUPA":
                            effect = "OCUPA"
                            confidence = "ALTA"
                        else:
                            # Si menciona un aula concreta pero la ubicación/horario está a confirmar, no la recomendamos.
                            effect = "REVISAR"
                            confidence = "BAJA"
                        effect_counts[effect] += 1
                        cap = ""
                        if room in catalog.entries and catalog.entries[room].capacidad is not None:
                            cap = catalog.entries[room].capacidad
                        occupancies.append({
                            "ID evento": eid,
                            "Fecha": date.isoformat() if date else "",
                            "Hora Desde": sched["start"],
                            "Hora Hasta": sched["end"],
                            "Espacio_ID": room,
                            "Capacidad": cap,
                            "Efecto": effect,
                            "Confianza": confidence,
                            "Evento": activity,
                            "Ubicación_original": where,
                            "Estado_ubicación": loc["estado"],
                            "Estado_horario": sched["status"],
                            "Sede": loc["sede"],
                            "Sector": clean_text(val(row, "Sector a cargo")),
                        })
    return {
        "events": events,
        "occupancies": occupancies,
        "events_count": len(events),
        "occupancy_count": len(occupancies),
        "effect_counts": effect_counts,
        "status_counts": status_counts,
    }


def make_report_r2(
    class_stats: Dict[str, Any],
    grid_stats: Dict[str, Any],
    comment_stats: Dict[str, Any],
    reservation_stats: Dict[str, Any],
    event_stats: Dict[str, Any],
    slot_map: Dict[str, Tuple[str, str]],
) -> str:
    lines = []
    lines.append("=" * 80)
    lines.append("UdeSA Horarios — DIAGNÓSTICO RONDA 2")
    lines.append(f"Autor: {AUTHOR}")
    lines.append(f"Generado: {dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append("=" * 80)
    lines.append("")
    lines.append("ETAPA 3")
    lines.append("Pendiente por diseño: catálogo académico completo de todas las carreras/camadas. No bloquea esta ronda.")
    lines.append("")
    lines.append("ETAPA 4 — CLASES REGULARES")
    lines.append(f"Filas de CURSOS procesadas: {class_stats['filas']}")
    lines.append(f"Códigos únicos: {class_stats['cursos_unicos']}")
    lines.append(f"Registros válidos para disponibilidad estudiantil: {class_stats['validas_estudiantes']}")
    lines.append(f"Registros válidos con aula física: {class_stats['validas_aulas']}")
    lines.append("Slots detectados:")
    for slot in sorted(slot_map, key=lambda x: int(x) if x.isdigit() else 999):
        start, end = slot_map[slot]
        lines.append(f"  - Slot {slot}: {start}–{end}")
    lines.append("")
    lines.append("ETAPA 5 — GRILLA AULAS")
    lines.append(f"Celdas no vacías procesadas: {grid_stats['total']}")
    for kind, count in grid_stats['counts'].most_common():
        lines.append(f"  - {kind}: {count}")
    lines.append("")
    lines.append("ETAPA 6 — COMENTARIOS")
    lines.append(f"Comentarios de celda originales: {comment_stats['comments_count']}")
    lines.append(f"Instrucciones fechadas/segmentadas: {len(comment_stats['items'])}")
    lines.append(f"Efectos de aula generados: {len(comment_stats['effects'])}")
    lines.append("Acciones interpretadas:")
    for kind, count in comment_stats['action_counts'].most_common():
        lines.append(f"  - {kind}: {count}")
    lines.append("Efectos:")
    for kind, count in comment_stats['effect_counts'].most_common():
        lines.append(f"  - {kind}: {count}")
    review_items = sum(1 for r in comment_stats['items'] if r['Acción'] == 'REVISAR' or r['Confianza'] == 'BAJA')
    lines.append(f"Instrucciones que requieren revisión: {review_items}")
    lines.append("")
    lines.append("ETAPA 7 — RESERVAS ESPECIALES")
    lines.append(f"Registros encontrados: {len(reservation_stats['rows'])}")
    lines.append(f"Reservas completas que ocupan: {reservation_stats['valid']}")
    lines.append(f"Reservas para revisar: {reservation_stats['review']}")
    lines.append("")
    lines.append("ETAPA 8 — EVENTOS")
    lines.append(f"Eventos procesados: {event_stats['events_count']}")
    lines.append(f"Efectos sobre espacios concretos: {event_stats['occupancy_count']}")
    for kind, count in event_stats['effect_counts'].most_common():
        lines.append(f"  - {kind}: {count}")
    all_day = sum(1 for r in event_stats['events'] if r['Estado_horario'] == 'TODO_EL_DIA')
    incomplete = sum(1 for r in event_stats['events'] if r['Estado_horario'] == 'INCOMPLETO')
    lines.append(f"Eventos interpretados como todo el día: {all_day}")
    lines.append(f"Eventos con una sola hora y por lo tanto REVISAR si tienen aula: {incomplete}")
    lines.append("")
    lines.append("ARCHIVOS GENERADOS")
    for name in [
        "clases_regulares.csv",
        "grilla_aulas_normalizada.csv",
        "comentarios_interpretados.csv",
        "efectos_comentarios.csv",
        "reservas_especiales_normalizadas.csv",
        "eventos_normalizados.csv",
        "ocupaciones_eventos.csv",
        "catalogo_maestro_espacios.csv",
        "diagnostico_ronda2.txt",
        "resumen_ronda2.json",
    ]:
        lines.append(f"  - {name}")
    lines.append("")
    lines.append("REGLA CONSERVADORA")
    lines.append("Un dato ambiguo con aula concreta nunca se convierte en LIBRE: queda como REVISAR para que el motor de ocupación lo bloquee hasta resolverlo.")
    return "\n".join(lines)


def main() -> None:
    base = Path(__file__).resolve().parent
    print(f"\n{APP_NAME} — Ronda 2")
    print(f"Autor: {AUTHOR}")
    print(f"Carpeta de trabajo: {base}")

    files = find_required_files(base)
    output_dir = base / OUTPUT_DIR_NAME_R2
    output_dir.mkdir(parents=True, exist_ok=True)

    # Reconstruimos el catálogo de Ronda 1 en memoria porque los parsers de Ronda 2 lo necesitan.
    problems: List[Dict[str, Any]] = []
    catalog = MasterCatalog()
    print("\n[1/8] Reconstruyendo catálogo maestro de espacios...")
    process_aux_catalog(files["catalogo_aux"], catalog, problems)
    base_course_stats = process_courses(files["cursos"], catalog, problems)
    base_aulas_stats = process_aulas(files["aulas"], catalog, problems)
    add_manual_spaces(catalog)
    for alias, key in MANUAL_ALIASES.items():
        if key in catalog.entries:
            catalog.add_alias(key, alias.title() if alias != "HAM" else alias)
    # Leer eventos una vez ayuda a validar el catálogo/aliases; sus filas de R1 no se exportan acá.
    process_events(files["eventos"], catalog, problems)
    finalize_catalog_problems(catalog, problems)

    print("[2/8] Etapa 4: normalizando clases regulares...")
    class_stats = process_regular_classes_r2(files["cursos"], catalog)
    slot_map = class_stats["slot_map"]
    course_names = class_stats["course_name_map"]
    known_codes = set(course_names.keys())

    print("[3/8] Etapa 5: interpretando grilla AULAS...")
    grid_stats = process_aulas_grid_r2(files["aulas"], catalog, slot_map, known_codes, course_names)

    print("[4/8] Etapa 6: interpretando comentarios y movimientos...")
    comment_stats = process_comments_r2(files["aulas"], catalog, slot_map, known_codes)

    print("[5/8] Etapa 7: normalizando reservas especiales...")
    reservation_stats = process_special_reservations_r2(files["aulas"], catalog)

    print("[6/8] Etapa 8: normalizando Eventos y ubicaciones...")
    event_stats = process_events_r2(files["eventos"], catalog)

    print("[7/8] Exportando tablas normalizadas...")
    write_csv(output_dir / "catalogo_maestro_espacios.csv", catalog.rows())
    write_csv(output_dir / "clases_regulares.csv", class_stats["rows"])
    write_csv(output_dir / "grilla_aulas_normalizada.csv", grid_stats["rows"])
    write_csv(output_dir / "comentarios_interpretados.csv", comment_stats["items"])
    write_csv(output_dir / "efectos_comentarios.csv", comment_stats["effects"])
    write_csv(output_dir / "reservas_especiales_normalizadas.csv", reservation_stats["rows"])
    write_csv(output_dir / "eventos_normalizados.csv", event_stats["events"])
    write_csv(output_dir / "ocupaciones_eventos.csv", event_stats["occupancies"])

    print("[8/8] Generando diagnóstico...")
    report = make_report_r2(class_stats, grid_stats, comment_stats, reservation_stats, event_stats, slot_map)
    (output_dir / "diagnostico_ronda2.txt").write_text(report, encoding="utf-8")
    summary = {
        "app": APP_NAME,
        "autor": AUTHOR,
        "generado": dt.datetime.now().isoformat(timespec="seconds"),
        "etapa_3": "PENDIENTE_POR_DISENO",
        "clases": {
            "filas": class_stats["filas"],
            "cursos_unicos": class_stats["cursos_unicos"],
            "validas_estudiantes": class_stats["validas_estudiantes"],
            "validas_aulas": class_stats["validas_aulas"],
            "slots": {k: {"desde": v[0], "hasta": v[1]} for k, v in slot_map.items()},
        },
        "grilla": {"total": grid_stats["total"], "tipos": dict(grid_stats["counts"])},
        "comentarios": {
            "originales": comment_stats["comments_count"],
            "instrucciones": len(comment_stats["items"]),
            "efectos": len(comment_stats["effects"]),
            "acciones": dict(comment_stats["action_counts"]),
            "tipos_efecto": dict(comment_stats["effect_counts"]),
        },
        "reservas": {"registros": len(reservation_stats["rows"]), "ocupa": reservation_stats["valid"], "revisar": reservation_stats["review"]},
        "eventos": {
            "registros": event_stats["events_count"],
            "efectos_espacios": event_stats["occupancy_count"],
            "tipos_efecto": dict(event_stats["effect_counts"]),
        },
    }
    (output_dir / "resumen_ronda2.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")

    print("\n" + report)
    print(f"\nListo. Los archivos de la Ronda 2 quedaron en:\n{output_dir}\n")
    try:
        input("Presioná ENTER para cerrar...")
    except EOFError:
        pass


if __name__ == "__main__":
    main()
