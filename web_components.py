from __future__ import annotations

import html
from typing import Any, Dict, List, Sequence

import streamlit as st

from web_logic import parse_extra_codes, population_payload, status_icon


def page_header(title: str, subtitle: str) -> None:
    st.markdown(f"# {title}")
    st.markdown(f"<div class='page-subtitle'>{html.escape(subtitle)}</div>", unsafe_allow_html=True)


def period_caption(period) -> None:
    st.caption(f"Período activo: {period.label} · {period.start.strftime('%d/%m/%Y')}–{period.end.strftime('%d/%m/%Y')}")


def render_warnings(warnings: Sequence[str], title: str = "Advertencias") -> None:
    rows = [str(x).strip() for x in warnings if str(x).strip()]
    if not rows:
        return
    with st.expander(f"{title} ({len(rows)})", expanded=False):
        for row in rows:
            st.warning(row, icon="⚠️")


def population_controls(service, career: str, prefix: str) -> Dict[str, Any]:
    st.markdown(f"**{career}**")
    window = service.resolver.current_cohort_window(career)
    cohorts = st.multiselect(
        "Camadas específicas",
        options=window,
        default=[],
        key=f"{prefix}_cohorts",
        help="Dejalo vacío para usar todas las camadas operativas del período.",
    )
    sites = service.listar_sedes(career)
    site_choice = st.selectbox(
        "Sede académica",
        options=["Automática"] + sites,
        key=f"{prefix}_academic_site",
        help="Automática prioriza Campus cuando existe y aprovecha las variantes disponibles de cada camada. Elegí una sede o variante solo si querés restringir la búsqueda.",
    )
    site_value = None if site_choice == "Automática" else site_choice
    variants = service.listar_variantes(career, sede=site_value)
    variant_choice = st.selectbox(
        "Variante",
        options=["Automática"] + variants,
        key=f"{prefix}_variant",
    )
    variant_value = None if variant_choice == "Automática" else variant_choice
    include_opt = st.checkbox(
        "Incluir optativas/electivas identificadas",
        value=False,
        key=f"{prefix}_optatives",
        help="Solo agrega optativas cuyo curso concreto puede identificarse. Nunca inventa requisitos genéricos.",
    )
    extra_text = st.text_input(
        "Códigos adicionales",
        value="",
        key=f"{prefix}_extra_codes",
        placeholder="Ej.: P318, E020",
        help="Opcional. Sirve para casos especiales sin convertir los códigos en la forma normal de usar la app.",
    )
    return population_payload(
        career,
        cohorts=cohorts,
        academic_site=site_value,
        variant=variant_value,
        include_optatives=include_opt,
        extra_codes=parse_extra_codes(extra_text),
    )


def render_population_errors(resolved: Sequence[Dict[str, Any]]) -> None:
    for result in resolved:
        if result.get("Estado") == "OK":
            continue
        career = result.get("Carrera", "Carrera")
        state = result.get("Estado", "")
        if state == "NECESITA_FILTRO":
            st.error(f"{career}: esta combinación específica requiere elegir sede/variante en ‘Más filtros’.")
            options = result.get("Opciones", [])
            if options:
                st.dataframe(options, use_container_width=True, hide_index=True)
        else:
            message = " · ".join(result.get("Advertencias", [])) or state
            st.error(f"{career}: {message}")


def room_card(row: Dict[str, Any]) -> None:
    with st.container(border=True):
        c1, c2, c3 = st.columns([1.2, 2.2, 1.2])
        with c1:
            st.markdown(f"### {row.get('Espacio_ID', '')}")
        with c2:
            name = row.get("Nombre", "")
            if name and str(name).upper() != str(row.get("Espacio_ID", "")).upper():
                st.write(name)
            st.caption(f"Sede: {row.get('Sede', 'N/A')}")
        with c3:
            cap = row.get("Capacidad", "") or "N/A"
            st.metric("Capacidad", cap)
        if row.get("Reservada para eventos") == "Sí":
            st.success("La grilla la reserva para eventos en este horario.", icon="⭐")
        if row.get("Estado") == "REVISAR":
            st.warning("Este espacio requiere revisión; no se considera libre de manera segura.")
        if row.get("Notas"):
            st.caption(str(row.get("Notas")))


def ranking_card(rank: int, row: Dict[str, Any]) -> None:
    with st.container(border=True):
        title = f"{rank}. {row.get('Día', '')} {row.get('Fecha', '')} · {row.get('Hora Desde', '')}–{row.get('Hora Hasta', '')}"
        st.markdown(f"### {title}")
        c1, c2, c3 = st.columns(3)
        c1.metric("Alumnos ocupados", int(row.get("Alumnos ocupados", 0) or 0))
        c2.metric("Materias afectadas", int(row.get("Materias afectadas", 0) or 0))
        c3.metric("Aulas disponibles", int(row.get("Aulas disponibles", 0) or 0))
        rooms = str(row.get("Aulas", "")).strip()
        if rooms:
            st.caption("Aulas: " + rooms)


def schedule_row(row: Dict[str, Any]) -> None:
    status = str(row.get("Estado", ""))
    with st.container(border=True):
        c1, c2, c3 = st.columns([1.4, 1.2, 4])
        c1.markdown(f"**{row.get('Hora Desde', '')}–{row.get('Hora Hasta', '')}**")
        c2.markdown(f"{status_icon(status)} **{status}**")
        detail = str(row.get("Detalle", "")).strip()
        notes = str(row.get("Notas", "")).strip()
        if detail:
            c3.write(detail)
        elif status == "LIBRE":
            c3.write("Sin ocupaciones registradas.")
        if notes:
            c3.caption(notes)
