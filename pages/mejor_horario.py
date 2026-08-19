from __future__ import annotations

import datetime as dt

import streamlit as st

from web_components import page_header, period_caption, population_controls, ranking_card, render_population_errors, render_warnings
from web_logic import clamp_default_window, dates_between, normalize_date_range, period_label_map, physical_sites, slot_options, validate_populations
from web_runtime import get_registry, get_runtime

page_header(
    "Buscar mejor horario",
    "Elegí una o más carreras. Sin tocar filtros avanzados usamos el período actual, todas las camadas operativas, materias obligatorias y la sede Campus para las aulas.",
)

registry_light = get_registry()
period_map = period_label_map(registry_light)
current_pid = registry_light.active_period_id
period_key = "mh_period"
selected_pid = st.session_state.get(period_key, current_pid)
if selected_pid not in registry_light.period_defs:
    selected_pid = current_pid

with st.spinner("Cargando fuentes y motor de horarios..."):
    registry, period, engine, resolver, service, stats = get_runtime(selected_pid)
period_caption(period)

careers = service.listar_carreras()
selected_careers = st.multiselect(
    "¿Para quién buscás horario?",
    options=careers,
    placeholder="Elegí una o más carreras",
    help="Podés combinar varias carreras. Las camadas se toman automáticamente salvo que uses Más filtros.",
)

c1, c2 = st.columns([2, 1])
def_start, def_end = clamp_default_window(period)
with c1:
    date_value = st.date_input(
        "Fechas a considerar",
        value=(def_start, def_end),
        min_value=period.start,
        max_value=period.end,
        key=f"mh_dates_{selected_pid}",
        format="DD/MM/YYYY",
    )
with c2:
    min_capacity = st.number_input("Capacidad mínima del aula", min_value=0, max_value=1000, value=0, step=5)

population_configs = []
weekdays_only = True
room_site = service.default_sede
slot_pairs = slot_options(engine.slot_map)
selected_slots = [slot for _, slot in slot_pairs]
top_n = 12

with st.expander("＋ Más filtros", expanded=False):
    labels = list(period_map)
    current_label = next((lab for lab, pid in period_map.items() if pid == selected_pid), labels[0])
    selected_label = st.selectbox("Período", options=labels, index=labels.index(current_label), key="mh_period_label")
    chosen_pid = period_map[selected_label]
    if chosen_pid != selected_pid:
        st.session_state[period_key] = chosen_pid
        st.rerun()
    st.session_state[period_key] = chosen_pid

    sites = physical_sites(engine)
    room_site = st.selectbox("Sede para las aulas", options=sites, index=sites.index(service.default_sede) if service.default_sede in sites else 0)
    weekdays_only = st.checkbox("Solo lunes a viernes", value=True)
    slot_labels = [label for label, _ in slot_pairs]
    chosen_slot_labels = st.multiselect("Slots", options=slot_labels, default=slot_labels)
    selected_slots = [dict(slot_pairs)[label] for label in chosen_slot_labels]
    top_n = st.slider("Cantidad máxima de resultados", min_value=5, max_value=30, value=12)

    if selected_careers:
        st.divider()
        st.caption("Filtros académicos por carrera. Si no tocás nada, se incluyen todas las camadas vigentes y se aprovecha toda la información curricular confiable disponible; las variantes solo restringen si las elegís explícitamente.")
        for i, career in enumerate(selected_careers):
            with st.container(border=True):
                population_configs.append(population_controls(service, career, f"mh_pop_{i}_{career}"))

if not population_configs:
    population_configs = [{"Carrera": career, "Camadas": "TODAS", "Sede": None, "Variante": None, "Incluir_optativas": False, "Códigos_extra": []} for career in selected_careers]

if st.button("Buscar mejores horarios", type="primary", use_container_width=True):
    if not selected_careers:
        st.error("Elegí al menos una carrera.")
    elif not selected_slots:
        st.error("Elegí al menos un slot.")
    else:
        start, end = normalize_date_range(date_value)
        if (end - start).days > 60:
            st.error("Para mantener la consulta ágil, elegí un rango de hasta 60 días.")
        else:
            dates = dates_between(start, end, weekdays_only=weekdays_only)
            if not dates:
                st.error("El rango seleccionado no contiene fechas para consultar.")
            else:
                valid, resolved = validate_populations(service, population_configs)
                if not valid:
                    render_population_errors(resolved)
                else:
                    with st.spinner("Combinando alumnos, clases, aulas, comentarios, reservas y Eventos..."):
                        result = service.buscar_mejor_horario(
                            population_configs,
                            dates,
                            capacidad_minima=int(min_capacity),
                            sede_aulas=room_site,
                            slots=selected_slots,
                            top_n=top_n,
                        )
                    ranking = result.get("Ranking", [])
                    render_warnings(result.get("Advertencias", []))
                    with st.expander("Ver poblaciones y códigos utilizados", expanded=False):
                        for pop in result.get("Poblaciones", []):
                            st.markdown(f"**{pop.get('Carrera')} · camadas {', '.join(map(str, pop.get('Camadas', [])))}**")
                            st.caption(
                                f"{pop.get('Sede')} · {pop.get('Variante')} · "
                                f"{len(pop.get('Códigos', []))} códigos del trayecto · "
                                f"{len(pop.get('Códigos_ofertados', []))} con oferta en {period.label}"
                            )
                            st.markdown("**Con oferta en el período**")
                            st.code(", ".join(pop.get("Códigos_ofertados", [])) or "Ninguno", language=None)
                            if pop.get("Códigos_sin_oferta"):
                                st.markdown("**Sin oferta en este período**")
                                st.code(", ".join(pop.get("Códigos_sin_oferta", [])), language=None)
                            if pop.get("Fuentes_académicas"):
                                st.caption("Fuentes académicas: " + " + ".join(pop.get("Fuentes_académicas", [])))
                    if not ranking:
                        st.warning("No se encontraron alternativas con los filtros seleccionados.")
                    else:
                        st.subheader("Mejores alternativas")
                        for i, row in enumerate(ranking, start=1):
                            ranking_card(i, row)
