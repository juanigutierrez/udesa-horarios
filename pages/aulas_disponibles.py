from __future__ import annotations

import datetime as dt

import streamlit as st

from web_components import page_header, period_caption, room_card
from web_logic import clamp_default_window, period_label_map, physical_sites, slot_options
from web_runtime import get_registry, get_runtime

page_header(
    "Buscar aulas disponibles",
    "Elegí fecha, horario y capacidad. Por defecto mostramos aulas seguras como LIBRE en Campus; REVISAR y CONFLICTO no se recomiendan.",
)

registry_light = get_registry()
period_map = period_label_map(registry_light)
current_pid = registry_light.active_period_id
period_key = "rooms_period"
selected_pid = st.session_state.get(period_key, current_pid)
if selected_pid not in registry_light.period_defs:
    selected_pid = current_pid

with st.spinner("Cargando datos de aulas..."):
    registry, period, engine, resolver, service, stats = get_runtime(selected_pid)
period_caption(period)

def_date, _ = clamp_default_window(period, days=1)
slot_pairs = slot_options(engine.slot_map)
slot_labels = [label for label, _ in slot_pairs]
slot_lookup = dict(slot_pairs)

c1, c2, c3 = st.columns([1.3, 1.6, 1])
with c1:
    query_date = st.date_input("Fecha", value=def_date, min_value=period.start, max_value=period.end, format="DD/MM/YYYY", key=f"rooms_date_{selected_pid}")
with c2:
    slot_label = st.selectbox("Horario", options=slot_labels, index=min(3, len(slot_labels)-1))
with c3:
    min_capacity = st.number_input("Capacidad mínima", min_value=0, max_value=1000, value=0, step=5, key="rooms_capacity")

room_site = service.default_sede
custom_time = False
start_time = dt.time(13, 0)
end_time = dt.time(14, 30)
include_non_classrooms = False
show_review = False

with st.expander("＋ Más filtros", expanded=False):
    labels = list(period_map)
    current_label = next((lab for lab, pid in period_map.items() if pid == selected_pid), labels[0])
    selected_label = st.selectbox("Período", options=labels, index=labels.index(current_label), key="rooms_period_label")
    chosen_pid = period_map[selected_label]
    if chosen_pid != selected_pid:
        st.session_state[period_key] = chosen_pid
        st.rerun()
    st.session_state[period_key] = chosen_pid
    sites = physical_sites(engine)
    room_site = st.selectbox("Sede", options=sites, index=sites.index(service.default_sede) if service.default_sede in sites else 0)
    custom_time = st.checkbox("Usar horario personalizado", value=False)
    if custom_time:
        cc1, cc2 = st.columns(2)
        start_time = cc1.time_input("Desde", value=dt.time(13, 0), step=300)
        end_time = cc2.time_input("Hasta", value=dt.time(14, 30), step=300)
    include_non_classrooms = st.checkbox("Incluir espacios que no son aulas", value=False)
    show_review = st.checkbox("Mostrar también espacios REVISAR", value=False, help="Se muestran separados y no deben considerarse libres de forma segura.")

if st.button("Buscar aulas", type="primary", use_container_width=True):
    try:
        kwargs = {"start": start_time, "end": end_time} if custom_time else {"slot": slot_lookup[slot_label]}
        if custom_time and start_time >= end_time:
            raise ValueError("La hora de inicio debe ser anterior a la hora de fin.")
        with st.spinner("Comprobando clases, comentarios, reservas y Eventos..."):
            rows = engine.free_rooms(
                query_date,
                sede=room_site,
                min_capacity=int(min_capacity),
                include_non_classrooms=include_non_classrooms,
                include_review=show_review,
                **kwargs,
            )
        safe = [r for r in rows if r.get("Estado") == "LIBRE"]
        review = [r for r in rows if r.get("Estado") == "REVISAR"]
        st.metric("Aulas/espacios libres", len(safe))
        if safe:
            for row in safe:
                room_card(row)
        else:
            st.warning("No se encontraron espacios LIBRE con esos filtros.")
        if review:
            st.divider()
            st.subheader("Requieren revisión")
            st.caption("Estos espacios se bloquean por seguridad; no son una recomendación de disponibilidad.")
            for row in review:
                room_card(row)
    except Exception as exc:
        st.error(str(exc))
