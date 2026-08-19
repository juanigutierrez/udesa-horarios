from __future__ import annotations

import datetime as dt

import streamlit as st

from web_components import page_header, period_caption, schedule_row
from web_logic import clamp_default_window, period_label_map, physical_sites, room_choices, status_icon
from web_runtime import get_registry, get_runtime

page_header(
    "Consultar un aula",
    "Elegí un espacio y una fecha para ver su día completo. Si necesitás un intervalo puntual, activalo en Más filtros.",
)

registry_light = get_registry()
period_map = period_label_map(registry_light)
current_pid = registry_light.active_period_id
period_key = "roomday_period"
selected_pid = st.session_state.get(period_key, current_pid)
if selected_pid not in registry_light.period_defs:
    selected_pid = current_pid

with st.spinner("Cargando catálogo y ocupaciones..."):
    registry, period, engine, resolver, service, stats = get_runtime(selected_pid)
period_caption(period)

sites = physical_sites(engine)
site = service.default_sede
custom_time = False
custom_start = dt.time(12, 15)
custom_end = dt.time(14, 0)
with st.expander("＋ Más filtros", expanded=False):
    labels = list(period_map)
    current_label = next((lab for lab, pid in period_map.items() if pid == selected_pid), labels[0])
    selected_label = st.selectbox("Período", options=labels, index=labels.index(current_label), key="roomday_period_label")
    chosen_pid = period_map[selected_label]
    if chosen_pid != selected_pid:
        st.session_state[period_key] = chosen_pid
        st.session_state.pop("roomday_last_query", None)
        st.rerun()
    st.session_state[period_key] = chosen_pid
    site = st.selectbox("Sede", options=sites, index=sites.index(service.default_sede) if service.default_sede in sites else 0)
    custom_time = st.checkbox(
        "Consultar también un horario personalizado",
        value=False,
        key="roomday_custom_enabled",
        help="El día completo siempre se muestra. Activá esto si además querés verificar un intervalo puntual.",
    )
    if custom_time:
        cc1, cc2 = st.columns(2)
        custom_start = cc1.time_input("Desde", dt.time(12, 15), step=300, key="roomday_custom_start")
        custom_end = cc2.time_input("Hasta", dt.time(14, 0), step=300, key="roomday_custom_end")

choices = room_choices(engine, site=site, include_non_classrooms=True)
choice_map = dict(choices)
labels = [x[0] for x in choices]
def_date, _ = clamp_default_window(period, days=1)

c1, c2 = st.columns([2, 1])
with c1:
    room_label = st.selectbox("Aula o espacio", options=["Elegí un espacio"] + labels, key=f"roomday_room_{selected_pid}_{site}")
with c2:
    query_date = st.date_input("Fecha", value=def_date, min_value=period.start, max_value=period.end, format="DD/MM/YYYY", key=f"roomday_date_{selected_pid}")

if st.button("Ver disponibilidad", type="primary", use_container_width=True):
    if room_label == "Elegí un espacio":
        st.error("Elegí un aula o espacio.")
    elif custom_time and custom_start >= custom_end:
        st.error("La hora de inicio debe ser anterior a la hora de fin.")
    else:
        st.session_state["roomday_last_query"] = {
            "period": selected_pid,
            "room": choice_map[room_label],
            "date": query_date.isoformat(),
            "custom": bool(custom_time),
            "start": custom_start.strftime("%H:%M"),
            "end": custom_end.strftime("%H:%M"),
        }

last = st.session_state.get("roomday_last_query")
if last and last.get("period") == selected_pid:
    room = last["room"]
    date_value = dt.date.fromisoformat(last["date"])
    meta = engine.catalog.get(room, {})
    a, b, c = st.columns(3)
    a.metric("Espacio", room)
    b.metric("Capacidad", meta.get("Capacidad") or "N/A")
    c.metric("Sede", meta.get("Sede") or "N/A")

    rows = engine.room_day(room, date_value)
    status_counts = {}
    for row in rows:
        status_counts[row["Estado"]] = status_counts.get(row["Estado"], 0) + 1
    st.caption(" · ".join(f"{status_icon(k)} {k}: {v}" for k, v in status_counts.items()))
    st.subheader("Día completo")
    for row in rows:
        schedule_row(row)

    if last.get("custom"):
        st.subheader("Horario personalizado")
        result = engine.query_room(room, date_value, start=last["start"], end=last["end"])
        schedule_row({
            "Hora Desde": last["start"],
            "Hora Hasta": last["end"],
            "Estado": result.status,
            "Detalle": " | ".join(
                [f"{g.representative.label}: {g.representative.description}" for g in result.occupation_groups]
                + [f"REVISAR: {x.description}" for x in result.review_claims]
            ),
            "Notas": " | ".join(result.notes),
        })

st.caption("La consulta mostrada permanece visible aunque abras/cierres filtros o cambies otros controles; solo se reemplaza al volver a tocar ‘Ver disponibilidad’.")
