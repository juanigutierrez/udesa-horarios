from __future__ import annotations

import datetime as dt

import streamlit as st

from web_components import page_header, period_caption, population_controls, render_population_errors, render_warnings
from web_logic import best_weekly_rows, clamp_default_window, period_label_map, slot_options, validate_populations, weekly_grid
from web_runtime import get_registry, get_runtime

page_header(
    "Disponibilidad de estudiantes",
    "Consultá una carrera completa o una camada específica. La vista semanal muestra cuántos alumnos están ocupados en cada slot; los filtros académicos quedan en Más filtros.",
)

registry_light = get_registry()
period_map = period_label_map(registry_light)
current_pid = registry_light.active_period_id
period_key = "students_period"
selected_pid = st.session_state.get(period_key, current_pid)
if selected_pid not in registry_light.period_defs:
    selected_pid = current_pid

with st.spinner("Cargando oferta académica..."):
    registry, period, engine, resolver, service, stats = get_runtime(selected_pid)
period_caption(period)

career_options = ["Elegí una carrera"] + service.listar_carreras()
career = st.selectbox("Carrera", options=career_options)
view = st.radio("Vista", options=["Semana", "Fecha puntual"], horizontal=True)
def_date, _ = clamp_default_window(period, days=1)

population = None
with st.expander("＋ Más filtros", expanded=False):
    labels = list(period_map)
    current_label = next((lab for lab, pid in period_map.items() if pid == selected_pid), labels[0])
    selected_label = st.selectbox("Período", options=labels, index=labels.index(current_label), key="students_period_label")
    chosen_pid = period_map[selected_label]
    if chosen_pid != selected_pid:
        st.session_state[period_key] = chosen_pid
        st.rerun()
    st.session_state[period_key] = chosen_pid
    if career != "Elegí una carrera":
        with st.container(border=True):
            population = population_controls(service, career, f"students_{career}")

if career != "Elegí una carrera" and population is None:
    population = {"Carrera": career, "Camadas": "TODAS", "Sede": None, "Variante": None, "Incluir_optativas": False, "Códigos_extra": []}

if view == "Semana":
    reference = st.date_input("Semana de referencia", value=def_date, min_value=period.start, max_value=period.end, format="DD/MM/YYYY", key=f"students_week_{selected_pid}")
    if st.button("Ver semana", type="primary", use_container_width=True):
        if career == "Elegí una carrera":
            st.error("Elegí una carrera.")
        else:
            valid, resolved = validate_populations(service, [population])
            if not valid:
                render_population_errors(resolved)
            else:
                pop = resolved[0]
                monday = reference - dt.timedelta(days=reference.weekday())
                rows = engine.weekly_student_ranking(pop["Códigos"], reference_week_start=monday)
                render_warnings(pop.get("Advertencias", []))
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Camadas", ", ".join(map(str, pop.get("Camadas", []))) or "N/A")
                c2.metric("Códigos del trayecto", len(pop.get("Códigos", [])))
                c3.metric("Con oferta en el período", len(pop.get("Códigos_ofertados", [])))
                c4.metric("Período", period.label)
                with st.expander("Ver códigos consultados", expanded=False):
                    st.markdown("**Con oferta en el período**")
                    st.code(", ".join(pop.get("Códigos_ofertados", [])) or "Ninguno", language=None)
                    if pop.get("Códigos_sin_oferta"):
                        st.markdown("**Sin oferta en este período**")
                        st.code(", ".join(pop.get("Códigos_sin_oferta", [])), language=None)
                    if pop.get("Fuentes_académicas"):
                        st.caption("Fuentes académicas: " + " + ".join(pop.get("Fuentes_académicas", [])))
                st.subheader("Mejores horarios de la semana")
                for row in best_weekly_rows(rows, 10):
                    with st.container(border=True):
                        st.markdown(f"**{row['Día']} · Slot {row['Slot']} · {row['Hora Desde']}–{row['Hora Hasta']}**")
                        cc1, cc2 = st.columns(2)
                        cc1.metric("Alumnos ocupados", row["Alumnos ocupados"])
                        cc2.metric("Materias afectadas", row["Materias afectadas"])
                        if row.get("Códigos afectados"):
                            st.caption("Materias: " + str(row["Códigos afectados"]))
                st.subheader("Grilla semanal")
                st.caption("Cada celda muestra alumnos ocupados. Un número menor implica mayor disponibilidad.")
                st.dataframe(weekly_grid(rows), use_container_width=True, hide_index=True)
else:
    slot_pairs = slot_options(engine.slot_map)
    lookup = dict(slot_pairs)
    c1, c2 = st.columns(2)
    query_date = c1.date_input("Fecha", value=def_date, min_value=period.start, max_value=period.end, format="DD/MM/YYYY", key=f"students_date_{selected_pid}")
    slot_label = c2.selectbox("Horario", options=[label for label, _ in slot_pairs])
    if st.button("Consultar", type="primary", use_container_width=True):
        if career == "Elegí una carrera":
            st.error("Elegí una carrera.")
        else:
            valid, resolved = validate_populations(service, [population])
            if not valid:
                render_population_errors(resolved)
            else:
                pop = resolved[0]
                result = engine.student_occupancy(pop["Códigos"], query_date, slot=lookup[slot_label])
                render_warnings(pop.get("Advertencias", []))
                c1, c2, c3 = st.columns(3)
                c1.metric("Alumnos ocupados", result["Alumnos ocupados"])
                c2.metric("Materias afectadas", result["Materias afectadas"])
                c3.metric("Clases afectadas", result["Clases afectadas"])
                if result.get("Detalle"):
                    table = []
                    for row in result["Detalle"]:
                        table.append({
                            "Código": row.get("Código", ""),
                            "Materia": row.get("Materia", ""),
                            "Horario": f"{row.get('Hora Desde', '')}–{row.get('Hora Hasta', '')}",
                            "Inscriptos": row.get("Cant Inscriptos", ""),
                            "Aula": row.get("Aula", ""),
                        })
                    st.dataframe(table, use_container_width=True, hide_index=True)
