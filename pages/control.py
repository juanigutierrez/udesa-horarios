from __future__ import annotations

import datetime as dt
from collections import Counter
from zoneinfo import ZoneInfo

import streamlit as st

from web_components import page_header, period_caption
from web_logic import period_label_map
from web_runtime import clear_all_caches, force_sync, get_conflicts, get_public_source_status, get_registry, get_runtime

ARGENTINA_TZ = ZoneInfo("America/Argentina/Buenos_Aires")


def format_datetime_argentina(value, include_seconds=False):
    if not value:
        return "—"

    try:
        text = str(value).strip().replace("Z", "+00:00")
        parsed = dt.datetime.fromisoformat(text)

        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=dt.timezone.utc)

        local = parsed.astimezone(ARGENTINA_TZ)

        if include_seconds:
            return local.strftime("%d/%m/%Y %H:%M:%S")

        return local.strftime("%d/%m/%Y %H:%M")
    except Exception:
        return str(value)

page_header(
    "Control de datos",
    "Estado de sincronización, fuentes activas, conflictos y registros académicos que requieren revisión.",
)

registry_light = get_registry()
period_map = period_label_map(registry_light)
current_pid = registry_light.active_period_id
period_ids = list(registry_light.list_period_ids())
selected_pid = st.selectbox(
    "Período a controlar",
    options=period_ids,
    index=period_ids.index(current_pid),
    format_func=lambda pid: next((label for label, value in period_map.items() if value == pid), pid),
)

with st.spinner("Cargando fuentes..."):
    registry, period, engine, resolver, service, stats = get_runtime(selected_pid)
period_caption(period)

source_rows, sync_state = get_public_source_status(selected_pid)

# Presentación de fechas en horario de Argentina.
source_rows_display = []
for row in source_rows:
    display_row = dict(row)
    display_row["Última actualización"] = format_datetime_argentina(
        display_row.get("Última actualización")
    )
    source_rows_display.append(display_row)

cols = st.columns(3)
cols[0].metric("Modo de datos", "Automático" if sync_state.get("bridge_configured") else "Local")
cols[1].metric("Fuentes", len(source_rows))
cols[2].metric(
    "Último chequeo",
    format_datetime_argentina(
        sync_state.get("checked_at"),
        include_seconds=True,
    ),
)

if st.button("Sincronizar ahora", type="secondary"):
    with st.spinner("Consultando fuentes y descargando solo lo que cambió..."):
        state = force_sync(selected_pid)
    st.success("Sincronización completada.")
    st.rerun()

sources_tab, conflicts_tab, reviews_tab = st.tabs(["Fuentes", "Conflictos", "Revisiones académicas"])

with sources_tab:
    st.dataframe(source_rows_display, use_container_width=True, hide_index=True)
    st.caption("Fechas y horarios mostrados en hora de Argentina.")
    st.caption("Solo se muestran nombres y estado. La interfaz nunca expone URLs, IDs de Drive, tokens, rutas locales ni credenciales.")

with conflicts_tab:
    st.caption("El detector no resuelve conflictos silenciosamente. Calculalos solo cuando necesites auditar superposiciones.")
    if st.button("Calcular conflictos", key="load_conflicts"):
        with st.spinner("Cruzando ocupaciones..."):
            rows = get_conflicts(selected_pid, sync_state.get("fingerprint", ""))
        st.metric("Conflictos detectados", len(rows))
        if rows:
            st.dataframe(rows, use_container_width=True, hide_index=True)

with reviews_tab:
    rows = [r for r in resolver.operational_rows if r.get("Estado_periodo") not in {"OFERTADO_CODIGO_PLAN", "RESUELTO_OFERTA_UNICA_EN_PERIODO", "RESUELTO_NOMBRE_EXACTO_EN_PERIODO"}]
    counts = Counter(r.get("Estado_periodo", "") for r in rows)
    cols = st.columns(min(4, max(1, len(counts))))
    for i, (state, count) in enumerate(counts.most_common(4)):
        cols[i].metric(state or "Sin estado", count)
    display = []
    for r in rows:
        display.append({
            "Carrera": r.get("Carrera", ""),
            "Camada": r.get("Camada", ""),
            "Sede": r.get("Sede", ""),
            "Variante": r.get("Variante_mostrar", "") or "Regular",
            "Materia": r.get("Materia", ""),
            "Estado": r.get("Estado_periodo", ""),
            "Detalle": r.get("Detalle_resolución", ""),
        })
    st.dataframe(display, use_container_width=True, hide_index=True)
