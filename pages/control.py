from __future__ import annotations

from collections import Counter

import streamlit as st

from web_components import page_header, period_caption
from web_logic import period_label_map
from web_runtime import clear_all_caches, force_sync, get_conflicts, get_public_source_status, get_registry, get_runtime

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
cols = st.columns(3)
cols[0].metric("Modo de datos", "Automático" if sync_state.get("bridge_configured") else "Local")
cols[1].metric("Fuentes", len(source_rows))
cols[2].metric("Último chequeo", str(sync_state.get("checked_at", ""))[11:19] or "—")

if st.button("Sincronizar ahora", type="secondary"):
    with st.spinner("Consultando fuentes y descargando solo lo que cambió..."):
        state = force_sync(selected_pid)
    st.success("Sincronización completada.")
    st.rerun()

sources_tab, conflicts_tab, reviews_tab = st.tabs(["Fuentes", "Conflictos", "Revisiones académicas"])

with sources_tab:
    st.dataframe(source_rows, use_container_width=True, hide_index=True)
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
