from __future__ import annotations

import streamlit as st

from web_runtime import deployment_readiness

st.set_page_config(
    page_title="UdeSA Horarios",
    page_icon="🗓️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
      .block-container {max-width: 1180px; padding-top: 1.4rem; padding-bottom: 4rem;}
      .page-subtitle {color:#5f6b76; font-size:1.02rem; margin-top:-0.55rem; margin-bottom:1.2rem;}
      .app-brand {font-weight:800; font-size:1.04rem; color:#14324a; letter-spacing:.01em;}
      .app-meta {font-size:.82rem; color:#6b7782;}
      [data-testid="stMetricValue"] {font-size:1.55rem;}
      div[data-testid="stExpander"] {border-radius:12px;}
      div[data-testid="stVerticalBlockBorderWrapper"] {border-radius:14px;}
      .footer {margin-top:2.5rem; padding-top:1rem; border-top:1px solid #e5e9ed; color:#707b85; font-size:.82rem;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown("<div class='app-brand'>UdeSA Horarios</div>", unsafe_allow_html=True)
st.markdown("<div class='app-meta'>Consulta integrada de aulas, clases y disponibilidad estudiantil</div>", unsafe_allow_html=True)

ready = deployment_readiness()
if not ready.get("ok"):
    st.error("UdeSA Horarios todavía no puede acceder a sus fuentes de datos.")
    st.info(ready.get("error", "Configurá el puente de fuentes en Streamlit Secrets o ejecutá localmente con snapshots."))
    st.stop()

pages = {
    "": [
        st.Page("pages/mejor_horario.py", title="Mejor horario", default=True),
        st.Page("pages/aulas_disponibles.py", title="Aulas disponibles"),
        st.Page("pages/consultar_aula.py", title="Consultar aula"),
        st.Page("pages/estudiantes.py", title="Estudiantes"),
    ],
    "Sistema": [
        st.Page("pages/control.py", title="Control de datos"),
    ],
}

pg = st.navigation(pages, position="top")
pg.run()

st.markdown(
    "<div class='footer'>UdeSA Horarios · Desarrollado por Juan Ignacio Gutiérrez Julián</div>",
    unsafe_allow_html=True,
)
