from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, Tuple

import streamlit as st

import udesa_horarios_ronda45 as r45
from source_sync import SyncError, public_sync_records, sync_sources

ROOT = Path(__file__).resolve().parent
CONFIG = ROOT / "fuentes_udesahorarios.json"


def get_registry() -> r45.SourceRegistry:
    return r45.SourceRegistry(ROOT, config_path=CONFIG)


def _period_id(period_id: str) -> str:
    reg = get_registry()
    return reg.active_period_id if period_id in (None, "", "ACTUAL") else str(period_id).upper()


@st.cache_data(ttl=120, show_spinner=False)
def get_sync_state(period_id: str = "ACTUAL", force: bool = False) -> Dict[str, Any]:
    pid = _period_id(period_id)
    return sync_sources(ROOT, CONFIG, pid, force=force)


@st.cache_resource(show_spinner=False)
def _build_runtime_cached(period_id: str, fingerprint: str) -> Tuple[Any, ...]:
    # fingerprint forma parte de la clave: si una fuente cambia, Streamlit reconstruye el motor.
    return r45.build_runtime(ROOT, config_path=CONFIG, period_id=period_id)


def get_runtime(period_id: str = "ACTUAL") -> Tuple[Any, ...]:
    pid = _period_id(period_id)
    state = get_sync_state(pid)
    return _build_runtime_cached(pid, state["fingerprint"])


@st.cache_data(ttl=600, show_spinner=False)
def get_conflicts(period_id: str = "ACTUAL", fingerprint: str = ""):
    pid = _period_id(period_id)
    if not fingerprint:
        fingerprint = get_sync_state(pid)["fingerprint"]
    _, _, engine, _, _, _ = _build_runtime_cached(pid, fingerprint)
    return engine.detect_conflicts()


def get_public_source_status(period_id: str = "ACTUAL"):
    state = get_sync_state(period_id)
    return public_sync_records(state), state


def clear_all_caches() -> None:
    get_sync_state.clear()
    _build_runtime_cached.clear()
    get_conflicts.clear()


def force_sync(period_id: str = "ACTUAL") -> Dict[str, Any]:
    clear_all_caches()
    state = sync_sources(ROOT, CONFIG, _period_id(period_id), force=True)
    # no cacheamos manualmente; el próximo get_runtime tomará el nuevo fingerprint.
    return state


def deployment_readiness() -> Dict[str, Any]:
    """Chequeo seguro para pantalla inicial; no expone secretos."""
    reg = get_registry()
    pid = reg.active_period_id
    try:
        state = get_sync_state(pid)
        return {"ok": True, "period_id": pid, "mode": state.get("mode"), "bridge": state.get("bridge_configured", False)}
    except SyncError as exc:
        return {"ok": False, "period_id": pid, "error": str(exc)}
    except Exception as exc:
        return {"ok": False, "period_id": pid, "error": f"No se pudieron preparar las fuentes: {exc}"}
