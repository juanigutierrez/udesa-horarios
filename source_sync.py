from __future__ import annotations

import base64
import datetime as dt
import hashlib
import json
import os
import tempfile
import urllib.error
import urllib.request
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping, Optional, Tuple


@dataclass
class SyncRecord:
    key: str
    label: str
    destino: str
    mode: str
    status: str
    remote_name: str = ""
    remote_modified: str = ""
    local_modified: str = ""
    downloaded: bool = False
    stale: bool = False
    message: str = ""

    def public_dict(self) -> Dict[str, Any]:
        # Deliberadamente no expone URL, token, file ID ni rutas absolutas.
        return {
            "Fuente": self.label,
            "Modo": self.mode,
            "Estado": self.status,
            "Archivo": self.remote_name or Path(self.destino).name,
            "Última actualización": self.remote_modified or self.local_modified,
            "Descargada ahora": "Sí" if self.downloaded else "No",
            "Usando respaldo": "Sí" if self.stale else "No",
            "Detalle": self.message,
        }


class SyncError(RuntimeError):
    pass


def _safe_iso_mtime(path: Path) -> str:
    try:
        return dt.datetime.fromtimestamp(path.stat().st_mtime, dt.timezone.utc).isoformat()
    except Exception:
        return ""


def _atomic_write_bytes(path: Path, data: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(prefix=path.name + ".", suffix=".tmp", dir=str(path.parent))
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(data)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, path)
    finally:
        try:
            if os.path.exists(tmp):
                os.unlink(tmp)
        except Exception:
            pass


def _load_json(path: Path) -> Dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def _post_json(url: str, payload: Mapping[str, Any], timeout: int = 45) -> Dict[str, Any]:
    raw = json.dumps(dict(payload), ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=raw,
        method="POST",
        headers={"Content-Type": "application/json", "User-Agent": "UdeSA-Horarios/1.0"},
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            body = resp.read().decode("utf-8")
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")[:500]
        raise SyncError(f"Bridge HTTP {exc.code}: {detail}") from exc
    except Exception as exc:
        raise SyncError(f"No se pudo conectar con el puente de fuentes: {exc}") from exc
    try:
        parsed = json.loads(body)
    except Exception as exc:
        raise SyncError("El puente respondió un formato inválido.") from exc
    if not parsed.get("ok"):
        raise SyncError(str(parsed.get("error") or "El puente rechazó la solicitud."))
    return parsed


def bridge_credentials_from_streamlit() -> Tuple[str, str]:
    """Lee secretos de Streamlit si existen. Se importa tarde para no requerir Streamlit en tests."""
    try:
        import streamlit as st  # type: ignore
        section = st.secrets.get("source_bridge", {})
        url = str(section.get("url", "")).strip()
        token = str(section.get("token", "")).strip()
        return url, token
    except Exception:
        return "", ""


def bridge_credentials_from_env() -> Tuple[str, str]:
    return os.environ.get("UDESA_BRIDGE_URL", "").strip(), os.environ.get("UDESA_BRIDGE_TOKEN", "").strip()


def bridge_credentials() -> Tuple[str, str]:
    url, token = bridge_credentials_from_env()
    if url and token:
        return url, token
    return bridge_credentials_from_streamlit()


def _sync_map(config: Mapping[str, Any], period_id: str) -> List[Dict[str, str]]:
    sync = dict(config.get("sincronizacion", {}))
    mappings = dict(sync.get("fuentes", {}))
    rows: List[Dict[str, str]] = []
    for key, raw in mappings.items():
        if not isinstance(raw, Mapping):
            continue
        periods = raw.get("periodos")
        if periods and period_id not in periods:
            continue
        rows.append({
            "key": str(key),
            "bridge_key": str(raw.get("bridge_key", key)),
            "destino": str(raw.get("destino", "")),
            "label": str(raw.get("label", key)),
        })
    return rows


def local_fingerprint(base: Path, config_path: Path, period_id: str) -> str:
    config = _load_json(config_path)
    parts = [period_id]
    for row in _sync_map(config, period_id):
        p = base / row["destino"]
        try:
            stat = p.stat()
            parts.append(f"{row['key']}:{stat.st_size}:{stat.st_mtime_ns}")
        except Exception:
            parts.append(f"{row['key']}:missing")
    return hashlib.sha256("|".join(parts).encode()).hexdigest()[:20]


def sync_sources(base: Path, config_path: Path, period_id: str, *, force: bool = False) -> Dict[str, Any]:
    """Sincroniza fuentes vía Apps Script bridge.

    Si no hay secretos configurados, usa archivos locales. Si el bridge falla pero existe una copia local,
    conserva esa copia y marca el estado como respaldo. Solo falla cuando una fuente necesaria no puede
    obtenerse ni existe localmente.
    """
    base = Path(base)
    config_path = Path(config_path)
    config = _load_json(config_path)
    sync_cfg = dict(config.get("sincronizacion", {}))
    enabled = bool(sync_cfg.get("habilitada", True))
    rows = _sync_map(config, period_id)
    url, token = bridge_credentials()
    bridge_ready = enabled and bool(url and token)
    records: List[SyncRecord] = []
    fingerprint_parts = [period_id]

    if bridge_ready:
        try:
            meta = _post_json(url, {"action": "metadata_all", "token": token})
            remote_map = {str(x.get("key")): x for x in meta.get("sources", [])}
        except Exception as exc:
            remote_map = {}
            bridge_error = str(exc)
        else:
            bridge_error = ""
    else:
        remote_map = {}
        bridge_error = "Puente no configurado; se usan snapshots locales." if enabled else "Sincronización remota deshabilitada."

    missing: List[str] = []
    for row in rows:
        dest = base / row["destino"]
        remote = remote_map.get(row["bridge_key"], {})
        local_exists = dest.is_file()
        local_mtime = _safe_iso_mtime(dest) if local_exists else ""
        rec = SyncRecord(
            key=row["key"], label=row["label"], destino=row["destino"],
            mode="Bridge Apps Script" if bridge_ready else "Local",
            status="OK" if local_exists else "FALTA",
            local_modified=local_mtime,
        )

        if bridge_ready and remote:
            rec.remote_name = str(remote.get("name", ""))
            rec.remote_modified = str(remote.get("modified", ""))
            rec.message = "Fuente remota disponible."
            fingerprint_parts.append(f"{row['key']}:{rec.remote_modified}:{remote.get('size','')}")
            marker = dest.with_suffix(dest.suffix + ".sync.json")
            previous = {}
            if marker.is_file():
                try:
                    previous = json.loads(marker.read_text(encoding="utf-8"))
                except Exception:
                    previous = {}
            needs_download = force or not local_exists or previous.get("modified") != rec.remote_modified
            if needs_download:
                try:
                    payload = _post_json(url, {"action": "download", "token": token, "source": row["bridge_key"]}, timeout=90)
                    content = base64.b64decode(payload["content_base64"])
                    _atomic_write_bytes(dest, content)
                    marker.write_text(json.dumps({
                        "modified": payload.get("modified", rec.remote_modified),
                        "name": payload.get("name", rec.remote_name),
                        "synced_at": dt.datetime.now(dt.timezone.utc).isoformat(),
                    }, ensure_ascii=False, indent=2), encoding="utf-8")
                    rec.downloaded = True
                    rec.status = "ACTUALIZADA"
                    rec.local_modified = _safe_iso_mtime(dest)
                except Exception as exc:
                    if local_exists:
                        rec.status = "RESPALDO_LOCAL"
                        rec.stale = True
                        rec.message = f"No se pudo actualizar; se conserva la última copia válida. {exc}"
                    else:
                        rec.status = "ERROR"
                        rec.message = str(exc)
                        missing.append(row["label"])
            else:
                rec.status = "AL_DÍA"
        else:
            if local_exists:
                fingerprint_parts.append(f"{row['key']}:local:{dest.stat().st_size}:{dest.stat().st_mtime_ns}")
                if bridge_error:
                    rec.message = bridge_error
                    rec.stale = bridge_ready
                    if bridge_ready:
                        rec.status = "RESPALDO_LOCAL"
            else:
                rec.status = "ERROR"
                rec.message = bridge_error or "Fuente no disponible."
                fingerprint_parts.append(f"{row['key']}:missing")
                missing.append(row["label"])
        records.append(rec)

    if missing:
        raise SyncError("No hay una copia disponible de: " + ", ".join(missing))

    return {
        "mode": "bridge" if bridge_ready else "local",
        "bridge_configured": bridge_ready,
        "bridge_error": bridge_error,
        "records": [asdict(r) for r in records],
        "public_records": [r.public_dict() for r in records],
        "fingerprint": hashlib.sha256("|".join(fingerprint_parts).encode()).hexdigest()[:20],
        "checked_at": dt.datetime.now(dt.timezone.utc).isoformat(),
    }


def public_sync_records(result: Mapping[str, Any]) -> List[Dict[str, Any]]:
    return list(result.get("public_records", []))
