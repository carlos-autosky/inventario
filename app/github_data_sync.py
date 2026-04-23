"""
Sync de datos entre AutoSky y un repo privado de GitHub (disco externo).
Usado para persistir archivos de negocio que deben sobrevivir redeploys de
Streamlit Cloud (filesystem efímero) y estar disponibles entre usuarios.

Archivos sincronizados (constante SYNC_FILES). Config vía st.secrets o env.
"""
import os
import base64
import requests
from typing import Optional

# ── Config default (overridable por st.secrets o env vars) ────────────────
DEFAULT_REPO = "carlos-autosky/autosky-inventario-data"
DEFAULT_BRANCH = "main"
API_BASE = "https://api.github.com"

# Archivos que se sincronizan con el repo-datos. Key = path remoto = nombre local.
SYNC_FILES = [
    "consolidado.xlsx",
    "ventas_consolidado.xlsx",
    "clasificacion_bodegas.json",
    "filtros_config.json",
    "importaciones.xlsx",
    "toma_fisica_rapida.xlsx",
    "ubicaciones_custom.json",
]


class SyncConfig:
    """Config de conexión al repo-datos. Lee de st.secrets o env vars."""

    def __init__(self):
        sec = {}
        try:
            import streamlit as st
            # st.secrets es un Mapping-like; get falla si no existe la key en Cloud
            sec = dict(st.secrets) if hasattr(st, "secrets") else {}
        except Exception:
            sec = {}

        def _first(*keys):
            for k in keys:
                v = sec.get(k) or os.getenv(k)
                if v:
                    return v
            return None

        self.token = _first("GITHUB_GIST_TOKEN", "GITHUB_TOKEN", "GH_TOKEN")
        self.repo = _first("GITHUB_DATA_REPO") or DEFAULT_REPO
        self.branch = _first("GITHUB_DATA_BRANCH") or DEFAULT_BRANCH

    def is_configured(self) -> bool:
        return bool(self.token and self.repo and "/" in self.repo)

    def headers(self) -> dict:
        return {
            "Authorization": f"Bearer {self.token}",
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
        }

    def status_summary(self) -> dict:
        return {
            "configured": self.is_configured(),
            "repo": self.repo,
            "branch": self.branch,
            "has_token": bool(self.token),
        }


# ── Operaciones sobre el repo-datos ────────────────────────────────────────

def list_remote_files(cfg: SyncConfig) -> list[dict]:
    """Lista archivos en el repo-datos (root solamente)."""
    if not cfg.is_configured():
        return []
    url = f"{API_BASE}/repos/{cfg.repo}/contents/?ref={cfg.branch}"
    try:
        r = requests.get(url, headers=cfg.headers(), timeout=30)
    except requests.RequestException:
        return []
    if r.status_code != 200:
        return []
    try:
        return [item for item in r.json() if item.get("type") == "file"]
    except Exception:
        return []


def get_remote_file_meta(cfg: SyncConfig, filename: str) -> Optional[dict]:
    """Obtiene metadata (sha, size, etc.) de un archivo remoto. None si no existe."""
    if not cfg.is_configured():
        return None
    url = f"{API_BASE}/repos/{cfg.repo}/contents/{filename}?ref={cfg.branch}"
    try:
        r = requests.get(url, headers=cfg.headers(), timeout=30)
    except requests.RequestException:
        return None
    if r.status_code == 200:
        try:
            return r.json()
        except Exception:
            return None
    return None


def pull_file(cfg: SyncConfig, filename: str, local_path: str) -> tuple[bool, str]:
    """Descarga un archivo del repo al filesystem local."""
    if not cfg.is_configured():
        return False, "Sync no configurado (falta token o repo)"
    url = f"{API_BASE}/repos/{cfg.repo}/contents/{filename}?ref={cfg.branch}"
    try:
        r = requests.get(url, headers=cfg.headers(), timeout=60)
    except requests.RequestException as ex:
        return False, f"Error de red: {ex}"
    if r.status_code == 404:
        return False, f"No existe en el repo: {filename}"
    if r.status_code != 200:
        return False, f"HTTP {r.status_code}: {r.text[:200]}"
    try:
        data = r.json()
        content = base64.b64decode(data.get("content", ""))
    except Exception as ex:
        return False, f"Error decodificando: {ex}"
    try:
        d = os.path.dirname(local_path)
        if d:
            os.makedirs(d, exist_ok=True)
        with open(local_path, "wb") as f:
            f.write(content)
        return True, f"Descargado ({len(content):,} bytes)"
    except Exception as ex:
        return False, f"Error escribiendo: {ex}"


def push_file(
    cfg: SyncConfig,
    filename: str,
    local_path: str,
    commit_msg: Optional[str] = None,
) -> tuple[bool, str]:
    """Sube (crea o actualiza) un archivo local al repo-datos."""
    if not cfg.is_configured():
        return False, "Sync no configurado"
    if not os.path.exists(local_path):
        return False, f"Archivo local no existe: {local_path}"
    try:
        with open(local_path, "rb") as f:
            content = f.read()
    except Exception as ex:
        return False, f"Error leyendo: {ex}"
    b64 = base64.b64encode(content).decode("ascii")
    url = f"{API_BASE}/repos/{cfg.repo}/contents/{filename}"
    # Intentar obtener el SHA existente para actualizar
    sha = None
    try:
        r_get = requests.get(f"{url}?ref={cfg.branch}", headers=cfg.headers(), timeout=30)
        if r_get.status_code == 200:
            sha = r_get.json().get("sha")
    except requests.RequestException:
        pass
    body = {
        "message": commit_msg or f"Update {filename} from AutoSky",
        "content": b64,
        "branch": cfg.branch,
    }
    if sha:
        body["sha"] = sha
    try:
        r = requests.put(url, headers=cfg.headers(), json=body, timeout=120)
    except requests.RequestException as ex:
        return False, f"Error de red: {ex}"
    if r.status_code in (200, 201):
        return True, f"Subido ({len(content):,} bytes)"
    return False, f"HTTP {r.status_code}: {r.text[:200]}"


def autopull_missing_files(
    cfg: SyncConfig,
    base_dir: str,
    files: Optional[list] = None,
) -> dict:
    """Descarga del repo los archivos que no existen localmente.
    Pensado para arranque de Streamlit Cloud (filesystem efímero).

    Returns:
        {
            "pulled": [str, ...],          # archivos descargados con éxito
            "skipped": [str, ...],         # ya existían local, no se tocan
            "failed": [(str, str), ...],   # error en el pull
            "not_in_remote": [str, ...],   # no existen en el repo
        }
    """
    result = {"pulled": [], "skipped": [], "failed": [], "not_in_remote": []}
    if not cfg.is_configured():
        return result
    files = files if files is not None else SYNC_FILES
    try:
        remote_names = {r["name"] for r in list_remote_files(cfg)}
    except Exception:
        remote_names = set()
    for fn in files:
        local_path = os.path.join(base_dir, fn)
        if os.path.exists(local_path):
            result["skipped"].append(fn)
            continue
        if fn not in remote_names:
            result["not_in_remote"].append(fn)
            continue
        ok, msg = pull_file(cfg, fn, local_path)
        if ok:
            result["pulled"].append(fn)
        else:
            result["failed"].append((fn, msg))
    return result


def delete_file(
    cfg: SyncConfig,
    filename: str,
    commit_msg: Optional[str] = None,
) -> tuple[bool, str]:
    """Borra un archivo del repo-datos."""
    if not cfg.is_configured():
        return False, "Sync no configurado"
    url = f"{API_BASE}/repos/{cfg.repo}/contents/{filename}"
    try:
        r_get = requests.get(f"{url}?ref={cfg.branch}", headers=cfg.headers(), timeout=30)
    except requests.RequestException as ex:
        return False, f"Error de red: {ex}"
    if r_get.status_code == 404:
        return False, f"No existe en el repo: {filename}"
    if r_get.status_code != 200:
        return False, f"HTTP {r_get.status_code} obteniendo SHA"
    sha = r_get.json().get("sha")
    body = {
        "message": commit_msg or f"Delete {filename} from AutoSky",
        "sha": sha,
        "branch": cfg.branch,
    }
    try:
        r = requests.delete(url, headers=cfg.headers(), json=body, timeout=60)
    except requests.RequestException as ex:
        return False, f"Error de red: {ex}"
    if r.status_code == 200:
        return True, "Borrado del repo"
    return False, f"HTTP {r.status_code}: {r.text[:200]}"
