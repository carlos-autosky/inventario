"""
Sistema de Inventario v4.3.3 — Interfaz Web (Streamlit)
"""
# ── Performance instrumentation (lo primero, para medir TODO el rerun) ──
import time as _ptime
_PERF_T0 = _ptime.perf_counter()
_PERF_CHECKPOINTS = []
def _perf(label):
    _PERF_CHECKPOINTS.append((label, _ptime.perf_counter() - _PERF_T0))
_perf("script_start")

import streamlit as st
import streamlit.components.v1 as _components
import tempfile, io, os, sys
import pandas as pd
from datetime import datetime, date, timedelta
from collections import defaultdict

sys.path.insert(0, os.path.dirname(__file__))
from app.engine import InventoryEngine
from app.config import PRIMARY_WAREHOUSE
from app.toma_fisica_module import DEFAULT_LOCATIONS
_perf("imports_done")

# Fragment decorator: aisla reruns al bloque decorado (Streamlit 1.37+ estable,
# 1.33-1.36 experimental, anterior: sin efecto — funciona como rerun global)
try:
    _fragment = st.fragment
except AttributeError:
    try:
        _fragment = st.experimental_fragment
    except AttributeError:
        def _fragment(fn): return fn

def _rerun_frag():
    """Rerun sólo del fragment actual (Streamlit 1.37+). Fallback: rerun global."""
    try:
        st.rerun(scope="fragment")
    except Exception:
        st.rerun()

APP_VERSION = "v4.3.3"
BUILD_TIME  = "21/04/2026 GMT-5"

# ── Diagnóstico de inicio (log) ──────────────────────────────
import logging as _logging
_log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "inicio.log")
try:
    with open(_log_path, "a", encoding="utf-8") as _lf:
        _lf.write(f"[STREAMLIT] app_web.py cargado correctamente\n")
        _lf.write(f"[STREAMLIT] __file__ = {os.path.abspath(__file__)}\n")
        _lf.write(f"[STREAMLIT] APP_VERSION = {APP_VERSION}\n")
        _lf.write(f"[STREAMLIT] BUILD_TIME  = {BUILD_TIME}\n")
except Exception: pass

# Forzar recarga: limpiar estado de sesión si la versión cambió
if st.session_state.get("_app_version") != "v4.3.3":
    st.session_state.clear()
    st.session_state["_app_version"] = "v4.3.3"

st.set_page_config(page_title="Inventario v4.3.3", page_icon="📦",
                   layout="wide", initial_sidebar_state="expanded")

# ── Estado compartido multi-sesión ──────────────────────────────
# El engine y la lista de archivos son UN SOLO objeto compartido por
# todas las sesiones (cache_resource = singleton). Así, cuando un cliente
# sube archivos, los demás ven los datos sin tener que recargarlos.
# Persistimos en disco (consolidado.xlsx / toma_fisica.xlsx) para que
# los datos sobrevivan reinicios del servidor.
import threading
_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONSOLIDADO_PATH = os.path.join(_BASE_DIR, "consolidado.xlsx")
TOMA_FISICA_PATH = os.path.join(_BASE_DIR, "toma_fisica.xlsx")
RAPIDA_PATH      = os.path.join(_BASE_DIR, "toma_fisica_rapida.xlsx")
_PERF_LOG_PATH   = os.path.join(_BASE_DIR, "perf.log")
_SHARED_WRITE_LOCK = threading.Lock()

@st.cache_resource
def _get_perf_history():
    return {"runs": []}

def _perf_flush():
    if not _PERF_CHECKPOINTS: return
    try:
        ts = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        total_ms = _PERF_CHECKPOINTS[-1][1] * 1000
        ck_list = []
        prev = 0.0
        parts = []
        for label, t in _PERF_CHECKPOINTS:
            dt_ms = (t - prev) * 1000
            ck_list.append((label, round(dt_ms)))
            parts.append(f"{label}={dt_ms:.0f}")
            prev = t
        with open(_PERF_LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"{ts} TOTAL={total_ms:.0f}ms | {' '.join(parts)}\n")
        hist = _get_perf_history()
        hist["runs"].append({"ts": ts, "total_ms": round(total_ms), "ck": ck_list})
        hist["runs"] = hist["runs"][-30:]
    except Exception: pass
_RAPIDA_COLS = ["Fecha","Ubicación","Código Producto","Nombre Producto",
                "Cantidad Física","Observación"]

@st.cache_resource
def _get_shared_engine():
    e = InventoryEngine()
    if os.path.exists(CONSOLIDADO_PATH):
        try: e.load_inventory_file(CONSOLIDADO_PATH)
        except Exception: pass
    if os.path.exists(TOMA_FISICA_PATH):
        try: e.load_physical_count(TOMA_FISICA_PATH)
        except Exception: pass
    return e

@st.cache_resource
def _get_shared_files():
    state = {"files_loaded": [], "files_stats": []}
    if os.path.exists(CONSOLIDADO_PATH):
        state["files_loaded"].append("consolidado.xlsx (persistido)")
    return state

def _persist_raw(df):
    try:
        with _SHARED_WRITE_LOCK:
            df.to_excel(CONSOLIDADO_PATH, index=False, engine="openpyxl")
    except Exception as ex:
        log(f"⚠ No se pudo persistir consolidado: {ex}")

def _persist_physical(df):
    try:
        with _SHARED_WRITE_LOCK:
            df.to_excel(TOMA_FISICA_PATH, index=False, engine="openpyxl")
    except Exception as ex:
        log(f"⚠ No se pudo persistir toma física: {ex}")

@st.cache_resource
def _get_shared_rapid():
    state = {"df": pd.DataFrame(columns=_RAPIDA_COLS)}
    if os.path.exists(RAPIDA_PATH):
        try:
            df = pd.read_excel(RAPIDA_PATH)
            # Compat: versión previa usaba "Bodega" en lugar de "Ubicación"
            if "Bodega" in df.columns and "Ubicación" not in df.columns:
                df = df.rename(columns={"Bodega":"Ubicación"})
            for c in _RAPIDA_COLS:
                if c not in df.columns: df[c] = ""
            state["df"] = df[_RAPIDA_COLS]
        except Exception:
            pass
    return state

def _persist_rapid(df):
    try:
        with _SHARED_WRITE_LOCK:
            df.to_excel(RAPIDA_PATH, index=False, engine="openpyxl")
    except Exception as ex:
        log(f"⚠ No se pudo persistir toma rápida: {ex}")

# Ubicaciones personalizadas — archivo JSON global compartido entre sesiones
UBIC_CUSTOM_PATH = os.path.join(_BASE_DIR, "ubicaciones_custom.json")

@st.cache_resource
def _get_custom_ubic():
    state = {"list": []}
    if os.path.exists(UBIC_CUSTOM_PATH):
        try:
            import json
            with open(UBIC_CUSTOM_PATH, encoding="utf-8") as f:
                state["list"] = json.load(f).get("ubicaciones", [])
        except Exception: pass
    return state

def _persist_custom_ubic(lst):
    try:
        import json
        with _SHARED_WRITE_LOCK:
            with open(UBIC_CUSTOM_PATH, "w", encoding="utf-8") as f:
                json.dump({"ubicaciones": list(lst)}, f, ensure_ascii=False, indent=2)
    except Exception as ex:
        log(f"⚠ No se pudo persistir ubicaciones: {ex}")

def _get_all_ubic():
    """Devuelve DEFAULT_LOCATIONS + custom, preservando orden y sin duplicados."""
    custom = _get_custom_ubic()["list"]
    seen = set(); result = []
    for u in list(DEFAULT_LOCATIONS) + list(custom):
        if u not in seen:
            result.append(u); seen.add(u)
    return result

VENTANA_OPTS = {
    "Últimos 30 días":  30,
    "Últimos 90 días":  90,
    "Últimos 180 días": 180,
    "Último año":       365,
    "Todo el período":  None,
    "Personalizado":    "custom",
}

def _compute_window_sales(r, ventana_label, custom_start=None, custom_end=None):
    """Calcula ventas por SKU y días de la ventana elegida.

    El consumo diario debe basarse en el ritmo RECIENTE (últimos N días),
    no en todo el histórico: si creciste 3x, las ventas antiguas diluyen
    el promedio y subestiman las compras.

    Devuelve: (days, ventas_by_sku, win_min, win_max)
    """
    df_f = r.filtered
    _data_max = df_f["Fecha"].max()
    _data_min = df_f["Fecha"].min()

    if ventana_label == "Personalizado":
        win_min = pd.Timestamp(custom_start) if custom_start else _data_min
        win_max = pd.Timestamp(custom_end)   if custom_end   else _data_max
    else:
        n_days = VENTANA_OPTS.get(ventana_label)
        if n_days is None:
            win_min = _data_min
            win_max = _data_max
        else:
            win_min = _data_max - pd.Timedelta(days=n_days)
            win_max = _data_max

    # Clamp dentro del rango de datos
    if win_min < _data_min: win_min = _data_min
    if win_max > _data_max: win_max = _data_max
    if win_min > win_max:   win_min = win_max

    df_win = df_f[(df_f["Fecha"] >= win_min) & (df_f["Fecha"] <= win_max)]
    typ = df_win["Tipo"].fillna("").astype(str).str.upper()
    ref = df_win["Referencia"].fillna("").astype(str).str.upper()
    mask = (typ == "EGR") & ref.str.startswith("FAC")
    ventas_win = df_win[mask]
    if ventas_win.empty:
        ventas_by_sku = {}
    else:
        ventas_by_sku = (ventas_win.groupby("Código Producto")["Cantidad"]
                         .sum().abs().to_dict())

    days = max(int((win_max - win_min).days) + 1, 1)
    return days, ventas_by_sku, win_min, win_max

def _ubic_sheet_name(ubic):
    """Reproduce la transformación que hace la plantilla al convertir
    un nombre de ubicación en nombre de hoja Excel (max 28 chars,
    ciertos caracteres reemplazados)."""
    s = str(ubic)[:28]
    for ch, rp in [("/","-"),("\\","-"),("?",""),("*",""),
                    ("[",""),("]",""),(":","")]:
        s = s.replace(ch, rp)
    return s

def _detect_historial_format(path):
    """True si el Excel es un historial exportado: 1 hoja con encabezados
    Fecha, Ubicación, Código Producto, Nombre Producto, Cantidad Física."""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        if len(wb.sheetnames) != 1:
            return False
        ws = wb[wb.sheetnames[0]]
        first_row = next(ws.iter_rows(values_only=True), None)
        if not first_row:
            return False
        expected = {"Fecha","Ubicación","Código Producto",
                    "Nombre Producto","Cantidad Física"}
        headers = {str(v).strip() for v in first_row if v is not None}
        return expected.issubset(headers)
    except Exception:
        return False

def _parse_historial(path):
    """Parsea el Excel de historial y devuelve filas listas para rap_df.
    Formato: 1 hoja con columnas Fecha, Ubicación, Código, Nombre, Cantidad, Obs."""
    df = pd.read_excel(path)
    req = ["Fecha","Ubicación","Código Producto","Nombre Producto","Cantidad Física"]
    if not all(c in df.columns for c in req):
        return [], set()
    rows = []
    ubicaciones = set()
    for _, r in df.iterrows():
        try:
            qty = float(r["Cantidad Física"])
        except (TypeError, ValueError):
            continue
        if qty <= 0:
            continue
        ubic = str(r.get("Ubicación") or "").strip()
        if not ubic:
            continue
        ubicaciones.add(ubic)
        fecha = r.get("Fecha")
        if pd.isna(fecha) or fecha is None or str(fecha).strip() == "":
            fecha_str = datetime.now().strftime("%Y-%m-%d %H:%M")
        elif isinstance(fecha, (datetime, date)):
            fecha_str = fecha.strftime("%Y-%m-%d %H:%M") if isinstance(fecha, datetime) else fecha.strftime("%Y-%m-%d")
        else:
            fecha_str = str(fecha)
        _obs_val = r.get("Observación","")
        _obs = str(_obs_val).strip() if pd.notna(_obs_val) else ""
        rows.append({
            "Fecha": fecha_str,
            "Ubicación": ubic,
            "Código Producto": str(r["Código Producto"] or "").strip(),
            "Nombre Producto": str(r["Nombre Producto"] or "").strip(),
            "Cantidad Física": qty,
            "Observación": _obs,
        })
    return rows, ubicaciones

def _parse_plantilla_toma(path, registered_ubic):
    """Parsea un Excel con el formato de la plantilla de toma física.

    Por cada hoja (excepto 'RESUMEN GENERAL'):
      - Valida que el nombre corresponda a una ubicación registrada.
      - Lee B2 como fecha de toma (opcional).
      - Lee filas desde la 4: A=Código, B=Nombre, C=Cantidad, D=Observación.
      - Ignora celdas vacías o cantidad ≤ 0 (significan "no contado").

    Devuelve:
      {
        "valid_sheets": {ubicacion: {"fecha": date|None, "rows": [...]}},
        "invalid_sheets": [nombres de hoja que no corresponden a ubicación registrada]
      }
    """
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)

    # Mapa: nombre_hoja (transformado) -> ubicación real
    reverse = {_ubic_sheet_name(u): u for u in registered_ubic}

    result = {"valid_sheets": {}, "invalid_sheets": []}

    for sname in wb.sheetnames:
        if sname.upper() == "RESUMEN GENERAL":
            continue
        if sname not in reverse:
            result["invalid_sheets"].append(sname)
            continue

        ubic = reverse[sname]
        ws = wb[sname]

        # Fecha desde B2
        b2 = ws["B2"].value
        fecha = None
        if b2 is not None and str(b2).strip():
            try:
                if isinstance(b2, datetime):
                    fecha = b2.date()
                elif isinstance(b2, date):
                    fecha = b2
                else:
                    fecha = pd.to_datetime(str(b2), errors="coerce")
                    fecha = fecha.date() if fecha is not pd.NaT and fecha is not None else None
            except Exception:
                fecha = None

        # Filas desde 4
        rows = []
        for r in range(4, (ws.max_row or 0) + 1):
            codigo = ws.cell(r, 1).value
            if codigo is None: continue
            _code = str(codigo).strip()
            if not _code: continue
            if _code.upper() in ("TOTAL", "TOTAL GENERAL"):
                break

            nombre = ws.cell(r, 2).value
            cantidad = ws.cell(r, 3).value
            obs = ws.cell(r, 4).value

            try:
                qty = float(cantidad) if cantidad is not None else 0.0
            except (ValueError, TypeError):
                qty = 0.0

            # Vacío o 0 = "no contado" → skip
            if qty <= 0:
                continue

            rows.append({
                "codigo":   _code,
                "nombre":   str(nombre or "").strip(),
                "cantidad": qty,
                "obs":      str(obs or "").strip(),
            })

        result["valid_sheets"][ubic] = {"fecha": fecha, "rows": rows}

    return result

# ── Session state ───────────────────────────────────────────────
def _init():
    shared_files = _get_shared_files()
    defs = {"engine": _get_shared_engine(), "result": None,
            "files_loaded": shared_files["files_loaded"],
            "files_stats":  shared_files["files_stats"],
            "log": [], "dark_mode": False,
            "excluded_skus": set(), "excl_wh": set()}
    for k,v in defs.items():
        if k not in st.session_state: st.session_state[k] = v
    # Si la sesión arranca con datos pre-cargados (el servidor ya los tenía)
    # pero sin cálculo, disparar auto-análisis para que los KPIs aparezcan
    # sin exigir al usuario pulsar "Calcular"
    if (st.session_state.engine.raw_df is not None
            and st.session_state.result is None):
        st.session_state["_recalc_pending"] = True
_init()
eng = st.session_state.engine
dark = st.session_state.dark_mode
_perf("session_init")

# ── Tema ────────────────────────────────────────────────────────
if dark:
    BG,PANEL,BORDER,TEXT,MUTED = "#0f172a","#1e293b","#334155","#f1f5f9","#94a3b8"
    TH,TDE,TDO,HOVER = "#0f172a","#1e293b","#162032","#243447"
    ACC,SUC,WRN,DNG = "#38bdf8","#4ade80","#fbbf24","#f87171"
else:
    BG,PANEL,BORDER,TEXT,MUTED = "#f8fafc","#ffffff","#e2e8f0","#0f172a","#64748b"
    TH,TDE,TDO,HOVER = "#f1f5f9","#ffffff","#f8fafc","#e0f2fe"
    ACC,SUC,WRN,DNG = "#0284c7","#16a34a","#d97706","#dc2626"

# ── Autosky Design System CSS ─────────────────────────────
_CSS = '''
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap');
:root {
  --sky:#0ea5e9; --sky-d:#0284c7; --sky-dd:#0369a1;
  --sky-l:#e0f2fe; --sky-ll:#f0f9ff;
  --bg:#f0f9ff; --surface:#ffffff; --surface2:#f8fafc;
  --border:#e2e8f0; --border2:#cbd5e1;
  --text:#0f172a; --text2:#475569; --text3:#94a3b8;
  --green:#059669; --green-l:#ecfdf5;
  --red:#dc2626; --red-l:#fef2f2;
  --amber:#d97706; --amber-l:#fffbeb;
  --purple:#7c3aed; --purple-l:#f3e8ff;
  --radius:10px; --radius-sm:6px;
  --shadow:0 1px 3px rgba(0,0,0,.08);
}
html,body,[data-testid='stAppViewContainer']{
  background:var(--bg)!important;
  font-family:'Inter','Segoe UI',system-ui,sans-serif;
  font-size:13px; color:var(--text);
}
.as-banner{
  background:linear-gradient(135deg,#0ea5e9,#38bdf8,#7dd3fc);
  padding:12px 22px; border-radius:10px; margin-bottom:16px;
  display:flex; align-items:center; justify-content:space-between;
  box-shadow:0 2px 8px rgba(14,165,233,.35);
}
.as-logo{font-size:18px;font-weight:800;color:#fff;letter-spacing:.06em;}
.as-logo span{font-weight:300;opacity:.85;}
.as-build{text-align:right;color:rgba(255,255,255,.9);
  font-size:10px;font-family:'JetBrains Mono',monospace;}
.as-build .v{font-size:13px;font-weight:700;}
section[data-testid='stSidebar']{
  background:var(--surface)!important;
  border-right:1px solid var(--border)!important;
}
section[data-testid='stSidebar'] *{color:var(--text)!important;}
section[data-testid='stSidebar'] .stSelectbox label,
section[data-testid='stSidebar'] .stMultiSelect label,
section[data-testid='stSidebar'] .stCheckbox label{
  font-size:10px!important; font-weight:600!important;
  color:var(--text3)!important; text-transform:uppercase; letter-spacing:.06em;
}
.kpi-row{display:flex;gap:10px;margin-bottom:10px;flex-wrap:wrap;}
.kpi-card{
  background:var(--surface); border:1px solid var(--border);
  border-radius:var(--radius); padding:14px 16px;
  flex:1 1 140px; min-width:0;
  position:relative; overflow:hidden; box-shadow:var(--shadow);
}
.kpi-card::before{
  content:''; position:absolute; top:0; left:0; right:0; height:3px;
  background:var(--accent,var(--sky));
}
.kpi-label{font-size:10px;font-weight:600;color:var(--text3);
  text-transform:uppercase;letter-spacing:.06em;margin-bottom:5px;
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.kpi-value{font-size:21px;font-weight:700;color:var(--text);
  font-family:'JetBrains Mono',monospace;
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.kpi-sub{font-size:10px;color:var(--text3);margin-top:3px;
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
@media (max-width: 640px){
  .kpi-row{gap:6px;}
  .kpi-card{flex:1 1 calc(50% - 3px); padding:10px 12px;}
  .kpi-label{font-size:9px;letter-spacing:.02em;margin-bottom:3px;}
  .kpi-value{font-size:17px;}
  .kpi-sub{font-size:9px;}
}
@media (max-width: 420px){
  .kpi-value{font-size:15px;}
}
.kpi-card.a{--accent:var(--sky);}
.kpi-card.s{--accent:var(--green);}
.kpi-card.w{--accent:var(--amber);}
.kpi-card.d{--accent:var(--red);}
.kpi-card.p{--accent:var(--purple);}
.tc{overflow:auto;max-height:540px;border-radius:var(--radius);
  border:1px solid var(--border);background:var(--surface);
  box-shadow:var(--shadow);
  /* CRÍTICO: position relative para que sticky funcione dentro */
  position:relative;}
.tc.piv{max-height:680px;}
.it{width:100%;border-collapse:separate;border-spacing:0;font-size:12px;
  color:var(--text);font-family:'Inter',sans-serif;}

/* ── Header sticky vertical (top:0 relativo al .tc) ── */
.it thead th{
  background:var(--surface2);color:var(--text3);
  font-size:10px;font-weight:600;letter-spacing:.06em;text-transform:uppercase;
  padding:8px 12px;
  position:sticky;top:0;z-index:3;
  border-bottom:2px solid var(--border2);
  white-space:nowrap;cursor:pointer;user-select:none;}

/* ── Col 1 sticky horizontal ── */
.it thead th:first-child,
.it tbody  td:first-child,
.it tfoot  td:first-child{
  position:sticky;left:0;z-index:2;
  background:var(--surface2);
  border-right:1px solid var(--border2);}

/* ── Col 2 sticky horizontal (para pivot con SKU + Nombre) ── */
.it thead th.sc2,
.it tbody  td.sc2,
.it tfoot  td.sc2{
  position:sticky;z-index:2;
  background:var(--surface2);
  border-right:2px solid var(--border2);}

/* ── Esquina (col1+header): z-index máximo ── */
.it thead th:first-child{z-index:5;}
.it thead th.sc2{z-index:5;}

/* ── Zebra ── */
.it tbody tr:nth-child(even) td            {background:var(--surface2);}
.it tbody tr:nth-child(odd)  td            {background:var(--surface);}
.it tbody tr:nth-child(even) td:first-child{background:var(--surface2);}
.it tbody tr:nth-child(odd)  td:first-child{background:var(--surface);}
.it tbody tr:nth-child(even) td.sc2        {background:var(--surface2);}
.it tbody tr:nth-child(odd)  td.sc2        {background:var(--surface);}

/* ── Hover sobre zebra ── */
.it tbody tr:hover td                      {background:var(--sky-ll)!important;}

.it thead th:hover{color:var(--sky);}
.it tbody td{padding:7px 12px;border-bottom:1px solid var(--surface2);color:var(--text);}
.it .n{text-align:right;font-family:'JetBrains Mono',monospace;}
/* Variante compacta para Rotación: filas densas, una por SKU (sin wrap) */
.it.it-rot thead th{padding:4px 6px!important;font-size:9px!important;
  letter-spacing:.04em;white-space:nowrap;
  /* Sticky top reforzado para que no desaparezca al scroll vertical */
  position:sticky!important;top:0!important;
  background:var(--surface2)!important;
  z-index:10!important;}
/* Esquina (Código): sticky top Y left al tiempo, z-index máximo */
.it.it-rot thead th:first-child{
  left:0!important;z-index:12!important;}
/* Nombre: segundo sticky top (sin left), z-index alto */
.it.it-rot thead th:nth-child(2){z-index:10!important;}
.it.it-rot tbody td{padding:2px 6px!important;font-size:10px;line-height:1.2;
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
/* Código (col 1) sticky left, fondo opaco, z-index por debajo del header */
.it.it-rot tbody td:first-child{
  font-family:'JetBrains Mono',monospace;font-size:10px;
  position:sticky!important;left:0!important;z-index:4!important;}
/* Columna Nombre (2da) con ancho flexible y truncado */
.it.it-rot tbody td.nom{max-width:320px;}
.it.it-rot tfoot td{padding:4px 6px!important;font-size:10px;}
.it.it-rot tfoot td:first-child{position:sticky!important;left:0!important;z-index:4!important;}
.tc.rot-tc{max-height:640px;}
.it tfoot tr.tot td{background:var(--sky-ll)!important;font-weight:700;
  border-top:2px solid var(--sky-l);color:var(--sky-d);}
.it tfoot tr.tot td{background:var(--sky-ll)!important;font-weight:700;
  border-top:2px solid var(--sky-l);color:var(--sky-d);}
.zb{display:flex;align-items:center;gap:6px;margin-bottom:6px;}
.zb button{background:var(--surface);border:1px solid var(--border);
  color:var(--text2);border-radius:var(--radius-sm);
  padding:2px 10px;cursor:pointer;font-weight:700;font-size:13px;}
.zb button:hover{background:var(--sky);color:#fff;border-color:var(--sky);}
.zb .zb-info{font-size:11px;color:var(--text3);}
.stTabs [data-baseweb='tab-list']{gap:2px;background:var(--surface2);
  border-radius:var(--radius-sm);padding:4px;border:1px solid var(--border);}
.stTabs [data-baseweb='tab']{border-radius:5px;font-size:12px;
  font-weight:600;color:var(--text2);padding:6px 14px;}
.stTabs [aria-selected='true']{background:var(--sky)!important;color:#fff!important;}
.stDataFrame td,.stDataFrame th{color:var(--text)!important;}
.stButton button[kind='primary']{background:var(--sky)!important;
  border-color:var(--sky)!important;border-radius:var(--radius-sm)!important;
  font-weight:600!important;}
.stButton button[kind='primary']:hover{background:var(--sky-d)!important;}
.stButton button{border-radius:var(--radius-sm)!important;font-weight:600!important;}
.stAlert{border-radius:var(--radius)!important;}
.stSelectbox>div>div,.stMultiSelect>div>div{border-radius:var(--radius-sm)!important;}
.stTextInput>div>div>input{border-radius:var(--radius-sm)!important;}

/* ═══════════════════════════════════════════════
   FORZAR TEMA CLARO EN TODOS LOS COMPONENTES
   ═══════════════════════════════════════════════ */

/* App background completo */
.stApp, .stApp > div,
[data-testid="stAppViewContainer"],
[data-testid="stAppViewBlockContainer"],
[data-testid="block-container"],
.main, .main .block-container {
  background-color: #f0f9ff !important;
  color: #0f172a !important;
}

/* Sidebar fondo blanco */
[data-testid="stSidebar"],
[data-testid="stSidebar"] > div,
[data-testid="stSidebar"] .sidebar-content {
  background-color: #ffffff !important;
}
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] div {
  color: #0f172a !important;
}

/* Todos los textos */
p, span, div, label, h1, h2, h3, h4 {
  color: #0f172a;
}

/* File uploader */
[data-testid="stFileUploader"],
[data-testid="stFileUploadDropzone"] {
  background-color: #ffffff !important;
  border: 2px dashed #bae6fd !important;
  border-radius: 10px !important;
  color: #0f172a !important;
}
[data-testid="stFileUploadDropzone"] * { color: #0f172a !important; }
[data-testid="stFileUploadDropzone"]:hover {
  border-color: #0ea5e9 !important;
  background-color: #f0f9ff !important;
}

/* Upload button dentro del dropzone */
[data-testid="stFileUploader"] button {
  background: #0ea5e9 !important;
  color: #fff !important;
  border-radius: 6px !important;
  border: none !important;
}

/* Uploaded file item */
[data-testid="stFileUploaderFile"],
[data-testid="uploadedFileData"] {
  background: #f0f9ff !important;
  border: 1px solid #bae6fd !important;
  border-radius: 8px !important;
}
[data-testid="stFileUploaderFile"] * { color: #0f172a !important; }

/* Inputs, selects */
.stTextInput input, .stNumberInput input,
.stDateInput input, .stTimeInput input {
  background: #ffffff !important;
  color: #0f172a !important;
  border: 1px solid #e2e8f0 !important;
  border-radius: 6px !important;
}
.stSelectbox > div > div > div,
.stMultiSelect > div > div > div {
  background: #ffffff !important;
  color: #0f172a !important;
}

/* Dropdown options */
[data-baseweb="popover"], [data-baseweb="menu"],
[role="listbox"], [role="option"] {
  background: #ffffff !important;
  color: #0f172a !important;
}
[role="option"]:hover { background: #f0f9ff !important; }

/* Number input */
[data-testid="stNumberInput"] input {
  background: #ffffff !important;
  color: #0f172a !important;
}
[data-testid="stNumberInput"] button {
  background: #f1f5f9 !important;
  color: #0f172a !important;
  border: 1px solid #e2e8f0 !important;
}

/* Checkbox, toggle */
[data-testid="stCheckbox"] label,
[data-testid="stToggle"] label { color: #0f172a !important; }

/* Metrics */
[data-testid="stMetric"] * { color: #0f172a !important; }
[data-testid="stMetricValue"] {
  color: #0f172a !important;
  font-family: 'JetBrains Mono', monospace !important;
}

/* Success/Info/Warning/Error alerts */
.stAlert > div { border-radius: 8px !important; }
[data-testid="stNotification"] { border-radius: 8px !important; }

/* Success boxes (like "1 archivo cargado") */
.element-container .stAlert [data-baseweb="notification"] {
  background: #ecfdf5 !important;
  color: #065f46 !important;
  border: 1px solid #6ee7b7 !important;
}

/* Dividers */
hr { border-color: #e2e8f0 !important; }

/* Spinner */
[data-testid="stSpinner"] { color: #0ea5e9 !important; }

/* Caption text */
.stCaption, [data-testid="stCaptionContainer"] {
  color: #64748b !important;
}

/* Markdown text */
.stMarkdown p, .stMarkdown span { color: #0f172a !important; }

/* Tabs area background */
[data-testid="stHorizontalBlock"],
[data-testid="stVerticalBlock"] {
  background: transparent !important;
}

/* Column backgrounds */
[data-testid="column"] { background: transparent !important; }

/* Expander */
[data-testid="stExpander"] {
  background: #ffffff !important;
  border: 1px solid #e2e8f0 !important;
  border-radius: 10px !important;
}
[data-testid="stExpander"] summary { color: #0f172a !important; }

/* Download button */
.stDownloadButton button {
  background: #f8fafc !important;
  color: #0284c7 !important;
  border: 1px solid #bae6fd !important;
  border-radius: 6px !important;
  font-weight: 600 !important;
}
.stDownloadButton button:hover {
  background: #e0f2fe !important;
  border-color: #0ea5e9 !important;
}
'''
# ── CSS adicional tema oscuro (se inyecta sobre el base claro) ──
_DARK_CSS = '''
html,body,[data-testid='stAppViewContainer'],
.stApp, .stApp > div,
[data-testid="stAppViewContainer"],
[data-testid="stAppViewBlockContainer"],
[data-testid="block-container"],
.main, .main .block-container {
  background-color: #0f172a !important;
  color: #f1f5f9 !important;
}
:root {
  --bg:#0f172a; --surface:#1e293b; --surface2:#162032;
  --border:#334155; --text:#f1f5f9; --text2:#cbd5e1; --text3:#94a3b8;
  --sky-ll:#162032; --sky-l:#1e3a5f;
  --green-l:#052e16; --red-l:#450a0a; --amber-l:#451a03; --purple-l:#2e1065;
}
section[data-testid='stSidebar'],
[data-testid="stSidebar"],
[data-testid="stSidebar"] > div {
  background-color: #1e293b !important;
}
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] div { color: #f1f5f9 !important; }
p, span, div, label, h1, h2, h3, h4 { color: #f1f5f9; }
.stTextInput input, .stNumberInput input,
.stDateInput input, .stTimeInput input {
  background: #1e293b !important; color: #f1f5f9 !important;
  border: 1px solid #334155 !important;
}
.stSelectbox > div > div > div,
.stMultiSelect > div > div > div {
  background: #1e293b !important; color: #f1f5f9 !important;
}
[data-baseweb="popover"], [data-baseweb="menu"],
[role="listbox"], [role="option"] {
  background: #1e293b !important; color: #f1f5f9 !important;
}
[role="option"]:hover { background: #0f172a !important; }
[data-testid="stFileUploader"],
[data-testid="stFileUploadDropzone"] {
  background-color: #1e293b !important;
  border: 2px dashed #334155 !important; color: #f1f5f9 !important;
}
[data-testid="stFileUploadDropzone"] * { color: #f1f5f9 !important; }
[data-testid="stFileUploaderFile"],
[data-testid="uploadedFileData"] {
  background: #162032 !important;
  border: 1px solid #334155 !important;
}
[data-testid="stFileUploaderFile"] * { color: #f1f5f9 !important; }
[data-testid="stFileUploader"] button {
  background: #0ea5e9 !important; color: #fff !important;
}
[data-testid="stExpander"] {
  background: #1e293b !important; border: 1px solid #334155 !important;
}
[data-testid="stExpander"] summary { color: #f1f5f9 !important; }
[data-testid="stMetric"] * { color: #f1f5f9 !important; }
.stMarkdown p, .stMarkdown span { color: #f1f5f9 !important; }
hr { border-color: #334155 !important; }
.stDownloadButton button {
  background: #1e293b !important; color: #38bdf8 !important;
  border: 1px solid #334155 !important;
}
.stDownloadButton button:hover {
  background: #0f172a !important; border-color: #38bdf8 !important;
}
.it tbody td { color: #f1f5f9 !important; }
.it thead th { background: #162032 !important; color: #94a3b8 !important; }
.it tbody tr:hover td { background: #162032 !important; }
.it tfoot tr.tot td { background: #1e3a5f !important; color: #38bdf8 !important; }
'''

_ACTIVE_CSS = _CSS + (_DARK_CSS if dark else "")
st.markdown(f'''<style>{_ACTIVE_CSS}</style>
<script>
// ── Zoom ──────────────────────────────────────────────────────
function asZoom(uid, delta) {{
  var t = document.getElementById("tbl_" + uid);
  if (!t) return;
  var sz = parseFloat(window.getComputedStyle(t).fontSize) || 12;
  t.style.fontSize = Math.min(20, Math.max(8, sz + delta)) + "px";
}}
function asZoomReset(uid) {{
  var t = document.getElementById("tbl_" + uid);
  if (t) t.style.fontSize = "12px";
}}

</script>''', unsafe_allow_html=True)

# ── Resize de columnas via window.parent (accede al DOM real) ───
_components.html("""<script>
(function() {
  var doc = window.parent.document;

  // ── Col resize ───────────────────────────────────────────────
  function initResize(table) {
    table.style.tableLayout = 'auto';
    table.querySelectorAll('thead th').forEach(function(th) {
      if (th.dataset.resizeInit) return;
      th.dataset.resizeInit = '1';
      th.style.position   = 'relative';
      th.style.overflow   = 'visible';
      th.style.userSelect = 'none';
      var handle = doc.createElement('div');
      handle.style.cssText = 'position:absolute;top:0;right:0;bottom:0;width:6px;cursor:col-resize;z-index:10;background:transparent';
      handle.addEventListener('mouseenter', function() { handle.style.background = 'rgba(14,165,233,.45)'; });
      handle.addEventListener('mouseleave', function() { if (!handle._drag) handle.style.background = 'transparent'; });
      var startX, startW, drag = false;
      handle.addEventListener('mousedown', function(e) {
        e.stopPropagation(); e.preventDefault();
        drag = true; handle._drag = true;
        startX = e.pageX; startW = th.offsetWidth;
        handle.style.background   = 'rgba(14,165,233,.7)';
        doc.body.style.cursor     = 'col-resize';
        doc.body.style.userSelect = 'none';
      });
      doc.addEventListener('mousemove', function(e) {
        if (!drag) return;
        var w = Math.max(40, startW + (e.pageX - startX));
        th.style.width = th.style.minWidth = th.style.maxWidth = w + 'px';
      });
      doc.addEventListener('mouseup', function() {
        if (!drag) return;
        drag = false; handle._drag = false;
        handle.style.background   = 'transparent';
        doc.body.style.cursor     = '';
        doc.body.style.userSelect = '';
      });
      th.appendChild(handle);
    });
  }

  // ── Two-finger / trackpad pan en contenedor .tc ──────────────
  // Detecta touchstart con 2 dedos O wheel con deltaX (trackpad horizontal)
  // y los traduce en scrollLeft/scrollTop del contenedor .tc
  function initPan(tc) {
    if (tc.dataset.panInit) return;
    tc.dataset.panInit = '1';

    // Trackpad horizontal (wheel event con deltaX)
    tc.addEventListener('wheel', function(e) {
      // Si hay desplazamiento horizontal real (trackpad/dos dedos) lo aplicamos
      if (Math.abs(e.deltaX) > Math.abs(e.deltaY)) {
        e.preventDefault();
        tc.scrollLeft += e.deltaX;
      }
      // Si hay desplazamiento vertical con shift → scroll horizontal
      else if (e.shiftKey) {
        e.preventDefault();
        tc.scrollLeft += e.deltaY;
      }
      // Vertical normal: dejar comportamiento por defecto (no interceptar)
    }, { passive: false });

    // Touch (móvil / tablet): dos dedos → pan libre
    var t0x, t0y, t0sl, t0st;
    tc.addEventListener('touchstart', function(e) {
      if (e.touches.length !== 2) return;
      t0x  = (e.touches[0].pageX + e.touches[1].pageX) / 2;
      t0y  = (e.touches[0].pageY + e.touches[1].pageY) / 2;
      t0sl = tc.scrollLeft;
      t0st = tc.scrollTop;
    }, { passive: true });

    tc.addEventListener('touchmove', function(e) {
      if (e.touches.length !== 2) return;
      e.preventDefault();
      var cx = (e.touches[0].pageX + e.touches[1].pageX) / 2;
      var cy = (e.touches[0].pageY + e.touches[1].pageY) / 2;
      tc.scrollLeft = t0sl - (cx - t0x);
      tc.scrollTop  = t0st - (cy - t0y);
    }, { passive: false });
  }

  // ── Scan: inicializar resize y pan en todas las tablas/contenedores ──
  function scanTables() {
    doc.querySelectorAll('table.it').forEach(initResize);
    doc.querySelectorAll('div.tc').forEach(initPan);
  }

  var obs = new MutationObserver(function(muts) {
    muts.forEach(function(m) {
      m.addedNodes.forEach(function(n) {
        if (n.nodeType !== 1) return;
        if (n.matches) {
          if (n.matches('table.it')) initResize(n);
          if (n.matches('div.tc'))   initPan(n);
        }
        if (n.querySelectorAll) {
          n.querySelectorAll('table.it').forEach(initResize);
          n.querySelectorAll('div.tc').forEach(initPan);
        }
      });
    });
  });
  obs.observe(doc.body, { childList: true, subtree: true });

  scanTables();
  setTimeout(scanTables, 800);
  setTimeout(scanTables, 2000);
})();
</script>""", height=0, scrolling=False)
_perf("css_js_emitted")

# ── Helpers ─────────────────────────────────────────────────────
def log(m):
    ts = datetime.now().strftime("%H:%M:%S")
    st.session_state.log.insert(0, f"[{ts}] {m}")
    st.session_state.log = st.session_state.log[:300]

def fmt(v, t="n"):
    try:
        f = float(v)
        if t=="i": return f"{int(f):,}"
        if t=="p": return f"{f:.1f}%"
        return f"{f:,.2f}"
    except: return "—"

def kc(label, val, cls="", sub=""):
    sub_html = f'<div class="kpi-sub">{sub}</div>' if sub else ''
    return f'<div class="kpi-card {cls}"><div class="kpi-label">{label}</div><div class="kpi-value">{val}</div>{sub_html}</div>'

def to_xl(df):
    b = io.BytesIO()
    with pd.ExcelWriter(b, engine="openpyxl") as w: df.to_excel(w, index=False)
    return b.getvalue()

def to_html(df, title="Reporte"):
    hdrs = "".join(f"<th style='background:#1e3a5f;color:#fff;padding:6px 10px;text-align:left'>{c}</th>" for c in df.columns)
    rows = ""
    for i,(_,r) in enumerate(df.iterrows()):
        bg="#f9fafb" if i%2==0 else "#fff"
        cells="".join(f"<td style='padding:4px 10px;border-bottom:1px solid #e5e7eb;background:{bg}'>{str(v) if str(v) not in ('nan','None','NaN') else ''}</td>" for v in r)
        rows+=f"<tr>{cells}</tr>"
    return f"""<!DOCTYPE html><html><head><meta charset='UTF-8'><title>{title}</title>
<style>body{{font-family:sans-serif;padding:20px;background:#fff;color:#111}}
h1{{color:#1e3a5f}}table{{border-collapse:collapse;width:100%;font-size:11px}}</style></head>
<body><h1>{title}</h1><p style='color:#6b7280;font-size:11px'>Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
<table><thead><tr>{hdrs}</tr></thead><tbody>{rows}</tbody></table></body></html>""".encode()

def to_pdf(df, title="Reporte"):
    try:
        from reportlab.lib.pagesizes import landscape, A4
        from reportlab.lib import colors
        from reportlab.lib.units import cm
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
        buf=io.BytesIO()
        doc=SimpleDocTemplate(buf,pagesize=landscape(A4),topMargin=1*cm,bottomMargin=1*cm,leftMargin=1.5*cm,rightMargin=1.5*cm)
        PW=landscape(A4)[0]-3*cm; cols=list(df.columns); cws=[PW/len(cols)]*len(cols)
        C_B=colors.HexColor("#1E3A5F"); C_E=colors.HexColor("#F9FAFB")
        data=[cols]+[list(map(lambda v: "" if str(v) in("nan","None","NaN") else str(v), row)) for _,row in df.iterrows()]
        ts=[("BACKGROUND",(0,0),(-1,0),C_B),("TEXTCOLOR",(0,0),(-1,0),colors.white),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),7),
            ("TOPPADDING",(0,0),(-1,-1),2),("BOTTOMPADDING",(0,0),(-1,-1),2),
            ("BOX",(0,0),(-1,-1),0.5,colors.HexColor("#D1D5DB")),
            ("INNERGRID",(0,0),(-1,-1),0.3,colors.HexColor("#E5E7EB"))]
        for i in range(1,len(data)):
            if i%2==0: ts.append(("BACKGROUND",(0,i),(-1,i),C_E))
        t=Table(data,colWidths=cws,repeatRows=1); t.setStyle(TableStyle(ts))
        sty=getSampleStyleSheet()
        doc.build([Paragraph(title,ParagraphStyle("t",fontSize=13,textColor=C_B,fontName="Helvetica-Bold",spaceAfter=6)),t])
        return buf.getvalue()
    except: return None

def dl3(df, name, key):
    c1,c2,c3=st.columns(3)
    with c1: st.download_button("📊 Excel",to_xl(df),f"{name}.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",key=f"{key}_xl",use_container_width=True)
    with c2: st.download_button("🌐 HTML",to_html(df,name),f"{name}.html","text/html",key=f"{key}_htm",use_container_width=True)
    with c3:
        pdf=to_pdf(df,name)
        if pdf: st.download_button("📄 PDF",pdf,f"{name}.pdf","application/pdf",key=f"{key}_pdf",use_container_width=True)
        else: st.button("📄 PDF",disabled=True,key=f"{key}_pdfx",use_container_width=True)

def _strict_opts(options, query, format_func=None, keep_selected=None):
    """Filtra opciones por substring estricto (no fuzzy). Streamlit usa fuzzy
    por subsecuencia en el dropdown interno (ej: 'AKE' matchea 'FRANK KEMPIN').

    Para evitar esto, si no hay query devuelve [] (más keep_selected) — fuerza
    al usuario a escribir en el text_input EXTERNO, cuyo filtrado es estricto.
    Si ya hay query, devuelve los items que realmente contienen ese texto."""
    fx = format_func if format_func else str
    if query:
        ql = query.lower()
        filt = [o for o in options if ql in fx(o).lower()]
    else:
        filt = []  # Sin búsqueda, sin opciones — se obliga a usar el buscador
    if keep_selected:
        for c in keep_selected:
            if c not in filt: filt.append(c)
    return filt

def pfilt(df, txt, cols=("Código Producto","Nombre Producto")):
    if not txt: return df
    s=txt.lower()
    m=None
    for c in cols:
        if c in df.columns:
            mk=df[c].fillna("").astype(str).str.lower().str.contains(s,regex=False)
            m=mk if m is None else m|mk
    return df[m] if m is not None else df


# Lista BLANCA de columnas monetarias (2 decimales).
# Todo lo demás en columnas numéricas se trata como entero (unidades).
_MONEY_COLS = {
    # valores financieros
    "valor inventario","valor compras","valor ventas","valor stock",
    "valor unitario","valor unitario promedio","valor total","v.unit",
    "v.total","valor inv.",
    # costos
    "costo","costo prom.","costo promedio","costo ponderado",
    # precios
    "pvp","pvp promedio","precio","precio promedio","p. unitario",
    # márgenes y rentabilidad
    "margen","margen%","rentabilidad","utilidad",
    # ratios con decimales
    "rotación","días inv.","cons/día","p.reorden","sug.compra",
    "marítimo","aéreo",
}

def _is_int_col(col_name):
    """Retorna True si la columna debe mostrarse como entero (sin decimales)."""
    return col_name.lower().strip() not in _MONEY_COLS

def tbl(df, nc=None, uid="t"):
    """Tabla HTML con zoom y totales. Stock=entero, Valores=2dec."""
    if df is None or df.empty: return "<p>Sin datos</p>"
    nc=nc or []
    hdrs="".join(f"<th class='{'n' if c in nc else ''}'>{c}</th>" for c in df.columns)
    rows=""; tots=defaultdict(float); first_col=df.columns[0] if len(df.columns)>0 else ""
    for _,row in df.iterrows():
        cells=""
        for c in df.columns:
            v=row[c]; disp=str(v) if str(v) not in("nan","None","NaN") else ""
            if c in nc:
                try:
                    fv=float(v); tots[c]+=fv
                    disp=f"{int(round(fv)):,}" if _is_int_col(c) else f"{fv:,.2f}"
                except: pass
                cells+=f"<td class='n'>{disp}</td>"
            else: cells+=f"<td>{disp}</td>"
        rows+=f"<tr>{cells}</tr>"
    def _tot_fmt(c):
        v=tots[c]
        return f"{int(round(v)):,}" if _is_int_col(c) else f"{v:,.2f}"
    tcells="".join(
        (f"<td class='n'>{_tot_fmt(c)}</td>" if c in nc and tots[c]!=0 else
         (f"<td><b>TOTAL</b></td>" if c==first_col else "<td></td>"))
        for c in df.columns)
    uid_safe=uid.replace("-","_")
    return f"""<div class="zb">
  <span style="color:var(--text3);font-size:11px;font-weight:700">ZOOM</span>
  <button onclick="asZoom('{uid_safe}',-1)">−</button>
  <button onclick="asZoom('{uid_safe}',1)">+</button>
  <button onclick="asZoomReset('{uid_safe}')">↺</button>
  <span style="color:var(--text3);font-size:10px">{len(df):,} registros</span>
</div>
<div class="tc"><table class="it" id="tbl_{uid_safe}">
<thead><tr>{hdrs}</tr></thead><tbody>{rows}</tbody>
<tfoot><tr class="tot">{tcells}</tr></tfoot>
</table></div>"""

# ══ Función universal: tabla modo componente (sticky header + cols fijas) ══
def _comp_tbl(df, nc, uid, freeze_cols=2, height=600, title="", groups=None, legend=""):
    """
    Renderiza df como st.components.html con:
    - Header siempre visible (sticky top)
    - freeze_cols columnas izquierda fijas (sticky left)
    - Zebra, ceros vacíos en nc, zoom, two-finger pan
    groups: lista de dicts {label, color_bg, color_text, rows_df} para separadores
    legend: HTML extra debajo de la tabla
    """
    import json
    if df is None or df.empty:
        _components.html("<p style='color:#94a3b8;padding:12px'>Sin datos.</p>", height=60)
        return

    all_cols = list(df.columns)
    # Anchos fijos para cols congeladas
    FW = [110, 240, 120, 120]   # col0, col1, col2, col3

    def col_left(i):
        return sum(FW[:i]) if i < len(FW) else 0

    def fmt_val(v, col):
        s = str(v)
        if s in ("nan","None","NaN",""): return ""
        if col in nc:
            try:
                fv = float(v)
                return f"{int(round(fv)):,}" if _is_int_col(col) else f"{fv:,.2f}"
            except: return s
        return s

    # ── Encabezados ───────────────────────────────────────────
    hdrs = ""
    for i,c in enumerate(all_cols):
        left  = f"left:{col_left(i)}px;" if i < freeze_cols else ""
        zidx  = "5" if i < freeze_cols else "3"
        bdr_r = "border-right:2px solid #94a3b8;" if i == freeze_cols-1 else ("border-right:1px solid #e2e8f0;" if i < freeze_cols else "")
        align = "text-align:right;" if c in nc else ""
        hdrs += (f'<th style="position:sticky;top:0;{left}z-index:{zidx};'
                 f'background:#f1f5f9;{bdr_r}border-bottom:2px solid #cbd5e1;'
                 f'padding:7px 10px;font-size:10px;font-weight:700;'
                 f'text-transform:uppercase;color:#64748b;white-space:nowrap;{align}">{c}</th>')

    # ── Filas ─────────────────────────────────────────────────
    rows = ""; tots = {c:0.0 for c in nc}
    ri = 0
    if groups:
        # Tabla agrupada con separadores de sección (T_INV)
        for g in groups:
            g_df  = g["df"]
            g_lbl = g["label"]
            g_bg  = g.get("bg","#e0f2fe")
            g_col = g.get("col","#0369a1")
            n_sub = {c:0.0 for c in nc}
            # Fila separadora de grupo
            rows += (f'<tr><td colspan="{len(all_cols)}" '
                     f'style="background:{g_bg};padding:5px 12px;font-weight:700;'
                     f'font-size:11px;color:{g_col};letter-spacing:.04em;'
                     f'position:sticky;left:0">'
                     f'🏪 {g_lbl}</td></tr>')
            for _,row in g_df.iterrows():
                bg = "#f8fafc" if ri%2==0 else "#ffffff"; ri+=1
                cells = ""
                for i,c in enumerate(all_cols):
                    v = row[c]; disp = fmt_val(v, c)
                    left  = f"left:{col_left(i)}px;" if i < freeze_cols else ""
                    zidx  = "2" if i < freeze_cols else "0"
                    bdr_r = "border-right:2px solid #94a3b8;" if i==freeze_cols-1 else ("border-right:1px solid #e2e8f0;" if i<freeze_cols else "")
                    align = "text-align:right;font-family:monospace;" if c in nc else ""
                    cells += (f'<td style="position:sticky;{left}z-index:{zidx};'
                              f'background:{bg};{bdr_r}border-bottom:1px solid #f1f5f9;'
                              f'padding:6px 10px;font-size:12px;{align}">{disp}</td>')
                    if c in nc:
                        try: n_sub[c]+=float(v); tots[c]+=float(v)
                        except: pass
                rows += f"<tr>{cells}</tr>"
            # Subtotal grupo
            sc = ""
            for i,c in enumerate(all_cols):
                left  = f"left:{col_left(i)}px;" if i < freeze_cols else ""
                zidx  = "2" if i < freeze_cols else "0"
                bdr_r = "border-right:2px solid #94a3b8;" if i==freeze_cols-1 else ("border-right:1px solid #e2e8f0;" if i<freeze_cols else "")
                if i==0:
                    sc += (f'<td style="position:sticky;{left}z-index:{zidx};'
                           f'background:#dbeafe;{bdr_r}border-top:1px solid #bfdbfe;'
                           f'padding:5px 10px;font-size:10px;font-weight:700;color:#1d4ed8">'
                           f'{g_lbl} — subtotal</td>')
                elif c in nc and n_sub[c]!=0:
                    d = f"{int(round(n_sub[c])):,}" if _is_int_col(c) else f"{n_sub[c]:,.2f}"
                    sc += (f'<td style="position:sticky;{left}z-index:{zidx};'
                           f'background:#dbeafe;{bdr_r}border-top:1px solid #bfdbfe;'
                           f'text-align:right;font-family:monospace;padding:5px 10px;'
                           f'font-weight:700;color:#1d4ed8">{d}</td>')
                else:
                    sc += (f'<td style="position:sticky;{left}z-index:{zidx};'
                           f'background:#dbeafe;{bdr_r}border-top:1px solid #bfdbfe;'
                           f'padding:5px 10px"></td>')
            rows += f"<tr>{sc}</tr>"
    else:
        # Tabla plana (T_SKU unidades, T_SAM)
        for _,row in df.iterrows():
            bg = "#f8fafc" if ri%2==0 else "#ffffff"; ri+=1
            cells = ""
            for i,c in enumerate(all_cols):
                v = row[c]; disp = fmt_val(v, c)
                left  = f"left:{col_left(i)}px;" if i < freeze_cols else ""
                zidx  = "2" if i < freeze_cols else "0"
                bdr_r = "border-right:2px solid #94a3b8;" if i==freeze_cols-1 else ("border-right:1px solid #e2e8f0;" if i<freeze_cols else "")
                align = "text-align:right;font-family:monospace;" if c in nc else ""
                cells += (f'<td style="position:sticky;{left}z-index:{zidx};'
                          f'background:{bg};{bdr_r}border-bottom:1px solid #f1f5f9;'
                          f'padding:6px 10px;font-size:12px;{align}">{disp}</td>')
                if c in nc:
                    try: tots[c]+=float(row[c])
                    except: pass
            rows += f"<tr>{cells}</tr>"

    # ── Fila TOTAL ────────────────────────────────────────────
    tfooter = ""
    for i,c in enumerate(all_cols):
        left  = f"left:{col_left(i)}px;" if i < freeze_cols else ""
        zidx  = "2" if i < freeze_cols else "0"
        bdr_r = "border-right:2px solid #94a3b8;" if i==freeze_cols-1 else ("border-right:1px solid #e2e8f0;" if i<freeze_cols else "")
        S = (f"position:sticky;{left}z-index:{zidx};{bdr_r}"
             f"background:#e0f2fe;font-weight:700;padding:7px 10px;"
             f"border-top:2px solid #7dd3fc;color:#0369a1")
        if i==0:
            tfooter += f'<td style="{S}">TOTAL GENERAL</td>'
        elif c in nc and tots[c]!=0:
            d = f"{int(round(tots[c])):,}" if _is_int_col(c) else f"{tots[c]:,.2f}"
            tfooter += f'<td style="{S};text-align:right;font-family:monospace">{d}</td>'
        else:
            tfooter += f'<td style="{S}"></td>'

    n_rows = len(df) if not groups else sum(len(g["df"]) for g in groups)
    n_info = title if title else f"{n_rows:,} registros"

    legend_html = f"<div style='font-size:10px;color:#64748b;margin-top:6px;padding:6px 10px;background:#f8fafc;border-radius:6px;border-left:3px solid #0ea5e9'>{legend}</div>" if legend else ""

    html = (
        "<!DOCTYPE html><html><head><meta charset=\"UTF-8\">"
        "<style>"
        "*{box-sizing:border-box;margin:0;padding:0}"
        "body{font-family:Inter,Segoe UI,sans-serif;background:#f0f9ff}"
        ".wrap{border:1px solid #e2e8f0;border-radius:8px;overflow:auto;"
        f"max-height:{height}px;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,.08)}}"
        ".zb{display:flex;align-items:center;gap:6px;margin-bottom:6px}"
        ".zb span{font-size:11px;font-weight:700;color:#94a3b8}"
        ".zb .inf{font-size:10px;font-weight:400}"
        ".zb button{background:#fff;border:1px solid #e2e8f0;border-radius:5px;"
        "padding:2px 10px;cursor:pointer;font-weight:700;font-size:13px;color:#475569}"
        ".zb button:hover{background:#0ea5e9;color:#fff;border-color:#0ea5e9}"
        "table{border-collapse:separate;border-spacing:0;width:max-content}"
        "tr:hover td{background:#e0f2fe!important}"
        "</style></head><body>"
        f'<div class=\"zb\">'
        "<span>ZOOM</span>"
        "<button onclick=\"var t=document.getElementById('_t');var s=parseFloat(t.style.fontSize||'12');t.style.fontSize=(s-1)+'px'\">−</button>"
        "<button onclick=\"var t=document.getElementById('_t');var s=parseFloat(t.style.fontSize||'12');t.style.fontSize=(s+1)+'px'\">+</button>"
        "<button onclick=\"document.getElementById('_t').style.fontSize='12px'\">↺</button>"
        f'<span class=\"inf\">{n_info}</span>'
        "</div>"
        '<div class=\"wrap\">'
        f'<table id=\"_t\"><thead><tr>{hdrs}</tr></thead>'
        f'<tbody>{rows}</tbody>'
        f'<tfoot><tr>{tfooter}</tr></tfoot>'
        "</table></div>"
        f"{legend_html}"
        "<script>"
        "var w=document.querySelector('.wrap');"
        "var t0x,t0y,t0sl,t0st;"
        "w.addEventListener('touchstart',function(e){if(e.touches.length!==2)return;"
        "t0x=(e.touches[0].pageX+e.touches[1].pageX)/2;"
        "t0y=(e.touches[0].pageY+e.touches[1].pageY)/2;"
        "t0sl=w.scrollLeft;t0st=w.scrollTop;},{passive:true});"
        "w.addEventListener('touchmove',function(e){if(e.touches.length!==2)return;"
        "e.preventDefault();"
        "var cx=(e.touches[0].pageX+e.touches[1].pageX)/2;"
        "var cy=(e.touches[0].pageY+e.touches[1].pageY)/2;"
        "w.scrollLeft=t0sl-(cx-t0x);w.scrollTop=t0st-(cy-t0y);"
        "},{passive:false});"
        "w.addEventListener('wheel',function(e){"
        "if(Math.abs(e.deltaX)>Math.abs(e.deltaY)){e.preventDefault();w.scrollLeft+=e.deltaX;}"
        "else if(e.shiftKey){e.preventDefault();w.scrollLeft+=e.deltaY;}"
        "},{passive:false});"
        "</script></body></html>"
    )
    _components.html(html, height=height+60, scrolling=False)


# ── Sidebar ─────────────────────────────────────────────────────
_perf("before_sidebar")
with st.sidebar:
    ca,cb=st.columns([4,1])
    with ca:
        st.markdown(
            f'''<div style="font-size:16px;font-weight:800;color:#0ea5e9;
            font-family:Inter,sans-serif;letter-spacing:.04em">
            AUTO<span style="font-weight:300;opacity:.8">SKY</span>
            </div>
            <div style="font-size:10px;font-family:monospace;color:#475569;margin-top:2px">
            <b style="color:#0ea5e9">{APP_VERSION}</b> &nbsp;|&nbsp; {BUILD_TIME}
            </div>''',
            unsafe_allow_html=True
        )
    with cb:
        st.markdown("")
        nd=st.toggle("🌙",value=dark,help="Tema oscuro/claro")
        if nd!=dark: st.session_state.dark_mode=nd; st.rerun()
    st.divider()

    st.markdown("### 📂 Cargar Excel")
    st.caption("XLS / XLSX del sistema contable — acumulativo")
    uploaded=st.file_uploader("Archivos",type=["xlsx","xls"],
                               accept_multiple_files=True,label_visibility="collapsed")
    if uploaded:
        # Al subir archivos reales, remover la entrada sintética del consolidado persistido
        st.session_state.files_loaded[:] = [
            f for f in st.session_state.files_loaded if "(persistido)" not in f
        ]
        new=[f for f in uploaded if f.name not in st.session_state.files_loaded]
        if new:
            for uf in new:
                with st.spinner(f"Procesando {uf.name}..."):
                    try:
                        suf=os.path.splitext(uf.name)[1]
                        with tempfile.NamedTemporaryFile(delete=False,suffix=suf) as tmp:
                            tmp.write(uf.getvalue()); tp=tmp.name
                        if eng.raw_df is None:
                            eng.load_inventory_file(tp)
                            _nuevos=len(eng.raw_df); _dupes=0
                        else:
                            from app.engine import InventoryEngine as _IE
                            te=_IE(); te.load_inventory_file(tp)
                            # Deduplicar por columna "Código" (N° de movimiento único)
                            if "Código" in eng.raw_df.columns and "Código" in te.raw_df.columns:
                                _exist=set(eng.raw_df["Código"].astype(str).str.strip())
                                _mask=~te.raw_df["Código"].astype(str).str.strip().isin(_exist)
                                _dupes=int((~_mask).sum()); _nuevos=int(_mask.sum())
                                if _nuevos>0:
                                    eng.raw_df=pd.concat([eng.raw_df,te.raw_df[_mask]],ignore_index=True)
                            else:
                                # Fallback: deduplicar por todas las columnas
                                _prev=len(eng.raw_df)
                                eng.raw_df=pd.concat([eng.raw_df,te.raw_df],ignore_index=True).drop_duplicates()
                                _nuevos=len(eng.raw_df)-_prev; _dupes=len(te.raw_df)-_nuevos
                        try: os.unlink(tp)
                        except: pass
                        st.session_state.files_loaded.append(uf.name)
                        # Guardar stats por archivo
                        st.session_state.setdefault("files_stats",[]).append({
                            "nombre": uf.name,
                            "nuevos": _nuevos,
                            "dupes":  _dupes,
                        })
                        st.session_state.result=None
                        _msg=f"✓ {uf.name} | {_nuevos:,} nuevos"
                        if _dupes: _msg+=f" | {_dupes} duplicados omitidos"
                        log(_msg)
                        # Marcar para recálculo automático
                        st.session_state["_recalc_pending"]=True
                    except Exception as e: st.error(f"{uf.name}: {e}")
            # Persistir consolidado para que todas las sesiones/clientes
            # vean los mismos datos y sobrevivan al reinicio del servidor
            if eng.raw_df is not None:
                _persist_raw(eng.raw_df)
            st.rerun()

    if st.session_state.files_loaded:
        # ── Rango de fechas consolidado ────────────────────────
        if eng.raw_df is not None and "Fecha" in eng.raw_df.columns:
            try:
                _dates=pd.to_datetime(eng.raw_df["Fecha"],errors="coerce").dropna()
                if not _dates.empty:
                    _d1=_dates.min().strftime("%d/%m/%Y")
                    _d2=_dates.max().strftime("%d/%m/%Y")
                    _pc_bg    = "#1e3a5f" if dark else "#f0f9ff"
                    _pc_bdr   = "#2d5a8e" if dark else "#bae6fd"
                    _pc_title = "#7dd3fc" if dark else "#0284c7"
                    _pc_date  = "#f1f5f9" if dark else "#0f172a"
                    _pc_muted = "#94a3b8" if dark else "#64748b"
                    st.markdown(
                        f"<div style='background:{_pc_bg};border:1px solid {_pc_bdr};"
                        f"border-radius:8px;padding:8px 12px;font-size:11px;margin-bottom:4px'>"
                        f"<b style='color:{_pc_title}'>📅 Período cargado</b><br>"
                        f"<span style='font-size:13px;font-weight:700;color:{_pc_date}'>"
                        f"{_d1} → {_d2}</span><br>"
                        f"<span style='color:{_pc_muted}'>{len(eng.raw_df):,} movimientos · "
                        f"{len(st.session_state.files_loaded)} archivo(s)</span>"
                        f"</div>",
                        unsafe_allow_html=True)
            except: pass

        # ── Detalle por archivo ────────────────────────────────
        for _st_item in st.session_state.get("files_stats",[]):
            _dup_txt=f" · <span style='color:#f59e0b'>{_st_item['dupes']} dup.</span>" if _st_item.get("dupes") else ""
            st.markdown(
                f"<div style='font-size:11px;color:{MUTED};padding:2px 4px'>"
                f"📄 {_st_item['nombre']} — "
                f"<b style='color:#059669'>{_st_item['nuevos']:,}</b> reg.{_dup_txt}"
                f"</div>",
                unsafe_allow_html=True)

        c1,c2=st.columns(2)
        with c1:
            if st.button("🗑 Limpiar",use_container_width=True):
                # Mutar en sitio para afectar a todas las sesiones
                eng.raw_df=None; eng.physical_df=None
                st.session_state.files_loaded.clear()
                st.session_state.files_stats.clear()
                st.session_state.result=None
                for _p in (CONSOLIDADO_PATH, TOMA_FISICA_PATH):
                    try:
                        if os.path.exists(_p): os.unlink(_p)
                    except Exception: pass
                st.rerun()
        with c2:
            if eng.raw_df is not None:
                st.download_button("📥 Exportar",to_xl(eng.raw_df),
                    "consolidado.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,key="exp_cons")
    st.divider()

    if eng.raw_df is not None:
        st.markdown("### 🚫 Exclusiones globales")
        # Mapa código → nombre para búsqueda predictiva por ambos
        _excl_map = (eng.raw_df[["Código Producto","Nombre Producto"]]
                     .drop_duplicates().sort_values("Nombre Producto"))
        _excl_map = dict(zip(_excl_map["Código Producto"].astype(str),
                             _excl_map["Nombre Producto"].astype(str)))
        all_skus = list(_excl_map.keys())

        _fmt_sku = lambda s: f"{s} — {_excl_map.get(s, '')}"
        excl_s=st.multiselect("Excluir SKUs", all_skus,
                               default=list(st.session_state.excluded_skus),
                               key="es_ms",
                               format_func=_fmt_sku,
                               placeholder="Escribe código o nombre…")
        _prev_excl_s = set(st.session_state.excluded_skus)
        st.session_state.excluded_skus=set(excl_s); eng.excluded_skus=set(excl_s)
        if _prev_excl_s != set(excl_s):
            st.session_state["_recalc_pending"] = True

        all_wh=eng.get_warehouses()
        excl_w=st.multiselect("Excluir Bodegas", all_wh,
                               default=list(st.session_state.excl_wh),
                               key="ew_ms",
                               placeholder="Escribe para filtrar…")
        _prev_excl_w = set(st.session_state.excl_wh)
        st.session_state.excl_wh=set(excl_w)
        # Exclusión GLOBAL: descarta movimientos (origen o destino) en esas bodegas.
        # Afecta KPIs, rotación, kardex, compras — todo el pipeline.
        eng.excluded_warehouses = set(excl_w)
        if _prev_excl_w != set(excl_w):
            st.session_state["_recalc_pending"] = True
        st.divider()

        st.markdown("### ⚙️ Calcular")
        cutoff=st.date_input("Fecha de corte",date.today(),format="DD/MM/YYYY")
        wh_mode=st.selectbox("Bodegas",["Todas","Solo principal","Selección manual"])
        sel_wh=[]
        if wh_mode=="Selección manual": sel_wh=st.multiselect("Bodegas",all_wh)
        if st.button("▶ Calcular",type="primary",use_container_width=True):
            with st.spinner("Calculando..."):
                try:
                    r=eng.analyze(str(cutoff),wh_mode,sel_wh)
                    st.session_state.result=r; log(f"OK | {cutoff} | {len(r.filtered):,} mov.")
                    st.success("✓")
                except Exception as e: st.error(str(e)); log(f"Error: {e}")

        # Auto-recálculo cuando se carga un archivo nuevo
        if st.session_state.pop("_recalc_pending", False) and eng.raw_df is not None:
            try:
                _ar=eng.analyze(str(cutoff),wh_mode,sel_wh)
                st.session_state.result=_ar
                log(f"Auto-calculado | {cutoff} | {len(_ar.filtered):,} mov.")
            except: pass

r=st.session_state.result

# ── Sin datos ────────────────────────────────────────────────────
if eng.raw_df is None:
    st.markdown("## 📦 Sistema de Inventario")
    _card_bg  = "#1e293b" if dark else "#f0f9ff"
    _card_bdr = "#334155" if dark else "#bae6fd"
    _card_ver = "#38bdf8" if dark else "#0284c7"
    _card_sep = "#475569" if dark else "#94a3b8"
    _card_txt = "#94a3b8" if dark else "#475569"
    st.markdown(f"""
<div style='display:inline-flex;align-items:center;gap:12px;
  background:{_card_bg};border:1px solid {_card_bdr};border-radius:8px;
  padding:8px 16px;margin-bottom:16px;'>
  <span style='color:{_card_ver};font-weight:700;font-family:monospace;font-size:14px'>{APP_VERSION}</span>
  <span style='color:{_card_sep}'>|</span>
  <span style='color:{_card_txt};font-family:monospace;font-size:12px'>{BUILD_TIME}</span>
</div>""", unsafe_allow_html=True)
    st.info("👈 Cargue uno o más archivos Excel. Puede cargar archivos adicionales y los datos se acumularán.")
    st.stop()

# ── Banner Autosky ───────────────────────────────────────────────
st.markdown(f"""<div class="as-banner">
  <div>
    <div class="as-logo">AUTO<span>SKY</span>&nbsp;
      <span style="font-size:12px;font-weight:400;opacity:.85">Sistema de Inventario</span>
    </div>
    <div style="font-size:10px;opacity:.8;color:#fff">Gestión · Análisis · Control de Inventario</div>
  </div>
  <div class="as-build">
    <div class="v">{APP_VERSION}</div>
    <div>{BUILD_TIME}</div>
  </div>
</div>""", unsafe_allow_html=True)

# ── Período de datos cargados ───────────────────────────────────
if eng.raw_df is not None and "Fecha" in eng.raw_df.columns:
    try:
        _all_d=pd.to_datetime(eng.raw_df["Fecha"],errors="coerce").dropna()
        if not _all_d.empty:
            _pd1=_all_d.min().strftime("%d/%m/%Y")
            _pd2=_all_d.max().strftime("%d/%m/%Y")
            _pb_bg  = "#1e3a5f" if dark else "#f0f9ff"
            _pb_bdr = "#2d5a8e" if dark else "#bae6fd"
            _pb_lbl = "#7dd3fc" if dark else "#475569"
            _pb_dt  = "#f1f5f9" if dark else "#0f172a"
            _pb_arr = "#7dd3fc" if dark else "#94a3b8"
            _pb_mut = "#94a3b8" if dark else "#64748b"
            st.markdown(
                f"<div style='display:inline-flex;align-items:center;gap:16px;"
                f"background:{_pb_bg};border:1px solid {_pb_bdr};border-radius:8px;"
                f"padding:7px 16px;margin-bottom:8px;font-size:12px'>"
                f"<span style='color:{_pb_lbl};font-weight:600'>📅 PERÍODO DE ANÁLISIS</span>"
                f"<span style='color:{_pb_dt};font-weight:700;font-size:14px'>{_pd1}</span>"
                f"<span style='color:{_pb_arr}'>→</span>"
                f"<span style='color:{_pb_dt};font-weight:700;font-size:14px'>{_pd2}</span>"
                f"<span style='color:{_pb_mut}'>{len(eng.raw_df):,} movimientos</span>"
                f"</div>",
                unsafe_allow_html=True)
    except: pass

# ── Aviso de exclusiones globales activas ──────────────────────
_act_excl_sku = list(st.session_state.get("excluded_skus", set()))
_act_excl_wh  = list(st.session_state.get("excl_wh", set()))
if _act_excl_sku or _act_excl_wh:
    _ex_lines = []
    if _act_excl_wh:
        _ex_lines.append(
            f"🏪 <b>{len(_act_excl_wh)} bodega(s) excluida(s):</b> "
            + ", ".join(f"<code>{b}</code>" for b in _act_excl_wh[:5])
            + (f" <i>+{len(_act_excl_wh)-5} más</i>" if len(_act_excl_wh) > 5 else "")
        )
    if _act_excl_sku:
        _ex_lines.append(
            f"🏷 <b>{len(_act_excl_sku)} SKU(s) excluido(s)</b>"
        )
    st.markdown(
        f"<div style='background:#fef3c7;border:1px solid #f59e0b;border-radius:8px;"
        f"padding:8px 14px;margin-bottom:10px;font-size:12px;color:#92400e;"
        f"line-height:1.5'>"
        f"<b>⚠ Exclusiones globales activas</b> — los movimientos relacionados se "
        f"descartan del pipeline (KPIs, Rotación, Kardex, Compras, todo).<br>"
        + "<br>".join(_ex_lines) +
        f"</div>",
        unsafe_allow_html=True
    )

# ── KPIs ────────────────────────────────────────────────────────
if r is not None:
    kpis=r.kpis
    r1="".join([kc("Stock Total",fmt(kpis.get("Stock total",0),"i"),"a"),
                kc("Disponible",fmt(kpis.get("Stock disponible",0),"i"),"s"),
                kc("Muestras",fmt(kpis.get("Stock en muestras",0),"i"),"w"),
                kc("Valor Inv.",f'${fmt(kpis.get("Valor inventario",0))}', "a"),
                kc("Compras Acum.",f'${fmt(kpis.get("Compras acumuladas",0))}', "d")])
    r2="".join([kc("Rotación",fmt(kpis.get("Rotación",0),"p")),
                kc("Días Inv.",f'{kpis.get("Días de inventario",0):.0f} d'),
                kc("Consumo/día",f'{kpis.get("Consumo promedio",0):.2f} u'),
                kc("Exactitud",fmt(kpis.get("Exactitud inventario",0),"p"))])
    st.markdown(f'<div class="kpi-row">{r1}</div>',unsafe_allow_html=True)
    st.markdown(f'<div class="kpi-row">{r2}</div>',unsafe_allow_html=True)
    st.markdown("---")


_perf("main_kpis_done")

# ── Pestañas ─────────────────────────────────────────────────────
tabs=st.tabs(["🏪 Inv×Bodega","🔍 Detalle SKU","📊 SKU×Bodega","👥 Muestras",
              "📈 Período","🔄 Rotación","📐 Cálculos","🧾 Compras","📋 Kardex","🏭 Toma Física","📝 Log"])
(T_INV,T_SKU,T_PIV,T_SAM,T_ANA,T_ROT,T_CAL,T_PUR,T_KDX,T_PHY,T_LOG)=tabs
_perf("tabs_defined")

excl_s=list(st.session_state.excluded_skus)
excl_w=list(st.session_state.excl_wh)

# ══ TAB 1 INVENTARIO ════════════════════════════════════════════
@_fragment
def _render_tab_inv():
    r = st.session_state.get("result")
    eng = st.session_state.engine
    if r is None: st.info("Ejecute el análisis.")
    else:
        df_inv=r.inventory_by_warehouse.copy()
        if excl_w: df_inv=df_inv[~df_inv["Bodega"].isin(excl_w)]
        for _drop in ("Grupo Visual","GRUPO VISUAL","Grupo Movimiento","Categoría Producto"):
            if _drop in df_inv.columns: df_inv=df_inv.drop(columns=[_drop])

        # ── Filtros ───────────────────────────────────────────────
        _all_skus=sorted(df_inv["Código Producto"].dropna().unique().tolist()) if "Código Producto" in df_inv.columns else []
        _all_bods=sorted(df_inv["Bodega"].dropna().unique().tolist()) if "Bodega" in df_inv.columns else []

        # Mapa SKU → nombre para búsqueda rápida
        _inv_name_map = dict(zip(df_inv["Código Producto"].astype(str),
                                 df_inv["Nombre Producto"].astype(str))) if "Código Producto" in df_inv.columns else {}
        _inv_fmt = lambda s: f"{s} — {_inv_name_map.get(s, '')}" if _inv_name_map.get(s) else s

        fc1,fc2,fc3,fc4=st.columns([4,4,1,1])
        with fc1:
            sel_skus=st.multiselect("🔍 SKU / Producto", _all_skus,
                key="i_skus", placeholder="Escribe código o nombre…",
                format_func=_inv_fmt)
        with fc2:
            sel_bods=st.multiselect("🏪 Bodega", _all_bods,
                key="i_bods", placeholder="Escribe para filtrar…")
        with fc3:
            st.markdown("")
            stk_only=st.checkbox("Solo con stock",True,key="i_s")
        with fc4:
            st.markdown("")
            st.button("✔ Aplicar",key="i_apply",use_container_width=True,
                      help="Cerrar selector y aplicar filtros")

        # Aplicar filtros
        df=df_inv.copy()
        if sel_skus: df=df[df["Código Producto"].isin(sel_skus)]
        if sel_bods: df=df[df["Bodega"].isin(sel_bods)]
        if stk_only and "Stock" in df.columns: df=df[df["Stock"]>0]

        # Columnas numéricas — excluir Bodega de la columna visible
        _skip_cols=("Código Producto","Nombre Producto","Bodega")
        nc=[c for c in df.columns if c not in _skip_cols]

        # ── Tabla agrupada por Bodega (sin columna Bodega) ────────
        def _inv_grouped(df, nc, uid):
            if df is None or df.empty: return "<p style='color:var(--text3);padding:12px'>Sin datos para los filtros seleccionados.</p>"
            uid_safe=uid.replace("-","_")
            bodegas=sorted(df["Bodega"].dropna().unique().tolist()) if "Bodega" in df.columns else []
            # Columnas a mostrar: excluir "Bodega" de la tabla
            show_cols=[c for c in df.columns if c!="Bodega"]
            nc_show=[c for c in nc if c in show_cols]
            hdrs="".join(f"<th class='{'n' if c in nc_show else ''}'>{c}</th>" for c in show_cols)
            body=""
            grand=defaultdict(float)
            for bod_name in bodegas:
                sub=df[df["Bodega"]==bod_name] if "Bodega" in df.columns else df
                if sub.empty: continue
                n_cols=len(show_cols)
                body+=f"<tr style='background:var(--sky-l)'><td colspan='{n_cols}' style='padding:5px 12px;font-weight:700;font-size:11px;color:var(--sky-d);letter-spacing:.04em'>🏪 {bod_name}</td></tr>"
                sub_tots=defaultdict(float)
                for _,row in sub.iterrows():
                    cells=""
                    for c in show_cols:
                        v=row[c]; raw=str(v)
                        disp="" if raw in("nan","None","NaN") else raw
                        if c in nc_show:
                            try:
                                fv=float(v); sub_tots[c]+=fv; grand[c]+=fv
                                disp=f"{int(round(fv)):,}" if _is_int_col(c) else f"{fv:,.2f}"
                            except: pass
                            cells+=f"<td class='n'>{disp}</td>"
                        else:
                            cells+=f"<td>{disp}</td>"
                    body+=f"<tr>{cells}</tr>"
                # Subtotal bodega
                sc=""
                for i,c in enumerate(show_cols):
                    if i==0:
                        sc+=f"<td style='font-weight:700;font-size:10px;color:var(--sky-d)'>{bod_name} — subtotal</td>"
                    elif c in nc_show and sub_tots[c]!=0:
                        v=sub_tots[c]
                        d=f"{int(round(v)):,}" if _is_int_col(c) else f"{v:,.2f}"
                        sc+=f"<td class='n' style='font-weight:700;color:var(--sky-d)'>{d}</td>"
                    else:
                        sc+="<td></td>"
                body+=f"<tr style='background:var(--sky-ll);border-top:1px solid var(--border2)'>{sc}</tr>"
            # TOTAL GENERAL
            gc=""
            for i,c in enumerate(show_cols):
                if i==0:
                    gc+=f"<td style='font-weight:800;color:var(--sky-d)'>TOTAL GENERAL</td>"
                elif c in nc_show and grand[c]!=0:
                    v=grand[c]
                    d=f"{int(round(v)):,}" if _is_int_col(c) else f"{v:,.2f}"
                    gc+=f"<td class='n' style='font-weight:800;color:var(--sky-d)'>{d}</td>"
                else:
                    gc+="<td></td>"
            n_shown=sum(1 for _,r in df.iterrows() for _ in [None])
            return f"""<div class="zb">
  <span style="color:var(--text3);font-size:11px;font-weight:700">ZOOM</span>
  <button onclick="asZoom('{uid_safe}',-1)">−</button>
  <button onclick="asZoom('{uid_safe}',1)">+</button>
  <button onclick="asZoomReset('{uid_safe}')">↺</button>
  <span style="color:var(--text3);font-size:10px">{len(df):,} registros · {len(bodegas)} bodega(s)</span>
</div>
<div class="tc"><table class="it" id="tbl_{uid_safe}">
<thead><tr>{hdrs}</tr></thead><tbody>{body}</tbody>
<tfoot><tr class="tot">{gc}</tr></tfoot>
</table></div>"""
        # Construir groups para _comp_tbl
        _bodegas = sorted(df["Bodega"].dropna().unique().tolist()) if "Bodega" in df.columns else []
        _show_cols = [c for c in df.columns if c != "Bodega"]
        _nc_show   = [c for c in nc if c in _show_cols]
        _groups = []
        for _bn in _bodegas:
            _sub = df[df["Bodega"]==_bn][_show_cols].copy() if "Bodega" in df.columns else df[_show_cols].copy()
            if not _sub.empty:
                _groups.append({"label":_bn, "df":_sub,
                                 "bg":"#e0f2fe", "col":"#0369a1"})
        _comp_tbl(df[_show_cols], _nc_show, "inv",
                  freeze_cols=2, height=580,
                  title=f"{len(df):,} registros · {len(_bodegas)} bodega(s)",
                  groups=_groups if len(_bodegas)>1 else None)
        dl3(df,"inventario_bodega","inv")

with T_INV:
    _render_tab_inv()

# ══ TAB 2 DETALLE SKU ═══════════════════════════════════════════
@_fragment
def _render_tab_sku():
    r = st.session_state.get("result")
    eng = st.session_state.engine
    if r is None: st.info("Ejecute el análisis.")
    else:
        df=r.sku_summary.copy()

        # ── Renombrar columnas del engine a labels legibles ───────
        _ren = {
            "Dev_Proveedor":     "Dev. Proveedor",
            "Dev_Cliente":       "Dev. Cliente",
            "Baja_Inventario":   "Baja Inv.",
            "Muestras_Enviadas": "Muestras Env.",
            "Muestras_Devueltas":"Muestras Dev.",
            "Valor_Compras":     "Valor Compras",
            "Valor_Ventas":      "Valor Ventas",
        }
        df = df.rename(columns={k:v for k,v in _ren.items() if k in df.columns})

        # ── Forzar enteros en todas las columnas de unidades ──────
        _unit_cols = ["Compras","Dev. Proveedor","Ventas","Dev. Cliente",
                      "Baja Inv.",
                      "Muestras Env.","Muestras Dev.",
                      "Stock Disponible","Stock Muestras","Stock Total"]
        for _uc in _unit_cols:
            if _uc in df.columns:
                df[_uc] = df[_uc].fillna(0).astype(int)

        # ── Cuadre: solo movimientos reales (muestras son internas)
        #   Compras − Dev.Proveedor − Ventas + Dev.Cliente = Stock Total
        #   Muestras Env./Dev. son transferencias internas — no afectan Stock Total
        # ── Costo Promedio Ponderado = Valor_Compras / Compras ───
        if "Valor Compras" in df.columns and "Compras" in df.columns:
            df["Costo Prom."] = (
                df["Valor Compras"] / df["Compras"].replace(0, float("nan"))
            ).round(2).fillna(0.0)
        elif "Valor_Compras" in df.columns and "Compras" in df.columns:
            df["Costo Prom."] = (
                df["Valor_Compras"] / df["Compras"].replace(0, float("nan"))
            ).round(2).fillna(0.0)

        def _safe(d, col): return d[col] if col in d.columns else 0
        df["✓ Cuadre"] = (
              _safe(df,"Compras")
            - _safe(df,"Dev. Proveedor")
            - _safe(df,"Baja Inv.")
            - _safe(df,"Ventas")
            + _safe(df,"Dev. Cliente")
        ).astype(int)
        df["Δ vs Stock"] = (df["✓ Cuadre"] - _safe(df,"Stock Total")).astype(int)

        # Multiselect predictivo con filtro estricto de substring
        _sk_opts_df = df[["Código Producto","Nombre Producto"]].drop_duplicates()
        _sk_opts_df = _sk_opts_df.sort_values("Código Producto")
        _sk_labels = [f"{c} — {n}" for c,n in zip(
            _sk_opts_df["Código Producto"].astype(str),
            _sk_opts_df["Nombre Producto"].astype(str))]

        sk_f1, sk_f2 = st.columns([3, 2])
        with sk_f1:
            _sk_sel = st.multiselect("🔍 SKU / Producto", _sk_labels, key="sk_f",
                                      placeholder="Escribe código o nombre…")
        with sk_f2:
            _sort_opts = {
                "Código ↑":            ("Código Producto",      True),
                "Código ↓":            ("Código Producto",      False),
                "Nombre ↑":            ("Nombre Producto",      True),
                "Nombre ↓":            ("Nombre Producto",      False),
                "Stock Total ↓":       ("Stock Total",          False),
                "Ventas ↓":            ("Ventas",               False),
                "Valor Inventario ↓":  ("Valor Inventario",     False),
                "Δ vs Stock ↓":        ("Δ vs Stock",           False),
            }
            _sort_label = st.selectbox("Ordenar por",
                                       list(_sort_opts.keys()),
                                       index=0, key="sk_sort",
                                       help="Default: Código ascendente (útil para búsqueda por SKU).")
        _sort_col, _sort_asc = _sort_opts[_sort_label]

        if _sk_sel:
            _sk_codes_sel = {s.split(" — ")[0] for s in _sk_sel}
            df = df[df["Código Producto"].astype(str).isin(_sk_codes_sel)]

        # Ordenamiento aplicado a las tablas que se muestran abajo
        if _sort_col in df.columns:
            df = df.sort_values(_sort_col, ascending=_sort_asc,
                                na_position="last").reset_index(drop=True)

        # ── Tabla Unidades ────────────────────────────────────────
        # Orden: flujo de entrada/salida | resultado | muestras (info) | cuadre
        st.markdown("##### 📦 Movimiento de Unidades")
        _u_cols = [
            "Código Producto","Nombre Producto",
            "Costo Prom.",               # último costo promedio ponderado
            "Compras","Dev. Proveedor","Baja Inv.",
            "Ventas","Dev. Cliente",
            "Stock Disponible","Stock Total",
            "✓ Cuadre","Δ vs Stock",
            "Stock Muestras",
        ]
        mu=[c for c in _u_cols if c in df.columns]
        nu=[c for c in mu if c not in("Código Producto","Nombre Producto")]
        _comp_tbl(df[mu], nu, "su", freeze_cols=2, height=520,
                  title=f"{len(df):,} SKUs",
                  legend="<b>✓ Cuadre</b> = Compras − Dev.Proveedor − <b>Baja Inv.</b> − Ventas + Dev.Cliente &nbsp;·&nbsp; "
                         "<b>Δ vs Stock</b> = Cuadre − Stock Total (0 = correcto) &nbsp;·&nbsp; "
                         "<b>Baja Inv.</b>: EGR con Descripción «BAJA DE INVENTARIO» (merma, deterioro, obsoleto) &nbsp;·&nbsp; "
                         "<b>Stock Muestras</b>: informativo, transferencias internas")

        st.markdown("")
        # ── Tabla Valores Financieros ─────────────────────────────
        st.markdown("##### 💰 Valores Financieros")
        _f_cols = [
            "Código Producto","Nombre Producto",
            "Valor Compras","Valor Ventas","Valor Inventario",
        ]
        mf=[c for c in _f_cols if c in df.columns]
        nf=[c for c in mf if c not in("Código Producto","Nombre Producto")]
        st.markdown(tbl(df[mf],nf,"sf"),unsafe_allow_html=True)

        dl3(df,"detalle_sku","sku")

with T_SKU:
    _render_tab_sku()

# ══ TAB 3 SKU×BODEGA ════════════════════════════════════════════
@_fragment
def _render_tab_piv():
    r = st.session_state.get("result")
    eng = st.session_state.engine
    if r is None: st.info("Ejecute el análisis.")
    else:
        df=r.inventory_by_warehouse.copy()
        if excl_w: df=df[~df["Bodega"].isin(excl_w)]

        # Filtros
        _pv_skus=sorted(df["Código Producto"].dropna().unique().tolist()) if "Código Producto" in df.columns else []
        # Mapa para búsqueda predictiva
        _pv_name_map = dict(zip(df["Código Producto"].astype(str),
                                df["Nombre Producto"].astype(str))) if "Código Producto" in df.columns else {}
        _pv_fmt = lambda s: f"{s} — {_pv_name_map.get(s, '')}" if _pv_name_map.get(s) else s

        pc1,pc2,pc3=st.columns([5,1,1])
        with pc1:
            sel_pv=st.multiselect("🔍 SKU", _pv_skus, key="pv_skus",
                placeholder="Escribe código o nombre…", format_func=_pv_fmt)
        with pc2:
            st.markdown("")
            excl_bp=st.checkbox("Excluir Bodega Ppal",True,key="pv_e")
        with pc3:
            st.markdown("")
            # Botón Aplicar: cierra el dropdown y fuerza rerun
            st.button("✔ Aplicar",key="pv_apply",
                      help="Cerrar selector y aplicar filtros",
                      use_container_width=True)

        if excl_bp: df=df[df["Bodega"]!=PRIMARY_WAREHOUSE]
        if sel_pv:  df=df[df["Código Producto"].isin(sel_pv)]
        try:
            pv=df.pivot_table(index=["Código Producto","Nombre Producto"],
                              columns="Bodega",values="Stock",
                              aggfunc="sum",fill_value=0).reset_index()
            bc=[c for c in pv.columns if c not in("Código Producto","Nombre Producto")]
            rn={c:c.replace("Bodega ","").replace("BODEGA ","") for c in bc}
            pv=pv.rename(columns=rn); bcr=list(rn.values())
            pv=pv[pv[bcr].sum(axis=1)>0]
            for _bc in bcr: pv[_bc]=pv[_bc].fillna(0).astype(int)
            pv["TOTAL"]=pv[bcr].sum(axis=1).astype(int)

            # ── Tabla pivot via st.components (sticky real garantizado) ──
            C1W, C2W  = 110, 240
            num_cols  = bcr + ["TOTAL"]
            all_cols  = list(pv.columns)

            def _td(style, content=""):
                return f'<td style="{style}">{content}</td>'
            def _th(style, content=""):
                return f'<th style="{style}">{content}</th>'

            S_STICKY_TOP   = "position:sticky;top:0;z-index:{z};background:#f1f5f9;border-bottom:2px solid #cbd5e1;padding:7px 10px;font-size:10px;font-weight:700;text-transform:uppercase;color:#64748b;white-space:nowrap"
            S_STICKY_LEFT0 = "position:sticky;left:0;z-index:{z};background:{bg};border-right:1px solid #e2e8f0;border-bottom:1px solid #f1f5f9;padding:6px 10px;font-size:12px;white-space:nowrap;font-weight:600;font-family:monospace"
            S_STICKY_LEFT1 = f"position:sticky;left:{C1W}px;z-index:{{z}};background:{{bg}};border-right:2px solid #94a3b8;border-bottom:1px solid #f1f5f9;padding:6px 10px;font-size:12px"
            S_NUM          = "text-align:right;padding:6px 10px;border-bottom:1px solid #f1f5f9;font-family:monospace;font-size:12px;background:{bg}"
            S_TXT          = "padding:6px 10px;border-bottom:1px solid #f1f5f9;font-size:12px;background:{bg}"

            # ── Encabezados ────────────────────────────────────────
            hdrs = ""
            for i,c in enumerate(all_cols):
                if i == 0:
                    hdrs += _th(S_STICKY_TOP.format(z=5) + f";left:0;border-right:1px solid #cbd5e1;width:{C1W}px;min-width:{C1W}px", c)
                elif i == 1:
                    hdrs += _th(S_STICKY_TOP.format(z=5) + f";left:{C1W}px;border-right:2px solid #94a3b8;width:{C2W}px;min-width:{C2W}px", c)
                else:
                    align = "text-align:right;" if c in num_cols else ""
                    hdrs += _th(S_STICKY_TOP.format(z=3) + ";" + align, c)

            # ── Filas de datos ─────────────────────────────────────
            rows = ""; tots = {c:0 for c in num_cols}
            for ri,(_,row) in enumerate(pv.iterrows()):
                bg = "#f8fafc" if ri%2==0 else "#ffffff"
                cells = ""
                for i,c in enumerate(all_cols):
                    v = row[c]
                    if i == 0:
                        txt = str(v) if str(v) not in("nan","None","NaN") else ""
                        cells += _td(S_STICKY_LEFT0.format(z=2, bg=bg), txt)
                    elif i == 1:
                        txt = str(v) if str(v) not in("nan","None","NaN") else ""
                        cells += _td(S_STICKY_LEFT1.format(z=2, bg=bg), txt)
                    elif c in num_cols:
                        try:
                            iv = int(round(float(v))); tots[c] += iv
                            disp = f"{iv:,}" if iv != 0 else ""
                        except:
                            disp = ""
                        cells += _td(S_NUM.format(bg=bg), disp)
                    else:
                        cells += _td(S_TXT.format(bg=bg), str(v) if str(v) not in("nan","None","NaN") else "")
                rows += f"<tr>{cells}</tr>"

            # ── Fila TOTAL ─────────────────────────────────────────
            S_TOT = "background:#e0f2fe;font-weight:700;padding:7px 10px;border-top:2px solid #7dd3fc;color:#0369a1"
            tfooter = ""
            for i,c in enumerate(all_cols):
                if i == 0:
                    tfooter += _td(S_TOT + f";position:sticky;left:0;z-index:2", "TOTAL")
                elif i == 1:
                    tfooter += _td(S_TOT + f";position:sticky;left:{C1W}px;z-index:2;border-right:2px solid #94a3b8", "")
                elif c in num_cols:
                    tfooter += _td(S_TOT + ";text-align:right;font-family:monospace", f"{tots[c]:,}")
                else:
                    tfooter += _td(S_TOT, "")

            html_piv = (
                "<!DOCTYPE html><html><head><meta charset=\"UTF-8\">"
                "<style>"
                "*{box-sizing:border-box;margin:0;padding:0}"
                "body{font-family:Inter,Segoe UI,sans-serif;background:#f0f9ff}"
                ".wrap{border:1px solid #e2e8f0;border-radius:8px;overflow:auto;"
                "max-height:620px;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,.08)}"
                ".zb{display:flex;align-items:center;gap:6px;margin-bottom:6px}"
                ".zb span{font-size:11px;font-weight:700;color:#94a3b8}"
                ".zb .inf{font-size:10px;font-weight:400}"
                ".zb button{background:#fff;border:1px solid #e2e8f0;border-radius:5px;"
                "padding:2px 10px;cursor:pointer;font-weight:700;font-size:13px;color:#475569}"
                ".zb button:hover{background:#0ea5e9;color:#fff;border-color:#0ea5e9}"
                "table{border-collapse:separate;border-spacing:0;width:max-content}"
                "tr:hover td{background:#e0f2fe!important}"
                "</style></head><body>"
                "<div class=\"zb\">"
                "<span>ZOOM</span>"
                "<button onclick=\"var t=document.getElementById('pt');var s=parseFloat(t.style.fontSize||'12');t.style.fontSize=(s-1)+'px'\">−</button>"
                "<button onclick=\"var t=document.getElementById('pt');var s=parseFloat(t.style.fontSize||'12');t.style.fontSize=(s+1)+'px'\">+</button>"
                "<button onclick=\"document.getElementById('pt').style.fontSize='12px'\">↺</button>"
                f"<span class=\"inf\">{len(pv):,} SKUs &middot; {len(bcr)} bodegas</span>"
                "</div>"
                "<div class=\"wrap\">"
                f"<table id=\"pt\"><thead><tr>{hdrs}</tr></thead>"
                f"<tbody>{rows}</tbody>"
                f"<tfoot><tr>{tfooter}</tr></tfoot>"
                "</table></div>"
                "<script>"
                "var w=document.querySelector('.wrap');"
                "var t0x,t0y,t0sl,t0st;"
                "w.addEventListener('touchstart',function(e){if(e.touches.length!==2)return;"
                "t0x=(e.touches[0].pageX+e.touches[1].pageX)/2;"
                "t0y=(e.touches[0].pageY+e.touches[1].pageY)/2;"
                "t0sl=w.scrollLeft;t0st=w.scrollTop;},{passive:true});"
                "w.addEventListener('touchmove',function(e){if(e.touches.length!==2)return;"
                "e.preventDefault();"
                "var cx=(e.touches[0].pageX+e.touches[1].pageX)/2;"
                "var cy=(e.touches[0].pageY+e.touches[1].pageY)/2;"
                "w.scrollLeft=t0sl-(cx-t0x);w.scrollTop=t0st-(cy-t0y);"
                "},{passive:false});"
                "w.addEventListener('wheel',function(e){"
                "if(Math.abs(e.deltaX)>Math.abs(e.deltaY)){e.preventDefault();w.scrollLeft+=e.deltaX;}"
                "else if(e.shiftKey){e.preventDefault();w.scrollLeft+=e.deltaY;}"
                "},{passive:false});"
                "</script></body></html>"
            )

            _components.html(html_piv, height=700, scrolling=False)
            dl3(pv,"sku_x_bodega","piv")
        except Exception as e: st.error(str(e))

with T_PIV:
    _render_tab_piv()

# ══ TAB 4 MUESTRAS ══════════════════════════════════════════════
@_fragment
def _render_tab_sam():
    r = st.session_state.get("result")
    eng = st.session_state.engine
    if r is None: st.info("Ejecute el análisis.")
    else:
        _s1,_s2=st.tabs(["👥 Resumen por Cliente","📋 Movimientos TRA"])

        # ── Sub-tab 1: Resumen por Cliente ────────────────────────
        with _s1:
            df=r.samples_by_client.copy()
            if df.empty:
                st.warning("Sin muestras.")
            else:
                c1,c2=st.columns([3,1])
                flt=c1.text_input("🔍 Cliente","",key="sa_f",placeholder="Nombre...")
                pend=c2.checkbox("Solo pendientes",True,key="sa_p")
                df=pfilt(df,flt,cols=("Cliente",))
                if pend and "Stock en Cliente" in df.columns:
                    df=df[df["Stock en Cliente"]>0]
                m1,m2,m3=st.columns(3)
                m1.metric("Enviadas",  int(df.get("Entregadas",    pd.Series([0])).sum()))
                m2.metric("Devueltas", int(df.get("Devueltas",     pd.Series([0])).sum()))
                m3.metric("Saldo",     int(df.get("Stock en Cliente",pd.Series([0])).sum()))
                nc=[c for c in df.columns if c not in("Código Producto","Nombre Producto","Última Devolución")]
                _freeze=2 if "Código Producto" in df.columns else 1
                _comp_tbl(df,nc,"sam",freeze_cols=_freeze,height=500,title=f"{len(df):,} clientes")
                dl3(df,"muestras","sam")

        # ── Sub-tab 2: Movimientos TRA detallados ─────────────────
        with _s2:
            if eng.raw_df is None:
                st.info("Cargue un archivo.")
            else:
                # Filtrar solo TRA del raw_df
                _raw=eng.raw_df.copy()
                if excl_s: _raw=_raw[~_raw["Código Producto"].isin(excl_s)]
                _tra=_raw[_raw["Tipo"].fillna("").str.upper()=="TRA"].copy()

                if _tra.empty:
                    st.warning("No hay movimientos de transferencia (TRA).")
                else:
                    # Normalizar fecha
                    _tra["Fecha"]=pd.to_datetime(_tra["Fecha"],errors="coerce")
                    _tra["Fecha_str"]=_tra["Fecha"].dt.strftime("%d/%m/%Y").fillna("")

                    # Tipo de movimiento — SIN emoji para comparaciones limpias
                    from app.config import PRIMARY_WAREHOUSE as _PW
                    _tra["Mov"]=_tra.apply(
                        lambda r: "Enviada"  if r["Bodega Origen"]==_PW else
                                  "Devuelta" if r["Bodega Destino"]==_PW else
                                  "Interna", axis=1)
                    _tra["Bodega"]=_tra.apply(
                        lambda r: r["Bodega Destino"] if r["Bodega Origen"]==_PW
                                  else r["Bodega Origen"], axis=1)

                    # Filtros
                    _fc1,_fc2,_fc3=st.columns([3,3,2])
                    _bods_tra=["Todas"]+sorted(_tra["Bodega"].dropna().unique().tolist())
                    _flt_cli =_fc1.text_input("🔍 Bodega/Cliente","",key="tra_f",placeholder="Nombre...")
                    _flt_sku =_fc2.text_input("🔍 SKU","",key="tra_s",placeholder="Código o nombre...")
                    _mov_flt =_fc3.selectbox("Movimiento",["Todos","📤 Enviada","📥 Devuelta","↔ Interna"],key="tra_m")

                    _t=_tra.copy()
                    if _flt_cli:
                        _t=_t[_t["Bodega"].fillna("").str.upper().str.contains(_flt_cli.upper())]
                    if _flt_sku:
                        _mask=(_t["Código Producto"].fillna("").str.upper().str.contains(_flt_sku.upper()) |
                               _t["Nombre Producto"].fillna("").str.upper().str.contains(_flt_sku.upper()))
                        _t=_t[_mask]
                    if _mov_flt!="Todos":
                        _t=_t[_t["Mov"]==_mov_flt]

                    # Orden: Bodega → SKU → Fecha
                    _t=_t.sort_values(["Bodega","Código Producto","Fecha"],na_position="last")

                    # Columnas: SKU | Nombre | N°Reg | MOV | Fecha | Desc | Cant
                    _COLS=["Código Producto","Nombre Producto","N° Registro","Mov","Fecha","Descripción","Cantidad"]
                    # Anchos MÍNIMOS — el usuario puede expandir arrastrando
                    _CW={"Código Producto":90,"Nombre Producto":200,
                         "N° Registro":138,"Mov":90,"Fecha":82,
                         "Descripción":280,"Cantidad":62}
                    _BOD_BG="#1e3a5f"; _BOD_FG="#7dd3fc"
                    _SKU_BG="#dbeafe"; _SKU_FG="#1d4ed8"
                    _TOT_BG="#1e40af"; _TOT_FG="#ffffff"
                    _DAT_BG=["#ffffff","#f0f9ff"]

                    # Helper <td> — siempre nowrap + overflow ellipsis para textos largos
                    def _td_tra(val, bg, extra_style=""):
                        v = str(val) if str(val) not in ("nan","None","NaN","") else ""
                        return ("<td style='background:"+bg
                                +";border-bottom:1px solid #f1f5f9;padding:4px 8px;"
                                "font-size:12px;white-space:nowrap;overflow:hidden;"
                                "text-overflow:ellipsis;max-width:400px;"
                                +extra_style+"' title='"+v.replace("'","")+"'>"+v+"</td>")

                    # Iconos MOV — flecha verde ↑ Enviada, flecha azul ↓ Devuelta
                    _MOV_ICON={"Enviada":"<span style='color:#059669;font-weight:900'>&#8593;</span> Enviada",
                               "Devuelta":"<span style='color:#0284c7;font-weight:900'>&#8595;</span> Devuelta",
                               "Interna":"&#8596; Interna"}

                    # Headers con ancho mínimo + position:relative para resize handle
                    TH_BASE=("position:sticky;top:0;z-index:3;background:#1e3a5f;"
                             "border-bottom:2px solid #0ea5e9;padding:6px 8px;font-size:10px;"
                             "font-weight:700;text-transform:uppercase;color:#e0f2fe;"
                             "white-space:nowrap;position:relative;overflow:visible")
                    _hdrs_tra="".join(
                        "<th style='"+TH_BASE
                        +";min-width:"+str(_CW.get(c,80))+"px"
                        +(";text-align:right" if c=="Cantidad" else "")
                        +"'>"+c
                        # Handle de resize (div invisible en borde derecho)
                        +"<div style='position:absolute;top:0;right:0;bottom:0;width:5px;"
                         "cursor:col-resize;z-index:10' "
                         "onmousedown='startResize(event,this.parentElement)'></div>"
                        +"</th>"
                        for c in _COLS)

                    _rows_html=""; _grand_env=0; _grand_dev=0; _ri=0

                    for _bod in _t["Bodega"].dropna().unique():
                        _df_b=_t[_t["Bodega"]==_bod]
                        _rows_html+=("<tr><td colspan='"+str(len(_COLS))
                                     +"' style='background:"+_BOD_BG
                                     +";padding:6px 12px;font-size:12px;font-weight:800;color:"+_BOD_FG
                                     +";border-top:3px solid #0ea5e9;border-bottom:2px solid #0ea5e9'>"
                                     +"🏪 "+str(_bod)+"</td></tr>")
                        _bod_env=0; _bod_dev=0

                        for _sku in _df_b["Código Producto"].dropna().unique():
                            _df_s=_df_b[_df_b["Código Producto"]==_sku]
                            _sku_nom=str(_df_s["Nombre Producto"].iloc[0]) if len(_df_s)>0 else ""
                            _sku_env=0; _sku_dev=0

                            for __,_row in _df_s.iterrows():
                                _bg=_DAT_BG[_ri%2]; _ri+=1
                                _qty=int(float(_row.get("Cantidad",0))) if str(_row.get("Cantidad","")) not in("","nan","None") else 0
                                _mov=str(_row.get("Mov",""))
                                # Enviada = positivo (+), Devuelta = negativo (−)
                                if _mov=="Enviada":
                                    _sku_env+=_qty
                                    _qty_disp="+"+str(_qty)
                                    _qty_col="#065f46"  # verde oscuro
                                elif _mov=="Devuelta":
                                    _sku_dev+=_qty
                                    _qty_disp="-"+str(_qty)
                                    _qty_col="#b91c1c"  # rojo oscuro
                                else:
                                    _qty_disp=str(_qty)
                                    _qty_col="#374151"
                                _cells=(
                                    _td_tra(_row.get("Código Producto",""), _bg)
                                    +_td_tra(_sku_nom, _bg)
                                    +_td_tra(_row.get("Código",""), _bg, "font-family:monospace;font-size:11px")
                                    # MOV con icono de flecha coloreada
                                    +"<td style='background:"+_bg+";border-bottom:1px solid #f1f5f9;"
                                     "padding:4px 8px;font-size:12px;white-space:nowrap'>"
                                     +_MOV_ICON.get(_mov,_mov)+"</td>"
                                    +_td_tra(_row.get("Fecha_str",""), _bg, "white-space:nowrap")
                                    +_td_tra(str(_row.get("Descripción",""))[:80], _bg)
                                    +"<td style='background:"+_bg+";border-bottom:1px solid #f1f5f9;"
                                     "padding:4px 8px;text-align:right;font-family:monospace;"
                                     "font-weight:700;color:"+_qty_col+";white-space:nowrap'>"
                                     +_qty_disp+"</td>"
                                )
                                _rows_html+="<tr>"+_cells+"</tr>"

                            # Subtotal SKU: total = Enviadas - Devueltas (suma simple con signo)
                            _sku_neto=_sku_env-_sku_dev
                            _bod_env+=_sku_env; _bod_dev+=_sku_dev
                            _neto_col="#065f46" if _sku_neto>0 else ("#b91c1c" if _sku_neto<0 else "#374151")
                            _ss=("background:"+_SKU_BG+";border-top:1px solid #93c5fd;"
                                 "border-bottom:1px solid #93c5fd;padding:5px 8px;"
                                 "font-size:11px;font-weight:700;color:"+_SKU_FG+";")
                            _s_cells=(
                                "<td style='"+_ss+"'>"+_sku+"</td>"
                                +"<td style='"+_ss+"'>"+_sku_nom+"</td>"
                                +"<td style='"+_ss+"'>Subtotal</td>"
                                +"<td style='"+_ss+"'></td>"
                                +"<td style='"+_ss+"'></td>"
                                +"<td style='"+_ss+"'>+"+str(_sku_env)+" / -"+str(_sku_dev)+"</td>"
                                +"<td style='"+_ss+"text-align:right;font-family:monospace;color:"+_neto_col+"'>"
                                +str(_sku_neto)+"</td>"
                            )
                            _rows_html+="<tr>"+_s_cells+"</tr>"

                        # Total Bodega
                        _bod_neto=_bod_env-_bod_dev
                        _grand_env+=_bod_env; _grand_dev+=_bod_dev
                        _bneto_col="#065f46" if _bod_neto>0 else ("#b91c1c" if _bod_neto<0 else "#374151")
                        _bs=("background:"+_TOT_BG+";border-top:2px solid #3b82f6;"
                             "border-bottom:2px solid #3b82f6;padding:6px 8px;"
                             "font-size:12px;font-weight:800;color:"+_TOT_FG+";")
                        _b_cells=(
                            "<td style='"+_bs+"'>TOTAL "+str(_bod)+"</td>"
                            +"<td style='"+_bs+"'></td>"
                            +"<td style='"+_bs+"'></td>"
                            +"<td style='"+_bs+"'></td>"
                            +"<td style='"+_bs+"'></td>"
                            +"<td style='"+_bs+"'>+"+str(_bod_env)+" / -"+str(_bod_dev)+"</td>"
                            +"<td style='"+_bs+"text-align:right;font-family:monospace;color:"
                            +_bneto_col+"'>"+str(_bod_neto)+"</td>"
                        )
                        _rows_html+="<tr>"+_b_cells+"</tr>"

                    # Gran total
                    _grand_neto=_grand_env-_grand_dev
                    _gneto_col="#065f46" if _grand_neto>0 else ("#b91c1c" if _grand_neto<0 else "#374151")
                    _gs=("background:#0ea5e9;color:#fff;font-weight:800;"
                         "border-top:3px solid #0284c7;padding:7px 8px;")
                    _gt_tra=(
                        "<td style='"+_gs+"'>TOTAL GENERAL</td>"
                        +"<td style='"+_gs+"'></td>"
                        +"<td style='"+_gs+"'></td>"
                        +"<td style='"+_gs+"'></td>"
                        +"<td style='"+_gs+"'></td>"
                        +"<td style='"+_gs+"'>+"+str(_grand_env)+" / -"+str(_grand_dev)+"</td>"
                        +"<td style='"+_gs+"text-align:right;font-family:monospace;color:"
                        +_gneto_col+"'>"+str(_grand_neto)+"</td>"
                    )
                    _n_mov=str(len(_t))
                    _html_tra=(
                        "<!DOCTYPE html><html><head><meta charset='UTF-8'>"
                        "<style>*{box-sizing:border-box;margin:0;padding:0}"
                        "body{font-family:Inter,sans-serif;background:#f0f9ff;padding:2px}"
                        ".zb{display:flex;align-items:center;gap:6px;margin-bottom:6px}"
                        ".zb span{font-size:11px;font-weight:700;color:#94a3b8}"
                        ".zb .i{font-size:10px;font-weight:400}"
                        ".zb button{background:#fff;border:1px solid #e2e8f0;border-radius:5px;"
                        "padding:2px 10px;cursor:pointer;font-weight:700;font-size:13px;color:#475569}"
                        ".zb button:hover{background:#0ea5e9;color:#fff}"
                        ".wt{overflow:auto;max-height:620px;border:1px solid #e2e8f0;"
                        "border-radius:8px;background:#fff}"
                        "table{border-collapse:separate;border-spacing:0;"
                        "width:max-content;font-size:12px;table-layout:auto}"
                        "td{white-space:nowrap}"
                        "tr:hover td{filter:brightness(.95)}"
                        "</style></head><body>"
                        "<div class='zb'><span>ZOOM</span>"
                        "<button onclick='tZ(-1)'>&#8722;</button>"
                        "<button onclick='tZ(1)'>+</button>"
                        "<button onclick='tR()'>&#8635;</button>"
                        "<span class='i'>"+_n_mov+" movimientos</span></div>"
                        "<div class='wt'><table id='trt'>"
                        "<thead><tr>"+_hdrs_tra+"</tr></thead>"
                        "<tbody>"+_rows_html+"</tbody>"
                        "<tfoot><tr>"+_gt_tra+"</tr></tfoot>"
                        "</table></div>"
                        "<script>"
                        "function tZ(d){var t=document.getElementById('trt');"
                        "var s=parseFloat(t.style.fontSize||'12');"
                        "t.style.fontSize=Math.min(20,Math.max(8,s+d))+'px';}"
                        "function tR(){document.getElementById('trt').style.fontSize='12px';}"
                        "var _rTh=null,_rX=0,_rW=0;"
                        "function startResize(e,th){"
                        "e.preventDefault();e.stopPropagation();"
                        "_rTh=th;_rX=e.pageX;_rW=th.offsetWidth;"
                        "document.body.style.cursor='col-resize';"
                        "document.body.style.userSelect='none';}"
                        "document.addEventListener('mousemove',function(e){"
                        "if(!_rTh)return;"
                        "var w=Math.max(40,_rW+(e.pageX-_rX));"
                        "_rTh.style.width=w+'px';_rTh.style.minWidth=w+'px';});"
                        "document.addEventListener('mouseup',function(){"
                        "if(!_rTh)return;_rTh=null;"
                        "document.body.style.cursor='';"
                        "document.body.style.userSelect='';});"
                        "var w=document.querySelector('.wt');"
                        "w.addEventListener('wheel',function(e){"
                        "if(Math.abs(e.deltaX)>Math.abs(e.deltaY)){e.preventDefault();w.scrollLeft+=e.deltaX;}"
                        "else if(e.shiftKey){e.preventDefault();w.scrollLeft+=e.deltaY;}"
                        "},{passive:false});"
                        "</script></body></html>"
                    )
                    st.markdown(f"**{len(_t):,} movimientos** encontrados")
                    _components.html(_html_tra, height=730, scrolling=False)
                    _exp_c=["Bodega","Código Producto","Nombre Producto","Código","Fecha_str","Descripción","Mov","Cantidad"]
                    _exp_df=_t[[c for c in _exp_c if c in _t.columns]].rename(columns={"Fecha_str":"Fecha","Código":"N° Registro"})
                    dl3(_exp_df,"muestras_tra","tra_exp")

with T_SAM:
    _render_tab_sam()

# ══ TAB 5 ANÁLISIS PERÍODO ══════════════════════════════════════
@_fragment
def _render_tab_ana():
    r = st.session_state.get("result")
    eng = st.session_state.engine
    if eng.raw_df is None: st.info("Cargue un archivo.")
    else:
        c1,c2,c3=st.columns([2,2,1])
        df=date.today()-timedelta(days=365)
        d_f=c1.date_input("Desde",df,format="DD/MM/YYYY",key="an_f")
        d_t=c2.date_input("Hasta",date.today(),format="DD/MM/YYYY",key="an_t")
        with c3:
            st.markdown("")
            go=st.button("▶ Calcular",type="primary",key="an_btn")
        if go:
            if d_f>d_t: st.error("Fecha inicial > final.")
            else:
                with st.spinner("Calculando..."):
                    dfa=eng.raw_df.copy()
                    if excl_s: dfa=dfa[~dfa["Código Producto"].isin(excl_s)]
                    dfa=dfa[(dfa["Fecha"]>=pd.Timestamp(d_f))&(dfa["Fecha"]<=pd.Timestamp(d_t))]
                    ref=dfa["Referencia"].fillna("").astype(str).str.upper()
                    typ=dfa["Tipo"].fillna("").astype(str).str.upper()
                    vdf=dfa[(typ=="EGR")&ref.str.startswith("FAC")].copy()
                    cdf=dfa[(typ=="ING")&ref.str.startswith("FAC")].copy()
                    cpq,cpv,cpm=defaultdict(float),defaultdict(float),{}
                    for _,row in cdf.sort_values(["Código Producto","Fecha"]).iterrows():
                        sku=row["Código Producto"]; qty=float(row["Cantidad"]) or 1
                        vt=float(row["Valor Total"]); cu=vt/qty if qty>0 and vt>0 else 0
                        if cu>0:
                            nq=cpq[sku]+qty; nv=cpv[sku]+vt; cpq[sku]=nq; cpv[sku]=nv; cpm[sku]=nv/nq
                    st.session_state["an_v"]=vdf; st.session_state["an_cpm"]=cpm
        if "an_v" in st.session_state and not st.session_state["an_v"].empty:
            vdf=st.session_state["an_v"]; cpm=st.session_state["an_cpm"]
            s1,s2=st.tabs(["📊 Top 10 Vendidos","💰 Top 10 Rentabilidad"])
            with s1:
                t10v=(vdf.groupby(["Código Producto","Nombre Producto"])
                      .agg(Unidades=("Cantidad","sum"),Ventas=("Valor Total","sum"))
                      .reset_index().nlargest(10,"Ventas")
                      .sort_values("Ventas",ascending=False))
                if not t10v.empty:
                    a,b=st.columns([1,1])
                    with a: st.markdown(tbl(t10v,["Unidades","Ventas"],"av"),unsafe_allow_html=True)
                    # Gráfica también ordenada de mayor a menor
                    with b: st.bar_chart(t10v.sort_values("Ventas",ascending=False)
                                            .set_index("Código Producto")["Ventas"])
                    dl3(t10v,"top10_vendidos","av")
            with s2:
                rows_r=[]
                for (sku,nom),grp in vdf.groupby(["Código Producto","Nombre Producto"]):
                    v=float(grp["Valor Total"].sum())
                    ct=sum(float(rv["Cantidad"])*cpm.get(sku,0) for _,rv in grp.iterrows())
                    rent=v-ct
                    rows_r.append({"Código":sku,"Nombre":nom,"Ventas":round(v,2),"Costo":round(ct,2),"Rentabilidad":round(rent,2),"Margen%":round(rent/v*100,2) if v else 0})
                t10r=pd.DataFrame(rows_r).nlargest(10,"Rentabilidad")
                if not t10r.empty:
                    a,b=st.columns([1,1])
                    with a: st.markdown(tbl(t10r,["Ventas","Costo","Rentabilidad","Margen%"],"ar"),unsafe_allow_html=True)
                    with b: st.bar_chart(t10r.set_index("Código")["Rentabilidad"])
                    dl3(t10r,"top10_rentabilidad","ar")

with T_ANA:
    _render_tab_ana()

# ══ TAB 6 ROTACIÓN ══════════════════════════════════════════════
@_fragment
def _render_tab_rot():
    r = st.session_state.get("result")
    eng = st.session_state.engine
    if r is None: st.info("Ejecute el análisis.")
    else:
        c1,c2,c3,c4,c5=st.columns([2,2,2,2,3])
        lt_m=c1.number_input("🚢 Marítimo(d)",1,365,45,key="rm")
        lt_a=c2.number_input("✈ Aéreo(d)",1,180,15,key="ra")
        saf=c3.number_input("Stock Seg.(d)",0,180,15,key="rs")
        cob=c4.number_input("🎯 Cobertura(d)",1,730,60,key="rcob",
                            help="Días totales de cobertura deseados. Sug.Cob. "
                                 "= Consumo_diario × Cobertura − Stock_actual.")
        flt_r=c5.text_input("🔍 Filtrar SKU","",key="rf",placeholder="Código o nombre...")

        # ── Ventana para el cálculo del consumo ─────────────────────
        vc1, vc2, vc3 = st.columns([2, 2, 2])
        with vc1:
            ventana = st.selectbox(
                "📅 Ventana consumo",
                list(VENTANA_OPTS.keys()), index=1,
                key="rc_ventana",
                help="Período para calcular el consumo promedio. Si la empresa "
                     "ha crecido, usar Últimos 90d es más realista que Todo el período.",
            )
        _cus_s = _cus_e = None
        if ventana == "Personalizado":
            _dmin = r.filtered["Fecha"].min()
            _dmax = r.filtered["Fecha"].max()
            with vc2:
                _cus_s = st.date_input("Desde", value=(_dmax - pd.Timedelta(days=90)).date(),
                                        format="DD/MM/YYYY", key="rc_win_s")
            with vc3:
                _cus_e = st.date_input("Hasta", value=_dmax.date(),
                                        format="DD/MM/YYYY", key="rc_win_e")

        calc_r=st.button("▶ Calcular Rotación",type="primary",key="rc")
        if calc_r:
            sku=r.sku_summary.copy()
            if excl_s: sku=sku[~sku["Código Producto"].isin(excl_s)]

            # Ventas por SKU SOLO en la ventana seleccionada
            days, _ventas_win, _win_min, _win_max = _compute_window_sales(
                r, ventana, custom_start=_cus_s, custom_end=_cus_e)
            st.session_state["_rc_win_info"] = {
                "label": ventana, "days": days,
                "min": _win_min, "max": _win_max,
                "n_skus_con_ventas": len(_ventas_win),
            }
            def make_rot(lt):
                rows=[]
                for _,row in sku.iterrows():
                    cod=row["Código Producto"]; nom=row["Nombre Producto"]
                    stk=max(0.0,float(row.get("Stock Disponible",0)))
                    vtas=float(_ventas_win.get(cod, 0))
                    # Costo promedio y valores $
                    _comp_u = float(row.get("Compras", 0) or 0)
                    _val_comp = float(row.get("Valor Compras", row.get("Valor_Compras", 0)) or 0)
                    costo_prom = (_val_comp / _comp_u) if _comp_u > 0 else 0.0
                    val_stock  = stk * costo_prom
                    cogs_aprox = vtas * costo_prom
                    cons=vtas/days
                    rot_u=vtas/stk if stk>0 else(999 if vtas>0 else 0)
                    rot_d=cogs_aprox/val_stock if val_stock>0 else 0.0
                    dinv=stk/cons if cons>0 else(0 if stk==0 else 9999)
                    sug_lt=max(0.0,cons*(lt+saf)-stk)
                    sug_cob=max(0.0,cons*cob-stk)
                    pr=cons*lt
                    if cons>0 and stk==0: alrt="SIN STOCK"
                    elif cons>0 and rot_u>2 and dinv<lt: alrt="BEST-SELLER"
                    elif cons>0 and dinv<lt: alrt="CRÍTICO"
                    elif cons>0 and dinv<lt+saf: alrt="BAJO"
                    elif cons==0: alrt="SIN VENTA"
                    else: alrt="OK"
                    rows.append({"Código":cod,"Nombre":nom,"Stock":int(stk),"Ventas(u)":int(vtas),
                                 "Cons/día":round(cons,3),"P.Reorden":round(pr,1),
                                 "Días Inv.":round(dinv,1),
                                 "Rotación":round(rot_u,2),
                                 "Rotación $":round(rot_d,2),
                                 "Sug.LT+SS":int(round(sug_lt)),
                                 "Sug.Cob.":int(round(sug_cob)),
                                 "Estado":alrt})
                order={"SIN STOCK":0,"BEST-SELLER":1,"CRÍTICO":2,"BAJO":3,"OK":4,"SIN VENTA":5}
                rows.sort(key=lambda x:(order.get(x["Estado"],6),-x["Ventas(u)"]))
                return pd.DataFrame(rows)
            st.session_state["rot_m"]=make_rot(lt_m)
            st.session_state["rot_a"]=make_rot(lt_a)

        CK={"SIN STOCK":"#fca5a5","BEST-SELLER":"#fed7aa","CRÍTICO":"#fca5a5",
            "BAJO":"#fde68a","OK":"#bbf7d0","SIN VENTA":"#e2e8f0"}
        nr=["Stock","Ventas(u)","Cons/día","P.Reorden","Días Inv.",
            "Rotación","Rotación $","Sug.LT+SS","Sug.Cob."]

        def rot_table(df_r, uid):
            hdrs="".join(f"<th class='{'n' if c in nr else ''}'>{c}</th>" for c in df_r.columns)
            rows=""; tot_lt=0; tot_cob=0
            for _,row in df_r.iterrows():
                cells=""
                for c in df_r.columns:
                    v=row[c]; disp=str(v) if str(v) not in("nan","None","NaN") else ""
                    if c=="Estado":
                        bg=CK.get(v,"")
                        cells+=f"<td style='background:{bg};color:#111;font-weight:700'>{disp}</td>"
                    elif c in nr:
                        cells+=f"<td class='n'>{disp}</td>"
                        if c=="Sug.LT+SS":
                            try: tot_lt+=int(v)
                            except: pass
                        elif c=="Sug.Cob.":
                            try: tot_cob+=int(v)
                            except: pass
                    elif c=="Nombre":
                        # Truncar nombre largo con tooltip (ver completo al hover)
                        _full = disp.replace('"','&quot;').replace('<','&lt;').replace('>','&gt;')
                        _disp = disp if len(disp) <= 50 else disp[:47] + "…"
                        cells+=f'<td class="nom" title="{_full}">{_disp}</td>'
                    else: cells+=f"<td>{disp}</td>"
                rows+=f"<tr>{cells}</tr>"
            def _tot_cell(c, i):
                if c=="Sug.LT+SS": return f"<td class='n'>{tot_lt:,}</td>"
                if c=="Sug.Cob.":  return f"<td class='n'>{tot_cob:,}</td>"
                if i==0: return "<td><b>TOTAL</b></td>"
                return "<td></td>"
            tcells="".join(_tot_cell(c, i) for i,c in enumerate(df_r.columns))
            return f"""<div class="zb">
  <button onclick="var t=document.getElementById('t_{uid}');t.style.fontSize=Math.max(8,parseFloat(getComputedStyle(t).fontSize)-1)+'px'">−</button>
  <button onclick="var t=document.getElementById('t_{uid}');t.style.fontSize=Math.min(20,parseFloat(getComputedStyle(t).fontSize)+1)+'px'">+</button>
  <button onclick="document.getElementById('t_{uid}').style.fontSize='12px'">↺</button>
  <span style="color:{MUTED};font-size:10px">{len(df_r):,} SKUs | ⚠️ {len(df_r[df_r['Estado'].isin(['SIN STOCK','CRÍTICO','BEST-SELLER'])])} críticos</span>
</div>
<div class="tc rot-tc"><table class="it it-rot" id="t_{uid}">
<thead><tr>{hdrs}</tr></thead><tbody>{rows}</tbody>
<tfoot><tr class="tot">{tcells}</tr></tfoot>
</table></div>"""

        if "rot_m" in st.session_state:
            _wi = st.session_state.get("_rc_win_info", {})
            if _wi:
                st.info(
                    f"📅 **Ventana de consumo**: {_wi['label']} · "
                    f"{_wi['min'].strftime('%d/%m/%Y')} → {_wi['max'].strftime('%d/%m/%Y')} "
                    f"({_wi['days']} días) · "
                    f"{_wi['n_skus_con_ventas']:,} SKUs con ventas en la ventana."
                )
            s1,s2=st.tabs(["🚢 Marítimo","✈ Aéreo"])
            for tab_r,key,lbl in [(s1,"rot_m","Marítimo"),(s2,"rot_a","Aéreo")]:
                with tab_r:
                    df_r=pfilt(st.session_state[key],flt_r,cols=("Código","Nombre"))
                    st.markdown(rot_table(df_r,f"rot_{lbl}"),unsafe_allow_html=True)
                    with st.expander("ℹ Leyenda — qué significa cada columna y cada estado"):
                        st.markdown("""
**Columnas de la tabla**

| Columna | Qué es | Cómo se calcula |
|---|---|---|
| **Código / Nombre** | Identificación del producto | — |
| **Stock** | Unidades disponibles hoy | Del análisis global |
| **Ventas (u)** | Unidades vendidas en la ventana elegida | Σ EGR FAC en la ventana |
| **Cons/día** | Ritmo actual de consumo | Ventas ÷ Días de la ventana |
| **P. Reorden** | Nivel de stock al que debes disparar una compra para que llegue justo antes de agotarse | Cons/día × Lead time |
| **Días Inv.** | Cuántos días te alcanza el stock al ritmo actual | Stock ÷ Cons/día |
| **Rotación** | Veces que rotas el stock en la ventana | Ventas ÷ Stock |
| **Rotación $** | Lo mismo pero en dinero (indicador financiero) | COGS ÷ Valor_stock |
| **Sug. LT+SS** | Compra para cubrir Lead time + Stock Seguridad | max(0, Cons/día × (LT+SS) − Stock) |
| **Sug. Cob.** | Compra para alcanzar la Cobertura objetivo (días) | max(0, Cons/día × Cobertura − Stock) |
| **Estado** | Clasificación por urgencia (ver abajo) | — |

---

**Estados — por qué se asigna cada uno**

| Estado | Significado | Regla que lo activa |
|---|---|---|
| 🔴 **SIN STOCK** | Hubo ventas pero el stock quedó en 0. Pierdes ventas ahora mismo. | `Ventas > 0` **AND** `Stock = 0` |
| 🔥 **BEST-SELLER** | Producto de alta demanda en riesgo de quiebre. Urgente reponer. | `Rotación > 2×` **AND** `Días Inv. < Lead time` |
| 🔴 **CRÍTICO** | No alcanzas a sobrevivir el tiempo de entrega si pides hoy. | `Días Inv. < Lead time` |
| 🟡 **BAJO** | Cubres el lead time pero no el stock de seguridad. Pide pronto. | `Días Inv. < Lead time + Stock Seguridad` |
| 🟢 **OK** | Stock suficiente para cubrir lead time + seguridad. Tranquilo. | `Días Inv. ≥ Lead time + Stock Seguridad` |
| ⚪ **SIN VENTA** | No hubo ventas en la ventana. Puede ser estacional o descontinuado. | `Ventas = 0` en la ventana |

**Tip**: si un producto aparece como 🔥 BEST-SELLER, **prioriza el aéreo** aunque sea más caro — el costo del quiebre supera al flete extra.
                        """)
                    dl3(df_r,f"rotacion_{lbl.lower()}",f"rot_{lbl}")

with T_ROT:
    _render_tab_rot()

# ══ TAB: CÁLCULOS PASO A PASO ═══════════════════════════════════
@_fragment
def _render_tab_cal():
    r = st.session_state.get("result")
    eng = st.session_state.engine
    if r is None or eng.raw_df is None:
        st.info("Ejecute el análisis primero (sidebar → Calcular).")
        return

    st.markdown("### 📐 Cálculos de Rotación y Compras — paso a paso")
    st.caption(
        "Elige un SKU y ajusta los parámetros. La tabla muestra cómo se "
        "obtiene cada valor, fórmula por fórmula."
    )

    # Selector de SKU (predictivo con format_func nativo)
    _opts = (eng.raw_df[["Código Producto","Nombre Producto"]]
             .drop_duplicates().sort_values("Nombre Producto"))
    _codes = _opts["Código Producto"].astype(str).tolist()
    _names = _opts["Nombre Producto"].astype(str).tolist()
    _labels = [f"{c} — {n}" for c,n in zip(_codes, _names)]
    if not _labels:
        st.warning("No hay SKUs disponibles.")
        return

    cs1, cs2 = st.columns([3, 2])
    with cs1:
        sel_label = st.selectbox("🔍 SKU a analizar", _labels, key="cal_sku",
                                  placeholder="Escribe código o nombre…")
    sel_idx = _labels.index(sel_label) if sel_label in _labels else 0
    sel_code = _codes[sel_idx]; sel_name = _names[sel_idx]

    p1, p2, p3 = st.columns(3)
    with p1:
        lt = st.number_input("🚢/✈ Lead time (d)", 1, 365, value=45, key="cal_lt",
                              help="Tiempo en días desde que pides hasta que recibes.")
    with p2:
        saf = st.number_input("🛡 Stock seguridad (d)", 0, 180, value=15, key="cal_saf",
                               help="Buffer en días contra variabilidad o demora.")
    with p3:
        cob = st.number_input("🎯 Cobertura objetivo (d)", 1, 730, value=60, key="cal_cob",
                               help="Cuántos días de demanda quieres tener cubiertos.")

    # ── Ventana para el consumo ──
    vc1, vc2, vc3 = st.columns([2, 2, 2])
    with vc1:
        ventana = st.selectbox(
            "📅 Ventana consumo",
            list(VENTANA_OPTS.keys()), index=1,
            key="cal_ventana",
            help="Período para consumo promedio. 'Últimos 90 días' es estándar "
                 "en retail; evita diluir el ritmo actual con ventas antiguas.",
        )
    _cus_s = _cus_e = None
    if ventana == "Personalizado":
        _dmin = r.filtered["Fecha"].min()
        _dmax = r.filtered["Fecha"].max()
        with vc2:
            _cus_s = st.date_input("Desde",
                value=(_dmax - pd.Timedelta(days=90)).date(),
                format="DD/MM/YYYY", key="cal_win_s")
        with vc3:
            _cus_e = st.date_input("Hasta", value=_dmax.date(),
                format="DD/MM/YYYY", key="cal_win_e")

    # ── Datos del SKU ──
    sku_sel = r.sku_summary[r.sku_summary["Código Producto"] == sel_code]
    if sku_sel.empty:
        st.warning(
            f"El SKU **{sel_code}** no aparece en el análisis actual "
            "(puede estar excluido o sin movimientos en el período)."
        )
        return
    row = sku_sel.iloc[0]

    # Ventas y días calculados desde la ventana (NO del histórico completo)
    days, _ventas_win, _win_min, _win_max = _compute_window_sales(
        r, ventana, custom_start=_cus_s, custom_end=_cus_e)

    stk   = max(0.0, float(row.get("Stock Disponible", 0)))
    vtas  = float(_ventas_win.get(sel_code, 0))
    comp_u= float(row.get("Compras", 0) or 0)
    val_c = float(row.get("Valor Compras", row.get("Valor_Compras", 0)) or 0)
    costo = (val_c / comp_u) if comp_u > 0 else 0.0
    val_stk  = stk * costo
    cogs_apx = vtas * costo

    cons    = vtas / days if days else 0
    rot_u   = (vtas / stk) if stk > 0 else (999 if vtas > 0 else 0)
    rot_d   = (cogs_apx / val_stk) if val_stk > 0 else 0.0
    dinv_p  = (stk / cons) if cons > 0 else (0 if stk == 0 else 9999)
    dinv_a  = (365 / rot_u) if (0 < rot_u < 999) else (9999 if rot_u == 0 else 0)
    dem_lt  = cons * lt
    ss_u    = cons * saf
    reord   = dem_lt + ss_u
    nec_cob = cons * cob
    sug_cob = max(0.0, nec_cob - stk)
    sug_lt  = max(0.0, cons * (lt + saf) - stk)

    # Alerta
    if cons > 0 and stk == 0:
        alerta = "🔴 SIN STOCK"
    elif cons > 0 and rot_u > 2 and dinv_p < lt:
        alerta = "🔥 BEST-SELLER EN RIESGO"
    elif cons > 0 and dinv_p < lt:
        alerta = "🔴 CRÍTICO"
    elif cons > 0 and dinv_p < lt + saf:
        alerta = "🟡 BAJO"
    elif cons == 0:
        alerta = "⚪ SIN VENTA"
    else:
        alerta = "🟢 OK"

    st.markdown(f"#### {sel_name}  ·  `{sel_code}`")
    st.caption(
        f"📅 **Ventana de consumo**: {ventana} · "
        f"{_win_min.strftime('%d/%m/%Y')} → {_win_max.strftime('%d/%m/%Y')} ({days} días)"
    )

    # ── Tabla paso a paso ──
    pasos = [
        ("1",  "Días de la ventana",      "fecha_fin − fecha_inicio + 1",     f"{days:,} d"),
        ("2",  "Stock actual",             "del análisis (hoy)",              f"{int(stk):,} u"),
        ("3",  "Ventas en la ventana",     "Σ unidades EGR FAC en ventana",   f"{int(vtas):,} u"),
        ("4",  "Compras del período",      "Σ unidades ING FAC",              f"{int(comp_u):,} u"),
        ("5",  "Costo promedio unit.",     "Valor_Compras / Compras",         f"${costo:,.2f}"),
        ("6",  "Valor del stock",          "Stock × Costo_prom",              f"${val_stk:,.2f}"),
        ("7",  "COGS aprox.",              "Ventas_u × Costo_prom",           f"${cogs_apx:,.2f}"),
        ("8",  "Consumo diario",           "Ventas / Días_período",           f"{cons:.2f} u/día"),
        ("9",  "Rotación (unidades)",      "Ventas / Stock",                  f"{rot_u:.2f}×"),
        ("10", "Rotación ($)",             "COGS / Valor_stock",              f"{rot_d:.2f}×"),
        ("11", "Días inv. (período)",      "Stock / Consumo_diario",          f"{dinv_p:.1f} d"),
        ("12", "Días inv. (anualizado)",   "365 / Rotación_u",                f"{dinv_a:.1f} d"),
        ("13", "Demanda en Lead time",     f"Cons × LT ({lt}d)",              f"{dem_lt:.1f} u"),
        ("14", "Stock de seguridad",       f"Cons × SS ({saf}d)",             f"{ss_u:.1f} u"),
        ("15", "Punto de reorden",         "Demanda_LT + SS",                 f"{reord:.1f} u"),
        ("16", f"Necesidad para {cob}d",   "Cons × Cobertura",                f"{nec_cob:.1f} u"),
        ("17", "Sug. compra (cobertura)",  "max(0, Necesidad − Stock)",       f"{int(round(sug_cob)):,} u"),
        ("18", "Sug. compra (LT+SS)",      "max(0, Cons×(LT+SS) − Stock)",    f"{int(round(sug_lt)):,} u"),
        ("19", "Estado / Alerta",          "evaluación de reglas",            alerta),
    ]
    cal_df = pd.DataFrame(pasos, columns=["#", "Variable", "Fórmula", "Valor"])
    st.dataframe(cal_df, use_container_width=True, hide_index=True, height=560,
                 column_config={
                     "#":        st.column_config.TextColumn("Paso", width="small"),
                     "Variable": st.column_config.TextColumn("Variable"),
                     "Fórmula":  st.column_config.TextColumn("Fórmula"),
                     "Valor":    st.column_config.TextColumn("Valor", width="medium"),
                 })

    # ── Interpretación narrativa ──
    st.markdown("#### 📝 Interpretación")
    lines = []
    if cons == 0:
        lines.append("Este SKU **no tuvo ventas** en el período. Revisa si está "
                     "descontinuado, fuera de temporada o si el filtro de fechas excluye ventas.")
    else:
        lines.append(
            f"- Vendiste **{int(vtas):,} u** en **{days} días** → consumo diario "
            f"promedio **{cons:.2f} u/día**."
        )
        lines.append(
            f"- Stock actual: **{int(stk):,} u** (valor **${val_stk:,.2f}** al "
            f"costo promedio **${costo:,.2f}**)."
        )
        lines.append(
            f"- Al ritmo actual te alcanza para **{dinv_p:.0f} días**. "
            f"Rotación del período: **{rot_u:.2f}×** "
            f"(anualizado: 1 rotación cada {dinv_a:.0f}d)."
        )
        lines.append(
            f"- Demanda durante los **{lt} días de lead time**: {dem_lt:.0f} u. "
            f"Punto de reorden (incluye {saf}d de seguridad): **{reord:.0f} u**."
        )
        if sug_cob > 0:
            lines.append(
                f"- Para cubrir **{cob} días** necesitas **{int(round(sug_cob)):,} u** "
                f"(ya tienes {int(stk):,})."
            )
        else:
            lines.append(
                f"- Tu stock actual cubre **más de {cob} días**. No hace falta comprar."
            )
        if dinv_p < lt:
            lines.append(
                f"- ⚠ Tu lead time ({lt}d) es **mayor** a los días de stock que te quedan "
                f"({dinv_p:.0f}d). Si pides hoy **te quedarás sin stock** antes de que llegue."
            )
        elif dinv_p < lt + saf:
            lines.append(
                f"- ⚠ Estás apenas cubriendo lead time + seguridad. Pide **pronto**."
            )
    st.markdown("\n".join(lines))

    # Export del cálculo
    st.download_button("📥 Exportar cálculo a Excel",
        to_xl(cal_df),
        f"calculo_{sel_code}.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="cal_xl", use_container_width=True)

with T_CAL:
    _render_tab_cal()

# ══ TAB 7 COMPRAS ═══════════════════════════════════════════════
@_fragment
def _render_tab_pur():
    r = st.session_state.get("result")
    eng = st.session_state.engine
    if eng.raw_df is None: st.info("Cargue un archivo.")
    else:
        # Multiselect predictivo con filtro estricto
        _pu_opts = eng.raw_df[["Código Producto","Nombre Producto"]].drop_duplicates()
        _pu_opts = _pu_opts.sort_values("Nombre Producto")
        _pu_labels = [f"{c} — {n}" for c,n in zip(
            _pu_opts["Código Producto"].astype(str), _pu_opts["Nombre Producto"].astype(str))]
        _pu_sel = st.multiselect("🔍 SKU / Producto", _pu_labels, key="pu_f",
                                  placeholder="Escribe código o nombre…")
        df=eng.raw_df.copy()
        if excl_s: df=df[~df["Código Producto"].isin(excl_s)]
        ref=df["Referencia"].fillna("").astype(str).str.upper()
        typ=df["Tipo"].fillna("").astype(str).str.upper()
        cdf=df[(typ=="ING")&ref.str.startswith("FAC")].copy()
        if _pu_sel:
            _pu_codes_sel = {s.split(" — ")[0] for s in _pu_sel}
            cdf = cdf[cdf["Código Producto"].astype(str).isin(_pu_codes_sel)]
        if cdf.empty: st.warning("Sin compras.")
        else:
            cdf=cdf.sort_values(["Código Producto","Fecha"])
            rows=[]
            cpq,cpv=defaultdict(float),defaultdict(float)
            for _,row in cdf.iterrows():
                sku=row["Código Producto"]; qty=float(row["Cantidad"]) or 1
                vt=float(row["Valor Total"]); vu=vt/qty
                nq=cpq[sku]+qty; nv=cpv[sku]+vt; cpq[sku]=nq; cpv[sku]=nv; cp=nv/nq
                rows.append({"Fecha":row["Fecha"].strftime("%d/%m/%Y") if pd.notna(row["Fecha"]) else "",
                             "Factura":str(row.get("Referencia","")).strip(),"Código":sku,
                             "Nombre":str(row.get("Nombre Producto",""))[:50],
                             "Desc.":str(row.get("Descripción",""))[:40],
                             "Cant.":int(qty),"V.Total":round(vt,2),
                             "V.Unit":round(vu,4),"Costo Prom.":round(cp,4),"_g":sku})
            all_r=[]
            for sku,grp in pd.DataFrame(rows).groupby("_g",sort=False):
                for _,row in grp.iterrows(): all_r.append(row.drop("_g"))
                s=grp.iloc[-1].copy(); s["Fecha"]="SUBTOTAL"; s["Factura"]=""
                s["Cant."]=int(grp["Cant."].sum()); s["V.Total"]=round(grp["V.Total"].sum(),2)
                s["V.Unit"]=""; s["Desc."]=""; all_r.append(s.drop("_g"))
            out=pd.DataFrame(all_r)

            # ── Tabla compras con color alternado por grupo SKU ──
            _nc_pur=["Cant.","V.Total","V.Unit","Costo Prom."]
            # Paleta de 2 fondos alternados por SKU
            _pal=["#f0f9ff","#ffffff"]   # celeste muy suave / blanco
            _sub_bg="#dbeafe"            # azul pastel para filas SUBTOTAL

            def _pur_table(df, nc, uid):
                if df is None or df.empty:
                    return "<p>Sin datos</p>"
                uid_s = uid.replace("-", "_")
                all_cols = list(df.columns)
                PAL = ["#f0f9ff", "#ffffff"]
                SUB_BG = "#dbeafe"
                TH_BASE = (
                    "position:sticky;top:0;z-index:3;background:#f1f5f9;"
                    "border-bottom:2px solid #cbd5e1;padding:7px 10px;font-size:10px;"
                    "font-weight:700;text-transform:uppercase;color:#64748b;white-space:nowrap"
                )
                hdrs = "".join(
                    "<th style='" + TH_BASE + (";text-align:right" if c in nc else "") + "'>" + c + "</th>"
                    for c in all_cols
                )
                rows = ""
                tots = {c: 0.0 for c in nc}
                gi = 0
                last_sku = ""
                for _, row in df.iterrows():
                    sku = str(row.get("Código", ""))
                    is_sub = (str(row.get("Fecha", "")) == "SUBTOTAL")
                    if not is_sub and sku != last_sku and sku:
                        if last_sku:
                            gi += 1
                        last_sku = sku
                    bg = SUB_BG if is_sub else PAL[gi % 2]
                    cells = ""
                    for c in all_cols:
                        v = row[c]
                        raw = str(v)
                        disp = "" if raw in ("nan", "None", "NaN") else raw
                        if c in nc:
                            try:
                                fv = float(v)
                                if not is_sub:
                                    tots[c] += fv
                                disp = ("{:,}".format(int(round(fv))) if _is_int_col(c)
                                        else "{:,.2f}".format(fv))
                            except:
                                pass
                        fw  = "font-weight:700;" if is_sub else ""
                        aln = "text-align:right;font-family:monospace;" if c in nc else ""
                        bdr = "border-top:1px solid #93c5fd;" if is_sub else "border-bottom:1px solid #f1f5f9;"
                        cells += (
                            "<td style='background:" + bg + ";padding:6px 10px;font-size:12px;"
                            + aln + fw + bdr + "'>" + disp + "</td>"
                        )
                    rows += "<tr>" + cells + "</tr>"

                def _tf(c):
                    v = tots[c]
                    return ("{:,}".format(int(round(v))) if _is_int_col(c)
                            else "{:,.2f}".format(v))

                first = all_cols[0] if all_cols else ""
                TOT = (
                    "background:#e0f2fe;font-weight:800;padding:7px 10px;"
                    "border-top:2px solid #7dd3fc;color:#0369a1"
                )
                tcells = "".join(
                    (
                        "<td style='" + TOT + ";text-align:right;font-family:monospace'>" + _tf(c) + "</td>"
                        if (c in nc and tots[c] != 0) else
                        ("<td style='" + TOT + "'>TOTAL</td>" if c == first
                         else "<td style='" + TOT + "'></td>")
                    )
                    for c in all_cols
                )
                # Barra ZOOM — onclick usa comillas simples para JS
                on_m = "asZoom('" + uid_s + "',-1)"
                on_p = "asZoom('" + uid_s + "',1)"
                on_r = "asZoomReset('" + uid_s + "')"
                zoom = (
                    "<div class='zb'>"
                    "<span style='color:var(--text3);font-size:11px;font-weight:700'>ZOOM</span>"
                    "<button onclick='" + on_m + "'>&#8722;</button>"
                    "<button onclick='" + on_p + "'>+</button>"
                    "<button onclick='" + on_r + "'>&#8635;</button>"
                    "<span style='color:var(--text3);font-size:10px'>"
                    + str(len(df)) + " registros</span></div>"
                )
                return (
                    zoom
                    + "<div class='tc'><table class='it' id='tbl_" + uid_s + "'>"
                    + "<thead><tr>" + hdrs + "</tr></thead>"
                    + "<tbody>" + rows + "</tbody>"
                    + "<tfoot><tr>" + tcells + "</tr></tfoot>"
                    + "</table></div>"
                )

            # Wrapper HTML completo para sticky header en st.components
            _pur_html = _pur_table(out, _nc_pur, "pur")
            _pur_doc = (
                "<!DOCTYPE html><html><head><meta charset='UTF-8'>"
                "<link rel='preconnect' href='https://fonts.googleapis.com'>"
                "<style>"
                "*{box-sizing:border-box;margin:0;padding:0}"
                "body{font-family:Inter,Segoe UI,sans-serif;background:#f0f9ff;padding:2px}"
                ".tc{overflow:auto;max-height:560px;border:1px solid #e2e8f0;"
                "border-radius:8px;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,.08);"
                "position:relative}"
                ".it{width:100%;border-collapse:separate;border-spacing:0;font-size:12px}"
                ".it thead th{position:sticky!important;top:0!important;z-index:3!important;"
                "background:#f1f5f9;border-bottom:2px solid #cbd5e1;padding:7px 10px;"
                "font-size:10px;font-weight:700;text-transform:uppercase;color:#64748b;"
                "white-space:nowrap}"
                ".it .n{text-align:right;font-family:monospace}"
                ".zb{display:flex;align-items:center;gap:6px;margin-bottom:6px;font-family:Inter,sans-serif}"
                ".zb span{font-size:11px;font-weight:700;color:#94a3b8}"
                ".zb button{background:#fff;border:1px solid #e2e8f0;border-radius:5px;"
                "padding:2px 10px;cursor:pointer;font-weight:700;font-size:13px;color:#475569}"
                ".zb button:hover{background:#0ea5e9;color:#fff}"
                "tr:hover td{background:#e0f2fe!important}"
                "</style></head><body>"
                + _pur_html
                + "<script>"
                "var w=document.querySelector('.tc');"
                "if(w){w.addEventListener('wheel',function(e){"
                "if(Math.abs(e.deltaX)>Math.abs(e.deltaY)){e.preventDefault();w.scrollLeft+=e.deltaX;}"
                "else if(e.shiftKey){e.preventDefault();w.scrollLeft+=e.deltaY;}"
                "},{passive:false});}"
                "</script>"
                "</body></html>"
            )
            _components.html(_pur_doc, height=640, scrolling=False)
            dl3(out, "historico_compras", "pur")

with T_PUR:
    _render_tab_pur()

# ══ TAB 8 KARDEX ════════════════════════════════════════════════
@_fragment
def _render_tab_kdx():
    r = st.session_state.get("result")
    eng = st.session_state.engine
    if eng.raw_df is None: st.info("Cargue un archivo.")
    else:
        c1,c2,c3,c4=st.columns([2,2,3,1])
        kdf=date.today()-timedelta(days=365)
        kd_f=c1.date_input("Desde",kdf,format="DD/MM/YYYY",key="kf")
        kd_t=c2.date_input("Hasta",date.today(),format="DD/MM/YYYY",key="kt")
        # Selector predictivo con filtro estricto de substring
        _kdx_opts = (eng.raw_df[["Código Producto","Nombre Producto"]]
                     .drop_duplicates().sort_values("Nombre Producto"))
        _kdx_codes  = _kdx_opts["Código Producto"].astype(str).tolist()
        _kdx_names  = _kdx_opts["Nombre Producto"].astype(str).tolist()
        _kdx_labels = [f"{c} — {n}" for c,n in zip(_kdx_codes, _kdx_names)]
        _opts_kx = ["— Todos —"] + _kdx_labels
        kd_s_label = c3.selectbox("SKU", _opts_kx, key="ks")
        kd_s = "— Todos —" if kd_s_label == "— Todos —" else kd_s_label.split(" — ")[0]
        with c4:
            st.markdown("")
            gen=st.button("▶ Generar",type="primary",key="kg")
        if gen:
            with st.spinner("Calculando..."):
                raw=eng.raw_df.copy()
                sf="" if kd_s=="— Todos —" else kd_s
                if sf: raw=raw[raw["Código Producto"]==sf]
                df_ts=pd.Timestamp(kd_f); dt_ts=pd.Timestamp(kd_t)
                ref=raw["Referencia"].fillna("").astype(str).str.upper()
                typ=raw["Tipo"].fillna("").astype(str).str.upper()
                desc=raw["Descripción"].fillna("").astype(str).str.upper()
                # Baja de inventario (EGR + descripción contiene "BAJA DE INVENTARIO")
                raw["_ib"]=(typ=="EGR") & desc.str.contains("BAJA DE INVENTARIO", regex=False, na=False)
                raw["_ic"]=(typ=="ING")&ref.str.startswith("NCT")
                # Compras: ref FAC o "FACTURA DE COMPRA" en desc (ref vacía)
                raw["_ip"]=(typ=="ING") & ~raw["_ic"] & (
                    ref.str.startswith("FAC") |
                    desc.str.contains("FACTURA DE COMPRA", regex=False, na=False)
                )
                # Excluir bajas de venta y dev. proveedor (prioridad)
                raw["_is"]=(typ=="EGR")&ref.str.startswith("FAC") & ~raw["_ib"]
                raw["_ir"]=(typ=="EGR")&ref.str.startswith("NCT") & ~raw["_ib"]
                raw["_it"]=typ=="TRA"
                raw=raw.sort_values(["Código Producto","Fecha"]).reset_index(drop=True)
                hist=raw[raw["Fecha"]<df_ts]
                cp={};cpq={};cpv={};stk={}
                for _,row in hist.iterrows():
                    sku=row["Código Producto"]; qty=float(row["Cantidad"])
                    stk[sku]=stk.get(sku,0)
                    if row["_ip"] or row["_ic"]: stk[sku]+=qty
                    elif row["_is"] or row["_ir"] or row["_ib"]: stk[sku]-=qty
                    if row["_ip"] and qty>0:
                        vt=float(row["Valor Total"]); cu=vt/qty if qty>0 else 0
                        if cu>0:
                            nq=cpq.get(sku,0)+qty; nv=cpv.get(sku,0)+vt
                            cpq[sku]=nq; cpv[sku]=nv; cp[sku]=nv/nq
                rows=[]
                for sku,grp in raw.groupby("Código Producto"):
                    nom=str(grp["Nombre Producto"].iloc[0])
                    s0=stk.get(sku,0); c0=cp.get(sku,0)
                    rows.append({"Fecha":kd_f.strftime("%d/%m/%Y"),"Código":sku,"Nombre":nom,
                                 "N°Reg":"","Referencia":"—","Descripción":"SALDO INICIAL","Tipo":"INICIO",
                                 "Cantidad":round(s0,2),"V.Unit":round(c0,4),
                                 "Costo Prom.":round(c0,4),"Saldo":round(s0,2),"Valor Inv.":round(s0*c0,2)})
                    cc=cp.get(sku,0); cq=cpq.get(sku,0); cv=cpv.get(sku,0); saldo=s0
                    for _,row in grp[(grp["Fecha"]>=df_ts)&(grp["Fecha"]<=dt_ts)].iterrows():
                        qty=float(row["Cantidad"]); vt=float(row["Valor Total"]); vu=vt/qty if qty>0 else 0
                        fd=row["Fecha"].strftime("%d/%m/%Y") if pd.notna(row["Fecha"]) else ""
                        if row["_ip"]: tipo="INGRESO"; ef=+qty; nq=cq+qty; nv=cv+vt; cc=nv/nq; cq=nq; cv=nv
                        elif row["_ic"]: tipo="ING DEV.CLI"; ef=+qty
                        elif row["_is"]: tipo="EGRESO"; ef=-qty
                        elif row["_ir"]: tipo="EGR DEV.PROV"; ef=-qty
                        elif row["_ib"]: tipo="BAJA INV."; ef=-qty
                        elif row["_it"]: tipo="TRANSFERENCIA"; ef=0
                        else: tipo="OTRO"; ef=0
                        saldo+=ef
                        # N° registro del movimiento (col "Código" del Excel)
                        n_reg=str(row.get("Código","")).strip()
                        ref_val=str(row.get("Referencia","")).strip()
                        # Para transferencias: mostrar N°Reg prominente en Referencia
                        if tipo=="TRANSFERENCIA" and n_reg:
                            ref_disp="[Reg:" + n_reg + "]"
                        else:
                            ref_disp=ref_val
                        rows.append({"Fecha":fd,"Código":sku,"Nombre":nom,
                                     "N°Reg":n_reg,
                                     "Referencia":ref_disp,
                                     "Descripción":str(row.get("Descripción","")).strip(),
                                     "Tipo":tipo,"Cantidad":round(abs(qty),2),
                                     "V.Unit":round(vu,4),"Costo Prom.":round(cc,4),
                                     "Saldo":round(saldo,2),"Valor Inv.":round(saldo*cc,2)})
                dk=pd.DataFrame(rows)
                st.session_state["kdx"]=dk; st.session_state["kdx_s"]=sf
                log(f"Kardex {sf or 'todos'}: {len(dk):,} filas")

        if "kdx" in st.session_state and not st.session_state["kdx"].empty:
            dk=st.session_state["kdx"]
            CK2={
                "INICIO":     ("#dbeafe","#1e40af"),
                "INGRESO":    ("#d1fae5","#065f46"),
                "ING DEV.CLI":("#fef9c3","#854d0e"),
                "EGRESO":     ("#f3f4f6","#374151"),
                "EGR DEV.PROV":("#f3f4f6","#374151"),
                "BAJA INV.":  ("#fee2e2","#991b1b"),
                "TRANSFERENCIA":("#ede9fe","#5b21b6"),
            }
            nk=["Cantidad","V.Unit","Costo Prom.","Saldo","Valor Inv."]
            all_cols=list(dk.columns)

            TH_CSS = (
                "position:sticky;top:0;z-index:3;background:#1e3a5f;"
                "border-bottom:2px solid #0ea5e9;padding:7px 10px;font-size:10px;"
                "font-weight:700;text-transform:uppercase;color:#e0f2fe;white-space:nowrap"
            )
            hdrs_parts = []
            for c in all_cols:
                align = ";text-align:right" if c in nk else ""
                hdrs_parts.append("<th style='" + TH_CSS + align + "'>" + c + "</th>")
            hdrs = "".join(hdrs_parts)

            # ── Filas con separadores entre SKUs ───────────────────
            rows_parts = []
            last_sku = ""
            sku_col = "Código" if "Código" in all_cols else ("Código Producto" if "Código Producto" in all_cols else "")
            for _, row in dk.iterrows():
                sku = str(row.get(sku_col, "")) if sku_col else ""
                tp  = str(row.get("Tipo", ""))
                bg, fg = CK2.get(tp, ("#ffffff","#111827"))
                if sku and sku != last_sku and last_sku:
                    # Fila separadora pronunciada al cambiar de SKU
                    rows_parts.append(
                        "<tr><td colspan='" + str(len(all_cols)) + "' "
                        "style='background:#1e3a5f;padding:5px 12px;"
                        "font-size:11px;font-weight:700;color:#7dd3fc;"
                        "letter-spacing:.06em;border-top:3px solid #0ea5e9;"
                        "border-bottom:2px solid #0ea5e9'>"
                        "&#9644;&#9644; " + sku + " &#9644;&#9644;"
                        "</td></tr>"
                    )
                last_sku = sku
                cells_parts = []
                for c in all_cols:
                    v    = row[c]
                    raw  = str(v)
                    disp = "" if raw in ("nan","None","NaN") else raw
                    if c in nk:
                        try:
                            fv = float(v)
                            disp = ("{:,.4f}".format(fv) if c in ("V.Unit","Costo Prom.")
                                    else "{:,.2f}".format(fv))
                        except:
                            pass
                        cells_parts.append(
                            "<td style='background:" + bg + ";color:" + fg + ";"
                            "text-align:right;font-family:monospace;padding:6px 10px;"
                            "border-bottom:1px solid #f1f5f9'>" + disp + "</td>"
                        )
                    else:
                        cells_parts.append(
                            "<td style='background:" + bg + ";color:" + fg + ";"
                            "padding:6px 10px;border-bottom:1px solid #f1f5f9'>" + disp + "</td>"
                        )
                rows_parts.append("<tr>" + "".join(cells_parts) + "</tr>")
            rows = "".join(rows_parts)

            # ── Leyenda ────────────────────────────────────────────
            leg_parts = []
            for t,(b,f) in CK2.items():
                leg_parts.append(
                    "<span style='background:" + b + ";color:" + f + ";"
                    "padding:2px 8px;border-radius:4px;font-size:10px;"
                    "border:1px solid #d1d5db;margin-right:4px'>"
                    "&#9632; " + t + "</span>"
                )
            leg = "".join(leg_parts)

            n_mov = str(len(dk))
            kdx_html = (
                "<!DOCTYPE html><html><head><meta charset='UTF-8'>"
                "<style>"
                "*{box-sizing:border-box;margin:0;padding:0}"
                "body{font-family:Inter,Segoe UI,sans-serif;background:#f0f9ff;padding:2px}"
                ".leg{display:flex;flex-wrap:wrap;gap:4px;margin-bottom:8px}"
                ".zb{display:flex;align-items:center;gap:6px;margin-bottom:6px}"
                ".zb span{font-size:11px;font-weight:700;color:#94a3b8}"
                ".zb .inf{font-size:10px;font-weight:400}"
                ".zb button{background:#fff;border:1px solid #e2e8f0;border-radius:5px;"
                "padding:2px 10px;cursor:pointer;font-weight:700;font-size:13px;color:#475569}"
                ".zb button:hover{background:#0ea5e9;color:#fff}"
                ".wrap{overflow:auto;max-height:580px;border:1px solid #e2e8f0;"
                "border-radius:8px;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,.08)}"
                "table{border-collapse:separate;border-spacing:0;"
                "width:max-content;font-size:12px}"
                "tr:hover td{filter:brightness(.93)}"
                "</style></head><body>"
                "<div class='leg'>" + leg + "</div>"
                "<div class='zb'>"
                "<span>ZOOM</span>"
                "<button onclick=\"var t=document.getElementById('kdt');"
                "var s=parseFloat(t.style.fontSize||'12');"
                "t.style.fontSize=(s-1)+'px'\">&#8722;</button>"
                "<button onclick=\"var t=document.getElementById('kdt');"
                "var s=parseFloat(t.style.fontSize||'12');"
                "t.style.fontSize=(s+1)+'px'\">+</button>"
                "<button onclick=\"document.getElementById('kdt').style.fontSize='12px'\">&#8635;</button>"
                "<span class='inf'>" + n_mov + " movimientos</span>"
                "</div>"
                "<div class='wrap'>"
                "<table id='kdt'>"
                "<thead><tr>" + hdrs + "</tr></thead>"
                "<tbody>" + rows + "</tbody>"
                "</table></div>"
                "<script>"
                "var w=document.querySelector('.wrap');"
                "w.addEventListener('wheel',function(e){"
                "if(Math.abs(e.deltaX)>Math.abs(e.deltaY)){"
                "e.preventDefault();w.scrollLeft+=e.deltaX;}"
                "else if(e.shiftKey){e.preventDefault();w.scrollLeft+=e.deltaY;}"
                "},{passive:false});"
                "</script>"
                "</body></html>"
            )
            _components.html(kdx_html, height=720, scrolling=False)
            sfn=st.session_state["kdx_s"].replace("/","_") or "todos"
            dl3(dk,f"kardex_{sfn}","kdx")

with T_KDX:
    _render_tab_kdx()

# ── Fragment: Toma Física por Ubicación (tabla tipo planilla) ───
# Aislado en @st.fragment: las ediciones de celdas solo re-ejecutan esta
# función, no toda la app.
def _build_toma_table(eng, rap_df, ubicacion):
    """Construye DataFrame con TODOS los SKUs + su cantidad anterior para la ubicación dada."""
    sk = (eng.raw_df[["Código Producto","Nombre Producto"]]
          .drop_duplicates()
          .sort_values("Nombre Producto")
          .reset_index(drop=True)).copy()
    sk["Código Producto"] = sk["Código Producto"].astype(str)
    sk["Nombre Producto"]  = sk["Nombre Producto"].astype(str)
    prev_map = {}
    if rap_df is not None and not rap_df.empty:
        pv = rap_df[rap_df["Ubicación"].astype(str) == str(ubicacion)].copy()
        if not pv.empty:
            pv = pv.sort_values("Fecha").drop_duplicates(
                subset=["Código Producto"], keep="last")
            prev_map = dict(zip(pv["Código Producto"].astype(str),
                                pd.to_numeric(pv["Cantidad Física"], errors="coerce")))
    sk["Anterior"]    = sk["Código Producto"].map(prev_map)
    sk["Nueva"]       = pd.NA
    sk["Observación"] = ""
    return sk[["Código Producto","Nombre Producto","Anterior","Nueva","Observación"]]

@_fragment
def _render_toma_fragment():
    _frag_t0 = _ptime.perf_counter()
    _eng = st.session_state.engine
    if _eng.raw_df is None:
        st.info("Primero carga los Excel de movimientos en el sidebar.")
        return
    rap_state = _get_shared_rapid()
    rap_df    = rap_state["df"]

    all_ubic = _get_all_ubic()

    # Selector + agregar ubicación
    col_u, col_add = st.columns([3, 1])
    with col_u:
        sel_ubic = st.selectbox("📍 Ubicación donde se hará la toma",
                                all_ubic, key="tu_ubic")
    with col_add:
        st.markdown("")
        with st.popover("➕ Nueva ubicación", use_container_width=True):
            _new_ub = st.text_input("Nombre de la nueva ubicación",
                                     key="tu_new_ubic_name",
                                     placeholder="Ej: Bodega Norte")
            if st.button("Agregar", key="tu_add_ubic", type="primary"):
                _n = (_new_ub or "").strip()
                if not _n:
                    st.error("Escribe un nombre.")
                elif _n in all_ubic:
                    st.warning(f"«Ya existe la ubicación {_n}».")
                else:
                    custom = _get_custom_ubic()
                    custom["list"].append(_n)
                    _persist_custom_ubic(custom["list"])
                    log(f"Nueva ubicación agregada: {_n}")
                    st.success(f"✓ Agregada: {_n}")
                    _rerun_frag()

    # Construir la tabla base para esta ubicación
    table_df = _build_toma_table(_eng, rap_df, sel_ubic)

    st.caption(
        "Escribe la cantidad en la columna **Nueva**. Tab avanza a Observación; "
        "Tab otra vez baja a la siguiente fila. Flechas para navegar libremente. "
        "Columna **Anterior** (gris) es la última toma registrada y no es editable."
    )

    # Key dinámica por ubicación: al cambiar de ubicación, el editor reinicia
    editor_key = f"tu_ed_{sel_ubic}"
    edited = st.data_editor(
        table_df,
        use_container_width=True,
        num_rows="fixed",
        hide_index=True,
        column_config={
            "Código Producto": st.column_config.TextColumn("SKU", width="small", disabled=True),
            "Nombre Producto":  st.column_config.TextColumn("Producto", disabled=True),
            "Anterior":         st.column_config.NumberColumn("Anterior", format="%d",
                                    disabled=True,
                                    help="Última toma registrada para este SKU en esta ubicación."),
            "Nueva":            st.column_config.NumberColumn("Nueva cantidad",
                                    format="%d", min_value=0,
                                    help="Escribe la cantidad contada. Vacío = pendiente."),
            "Observación":      st.column_config.TextColumn("Observación",
                                    help="Opcional. Ej: dñado, en caja master..."),
        },
        key=editor_key,
        height=520,
    )

    # Progreso
    total    = len(edited)
    contados = int(edited["Nueva"].notna().sum())
    faltan   = total - contados
    pct      = (contados / total) if total else 0.0
    st.progress(pct,
        text=f"✅ Contados: {contados:,}  |  ⏳ Pendientes: {faltan:,}  "
             f"|  {pct*100:.0f}% de {total:,}")

    # Guardar / Descartar / Borrar
    cg1, cg2, cg3 = st.columns([2,1,1])
    with cg1:
        if st.button(f"💾 Guardar toma de «{sel_ubic}»",
                     type="primary", use_container_width=True,
                     key=f"tu_save_{sel_ubic}"):
            to_save = edited[edited["Nueva"].notna()].copy()
            if to_save.empty:
                st.warning("No hay cantidades registradas para guardar.")
            else:
                _now = datetime.now().strftime("%Y-%m-%d %H:%M")
                _rows = []
                for _, row in to_save.iterrows():
                    _obs = row.get("Observación","")
                    _rows.append({
                        "Fecha": _now,
                        "Ubicación": sel_ubic,
                        "Código Producto": str(row["Código Producto"]),
                        "Nombre Producto":  str(row["Nombre Producto"]),
                        "Cantidad Física": float(row["Nueva"]),
                        "Observación": _obs if pd.notna(_obs) else "",
                    })
                new_df = pd.concat([rap_df, pd.DataFrame(_rows)], ignore_index=True)
                rap_state["df"] = new_df
                _persist_rapid(new_df)
                log(f"Toma «{sel_ubic}»: {len(_rows)} items guardados")
                # Limpiar estado del editor para esta ubicación
                st.session_state.pop(editor_key, None)
                st.success(f"✓ Guardados {len(_rows)} items en «{sel_ubic}». "
                           f"Cambia de ubicación arriba para continuar.")
                _rerun_frag()
    with cg2:
        if st.button("↺ Descartar edición", use_container_width=True,
                     key=f"tu_reset_{sel_ubic}",
                     help="Limpia los cambios sin guardar en la tabla de arriba."):
            st.session_state.pop(editor_key, None)
            _rerun_frag()
    with cg3:
        # Eliminar TODA la toma guardada de esta ubicación — acción destructiva
        _n_saved = int((rap_df["Ubicación"].astype(str) == str(sel_ubic)).sum()) \
                   if not rap_df.empty else 0
        with st.popover(f"🔥 Borrar toma ({_n_saved})",
                        use_container_width=True,
                        disabled=(_n_saved == 0),
                        help=("No hay tomas guardadas en esta ubicación"
                              if _n_saved == 0 else
                              f"Elimina todas las {_n_saved} filas guardadas en «{sel_ubic}»")):
            st.warning(
                f"⚠ Vas a borrar **TODAS** las tomas guardadas en "
                f"**«{sel_ubic}»** ({_n_saved} registro(s)). Acción **irreversible**. "
                f"El resto de ubicaciones no se tocará."
            )
            if st.button("✔ Confirmar eliminación",
                         type="primary",
                         key=f"tu_del_confirm_{sel_ubic}"):
                new_df = rap_df[rap_df["Ubicación"].astype(str) != str(sel_ubic)].copy()
                rap_state["df"] = new_df
                _persist_rapid(new_df)
                log(f"Toma «{sel_ubic}» eliminada: {_n_saved} registro(s) borrados")
                st.session_state.pop(editor_key, None)
                st.success(f"✓ Toma de «{sel_ubic}» eliminada ({_n_saved} registros).")
                _rerun_frag()

    # Log fragment timing
    _frag_ms = (_ptime.perf_counter() - _frag_t0) * 1000
    try:
        _fts = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        with open(_PERF_LOG_PATH, "a", encoding="utf-8") as _ff:
            _ff.write(f"{_fts} FRAGMENT=toma TOTAL={_frag_ms:.0f}ms\n")
        _fh = _get_perf_history()
        _fh["runs"].append({"ts": _fts, "total_ms": round(_frag_ms),
                            "ck": [("fragment_toma", round(_frag_ms))]})
        _fh["runs"] = _fh["runs"][-30:]
    except Exception: pass


# ── Fragment: Resumen de Tomas Físicas (pivot SKU × Ubicación) ──
@_fragment
def _render_resumen_fragment():
    rap_state = _get_shared_rapid()
    rap_df    = rap_state["df"]

    if rap_df is None or rap_df.empty:
        st.info("Aún no hay tomas registradas. Ve a «⚡ Toma» para empezar.")
        return

    # Para cada (SKU, Ubicación) conservar sólo la última toma
    latest = rap_df.sort_values("Fecha").drop_duplicates(
        subset=["Código Producto","Ubicación"], keep="last")

    # Pivot: filas=SKU, columnas=Ubicación, valores=Cantidad
    pivot = latest.pivot_table(
        index=["Código Producto","Nombre Producto"],
        columns="Ubicación",
        values="Cantidad Física",
        aggfunc="sum",
    )
    # Ordenar columnas: primero DEFAULT_LOCATIONS (las que existan), luego resto alfabético
    _cols_used = list(pivot.columns)
    _ordered = [u for u in DEFAULT_LOCATIONS if u in _cols_used]
    _ordered += sorted([u for u in _cols_used if u not in _ordered])
    pivot = pivot[_ordered]

    # Total por fila (tratando NaN como 0 para la suma, sin alterar display)
    pivot_total = pivot.fillna(0).sum(axis=1).astype(int)
    pivot_display = pivot.reset_index()
    pivot_display["Total"] = pivot_total.values

    # KPIs arriba
    k1,k2,k3,k4 = st.columns(4)
    k1.metric("SKUs contados", f"{len(pivot_display):,}")
    k2.metric("Ubicaciones", f"{len(_ordered)}")
    k3.metric("Unidades totales", f"{int(pivot_total.sum()):,}")
    k4.metric("Registros", f"{len(rap_df):,}")

    # Configuración de columnas para display (números con formato)
    _col_cfg = {
        "Código Producto": st.column_config.TextColumn("SKU", width="small"),
        "Nombre Producto":  st.column_config.TextColumn("Producto"),
        "Total":            st.column_config.NumberColumn("Σ Total", format="%d"),
    }
    for _c in _ordered:
        _col_cfg[_c] = st.column_config.NumberColumn(_c, format="%d",
                         help=f"Cantidad contada en {_c}")

    st.dataframe(pivot_display, use_container_width=True, hide_index=True,
                 column_config=_col_cfg, height=560)

    # Export
    st.markdown("#### 📥 Exportar resumen")
    ec1, ec2 = st.columns(2)
    with ec1:
        st.download_button("📊 Excel",
            to_xl(pivot_display),
            "resumen_toma_fisica.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, key="res_xl")
    with ec2:
        _pdf = to_pdf(pivot_display, "Resumen Toma Física")
        if _pdf:
            st.download_button("📄 PDF",
                _pdf, "resumen_toma_fisica.pdf",
                "application/pdf",
                use_container_width=True, key="res_pdf")
        else:
            st.button("📄 PDF", disabled=True, use_container_width=True,
                     key="res_pdf_disabled")

    with st.expander("🧾 Historial completo de movimientos"):
        _hist = rap_df.sort_values("Fecha", ascending=False).reset_index(drop=True)
        # Tipos consistentes
        for _tc in ["Fecha","Ubicación","Código Producto","Nombre Producto","Observación"]:
            if _tc in _hist.columns:
                _hist[_tc] = _hist[_tc].fillna("").astype(str)
        st.dataframe(_hist, use_container_width=True, hide_index=True, height=320)
        st.download_button("📥 Exportar historial",
            to_xl(_hist), "historial_toma_fisica.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, key="res_hist_xl")

    # ── Administración: borrar tomas guardadas ──────────────────
    with st.expander("🧹 Administrar datos guardados — borrar tomas"):
        st.caption("Para borrar registros guardados (ej: pruebas). Esta acción **NO** se puede deshacer.")

        _all_ubic_saved = sorted(rap_df["Ubicación"].astype(str).unique().tolist())

        # A) Borrar por ubicación(es)
        st.markdown("##### 🗂 Borrar por ubicación(es)")
        _to_del = st.multiselect(
            "Ubicaciones a borrar",
            _all_ubic_saved,
            key="res_del_ubic",
            placeholder="Elige una o más ubicaciones…",
        )
        if _to_del:
            _n_afect = int(rap_df["Ubicación"].astype(str).isin(_to_del).sum())
            with st.popover(f"🔥 Borrar {len(_to_del)} ubicación(es) ({_n_afect} registro(s))",
                            use_container_width=True):
                st.warning(
                    f"⚠ Se borrarán **{_n_afect}** registro(s) de "
                    f"**{len(_to_del)}** ubicación(es): "
                    + ", ".join(f"«{u}»" for u in _to_del)
                )
                if st.button("✔ Confirmar", type="primary", key="res_del_sel_confirm"):
                    new_df = rap_df[~rap_df["Ubicación"].astype(str).isin(_to_del)].copy()
                    rap_state["df"] = new_df
                    _persist_rapid(new_df)
                    log(f"Resumen: borrados {_n_afect} registro(s) "
                        f"de {len(_to_del)} ubicación(es)")
                    # Limpiar editores activos de estas ubicaciones
                    for _u in _to_del:
                        st.session_state.pop(f"tu_ed_{_u}", None)
                    st.success(f"✓ {_n_afect} registro(s) eliminados.")
                    _rerun_frag()

        st.divider()

        # B) Borrar TODO
        st.markdown("##### 🔥 Borrar TODO el historial")
        st.caption(f"Actualmente hay **{len(rap_df):,}** registro(s) en **{len(_all_ubic_saved)}** ubicación(es).")
        with st.popover("🔥 Borrar TODAS las tomas guardadas",
                        use_container_width=True):
            st.error(
                f"⚠ Esto eliminará **TODOS** los {len(rap_df):,} registro(s) de "
                f"toma física guardados, de **TODAS** las ubicaciones. "
                f"El archivo `toma_fisica_rapida.xlsx` quedará vacío. "
                f"Los archivos de inventario (consolidado.xlsx) NO se afectan."
            )
            _typed = st.text_input(
                "Escribe **BORRAR** para confirmar",
                key="res_del_all_confirm_txt",
                placeholder="BORRAR",
            )
            if st.button("✔ Confirmar borrado total",
                         type="primary",
                         disabled=(_typed.strip().upper() != "BORRAR"),
                         key="res_del_all_confirm_btn"):
                empty = pd.DataFrame(columns=_RAPIDA_COLS)
                rap_state["df"] = empty
                _persist_rapid(empty)
                log(f"Resumen: borrado TOTAL ({len(rap_df):,} registros eliminados)")
                # Limpiar TODOS los editores de tomas activos
                for _k in [k for k in list(st.session_state.keys())
                           if k.startswith("tu_ed_")]:
                    st.session_state.pop(_k, None)
                st.success("✓ Todo el historial de tomas físicas fue eliminado.")
                _rerun_frag()

# -- Fragment: sub-pestaña Importar --
@_fragment
def _render_importar_fragment():
    r = st.session_state.get("result")
    eng = st.session_state.engine
    st.markdown("#### 📥 Importar toma física desde Excel")
    st.caption(
        "Detecta automáticamente dos formatos:\n\n"
        "• **Plantilla** (una hoja por ubicación, con columna Cantidad) — generada en «📋 Plantilla».\n"
        "• **Historial** (una sola hoja con columnas Fecha, Ubicación, SKU, Nombre, Cantidad, Observación) "
        "— exportada desde «📊 Resumen → Exportar historial». Útil para restaurar backups tras un reboot."
    )

    pf = st.file_uploader("Archivo (.xlsx / .xls)",
                          type=["xlsx","xls"], key="ph_up")

    if pf is not None:
        # Parsear en memoria
        _tp_ext = os.path.splitext(pf.name)[1]
        with tempfile.NamedTemporaryFile(delete=False, suffix=_tp_ext) as _tmp:
            _tmp.write(pf.getvalue())
            _tp = _tmp.name

        # Detección de formato
        _is_historial = False
        try:
            _is_historial = _detect_historial_format(_tp)
        except Exception:
            pass

        if _is_historial:
            # ── Flujo HISTORIAL (restauración de backup) ──────────
            try:
                _hist_rows, _hist_ubic = _parse_historial(_tp)
            except Exception as _ex:
                st.error(f"Error al leer el historial: {_ex}")
                _hist_rows, _hist_ubic = [], set()
            try: os.unlink(_tp)
            except: pass

            if not _hist_rows:
                st.warning("El archivo no contiene filas válidas para importar.")
                return

            st.info(
                f"📘 Detectado formato **Historial** (restauración de backup). "
                f"Contiene **{len(_hist_rows):,}** filas en "
                f"**{len(_hist_ubic)}** ubicación(es). Las fechas originales se preservan."
            )

            # Validar que las ubicaciones existan (si no, ofrecer crearlas)
            _registered = set(_get_all_ubic())
            _missing = sorted([u for u in _hist_ubic if u not in _registered])

            if _missing:
                st.warning(
                    f"⚠ El historial incluye {len(_missing)} ubicación(es) no "
                    f"registradas en el sistema:\n\n"
                    + "\n".join(f"- `{u}`" for u in _missing)
                )
                _to_create_h = st.multiselect(
                    "Ubicaciones a crear automáticamente",
                    _missing, default=_missing,
                    key="imp_hist_create",
                    help="Las no seleccionadas omitirán sus filas al importar."
                )
                _btn_label = "✔ Crear ubicación(es) e importar historial"
            else:
                _to_create_h = []
                _btn_label = f"✔ Importar historial ({len(_hist_rows):,} filas)"

            # Preview resumen por ubicación
            _resumen = {}
            for _r in _hist_rows:
                _u = _r["Ubicación"]
                if _u not in _resumen: _resumen[_u] = {"rows":0,"units":0}
                _resumen[_u]["rows"]  += 1
                _resumen[_u]["units"] += _r["Cantidad Física"]
            _prev_df = pd.DataFrame([
                {"Ubicación": u, "Registros": v["rows"], "Unidades": int(v["units"]),
                 "Estado": "✓ Registrada" if u in _registered else
                           ("➕ Se creará" if u in _to_create_h else "⏭ Se omitirá")}
                for u, v in _resumen.items()
            ])
            st.dataframe(_prev_df, use_container_width=True, hide_index=True)

            if st.button(_btn_label, type="primary", use_container_width=True,
                         key="imp_hist_confirm"):
                # Crear ubicaciones nuevas seleccionadas
                if _to_create_h:
                    _custom = _get_custom_ubic()
                    _existing = set(_custom["list"])
                    for _n in _to_create_h:
                        if _n not in _existing:
                            _custom["list"].append(_n)
                    _persist_custom_ubic(_custom["list"])

                # Filtrar filas: solo ubicaciones registradas o recién creadas
                _valid_ubic = _registered | set(_to_create_h)
                _to_save = [r for r in _hist_rows if r["Ubicación"] in _valid_ubic]
                _skipped = len(_hist_rows) - len(_to_save)

                rap_state = _get_shared_rapid()
                rap_df    = rap_state["df"]
                if _to_save:
                    _merged = pd.concat([rap_df, pd.DataFrame(_to_save)],
                                        ignore_index=True)
                    rap_state["df"] = _merged
                    _persist_rapid(_merged)
                log(f"Import historial ({pf.name}): {len(_to_save)} filas restauradas"
                    + (f" ({_skipped} omitidas)" if _skipped else ""))
                _msg = f"✅ Restauradas {len(_to_save):,} filas del historial."
                if _skipped:
                    _msg += f" Se omitieron {_skipped} por ubicaciones descartadas."
                st.success(_msg)
                _rerun_frag()
            return   # Terminar aquí el flujo historial

        try:
            parsed = _parse_plantilla_toma(_tp, _get_all_ubic())
        except Exception as _ex:
            st.error(f"Error al leer el Excel: {_ex}")
            parsed = None
        finally:
            try: os.unlink(_tp)
            except: pass

        if parsed is not None:
            # Reset de lista de "ignorados" si cambió el archivo
            _fhash = hash(pf.getvalue()[:50000])
            if st.session_state.get("_imp_fhash") != _fhash:
                st.session_state.pop("_imp_ignore", None)
                st.session_state["_imp_fhash"] = _fhash

            _ignored = set(st.session_state.get("_imp_ignore", []))
            # Hojas inválidas aún visibles (no marcadas como ignoradas)
            _invalid_vis = [s for s in parsed["invalid_sheets"]
                            if s not in _ignored]

            if _invalid_vis:
                st.warning(
                    "⚠ El Excel contiene hojas que **no corresponden a "
                    "ubicaciones registradas**:\n\n"
                    + "\n".join(f"- `{s}`" for s in _invalid_vis)
                    + "\n\nPara las que marques, se **crearán nuevas ubicaciones**. "
                      "Las que **desmarques**, se **omitirán** de la importación."
                )
                _to_create = st.multiselect(
                    "Ubicaciones a crear en el sistema (desmarcar = omitir)",
                    _invalid_vis,
                    default=_invalid_vis,
                    key="imp_create_ubic",
                )
                _to_skip = [s for s in _invalid_vis if s not in _to_create]
                if _to_skip:
                    st.info(
                        "Se **omitirán** estas hojas (no se importan): "
                        + ", ".join(f"`{s}`" for s in _to_skip)
                    )

                _btn_label = "✔ Continuar con la importación"
                if _to_create:
                    _btn_label = (f"➕ Crear {len(_to_create)} y continuar"
                                   + (f" (omitir {len(_to_skip)})" if _to_skip else ""))
                if st.button(_btn_label, type="primary", use_container_width=True,
                             key="imp_create_and_continue"):
                    if _to_create:
                        _custom = _get_custom_ubic()
                        _already = set(_custom["list"])
                        _added = 0
                        for _name in _to_create:
                            if _name not in _already:
                                _custom["list"].append(_name)
                                _added += 1
                        _persist_custom_ubic(_custom["list"])
                        log(f"Importar: creadas {_added} ubicación(es) desde Excel")
                    if _to_skip:
                        st.session_state["_imp_ignore"] = list(_ignored | set(_to_skip))
                        log(f"Importar: omitidas {len(_to_skip)} hoja(s) — {', '.join(_to_skip)}")
                    _rerun_frag()

                with st.expander("ℹ Ubicaciones registradas actualmente"):
                    st.caption(", ".join(f"`{u}`" for u in _get_all_ubic()))
            elif not parsed["valid_sheets"] or all(
                    len(info["rows"]) == 0
                    for info in parsed["valid_sheets"].values()):
                st.warning(
                    "El archivo no contiene cantidades para importar. "
                    "Recuerda: celdas vacías o con valor 0 se consideran «no contadas» y se omiten."
                )
            else:
                # Preview
                st.markdown("##### 👁 Vista previa")
                _prev = []
                _need_date = []
                for ubic, info in parsed["valid_sheets"].items():
                    _n = len(info["rows"])
                    if _n == 0: continue
                    _prev.append({
                        "Ubicación": ubic,
                        "Items a importar": _n,
                        "Fecha toma (B2)": info["fecha"].strftime("%d/%m/%Y")
                                           if info["fecha"] else "— sin fecha —",
                    })
                    if info["fecha"] is None:
                        _need_date.append(ubic)
                _prev_df = pd.DataFrame(_prev)
                st.dataframe(_prev_df, use_container_width=True, hide_index=True)

                # Fecha fallback para hojas sin B2
                _fallback_date = None
                if _need_date:
                    st.warning(
                        "⚠ Estas hojas no traen fecha en la celda B2: "
                        + ", ".join(f"«{u}»" for u in _need_date)
                    )
                    _fallback_date = st.date_input(
                        "Fecha para estas hojas (si la dejas vacía, se usa la de importación)",
                        value=date.today(),
                        format="DD/MM/YYYY",
                        key="imp_fallback_date",
                    )

                _total_items = sum(len(info["rows"])
                                   for info in parsed["valid_sheets"].values())

                if st.button(f"✔ Confirmar importación ({_total_items} items en "
                             f"{len(_prev)} ubicación(es))",
                             type="primary", key="imp_confirm",
                             use_container_width=True):
                    rap_state = _get_shared_rapid()
                    rap_df = rap_state["df"]
                    _now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
                    _new_rows = []
                    for ubic, info in parsed["valid_sheets"].items():
                        if not info["rows"]: continue
                        if info["fecha"] is not None:
                            _fecha_str = info["fecha"].strftime("%Y-%m-%d") + " 00:00"
                        elif _fallback_date is not None:
                            _fecha_str = _fallback_date.strftime("%Y-%m-%d") + " 00:00"
                        else:
                            _fecha_str = _now_str
                        for _row in info["rows"]:
                            _new_rows.append({
                                "Fecha": _fecha_str,
                                "Ubicación": ubic,
                                "Código Producto": _row["codigo"],
                                "Nombre Producto": _row["nombre"],
                                "Cantidad Física": float(_row["cantidad"]),
                                "Observación": _row["obs"],
                            })
                    if _new_rows:
                        _merged = pd.concat([rap_df, pd.DataFrame(_new_rows)],
                                            ignore_index=True)
                        rap_state["df"] = _merged
                        _persist_rapid(_merged)
                        log(f"Import Excel ({pf.name}): {len(_new_rows)} items "
                            f"en {len(_prev)} ubicaciones")
                        st.success(
                            f"✅ Importados {len(_new_rows)} items en "
                            f"{len(_prev)} ubicación(es). Revisa «📊 Resumen» "
                            f"y «📊 Comparación»."
                        )
                        _rerun_frag()

# -- Fragment: sub-pestaña Plantilla --
@_fragment
def _render_plantilla_fragment():
    r = st.session_state.get("result")
    eng = st.session_state.engine
    if eng.raw_df is not None and st.button("📄 Generar plantilla",key="ph_t"):
        try:
            import openpyxl
            from openpyxl.styles import PatternFill,Font,Alignment,Border,Side
            from openpyxl.utils import get_column_letter,quote_sheetname
            sd=(eng.raw_df[["Código Producto","Nombre Producto"]].drop_duplicates()
                .sort_values("Nombre Producto").reset_index(drop=True))
            n=len(sd); DR=4
            H=PatternFill("solid",fgColor="1E3A5F"); Y=PatternFill("solid",fgColor="FEF9C3")
            G=PatternFill("solid",fgColor="D1FAE5"); Q=PatternFill("solid",fgColor="DBEAFE")
            F=PatternFill("solid",fgColor="F0FDF4"); E=PatternFill("solid",fgColor="F8FAFC")
            O=PatternFill("solid",fgColor="FFFFFF"); t=Side(style="thin",color="CBD5E1")
            brd=Border(t,t,t,t); tg=Side(style="thin",color="BBF7D0"); bg=Border(tg,tg,tg,tg)
            wb=openpyxl.Workbook(); ls={}
            _tpl_ubic = _get_all_ubic()
            for loc in _tpl_ubic:
                sn=loc[:28].replace("/","-").replace("\\","-").replace("?","").replace("*","").replace("[","").replace("]","").replace(":","")
                ws2=wb.create_sheet(sn); ls[loc]=(sn,ws2)
                ws2.merge_cells("A1:D1"); ws2["A1"]=f"TOMA FISICA — {loc.upper()}"
                ws2["A1"].font=Font(bold=True,size=12,color="FFFFFF"); ws2["A1"].fill=H
                ws2["A1"].alignment=Alignment(horizontal="center"); ws2.row_dimensions[1].height=24
                ws2["A2"]="FECHA TOMA:"; ws2["A2"].font=Font(bold=True,size=10,color="FFFFFF"); ws2["A2"].fill=H
                ws2["B2"].fill=PatternFill("solid",fgColor="EFF6FF"); ws2["B2"].font=Font(size=10,color="1E3A5F")
                for ci,h in enumerate(["Código","Nombre","Cantidad","Observación"],1):
                    ce=ws2.cell(3,ci,h); ce.font=Font(bold=True,size=9,color="1E3A5F")
                    ce.fill=Y; ce.alignment=Alignment(horizontal="center",wrap_text=True); ce.border=brd
                for ri,(_,row) in enumerate(sd.iterrows(),DR):
                    fill=E if ri%2==0 else O
                    ws2.cell(ri,1,str(row["Código Produto"] if "Código Produto" in row else row.get("Código Producto",""))).fill=fill
                    ws2.cell(ri,1).font=Font(size=9,color="1E40AF")
                    ws2.cell(ri,2,str(row["Nombre Producto"])).fill=fill; ws2.cell(ri,2).font=Font(size=9,color="111827")
                    qc=ws2.cell(ri,3,""); qc.fill=Q; qc.alignment=Alignment(horizontal="right"); qc.font=Font(size=10,bold=True,color="1E3A5F")
                    ws2.cell(ri,4,"").fill=fill
                    for ci in range(1,5): ws2.cell(ri,ci).border=brd
                    ws2.row_dimensions[ri].height=16
                tr2=n+DR; ws2.merge_cells(f"A{tr2}:B{tr2}")
                ws2.cell(tr2,1,"TOTAL").font=Font(bold=True,size=9,color="FFFFFF"); ws2.cell(tr2,1).fill=H
                tc2=ws2.cell(tr2,3,f"=SUM(C{DR}:C{tr2-1})")
                tc2.font=Font(bold=True,size=9,color="065F46"); tc2.fill=G
                tc2.alignment=Alignment(horizontal="right"); tc2.border=brd
                ws2.column_dimensions["A"].width=13; ws2.column_dimensions["B"].width=40
                ws2.column_dimensions["C"].width=11; ws2.column_dimensions["D"].width=30; ws2.freeze_panes="C4"
            ac=["Código","Nombre"]+_tpl_ubic+["TOTAL"]
            wr=wb.create_sheet("RESUMEN GENERAL",0); nc=len(ac)
            wr.merge_cells(f"A1:{get_column_letter(nc)}1")
            wr["A1"]="RESUMEN — Actualiza automáticamente"; wr["A1"].font=Font(bold=True,size=12,color="FFFFFF"); wr["A1"].fill=H
            wr["A1"].alignment=Alignment(horizontal="center"); wr.row_dimensions[1].height=24
            wr.merge_cells(f"A2:{get_column_letter(nc)}2")
            wr["A2"]="⚠ NO editar — calculado automáticamente desde cada hoja"
            wr["A2"].font=Font(bold=True,size=9,color="92400E"); wr["A2"].fill=PatternFill("solid",fgColor="FEF3C7")
            wr["A2"].alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); wr.row_dimensions[2].height=28
            for ci,col in enumerate(ac,1):
                ce=wr.cell(3,ci,col); ce.font=Font(bold=True,size=9,color="1E3A5F"); ce.fill=Y
                ce.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); ce.border=brd
            wr.row_dimensions[3].height=30
            for ri,(_,row) in enumerate(sd.iterrows(),DR):
                fill=E if ri%2==0 else O
                wr.cell(ri,1,str(row.get("Código Producto",""))).fill=fill; wr.cell(ri,1).font=Font(size=9,color="1E40AF")
                wr.cell(ri,2,str(row["Nombre Producto"])).fill=fill; wr.cell(ri,2).font=Font(size=9,color="111827")
                for li,loc in enumerate(_tpl_ubic):
                    ci=3+li; sn,_=ls[loc]
                    ce=wr.cell(ri,ci,f"=IFERROR({quote_sheetname(sn)}!C{ri},0)")
                    ce.fill=F; ce.font=Font(size=9,color="065F46"); ce.alignment=Alignment(horizontal="right"); ce.border=bg
                cf=get_column_letter(3); cl=get_column_letter(3+len(_tpl_ubic)-1)
                tc=wr.cell(ri,nc,f"=SUM({cf}{ri}:{cl}{ri})")
                tc.fill=G; tc.font=Font(bold=True,size=9,color="065F46"); tc.alignment=Alignment(horizontal="right"); tc.border=brd
                for ci in range(1,3): wr.cell(ri,ci).border=brd
                wr.row_dimensions[ri].height=16
            tr=n+DR; wr.merge_cells(f"A{tr}:B{tr}")
            wr.cell(tr,1,"TOTAL GENERAL").font=Font(bold=True,size=9,color="FFFFFF"); wr.cell(tr,1).fill=H
            for ci in range(3,nc+1):
                cl=get_column_letter(ci); ce=wr.cell(tr,ci,f"=SUM({cl}{DR}:{cl}{tr-1})")
                ce.font=Font(bold=True,size=9,color="FFFFFF"); ce.fill=H
                ce.alignment=Alignment(horizontal="right"); ce.border=brd
            wr.column_dimensions["A"].width=13; wr.column_dimensions["B"].width=38
            for ci in range(3,nc+1): wr.column_dimensions[get_column_letter(ci)].width=13
            wr.freeze_panes="C4"
            buf=io.BytesIO(); wb.save(buf)
            st.download_button("📥 Descargar plantilla",buf.getvalue(),"plantilla_toma_fisica.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.success("✓ Plantilla lista")
        except Exception as e: st.error(str(e))

# -- Fragment: sub-pestaña Comparacion --
@_fragment
def _render_comparacion_fragment():
    r = st.session_state.get("result")
    eng = st.session_state.engine
    # Comparación unificada: lee de toma_fisica_rapida.xlsx
    # (mismo archivo que usan «⚡ Toma» e «📥 Importar»).
    _rap_state = _get_shared_rapid()
    _rap_df    = _rap_state["df"]
    if r is None:
        st.info("Ejecute el análisis primero.")
        return
    if _rap_df is None or _rap_df.empty:
        st.info("Aún no hay tomas registradas. Usa «⚡ Toma» o «📥 Importar».")
        return

    _inv = r.inventory_by_warehouse.copy()
    if excl_w: _inv = _inv[~_inv["Bodega"].isin(excl_w)]
    _all_bod = sorted(_inv["Bodega"].dropna().unique().tolist())

    # Selector de "bodegas contables" (físicamente contables vs consignadas en clientes).
    # Default: solo Bodega Principal. Se persiste en session_state.
    if "_cmp_contables" not in st.session_state:
        default_contables = [b for b in _all_bod if b == PRIMARY_WAREHOUSE]
        if not default_contables and _all_bod:
            default_contables = [_all_bod[0]]
        st.session_state["_cmp_contables"] = default_contables

    sc1, sc2 = st.columns([3, 2])
    with sc1:
        contables = st.multiselect(
            "🏢 Bodegas contables (físicamente contables — se comparan contra la toma)",
            _all_bod,
            default=st.session_state["_cmp_contables"],
            key="_cmp_contables",
            help="Bodegas donde el inventario está en tu posesión y se puede contar. "
                 "El resto se considera 'consignado en clientes' y se reporta aparte.",
        )
    with sc2:
        _tol = st.number_input(
            "Tolerancia (± unidades para 'coincide')",
            min_value=0.0, max_value=100.0, value=0.5, step=0.5,
            key="_cmp_tol",
        )

    if not contables:
        st.warning("Selecciona al menos una bodega contable.")
        return

    # ── FÍSICO: suma por SKU de la última toma en cada ubicación
    _latest = _rap_df.sort_values("Fecha").drop_duplicates(
        subset=["Código Producto","Ubicación"], keep="last")
    _phys_sku = (_latest.groupby(["Código Producto","Nombre Producto"])
                 ["Cantidad Física"].sum().reset_index())

    # ── CALCULADO CONTABLE: stock SOLO en bodegas contables
    _inv_ct  = _inv[_inv["Bodega"].isin(contables)]
    _calc_ct = (_inv_ct.groupby(["Código Producto","Nombre Producto"])
                ["Stock"].sum().reset_index()
                .rename(columns={"Stock":"Cantidad Calculada"}))

    # ── CONSIGNADO EN CLIENTES: stock en bodegas NO contables (separado)
    _inv_cl  = _inv[~_inv["Bodega"].isin(contables)]
    _cons_cl = (_inv_cl.groupby(["Código Producto","Nombre Producto"])
                ["Stock"].sum().reset_index()
                .rename(columns={"Stock":"En Clientes"}))

    # Merge outer para la comparación
    _cmp = _phys_sku.merge(_calc_ct,
                           on=["Código Producto","Nombre Producto"],
                           how="outer").fillna(0)
    _cmp = _cmp.merge(_cons_cl,
                      on=["Código Producto","Nombre Producto"],
                      how="left").fillna(0)
    # Diferencia = Sistema − Físico
    #   > 0 → FALTANTE (sistema dice que debería haber más de lo contado)
    #   < 0 → SOBRANTE (hay más físicamente de lo que el sistema dice)
    _cmp["Diferencia"] = _cmp["Cantidad Calculada"] - _cmp["Cantidad Física"]
    _cmp["Coincide"]   = _cmp["Diferencia"].abs() <= _tol
    _cmp = _cmp.sort_values("Diferencia",
                            key=lambda s: s.abs(), ascending=False)

    # Banner explicativo
    st.markdown(
        f"<div style='background:#e0f2fe;border:1px solid #7dd3fc;border-radius:8px;"
        f"padding:8px 14px;margin:8px 0;font-size:12px;color:#075985;line-height:1.5'>"
        f"<b>📘 Lógica de la comparación</b><br>"
        f"• <b>Físico</b> = suma de la última cantidad contada por SKU en todas las "
        f"ubicaciones de la toma.<br>"
        f"• <b>Calculado</b> = stock del sistema SOLO en las bodegas marcadas como "
        f"«contables» ({', '.join(f'<code>{b}</code>' for b in contables)}).<br>"
        f"• Las bodegas <b>no contables</b> (consignadas en clientes) se listan "
        f"aparte abajo — su stock es tuyo pero no se puede contar físicamente.<br>"
        f"• <b>Diferencia</b> = Sistema − Físico. "
        f"<b>+</b> = faltante (sistema dice que debería haber más · investigar). "
        f"<b>−</b> = sobrante (hay más físicamente). "
        f"Coincide si |Diff| ≤ {_tol:g}."
        f"</div>",
        unsafe_allow_html=True
    )

    # KPIs
    _ex   = (_cmp["Coincide"].mean()*100) if len(_cmp) else 0
    _dif  = int((_cmp["Diferencia"].abs() > _tol).sum())
    _ok   = int(_cmp["Coincide"].sum())
    _total_fis = int(_cmp["Cantidad Física"].sum())
    _total_ct  = int(_cmp["Cantidad Calculada"].sum())
    _total_cl  = int(_cmp["En Clientes"].sum())
    m1,m2,m3,m4 = st.columns(4)
    m1.metric("Exactitud", f"{_ex:.1f}%")
    m2.metric("Con diferencia", _dif)
    m3.metric("Coinciden", _ok)
    m4.metric("SKUs", len(_cmp))

    n1,n2,n3 = st.columns(3)
    n1.metric("Σ Físico contado", f"{_total_fis:,} u")
    n2.metric("Σ Calculado contable", f"{_total_ct:,} u")
    n3.metric("Σ En clientes", f"{_total_cl:,} u",
              help="Stock del sistema en bodegas consignadas a clientes (no contables)")

    # Filtro de vista
    _show_mode = st.radio("Mostrar",
        ["Con diferencia","Todo","Coinciden"],
        horizontal=True, key="cmp_mode")
    if _show_mode == "Con diferencia":
        _cmp_show = _cmp[_cmp["Diferencia"].abs() > _tol].copy()
    elif _show_mode == "Coinciden":
        _cmp_show = _cmp[_cmp["Coincide"]].copy()
    else:
        _cmp_show = _cmp.copy()

    _cmp_show = _cmp_show[["Código Producto","Nombre Producto",
                           "Cantidad Calculada","Cantidad Física",
                           "Diferencia","En Clientes","Coincide"]].reset_index(drop=True)
    for _c in ("Cantidad Calculada","Cantidad Física","Diferencia","En Clientes"):
        _cmp_show[_c] = pd.to_numeric(_cmp_show[_c], errors="coerce").fillna(0).astype(int)
    _cmp_show["Coincide"] = _cmp_show["Coincide"].map({True:"✓", False:"✗"})

    st.dataframe(_cmp_show, use_container_width=True, hide_index=True, height=520,
                 column_config={
                     "Código Producto": st.column_config.TextColumn("SKU", width="small"),
                     "Nombre Producto": st.column_config.TextColumn("Producto"),
                     "Cantidad Calculada": st.column_config.NumberColumn(
                         "Sistema (contable)", format="%d",
                         help=f"Stock en: {', '.join(contables)}"),
                     "Cantidad Física":    st.column_config.NumberColumn("Físico", format="%d"),
                     "Diferencia":         st.column_config.NumberColumn("Diferencia", format="%+d"),
                     "En Clientes":        st.column_config.NumberColumn(
                         "En clientes", format="%d",
                         help="Stock en bodegas consignadas (informativo)"),
                     "Coincide":           st.column_config.TextColumn("OK", width="small"),
                 })
    dl3(_cmp_show, "comparacion_toma", "ph")

    # ── Detalle de Consignado en Clientes ──────────────────────
    st.markdown("---")
    with st.expander(f"🏪 Stock consignado en clientes ({len(_inv_cl['Bodega'].unique())} bodegas · "
                     f"{_total_cl:,} u totales) — no se cuenta físicamente"):
        if _inv_cl.empty:
            st.caption("No hay stock en bodegas consignadas con los filtros actuales.")
        else:
            # Pivot SKU × Bodega de clientes
            _pv_cl = _inv_cl.pivot_table(
                index=["Código Producto","Nombre Producto"],
                columns="Bodega", values="Stock", aggfunc="sum"
            ).fillna(0).astype(int).reset_index()
            # Total por SKU
            _cl_cols = [c for c in _pv_cl.columns
                        if c not in ("Código Producto","Nombre Producto")]
            _pv_cl["Σ En clientes"] = _pv_cl[_cl_cols].sum(axis=1)
            _pv_cl = _pv_cl.sort_values("Σ En clientes", ascending=False)
            st.dataframe(_pv_cl, use_container_width=True, hide_index=True, height=400,
                         column_config={
                             "Código Producto": st.column_config.TextColumn("SKU", width="small"),
                             "Nombre Producto": st.column_config.TextColumn("Producto"),
                             "Σ En clientes":   st.column_config.NumberColumn("Σ Total", format="%d"),
                             **{c: st.column_config.NumberColumn(c, format="%d") for c in _cl_cols},
                         })
            st.download_button("📥 Exportar consignado a Excel",
                to_xl(_pv_cl), "consignado_clientes.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, key="cons_cl_xl")


_perf("before_tab_phy")
# ══ TAB 9 TOMA FÍSICA ═══════════════════════════════════════════
with T_PHY:
    st.markdown("### 🏭 Toma Física")
    p0,p_res,p1,p2,p3=st.tabs([
        "⚡ Toma", "📊 Resumen",
        "📥 Importar", "📋 Plantilla", "📊 Comparación"
    ])

    # ── Sub-pestaña: TOMA POR UBICACIÓN (aislada en @st.fragment) ──
    with p0:
        _render_toma_fragment()

    # ── Sub-pestaña: RESUMEN (aislado en @st.fragment) ────────────
    with p_res:
        _render_resumen_fragment()


    with p1:
        _render_importar_fragment()
    with p2:
        _render_plantilla_fragment()
    with p3:
        _render_comparacion_fragment()

# ══ TAB 10 LOG ══════════════════════════════════════════════════
with T_LOG:
    if st.button("🗑 Limpiar",key="lc"): st.session_state.log=[]
    for e in st.session_state.log: st.text(e)
    if not st.session_state.log: st.info("Sin actividad.")

# ── Panel visual de performance en el sidebar ───────────────────
_perf("end")
with st.sidebar:
    with st.expander("⏱ Performance (últimos reruns)", expanded=False):
        _hist = _get_perf_history()["runs"]
        if _hist:
            _rows = []
            for _r in _hist[-15:]:
                _d = {"hora": _r["ts"], "total_ms": _r["total_ms"]}
                for _lbl, _ms in _r["ck"]:
                    _d[_lbl] = _ms
                _rows.append(_d)
            _pdf = pd.DataFrame(_rows).iloc[::-1]
            st.dataframe(_pdf, use_container_width=True, hide_index=True, height=300)
            st.caption(f"Archivo: `perf.log` ({_PERF_LOG_PATH})")
            if st.button("🗑 Limpiar historial", key="perf_clear"):
                _get_perf_history()["runs"] = []
                try: os.remove(_PERF_LOG_PATH)
                except: pass
                st.rerun()
        else:
            st.caption("Sin datos todavía. Interactúa con la app para ver mediciones.")

_perf_flush()
