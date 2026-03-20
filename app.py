import streamlit as st
import numpy as np
import pandas as pd
import io
import json
import os
import hashlib
import base64
from dataclasses import dataclass
from io import BytesIO
from openpyxl import load_workbook
import datetime as dt
import plotly.express as px
import streamlit.components.v1 as components

try:
    import xlsxwriter  # noqa: F401
    HAS_XLSXWRITER = True
except Exception:
    HAS_XLSXWRITER = False

try:
    from statsmodels.tsa.holtwinters import ExponentialSmoothing
    HAS_STATSMODELS = True
except Exception:
    ExponentialSmoothing = None
    HAS_STATSMODELS = False


# ---------- ВНЕШНЙ ВД (CSS) ----------

st.set_page_config(page_title="Медиапланер 12 месяцев (кампании)", layout="wide")

THEME_BORDER = "#1D2A44"
THEME_PLOT_TEXT = "#EAF0FF"
THEME_LEGEND_TEXT = "#D4DDF2"
THEME_CARD_BG = "#111A2E"
VAT_RATE = 0.22
USE_EXCEL_ROUNDDOWN = True

METRIC_PRESETS = {
    "ecom": {
        "label": "E-com",
        "description": "Базовый пресет для e-commerce: показы, CTR, CPC, CR, AOV и производные метрики.",
    },
    "diy": {
        "label": "DIY",
        "description": "Пресет для DIY: продажи (от кликов), CR в продажи, SOV и доля новых клиентов.",
    },
}

st.markdown(
    """
    <style>
    .main {
        background-color: #0B1220;
        color: #EAF0FF;
    }
    .stApp {
        background-color: #0B1220;
        color: #EAF0FF;
    }
    [data-testid="block-container"] {
        padding-top: 1.05rem;
        padding-bottom: 0.9rem;
    }
    section[data-testid="stSidebar"] {
        background-color: #111A2E;
        border-right: 1px solid #1D2A44;
    }
    h1, h2, h3, h4 {
        font-weight: 700;
        letter-spacing: 0.02em;
        color: #EAF0FF;
    }
    p, li, label {
        color: #D4DDF2;
    }
    .stDataFrame {
        border-radius: 12px;
        border: 1px solid #1D2A44;
    }
    .stDownloadButton button {
        background: #0066E0;
        color: #FFFFFF;
        font-weight: 700;
        border-radius: 999px;
        border: 1px solid #2B7EE8;
    }
    .stDownloadButton button:hover {
        background: #0A74F2;
        border-color: #0A74F2;
    }
    .stButton > button, .stFormSubmitButton > button {
        border-radius: 10px;
        border: 1px solid #27406F;
        background: #111A2E;
        color: #EAF0FF;
        font-weight: 600;
    }
    .stButton > button:hover, .stFormSubmitButton > button:hover {
        border-color: #0066E0;
        background: #14203A;
    }
    button[kind="primary"],
    button[kind="primaryFormSubmit"],
    .stButton > button[kind="primary"],
    .stFormSubmitButton > button[kind="primary"],
    .stFormSubmitButton > button[kind="primaryFormSubmit"],
    button[data-testid="baseButton-primary"],
    button[data-testid="stBaseButton-primary"] {
        background: linear-gradient(180deg, #0A74F2 0%, #0066E0 100%) !important;
        border: 1px solid #3D8EF0 !important;
        color: #FFFFFF !important;
        box-shadow: 0 0 0 1px rgba(61, 142, 240, 0.25), 0 6px 16px rgba(0, 102, 224, 0.35);
    }
    button[kind="primary"]:hover,
    button[kind="primaryFormSubmit"]:hover,
    .stButton > button[kind="primary"]:hover,
    .stFormSubmitButton > button[kind="primary"]:hover,
    .stFormSubmitButton > button[kind="primaryFormSubmit"]:hover,
    button[data-testid="baseButton-primary"]:hover,
    button[data-testid="stBaseButton-primary"]:hover {
        background: linear-gradient(180deg, #1C82F6 0%, #0A74F2 100%) !important;
        border-color: #74AEF6 !important;
        color: #FFFFFF !important;
    }
    [data-baseweb="tag"] {
        background-color: #1A2A47 !important;
        border: 1px solid #2C4D82 !important;
        color: #9EC5FF !important;
    }
    button[data-baseweb="tab"] {
        color: #9FB0D1 !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] {
        color: #0066E0 !important;
    }
    [data-baseweb="tab-highlight"] {
        background-color: #0066E0 !important;
    }
    .bottom-tab-nav {
        display: flex;
        gap: 1rem;
        align-items: end;
        border-bottom: 1px solid #1D2A44;
        padding-bottom: 6px;
        margin-top: 8px;
    }
    .bottom-tab-btn {
        background: transparent;
        border: 0;
        border-bottom: 2px solid transparent;
        color: #9FB0D1;
        font-weight: 650;
        font-size: 1rem;
        padding: 0 0 8px 0;
        cursor: pointer;
    }
    .bottom-tab-btn.is-active {
        color: #EAF0FF;
        border-bottom-color: #0066E0;
    }
    .ui-section-title {
        margin: 0.22rem 0 0.42rem 0;
        font-size: 1.35rem;
        font-weight: 700;
        letter-spacing: 0.01em;
        line-height: 1.2;
    }
    .tab-intro {
        margin: 0.2rem 0 0.7rem 0;
        padding: 10px 12px;
        border: 1px solid #1D2A44;
        border-radius: 10px;
        background: rgba(17, 26, 46, 0.55);
        animation: introIn 220ms ease both;
    }
    .tab-intro p {
        margin: 0.12rem 0;
        color: #D4DDF2;
        line-height: 1.45;
    }
    [data-testid="stExpander"] {
        margin-bottom: 0.3rem;
    }
    [data-testid="stExpander"] summary p {
        font-weight: 650;
        letter-spacing: 0.01em;
    }
    /* MOTION START */
    .stButton > button,
    .stFormSubmitButton > button,
    .stDownloadButton > button {
        transition: transform 140ms ease, box-shadow 180ms ease, background-color 180ms ease, border-color 180ms ease;
    }
    .stButton > button:hover,
    .stFormSubmitButton > button:hover,
    .stDownloadButton > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 6px 14px rgba(0, 0, 0, 0.22);
    }
    .stButton > button:active,
    .stFormSubmitButton > button:active,
    .stDownloadButton > button:active {
        transform: translateY(0);
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.16);
    }
    [data-testid="stExpander"] {
        transition: border-color 180ms ease, box-shadow 180ms ease, transform 160ms ease;
    }
    [data-testid="stExpander"]:hover {
        border-color: #2C4D82;
        box-shadow: 0 6px 16px rgba(0, 0, 0, 0.18);
        transform: translateY(-1px);
    }
    [data-testid="stDataFrame"] {
        transition: border-color 180ms ease, box-shadow 180ms ease;
    }
    [data-testid="stDataFrame"]:hover {
        border-color: #2C4D82;
        box-shadow: 0 8px 20px rgba(0, 0, 0, 0.2);
    }
    [data-baseweb="tab"] {
        transition: color 160ms ease;
    }
    [data-baseweb="tab-highlight"] {
        transition: transform 180ms ease, width 180ms ease, left 180ms ease;
    }
    .ui-section-title {
        animation: fadeInUp 260ms ease both;
    }
    @keyframes fadeInUp {
        from { opacity: 0; transform: translateY(4px); }
        to { opacity: 1; transform: translateY(0); }
    }
    @keyframes introIn {
        from { opacity: 0; transform: translateY(6px); }
        to { opacity: 1; transform: translateY(0); }
    }
    /* MOTION END */
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------- МОДЕЛЬ  ФУНКЦ РАСЧЁТА ----------

@dataclass
class PlanInput:
    impressions: float  # Показы
    ctr: float          # CTR как доля (0.01 = 1%)
    cpc: float          # ₽
    cr: float           # CR как доля (0.02 = 2%)
    aov: float          # ₽


def calculate_plan_month(inp: PlanInput) -> dict:
    raw_clicks = inp.impressions * inp.ctr
    clicks = np.floor(raw_clicks) if USE_EXCEL_ROUNDDOWN else raw_clicks
    # В шаблоне Excel стоимость считается от S*CTR (до округления кликов).
    cost = raw_clicks * inp.cpc
    conv = np.floor(clicks * inp.cr) if USE_EXCEL_ROUNDDOWN else (clicks * inp.cr)
    revenue = conv * inp.aov

    cpm = cost / (inp.impressions / 1000) if inp.impressions > 0 else 0
    cpa = cost / conv if conv > 0 else 0
    roas = revenue / cost if cost > 0 else 0
    drr = cost / revenue if revenue > 0 else 0

    return {
        "impressions": inp.impressions,
        "ctr": inp.ctr,
        "cpc": inp.cpc,
        "cr": inp.cr,
        "aov": inp.aov,
        "clicks": clicks,
        "conversions": conv,
        "cost": cost,
        "revenue": revenue,
        "cpm": cpm,
        "cpa": cpa,
        "roas": roas,
        "drr": drr,
    }


def calc_month_for_all_campaigns(base_campaigns: pd.DataFrame,
                                 coeffs_month: pd.DataFrame,
                                 month_num: int,
                                 month_name: str) -> pd.DataFrame:
    rows = []
    for _, base_row in base_campaigns.iterrows():
        campaign_type = str(base_row["campaign_type"])
        k_row = coeffs_month[coeffs_month["campaign_type"] == campaign_type]
        if k_row.empty:
            k_imp = k_ctr = k_cpc = k_cr = k_aov = 1.0
        else:
            k_row = k_row.iloc[0]
            k_imp = k_row["k_imp"]
            k_ctr = k_row["k_ctr"]
            k_cpc = k_row["k_cpc"]
            k_cr = k_row["k_cr"]
            k_aov = k_row["k_aov"]

        base = PlanInput(
            impressions=base_row["impressions_avg"],
            ctr=base_row["ctr_avg_percent"] / 100.0,
            cpc=base_row["cpc_avg"],
            cr=base_row["cr_avg_percent"] / 100.0,
            aov=base_row["aov_avg"],
        )

        month_inp = PlanInput(
            impressions=base.impressions * k_imp,
            ctr=base.ctr * k_ctr,
            cpc=base.cpc * k_cpc,
            cr=base.cr * k_cr,
            aov=base.aov * k_aov,
        )

        out = calculate_plan_month(month_inp)
        out["month_num"] = month_num
        out["month_name"] = month_name
        out["campaign_type"] = campaign_type
        out["system"] = base_row["system"]
        out["format"] = base_row["format"]

        rows.append(out)

    return pd.DataFrame(rows)


def normalize_coeff_set_type(raw_type: str) -> str:
    """
    Normalize legacy/custom spelling variants of coefficient set types.
    This keeps older session_state values compatible after UI text changes.
    """
    val = str(raw_type or "").strip()
    compact = val.lower().replace(" ", "")
    if compact in {"спрос(позапросам)", "спроспозапросам", "demand"}:
        return "Спрос (по запросам)"
    if compact in {"aov(среднийчек)", "aovсреднийчек", "aov"}:
        return "AOV (средний чек)"
    if compact in {"кастомныйнабор", "custom", "customset"}:
        return "Кастомный набор"
    if compact in {"медийныехвосты", "медийныйхвост", "mediatails", "media_tail", "media tails"}:
        return "Медийные хвосты"
    return val


def ui_section_title(text: str) -> None:
    st.markdown(f"<div class='ui-section-title'>{text}</div>", unsafe_allow_html=True)


def parse_float_loose(value, default: float = 0.0) -> float:
    """Parse numeric values robustly (supports spaces, commas and percent sign)."""
    try:
        if value is None:
            return float(default)
        if isinstance(value, str):
            cleaned = value.strip().replace(" ", "").replace("%", "").replace(",", ".")
            if cleaned == "":
                return float(default)
            return float(cleaned)
        if pd.isna(value):
            return float(default)
        return float(value)
    except Exception:
        return float(default)


def image_file_to_data_uri(image_path: str) -> str | None:
    if not image_path or not os.path.exists(image_path):
        return None
    ext = os.path.splitext(image_path)[1].lower()
    mime = {
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".webp": "image/webp",
        ".svg": "image/svg+xml",
    }.get(ext)
    if not mime:
        return None
    try:
        raw = open(image_path, "rb").read()
    except Exception:
        return None
    encoded = base64.b64encode(raw).decode("ascii")
    return f"data:{mime};base64,{encoded}"


def _df_to_payload(df: pd.DataFrame | None) -> dict:
    if df is None or not isinstance(df, pd.DataFrame):
        return {"columns": [], "rows": []}
    safe_df = df.replace({np.nan: None})
    return {
        "columns": list(safe_df.columns),
        "rows": safe_df.to_dict(orient="records"),
    }


def _df_from_payload(payload: dict | None) -> pd.DataFrame:
    if not payload or not isinstance(payload, dict):
        return pd.DataFrame()
    cols = payload.get("columns", [])
    rows = payload.get("rows", [])
    if not isinstance(cols, list) or not isinstance(rows, list):
        return pd.DataFrame()
    return pd.DataFrame(rows, columns=cols if cols else None)


def _bootstrap_reference_from_campaigns(campaigns_df: pd.DataFrame | None) -> None:
    """
    Prepare quick-reference sidebar data right after import.
    This keeps "Быстрая сверка" usable even before full recalculation blocks run.
    """
    if campaigns_df is None or not isinstance(campaigns_df, pd.DataFrame) or campaigns_df.empty:
        return

    required_cols = {"campaign_type", "impressions_avg", "ctr_avg_percent", "cpc_avg", "cr_avg_percent", "aov_avg"}
    if not required_cols.issubset(set(campaigns_df.columns)):
        return

    work = campaigns_df.copy()
    work["campaign_type"] = work["campaign_type"].astype(str).str.strip()
    work = work[work["campaign_type"] != ""]
    if work.empty:
        return

    use_vat = bool(st.session_state.get("use_vat_budget_metrics", True))
    base_by_campaign = {}
    total_imp = total_clicks = total_conv = total_cost = total_rev = 0.0

    for _, row in work.iterrows():
        inp = PlanInput(
            impressions=max(0.0, parse_float_loose(row.get("impressions_avg"), 0.0)),
            ctr=max(0.0, parse_float_loose(row.get("ctr_avg_percent"), 0.0) / 100.0),
            cpc=max(0.0, parse_float_loose(row.get("cpc_avg"), 0.0)),
            cr=max(0.0, parse_float_loose(row.get("cr_avg_percent"), 0.0) / 100.0),
            aov=max(0.0, parse_float_loose(row.get("aov_avg"), 0.0)),
        )
        out = calculate_plan_month(inp)
        camp = str(row.get("campaign_type", "")).strip()
        if not camp:
            continue

        imp = float(out.get("impressions", 0.0))
        clicks = float(out.get("clicks", 0.0))
        conv = float(out.get("conversions", 0.0))
        cost = float(out.get("cost", 0.0))
        rev = float(out.get("revenue", 0.0))
        cost_with_vat = cost * (1.0 + VAT_RATE)
        budget_basis = cost_with_vat if use_vat else cost
        cpo = (budget_basis / conv) if conv > 0 else 0.0
        roas = (rev / budget_basis * 100.0) if budget_basis > 0 else 0.0
        drr = (budget_basis / rev * 100.0) if rev > 0 else 0.0

        base_by_campaign[camp] = {
            "impressions": imp,
            "clicks": clicks,
            "conversions": conv,
            "cost": cost,
            "cost_with_vat": cost_with_vat,
            "cost_with_vat_ak": cost_with_vat,
            "revenue": rev,
            "ctr": (clicks / imp * 100.0) if imp > 0 else 0.0,
            "cpc": (cost / clicks) if clicks > 0 else 0.0,
            "cr": (conv / clicks * 100.0) if clicks > 0 else 0.0,
            "cpo": cpo,
            "roas": roas,
            "drr": drr,
        }

        total_imp += imp
        total_clicks += clicks
        total_conv += conv
        total_cost += cost
        total_rev += rev

    if not base_by_campaign:
        return

    total_cost_with_vat = total_cost * (1.0 + VAT_RATE)
    total_budget_basis = total_cost_with_vat if use_vat else total_cost
    st.session_state["mp_ref_base_by_campaign"] = base_by_campaign
    st.session_state["mp_ref_base"] = {
        "impressions": float(total_imp),
        "clicks": float(total_clicks),
        "conversions": float(total_conv),
        "cost": float(total_cost),
        "cost_with_vat": float(total_cost_with_vat),
        "cost_with_vat_ak": float(total_cost_with_vat),
        "revenue": float(total_rev),
        "ctr": float((total_clicks / total_imp * 100.0) if total_imp > 0 else 0.0),
        "cpc": float((total_cost / total_clicks) if total_clicks > 0 else 0.0),
        "cr": float((total_conv / total_clicks * 100.0) if total_clicks > 0 else 0.0),
        "cpo": float((total_budget_basis / total_conv) if total_conv > 0 else 0.0),
        "roas": float((total_rev / total_budget_basis * 100.0) if total_budget_basis > 0 else 0.0),
        "drr": float((total_budget_basis / total_rev * 100.0) if total_rev > 0 else 0.0),
    }


def export_project_state_payload() -> dict:
    coeff_sets_src = st.session_state.get("coeff_sets", [])
    coeff_sets_payload = []
    for cs in coeff_sets_src:
        if not isinstance(cs, dict):
            continue
        coeff_sets_payload.append(
            {
                "id": int(cs.get("id", 0) or 0),
                "name": str(cs.get("name", "")),
                "type": str(cs.get("type", "")),
                "start_month": int(cs.get("start_month", 1) or 1),
                "start_year": int(cs.get("start_year", dt.date.today().year) or dt.date.today().year),
                "queries": [str(q) for q in cs.get("queries", []) if str(q).strip()],
                "df_data": _df_to_payload(cs.get("df_data")),
                "result": _df_to_payload(cs.get("result")),
            }
        )

    payload = {
        "schema_version": 1,
        "exported_at": dt.datetime.now().isoformat(timespec="seconds"),
        "app": "media_planner",
        "state": {
            "campaigns_df": _df_to_payload(st.session_state.get("campaigns_df")),
            "coeff_sets": coeff_sets_payload,
            "coeff_sets_links_new": _df_to_payload(st.session_state.get("coeff_sets_links_new")),
            "elasticity_df": _df_to_payload(st.session_state.get("elasticity_df")),
            "ak_rules_df": _df_to_payload(st.session_state.get("ak_rules_df")),
            "use_vat_budget_metrics": bool(st.session_state.get("use_vat_budget_metrics", True)),
            "use_ak_budget_metrics": bool(st.session_state.get("use_ak_budget_metrics", False)),
            "ak_mode": str(st.session_state.get("ak_mode", "percent")),
            "ak_fixed_month_wo_vat": float(st.session_state.get("ak_fixed_month_wo_vat", 200000.0) or 0.0),
            "planning_months_multiselect": list(st.session_state.get("planning_months_multiselect", [])),
            "metric_preset_key": str(st.session_state.get("metric_preset_key", "ecom")),
        },
    }
    return payload


def import_project_state_payload(payload: dict) -> tuple[bool, str]:
    if not isinstance(payload, dict):
        return False, "Неверный формат файла импорта."
    if payload.get("schema_version") != 1:
        return False, "Неподдерживаемая версия файла импорта."

    state = payload.get("state")
    if not isinstance(state, dict):
        return False, "В файле нет блока state."

    st.session_state["campaigns_df"] = _df_from_payload(state.get("campaigns_df"))
    st.session_state["coeff_sets_links_new"] = _df_from_payload(state.get("coeff_sets_links_new"))
    st.session_state["elasticity_df"] = _df_from_payload(state.get("elasticity_df"))
    st.session_state["ak_rules_df"] = _df_from_payload(state.get("ak_rules_df"))

    imported_sets = []
    for idx, cs in enumerate(state.get("coeff_sets", []), start=1):
        if not isinstance(cs, dict):
            continue
        imported_sets.append(
            {
                "id": int(cs.get("id", idx) or idx),
                "name": str(cs.get("name", f"Набор {idx}")),
                "type": str(cs.get("type", "Спрос (по запросам)")),
                "start_month": int(cs.get("start_month", 1) or 1),
                "start_year": int(cs.get("start_year", dt.date.today().year) or dt.date.today().year),
                "queries": [str(q) for q in cs.get("queries", []) if str(q).strip()],
                "df_data": _df_from_payload(cs.get("df_data")),
                "result": _df_from_payload(cs.get("result")),
            }
        )
    st.session_state["coeff_sets"] = imported_sets

    st.session_state["use_vat_budget_metrics"] = bool(state.get("use_vat_budget_metrics", True))
    st.session_state["use_ak_budget_metrics"] = bool(state.get("use_ak_budget_metrics", False))
    st.session_state["ak_mode"] = str(state.get("ak_mode", "percent"))
    st.session_state["ak_fixed_month_wo_vat"] = float(state.get("ak_fixed_month_wo_vat", 200000.0) or 0.0)
    st.session_state["planning_months_multiselect"] = list(state.get("planning_months_multiselect", []))
    preset_key = str(state.get("metric_preset_key", "ecom"))
    st.session_state["metric_preset_key"] = preset_key if preset_key in METRIC_PRESETS else "ecom"
    _bootstrap_reference_from_campaigns(st.session_state.get("campaigns_df"))

    return True, "Данные проекта импортированы."


if "_pending_project_import_payload" in st.session_state:
    _pending_payload = st.session_state.pop("_pending_project_import_payload")
    _ok, _msg = import_project_state_payload(_pending_payload)
    st.session_state["_pending_project_import_result"] = {"ok": bool(_ok), "msg": str(_msg)}


def queue_project_import_from_upload(uploaded_file, source_key: str) -> None:
    """
    Queue project import from a Streamlit UploadedFile only once per file content.
    Prevents rerun loops when uploader keeps the selected file between reruns.
    """
    sig_key = f"_project_upload_sig_{source_key}"
    if uploaded_file is None:
        st.session_state.pop(sig_key, None)
        return

    raw_bytes = uploaded_file.getvalue()
    payload_sig = hashlib.sha1(raw_bytes).hexdigest()
    if st.session_state.get(sig_key) == payload_sig:
        return

    st.session_state[sig_key] = payload_sig
    try:
        imported_payload = json.loads(raw_bytes.decode("utf-8"))
        st.session_state["_pending_project_import_payload"] = imported_payload
        st.rerun()
    except Exception as e:
        st.error(f"Не удалось прочитать файл проекта: {e}")


def make_ak_rules_signature(ak_rules_df: pd.DataFrame | None) -> tuple:
    if ak_rules_df is None or not isinstance(ak_rules_df, pd.DataFrame) or ak_rules_df.empty:
        return tuple()
    work = ak_rules_df.copy()
    if "min_budget_wo_vat" not in work.columns:
        work["min_budget_wo_vat"] = 0.0
    if "ak_percent" not in work.columns:
        work["ak_percent"] = 0.0
    work["min_budget_wo_vat"] = work["min_budget_wo_vat"].map(lambda v: round(parse_float_loose(v, 0.0), 4))
    work["ak_percent"] = work["ak_percent"].map(lambda v: round(parse_float_loose(v, 0.0), 4))
    work = work[["min_budget_wo_vat", "ak_percent"]].sort_values("min_budget_wo_vat", kind="stable")
    return tuple((float(r["min_budget_wo_vat"]), float(r["ak_percent"])) for _, r in work.iterrows())


def resolve_ak_rate(total_cost_wo_vat: float, ak_rules_df: pd.DataFrame) -> float:
    if ak_rules_df is None or ak_rules_df.empty:
        return 0.0
    rules = ak_rules_df.copy()
    rules["min_budget_wo_vat"] = rules["min_budget_wo_vat"].map(lambda v: parse_float_loose(v, 0.0))
    rules["ak_percent"] = rules["ak_percent"].map(lambda v: parse_float_loose(v, 0.0))
    rules = rules.sort_values("min_budget_wo_vat")
    matched = rules[rules["min_budget_wo_vat"] <= float(total_cost_wo_vat)]
    if matched.empty:
        return 0.0
    return max(0.0, float(matched.iloc[-1]["ak_percent"]) / 100.0)


def apply_budget_basis_metrics(
    df: pd.DataFrame,
    use_vat: bool,
    use_ak: bool = False,
    ak_mode: str = "percent",
    ak_rate_by_month: dict | None = None,
    default_ak_rate: float = 0.0,
    ak_fixed_by_month: dict | None = None,
    default_ak_fixed_wo_vat: float = 0.0,
    vat_rate: float = VAT_RATE,
) -> pd.DataFrame:
    out = df.copy()
    out["cost_with_vat"] = out["cost"].astype(float) * (1.0 + vat_rate)
    out["ak_rate"] = 0.0
    out["ak_fixed_month_wo_vat"] = 0.0
    out["ak_cost_wo_vat"] = 0.0

    if use_ak:
        if ak_mode == "fixed":
            if "month_num" in out.columns:
                if ak_fixed_by_month:
                    out["ak_fixed_month_wo_vat"] = out["month_num"].map(
                        lambda m: float(ak_fixed_by_month.get(int(m), default_ak_fixed_wo_vat))
                    )
                else:
                    out["ak_fixed_month_wo_vat"] = float(default_ak_fixed_wo_vat)
                month_sum = out.groupby("month_num")["cost"].transform("sum")
                share = np.where(month_sum > 0, out["cost"] / month_sum, 0.0)
                out["ak_cost_wo_vat"] = out["ak_fixed_month_wo_vat"] * share
            else:
                total_cost = float(out["cost"].sum())
                fixed_total = float(default_ak_fixed_wo_vat)
                out["ak_fixed_month_wo_vat"] = fixed_total
                share = np.where(total_cost > 0, out["cost"] / total_cost, 0.0)
                out["ak_cost_wo_vat"] = fixed_total * share
            out["ak_rate"] = np.where(out["cost"] > 0, out["ak_cost_wo_vat"] / out["cost"], 0.0)
        else:
            if "month_num" in out.columns and ak_rate_by_month:
                out["ak_rate"] = out["month_num"].map(
                    lambda m: float(ak_rate_by_month.get(int(m), default_ak_rate))
                )
            else:
                out["ak_rate"] = float(default_ak_rate)
            out["ak_rate"] = pd.to_numeric(out["ak_rate"], errors="coerce").fillna(float(default_ak_rate))
            out["ak_cost_wo_vat"] = out["cost"] * out["ak_rate"]

    out["ak_cost_with_vat"] = out["ak_cost_wo_vat"] * (1.0 + vat_rate)
    out["cost_with_vat_ak"] = (out["cost"] + out["ak_cost_wo_vat"]) * (1.0 + vat_rate)
    if use_ak:
        budget_basis = out["cost_with_vat_ak"] if use_vat else (out["cost"] + out["ak_cost_wo_vat"])
    else:
        budget_basis = out["cost_with_vat"] if use_vat else out["cost"]
    out["cpm"] = np.where(
        out["impressions"] > 0,
        budget_basis / (out["impressions"] / 1000.0),
        0.0,
    )
    out["cpa"] = np.where(
        out["conversions"] > 0,
        budget_basis / out["conversions"],
        0.0,
    )
    out["roas"] = np.where(
        budget_basis > 0,
        out["revenue"] / budget_basis,
        0.0,
    )
    out["drr"] = np.where(
        out["revenue"] > 0,
        budget_basis / out["revenue"],
        0.0,
    )
    out["budget_basis"] = budget_basis
    return out


def render_bottom_tab_switcher(current_tab: str, key_suffix: str) -> None:
    st.markdown("---")
    nav_id = f"bottom-tab-nav-{key_suffix}"
    st.markdown(
        f"""
        <div id="{nav_id}" class="bottom-tab-nav">
            <button type="button" class="bottom-tab-btn" data-tab-idx="0">Коэффициенты</button>
            <button type="button" class="bottom-tab-btn" data-tab-idx="1">Медиаплан</button>
            <button type="button" class="bottom-tab-btn" data-tab-idx="2">Диаграммы</button>
            <button type="button" class="bottom-tab-btn" data-tab-idx="3">Export/Import</button>
        </div>
        """,
        unsafe_allow_html=True,
    )
    components.html(
        f"""
        <script>
        (function() {{
            const tablists = Array.from(window.parent.document.querySelectorAll('[role="tablist"]'));
            if (!tablists.length) return;
            const topmostTablist = tablists
                .map(tl => ({{ tl, top: tl.getBoundingClientRect().top }}))
                .sort((a, b) => a.top - b.top)[0]?.tl;
            const nav = window.parent.document.getElementById("{nav_id}");
            if (!topmostTablist || !nav) return;

            const topTabs = topmostTablist.querySelectorAll('button[role="tab"]');
            const navBtns = nav.querySelectorAll('.bottom-tab-btn');
            if (!topTabs || topTabs.length < 4 || !navBtns.length) return;

            const syncActive = () => {{
                let activeIdx = 0;
                for (let i = 0; i < topTabs.length; i++) {{
                    if (topTabs[i].getAttribute('aria-selected') === 'true') {{
                        activeIdx = i;
                        break;
                    }}
                }}
                navBtns.forEach((btn, i) => {{
                    if (i === activeIdx) btn.classList.add('is-active');
                    else btn.classList.remove('is-active');
                }});
            }};

            navBtns.forEach((btn) => {{
                btn.onclick = () => {{
                    const idx = Number(btn.getAttribute('data-tab-idx') || 0);
                    if (topTabs[idx]) {{
                        topTabs[idx].click();
                        setTimeout(syncActive, 30);
                    }}
                }};
            }});

            syncActive();
            setTimeout(syncActive, 40);
            const obs = new MutationObserver(() => syncActive());
            obs.observe(topmostTablist, {{
                subtree: true,
                attributes: true,
                attributeFilter: ['aria-selected', 'class']
            }});
        }})();
        </script>
        """,
        height=0,
        width=0,
    )


DISPLAY_COL_RENAME = {
    "segment": "Сегмент",
    "campaign_type": "Название кампании",
    "system": "Рекламная система",
    "format": "Формат/таргетинги",
    "month_name": "Месяц",
    "impressions": "Показы",
    "ctr_pct": "CTR",
    "ctr": "CTR",
    "cpc": "CPC",
    "cr_pct": "CR",
    "cr": "CR",
    "aov": "AOV",
    "clicks": "Клики",
    "conversions": "Конверсии",
    "cost": "Бюджет",
    "cost_with_vat": "Бюджет с НДС",
    "cost_with_vat_ak": "Бюджет с НДС и АК",
    "ak_cost_wo_vat": "АК без НДС",
    "ak_rate_pct": "АК, %",
    "revenue": "Доход",
    "cpm": "CPM",
    "cpa": "CPO",
    "drr_pct": "ДРР",
    "drr": "ДРР",
    "ROAS": "ROAS",
    "sov_pct": "SOV, %",
    "available_capacity": "Доступная емкость",
    "new_clients_share_pct": "Доля новых клиентов, %",
}

METRIC_HELP = {
    "impressions": "Показы: количество рекламных показов.",
    "clicks": "Клики: Показы × CTR.",
    "ctr": "CTR, % = Клики / Показы × 100%.",
    "cpc": "CPC = Бюджет / Клики.",
    "cr": "CR, % = Конверсии / Клики × 100%.",
    "aov": "AOV = Доход / Конверсии.",
    "conversions": "Конверсии = Клики × CR.",
    "cost": "Бюджет = Клики × CPC.",
    "cost_with_vat": "Бюджет с НДС = Бюджет × 1.22 (при включенном учете НДС).",
    "cost_with_vat_ak": "Бюджет с НДС и АК = (Бюджет + АК без НДС) × 1.22.",
    "revenue": "Доход = Конверсии × AOV.",
    "cpm": "CPM = Бюджетная база / (Показы / 1000). База зависит от режима НДС.",
    "cpo": "CPO = Бюджетная база / Конверсии. База зависит от режима НДС.",
    "roas": "ROAS, % = Доход / Бюджетная база × 100%. База зависит от режима НДС.",
    "drr": "ДРР, % = Бюджетная база / Доход × 100%. База зависит от режима НДС.",
    "k_imp": "k_imp: коэффициент показов. Применяется к Показы.",
    "k_ctr": "k_ctr: коэффициент CTR. Применяется к CTR.",
    "k_cpc": "k_cpc: коэффициент CPC. Применяется к CPC.",
    "k_cr": "k_cr: коэффициент CR. Применяется к CR.",
    "k_aov": "k_aov: коэффициент AOV. Применяется к AOV.",
    "cpc_div": "Делитель влияния на CPC: k_cpc = 1 + (k_demand - 1) / div.",
    "ctr_div": "Делитель влияния на CTR: k_ctr = 1 - (k_demand - 1) / div.",
    "cr_div": "Делитель влияния на CR: k_cr = 1 - (k_demand - 1) / div.",
}


def mhelp(key: str, fallback: str = "") -> str:
    return METRIC_HELP.get(key, fallback)


def reorder_rows_with_segment_subtotals(
    df_show: pd.DataFrame,
    campaign_col: str,
    segment_col: str,
) -> pd.DataFrame:
    """Order rows as: campaigns by segment -> segment subtotal -> ... -> grand total."""
    if df_show is None or df_show.empty:
        return df_show
    if campaign_col not in df_show.columns or segment_col not in df_show.columns:
        return df_show

    work = df_show.copy()
    campaign_vals = work[campaign_col].astype(str)
    is_total_any = campaign_vals.str.startswith("того")
    is_total_grand = campaign_vals.eq("того")
    is_total_segment = is_total_any & ~is_total_grand
    is_campaign = ~is_total_any

    seg_order = ["B2C", "B2B"]
    segments_present = work.loc[is_campaign, segment_col].astype(str).unique().tolist()
    for seg in segments_present:
        if seg not in seg_order:
            seg_order.append(seg)

    ordered_parts = []
    used_idx: set[int] = set()
    for seg in seg_order:
        seg_campaigns = work[is_campaign & (work[segment_col].astype(str) == seg)]
        if not seg_campaigns.empty:
            ordered_parts.append(seg_campaigns)
            used_idx.update(seg_campaigns.index.tolist())
            seg_total = work[is_total_segment & (campaign_vals == f"того {seg}")]
            if not seg_total.empty:
                ordered_parts.append(seg_total)
                used_idx.update(seg_total.index.tolist())

    # Fallback for any rows that didn't get included.
    if ordered_parts:
        ordered = pd.concat(ordered_parts, ignore_index=True)
        campaign_rows_left = work[is_campaign & ~work.index.isin(list(used_idx))]
        seg_totals_left = work[is_total_segment & ~work.index.isin(list(used_idx))]
        if not campaign_rows_left.empty:
            ordered = pd.concat([ordered, campaign_rows_left], ignore_index=True)
            used_idx.update(campaign_rows_left.index.tolist())
        if not seg_totals_left.empty:
            ordered = pd.concat([ordered, seg_totals_left], ignore_index=True)
            used_idx.update(seg_totals_left.index.tolist())
    else:
        ordered = work[is_campaign].copy()
        used_idx.update(ordered.index.tolist())

    grand_total = work[is_total_grand & ~work.index.isin(list(used_idx))]
    if not grand_total.empty:
        ordered = pd.concat([ordered, grand_total], ignore_index=True)
    return ordered


def forecast_ets_like(
    series: list[float],
    periods: int = 12,
    season_length: int = 12,
    alpha: float = 0.35,
    beta: float = 0.10,
    gamma: float = 0.25,
) -> list[float]:
    """
    ETS forecast with statsmodels when available; otherwise fallback to a
    lightweight additive Holt-Winters implementation.
    """
    vals = [float(x) for x in series if pd.notna(x)]
    n = len(vals)
    if n == 0:
        return [0.0] * periods

    # Primary path: statsmodels ETS (closest to Excel FORECAST.ETS behavior).
    if HAS_STATSMODELS and n >= season_length * 2:
        try:
            # Keep smoothing params inside valid open interval for stable fitting.
            a = min(max(float(alpha), 1e-4), 0.9999)
            b = min(max(float(beta), 1e-4), 0.9999)
            g = min(max(float(gamma), 1e-4), 0.9999)

            model = ExponentialSmoothing(
                vals,
                trend="add",
                seasonal="add",
                seasonal_periods=season_length,
                initialization_method="estimated",
            )
            # Use user-selected alpha/beta/gamma (instead of auto-optimized fit).
            fitted = model.fit(
                optimized=False,
                smoothing_level=a,
                smoothing_trend=b,
                smoothing_seasonal=g,
            )
            fc = fitted.forecast(periods)
            return [max(0.0, float(v)) for v in fc]
        except Exception:
            # fallback below
            pass

    if n < season_length * 2:
        # For short history, repeat last available seasonal pattern.
        tail = vals[-season_length:] if n >= season_length else vals
        rep = []
        while len(rep) < periods:
            rep.extend(tail)
        return [max(0.0, float(v)) for v in rep[:periods]]

    n_seasons = n // season_length
    season_avgs = []
    for j in range(n_seasons):
        chunk = vals[j * season_length:(j + 1) * season_length]
        season_avgs.append(float(np.mean(chunk)))

    seasonals = [0.0] * season_length
    for i in range(season_length):
        acc = 0.0
        cnt = 0
        for j in range(n_seasons):
            idx = j * season_length + i
            if idx < n:
                acc += vals[idx] - season_avgs[j]
                cnt += 1
        seasonals[i] = acc / max(cnt, 1)

    level = vals[0] - seasonals[0]
    trend = (
        ((vals[season_length] - seasonals[0]) - (vals[0] - seasonals[0])) / season_length
        if n > season_length else 0.0
    )

    for t in range(n):
        s = seasonals[t % season_length]
        prev_level = level
        level = alpha * (vals[t] - s) + (1 - alpha) * (level + trend)
        trend = beta * (level - prev_level) + (1 - beta) * trend
        seasonals[t % season_length] = gamma * (vals[t] - level) + (1 - gamma) * s

    fc = []
    for m in range(1, periods + 1):
        idx = (n + m - 1) % season_length
        yhat = level + m * trend + seasonals[idx]
        fc.append(max(0.0, float(yhat)))
    return fc


# ---------- ЭКСПОРТ В ШАБЛОН EXCEL ----------

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATHS_ECOM = [
    os.path.join(BASE_DIR, "templates", "Шаблоны МП.xlsx"),
    os.path.join(BASE_DIR, "Шаблоны МП.xlsx"),
    os.path.join(BASE_DIR, "Shablony-MP.xlsx"),
]
TEMPLATE_PATHS_DIY = [
    os.path.join(BASE_DIR, "templates", "Шаблоны МП DIY.xlsx"),
    os.path.join(BASE_DIR, "Шаблоны МП DIY.xlsx"),
    os.path.join(BASE_DIR, "templates", "Shablony-MP-DIY.xlsx"),
]
TEMPLATE_PATHS_BY_KIND = {
    "ecom": TEMPLATE_PATHS_ECOM,
    "diy": TEMPLATE_PATHS_DIY,
}


def build_excel_from_template(df_all: pd.DataFrame,
                              campaigns: pd.DataFrame,
                              selected_month_nums: list[int],
                              template_kind: str = "ecom",
                              compact_empty_rows: bool = True) -> BytesIO:

    template_paths = TEMPLATE_PATHS_BY_KIND.get(template_kind, TEMPLATE_PATHS_ECOM)
    existing_templates = [p for p in template_paths if os.path.exists(p)]
    template_path = (
        max(existing_templates, key=os.path.getmtime)
        if existing_templates
        else template_paths[0]
    )

    wb = load_workbook(template_path)

    periods_candidates = [s for s in wb.sheetnames if str(s).endswith("_Periods")]
    total_candidates = [s for s in wb.sheetnames if str(s).endswith("_Total")]

    if periods_candidates:
        ws = wb[periods_candidates[0]]
    elif len(wb.worksheets) >= 2:
        ws = wb.worksheets[1]
    else:
        ws = wb.worksheets[0]

    if total_candidates:
        ws_total = wb[total_candidates[0]]
    else:
        ws_total = wb.worksheets[0]

    DEFAULT_START_ROW_JAN = 3
    DEFAULT_BLOCK_STEP = 13
    DEFAULT_ROWS_PER_MONTH = 10
    DEFAULT_START_ROW_TOTAL = 3

    COL_SYSTEM = "B"
    COL_FORMAT = "C"
    COL_TARGETING = "D"
    COL_PERIOD = "F"
    COL_MODEL = "G"

    COL_CPC = "H"
    COL_AK = "N"
    COL_IMPRESSIONS = "S"
    COL_CTR = "AA"
    COL_CR = "AE"
    # In DIY template AN is a formula column ("share of revenue"), while AOV is in AL.
    COL_AOV = "AL" if template_kind == "diy" else "AN"
    COL_NEW_CLIENTS_SHARE = "W" if template_kind == "diy" else None
    COL_AVAILABLE_CAPACITY = "X" if template_kind == "diy" else None

    def _safe_text(v) -> str:
        if v is None:
            return ""
        return str(v).strip().lower()

    def _norm_segment(v) -> str:
        t = _safe_text(v).replace(" ", "")
        if "b2b" in t:
            return "B2B"
        if "b2c" in t:
            return "B2C"
        return ""

    def _is_header_row(sheet, r: int) -> bool:
        b = sheet[f"{COL_SYSTEM}{r}"].value
        c = sheet[f"{COL_FORMAT}{r}"].value
        d = sheet[f"{COL_TARGETING}{r}"].value
        f = sheet[f"{COL_PERIOD}{r}"].value
        s = sheet[f"{COL_IMPRESSIONS}{r}"].value
        if not all(v is not None and str(v).strip() != "" for v in (b, c, d, f, s)):
            return False
        # Exclude subtotal/total rows where template keeps dashes in C/D.
        if str(c).strip() == "-" and str(d).strip() == "-":
            return False
        return True

    def _detect_periods_layout() -> tuple[int, int, int]:
        header_rows = []
        for r in range(1, min(ws.max_row, 500) + 1):
            if _is_header_row(ws, r):
                header_rows.append(r)
                if len(header_rows) >= 3:
                    break

        if not header_rows:
            return DEFAULT_START_ROW_JAN, DEFAULT_BLOCK_STEP, DEFAULT_ROWS_PER_MONTH

        header_row = header_rows[0]
        next_header_row = header_rows[1] if len(header_rows) > 1 else None
        start_row = header_row + 1

        if next_header_row:
            rows_per_month = max(1, next_header_row - start_row - 2)
            block_step = next_header_row - header_row
        else:
            search_end = min(ws.max_row, start_row + 50)
            total_row = None
            for r in range(start_row, search_end + 1):
                c = _safe_text(ws[f"{COL_FORMAT}{r}"].value)
                d = _safe_text(ws[f"{COL_TARGETING}{r}"].value)
                b = _safe_text(ws[f"{COL_SYSTEM}{r}"].value)
                if b and c == "-" and d == "-":
                    total_row = r
                    break
            rows_per_month = max(1, total_row - start_row) if total_row else DEFAULT_ROWS_PER_MONTH
            block_step = max(DEFAULT_BLOCK_STEP, rows_per_month + 2)

        return start_row, block_step, rows_per_month

    def _detect_total_layout(default_rows_per_month: int) -> tuple[int, int]:
        header_row = None
        for r in range(1, min(ws_total.max_row, 200) + 1):
            if _is_header_row(ws_total, r):
                header_row = r
                break

        if header_row is None:
            return DEFAULT_START_ROW_TOTAL, default_rows_per_month

        start_row = header_row + 1
        total_like_row = None
        for r in range(start_row, min(ws_total.max_row, start_row + 100) + 1):
            c = _safe_text(ws_total[f"{COL_FORMAT}{r}"].value)
            d = _safe_text(ws_total[f"{COL_TARGETING}{r}"].value)
            b = _safe_text(ws_total[f"{COL_SYSTEM}{r}"].value)
            if b and c == "-" and d == "-":
                total_like_row = r
                break

        rows_per_month = max(1, total_like_row - start_row) if total_like_row else default_rows_per_month
        return start_row, rows_per_month

    START_ROW_JAN, BLOCK_STEP, ROWS_PER_MONTH = _detect_periods_layout()
    START_ROW_TOTAL, ROWS_PER_MONTH_TOTAL = _detect_total_layout(ROWS_PER_MONTH)

    def _write_period_row(row_excel: int, camp_row: pd.Series | None, row_data: pd.Series | None, period_str: str):
        ws.row_dimensions[row_excel].hidden = False
        if camp_row is None or row_data is None:
            ws[f"{COL_SYSTEM}{row_excel}"] = None
            ws[f"{COL_FORMAT}{row_excel}"] = None
            ws[f"{COL_TARGETING}{row_excel}"] = None
            ws[f"{COL_PERIOD}{row_excel}"] = None
            ws[f"{COL_MODEL}{row_excel}"] = None
            ws[f"{COL_IMPRESSIONS}{row_excel}"] = None
            ws[f"{COL_CTR}{row_excel}"] = None
            ws[f"{COL_CPC}{row_excel}"] = None
            ws[f"{COL_AK}{row_excel}"] = None
            ws[f"{COL_CR}{row_excel}"] = None
            ws[f"{COL_AOV}{row_excel}"] = None
            if COL_NEW_CLIENTS_SHARE:
                ws[f"{COL_NEW_CLIENTS_SHARE}{row_excel}"] = None
            if COL_AVAILABLE_CAPACITY:
                ws[f"{COL_AVAILABLE_CAPACITY}{row_excel}"] = None
            return

        impressions = float(row_data["impressions"])
        ctr = float(row_data["ctr"])
        cpc = float(row_data["cpc"])
        cr = float(row_data["cr"])
        aov = float(row_data["aov"])
        ak_rate = float(row_data.get("ak_rate", 0.0))
        if pd.isna(ak_rate):
            ak_rate = 0.0
        new_clients_share_pct = float(row_data.get("new_clients_share_pct", 0.0))
        if pd.isna(new_clients_share_pct):
            new_clients_share_pct = 0.0
        # App stores this metric in percents (e.g. 25), while Excel percent cells
        # typically expect a fraction (0.25) to display 25%.
        if new_clients_share_pct > 1:
            new_clients_share_pct = new_clients_share_pct / 100.0
        available_capacity = float(row_data.get("available_capacity", 0.0))
        if pd.isna(available_capacity):
            available_capacity = 0.0

        ws[f"{COL_SYSTEM}{row_excel}"] = camp_row.get("system", "")
        ws[f"{COL_FORMAT}{row_excel}"] = camp_row.get("campaign_type", "")
        ws[f"{COL_TARGETING}{row_excel}"] = camp_row.get("format", "")
        ws[f"{COL_PERIOD}{row_excel}"] = period_str
        ws[f"{COL_MODEL}{row_excel}"] = "CPC"
        ws[f"{COL_IMPRESSIONS}{row_excel}"] = impressions
        ws[f"{COL_CTR}{row_excel}"] = ctr
        ws[f"{COL_CPC}{row_excel}"] = cpc
        ws[f"{COL_AK}{row_excel}"] = ak_rate
        ws[f"{COL_CR}{row_excel}"] = cr
        ws[f"{COL_AOV}{row_excel}"] = aov
        if COL_NEW_CLIENTS_SHARE:
            ws[f"{COL_NEW_CLIENTS_SHARE}{row_excel}"] = new_clients_share_pct
        if COL_AVAILABLE_CAPACITY:
            ws[f"{COL_AVAILABLE_CAPACITY}{row_excel}"] = available_capacity

    def _write_total_row(row_excel: int, camp_row: pd.Series | None):
        ws_total.row_dimensions[row_excel].hidden = False
        if camp_row is None:
            ws_total[f"{COL_SYSTEM}{row_excel}"] = None
            ws_total[f"{COL_FORMAT}{row_excel}"] = None
            ws_total[f"{COL_TARGETING}{row_excel}"] = None
            return
        ws_total[f"{COL_SYSTEM}{row_excel}"] = camp_row.get("system", "")
        ws_total[f"{COL_FORMAT}{row_excel}"] = camp_row.get("campaign_type", "")
        ws_total[f"{COL_TARGETING}{row_excel}"] = camp_row.get("format", "")

    def _collect_diy_period_rows(block_start_row: int) -> tuple[list[int], list[int]] | None:
        scan_from = block_start_row
        scan_to = min(ws.max_row, block_start_row + BLOCK_STEP + 8)
        row_b2c_total = None
        row_b2b_total = None
        for r in range(scan_from, scan_to + 1):
            t = _safe_text(ws[f"{COL_SYSTEM}{r}"].value)
            c = _safe_text(ws[f"{COL_FORMAT}{r}"].value)
            d = _safe_text(ws[f"{COL_TARGETING}{r}"].value)
            if c != "-" or d != "-":
                continue
            if "b2c" in t and row_b2c_total is None:
                row_b2c_total = r
            elif "b2b" in t and row_b2b_total is None:
                row_b2b_total = r

        if row_b2c_total is None or row_b2b_total is None or row_b2b_total <= row_b2c_total:
            return None

        b2c_rows = list(range(block_start_row, row_b2c_total))
        b2b_rows = list(range(row_b2c_total + 1, row_b2b_total))
        return b2c_rows, b2b_rows

    def _collect_diy_total_rows() -> tuple[list[int], list[int]] | None:
        scan_from = START_ROW_TOTAL
        scan_to = min(ws_total.max_row, START_ROW_TOTAL + ROWS_PER_MONTH_TOTAL + 40)
        row_b2c_total = None
        row_b2b_total = None
        for r in range(scan_from, scan_to + 1):
            t = _safe_text(ws_total[f"{COL_SYSTEM}{r}"].value)
            c = _safe_text(ws_total[f"{COL_FORMAT}{r}"].value)
            d = _safe_text(ws_total[f"{COL_TARGETING}{r}"].value)
            if c != "-" or d != "-":
                continue
            if "b2c" in t and row_b2c_total is None:
                row_b2c_total = r
            elif "b2b" in t and row_b2b_total is None:
                row_b2b_total = r

        if row_b2c_total is None or row_b2b_total is None or row_b2b_total <= row_b2c_total:
            return None

        b2c_rows = list(range(START_ROW_TOTAL, row_b2c_total))
        b2b_rows = list(range(row_b2c_total + 1, row_b2b_total))
        return b2c_rows, b2b_rows

    year = dt.date.today().year

    for block_index, month_num in enumerate(selected_month_nums):
        block_start_row = START_ROW_JAN + block_index * BLOCK_STEP
        start = dt.date(year, month_num, 1)
        end = dt.date(year, 12, 31) if month_num == 12 else dt.date(year, month_num + 1, 1) - dt.timedelta(days=1)
        period_str = f"{start.strftime('%d.%m.%Y')} - {end.strftime('%d.%m.%Y')}"
        df_month = df_all[df_all["month_num"] == month_num]

        diy_period_rows = _collect_diy_period_rows(block_start_row) if (template_kind == "diy" and "segment" in campaigns.columns) else None

        if diy_period_rows is not None:
            b2c_rows, b2b_rows = diy_period_rows
            camps = campaigns.copy()
            camps["_seg"] = camps["segment"].map(_norm_segment)
            camps_b2c = camps[camps["_seg"] == "B2C"]
            camps_b2b = camps[camps["_seg"] == "B2B"]

            for row_excel, (_, camp) in zip(b2c_rows, camps_b2c.iterrows()):
                row_data = df_month[df_month["campaign_type"] == camp["campaign_type"]]
                _write_period_row(row_excel, camp, (None if row_data.empty else row_data.iloc[0]), period_str)
            for row_excel in b2c_rows[len(camps_b2c):]:
                _write_period_row(row_excel, None, None, period_str)

            for row_excel, (_, camp) in zip(b2b_rows, camps_b2b.iterrows()):
                row_data = df_month[df_month["campaign_type"] == camp["campaign_type"]]
                _write_period_row(row_excel, camp, (None if row_data.empty else row_data.iloc[0]), period_str)
            for row_excel in b2b_rows[len(camps_b2b):]:
                _write_period_row(row_excel, None, None, period_str)
        else:
            for i, (_, camp) in enumerate(campaigns.iterrows()):
                if i >= ROWS_PER_MONTH:
                    break
                row_excel = block_start_row + i
                row_data = df_month[df_month["campaign_type"] == camp["campaign_type"]]
                _write_period_row(row_excel, camp, (None if row_data.empty else row_data.iloc[0]), period_str)
            for i in range(min(len(campaigns), ROWS_PER_MONTH), ROWS_PER_MONTH):
                _write_period_row(block_start_row + i, None, None, period_str)

    if compact_empty_rows:
        rows_to_hide_periods = []
        for block_index, _ in enumerate(selected_month_nums):
            block_start_row = START_ROW_JAN + block_index * BLOCK_STEP
            for i in range(ROWS_PER_MONTH):
                row_excel = block_start_row + i
                is_empty_main = (
                    ws[f"{COL_SYSTEM}{row_excel}"].value in (None, "")
                    and ws[f"{COL_FORMAT}{row_excel}"].value in (None, "")
                    and ws[f"{COL_TARGETING}{row_excel}"].value in (None, "")
                )
                is_empty_metrics = (
                    ws[f"{COL_IMPRESSIONS}{row_excel}"].value in (None, "")
                    and ws[f"{COL_CTR}{row_excel}"].value in (None, "")
                    and ws[f"{COL_CPC}{row_excel}"].value in (None, "")
                    and ws[f"{COL_AK}{row_excel}"].value in (None, "")
                    and ws[f"{COL_CR}{row_excel}"].value in (None, "")
                    and ws[f"{COL_AOV}{row_excel}"].value in (None, "")
                    and (not COL_NEW_CLIENTS_SHARE or ws[f"{COL_NEW_CLIENTS_SHARE}{row_excel}"].value in (None, ""))
                    and (not COL_AVAILABLE_CAPACITY or ws[f"{COL_AVAILABLE_CAPACITY}{row_excel}"].value in (None, ""))
                )
                if is_empty_main and is_empty_metrics:
                    rows_to_hide_periods.append(row_excel)
        for row_idx in sorted(set(rows_to_hide_periods)):
            ws.row_dimensions[row_idx].hidden = True

    diy_total_rows = _collect_diy_total_rows() if (template_kind == "diy" and "segment" in campaigns.columns) else None

    if diy_total_rows is not None:
        b2c_rows_t, b2b_rows_t = diy_total_rows
        total_rows_for_hide = list(b2c_rows_t) + list(b2b_rows_t)
        camps = campaigns.copy()
        camps["_seg"] = camps["segment"].map(_norm_segment)
        camps_b2c = camps[camps["_seg"] == "B2C"]
        camps_b2b = camps[camps["_seg"] == "B2B"]

        for row_excel, (_, camp) in zip(b2c_rows_t, camps_b2c.iterrows()):
            _write_total_row(row_excel, camp)
        for row_excel in b2c_rows_t[len(camps_b2c):]:
            _write_total_row(row_excel, None)

        for row_excel, (_, camp) in zip(b2b_rows_t, camps_b2b.iterrows()):
            _write_total_row(row_excel, camp)
        for row_excel in b2b_rows_t[len(camps_b2b):]:
            _write_total_row(row_excel, None)
    else:
        total_rows_for_hide = list(range(START_ROW_TOTAL, START_ROW_TOTAL + ROWS_PER_MONTH_TOTAL))
        for i in range(ROWS_PER_MONTH_TOTAL):
            row_excel = START_ROW_TOTAL + i
            if i < len(campaigns):
                _write_total_row(row_excel, campaigns.iloc[i])
            else:
                _write_total_row(row_excel, None)

    if compact_empty_rows:
        rows_to_hide_total = []
        for row_excel in total_rows_for_hide:
            if (
                ws_total[f"{COL_SYSTEM}{row_excel}"].value in (None, "")
                and ws_total[f"{COL_FORMAT}{row_excel}"].value in (None, "")
                and ws_total[f"{COL_TARGETING}{row_excel}"].value in (None, "")
            ):
                rows_to_hide_total.append(row_excel)
        for row_idx in sorted(set(rows_to_hide_total)):
            ws_total.row_dimensions[row_idx].hidden = True

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output
def resolve_template_path(template_kind: str = "ecom") -> str | None:
    template_paths = TEMPLATE_PATHS_BY_KIND.get(template_kind, TEMPLATE_PATHS_ECOM)
    existing_templates = [p for p in template_paths if os.path.exists(p)]
    if not existing_templates:
        return None
    return max(existing_templates, key=os.path.getmtime)


# ---------- ЗАГОЛОВОК  ТАБЫ ----------

coeff_header_uri = image_file_to_data_uri(os.path.join(BASE_DIR, "assets", "coeff_header_mascot.png"))
plan_header_uri = image_file_to_data_uri(os.path.join(BASE_DIR, "assets", "plan_header_mascot.png"))
charts_header_uri = image_file_to_data_uri(os.path.join(BASE_DIR, "assets", "charts_header_mascot.png"))
export_header_uri = image_file_to_data_uri(os.path.join(BASE_DIR, "assets", "export_header_mascot.png"))
top_header_uri = coeff_header_uri or plan_header_uri or charts_header_uri or export_header_uri
if top_header_uri:
    top_header_html = """
        <style>
        .app-top-hero {
            position: relative;
            border: 1px solid #1D2A44;
            border-radius: 14px;
            overflow: hidden;
            margin: 0.15rem 0 0.85rem 0;
            min-height: 250px;
            background: #0B1220;
        }
        .app-top-hero-bg {
            position: absolute;
            inset: 0;
            background-image: url("__TOP_HEADER_URI__");
            background-size: contain;
            background-repeat: no-repeat;
            background-position: calc(100% - 20px) center;
            opacity: 0.0;
            transform: scale(1.0);
            filter: none;
            transition: opacity 220ms ease;
        }
        .app-top-hero-fade {
            position: absolute;
            inset: 0;
            background: linear-gradient(90deg, rgba(8, 16, 30, 0.90) 0%, rgba(8, 16, 30, 0.48) 42%, rgba(8, 16, 30, 0.02) 72%, rgba(8, 16, 30, 0.00) 100%);
        }
        .app-top-hero-content {
            position: relative;
            z-index: 2;
            padding: 14px 16px 14px 16px;
        }
        .app-top-hero-content .intro-card {
            margin: 0.2rem 0 0.2rem 0;
            padding: 12px 14px;
            width: fit-content;
            max-width: calc(100% - 360px);
            border-radius: 12px;
            border: 1px solid #3D8EF0;
            border-left: 4px solid #00CDC5;
            background: linear-gradient(180deg, rgba(0, 102, 224, 0.20) 0%, rgba(17, 26, 46, 0.84) 100%);
            box-shadow: 0 8px 18px rgba(0, 0, 0, 0.20);
        }
        @media (max-width: 1150px) {
            .app-top-hero-content .intro-card {
                width: 100%;
            }
            .app-top-hero-bg {
                background-position: calc(100% - 12px) center;
            }
        }
        .app-top-hero-content .intro-card p {
            margin: 0 0 6px 0;
            color: #EAF0FF;
            line-height: 1.38;
            font-size: clamp(0.93rem, 0.92vw, 1.06rem);
            overflow-wrap: anywhere;
        }
        .app-top-hero-content .intro-card p.one-line {
            white-space: normal;
        }
        .app-top-hero-content .intro-card p:last-child {
            margin-bottom: 0;
        }
        @media (max-width: 1150px) {
            .app-top-hero-content .intro-card p.one-line {
                white-space: normal;
            }
        }
        </style>
        <div class="app-top-hero">
            <div id="app-top-hero-bg" class="app-top-hero-bg"></div>
            <div class="app-top-hero-fade"></div>
            <div class="app-top-hero-content">
                <h1 style="font-weight: 700; letter-spacing: 0.02em; margin: 0 0 6px 0; max-width: 70%;">
                    Медиапланер <span style="color: #00CDC5; font-size: 1.25em;">E-promo</span>
                </h1>
                <div class="intro-card">
                    <p><span style="font-weight: 800; color: #9EC5FF;">Что это:</span>
                    инструмент для расчета медиаплана на выбранный период (от 1 до 12 месяцев) по типам рекламных кампаний.</p>
                    <p class="one-line"><span style="font-weight: 800; color: #9EC5FF;">Зачем нужен:</span>
                    чтобы упростить алгоритм расчета медиаплана, сократить время на первичный расчет и ускорить внесение последующих правок.</p>
                </div>
            </div>
        </div>
        """
    st.markdown(
        top_header_html.replace("__TOP_HEADER_URI__", top_header_uri),
        unsafe_allow_html=True,
    )
else:
    st.markdown(
        """
        <h1 style="font-weight: 700; letter-spacing: 0.02em;">
            Медиапланер <span style="color: #00CDC5; font-size: 1.25em;">E-promo</span>
        </h1>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div style="
            margin: 0.2rem 0 0.8rem 0;
            padding: 12px 14px;
            border-radius: 12px;
            border: 1px solid #3D8EF0;
            border-left: 4px solid #00CDC5;
            background: linear-gradient(180deg, rgba(0, 102, 224, 0.18) 0%, rgba(17, 26, 46, 0.82) 100%);
            box-shadow: 0 8px 18px rgba(0, 0, 0, 0.20);
        ">
            <p style="margin: 0 0 6px 0; color: #EAF0FF; line-height: 1.45;">
                <span style="font-weight: 800; color: #9EC5FF;">??? ???:</span>
                ?????????? ??? ??????? ?????????? ?? ????????? ?????? (?? 1 ?? 12 ???????) ?? ????? ????????? ????????.
            </p>
            <p style="margin: 0; color: #EAF0FF; line-height: 1.45;">
                <span style="font-weight: 800; color: #9EC5FF;">????? ?????:</span>
                ????? ????????? ???????? ??????? ??????????, ????????? ????? ?? ????????? ?????? ? ???????? ???????? ??????????? ??????.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

ui_compact_tables = True
ui_editor_row_height = 28
ui_table_row_px = 29
ui_table_font_px = "12px"

st.markdown(
    f"""
    <style>
    [data-testid="stDataFrame"] [role="gridcell"],
    [data-testid="stDataFrame"] [role="columnheader"] {{
        font-size: {ui_table_font_px};
        line-height: 1.0 !important;
        padding-top: 2px !important;
        padding-bottom: 2px !important;
        min-height: {ui_table_row_px}px !important;
        height: {ui_table_row_px}px !important;
        max-height: {ui_table_row_px}px !important;
        box-sizing: border-box !important;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

tab_coeffs, tab_plan, tab_charts, tab_export = st.tabs(
    ["Коэффициенты", "Медиаплан", "Диаграммы", "Export/Import"]
)
if top_header_uri:
    header_tab_bg_map_json = json.dumps(
        {
            0: coeff_header_uri or "",
            1: plan_header_uri or "",
            2: charts_header_uri or "",
            3: export_header_uri or "",
        },
        ensure_ascii=False,
    )
    components.html(
        """
        <script>
        (function() {
            const TAB_BG_MAP = __TAB_BG_MAP__;
            const VISIBLE_OPACITY = "1.00";
            const COEFF_TAB_INDEX = 0;
            const PLAN_TAB_INDEX = 1;
            const CHARTS_TAB_INDEX = 2;
            const EXPORT_TAB_INDEX = 3;
            const BASE_BG_POSITION = "calc(100% - 20px) center";
            const COEFF_BG_POSITION = "calc(100% + 95px) calc(50% + 77px)";
            const PLAN_BG_POSITION = "calc(100% + 120px) center";
            const CHARTS_BG_SIZE = "47% auto";
            const CHARTS_BG_POSITION = "calc(100% + 120px) calc(50% + 30px)";
            const EXPORT_BG_SIZE = "24.8% auto";
            const EXPORT_BG_POSITION = "calc(100% - 20px) calc(50% + 30px)";

            const getTopTablist = (doc) => {
                const tablists = Array.from(doc.querySelectorAll('[role="tablist"]'));
                if (!tablists.length) return null;
                return tablists
                    .map(tl => ({ tl, top: tl.getBoundingClientRect().top }))
                    .sort((a, b) => a.top - b.top)[0]?.tl || null;
            };

            const applyHeaderImage = () => {
                const doc = window.parent.document;
                const bg = doc.getElementById("app-top-hero-bg");
                if (!bg) return false;
                const topTablist = getTopTablist(doc);
                if (!topTablist) return false;
                const tabs = topTablist.querySelectorAll('button[role="tab"]');
                if (!tabs || !tabs.length) return false;

                let activeIdx = 0;
                for (let i = 0; i < tabs.length; i++) {
                    if (tabs[i].getAttribute("aria-selected") === "true") {
                        activeIdx = i;
                        break;
                    }
                }
                const key = String(activeIdx);
                const imgUrl = TAB_BG_MAP[key] || "";
                bg.style.backgroundImage = imgUrl ? `url(${imgUrl})` : "";
                if (activeIdx === COEFF_TAB_INDEX) {
                    bg.style.backgroundSize = "52% auto";
                    bg.style.backgroundPosition = COEFF_BG_POSITION;
                } else if (activeIdx === PLAN_TAB_INDEX) {
                    bg.style.backgroundSize = "58% auto";
                    bg.style.backgroundPosition = PLAN_BG_POSITION;
                } else if (activeIdx === CHARTS_TAB_INDEX) {
                    bg.style.backgroundSize = CHARTS_BG_SIZE;
                    bg.style.backgroundPosition = CHARTS_BG_POSITION;
                } else if (activeIdx === EXPORT_TAB_INDEX) {
                    bg.style.backgroundSize = EXPORT_BG_SIZE;
                    bg.style.backgroundPosition = EXPORT_BG_POSITION;
                } else {
                    bg.style.backgroundSize = "58% auto";
                    bg.style.backgroundPosition = BASE_BG_POSITION;
                }
                bg.style.opacity = imgUrl ? VISIBLE_OPACITY : "0.0";
                return true;
            };

            let attempts = 0;
            const init = () => {
                applyHeaderImage();
                attempts += 1;
                if (attempts < 30) setTimeout(init, 120);
            };
            init();

            const topTablist = getTopTablist(window.parent.document);
            if (topTablist) {
                const obs = new MutationObserver(() => applyHeaderImage());
                obs.observe(topTablist, {
                    subtree: true,
                    attributes: true,
                    attributeFilter: ["aria-selected", "class"],
                });
            }
        })();
        </script>
        """.replace("__TAB_BG_MAP__", header_tab_bg_map_json),
        height=0,
        width=0,
    )

# Быстрый импорт доступен всегда (даже до первого расчета на вкладке "Медиаплан").
with st.sidebar:
    st.markdown("### Быстрый импорт проекта")
    st.caption("Загрузите JSON проекта сразу после запуска приложения.")
    uploaded_project_quick = st.file_uploader(
        "Импорт проекта (JSON)",
        type=["json"],
        key="upload_project_json_quick",
    )
    queue_project_import_from_upload(uploaded_project_quick, "quick_sidebar")
# ====================================================================
#                        ТАБ «КОЭФФИЦИЕНТЫ»
# ====================================================================

with tab_coeffs:
    st.markdown(
        """
        <div class="tab-intro">
            <p>1) Создайте набор коэффициентов и выберите тип: <b>Спрос</b>, <b>AOV</b>, <b>Кастомный набор</b> или <b>Медийные хвосты</b>.</p>
            <p>2) Укажите стартовый месяц и год, затем заполните значения по периодам в таблице по данным из <a href="https://wordstat.yandex.ru/" target="_blank">Вордстат</a>.</p>
            <p>3) Нажмите кнопку применения/расчета, чтобы сохранить набор и получить итоговые коэффициенты.</p>
            <p>4) Повторите для всех наборов, которые планируете использовать в «Медиаплане».</p>
            <p>5) Проверьте, что у каждого набора заполнены данные, иначе он не будет влиять на расчет.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    ui_section_title("Коэффициенты сезонности")
    st.caption(
        "Создавайте один или несколько наборов коэффициентов сезонности: "
        "по спросу (запросы) или по AOV (средний чек). Для каждого набора считайте индексы по месяцам."
    )
    st.markdown(
        """
        <div style="
            margin: 8px 0 14px 0;
            display: inline-block;
            width: fit-content;
            max-width: 100%;
            white-space: nowrap;
            padding: 10px 12px;
            border-radius: 10px;
            border: 1px solid #FF8A66;
            background: rgba(255, 99, 51, 0.14);
            color: #FFD9CC;
            font-weight: 600;
        ">
            Важно: для корректного расчета медиаплана сначала обязательно рассчитайте минимум два набора:
            «Сезонность» (спрос) и «AOV» (средний чек).
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.caption(
        "сточники данных: для «Спрос» — Wordstat; для «AOV» — фактический средний чек из аналитики; "
        "для «Медийных хвостов» — обратитесь к специалистам по медийной рекламе для расчета влияния."
    )

    # ---------------- ВСПОМОГАТЕЛЬНЫЕ ФУНКЦ ----------------

    def generate_months_list(start_month: int, start_year: int, count: int = 24):
        month_names = [
            "Январь", "Февраль", "Март", "Апрель", "Май", "юнь",
            "юль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
        ]
        months = []
        current_month = start_month
        current_year = start_year
        for _ in range(count):
            months.append(
                {
                    "period": f"{month_names[current_month - 1]} {current_year}",
                    "month_num": current_month,
                    "year": current_year,
                }
            )
            current_month += 1
            if current_month > 12:
                current_month = 1
                current_year += 1
        return months

    def calculate_seasonality_coeffs_demand(df_data: pd.DataFrame, query_cols: list[str]) -> pd.DataFrame:
        """
        Спрос (по запросам), логика как в Excel-файле:
        1) Для каждого календарного месяца суммируем по каждому запросу (по всем годам).
        2) Внутри месяца считаем веса запросов как доли (сумма по запросу / сумма по всем запросам месяца).
        3) Сред. взв. месяца = сумма(сумма_по_запросу_месяца * вес_запроса_в_месяце).
        4) Среднее по 12 месяцам = СРЗНАЧ(Сред. взв. по месяцам).
        5) Коэф. месяца = Сред. взв. месяца / Среднее по 12 месяцам.
        """
        month_names_map = {
            1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель",
            5: "Май", 6: "юнь", 7: "юль", 8: "Август",
            9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь",
        }

        df = df_data.copy()

        avg_weighted_by_month: dict[int, float] = {}

        for m in range(1, 13):
            df_m = df[df["month_num"] == m]
            if df_m.empty:
                avg_weighted_by_month[m] = 0.0
                continue

            sums_per_query = df_m[query_cols].sum(axis=0)
            total_sum_month = float(sums_per_query.sum())
            if total_sum_month <= 0:
                avg_weighted_by_month[m] = 0.0
                continue

            weights = sums_per_query / total_sum_month
            avg_weighted = float((sums_per_query * weights).sum())
            avg_weighted_by_month[m] = avg_weighted

        values = list(avg_weighted_by_month.values())
        avg_all_months = float(np.mean(values)) if values else 1.0
        if avg_all_months == 0:
            avg_all_months = 1.0

        rows = []
        for m in range(1, 13):
            avg_w = avg_weighted_by_month.get(m, 0.0)
            coeff = avg_w / avg_all_months if avg_all_months > 0 else 1.0
            rows.append(
                {
                    "Номер месяца": m,
                    "Месяц": month_names_map[m],
                    "Сред. взв.": round(avg_w, 2),
                    "Коэф.": round(coeff, 2),
                }
            )

        return pd.DataFrame(rows)

    def calculate_seasonality_coeffs_aov(df_data: pd.DataFrame) -> pd.DataFrame:
        """
        AOV (средний чек) по месяцам:
        1) Для каждого календарного месяца считаем средний AOV по строкам (по годам).
        2) Считаем средний AOV по 12 месяцам.
        3) Коэф. AOV месяца = Сред. AOV месяца / средний AOV по 12 месяцам.
        """
        month_names_map = {
            1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель",
            5: "Май", 6: "юнь", 7: "юль", 8: "Август",
            9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь",
        }

        df = df_data.copy()

        # средний AOV по каждому календарному месяцу
        if "AOV" not in df.columns:
            df["AOV"] = 0.0
        avg_aov_by_month = df.groupby("month_num")["AOV"].mean()

        avg_all = float(avg_aov_by_month.mean()) if not avg_aov_by_month.empty else 1.0
        if avg_all == 0:
            avg_all = 1.0

        rows = []
        for m in range(1, 13):
            avg_m = float(avg_aov_by_month.get(m, 0.0))
            coeff = avg_m / avg_all if avg_all > 0 else 1.0
            rows.append(
                {
                    "Номер месяца": m,
                    "Месяц": month_names_map[m],
                    "Сред. AOV": round(avg_m, 2),
                    "Коэф. AOV": round(coeff, 2),
                }
            )

        return pd.DataFrame(rows)

    # ---------------- НЦАЛЗАЦЯ СОСТОЯНЯ ----------------

    if "coeff_sets" not in st.session_state:
        st.session_state["coeff_sets"] = []  # список наборов

    # ---------------- ДОБАВЛЕНЕ НОВОГО НАБОРА ----------------

    col_add_set, col_help = st.columns([2, 3])
    with col_add_set:
        if st.button("➕ Добавить новый набор коэффициентов", key="add_coeff_set"):
            new_id = len(st.session_state["coeff_sets"]) + 1
            st.session_state["coeff_sets"].append(
                {
                    "id": new_id,
                    "name": f"Набор {new_id}",
                    "type": "Спрос (по запросам)",  # или "AOV (средний чек)"
                    "start_month": 1,
                    "start_year": 2024,
                    "queries": [],
                    "df_data": None,
                    "result": None,
                }
            )
            st.rerun()
    with col_help:
        st.empty()

    # ---------------- СПСОК НАБОРОВ ----------------

    if not st.session_state["coeff_sets"]:
        st.info("Пока нет ни одного набора коэффициентов. Нажмите «➕ Добавить новый набор коэффициентов».")
    else:
        for idx, cs in enumerate(st.session_state["coeff_sets"]):
            set_id = cs["id"]
            with st.expander(f"Набор {set_id}: {cs['name']}", expanded=False):

                # ---- 1-2. Название и тип набора (компактный layout) ----
                meta_c1, meta_c2, _meta_spacer = st.columns([2.2, 2.2, 1.6], vertical_alignment="bottom")
                with meta_c1:
                    cs["name"] = st.text_input(
                        "Название набора",
                        value=cs["name"],
                        key=f"cs_name_{set_id}",
                    )
                with meta_c2:
                    cs["type"] = normalize_coeff_set_type(cs.get("type"))
                    cs["type"] = st.selectbox(
                        "Тип набора:",
                        options=["Спрос (по запросам)", "AOV (средний чек)", "Кастомный набор", "Медийные хвосты"],
                        help=(
                            "• Спрос (по запросам)\n"
                            "  Рассчитывает сезонность по данным Wordstat.\n\n"
                            "• AOV (средний чек)\n"
                            "  Рассчитывает сезонность среднего чека.\n\n"
                            "• Кастомный набор\n"
                            "  Ручной ввод коэффициентов на 12 месяцев.\n\n"
                            "• Медийные хвосты\n"
                            "  Дополнительный множитель к показам, накладывается "
                            "поверх сезонности спроса."
                        ),
                        index=(
                            0 if cs.get("type") == "Спрос (по запросам)"
                            else 1 if cs.get("type") == "AOV (средний чек)"
                            else 2 if cs.get("type") == "Кастомный набор"
                            else 3
                        ),
                        key=f"cs_type_{set_id}",
                    )

                # ---- 3. Период 24 месяца (компактный layout) ----
                col_m, col_y, _period_spacer = st.columns([1.8, 1.8, 2.4], vertical_alignment="bottom")
                with col_m:
                    cs["start_month"] = st.selectbox(
                        "Начальный месяц периода:",
                        options=list(range(1, 13)),
                        format_func=lambda x: [
                            "Январь", "Февраль", "Март", "Апрель", "Май", "юнь",
                            "юль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь",
                        ][x - 1],
                        index=cs["start_month"] - 1,
                        key=f"cs_month_{set_id}",
                    )
                with col_y:
                    cs["start_year"] = st.number_input(
                        "Начальный год периода:",
                        min_value=2020,
                        max_value=2030,
                        value=cs["start_year"],
                        step=1,
                        key=f"cs_year_{set_id}",
                    )

                # ---- 4. нтерфейс в зависимости от типа набора ----

                if cs["type"] == "Спрос (по запросам)":
                    # ===== РЕЖМ СПРОС (ПО ЗАПРОСАМ) =====
                    st.markdown("**Поисковые запросы:**")

                    if not cs.get("queries"):
                        cs["queries"] = ["Запрос 1"]

                    to_remove = []
                    for q_idx, q in enumerate(cs["queries"]):
                        query_row, _query_spacer = st.columns([1.6, 4.4])
                        with query_row:
                            col_q, col_del = st.columns([5, 1])
                            with col_q:
                                new_q = st.text_input(
                                    f"Запрос {q_idx + 1}:",
                                    value=q,
                                    key=f"cs_q_{set_id}_{q_idx}",
                                    label_visibility="collapsed",
                                )
                                cs["queries"][q_idx] = new_q
                            with col_del:
                                if st.button("🗑️", key=f"cs_del_{set_id}_{q_idx}"):
                                    to_remove.append(q_idx)

                    for r in sorted(to_remove, reverse=True):
                        cs["queries"].pop(r)
                        st.rerun()

                    if st.button("➕ Добавить запрос", key=f"cs_add_q_{set_id}"):
                        cs["queries"].append("")
                        st.rerun()

                    # Генерация таблицы на 24 месяца
                    if st.button("Применить настройки набора", key=f"cs_apply_{set_id}", type="primary"):
                        queries_clean = [q.strip() for q in cs["queries"] if q.strip()]
                        if not queries_clean:
                            st.error("Добавьте хотя бы один поисковый запрос.")
                        else:
                            months_24 = generate_months_list(cs["start_month"], cs["start_year"], 24)
                            df_cs = pd.DataFrame(months_24)
                            for q in queries_clean:
                                df_cs[q] = 0
                            cs["df_data"] = df_cs
                            st.success("Настроен период и запросы, заполните таблицу ниже.")
                            st.rerun()

                    # Редактор данных и расчёт
                    if cs.get("df_data") is not None:
                        st.markdown("**Заполните данные (можно копировать из Excel через Ctrl+C → Ctrl+V):**")
                        df_edit = st.data_editor(
                            cs["df_data"],
                            use_container_width=True,
                            row_height=ui_editor_row_height,
                            num_rows="fixed",
                            column_config={
                                "period": st.column_config.TextColumn("Период", disabled=True),
                                "month_num": st.column_config.NumberColumn("Месяц №", disabled=True),
                                "year": st.column_config.NumberColumn("Год", disabled=True),
                            },
                            key=f"cs_editor_{set_id}",
                        )
                        cs["df_data"] = df_edit

                        if st.button("Рассчитать коэффициенты для этого набора", key=f"cs_calc_{set_id}", type="primary"):
                            queries_clean = [q.strip() for q in cs["queries"] if q.strip()]
                            if not queries_clean:
                                st.error("Добавьте хотя бы один поисковый запрос.")
                            else:
                                df_coeffs = calculate_seasonality_coeffs_demand(df_edit, queries_clean)
                                cs["result"] = df_coeffs
                                st.success("Коэффициенты рассчитаны!")

                elif cs["type"] == "AOV (средний чек)":
                    # ===== РЕЖМ AOV (СРЕДНЙ ЧЕК) =====
                    st.markdown("**Данные по среднему чеку (AOV):**")
                    st.caption("Для каждого из 24 месяцев укажите средний чек (AOV).")

                    # Генерация таблицы на 24 месяца (только AOV)
                    if st.button("Применить настройки набора (AOV)", key=f"cs_apply_aov_{set_id}", type="primary"):
                        months_24 = generate_months_list(cs["start_month"], cs["start_year"], 24)
                        df_cs = pd.DataFrame(months_24)
                        df_cs["AOV"] = 0.0
                        cs["df_data"] = df_cs
                        st.success("Настроен период. Заполните AOV по месяцам ниже.")
                        st.rerun()

                    # Редактор данных и расчёт для AOV
                    if cs.get("df_data") is not None:
                        aov_left, aov_right = st.columns([1.65, 1.0], vertical_alignment="top")
                        with aov_left:
                            st.markdown("**Заполните AOV (средний чек) по месяцам:**")
                            df_edit = st.data_editor(
                                cs["df_data"],
                                use_container_width=True,
                                row_height=ui_editor_row_height,
                                num_rows="fixed",
                                column_config={
                                    "period": st.column_config.TextColumn("Период", disabled=True),
                                    "month_num": st.column_config.NumberColumn("Месяц №", disabled=True),
                                    "year": st.column_config.NumberColumn("Год", disabled=True),
                                    "AOV": st.column_config.NumberColumn(
                                        "AOV (средний чек)", format="%.2f", help=mhelp("aov")
                                    ),
                                },
                                key=f"cs_editor_aov_{set_id}",
                            )
                            cs["df_data"] = df_edit

                        if st.button("Рассчитать коэффициенты AOV для этого набора", key=f"cs_calc_aov_{set_id}", type="primary"):
                                df_coeffs = calculate_seasonality_coeffs_aov(df_edit)
                                cs["result"] = df_coeffs
                                st.success("Коэффициенты AOV рассчитаны!")
                        with aov_right:
                            st.markdown("**тоговые коэффициенты AOV**")
                            if cs.get("result") is not None:
                                st.dataframe(cs["result"], use_container_width=True, height=420)
                            else:
                                st.info("После расчета здесь появится итоговая таблица коэффициентов.")
                elif cs["type"] == "Кастомный набор":
                    # ===== РЕЖМ КАСТОМНЫЙ НАБОР =====
                    st.markdown("**Кастомные коэффициенты на 12 месяцев:**")
                    st.caption(
                        "Заполните коэффициенты вручную. Этот набор можно использовать и для спроса, и для AOV."
                    )

                    month_names_map = {
                        1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель",
                        5: "Май", 6: "юнь", 7: "юль", 8: "Август",
                        9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь",
                    }

                    if (
                        cs.get("df_data") is None
                        or "month_num" not in cs["df_data"].columns
                        or "Коэф." not in cs["df_data"].columns
                    ):
                        cs["df_data"] = pd.DataFrame(
                            {
                                "month_num": list(range(1, 13)),
                                "Месяц": [month_names_map[m] for m in range(1, 13)],
                                "Коэф.": [1.0] * 12,
                            }
                        )

                    custom_left, custom_right = st.columns([1.65, 1.0], vertical_alignment="top")
                    with custom_left:
                        df_edit = st.data_editor(
                            cs["df_data"],
                            use_container_width=True,
                            row_height=ui_editor_row_height,
                            num_rows="fixed",
                            column_config={
                                "month_num": st.column_config.NumberColumn("Месяц №", disabled=True),
                                "Месяц": st.column_config.TextColumn("Месяц", disabled=True),
                                "Коэф.": st.column_config.NumberColumn(
                                    "Коэффициент", format="%.2f",
                                    help="Коэффициент месяца: 1.00 = без изменений, 1.20 = +20%, 0.80 = -20%."
                                ),
                            },
                            key=f"cs_editor_custom_{set_id}",
                        )
                        cs["df_data"] = df_edit
                    with custom_right:
                        st.markdown("**Действия**")
                        if st.button("Применить кастомный набор", key=f"cs_apply_custom_{set_id}", type="primary"):
                            df_coeffs = df_edit.copy()
                            if "Коэф." not in df_coeffs.columns:
                                df_coeffs["Коэф."] = 1.0
                            df_coeffs["Номер месяца"] = df_coeffs["month_num"].astype(int)
                            df_coeffs["Коэф. AOV"] = df_coeffs["Коэф."].astype(float)
                            df_coeffs = df_coeffs[["Номер месяца", "Месяц", "Коэф.", "Коэф. AOV"]]
                            cs["result"] = df_coeffs
                            st.success("Кастомные коэффициенты сохранены!")
                        else:
                            st.caption("Вносите значения слева и применяйте набор кнопкой выше.")
                else:
                    # ===== РЕЖМ МЕДЙНЫЕ ХВОСТЫ =====
                    st.markdown("**Медийные хвосты (коэффициенты показов на 12 месяцев):**")
                    st.caption(
                        "Этот набор накладывается на спрос и влияет только на показы "
                        "(k_imp = k_demand × k_media_tail)."
                    )

                    month_names_map = {
                        1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель",
                        5: "Май", 6: "юнь", 7: "юль", 8: "Август",
                        9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь",
                    }

                    if (
                        cs.get("df_data") is None
                        or "month_num" not in cs["df_data"].columns
                        or "Коэф." not in cs["df_data"].columns
                    ):
                        cs["df_data"] = pd.DataFrame(
                            {
                                "month_num": list(range(1, 13)),
                                "Месяц": [month_names_map[m] for m in range(1, 13)],
                                "Коэф.": [1.0] * 12,
                            }
                        )

                    media_left, media_right = st.columns([1.65, 1.0], vertical_alignment="top")
                    with media_left:
                        df_edit = st.data_editor(
                            cs["df_data"],
                            use_container_width=True,
                            row_height=ui_editor_row_height,
                            num_rows="fixed",
                            column_config={
                                "month_num": st.column_config.NumberColumn("Месяц №", disabled=True),
                                "Месяц": st.column_config.TextColumn("Месяц", disabled=True),
                                "Коэф.": st.column_config.NumberColumn(
                                    "Коэффициент хвостов", format="%.2f",
                                    help="Медийный хвост для показов: k_imp = k_demand × k_media_tail."
                                ),
                            },
                            key=f"cs_editor_media_tail_{set_id}",
                        )
                        cs["df_data"] = df_edit
                    with media_right:
                        st.markdown("**Действия**")
                        if st.button("Применить набор медийных хвостов", key=f"cs_apply_media_tail_{set_id}", type="primary"):
                            df_coeffs = df_edit.copy()
                            if "Коэф." not in df_coeffs.columns:
                                df_coeffs["Коэф."] = 1.0
                            df_coeffs["Номер месяца"] = df_coeffs["month_num"].astype(int)
                            df_coeffs = df_coeffs[["Номер месяца", "Месяц", "Коэф."]]
                            cs["result"] = df_coeffs
                            st.success("Набор медийных хвостов сохранен!")
                        else:
                            st.caption("Набор влияет только на показы и накладывается поверх спроса.")

                # ---- 5. Показ результата + скачивание ----
                if cs.get("result") is not None:
                    cs_type_norm = normalize_coeff_set_type(cs.get("type"))
                    if cs_type_norm == "Спрос (по запросам)":
                        st.markdown("### тоговые коэффициенты")
                        st.dataframe(cs["result"], use_container_width=True)

                    buf = io.BytesIO()
                    sheet_name = (cs["name"] or "Коэффициенты")[:31]
                    excel_engine = "xlsxwriter" if HAS_XLSXWRITER else "openpyxl"
                    with pd.ExcelWriter(buf, engine=excel_engine) as writer:
                        cs["result"].to_excel(writer, sheet_name=sheet_name, index=False)
                    buf.seek(0)

                    st.download_button(
                        label="📥 Скачать коэффициенты этого набора",
                        data=buf,
                        file_name=f"coeffs_set_{set_id}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"cs_dl_{set_id}",
                    )
    render_bottom_tab_switcher("Коэффициенты", "coeffs")
# ====================================================================
#                           ТАБ «МЕДАПЛАН»
# ====================================================================

with tab_plan:
    st.markdown(
        """
        <div class="tab-intro">
            <p>1) Выберите месяцы, для которых нужно рассчитать медиаплан.</p>
            <p>2) Заполните данные по типам рекламных кампаний и базовым метрикам (Показы, CPC, CR, CTR, AOV) для расчета среднего месяца. Значения в блок «Средний месяц» вносятся <b><span style="color:#9EC5FF;">без НДС</span></b>. Помните: сезонность среднего месяца равна 1, и от нее рассчитываются все 12 месяцев через коэффициенты.</p>
            <p>3) Назначьте для каждого типа РК наборы коэффициентов: Спрос, AOV, Кастомный набор и при необходимости Медийные хвосты.</p>
            <p>4) Настройте эластичность метрик (CPC, CTR, CR), чтобы учесть реакцию показателей на сезонность спроса.</p>
            <p>5) Проверьте расчеты по каждому месяцу: слева результаты, справа примененные коэффициенты. Все коэффициенты справа редактируемые: вносите изменения прямо в ячейки и сразу смотрите на результат.</p>
            <p>6) Опционально используйте левое меню «Быстрая сверка», чтобы закрепить перед глазами средний месяц или TOTAL выбранных месяцев и удобнее сравнивать данные при редактировании.</p>
            <p>7) Сверьте итоги в блоке TOTAL по выбранным месяцам и при необходимости скорректируйте настройки.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if "use_vat_budget_metrics" not in st.session_state:
        st.session_state["use_vat_budget_metrics"] = True
    if "use_ak_budget_metrics" not in st.session_state:
        st.session_state["use_ak_budget_metrics"] = False
    if "ak_mode" not in st.session_state:
        st.session_state["ak_mode"] = "percent"
    if "ak_fixed_month_wo_vat" not in st.session_state:
        st.session_state["ak_fixed_month_wo_vat"] = 200000.0
    if "ak_rules_editor_nonce" not in st.session_state:
        st.session_state["ak_rules_editor_nonce"] = 0
    default_ak_rules = pd.DataFrame(
        [
            {"min_budget_wo_vat": 0.0, "ak_percent": 0.0},
            {"min_budget_wo_vat": 1500000.0, "ak_percent": 8.0},
            {"min_budget_wo_vat": 3000000.0, "ak_percent": 4.0},
            {"min_budget_wo_vat": 6000000.0, "ak_percent": 2.0},
            {"min_budget_wo_vat": 10000000.0, "ak_percent": 0.0},
        ]
    )
    if "ak_rules_df" not in st.session_state:
        st.session_state["ak_rules_df"] = default_ak_rules.copy()
    use_vat_budget_metrics = bool(st.session_state.get("use_vat_budget_metrics", True))
    use_ak_budget_metrics = bool(st.session_state.get("use_ak_budget_metrics", False))
    ak_mode = str(st.session_state.get("ak_mode", "fixed"))
    ak_fixed_month_wo_vat = float(st.session_state.get("ak_fixed_month_wo_vat", 200000.0))
    ak_rules_df = st.session_state.get("ak_rules_df", pd.DataFrame())
    ak_rules_invalid = (
        not isinstance(ak_rules_df, pd.DataFrame)
        or ak_rules_df.empty
        or "min_budget_wo_vat" not in ak_rules_df.columns
        or "ak_percent" not in ak_rules_df.columns
    )
    if not ak_rules_invalid:
        _min_col = pd.to_numeric(ak_rules_df["min_budget_wo_vat"], errors="coerce")
        _pct_col = pd.to_numeric(ak_rules_df["ak_percent"], errors="coerce")
        ak_rules_invalid = _min_col.isna().all() or _pct_col.isna().all()
    if ak_rules_invalid:
        st.session_state["ak_rules_df"] = default_ak_rules.copy()
        ak_rules_df = st.session_state["ak_rules_df"]
        st.session_state["ak_rules_editor_nonce"] = int(st.session_state.get("ak_rules_editor_nonce", 0)) + 1

    month_names_full = {
        1: "Январь",
        2: "Февраль",
        3: "Март",
        4: "Апрель",
        5: "Май",
        6: "юнь",
        7: "юль",
        8: "Август",
        9: "Сентябрь",
        10: "Октябрь",
        11: "Ноябрь",
        12: "Декабрь",
    }

    def _fmt_ref_int(v: float) -> str:
        return f"{int(round(v)):,}".replace(",", " ")

    def _fmt_ref_rub(v: float) -> str:
        return f"{int(round(v)):,} ₽".replace(",", " ")

    def _fmt_ref_rub2(v: float) -> str:
        return f"{float(v):.2f} ₽"

    def _fmt_ref_pct(v: float) -> str:
        return f"{float(v):.2f}%"

    def _fmt_ref_roas(v: float) -> str:
        return f"{float(v):.2f}%"

    def _build_ref_df(ref: dict) -> pd.DataFrame:
        if not ref:
            return pd.DataFrame(columns=["Показатель", "Значение"])
        rows = [
            {"Показатель": "Показы", "Значение": _fmt_ref_int(ref.get("impressions", 0.0))},
            {"Показатель": "Клики", "Значение": _fmt_ref_int(ref.get("clicks", 0.0))},
            {"Показатель": "Конверсии", "Значение": _fmt_ref_int(ref.get("conversions", 0.0))},
            {"Показатель": "Бюджет", "Значение": _fmt_ref_rub(ref.get("cost", 0.0))},
            {"Показатель": "Бюджет с НДС", "Значение": _fmt_ref_rub(ref.get("cost_with_vat", 0.0))},
            {"Показатель": "Бюджет с НДС и АК", "Значение": _fmt_ref_rub(ref.get("cost_with_vat_ak", 0.0))},
            {"Показатель": "Доход", "Значение": _fmt_ref_rub(ref.get("revenue", 0.0))},
            {"Показатель": "CTR", "Значение": _fmt_ref_pct(ref.get("ctr", 0.0))},
            {"Показатель": "CPC", "Значение": _fmt_ref_rub2(ref.get("cpc", 0.0))},
            {"Показатель": "CR", "Значение": _fmt_ref_pct(ref.get("cr", 0.0))},
            {
                "Показатель": "CPO",
                "Значение": _fmt_ref_rub(ref.get("cpo", 0.0)),
            },
            {
                "Показатель": "ROAS",
                "Значение": _fmt_ref_roas(ref.get("roas", 0.0)),
            },
            {
                "Показатель": "ДРР",
                "Значение": _fmt_ref_pct(ref.get("drr", 0.0)),
            },
        ]
        return pd.DataFrame(rows)

    # ---------- 0. Горизонт планирования ----------

    ui_section_title("0. Горизонт планирования")
    st.caption(
        "Выбранные месяцы напрямую влияют на помесячный расчет, итоги в блоке TOTAL "
        "и данные на вкладке «Диаграммы»."
    )

    all_month_nums = list(month_names_full.keys())
    all_month_labels = [f"{i}. {month_names_full[i]}" for i in all_month_nums]

    selected_month_labels = st.multiselect(
        "Выберите месяцы, для которых считать медиаплан:",
        options=all_month_labels,
        default=all_month_labels,
        key="planning_months_multiselect",
    )

    if not selected_month_labels:
        st.warning("Выберите хотя бы один месяц для расчёта медиаплана.")
        st.stop()

    selected_month_nums = [int(label.split(".")[0]) for label in selected_month_labels]
    months_count = len(selected_month_nums)

    ui_section_title("0.1 Пресет метрик")
    st.caption(
        "После выбора горизонта планирования выберите пресет метрик. "
        "Доступны пресеты: E-com и DIY."
    )
    preset_keys = list(METRIC_PRESETS.keys())
    if "metric_preset_key" not in st.session_state or st.session_state["metric_preset_key"] not in preset_keys:
        st.session_state["metric_preset_key"] = "ecom"
    preset_col, _preset_spacer = st.columns([1, 3])
    with preset_col:
        active_preset_key = st.selectbox(
            "Выберите пресет метрик:",
            options=preset_keys,
            index=preset_keys.index(st.session_state["metric_preset_key"]),
            format_func=lambda k: METRIC_PRESETS[k]["label"],
            key="metric_preset_key",
        )
    st.caption(f"Активный пресет: {METRIC_PRESETS[active_preset_key]['label']}")
    is_diy_preset = active_preset_key == "diy"
    if is_diy_preset:
        DISPLAY_COL_RENAME["conversions"] = "Продажи"
        DISPLAY_COL_RENAME["cr"] = "CR в продажи"
        DISPLAY_COL_RENAME["cr_pct"] = "CR в продажи"
    else:
        DISPLAY_COL_RENAME["conversions"] = "Конверсии"
        DISPLAY_COL_RENAME["cr"] = "CR"
        DISPLAY_COL_RENAME["cr_pct"] = "CR"

    def _df_for_compare(df: pd.DataFrame) -> pd.DataFrame:
        if df is None:
            return pd.DataFrame()
        tmp = df.copy().reset_index(drop=True)
        tmp = tmp.reindex(sorted(tmp.columns), axis=1)
        return tmp.fillna("")

    def _has_unsaved_changes(draft_df: pd.DataFrame, saved_df: pd.DataFrame) -> bool:
        return not _df_for_compare(draft_df).equals(_df_for_compare(saved_df))

    def _table_height_for_rows(rows_count: int, min_height: int = 120, max_height: int = 560) -> int:
        header_px = 34 if ui_compact_tables else 38
        row_px = ui_table_row_px
        h = header_px + max(rows_count, 1) * row_px + 6
        return max(min_height, min(h, max_height))

    # ---------- 1. Типы РК (средний месяц) ----------

    ui_section_title("1. Типы рекламных кампаний в среднем месяце")
    with st.expander("Типы рекламных кампаний в среднем месяце", expanded=True):

        default_campaigns = pd.DataFrame(
            [
                ["Поиск бренд", "B2C", "Яндекс", "Поиск",        500_000, 5.0, 15.0, 5.0, 5000.0, 0.0, 0.0],
                ["РСЯ баннеры", "B2C", "Яндекс", "РСЯ баннеры", 1_000_000, 1.0, 10.0, 2.0, 3000.0, 0.0, 0.0],
                ["Видео YouTube", "B2B", "YouTube", "Видео",     300_000, 0.7, 20.0, 1.5, 4000.0, 0.0, 0.0],
            ],
            columns=[
                "campaign_type", "segment", "system", "format",
                "impressions_avg", "ctr_avg_percent", "cpc_avg", "cr_avg_percent", "aov_avg",
                "available_capacity_avg", "new_clients_share_avg_percent",
            ],
        )

        if "campaigns_df" not in st.session_state:
            st.session_state["campaigns_df"] = default_campaigns.copy()
        if "segment" not in st.session_state["campaigns_df"].columns:
            st.session_state["campaigns_df"]["segment"] = "B2C"
        if "available_capacity_avg" not in st.session_state["campaigns_df"].columns:
            st.session_state["campaigns_df"]["available_capacity_avg"] = 0.0
        if "new_clients_share_avg_percent" not in st.session_state["campaigns_df"].columns:
            st.session_state["campaigns_df"]["new_clients_share_avg_percent"] = 0.0

        campaign_cols_order = [
            "campaign_type",
            "system",
            "format",
            "segment",
            "impressions_avg",
            "ctr_avg_percent",
            "cpc_avg",
            "cr_avg_percent",
            "aov_avg",
            "available_capacity_avg",
            "new_clients_share_avg_percent",
        ]
        st.session_state["campaigns_df"] = st.session_state["campaigns_df"].reindex(
            columns=[c for c in campaign_cols_order if c in st.session_state["campaigns_df"].columns]
        )

        campaigns_column_config = {
            "campaign_type": st.column_config.TextColumn("Название кампании"),
            "system": st.column_config.TextColumn("Рекламная система"),
            "format": st.column_config.TextColumn("Формат/таргетинги"),
            "impressions_avg": st.column_config.NumberColumn(
                "Показы (средний месяц)", format="%.0f", help=mhelp("impressions")
            ),
            "ctr_avg_percent": st.column_config.NumberColumn(
                "CTR, % (средний месяц)", format="%.2f", help=mhelp("ctr")
            ),
            "cpc_avg": st.column_config.NumberColumn(
                "CPC, ₽ (средний месяц)", format="%.2f", help=mhelp("cpc")
            ),
            "cr_avg_percent": st.column_config.NumberColumn(
                "CR в продажи, % (средний месяц)" if is_diy_preset else "CR, % (средний месяц)",
                format="%.2f",
                help=mhelp("cr"),
            ),
            "aov_avg": st.column_config.NumberColumn(
                "Средний чек, ₽ (средний месяц)" if is_diy_preset else "AOV, ₽ (средний месяц)",
                format="%.2f",
                help=mhelp("aov"),
            ),
        }
        campaigns_editor_columns = [
            "campaign_type",
            "system",
            "format",
            "impressions_avg",
            "ctr_avg_percent",
            "cpc_avg",
            "cr_avg_percent",
            "aov_avg",
        ]
        if is_diy_preset:
            campaigns_column_config["segment"] = st.column_config.SelectboxColumn("Сегмент", options=["B2B", "B2C"])
            campaigns_editor_columns.insert(3, "segment")
            campaigns_column_config["available_capacity_avg"] = st.column_config.NumberColumn(
                "Доступная емкость (база)", format="%.0f"
            )
            campaigns_column_config["new_clients_share_avg_percent"] = st.column_config.NumberColumn(
                "Доля новых клиентов, % (база)", format="%.2f"
            )
            campaigns_editor_columns += ["available_capacity_avg", "new_clients_share_avg_percent"]

        with st.form("campaigns_editor_form"):
            campaigns_draft = st.data_editor(
                st.session_state["campaigns_df"],
                num_rows="dynamic",
                use_container_width=True,
                row_height=ui_editor_row_height,
                column_order=campaigns_editor_columns,
                column_config=campaigns_column_config,
                key="campaigns_editor",
            )
            if _has_unsaved_changes(campaigns_draft, st.session_state["campaigns_df"]):
                st.caption("Есть несохраненные изменения в таблице РК.")
            camp_btn_c1, camp_btn_c2 = st.columns([1.2, 2.8])
            with camp_btn_c1:
                campaigns_saved = st.form_submit_button("Сохранить таблицу РК", type="primary")
            with camp_btn_c2:
                campaign_row_added = st.form_submit_button("➕ Добавить тип РК")

        if campaigns_saved:
            st.session_state["campaigns_df"] = campaigns_draft.copy()

        if campaign_row_added:
            new_row = {
                "campaign_type": "",
                "segment": "B2C",
                "system": "",
                "format": "",
                "impressions_avg": 0,
                "ctr_avg_percent": 0.0,
                "cpc_avg": 0.0,
                "cr_avg_percent": 0.0,
                "aov_avg": 0.0,
                "available_capacity_avg": 0.0,
                "new_clients_share_avg_percent": 0.0,
            }
            st.session_state["campaigns_df"] = pd.concat(
                [campaigns_draft, pd.DataFrame([new_row])],
                ignore_index=True,
            )
            st.rerun()

        campaigns = st.session_state["campaigns_df"].copy()
        campaigns = campaigns.dropna(subset=["campaign_type"])
        campaigns = campaigns[campaigns["campaign_type"].astype(str).str.strip() != ""]

        if campaigns.empty:
            st.warning("Добавьте хотя бы один тип РК.")
            st.stop()

        required_cols = ["impressions_avg", "ctr_avg_percent", "cpc_avg", "cr_avg_percent", "aov_avg"]

        def row_has_missing(row):
            for col in required_cols:
                val = row.get(col)
                if val is None:
                    return True
                if isinstance(val, str) and val.strip() == "":
                    return True
                try:
                    if pd.isna(val):
                        return True
                except Exception:
                    pass
            return False

        missing_mask = campaigns.apply(row_has_missing, axis=1)

        if missing_mask.any():
            st.error(
                "Заполни пожалуйста все поля в среднем месяце (Показы, CTR, CPC, CR, AOV) "
                "для каждого типа РК. Проверь строки, где есть пустые ячейки."
            )
            st.stop()

        segment_filter_options = ["Все", "B2B", "B2C"]
        if "plan_segment_filter" not in st.session_state:
            st.session_state["plan_segment_filter"] = "Все"
        if "plan_segment_filter_top" not in st.session_state:
            st.session_state["plan_segment_filter_top"] = st.session_state["plan_segment_filter"]
        if "plan_segment_filter_sidebar" not in st.session_state:
            st.session_state["plan_segment_filter_sidebar"] = st.session_state["plan_segment_filter"]

        def _sync_segment_from_top():
            val = str(st.session_state.get("plan_segment_filter_top", "Все"))
            if val not in segment_filter_options:
                val = "Все"
            st.session_state["plan_segment_filter"] = val
            st.session_state["plan_segment_filter_sidebar"] = val

        def _sync_segment_from_sidebar():
            val = str(st.session_state.get("plan_segment_filter_sidebar", "Все"))
            if val not in segment_filter_options:
                val = "Все"
            st.session_state["plan_segment_filter"] = val
            st.session_state["plan_segment_filter_top"] = val

        if is_diy_preset:
            st.radio(
                "Показывать кампании сегмента:",
                options=segment_filter_options,
                horizontal=True,
                key="plan_segment_filter_top",
                on_change=_sync_segment_from_top,
            )
            segment_filter = str(st.session_state.get("plan_segment_filter", "Все"))
            if segment_filter != "Все":
                campaigns = campaigns[campaigns["segment"].astype(str).str.upper() == segment_filter]
                if campaigns.empty:
                    st.warning(f"Для сегмента {segment_filter} нет кампаний.")
                    st.stop()
            visible_segments = sorted(campaigns["segment"].astype(str).str.upper().unique().tolist())
            show_segment_subtotals = len(visible_segments) > 1
        else:
            st.session_state["plan_segment_filter"] = "Все"
            st.session_state["plan_segment_filter_top"] = "Все"
            st.session_state["plan_segment_filter_sidebar"] = "Все"
            show_segment_subtotals = False

        existing_ctypes = campaigns["campaign_type"].astype(str).unique().tolist()

    with st.expander("1.2 НДС и АК", expanded=True):
        st.markdown(
            """
            <div style="
                margin: 10px 0 8px 0;
                display: inline-block;
                width: fit-content;
                max-width: 100%;
                padding: 10px 12px;
                border-radius: 10px;
                border: 1px solid #FF8A66;
                background: rgba(255, 99, 51, 0.14);
                color: #FFD9CC;
                font-weight: 600;
            ">
                <b>Учитываеть НДС 22% в расчетах</b>: настройка влияет на CPM, CPO, ROAS и ДРР.
            </div>
            """,
            unsafe_allow_html=True,
        )
        use_vat_budget_metrics = st.checkbox(
            "Учитываеть НДС 22% в расчетах",
            key="use_vat_budget_metrics",
        )
        use_ak_budget_metrics = st.checkbox(
            "Учитывать АК в расчетах",
            key="use_ak_budget_metrics",
        )
        ak_mode_label = st.selectbox(
            "Способ расчета АК",
            options=["Процент по шкале", "Фиксированная сумма в месяц"],
            index=0 if ak_mode == "percent" else 1,
            key="ak_mode_select",
        )
        ak_mode = "percent" if ak_mode_label == "Процент по шкале" else "fixed"
        st.session_state["ak_mode"] = ak_mode
        if ak_mode == "fixed":
            ak_fixed_month_wo_vat = st.number_input(
                "АК в месяц (без НДС), ₽",
                min_value=0.0,
                step=10000.0,
                format="%.0f",
                key="ak_fixed_month_wo_vat",
            )
            st.caption("Фиксированная АК месяца распределяется между типами РК пропорционально их бюджету без НДС.")
        if ak_mode == "percent":
            with st.expander("Настройка АК по шкале (TOTAL бюджета месяца без НДС)", expanded=use_ak_budget_metrics):
                st.caption(
                    "Для режима «Процент по шкале»: АК выбирается по наибольшему подходящему порогу TOTAL бюджета месяца без НДС."
                )
                ak_editor_key = f"ak_rules_editor_{int(st.session_state.get('ak_rules_editor_nonce', 0))}"
                scale_col, _spacer_col = st.columns([1, 2], vertical_alignment="top")
                with scale_col:
                    ak_rules_draft = st.data_editor(
                        st.session_state["ak_rules_df"],
                        num_rows="dynamic",
                        use_container_width=True,
                        row_height=ui_editor_row_height,
                        column_config={
                            "min_budget_wo_vat": st.column_config.NumberColumn(
                                "Бюджет от (без НДС), ₽",
                                min_value=0.0,
                                step=10000.0,
                                format="%.0f ₽",
                            ),
                            "ak_percent": st.column_config.NumberColumn(
                                "АК, %",
                                min_value=0.0,
                                step=0.5,
                                format="%.2f%%",
                            ),
                        },
                        key=ak_editor_key,
                    )
                if isinstance(ak_rules_draft, pd.DataFrame):
                    ak_rules_clean = ak_rules_draft.copy()
                    for c in ["min_budget_wo_vat", "ak_percent"]:
                        if c not in ak_rules_clean.columns:
                            ak_rules_clean[c] = 0.0
                    ak_rules_clean["min_budget_wo_vat"] = ak_rules_clean["min_budget_wo_vat"].map(
                        lambda v: parse_float_loose(v, 0.0)
                    )
                    ak_rules_clean["ak_percent"] = ak_rules_clean["ak_percent"].map(
                        lambda v: parse_float_loose(v, 0.0)
                    )
                    ak_rules_clean = ak_rules_clean[
                        ["min_budget_wo_vat", "ak_percent"]
                    ].sort_values("min_budget_wo_vat", kind="stable").reset_index(drop=True)
                    if ak_rules_clean.empty:
                        ak_rules_clean = pd.DataFrame([{"min_budget_wo_vat": 0.0, "ak_percent": 0.0}])
                    st.session_state["ak_rules_df"] = ak_rules_clean
                    ak_rules_df = ak_rules_clean
        current_vat_ak_sig = (
            bool(use_vat_budget_metrics),
            bool(use_ak_budget_metrics),
            str(ak_mode),
            round(float(ak_fixed_month_wo_vat), 4) if ak_mode == "fixed" else None,
            make_ak_rules_signature(st.session_state.get("ak_rules_df")) if ak_mode == "percent" else tuple(),
        )
        last_applied_vat_ak_sig = st.session_state.get("vat_ak_last_applied_sig")
        vat_ak_dirty = current_vat_ak_sig != last_applied_vat_ak_sig
        apply_vat_ak_clicked = st.button(
            "Применить настройки НДС и АК",
            type="primary" if vat_ak_dirty else "secondary",
            key="apply_vat_ak_settings_btn",
        )
        if apply_vat_ak_clicked:
            st.session_state["vat_ak_last_applied_sig"] = current_vat_ak_sig
            if vat_ak_dirty:
                st.success("Настройки НДС и АК применены.")
            else:
                st.info("зменений нет: текущие настройки уже применены.")
        st.caption(
            "CPC всегда считается от бюджета без НДС. При выключенном режиме НДС расчеты возвращаются к старой логике."
        )

    # ---------- 1.3. Наборы коэффициентов для типов РК ----------

    with st.expander("1.3. Наборы коэффициентов для типов РК", expanded=True):
        st.markdown(
            """
            <div style="
                margin: 8px 0 14px 0;
                display: inline-block;
                width: fit-content;
                max-width: 100%;
                white-space: nowrap;
                padding: 10px 12px;
                border-radius: 10px;
                border: 1px solid #FF8A66;
                background: rgba(255, 99, 51, 0.14);
                color: #FFD9CC;
                font-weight: 600;
            ">
                Важно: без назначенных наборов расчет не пойдет.
            </div>
            """,
            unsafe_allow_html=True,
        )

        coeff_sets = st.session_state.get("coeff_sets", [])

        demand_set_names = [
            cs["name"] for cs in coeff_sets
            if normalize_coeff_set_type(cs.get("type")) in ["Спрос (по запросам)", "Кастомный набор"]
        ]
        aov_set_names = [
            cs["name"] for cs in coeff_sets
            if normalize_coeff_set_type(cs.get("type")) in ["AOV (средний чек)", "Кастомный набор"]
        ]
        media_tail_set_names = [
            cs["name"] for cs in coeff_sets
            if normalize_coeff_set_type(cs.get("type")) == "Медийные хвосты"
        ]
        capacity_set_names = [cs["name"] for cs in coeff_sets]

        if "coeff_sets_links_new" not in st.session_state:
            st.session_state["coeff_sets_links_new"] = pd.DataFrame(
                columns=["campaign_type", "demand_set", "aov_set", "media_tail_set", "capacity_set"]
            )

        coeff_links_prev = st.session_state["coeff_sets_links_new"].copy()
        prev_demand_map = {}
        prev_aov_map = {}
        prev_media_tail_map = {}
        prev_capacity_map = {}
        for _, r in coeff_links_prev.iterrows():
            ct = str(r.get("campaign_type", "")).strip()
            if not ct:
                continue
            if ct not in prev_demand_map:
                prev_demand_map[ct] = str(r.get("demand_set", "")).strip()
            if ct not in prev_aov_map:
                prev_aov_map[ct] = str(r.get("aov_set", "")).strip()
            if ct not in prev_media_tail_map:
                prev_media_tail_map[ct] = str(r.get("media_tail_set", "")).strip()
            if ct not in prev_capacity_map:
                prev_capacity_map[ct] = str(r.get("capacity_set", "")).strip()

        coeff_links_new = pd.DataFrame(
            {
                "campaign_type": existing_ctypes,
                "demand_set": [prev_demand_map.get(ct, "") for ct in existing_ctypes],
                "aov_set": [prev_aov_map.get(ct, "") for ct in existing_ctypes],
                "media_tail_set": [prev_media_tail_map.get(ct, "") for ct in existing_ctypes],
                "capacity_set": [prev_capacity_map.get(ct, prev_demand_map.get(ct, "")) for ct in existing_ctypes],
            }
        )
        if not is_diy_preset:
            coeff_links_new = coeff_links_new.drop(columns=["capacity_set"], errors="ignore")

        st.markdown(
            "Для каждого **типа РК** выберите набор сезонности **спроса** и набор сезонности **AOV (средний чек)**."
        )

        with st.form("coeff_sets_links_form"):
            bulk_c1, bulk_c2, bulk_c3 = st.columns([2, 2, 1.2], vertical_alignment="bottom")
            with bulk_c1:
                bulk_demand_set = st.selectbox(
                    "Набор спроса для всех РК",
                    options=[""] + demand_set_names,
                    help="Быстро назначает один набор сезонности спроса всем типам РК в таблице ниже.",
                    key="bulk_demand_set_for_all",
                )
            with bulk_c2:
                bulk_aov_set = st.selectbox(
                    "Набор AOV для всех РК",
                    options=[""] + aov_set_names,
                    help="Быстро назначает один набор сезонности AOV всем типам РК в таблице ниже.",
                    key="bulk_aov_set_for_all",
                )
            with bulk_c3:
                apply_all_clicked = st.form_submit_button("Применить ко всем", type="primary")

            links_column_config = {
                "campaign_type": st.column_config.TextColumn(
                    "Тип РК",
                    disabled=True,
                ),
                "demand_set": st.column_config.SelectboxColumn(
                    "Набор сезонности спроса",
                    options=demand_set_names,
                    help="Влияет на показы (k_imp), а также через эластичность на CTR/CPC/CR.",
                ),
                "aov_set": st.column_config.SelectboxColumn(
                    "Набор сезонности AOV",
                    options=aov_set_names,
                    help="Влияет на средний чек: k_aov.",
                ),
                "media_tail_set": st.column_config.SelectboxColumn(
                    "Набор медийных хвостов (показы)",
                    options=[""] + media_tail_set_names,
                    help="Множитель показов поверх спроса: k_imp = k_demand * k_media_tail.",
                ),
            }
            if is_diy_preset:
                links_column_config["capacity_set"] = st.column_config.SelectboxColumn(
                    "Набор для емкости",
                    options=capacity_set_names,
                    help="Коэффициенты этого набора применяются к доступной емкости (k_capacity).",
                )

            coeff_links_draft = st.data_editor(
                coeff_links_new,
                num_rows="fixed",
                use_container_width=True,
                row_height=ui_editor_row_height,
                column_config=links_column_config,
                key="coeff_sets_links_new_editor",
            )
            if _has_unsaved_changes(coeff_links_draft, coeff_links_new):
                st.caption("Есть несохраненные изменения в выборе наборов.")
            links_saved = st.form_submit_button("Сохранить выбор наборов", type="primary")

        if apply_all_clicked or links_saved:
            coeff_links_new = coeff_links_draft.copy()
            if apply_all_clicked:
                if bulk_demand_set:
                    coeff_links_new["demand_set"] = bulk_demand_set
                if bulk_aov_set:
                    coeff_links_new["aov_set"] = bulk_aov_set
            st.session_state["coeff_sets_links_new"] = coeff_links_new
            if apply_all_clicked:
                # После массового применения перерисовываем форму,
                # чтобы data_editor отобразил новые значения в строках.
                st.rerun()

        aov_link_map = {}
        demand_link_map = {}
        media_tail_link_map = {}
        capacity_link_map = {}

        for _, r in coeff_links_new.iterrows():
            camp = str(r.get("campaign_type", "")).strip()
            if not camp:
                continue
            demand_link_map[camp] = str(r.get("demand_set", "")).strip()
            aov_link_map[camp] = str(r.get("aov_set", "")).strip()
            media_tail_link_map[camp] = str(r.get("media_tail_set", "")).strip()
            capacity_link_map[camp] = str(r.get("capacity_set", "")).strip()

        missing_demand = [ct for ct in existing_ctypes if not demand_link_map.get(ct)]
        if missing_demand:
            st.error(
                "Для следующих типов РК не выбран набор сезонности спроса, поэтому медиаплан дальше считаться не будет: "
                + ", ".join(missing_demand)
            )
            st.stop()

        missing_aov = [ct for ct in existing_ctypes if not aov_link_map.get(ct)]
        if missing_aov:
            st.error(
                "Для следующих типов РК не выбран набор сезонности AOV (средний чек), "
                "поэтому медиаплан дальше считаться не будет: "
                + ", ".join(missing_aov)
            )
            st.stop()

    # ---------- 1.4. Эластичность метрик к сезонности спроса ----------

    with st.expander("1.4. Эластичность метрик к сезонности спроса", expanded=False):

        st.caption("Памятка: меньше делитель = сильнее влияние.")
        st.caption(
            "Задайте делители влияния сезонности спроса на метрики для каждого типа РК. "
            "Чем меньше делитель, тем сильнее влияние. Процентный результат смотрите в превью справа."
        )
        st.markdown(
            "**Как заполнять поля делителей:**\n"
            "1. `CPC` — делитель роста CPC: формула `(k-1)/div + 1`.\n"
            "2. `CTR` — делитель снижения CTR: формула `1 - (k-1)/div`.\n"
            "3. `CR` — делитель снижения CR: формула `1 - (k-1)/div`.\n"
            "Пример: при `k=1.10` и `CPC div=2` получим `+5%` к CPC."
        )

        if "elasticity_df" not in st.session_state:
            st.session_state["elasticity_df"] = pd.DataFrame(
                {
                    "campaign_type": existing_ctypes,
                    "cpc_div": [1.0] * len(existing_ctypes),
                    "ctr_div": [2.0] * len(existing_ctypes),
                    "cr_div": [10.0] * len(existing_ctypes),
                }
            )

        elasticity_df = st.session_state["elasticity_df"]

        # Миграция старого формата процентов/legacy в делители.
        if "cpc_div" not in elasticity_df.columns and "cpc_impact_pct_per_10" in elasticity_df.columns:
            src = pd.to_numeric(elasticity_df["cpc_impact_pct_per_10"], errors="coerce")
            elasticity_df["cpc_div"] = np.where(src > 0, 10.0 / src, 1.0)
        if "ctr_div" not in elasticity_df.columns and "ctr_impact_pct_per_10" in elasticity_df.columns:
            src = pd.to_numeric(elasticity_df["ctr_impact_pct_per_10"], errors="coerce").abs()
            elasticity_df["ctr_div"] = np.where(src > 0, 10.0 / src, 2.0)
        if "cr_div" not in elasticity_df.columns and "cr_impact_pct_per_10" in elasticity_df.columns:
            src = pd.to_numeric(elasticity_df["cr_impact_pct_per_10"], errors="coerce").abs()
            elasticity_df["cr_div"] = np.where(src > 0, 10.0 / src, 10.0)
        if "cpc_div" not in elasticity_df.columns and "cpc_up_pct_per_10" in elasticity_df.columns:
            src = pd.to_numeric(elasticity_df["cpc_up_pct_per_10"], errors="coerce")
            elasticity_df["cpc_div"] = np.where(src > 0, 10.0 / src, 1.0)
        if "ctr_div" not in elasticity_df.columns and "ctr_down_pct_per_10" in elasticity_df.columns:
            src = pd.to_numeric(elasticity_df["ctr_down_pct_per_10"], errors="coerce").abs()
            elasticity_df["ctr_div"] = np.where(src > 0, 10.0 / src, 2.0)
        if "cr_div" not in elasticity_df.columns and "cr_down_pct_per_10" in elasticity_df.columns:
            src = pd.to_numeric(elasticity_df["cr_down_pct_per_10"], errors="coerce").abs()
            elasticity_df["cr_div"] = np.where(src > 0, 10.0 / src, 10.0)
        for col_name, default_val in [("cpc_div", 1.0), ("ctr_div", 2.0), ("cr_div", 10.0)]:
            if col_name in elasticity_df.columns:
                elasticity_df[col_name] = pd.to_numeric(elasticity_df[col_name], errors="coerce").fillna(default_val)

        elasticity_df = elasticity_df[
            elasticity_df["campaign_type"].isin(existing_ctypes)
        ]

        keep_cols = ["campaign_type", "preset", "cpc_div", "ctr_div", "cr_div"]
        for c in keep_cols:
            if c not in elasticity_df.columns:
                if c == "campaign_type":
                    elasticity_df[c] = ""
                elif c == "preset":
                    elasticity_df[c] = "Кастом"
                elif c == "cpc_div":
                    elasticity_df[c] = 1.0
                elif c == "ctr_div":
                    elasticity_df[c] = 2.0
                elif c == "cr_div":
                    elasticity_df[c] = 10.0
                else:
                    elasticity_df[c] = 0.0
        elasticity_df = elasticity_df[keep_cols]

        for ct in existing_ctypes:
            if ct not in elasticity_df["campaign_type"].tolist():
                elasticity_df = pd.concat(
                    [
                        elasticity_df,
                        pd.DataFrame(
                            [{
                                "campaign_type": ct,
                                "preset": "Среднее",
                                "cpc_div": 1.0,
                                "ctr_div": 2.0,
                                "cr_div": 10.0,
                            }]
                        ),
                    ],
                    ignore_index=True,
                )

        def _style_impact_col(col: pd.Series) -> pd.Series:
            out = pd.Series([""] * len(col), index=col.index)
            vals = pd.to_numeric(col, errors="coerce")
            for idx_v, v in vals.items():
                if pd.isna(v):
                    continue
                if v > 0:
                    out.loc[idx_v] = "color: #00CDC5; font-weight: 700;"
                elif v < 0:
                    out.loc[idx_v] = "color: #FF6333; font-weight: 700;"
                else:
                    out.loc[idx_v] = "color: #D0D6DF;"
            return out

        with st.form("elasticity_editor_form"):
            preset_map = {
                "Слабое": {"cpc_div": 2.0, "ctr_div": 5.0, "cr_div": 15.0},
                "Среднее": {"cpc_div": 1.0, "ctr_div": 2.0, "cr_div": 10.0},
                "Сильное": {"cpc_div": 0.5, "ctr_div": 1.0, "cr_div": 5.0},
            }
            st.caption("Для каждого типа РК можно выбрать пресет в колонке «Пресет» или задать значения вручную (Кастом).")
            st.info("Чтобы выбрать пресет: кликните ячейку в колонке «Пресет ▼» и выберите вариант из списка.")
            st.caption("В колонках CPC/CTR/CR указываются делители влияния. Процентный эффект для выбранного коэффициента спроса показывается в превью справа.")
            st.markdown(
                "**Описание пресетов:**\n"
                "1. `Слабое` — мягкая реакция на сезонность (CPC растет слабо, CTR/CR падают слабо).\n"
                "2. `Среднее` — сбалансированная реакция (базовый рекомендуемый вариант).\n"
                "3. `Сильное` — агрессивная реакция (резче рост CPC и падение CTR/CR).\n"
            )
            st.caption(
                "Пример влияния при росте спроса на +10%:\n"
                "Слабое: CPC +5.0%, CTR -2.0%, CR -0.67% | "
                "Среднее: CPC +10.0%, CTR -5.0%, CR -1.0% | "
                "Сильное: CPC +20.0%, CTR -10.0%, CR -2.0%."
            )
            st.caption(
                "Пример при +30% спроса (в 3 раза сильнее): "
                "для пресета «Среднее» ориентир: CPC +30.0%, CTR -15.0%, CR -3.0%."
            )
            st.caption(
                "Пример при снижении спроса до k=0.85 (−15%):\n"
                "Слабое: CPC -7.5%, CTR +3.0%, CR +1.0% | "
                "Среднее: CPC -15.0%, CTR +7.5%, CR +1.5% | "
                "Сильное: CPC -30.0%, CTR +15.0%, CR +3.0%."
            )
            top_left, top_right = st.columns([1.55, 1.05], vertical_alignment="bottom")
            with top_left:
                bulk_c1, bulk_c2, _bulk_sp = st.columns([1.2, 0.9, 1.2], vertical_alignment="bottom")
                with bulk_c1:
                    bulk_preset_name = st.selectbox(
                        "Массовый пресет",
                        options=[""] + list(preset_map.keys()),
                        key="elasticity_bulk_preset_select",
                    )
                with bulk_c2:
                    bulk_apply_clicked = st.form_submit_button("Применить массово", type="primary")
            with top_right:
                preview_k_col, preview_btn_col, _preview_k_sp = st.columns([0.34, 0.20, 0.46], vertical_alignment="bottom")
                with preview_k_col:
                    preview_k_demand = st.number_input(
                        "Коэффициент спроса для расчета превью",
                        min_value=0.10,
                        max_value=5.00,
                        value=float(st.session_state.get("elasticity_preview_k_demand", 1.10)),
                        step=0.01,
                        format="%.2f",
                        key="elasticity_preview_k_demand",
                        help="1.10 = +10% спроса, 1.30 = +30%, 0.90 = -10%.",
                    )
                with preview_btn_col:
                    preview_recalc_clicked = st.form_submit_button("Обновить")

            elasticity_row_h = max(28, int(ui_editor_row_height) - 8)
            compact_h = _table_height_for_rows(len(elasticity_df), min_height=140, max_height=280)

            edit_left, preview_right = st.columns([1.55, 1.05], vertical_alignment="top")
            with edit_left:
                st.caption("Настройка по типам РК")
                elasticity_draft = st.data_editor(
                    elasticity_df,
                    num_rows="fixed",
                    use_container_width=True,
                    row_height=elasticity_row_h,
                    height=compact_h,
                    column_config={
                        "campaign_type": st.column_config.TextColumn("Тип РК", disabled=True),
                        "preset": st.column_config.SelectboxColumn(
                            "Пресет ▼ (кликните ячейку)",
                            options=["Кастом", "Слабое", "Среднее", "Сильное"],
                            help="При сохранении пресет подставит значения влияния в этой строке.",
                        ),
                        "cpc_div": st.column_config.NumberColumn(
                            "CPC",
                            format="%.2f",
                            help=mhelp("cpc_div"),
                        ),
                        "ctr_div": st.column_config.NumberColumn(
                            "CTR",
                            format="%.2f",
                            help=mhelp("ctr_div"),
                        ),
                        "cr_div": st.column_config.NumberColumn(
                            "CR",
                            format="%.2f",
                            help=mhelp("cr_div"),
                        ),
                    },
                    key="elasticity_editor",
                )
            with preview_right:
                st.caption("Результат (с учетом пресета в строке)")
                preview_effective = elasticity_draft[
                    ["campaign_type", "preset", "cpc_div", "ctr_div", "cr_div"]
                ].copy()
                for idx_row, row in preview_effective.iterrows():
                    p_name = str(row.get("preset", "Кастом")).strip()
                    if p_name in preset_map:
                        p_vals = preset_map[p_name]
                        preview_effective.at[idx_row, "cpc_div"] = p_vals["cpc_div"]
                        preview_effective.at[idx_row, "ctr_div"] = p_vals["ctr_div"]
                        preview_effective.at[idx_row, "cr_div"] = p_vals["cr_div"]
                # Пересчет под выбранный коэффициент спроса:
                # значения в таблице задаются как эффект на +10% спроса, масштабируем линейно.
                # scale = 1 при k=1.10; scale = 3 при k=1.30; scale = -1 при k=0.90.
                demand_delta = float(preview_k_demand) - 1.0
                cpc_div_vals = pd.to_numeric(preview_effective["cpc_div"], errors="coerce").replace(0, np.nan)
                ctr_div_vals = pd.to_numeric(preview_effective["ctr_div"], errors="coerce").replace(0, np.nan)
                cr_div_vals = pd.to_numeric(preview_effective["cr_div"], errors="coerce").replace(0, np.nan)
                preview_effective["cpc_pct"] = (demand_delta / cpc_div_vals) * 100.0
                preview_effective["ctr_pct"] = -(demand_delta / ctr_div_vals) * 100.0
                preview_effective["cr_pct"] = -(demand_delta / cr_div_vals) * 100.0
                preview_effective = preview_effective[
                    ["campaign_type", "cpc_pct", "ctr_pct", "cr_pct"]
                ].rename(
                    columns={
                        "campaign_type": "Тип РК",
                        "cpc_pct": "CPC, %",
                        "ctr_pct": "CTR, %",
                        "cr_pct": "CR, %",
                    }
                )
                st.dataframe(
                    preview_effective.style
                    .format({"CPC, %": "{:+.2f}%", "CTR, %": "{:+.2f}%", "CR, %": "{:+.2f}%"})
                    .apply(_style_impact_col, axis=0, subset=["CPC, %", "CTR, %", "CR, %"]),
                    use_container_width=True,
                    hide_index=True,
                    height=compact_h,
                )
                st.caption("Зеленый = рост метрики, красный = снижение. Значения пересчитаны под выбранный коэффициент спроса.")
            if _has_unsaved_changes(elasticity_draft, elasticity_df):
                st.caption("Есть несохраненные изменения эластичности.")
            btn_save_col, btn_reset_col, _btn_sp = st.columns([0.28, 0.34, 0.38])
            with btn_save_col:
                elasticity_saved = st.form_submit_button("Сохранить эластичность", type="primary")
            with btn_reset_col:
                reset_recommended_clicked = st.form_submit_button("Вернуть рекомендуемые значения")

        if bulk_apply_clicked:
            if bulk_preset_name in preset_map:
                elasticity_draft = elasticity_draft.copy()
                elasticity_draft["preset"] = bulk_preset_name
                preset_vals = preset_map[bulk_preset_name]
                elasticity_draft["cpc_div"] = preset_vals["cpc_div"]
                elasticity_draft["ctr_div"] = preset_vals["ctr_div"]
                elasticity_draft["cr_div"] = preset_vals["cr_div"]
                st.session_state["elasticity_df"] = elasticity_draft.copy()
                elasticity_df = elasticity_draft.copy()
                st.success(f"Массово применен пресет «{bulk_preset_name}».")
                st.rerun()
            else:
                st.info("Выберите пресет для массового применения.")

        if elasticity_saved:
            elasticity_draft = elasticity_draft.copy()
            for idx_row, row in elasticity_draft.iterrows():
                preset_name = str(row.get("preset", "Кастом")).strip()
                if preset_name in preset_map:
                    preset_vals = preset_map[preset_name]
                    elasticity_draft.at[idx_row, "cpc_div"] = preset_vals["cpc_div"]
                    elasticity_draft.at[idx_row, "ctr_div"] = preset_vals["ctr_div"]
                    elasticity_draft.at[idx_row, "cr_div"] = preset_vals["cr_div"]
            st.session_state["elasticity_df"] = elasticity_draft.copy()
            elasticity_df = elasticity_draft.copy()

        if reset_recommended_clicked:
            st.session_state["elasticity_df"] = pd.DataFrame(
                {
                    "campaign_type": existing_ctypes,
                    "preset": ["Среднее"] * len(existing_ctypes),
                    "cpc_div": [1.0] * len(existing_ctypes),
                    "ctr_div": [2.0] * len(existing_ctypes),
                    "cr_div": [10.0] * len(existing_ctypes),
                }
            )
            st.rerun()

        def _safe_div(val: float, default_val: float) -> float:
            try:
                num = float(val)
                if pd.isna(num) or num <= 0:
                    return default_val
                return num
            except Exception:
                return default_val

        def _is_invalid_div(val: float) -> bool:
            try:
                num = float(val)
                return pd.isna(num) or num <= 0
            except Exception:
                return True

        had_invalid_values = False
        elasticity_map = {}
        for _, r in elasticity_df.iterrows():
            camp = str(r["campaign_type"])
            raw_cpc_div = r.get("cpc_div")
            raw_ctr_div = r.get("ctr_div")
            raw_cr_div = r.get("cr_div")

            cpc_div = _safe_div(raw_cpc_div, 1.0)
            ctr_div = _safe_div(raw_ctr_div, 2.0)
            cr_div = _safe_div(raw_cr_div, 10.0)

            if (
                _is_invalid_div(raw_cpc_div)
                or _is_invalid_div(raw_ctr_div)
                or _is_invalid_div(raw_cr_div)
            ):
                had_invalid_values = True

            elasticity_map[camp] = {
                "cpc_div": cpc_div,
                "ctr_div": ctr_div,
                "cr_div": cr_div,
            }

        if had_invalid_values:
            st.warning(
                "Некорректные значения делителей эластичности (пустые, нечисловые или <= 0) автоматически заменены: "
                "CPC=1.0, CTR=2.0, CR=10.0."
            )

    # ---------- 2. Средний месяц (базовые значения) ----------

    ui_section_title("2. Средний месяц (базовые значения)")
    # Берем актуальные правила из session_state перед расчетами,
    # чтобы исключить рассинхрон с data_editor.
    ak_rules_df = st.session_state.get("ak_rules_df", ak_rules_df)

    base_rows = []
    for _, base_row in campaigns.iterrows():
        base = PlanInput(
            impressions=base_row["impressions_avg"],
            ctr=base_row["ctr_avg_percent"] / 100.0,
            cpc=base_row["cpc_avg"],
            cr=base_row["cr_avg_percent"] / 100.0,
            aov=base_row["aov_avg"],
        )
        out = calculate_plan_month(base)
        out["campaign_type"] = base_row["campaign_type"]
        out["segment"] = str(base_row.get("segment", "B2C"))
        out["system"] = base_row["system"]
        out["format"] = base_row["format"]
        if is_diy_preset:
            cap_avg = float(base_row.get("available_capacity_avg", 0.0) or 0.0)
            out["available_capacity"] = cap_avg
            out["new_clients_share_pct"] = float(base_row.get("new_clients_share_avg_percent", 0.0) or 0.0)
            out["sov_pct"] = (float(out["impressions"]) / cap_avg * 100.0) if cap_avg > 0 else 0.0
        base_rows.append(out)

    df_base = pd.DataFrame(base_rows)
    if not df_base.empty:
        base_ak_rate = resolve_ak_rate(float(df_base["cost"].sum()), ak_rules_df) if (use_ak_budget_metrics and ak_mode == "percent") else 0.0
        df_base = apply_budget_basis_metrics(
            df_base,
            use_vat_budget_metrics,
            use_ak=use_ak_budget_metrics,
            ak_mode=ak_mode,
            default_ak_rate=base_ak_rate,
            default_ak_fixed_wo_vat=float(ak_fixed_month_wo_vat),
        )
        df_base["ak_rate_pct"] = df_base["ak_rate"] * 100.0

    if df_base.empty:
        st.info("Нет данных для среднего месяца.")
    else:
        df_base_show = df_base.copy()
        df_base_show["ctr_pct"] = df_base_show["ctr"] * 100
        df_base_show["cr_pct"] = df_base_show["cr"] * 100
        df_base_show["drr_pct"] = df_base_show["drr"] * 100

        total_imp = df_base["impressions"].sum()
        total_clicks = df_base["clicks"].sum()
        total_conv = df_base["conversions"].sum()
        total_cost = df_base["cost"].sum()
        total_cost_with_vat = df_base["cost_with_vat"].sum()
        total_ak_wo_vat = df_base["ak_cost_wo_vat"].sum()
        total_cost_with_vat_ak = df_base["cost_with_vat_ak"].sum()
        total_rev = df_base["revenue"].sum()
        if use_ak_budget_metrics:
            total_budget_basis = total_cost_with_vat_ak if use_vat_budget_metrics else (total_cost + total_ak_wo_vat)
        else:
            total_budget_basis = total_cost_with_vat if use_vat_budget_metrics else total_cost

        total_ctr = (total_clicks / total_imp * 100) if total_imp > 0 else 0
        total_cpc = (total_cost / total_clicks) if total_clicks > 0 else 0
        total_cr = (total_conv / total_clicks * 100) if total_clicks > 0 else 0
        total_cpm = (total_budget_basis / (total_imp / 1000)) if total_imp > 0 else 0
        total_cpa = (total_budget_basis / total_conv) if total_conv > 0 else 0
        total_roas = (total_rev / total_budget_basis) if total_budget_basis > 0 else 0
        total_drr = (total_budget_basis / total_rev * 100) if total_rev > 0 else 0
        st.session_state["mp_ref_base"] = {
            "impressions": float(total_imp),
            "clicks": float(total_clicks),
            "conversions": float(total_conv),
            "cost": float(total_cost),
            "cost_with_vat": float(total_cost_with_vat),
            "cost_with_vat_ak": float(total_cost_with_vat_ak),
            "revenue": float(total_rev),
            "ctr": float(total_ctr),
            "cpc": float(total_cpc),
            "cr": float(total_cr),
            "cpo": float(total_cpa),
            "roas": float(total_roas * 100.0),
            "drr": float(total_drr),
        }
        base_by_campaign = {}
        for _, r in df_base.iterrows():
            camp = str(r.get("campaign_type", "")).strip()
            if not camp:
                continue
            imp = float(r.get("impressions", 0.0))
            clicks = float(r.get("clicks", 0.0))
            conv = float(r.get("conversions", 0.0))
            cost = float(r.get("cost", 0.0))
            cost_with_vat = float(r.get("cost_with_vat", 0.0))
            cost_with_vat_ak = float(r.get("cost_with_vat_ak", 0.0))
            rev = float(r.get("revenue", 0.0))
            if use_ak_budget_metrics:
                budget_basis = cost_with_vat_ak if use_vat_budget_metrics else float(r.get("cost", 0.0) + r.get("ak_cost_wo_vat", 0.0))
            else:
                budget_basis = cost_with_vat if use_vat_budget_metrics else cost
            ctr = (clicks / imp * 100.0) if imp > 0 else 0.0
            cpc = (cost / clicks) if clicks > 0 else 0.0
            cr = (conv / clicks * 100.0) if clicks > 0 else 0.0
            cpo = (budget_basis / conv) if conv > 0 else 0.0
            roas = (rev / budget_basis * 100.0) if budget_basis > 0 else 0.0
            drr = (budget_basis / rev * 100.0) if rev > 0 else 0.0
            base_by_campaign[camp] = {
                "impressions": imp,
                "clicks": clicks,
                "conversions": conv,
                "cost": cost,
                "cost_with_vat": cost_with_vat,
                "cost_with_vat_ak": cost_with_vat_ak,
                "revenue": rev,
                "ctr": ctr,
                "cpc": cpc,
                "cr": cr,
                "cpo": cpo,
                "roas": roas,
                "drr": drr,
            }
        st.session_state["mp_ref_base_by_campaign"] = base_by_campaign

        total_row_raw = {
            "campaign_type": "того",
            "segment": "ALL",
            "system": "",
            "format": "",
            "impressions": total_imp,
            "ctr": total_ctr / 100,
            "cpc": total_cpc,
            "cr": total_cr / 100,
            "aov": None,
            "clicks": total_clicks,
            "conversions": total_conv,
            "cost": total_cost,
            "cost_with_vat": total_cost_with_vat,
            "cost_with_vat_ak": total_cost_with_vat_ak,
            "revenue": total_rev,
            "cpm": total_cpm,
            "cpa": total_cpa,
            "roas": total_roas,
            "drr": total_drr / 100,
            "ctr_pct": total_ctr,
            "cr_pct": total_cr,
            "drr_pct": total_drr,
        }
        if is_diy_preset:
            total_capacity = float(df_base["available_capacity"].sum()) if "available_capacity" in df_base.columns else 0.0
            total_row_raw["available_capacity"] = total_capacity
            total_row_raw["sov_pct"] = (total_imp / total_capacity * 100.0) if total_capacity > 0 else 0.0
            total_row_raw["new_clients_share_pct"] = float(df_base["new_clients_share_pct"].mean()) if "new_clients_share_pct" in df_base.columns else 0.0
        segment_total_rows = []
        if show_segment_subtotals and "segment" in df_base.columns:
            for seg_name, seg_df in df_base.groupby("segment"):
                seg_imp = seg_df["impressions"].sum()
                seg_clicks = seg_df["clicks"].sum()
                seg_conv = seg_df["conversions"].sum()
                seg_cost = seg_df["cost"].sum()
                seg_cost_with_vat = seg_df["cost_with_vat"].sum()
                seg_cost_with_vat_ak = seg_df["cost_with_vat_ak"].sum()
                seg_ak_wo_vat = seg_df["ak_cost_wo_vat"].sum()
                seg_rev = seg_df["revenue"].sum()
                if use_ak_budget_metrics:
                    seg_budget_basis = seg_cost_with_vat_ak if use_vat_budget_metrics else (seg_cost + seg_ak_wo_vat)
                else:
                    seg_budget_basis = seg_cost_with_vat if use_vat_budget_metrics else seg_cost
                seg_ctr = (seg_clicks / seg_imp * 100) if seg_imp > 0 else 0
                seg_cpc = (seg_cost / seg_clicks) if seg_clicks > 0 else 0
                seg_cr = (seg_conv / seg_clicks * 100) if seg_clicks > 0 else 0
                seg_cpm = (seg_budget_basis / (seg_imp / 1000)) if seg_imp > 0 else 0
                seg_cpa = (seg_budget_basis / seg_conv) if seg_conv > 0 else 0
                seg_roas = (seg_rev / seg_budget_basis) if seg_budget_basis > 0 else 0
                seg_drr = (seg_budget_basis / seg_rev * 100) if seg_rev > 0 else 0
                segment_total_rows.append(
                    {
                        "campaign_type": f"того {seg_name}",
                        "segment": seg_name,
                        "system": "",
                        "format": "",
                        "impressions": seg_imp,
                        "ctr": seg_ctr / 100,
                        "cpc": seg_cpc,
                        "cr": seg_cr / 100,
                        "aov": None,
                        "clicks": seg_clicks,
                        "conversions": seg_conv,
                        "cost": seg_cost,
                        "cost_with_vat": seg_cost_with_vat,
                        "cost_with_vat_ak": seg_cost_with_vat_ak,
                        "revenue": seg_rev,
                        "cpm": seg_cpm,
                        "cpa": seg_cpa,
                        "roas": seg_roas,
                        "drr": seg_drr / 100,
                        "ctr_pct": seg_ctr,
                        "cr_pct": seg_cr,
                        "drr_pct": seg_drr,
                        "ak_rate_pct": (seg_ak_wo_vat / seg_cost * 100.0) if seg_cost > 0 else 0.0,
                    }
                )
                if is_diy_preset:
                    seg_capacity = float(seg_df["available_capacity"].sum()) if "available_capacity" in seg_df.columns else 0.0
                    segment_total_rows[-1]["available_capacity"] = seg_capacity
                    segment_total_rows[-1]["sov_pct"] = (seg_imp / seg_capacity * 100.0) if seg_capacity > 0 else 0.0
                    segment_total_rows[-1]["new_clients_share_pct"] = float(seg_df["new_clients_share_pct"].mean()) if "new_clients_share_pct" in seg_df.columns else 0.0

        df_base_show = pd.concat(
            [df_base_show, pd.DataFrame(segment_total_rows + [total_row_raw])],
            ignore_index=True,
        )

        df_base_show["impressions"] = df_base_show["impressions"].map(
            lambda x: "" if pd.isna(x) else f"{round(x):,}".replace(",", " ")
        )
        df_base_show["clicks"] = df_base_show["clicks"].map(
            lambda x: "" if pd.isna(x) else f"{round(x):,}".replace(",", " ")
        )
        df_base_show["conversions"] = df_base_show["conversions"].map(
            lambda x: "" if pd.isna(x) else f"{round(x):,}".replace(",", " ")
        )

        df_base_show["cost"] = df_base_show["cost"].map(
            lambda x: "" if pd.isna(x) else f"{round(x):,} ₽".replace(",", " ")
        )
        df_base_show["cost_with_vat"] = df_base_show["cost_with_vat"].map(
            lambda x: "" if pd.isna(x) else f"{round(x):,} ₽".replace(",", " ")
        )
        df_base_show["cost_with_vat_ak"] = df_base_show["cost_with_vat_ak"].map(
            lambda x: "" if pd.isna(x) else f"{round(x):,} ₽".replace(",", " ")
        )
        df_base_show["ak_rate_pct"] = df_base_show["ak_rate_pct"].map(
            lambda x: "" if pd.isna(x) else f"{x:.2f} %"
        )
        df_base_show["revenue"] = df_base_show["revenue"].map(
            lambda x: "" if pd.isna(x) else f"{round(x):,} ₽".replace(",", " ")
        )
        df_base_show["cpc"] = df_base_show["cpc"].map(
            lambda x: "" if pd.isna(x) else f"{x:.2f} ₽".replace(",", " ")
        )
        df_base_show["cpm"] = df_base_show["cpm"].map(
            lambda x: "" if pd.isna(x) else f"{round(x):,} ₽".replace(",", " ")
        )
        df_base_show["cpa"] = df_base_show["cpa"].map(
            lambda x: "" if pd.isna(x) else f"{round(x):,} ₽".replace(",", " ")
        )
        df_base_show["aov"] = df_base_show["aov"].map(
            lambda x: "" if pd.isna(x) else f"{round(x):,} ₽".replace(",", " ")
        )

        df_base_show["ctr_pct"] = df_base_show["ctr_pct"].map(
            lambda x: "" if pd.isna(x) else f"{x:.2f} %"
        )
        df_base_show["cr_pct"] = df_base_show["cr_pct"].map(
            lambda x: "" if pd.isna(x) else f"{x:.2f} %"
        )
        df_base_show["drr_pct"] = df_base_show["drr_pct"].map(
            lambda x: "" if pd.isna(x) else f"{x:.2f} %"
        )
        df_base_show["ROAS"] = df_base_show["roas"].map(
            lambda x: "" if pd.isna(x) else f"{x * 100:.2f} %"
        )

        base_show_cols = [
            "campaign_type",
            "system",
            "format",
            "cost",
            "cost_with_vat",
            "cost_with_vat_ak",
            "ak_rate_pct",
            "impressions",
            "clicks",
            "cpc",
            "ctr_pct",
            "cr_pct",
            "conversions",
            "cpa",
            "aov",
            "revenue",
            "drr_pct",
            "ROAS",
        ]
        if is_diy_preset:
            base_show_cols.insert(3, "segment")
        if is_diy_preset:
            df_base_show["available_capacity"] = df_base_show["available_capacity"].map(
                lambda x: "" if pd.isna(x) else f"{round(x):,}".replace(",", " ")
            )
            df_base_show["sov_pct"] = df_base_show["sov_pct"].map(
                lambda x: "" if pd.isna(x) else f"{x:.2f} %"
            )
            df_base_show["new_clients_share_pct"] = df_base_show["new_clients_share_pct"].map(
                lambda x: "" if pd.isna(x) else f"{x:.2f} %"
            )
            base_show_cols += ["available_capacity", "sov_pct", "new_clients_share_pct"]
        # Порядок метрик задан под бизнес-логику проверки.
        df_base_show = df_base_show[base_show_cols]
        df_base_show = df_base_show.rename(columns=DISPLAY_COL_RENAME)
        if is_diy_preset:
            df_base_show = reorder_rows_with_segment_subtotals(
                df_base_show,
                DISPLAY_COL_RENAME["campaign_type"],
                DISPLAY_COL_RENAME["segment"],
            )
        # Принудительно фиксируем значения в последней строке того (после всех преобразований).
        if len(df_base_show) > 0:
            li = len(df_base_show) - 1
            df_base_show.at[li, DISPLAY_COL_RENAME["campaign_type"]] = "того"
            if DISPLAY_COL_RENAME.get("segment") in df_base_show.columns:
                df_base_show.at[li, DISPLAY_COL_RENAME["segment"]] = "ALL"
            df_base_show.at[li, DISPLAY_COL_RENAME["system"]] = ""
            df_base_show.at[li, DISPLAY_COL_RENAME["format"]] = ""
            if DISPLAY_COL_RENAME["impressions"] in df_base_show.columns:
                df_base_show.at[li, DISPLAY_COL_RENAME["impressions"]] = f"{round(total_imp):,}".replace(",", " ")
            if DISPLAY_COL_RENAME["clicks"] in df_base_show.columns:
                df_base_show.at[li, DISPLAY_COL_RENAME["clicks"]] = f"{round(total_clicks):,}".replace(",", " ")
            if DISPLAY_COL_RENAME["conversions"] in df_base_show.columns:
                df_base_show.at[li, DISPLAY_COL_RENAME["conversions"]] = f"{round(total_conv):,}".replace(",", " ")
            if DISPLAY_COL_RENAME["cost"] in df_base_show.columns:
                df_base_show.at[li, DISPLAY_COL_RENAME["cost"]] = f"{round(total_cost):,} ₽".replace(",", " ")
            if DISPLAY_COL_RENAME["cost_with_vat"] in df_base_show.columns:
                df_base_show.at[li, DISPLAY_COL_RENAME["cost_with_vat"]] = f"{round(total_cost_with_vat):,} ₽".replace(",", " ")
            if DISPLAY_COL_RENAME["cost_with_vat_ak"] in df_base_show.columns:
                df_base_show.at[li, DISPLAY_COL_RENAME["cost_with_vat_ak"]] = f"{round(total_cost_with_vat_ak):,} ₽".replace(",", " ")
            total_ak_rate_pct = (total_ak_wo_vat / total_cost * 100.0) if total_cost > 0 else 0.0
            if DISPLAY_COL_RENAME["ak_rate_pct"] in df_base_show.columns:
                df_base_show.at[li, DISPLAY_COL_RENAME["ak_rate_pct"]] = f"{total_ak_rate_pct:.2f} %"
            if DISPLAY_COL_RENAME["revenue"] in df_base_show.columns:
                df_base_show.at[li, DISPLAY_COL_RENAME["revenue"]] = f"{round(total_rev):,} ₽".replace(",", " ")
            if DISPLAY_COL_RENAME["ctr"] in df_base_show.columns:
                df_base_show.at[li, DISPLAY_COL_RENAME["ctr"]] = f"{total_ctr:.2f} %"
            if DISPLAY_COL_RENAME["cpc"] in df_base_show.columns:
                df_base_show.at[li, DISPLAY_COL_RENAME["cpc"]] = f"{total_cpc:.2f} ₽"
            if DISPLAY_COL_RENAME["cr"] in df_base_show.columns:
                df_base_show.at[li, DISPLAY_COL_RENAME["cr"]] = f"{total_cr:.2f} %"
            if DISPLAY_COL_RENAME["cpa"] in df_base_show.columns:
                df_base_show.at[li, DISPLAY_COL_RENAME["cpa"]] = f"{round(total_cpa):,} ₽".replace(",", " ")
            if "ROAS" in df_base_show.columns:
                df_base_show.at[li, "ROAS"] = f"{total_roas * 100:.2f} %"
            if DISPLAY_COL_RENAME["drr"] in df_base_show.columns:
                df_base_show.at[li, DISPLAY_COL_RENAME["drr"]] = f"{total_drr:.2f} %"
            if DISPLAY_COL_RENAME["aov"] in df_base_show.columns:
                df_base_show.at[li, DISPLAY_COL_RENAME["aov"]] = ""

        def highlight_total_row(row):
            style = [""] * len(row)
            camp_col = DISPLAY_COL_RENAME.get("campaign_type", "Название кампании")
            camp_val = str(row.get(camp_col, ""))
            if camp_val == "того":
                style = [
                    f"background-color: #00CDC5; color: #081521; font-weight: 700; border-top: 2px solid {THEME_BORDER};"
                ] * len(row)
            elif camp_val.startswith("того "):
                style = [
                    f"background-color: #2C6E75; color: #DDEAF0; font-weight: 650; border-top: 1px solid {THEME_BORDER};"
                ] * len(row)
            return style

        styled_base = df_base_show.style.apply(highlight_total_row, axis=1)
        st.dataframe(styled_base, use_container_width=True)

    # ---------- 3. Коэффициенты и результаты по месяцам ----------

    with st.expander("3. Коэффициенты и результаты по месяцам", expanded=True):
        show_month_details = True

        coeff_values = {}
        for cs in coeff_sets:
            cs_name = cs["name"]
            cs_type = normalize_coeff_set_type(cs.get("type", "Спрос (по запросам)"))
            df_res = cs.get("result")
            if df_res is None or df_res.empty:
                continue

            if cs_type == "Спрос (по запросам)":
                col_coeff = "Коэф."
                metric_type = "demand"
            elif cs_type == "AOV (средний чек)":
                col_coeff = "Коэф. AOV"
                metric_type = "aov"
            elif cs_type == "Медийные хвосты":
                col_coeff = "Коэф."
                metric_type = "media_tail"
            else:
                col_coeff = "Коэф."
                metric_type = "custom"

            if col_coeff not in df_res.columns or "Номер месяца" not in df_res.columns:
                continue

            for _, r in df_res.iterrows():
                try:
                    m = int(r["Номер месяца"])
                    k = float(r[col_coeff])
                except Exception:
                    continue
                if metric_type == "custom":
                    coeff_values[(cs_name, m, "demand")] = k
                    coeff_values[(cs_name, m, "aov")] = float(
                        r["Коэф. AOV"] if "Коэф. AOV" in df_res.columns else k
                    )
                else:
                    coeff_values[(cs_name, m, metric_type)] = k

        def get_k_demand(campaign_type: str, month_num: int) -> float:
            set_name = demand_link_map.get(campaign_type)
            if not set_name:
                return 1.0
            return coeff_values.get((set_name, month_num, "demand"), 1.0)

        def get_k_aov(campaign_type: str, month_num: int) -> float:
            set_name = aov_link_map.get(campaign_type)
            if not set_name:
                return 1.0
            return coeff_values.get((set_name, month_num, "aov"), 1.0)

        def get_k_media_tail(campaign_type: str, month_num: int) -> float:
            set_name = media_tail_link_map.get(campaign_type)
            if not set_name:
                return 1.0
            return coeff_values.get((set_name, month_num, "media_tail"), 1.0)

        def get_k_capacity(campaign_type: str, month_num: int) -> float:
            set_name = capacity_link_map.get(campaign_type)
            if not set_name:
                set_name = demand_link_map.get(campaign_type, "")
            if not set_name:
                return 1.0
            if (set_name, month_num, "demand") in coeff_values:
                return coeff_values[(set_name, month_num, "demand")]
            if (set_name, month_num, "media_tail") in coeff_values:
                return coeff_values[(set_name, month_num, "media_tail")]
            if (set_name, month_num, "aov") in coeff_values:
                return coeff_values[(set_name, month_num, "aov")]
            return 1.0

        all_months_results = []

        for month in selected_month_nums:
            month_name = month_names_full[month]

            coeff_rows = []
            base_capacity_map = {
                str(r["campaign_type"]): float(r.get("available_capacity_avg", 0.0) or 0.0)
                for _, r in campaigns.iterrows()
            }
            for _, base_row in campaigns.iterrows():
                camp_type = str(base_row["campaign_type"])

                k_demand = get_k_demand(camp_type, month)
                k_media_tail = get_k_media_tail(camp_type, month)
                el = elasticity_map.get(
                    camp_type,
                    {"cpc_div": 1.0, "ctr_div": 2.0, "cr_div": 10.0},
                )

                k_imp = k_demand * k_media_tail
                demand_delta = (k_demand - 1.0)
                k_cpc = 1.0 + demand_delta / el["cpc_div"]
                k_ctr = 1.0 - demand_delta / el["ctr_div"]
                k_cr = 1.0 - demand_delta / el["cr_div"]
                k_ctr = max(0.0, k_ctr)
                k_cr = max(0.0, k_cr)
                k_aov = get_k_aov(camp_type, month)
                k_capacity = get_k_capacity(camp_type, month)

                coeff_rows.append(
                    {
                        "campaign_type": camp_type,
                        "k_imp": k_imp,
                        "k_capacity": k_capacity,
                        "available_capacity": float(base_row.get("available_capacity_avg", 0.0) or 0.0) * k_capacity,
                        "new_clients_share_pct": float(base_row.get("new_clients_share_avg_percent", 0.0) or 0.0),
                        "k_ctr": k_ctr,
                        "k_cpc": k_cpc,
                        "k_cr": k_cr,
                        "k_aov": k_aov,
                    }
                )

            coeffs_default = pd.DataFrame(coeff_rows)
            if not is_diy_preset:
                coeffs_default = coeffs_default.drop(
                    columns=["k_capacity", "available_capacity", "new_clients_share_pct"],
                    errors="ignore",
                )
            else:
                coeffs_default = coeffs_default[
                    [
                        "campaign_type",
                        "k_imp",
                        "k_cpc",
                        "k_ctr",
                        "k_cr",
                        "k_aov",
                        "k_capacity",
                        "available_capacity",
                        "new_clients_share_pct",
                    ]
                ]
            coeffs_month = coeffs_default.copy()
            coeffs_month["campaign_type"] = coeffs_month["campaign_type"].astype(str)

            rows = []
            # если хотим показывать детали по месяцам, создаём редактор
            if show_month_details:
                st.markdown(f"#### Месяц {month}: {month_name}")
                head_left, head_right = st.columns([1.8, 1.0], vertical_alignment="top")
                with head_left:
                    st.caption("Результаты для этого месяца")
                with head_right:
                    st.caption("Коэффициенты месяца (k_imp: показы, k_ctr: CTR, k_cpc: CPC, k_cr: CR, k_aov: AOV)")
                    if is_diy_preset:
                        st.caption("Для DIY: емкость = база × k_capacity; коэффициент k_capacity можно менять вручную.")

                col_left, col_right = st.columns([1.8, 1.0], vertical_alignment="top")
                month_row_height = max(36, int(ui_table_row_px))
                month_rows_count = max(len(coeffs_month), 1)
                month_table_height = max(140, min(520, 38 + month_rows_count * month_row_height + 6))
                with col_right:
                    edited_coeffs_month = st.data_editor(
                        coeffs_month,
                        num_rows="fixed",
                        use_container_width=True,
                        row_height=month_row_height,
                        height=month_table_height,
                        key=f"coeffs_month_{month}_v2",
                        column_config={
                            "campaign_type": st.column_config.TextColumn(
                                "Тип РК / Название", disabled=True, width="medium"
                            ),
                            "k_imp": st.column_config.NumberColumn(
                                "k_imp", format="%.2f", width="small",
                                help=mhelp("k_imp")
                            ),
                            "k_cpc": st.column_config.NumberColumn(
                                "k_cpc", format="%.2f", width="small",
                                help=mhelp("k_cpc")
                            ),
                            "k_ctr": st.column_config.NumberColumn(
                                "k_ctr", format="%.2f", width="small",
                                help=mhelp("k_ctr")
                            ),
                            "k_cr": st.column_config.NumberColumn(
                                "k_cr", format="%.2f", width="small",
                                help=mhelp("k_cr")
                            ),
                            "k_aov": st.column_config.NumberColumn(
                                "k_aov", format="%.2f", width="small",
                                help=mhelp("k_aov")
                            ),
                            "k_capacity": st.column_config.NumberColumn(
                                "k_capacity", format="%.2f", width="small",
                                disabled=not is_diy_preset,
                            ),
                            "available_capacity": st.column_config.NumberColumn(
                                "Доступная емкость", format="%.0f", width="small",
                                disabled=True,
                            ),
                            "new_clients_share_pct": st.column_config.NumberColumn(
                                "Доля новых клиентов, %", format="%.2f", width="small",
                                disabled=not is_diy_preset,
                            ),
                        },
                    )
                    edited_coeffs_month = edited_coeffs_month.copy()
                    edited_coeffs_month["campaign_type"] = edited_coeffs_month["campaign_type"].astype(str)
                    if is_diy_preset and "available_capacity" in edited_coeffs_month.columns:
                        edited_coeffs_month["available_capacity"] = edited_coeffs_month.apply(
                            lambda r: float(base_capacity_map.get(str(r.get("campaign_type", "")), 0.0))
                            * float(r.get("k_capacity", 1.0) or 1.0),
                            axis=1,
                        )
                    coeffs_month = edited_coeffs_month
            else:
                # если подробности скрыты — берём coeffs_default как есть
                coeffs_month = coeffs_month.copy()
                col_left = st  # заглушка, чтобы ниже не падало

            for _, base_row in campaigns.iterrows():
                campaign_type = str(base_row["campaign_type"])

                k_row = coeffs_month[coeffs_month["campaign_type"] == campaign_type]
                if k_row.empty:
                    k_imp = k_capacity = k_ctr = k_cpc = k_cr = k_aov = 1.0
                    available_capacity_month = float(base_row.get("available_capacity_avg", 0.0) or 0.0)
                    new_clients_share_month = 0.0
                else:
                    k_row = k_row.iloc[0]
                    k_imp = float(k_row["k_imp"])
                    k_capacity = float(k_row.get("k_capacity", 1.0) or 1.0)
                    k_ctr = float(k_row["k_ctr"])
                    k_cpc = float(k_row["k_cpc"])
                    k_cr = float(k_row["k_cr"])
                    k_aov = float(k_row["k_aov"])
                    available_capacity_month = float(base_row.get("available_capacity_avg", 0.0) or 0.0) * k_capacity
                    new_clients_share_month = float(k_row.get("new_clients_share_pct", 0.0) or 0.0)

                base = PlanInput(
                    impressions=base_row["impressions_avg"],
                    ctr=base_row["ctr_avg_percent"] / 100.0,
                    cpc=base_row["cpc_avg"],
                    cr=base_row["cr_avg_percent"] / 100.0,
                    aov=base_row["aov_avg"],
                )

                month_inp = PlanInput(
                    impressions=base.impressions * k_imp,
                    ctr=base.ctr * k_ctr,
                    cpc=base.cpc * k_cpc,
                    cr=base.cr * k_cr,
                    aov=base.aov * k_aov,
                )

                out = calculate_plan_month(month_inp)
                out["month_num"] = month
                out["month_name"] = month_name
                out["campaign_type"] = campaign_type
                out["segment"] = str(base_row.get("segment", "B2C"))
                out["system"] = base_row["system"]
                out["format"] = base_row["format"]
                if is_diy_preset:
                    out["available_capacity"] = available_capacity_month
                    out["sov_pct"] = (float(out["impressions"]) / available_capacity_month * 100.0) if available_capacity_month > 0 else 0.0
                    out["new_clients_share_pct"] = new_clients_share_month
                rows.append(out)

            df_month = pd.DataFrame(rows)
            if not df_month.empty:
                month_total_wo_vat = float(df_month["cost"].sum())
                month_ak_rate = resolve_ak_rate(month_total_wo_vat, ak_rules_df) if (use_ak_budget_metrics and ak_mode == "percent") else 0.0
                df_month = apply_budget_basis_metrics(
                    df_month,
                    use_vat_budget_metrics,
                    use_ak=use_ak_budget_metrics,
                    ak_mode=ak_mode,
                    default_ak_rate=month_ak_rate,
                    default_ak_fixed_wo_vat=float(ak_fixed_month_wo_vat),
                )
                df_month["ak_rate_pct"] = df_month["ak_rate"] * 100.0
            all_months_results.append(df_month)

            if show_month_details:
                with col_left:
                    if not df_month.empty:
                        df_rows_show = df_month.copy()

                        df_rows_show["ctr_pct"] = df_rows_show["ctr"] * 100
                        df_rows_show["cr_pct"] = df_rows_show["cr"] * 100
                        df_rows_show["drr_pct"] = df_rows_show["drr"] * 100

                        total_imp = df_month["impressions"].sum()
                        total_clicks = df_month["clicks"].sum()
                        total_conv = df_month["conversions"].sum()
                        total_cost = df_month["cost"].sum()
                        total_cost_with_vat = df_month["cost_with_vat"].sum()
                        total_cost_with_vat_ak = df_month["cost_with_vat_ak"].sum()
                        total_ak_wo_vat = df_month["ak_cost_wo_vat"].sum()
                        month_ak_rate_effective = (total_ak_wo_vat / total_cost) if total_cost > 0 else 0.0
                        total_rev = df_month["revenue"].sum()
                        if use_ak_budget_metrics:
                            total_budget_basis = total_cost_with_vat_ak if use_vat_budget_metrics else (total_cost + total_ak_wo_vat)
                        else:
                            total_budget_basis = total_cost_with_vat if use_vat_budget_metrics else total_cost

                        total_ctr = (total_clicks / total_imp * 100) if total_imp > 0 else 0
                        total_cpc = (total_cost / total_clicks) if total_clicks > 0 else 0
                        total_cr = (total_conv / total_clicks * 100) if total_clicks > 0 else 0
                        total_cpm = (total_budget_basis / (total_imp / 1000)) if total_imp > 0 else 0
                        total_cpa = (total_budget_basis / total_conv) if total_conv > 0 else 0
                        total_roas = (total_rev / total_budget_basis) if total_budget_basis > 0 else 0
                        total_drr = (total_budget_basis / total_rev * 100) if total_rev > 0 else 0

                        total_row_month = {
                            "campaign_type": "того",
                            "segment": "ALL",
                            "system": "",
                            "format": "",
                            "impressions": total_imp,
                            "ctr": total_ctr / 100,
                            "cpc": total_cpc,
                            "cr": total_cr / 100,
                            "aov": None,
                            "clicks": total_clicks,
                            "conversions": total_conv,
                            "cost": total_cost,
                            "cost_with_vat": total_cost_with_vat,
                            "cost_with_vat_ak": total_cost_with_vat_ak,
                            "ak_rate_pct": month_ak_rate_effective * 100.0,
                            "revenue": total_rev,
                            "cpm": total_cpm,
                            "cpa": total_cpa,
                            "roas": total_roas,
                            "drr": total_drr / 100,
                            "ctr_pct": total_ctr,
                            "cr_pct": total_cr,
                            "drr_pct": total_drr,
                        }
                        if is_diy_preset:
                            total_capacity = float(df_month["available_capacity"].sum()) if "available_capacity" in df_month.columns else 0.0
                            total_row_month["available_capacity"] = total_capacity
                            total_row_month["sov_pct"] = (total_imp / total_capacity * 100.0) if total_capacity > 0 else 0.0
                            total_row_month["new_clients_share_pct"] = float(df_month["new_clients_share_pct"].mean()) if "new_clients_share_pct" in df_month.columns else 0.0
                        segment_total_rows_month = []
                        if show_segment_subtotals and "segment" in df_month.columns:
                            for seg_name, seg_df in df_month.groupby("segment"):
                                seg_imp = seg_df["impressions"].sum()
                                seg_clicks = seg_df["clicks"].sum()
                                seg_conv = seg_df["conversions"].sum()
                                seg_cost = seg_df["cost"].sum()
                                seg_cost_with_vat = seg_df["cost_with_vat"].sum()
                                seg_cost_with_vat_ak = seg_df["cost_with_vat_ak"].sum()
                                seg_ak_wo_vat = seg_df["ak_cost_wo_vat"].sum()
                                seg_rev = seg_df["revenue"].sum()
                                if use_ak_budget_metrics:
                                    seg_budget_basis = seg_cost_with_vat_ak if use_vat_budget_metrics else (seg_cost + seg_ak_wo_vat)
                                else:
                                    seg_budget_basis = seg_cost_with_vat if use_vat_budget_metrics else seg_cost
                                seg_ctr = (seg_clicks / seg_imp * 100) if seg_imp > 0 else 0
                                seg_cpc = (seg_cost / seg_clicks) if seg_clicks > 0 else 0
                                seg_cr = (seg_conv / seg_clicks * 100) if seg_clicks > 0 else 0
                                seg_cpm = (seg_budget_basis / (seg_imp / 1000)) if seg_imp > 0 else 0
                                seg_cpa = (seg_budget_basis / seg_conv) if seg_conv > 0 else 0
                                seg_roas = (seg_rev / seg_budget_basis) if seg_budget_basis > 0 else 0
                                seg_drr = (seg_budget_basis / seg_rev * 100) if seg_rev > 0 else 0
                                seg_row = {
                                    "campaign_type": f"того {seg_name}",
                                    "segment": seg_name,
                                    "system": "",
                                    "format": "",
                                    "impressions": seg_imp,
                                    "ctr": seg_ctr / 100,
                                    "cpc": seg_cpc,
                                    "cr": seg_cr / 100,
                                    "aov": None,
                                    "clicks": seg_clicks,
                                    "conversions": seg_conv,
                                    "cost": seg_cost,
                                    "cost_with_vat": seg_cost_with_vat,
                                    "cost_with_vat_ak": seg_cost_with_vat_ak,
                                    "ak_rate_pct": (seg_ak_wo_vat / seg_cost * 100.0) if seg_cost > 0 else 0.0,
                                    "revenue": seg_rev,
                                    "cpm": seg_cpm,
                                    "cpa": seg_cpa,
                                    "roas": seg_roas,
                                    "drr": seg_drr / 100,
                                    "ctr_pct": seg_ctr,
                                    "cr_pct": seg_cr,
                                    "drr_pct": seg_drr,
                                }
                                if is_diy_preset:
                                    seg_cap = float(seg_df["available_capacity"].sum()) if "available_capacity" in seg_df.columns else 0.0
                                    seg_row["available_capacity"] = seg_cap
                                    seg_row["sov_pct"] = (seg_imp / seg_cap * 100.0) if seg_cap > 0 else 0.0
                                    seg_row["new_clients_share_pct"] = float(seg_df["new_clients_share_pct"].mean()) if "new_clients_share_pct" in seg_df.columns else 0.0
                                segment_total_rows_month.append(seg_row)

                        # Основная таблица: только типы РК (без TOTAL), чтобы построчно
                        # совпадать с таблицей коэффициентов справа.
                        df_rows_show["impressions"] = df_rows_show["impressions"].map(
                            lambda x: "" if pd.isna(x) else f"{round(x):,}".replace(",", " ")
                        )
                        df_rows_show["clicks"] = df_rows_show["clicks"].map(
                            lambda x: "" if pd.isna(x) else f"{round(x):,}".replace(",", " ")
                        )
                        df_rows_show["conversions"] = df_rows_show["conversions"].map(
                            lambda x: "" if pd.isna(x) else f"{round(x):,}".replace(",", " ")
                        )

                        df_rows_show["cost"] = df_rows_show["cost"].map(
                            lambda x: "" if pd.isna(x) else f"{round(x):,} ₽".replace(",", " ")
                        )
                        df_rows_show["cost_with_vat"] = df_rows_show["cost_with_vat"].map(
                            lambda x: "" if pd.isna(x) else f"{round(x):,} ₽".replace(",", " ")
                        )
                        df_rows_show["cost_with_vat_ak"] = df_rows_show["cost_with_vat_ak"].map(
                            lambda x: "" if pd.isna(x) else f"{round(x):,} ₽".replace(",", " ")
                        )
                        df_rows_show["ak_rate_pct"] = df_rows_show["ak_rate_pct"].map(
                            lambda x: "" if pd.isna(x) else f"{x:.2f} %"
                        )
                        df_rows_show["revenue"] = df_rows_show["revenue"].map(
                            lambda x: "" if pd.isna(x) else f"{round(x):,} ₽".replace(",", " ")
                        )
                        df_rows_show["cpc"] = df_rows_show["cpc"].map(
                            lambda x: "" if pd.isna(x) else f"{x:.2f} ₽".replace(",", " ")
                        )
                        df_rows_show["cpm"] = df_rows_show["cpm"].map(
                            lambda x: "" if pd.isna(x) else f"{round(x):,} ₽".replace(",", " ")
                        )
                        df_rows_show["cpa"] = df_rows_show["cpa"].map(
                            lambda x: "" if pd.isna(x) else f"{round(x):,} ₽".replace(",", " ")
                        )
                        df_rows_show["aov"] = df_rows_show["aov"].map(
                            lambda x: "" if pd.isna(x) else f"{round(x):,} ₽".replace(",", " ")
                        )

                        df_rows_show["ctr_pct"] = df_rows_show["ctr_pct"].map(
                            lambda x: "" if pd.isna(x) else f"{x:.2f} %"
                        )
                        df_rows_show["cr_pct"] = df_rows_show["cr_pct"].map(
                            lambda x: "" if pd.isna(x) else f"{x:.2f} %"
                        )
                        df_rows_show["drr_pct"] = df_rows_show["drr_pct"].map(
                            lambda x: "" if pd.isna(x) else f"{x:.2f} %"
                        )
                        df_rows_show["ROAS"] = df_rows_show["roas"].map(
                            lambda x: "" if pd.isna(x) else f"{x * 100:.2f} %"
                        )
                        if is_diy_preset:
                            df_rows_show["available_capacity"] = df_rows_show["available_capacity"].map(
                                lambda x: "" if pd.isna(x) else f"{round(x):,}".replace(",", " ")
                            )
                            df_rows_show["sov_pct"] = df_rows_show["sov_pct"].map(
                                lambda x: "" if pd.isna(x) else f"{x:.2f} %"
                            )
                            df_rows_show["new_clients_share_pct"] = df_rows_show["new_clients_share_pct"].map(
                                lambda x: "" if pd.isna(x) else f"{x:.2f} %"
                            )

                        month_show_cols = [
                            "campaign_type",
                            "system",
                            "format",
                            "cost",
                            "cost_with_vat",
                            "cost_with_vat_ak",
                            "ak_rate_pct",
                            "impressions",
                            "clicks",
                            "cpc",
                            "ctr_pct",
                            "cr_pct",
                            "conversions",
                            "cpa",
                            "aov",
                            "revenue",
                            "drr_pct",
                            "ROAS",
                        ]
                        if is_diy_preset:
                            month_show_cols.insert(3, "segment")
                        if is_diy_preset:
                            month_show_cols += ["available_capacity", "sov_pct", "new_clients_share_pct"]
                        df_rows_show = df_rows_show[month_show_cols]
                        df_rows_show = df_rows_show.rename(columns=DISPLAY_COL_RENAME)

                        # TOTAL отдельной строкой ниже основной таблицы.
                        total_month_show = pd.DataFrame(segment_total_rows_month + [total_row_month])
                        total_month_show["impressions"] = total_month_show["impressions"].map(lambda x: f"{round(x):,}".replace(",", " "))
                        total_month_show["clicks"] = total_month_show["clicks"].map(lambda x: f"{round(x):,}".replace(",", " "))
                        total_month_show["conversions"] = total_month_show["conversions"].map(lambda x: f"{round(x):,}".replace(",", " "))
                        total_month_show["cost"] = total_month_show["cost"].map(lambda x: f"{round(x):,} ₽".replace(",", " "))
                        total_month_show["cost_with_vat"] = total_month_show["cost_with_vat"].map(lambda x: f"{round(x):,} ₽".replace(",", " "))
                        total_month_show["cost_with_vat_ak"] = total_month_show["cost_with_vat_ak"].map(lambda x: f"{round(x):,} ₽".replace(",", " "))
                        total_month_show["ak_rate_pct"] = total_month_show["ak_rate_pct"].map(lambda x: f"{x:.2f} %")
                        total_month_show["revenue"] = total_month_show["revenue"].map(lambda x: f"{round(x):,} ₽".replace(",", " "))
                        total_month_show["cpc"] = total_month_show["cpc"].map(lambda x: f"{x:.2f} ₽")
                        total_month_show["cpm"] = total_month_show["cpm"].map(lambda x: f"{round(x):,} ₽".replace(",", " "))
                        total_month_show["cpa"] = total_month_show["cpa"].map(lambda x: f"{round(x):,} ₽".replace(",", " "))
                        total_month_show["aov"] = ""
                        total_month_show["ctr_pct"] = total_month_show["ctr_pct"].map(lambda x: f"{x:.2f} %")
                        total_month_show["cr_pct"] = total_month_show["cr_pct"].map(lambda x: f"{x:.2f} %")
                        total_month_show["drr_pct"] = total_month_show["drr_pct"].map(lambda x: f"{x:.2f} %")
                        total_month_show["ROAS"] = total_month_show["roas"].map(lambda x: f"{x * 100:.2f} %")
                        if is_diy_preset:
                            total_month_show["available_capacity"] = total_month_show["available_capacity"].map(
                                lambda x: f"{round(x):,}".replace(",", " ")
                            )
                            total_month_show["sov_pct"] = total_month_show["sov_pct"].map(lambda x: f"{x:.2f} %")
                            total_month_show["new_clients_share_pct"] = total_month_show["new_clients_share_pct"].map(lambda x: f"{x:.2f} %")
                        total_month_cols = [
                            "campaign_type", "system", "format",
                            "cost", "cost_with_vat", "cost_with_vat_ak", "ak_rate_pct",
                            "impressions", "clicks", "cpc", "ctr_pct", "cr_pct",
                            "conversions", "cpa", "aov", "revenue", "drr_pct", "ROAS",
                        ]
                        if is_diy_preset:
                            total_month_cols.insert(3, "segment")
                        if is_diy_preset:
                            total_month_cols += ["available_capacity", "sov_pct", "new_clients_share_pct"]
                        total_month_show = total_month_show[total_month_cols].rename(columns=DISPLAY_COL_RENAME)

                        # Простой вариант: возвращаем TOTAL обратно в основную таблицу.
                        df_month_with_total_show = pd.concat([df_rows_show, total_month_show], ignore_index=True)
                        if is_diy_preset:
                            df_month_with_total_show = reorder_rows_with_segment_subtotals(
                                df_month_with_total_show,
                                DISPLAY_COL_RENAME["campaign_type"],
                                DISPLAY_COL_RENAME["segment"],
                            )

                        def _highlight_total_row_merged(row):
                            style = [""] * len(row)
                            camp_col = DISPLAY_COL_RENAME.get("campaign_type", "Название кампании")
                            camp_val = str(row.get(camp_col, ""))
                            if camp_val == "того":
                                style = [
                                    f"background-color: #00CDC5; color: #081521; font-weight: 700; border-top: 2px solid {THEME_BORDER};"
                                ] * len(row)
                            elif camp_val.startswith("того "):
                                style = [
                                    f"background-color: #2C6E75; color: #DDEAF0; font-weight: 650; border-top: 1px solid {THEME_BORDER};"
                                ] * len(row)
                            return style

                        month_table_height_with_total = max(
                            140,
                            min(560, month_table_height + month_row_height + 4),
                        )
                        st.dataframe(
                            df_month_with_total_show.style.apply(_highlight_total_row_merged, axis=1),
                            use_container_width=True,
                            height=month_table_height_with_total,
                        )
                    else:
                        st.info("Нет данных для этого месяца.")

    # ---------- 4. тоги по выбранным месяцам (TOTAL) ----------

    if all_months_results:
        df_all = pd.concat(all_months_results, ignore_index=True)
    else:
        df_all = pd.DataFrame()

    if not df_all.empty:
        ref_imp = float(df_all["impressions"].sum())
        ref_clicks = float(df_all["clicks"].sum())
        ref_conv = float(df_all["conversions"].sum())
        ref_cost = float(df_all["cost"].sum())
        ref_cost_with_vat = float(df_all["cost_with_vat"].sum())
        ref_cost_with_vat_ak = float(df_all["cost_with_vat_ak"].sum())
        ref_ak_wo_vat = float(df_all["ak_cost_wo_vat"].sum())
        ref_rev = float(df_all["revenue"].sum())
        if use_ak_budget_metrics:
            ref_budget_basis = ref_cost_with_vat_ak if use_vat_budget_metrics else (ref_cost + ref_ak_wo_vat)
        else:
            ref_budget_basis = ref_cost_with_vat if use_vat_budget_metrics else ref_cost
        ref_ctr = (ref_clicks / ref_imp * 100.0) if ref_imp > 0 else 0.0
        ref_cpc = (ref_cost / ref_clicks) if ref_clicks > 0 else 0.0
        ref_cr = (ref_conv / ref_clicks * 100.0) if ref_clicks > 0 else 0.0
        ref_cpo = (ref_budget_basis / ref_conv) if ref_conv > 0 else 0.0
        ref_roas = (ref_rev / ref_budget_basis * 100.0) if ref_budget_basis > 0 else 0.0
        ref_drr = (ref_budget_basis / ref_rev * 100.0) if ref_rev > 0 else 0.0
        st.session_state["mp_ref_total"] = {
            "impressions": ref_imp,
            "clicks": ref_clicks,
            "conversions": ref_conv,
            "cost": ref_cost,
            "cost_with_vat": ref_cost_with_vat,
            "cost_with_vat_ak": ref_cost_with_vat_ak,
            "revenue": ref_rev,
            "ctr": ref_ctr,
            "cpc": ref_cpc,
            "cr": ref_cr,
            "cpo": ref_cpo,
            "roas": ref_roas,
            "drr": ref_drr,
        }
        total_by_campaign = {}
        agg_ct = df_all.groupby("campaign_type", as_index=False).agg(
            impressions=("impressions", "sum"),
            clicks=("clicks", "sum"),
            conversions=("conversions", "sum"),
            cost=("cost", "sum"),
            cost_with_vat=("cost_with_vat", "sum"),
            cost_with_vat_ak=("cost_with_vat_ak", "sum"),
            ak_cost_wo_vat=("ak_cost_wo_vat", "sum"),
            revenue=("revenue", "sum"),
        )
        for _, r in agg_ct.iterrows():
            camp = str(r.get("campaign_type", "")).strip()
            if not camp:
                continue
            imp = float(r.get("impressions", 0.0))
            clicks = float(r.get("clicks", 0.0))
            conv = float(r.get("conversions", 0.0))
            cost = float(r.get("cost", 0.0))
            cost_with_vat = float(r.get("cost_with_vat", 0.0))
            cost_with_vat_ak = float(r.get("cost_with_vat_ak", 0.0))
            ak_wo_vat = float(r.get("ak_cost_wo_vat", 0.0))
            rev = float(r.get("revenue", 0.0))
            if use_ak_budget_metrics:
                budget_basis = cost_with_vat_ak if use_vat_budget_metrics else (cost + ak_wo_vat)
            else:
                budget_basis = cost_with_vat if use_vat_budget_metrics else cost
            ctr = (clicks / imp * 100.0) if imp > 0 else 0.0
            cpc = (cost / clicks) if clicks > 0 else 0.0
            cr = (conv / clicks * 100.0) if clicks > 0 else 0.0
            cpo = (budget_basis / conv) if conv > 0 else 0.0
            roas = (rev / budget_basis * 100.0) if budget_basis > 0 else 0.0
            drr = (budget_basis / rev * 100.0) if rev > 0 else 0.0
            total_by_campaign[camp] = {
                "impressions": imp,
                "clicks": clicks,
                "conversions": conv,
                "cost": cost,
                "cost_with_vat": cost_with_vat,
                "cost_with_vat_ak": cost_with_vat_ak,
                "revenue": rev,
                "ctr": ctr,
                "cpc": cpc,
                "cr": cr,
                "cpo": cpo,
                "roas": roas,
                "drr": drr,
            }
        st.session_state["mp_ref_total_by_campaign"] = total_by_campaign
    else:
        st.session_state["mp_ref_total"] = None
        st.session_state["mp_ref_total_by_campaign"] = {}

    with st.sidebar:
        st.markdown("---")
        st.markdown("### Быстрая сверка")
        # Import compatibility: old state could store a legacy mode value.
        if st.session_state.get("mp_ref_mode") == "Включено":
            st.session_state["mp_ref_mode"] = "Средний месяц"
        # If references are missing after import/rerun, rebuild them from base campaigns.
        if not st.session_state.get("mp_ref_base"):
            _bootstrap_reference_from_campaigns(st.session_state.get("campaigns_df"))
        if is_diy_preset:
            st.selectbox(
                "Сегмент в расчете",
                options=["Все", "B2B", "B2C"],
                key="plan_segment_filter_sidebar",
                on_change=_sync_segment_from_sidebar,
            )
        ref_mode = st.selectbox(
            "Что закрепить",
            options=["Выключено", "Средний месяц", "TOTAL выбранных месяцев"],
            key="mp_ref_mode",
        )
        all_ref_campaigns = sorted(
            set(st.session_state.get("mp_ref_base_by_campaign", {}).keys())
            | set(st.session_state.get("mp_ref_total_by_campaign", {}).keys())
        )
        valid_ref_campaigns = ["Все типы РК"] + all_ref_campaigns
        if st.session_state.get("mp_ref_campaign") not in valid_ref_campaigns:
            st.session_state["mp_ref_campaign"] = "Все типы РК"
        ref_campaign = st.selectbox(
            "Тип РК",
            options=valid_ref_campaigns,
            key="mp_ref_campaign",
        )
        ref_src = None
        if ref_mode == "Средний месяц":
            if ref_campaign == "Все типы РК":
                ref_src = st.session_state.get("mp_ref_base")
            else:
                ref_src = st.session_state.get("mp_ref_base_by_campaign", {}).get(ref_campaign)
        elif ref_mode == "TOTAL выбранных месяцев":
            if ref_campaign == "Все типы РК":
                ref_src = st.session_state.get("mp_ref_total")
            else:
                ref_src = st.session_state.get("mp_ref_total_by_campaign", {}).get(ref_campaign)

        # Fallback: right after import TOTAL may be unavailable until full plan calc.
        # Show base reference instead of empty block.
        if ref_src is None and ref_mode == "TOTAL выбранных месяцев":
            if ref_campaign == "Все типы РК":
                ref_src = st.session_state.get("mp_ref_base")
            else:
                ref_src = st.session_state.get("mp_ref_base_by_campaign", {}).get(ref_campaign)

        if ref_mode != "Выключено":
            if ref_src:
                st.dataframe(_build_ref_df(ref_src), use_container_width=True, hide_index=True)
            else:
                st.caption("Нет данных для выбранного режима сверки.")

    with st.expander("4. тоги по выбранным месяцам (TOTAL)", expanded=True):
        if df_all.empty:
            st.info("Нет данных для итогов. Заполните медиаплан по месяцам.")
        else:
            st.caption("Сводные показатели по всем выбранным месяцам и типам РК.")

            agg = df_all.groupby(["month_num", "month_name"], as_index=False).agg(
                impressions=("impressions", "sum"),
                clicks=("clicks", "sum"),
                conversions=("conversions", "sum"),
                cost=("cost", "sum"),
                cost_with_vat=("cost_with_vat", "sum"),
                cost_with_vat_ak=("cost_with_vat_ak", "sum"),
                ak_cost_wo_vat=("ak_cost_wo_vat", "sum"),
                revenue=("revenue", "sum"),
            )
            if use_ak_budget_metrics:
                budget_series = agg["cost_with_vat_ak"] if use_vat_budget_metrics else (agg["cost"] + agg["ak_cost_wo_vat"])
            else:
                budget_series = agg["cost_with_vat"] if use_vat_budget_metrics else agg["cost"]

            agg["ctr"] = np.where(
                agg["impressions"] > 0,
                agg["clicks"] / agg["impressions"] * 100,
                0.0,
            )
            agg["cr"] = np.where(
                agg["clicks"] > 0,
                agg["conversions"] / agg["clicks"] * 100,
                0.0,
            )
            agg["cpc"] = np.where(
                agg["clicks"] > 0,
                agg["cost"] / agg["clicks"],
                0.0,
            )
            agg["cpm"] = np.where(
                agg["impressions"] > 0,
                budget_series / (agg["impressions"] / 1000),
                0.0,
            )
            agg["cpa"] = np.where(
                agg["conversions"] > 0,
                budget_series / agg["conversions"],
                0.0,
            )
            agg["roas"] = np.where(
                budget_series > 0,
                agg["revenue"] / budget_series,
                0.0,
            )
            agg["drr"] = np.where(
                agg["revenue"] > 0,
                budget_series / agg["revenue"] * 100,
                0.0,
            )

            total_row = {
                "month_num": 999,
                "month_name": "TOTAL",
                "impressions": agg["impressions"].sum(),
                "clicks": agg["clicks"].sum(),
                "conversions": agg["conversions"].sum(),
                "cost": agg["cost"].sum(),
                "cost_with_vat": agg["cost_with_vat"].sum(),
                "cost_with_vat_ak": agg["cost_with_vat_ak"].sum(),
                "ak_cost_wo_vat": agg["ak_cost_wo_vat"].sum(),
                "revenue": agg["revenue"].sum(),
            }
            if use_ak_budget_metrics:
                total_budget_basis = total_row["cost_with_vat_ak"] if use_vat_budget_metrics else (total_row["cost"] + total_row["ak_cost_wo_vat"])
            else:
                total_budget_basis = total_row["cost_with_vat"] if use_vat_budget_metrics else total_row["cost"]
            if total_row["impressions"] > 0:
                total_row["ctr"] = total_row["clicks"] / total_row["impressions"] * 100
                total_row["cpm"] = total_budget_basis / (total_row["impressions"] / 1000)
            else:
                total_row["ctr"] = 0.0
                total_row["cpm"] = 0.0
            if total_row["clicks"] > 0:
                total_row["cr"] = total_row["conversions"] / total_row["clicks"] * 100
                total_row["cpc"] = total_row["cost"] / total_row["clicks"]
            else:
                total_row["cr"] = 0.0
                total_row["cpc"] = 0.0
            if total_row["conversions"] > 0:
                total_row["cpa"] = total_budget_basis / total_row["conversions"]
            else:
                total_row["cpa"] = 0.0
            if total_budget_basis > 0:
                total_row["roas"] = total_row["revenue"] / total_budget_basis
            else:
                total_row["roas"] = 0.0
            if total_row["revenue"] > 0:
                total_row["drr"] = total_budget_basis / total_row["revenue"] * 100
            else:
                total_row["drr"] = 0.0

            segment_total_rows = []
            if show_segment_subtotals and "segment" in df_all.columns:
                for seg_name, seg_df in df_all.groupby("segment"):
                    seg_cost = float(seg_df["cost"].sum())
                    seg_cost_with_vat = float(seg_df["cost_with_vat"].sum())
                    seg_cost_with_vat_ak = float(seg_df["cost_with_vat_ak"].sum())
                    seg_ak_wo_vat = float(seg_df["ak_cost_wo_vat"].sum())
                    if use_ak_budget_metrics:
                        seg_budget_basis = seg_cost_with_vat_ak if use_vat_budget_metrics else (seg_cost + seg_ak_wo_vat)
                    else:
                        seg_budget_basis = seg_cost_with_vat if use_vat_budget_metrics else seg_cost
                    seg_row = {
                        "month_num": 998,
                        "month_name": f"того {seg_name}",
                        "impressions": float(seg_df["impressions"].sum()),
                        "clicks": float(seg_df["clicks"].sum()),
                        "conversions": float(seg_df["conversions"].sum()),
                        "cost": seg_cost,
                        "cost_with_vat": seg_cost_with_vat,
                        "cost_with_vat_ak": seg_cost_with_vat_ak,
                        "ak_cost_wo_vat": seg_ak_wo_vat,
                        "revenue": float(seg_df["revenue"].sum()),
                    }
                    seg_row["ctr"] = (seg_row["clicks"] / seg_row["impressions"] * 100.0) if seg_row["impressions"] > 0 else 0.0
                    seg_row["cr"] = (seg_row["conversions"] / seg_row["clicks"] * 100.0) if seg_row["clicks"] > 0 else 0.0
                    seg_row["cpc"] = (seg_row["cost"] / seg_row["clicks"]) if seg_row["clicks"] > 0 else 0.0
                    seg_row["cpm"] = (seg_budget_basis / (seg_row["impressions"] / 1000.0)) if seg_row["impressions"] > 0 else 0.0
                    seg_row["cpa"] = (seg_budget_basis / seg_row["conversions"]) if seg_row["conversions"] > 0 else 0.0
                    seg_row["roas"] = (seg_row["revenue"] / seg_budget_basis) if seg_budget_basis > 0 else 0.0
                    seg_row["drr"] = (seg_budget_basis / seg_row["revenue"] * 100.0) if seg_row["revenue"] > 0 else 0.0
                    segment_total_rows.append(seg_row)

            agg = pd.concat([agg, pd.DataFrame(segment_total_rows + [total_row])], ignore_index=True)

            agg_cols = [
                "month_name",
                "impressions",
                "clicks",
                "conversions",
            ]
            agg_cols += [
                "cost",
                "cost_with_vat",
                "cost_with_vat_ak",
                "revenue",
                "ctr",
                "cpc",
                "cr",
                "cpm",
                "cpa",
                "roas",
                "drr",
            ]
            agg_show = agg[agg_cols]
            agg_show = agg_show.rename(columns={"roas": "ROAS"})
            agg_show = agg_show.rename(columns=DISPLAY_COL_RENAME)

            numeric_cols = [
                "impressions", "clicks", "conversions", "cost", "cost_with_vat", "cost_with_vat_ak", "revenue",
                "ctr", "cpc", "cr", "cpm", "cpa", "ROAS", "drr"
            ]
            numeric_cols = [DISPLAY_COL_RENAME.get(c, c) for c in numeric_cols]

            def _blend_hex(c1: str, c2: str, t: float) -> str:
                t = max(0.0, min(1.0, t))
                r1, g1, b1 = int(c1[1:3], 16), int(c1[3:5], 16), int(c1[5:7], 16)
                r2, g2, b2 = int(c2[1:3], 16), int(c2[3:5], 16), int(c2[5:7], 16)
                r = int(r1 + (r2 - r1) * t)
                g = int(g1 + (g2 - g1) * t)
                b = int(b1 + (b2 - b1) * t)
                return f"#{r:02x}{g:02x}{b:02x}"

            def _style_metric_col(col: pd.Series) -> pd.Series:
                styles = pd.Series([""] * len(col), index=col.index)
                mask = ~agg_show["Месяц"].astype(str).str.startswith("того")
                mask &= agg_show["Месяц"] != "TOTAL"
                vals = pd.to_numeric(col[mask], errors="coerce").dropna()
                if vals.empty:
                    return styles
                vmin, vmax = float(vals.min()), float(vals.max())
                span = (vmax - vmin) if vmax != vmin else 1.0
                base_low = "#173058"
                base_high = "#19B8B2"
                for idx in vals.index:
                    t = (float(col.loc[idx]) - vmin) / span
                    color = _blend_hex(base_low, base_high, t)
                    styles.loc[idx] = (
                        f"background-color: {color}; color: #EAF3EE; border: 1px solid {THEME_BORDER};"
                    )
                return styles

            def _highlight_total_row_total(row):
                style = [""] * len(row)
                month_val = str(row["Месяц"])
                if month_val == "TOTAL":
                    style = [
                        f"background-color: #00CDC5; color: #081521; font-weight: 700; border-top: 2px solid {THEME_BORDER};"
                    ] * len(row)
                elif month_val.startswith("того "):
                    style = [
                        f"background-color: #2C6E75; color: #DDEAF0; font-weight: 650; border-top: 1px solid {THEME_BORDER};"
                    ] * len(row)
                return style

            total_formatters = {
                DISPLAY_COL_RENAME["impressions"]: lambda x: "" if pd.isna(x) else f"{round(x):,}".replace(",", " "),
                DISPLAY_COL_RENAME["clicks"]: lambda x: "" if pd.isna(x) else f"{round(x):,}".replace(",", " "),
                DISPLAY_COL_RENAME["conversions"]: lambda x: "" if pd.isna(x) else f"{round(x):,}".replace(",", " "),
                DISPLAY_COL_RENAME["cost"]: lambda x: "" if pd.isna(x) else f"{round(x):,} \u20BD".replace(",", " "),
                DISPLAY_COL_RENAME["cost_with_vat"]: lambda x: "" if pd.isna(x) else f"{round(x):,} \u20BD".replace(",", " "),
                DISPLAY_COL_RENAME["cost_with_vat_ak"]: lambda x: "" if pd.isna(x) else f"{round(x):,} \u20BD".replace(",", " "),
                DISPLAY_COL_RENAME["revenue"]: lambda x: "" if pd.isna(x) else f"{round(x):,} \u20BD".replace(",", " "),
                DISPLAY_COL_RENAME["cpm"]: lambda x: "" if pd.isna(x) else f"{round(x):,} \u20BD".replace(",", " "),
                DISPLAY_COL_RENAME["cpa"]: lambda x: "" if pd.isna(x) else f"{round(x):,} \u20BD".replace(",", " "),
                DISPLAY_COL_RENAME["cpc"]: lambda x: "" if pd.isna(x) else f"{x:.2f} \u20BD",
                DISPLAY_COL_RENAME["ctr"]: lambda x: "" if pd.isna(x) else f"{x:.2f} %",
                DISPLAY_COL_RENAME["cr"]: lambda x: "" if pd.isna(x) else f"{x:.2f} %",
                DISPLAY_COL_RENAME["drr"]: lambda x: "" if pd.isna(x) else f"{x:.2f} %",
                "ROAS": lambda x: "" if pd.isna(x) else f"{x * 100:.2f} %",
            }
            styled_total = (
                agg_show.style
                .format(total_formatters)
                .apply(_highlight_total_row_total, axis=1)
            )
            for c in numeric_cols:
                styled_total = styled_total.apply(_style_metric_col, axis=0, subset=[c])

            st.dataframe(styled_total, use_container_width=True)

    render_bottom_tab_switcher("Медиаплан", "plan")
# ====================================================================
#                           ТАБ «ДАГРАММЫ»
# ====================================================================

with tab_charts:
    ui_section_title("Сводная таблица и диаграммы")
    use_vat_budget_metrics = st.session_state.get("use_vat_budget_metrics", True)
    use_ak_budget_metrics = st.session_state.get("use_ak_budget_metrics", False)
    st.markdown(
        """
        <div style="
            margin: 0.15rem 0 0.55rem 0;
            padding: 10px 12px;
            border-radius: 10px;
            border: 1px solid #2B7EE8;
            background: rgba(10, 116, 242, 0.12);
        ">
            <span style="font-weight: 800; color: #9EC5FF;">Назначение блока:</span>
            <span style="color: #EAF0FF; font-weight: 560;">
                Вкладка используется для финальной проверки расчетов медиаплана: через фильтры и визуализации можно валидировать корректность формул, распределений и итоговых значений.
            </span>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        """
        <div class="tab-intro">
            <p>1) Сначала выберите период (месяцы), который хотите проверить, или оставьте все настройки и проверьте весь период расчета медиаплана целиком.</p>
            <p>2) Примените фильтры по нужным разрезам (тип кампании, система, формат), чтобы сузить данные.</p>
            <p>3) Проверьте сводную таблицу по месяцам: она показывает итоговые метрики и помогает быстро увидеть отклонения.</p>
            <p>4) Сравните динамику ключевых показателей (показы, клики, расход, конверсии, выручка, ROAS/ДРР) между месяцами.</p>
            <p>5) При необходимости включите круговые диаграммы, чтобы проверить структуру распределения бюджета, конверсий и выручки по типам РК.</p>
            <p>6) Если цифры выглядят нелогично, вернитесь во вкладку «Медиаплан» и внесите точечные корректировки в нужные месяцы, затем вернитесь в «Диаграммы» и проведите повторную проверку.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if "df_all" not in locals() or df_all.empty:
        st.info("Сначала заполните данные на вкладке «Медиаплан» и рассчитайте месяцы.")
    else:
        st.caption("Выберите месяцы и типы РК для сводной таблицы и круговых диаграмм.")
        st.caption(
            "Режим бюджетных метрик: "
            + (
                "с НДС 22% и АК"
                if use_ak_budget_metrics and use_vat_budget_metrics
                else "с НДС 22%"
                if use_vat_budget_metrics
                else "без НДС"
            )
        )

        month_names_full = {
            1: "Январь",
            2: "Февраль",
            3: "Март",
            4: "Апрель",
            5: "Май",
            6: "юнь",
            7: "юль",
            8: "Август",
            9: "Сентябрь",
            10: "Октябрь",
            11: "Ноябрь",
            12: "Декабрь",
        }

        available_months = sorted(df_all["month_num"].unique())
        month_options = [f"{m}. {month_names_full[m]}" for m in available_months]

        selected_labels = st.multiselect(
            "Месяцы в сводной таблице и диаграммах:",
            options=month_options,
            default=month_options,
            key="charts_months_multiselect",
        )

        if not selected_labels:
            st.info("Выберите хотя бы один месяц.")
        else:
            selected_month_nums_charts = [int(label.split(".")[0]) for label in selected_labels]

            all_campaign_types = df_all["campaign_type"].unique().tolist()
            selected_campaign_types = st.multiselect(
                "Типы РК для включения:",
                options=all_campaign_types,
                default=all_campaign_types,
                key="charts_campaign_types_multiselect",
            )

            if not selected_campaign_types:
                st.info("Выберите хотя бы один тип РК.")
            else:
                selected_segments_charts = None
                if is_diy_preset and "segment" in df_all.columns:
                    all_segments = sorted(df_all["segment"].dropna().astype(str).unique().tolist())
                    selected_segments_charts = st.multiselect(
                        "Сегменты для включения:",
                        options=all_segments,
                        default=all_segments,
                        key="charts_segments_multiselect",
                    )
                    if not selected_segments_charts:
                        st.info("Выберите хотя бы один сегмент.")
                        st.stop()

                mask = df_all["month_num"].isin(selected_month_nums_charts) & df_all[
                    "campaign_type"
                ].isin(selected_campaign_types)
                if selected_segments_charts is not None:
                    mask &= df_all["segment"].isin(selected_segments_charts)
                df_sel = df_all[mask]

                if df_sel.empty:
                    st.info("Для выбранных месяцев и типов РК нет данных.")
                else:
                    # ---------- Сводная таблица по месяцам ----------
                    agg_m = df_sel.groupby(["month_num", "month_name"], as_index=False).agg(
                        impressions=("impressions", "sum"),
                        clicks=("clicks", "sum"),
                        conversions=("conversions", "sum"),
                        cost=("cost", "sum"),
                        cost_with_vat=("cost_with_vat", "sum"),
                        cost_with_vat_ak=("cost_with_vat_ak", "sum"),
                        ak_cost_wo_vat=("ak_cost_wo_vat", "sum"),
                        revenue=("revenue", "sum"),
                    )
                    if use_ak_budget_metrics:
                        budget_series = agg_m["cost_with_vat_ak"] if use_vat_budget_metrics else (agg_m["cost"] + agg_m["ak_cost_wo_vat"])
                    else:
                        budget_series = agg_m["cost_with_vat"] if use_vat_budget_metrics else agg_m["cost"]

                    agg_m["ctr"] = np.where(
                        agg_m["impressions"] > 0,
                        agg_m["clicks"] / agg_m["impressions"] * 100,
                        0.0,
                    )
                    agg_m["cr"] = np.where(
                        agg_m["clicks"] > 0,
                        agg_m["conversions"] / agg_m["clicks"] * 100,
                        0.0,
                    )
                    agg_m["cpm"] = np.where(
                        agg_m["impressions"] > 0,
                        budget_series / (agg_m["impressions"] / 1000),
                        0.0,
                    )
                    agg_m["cpa"] = np.where(
                        agg_m["conversions"] > 0,
                        budget_series / agg_m["conversions"],
                        0.0,
                    )
                    agg_m["aov"] = np.where(
                        agg_m["conversions"] > 0,
                        agg_m["revenue"] / agg_m["conversions"],
                        0.0,
                    )
                    agg_m["cpc"] = np.where(
                        agg_m["clicks"] > 0,
                        agg_m["cost"] / agg_m["clicks"],
                        0.0,
                    )
                    agg_m["roas"] = np.where(
                        budget_series > 0,
                        agg_m["revenue"] / budget_series,
                        0.0,
                    )
                    agg_m["drr"] = np.where(
                        agg_m["revenue"] > 0,
                        budget_series / agg_m["revenue"] * 100,
                        0.0,
                    )

                    selected_month_nums_charts = sorted(selected_month_nums_charts)
                    month_headers = [month_names_full[m] for m in selected_month_nums_charts]

                    # Коэффициенты, использованные в расчетах (в среднем по выбранным типам РК)
                    if "campaigns" in locals() and not campaigns.empty:
                        base_for_coeffs = campaigns[
                            campaigns["campaign_type"].isin(selected_campaign_types)
                        ][
                            [
                                "campaign_type",
                                "impressions_avg",
                                "ctr_avg_percent",
                                "cpc_avg",
                                "cr_avg_percent",
                                "aov_avg",
                            ]
                        ].copy()
                        if not base_for_coeffs.empty:
                            coeff_raw = df_sel.merge(base_for_coeffs, on="campaign_type", how="left")
                            coeff_raw["base_ctr"] = coeff_raw["ctr_avg_percent"] / 100.0
                            coeff_raw["base_cr"] = coeff_raw["cr_avg_percent"] / 100.0

                            coeff_raw["k_imp"] = np.where(
                                coeff_raw["impressions_avg"] > 0,
                                coeff_raw["impressions"] / coeff_raw["impressions_avg"],
                                np.nan,
                            )
                            coeff_raw["k_ctr"] = np.where(
                                coeff_raw["base_ctr"] > 0,
                                coeff_raw["ctr"] / coeff_raw["base_ctr"],
                                np.nan,
                            )
                            coeff_raw["k_cpc"] = np.where(
                                coeff_raw["cpc_avg"] > 0,
                                coeff_raw["cpc"] / coeff_raw["cpc_avg"],
                                np.nan,
                            )
                            coeff_raw["k_cr"] = np.where(
                                coeff_raw["base_cr"] > 0,
                                coeff_raw["cr"] / coeff_raw["base_cr"],
                                np.nan,
                            )
                            coeff_raw["k_aov"] = np.where(
                                coeff_raw["aov_avg"] > 0,
                                coeff_raw["aov"] / coeff_raw["aov_avg"],
                                np.nan,
                            )

                            coeff_month = coeff_raw.groupby("month_num", as_index=False).agg(
                                k_imp=("k_imp", "mean"),
                                k_ctr=("k_ctr", "mean"),
                                k_cpc=("k_cpc", "mean"),
                                k_cr=("k_cr", "mean"),
                                k_aov=("k_aov", "mean"),
                            )
                            coeff_month_map = coeff_month.set_index("month_num").to_dict("index")

                            # Отдельно показываем коэффициент "медийных хвостов" как значение из набора
                            # (без усреднения по типам РК)
                            media_tail_month_map = {}
                            media_tail_values_by_set = {}
                            coeff_sets_for_tail = st.session_state.get("coeff_sets", [])
                            for cs_tail in coeff_sets_for_tail:
                                if normalize_coeff_set_type(cs_tail.get("type")) != "Медийные хвосты":
                                    continue
                                cs_tail_name = str(cs_tail.get("name", "")).strip()
                                df_tail = cs_tail.get("result")
                                if not cs_tail_name or df_tail is None or getattr(df_tail, "empty", True):
                                    continue
                                if "Номер месяца" not in df_tail.columns or "Коэф." not in df_tail.columns:
                                    continue
                                mm = {}
                                for _, rtail in df_tail.iterrows():
                                    try:
                                        mm[int(rtail["Номер месяца"])] = float(rtail["Коэф."])
                                    except Exception:
                                        continue
                                media_tail_values_by_set[cs_tail_name] = mm

                            coeff_links_for_tail = st.session_state.get("coeff_sets_links_new", pd.DataFrame())
                            coeff_links_for_tail = coeff_links_for_tail[
                                coeff_links_for_tail["campaign_type"].isin(selected_campaign_types)
                            ] if not coeff_links_for_tail.empty else pd.DataFrame()

                            for m in selected_month_nums_charts:
                                set_names = []
                                if not coeff_links_for_tail.empty and "media_tail_set" in coeff_links_for_tail.columns:
                                    for _, rr in coeff_links_for_tail.iterrows():
                                        set_name = str(rr.get("media_tail_set", "")).strip()
                                        if set_name:
                                            set_names.append(set_name)

                                uniq_set_names = sorted(set(set_names))
                                if len(uniq_set_names) == 1:
                                    media_tail_month_map[m] = float(
                                        media_tail_values_by_set.get(uniq_set_names[0], {}).get(m, 1.0)
                                    )
                                elif len(uniq_set_names) == 0:
                                    media_tail_month_map[m] = 1.0
                                else:
                                    media_tail_month_map[m] = None

                            coeff_rows = []
                            coeff_labels = [
                                ("k_imp", "Коэф. показы"),
                                ("k_ctr", "Коэф. CTR"),
                                ("k_cpc", "Коэф. CPC"),
                                ("k_cr", "Коэф. CR"),
                                ("k_aov", "Коэф. AOV"),
                            ]
                            for key_name, label in coeff_labels:
                                row = {"Метрика": label}
                                for m in selected_month_nums_charts:
                                    val = coeff_month_map.get(m, {}).get(key_name, np.nan)
                                    row[month_names_full[m]] = "" if pd.isna(val) else f"{val:.2f}"
                                coeff_rows.append(row)

                            row_media_tail = {"Метрика": "Коэф. медийные хвосты"}
                            for m in selected_month_nums_charts:
                                v = media_tail_month_map.get(m, 1.0)
                                row_media_tail[month_names_full[m]] = "неск." if v is None else f"{v:.2f}"
                            coeff_rows.append(row_media_tail)

                            coeff_table = pd.DataFrame(coeff_rows).set_index("Метрика")
                            with st.expander("Коэффициенты, использованные в расчете по месяцам", expanded=False):
                                st.dataframe(coeff_table, use_container_width=True)
                            st.caption("Коэф. медийные хвосты: 1.00 = хвосты не применяются.")

                    # Транспонированная таблица: месяцы в колонках, метрики в строках
                    agg_by_month = agg_m.set_index("month_num").to_dict("index")

                    total_impressions = agg_m["impressions"].sum()
                    total_clicks = agg_m["clicks"].sum()
                    total_conversions = agg_m["conversions"].sum()
                    total_cost = agg_m["cost"].sum()
                    total_cost_with_vat = agg_m["cost_with_vat"].sum()
                    total_cost_with_vat_ak = agg_m["cost_with_vat_ak"].sum()
                    total_ak_wo_vat = agg_m["ak_cost_wo_vat"].sum()
                    total_revenue = agg_m["revenue"].sum()
                    if use_ak_budget_metrics:
                        total_budget_basis = total_cost_with_vat_ak if use_vat_budget_metrics else (total_cost + total_ak_wo_vat)
                    else:
                        total_budget_basis = total_cost_with_vat if use_vat_budget_metrics else total_cost
                    total_metrics = {
                        "impressions": total_impressions,
                        "clicks": total_clicks,
                        "ctr": (total_clicks / total_impressions * 100) if total_impressions > 0 else 0.0,
                        "cpc": (total_cost / total_clicks) if total_clicks > 0 else 0.0,
                        "cost": total_cost,
                        "cost_with_vat": total_cost_with_vat,
                        "cost_with_vat_ak": total_cost_with_vat_ak,
                        "cr": (total_conversions / total_clicks * 100) if total_clicks > 0 else 0.0,
                        "aov": (total_revenue / total_conversions) if total_conversions > 0 else 0.0,
                        "conversions": total_conversions,
                        "cpa": (total_budget_basis / total_conversions) if total_conversions > 0 else 0.0,
                        "revenue": total_revenue,
                        "roas": (total_revenue / total_budget_basis) if total_budget_basis > 0 else 0.0,
                        "drr": (total_budget_basis / total_revenue * 100) if total_revenue > 0 else 0.0,
                    }

                    def _fmt_int(v):
                        return f"{round(v):,}".replace(",", " ")

                    def _fmt_rub(v):
                        return f"{round(v):,} ₽".replace(",", " ")

                    def _fmt_rub2(v):
                        return f"{v:.2f} ₽"

                    def _fmt_pct(v):
                        return f"{v:.2f} %"

                    def _fmt_mult(v):
                        return f"{v * 100:.2f} %"

                    metric_specs = [
                        ("Показы", "impressions"),
                        ("Клики", "clicks"),
                        ("CTR", "ctr"),
                        ("CPC", "cpc"),
                        ("Расход", "cost"),
                        ("Бюджет с НДС", "cost_with_vat"),
                        ("Бюджет с НДС и АК", "cost_with_vat_ak"),
                        ("CR", "cr"),
                        ("AOV", "aov"),
                        ("Конверсии", "conversions"),
                        ("CPO", "cpa"),
                        ("Выручка", "revenue"),
                        ("ROAS", "roas"),
                        ("ДРР", "drr"),
                    ]

                    rows = []
                    for metric_label, metric_key in metric_specs:
                        row = {"Метрика": metric_label}
                        for m in selected_month_nums_charts:
                            month_col = month_names_full[m]
                            val = agg_by_month.get(m, {}).get(metric_key, 0.0)
                            if metric_key in ["impressions", "clicks", "conversions"]:
                                row[month_col] = _fmt_int(val)
                            elif metric_key in ["cost", "cost_with_vat", "cost_with_vat_ak", "revenue", "cpa", "aov"]:
                                row[month_col] = _fmt_rub(val)
                            elif metric_key == "cpc":
                                row[month_col] = _fmt_rub2(val)
                            elif metric_key in ["ctr", "cr", "drr"]:
                                row[month_col] = _fmt_pct(val)
                            elif metric_key == "roas":
                                row[month_col] = _fmt_mult(val)
                            else:
                                row[month_col] = val

                        total_val = total_metrics.get(metric_key, 0.0)
                        if metric_key in ["impressions", "clicks", "conversions"]:
                            row["TOTAL"] = _fmt_int(total_val)
                        elif metric_key in ["cost", "cost_with_vat", "cost_with_vat_ak", "revenue", "cpa", "aov"]:
                            row["TOTAL"] = _fmt_rub(total_val)
                        elif metric_key == "cpc":
                            row["TOTAL"] = _fmt_rub2(total_val)
                        elif metric_key in ["ctr", "cr", "drr"]:
                            row["TOTAL"] = _fmt_pct(total_val)
                        elif metric_key == "roas":
                            row["TOTAL"] = _fmt_mult(total_val)
                        else:
                            row["TOTAL"] = _fmt_int(total_val)
                        rows.append(row)

                    summary_matrix = pd.DataFrame(rows).set_index("Метрика")
                    ui_section_title("Сводная таблица (месяцы в колонках)")
                    heat_cols = month_headers

                    def _parse_num_for_heat(v):
                        if pd.isna(v):
                            return None
                        if isinstance(v, (int, float, np.integer, np.floating)):
                            return float(v)
                        s = str(v).strip()
                        if not s:
                            return None
                        for token in ["₽", "%", "×"]:
                            s = s.replace(token, "")
                        s = s.replace(" ", "").replace(",", ".")
                        try:
                            return float(s)
                        except Exception:
                            return None

                    def _blend_hex(c1: str, c2: str, t: float) -> str:
                        t = max(0.0, min(1.0, t))
                        r1, g1, b1 = int(c1[1:3], 16), int(c1[3:5], 16), int(c1[5:7], 16)
                        r2, g2, b2 = int(c2[1:3], 16), int(c2[3:5], 16), int(c2[5:7], 16)
                        r = int(r1 + (r2 - r1) * t)
                        g = int(g1 + (g2 - g1) * t)
                        b = int(b1 + (b2 - b1) * t)
                        return f"#{r:02x}{g:02x}{b:02x}"

                    base_low = "#173058"
                    base_high = "#19B8B2"

                    def _row_heat_style(row):
                        styles = [""] * len(row)
                        vals = [_parse_num_for_heat(row.get(col)) for col in heat_cols]
                        nums = [v for v in vals if v is not None]
                        if not nums:
                            return styles

                        vmin, vmax = min(nums), max(nums)
                        span = (vmax - vmin) if vmax != vmin else 1.0

                        for i, col in enumerate(row.index):
                            if col not in heat_cols:
                                continue
                            value = _parse_num_for_heat(row[col])
                            if value is None:
                                continue
                            t = (value - vmin) / span
                            color = _blend_hex(base_low, base_high, t)
                            styles[i] = (
                                f"background-color: {color}; color: #EAF3EE; "
                                f"border: 1px solid {THEME_BORDER};"
                            )
                        return styles

                    legend_c1, legend_c2 = st.columns([4.5, 1.5], vertical_alignment="center")
                    with legend_c2:
                        st.markdown(
                            f"""
                            <div style="border:1px solid {THEME_BORDER}; border-radius:10px; padding:8px 10px; background:{THEME_CARD_BG};">
                                <div style="font-size:12px; color:{THEME_LEGEND_TEXT}; margin-bottom:5px;">Heatmap: min -> max</div>
                                <div style="height:10px; border-radius:999px; border:1px solid {THEME_BORDER}; background: linear-gradient(90deg, {base_low}, {base_high});"></div>
                                <div style="display:flex; justify-content:space-between; font-size:11px; color:{THEME_LEGEND_TEXT}; margin-top:4px;">
                                    <span>min</span><span>max</span>
                                </div>
                            </div>
                            """,
                            unsafe_allow_html=True,
                        )
                    with legend_c1:
                        st.caption("Шкала применяется отдельно в каждой строке метрик.")

                    styled_summary = summary_matrix.style.apply(_row_heat_style, axis=1)
                    st.dataframe(styled_summary, use_container_width=True)

                    # ---------- Круговые диаграммы по типам РК ----------

                    ui_section_title("Распределение по типам РК")

                    agg_ct = df_sel.groupby("campaign_type", as_index=False).agg(
                        cost=("cost", "sum"),
                        conversions=("conversions", "sum"),
                        revenue=("revenue", "sum"),
                    )
                    show_pies = st.checkbox(
                        "Показывать круговые диаграммы",
                        value=True,
                        key="charts_show_pies",
                    )

                    if show_pies:
                        pie_palette = [
                            "#0066E0", "#00CDC5", "#9747FF", "#FF6333",
                            "#3D8EF0", "#42DDD6", "#B07BFF", "#FF8A66",
                        ]
                        st.caption("Общая легенда")
                        legend_cols_count = min(4, max(1, len(agg_ct)))
                        legend_cols = st.columns(legend_cols_count)
                        for i, ct_name in enumerate(agg_ct["campaign_type"].tolist()):
                            c = pie_palette[i % len(pie_palette)]
                            with legend_cols[i % legend_cols_count]:
                                st.markdown(
                                    f"<span style='color:{c}; font-weight:700;'>●</span> "
                                    f"<span style='color:{THEME_LEGEND_TEXT};'>{ct_name}</span>",
                                    unsafe_allow_html=True,
                                )

                        col_pie1, col_pie2, col_pie3 = st.columns(3)
                        with col_pie1:
                            fig_cost = px.pie(
                                agg_ct,
                                names="campaign_type",
                                values="cost",
                                title="Доля бюджета по типам РК",
                                color_discrete_sequence=pie_palette,
                            )
                            fig_cost.update_traces(textposition="inside", textinfo="percent+label")
                            fig_cost.update_layout(
                                height=520,
                                margin=dict(l=12, r=12, t=64, b=12),
                                showlegend=False,
                                paper_bgcolor="rgba(0,0,0,0)",
                                plot_bgcolor="rgba(0,0,0,0)",
                                font=dict(color=THEME_PLOT_TEXT),
                            )
                            st.plotly_chart(fig_cost, use_container_width=True)

                        with col_pie2:
                            fig_conv = px.pie(
                                agg_ct,
                                names="campaign_type",
                                values="conversions",
                                title="Доля конверсий по типам РК",
                                color_discrete_sequence=pie_palette,
                            )
                            fig_conv.update_traces(textposition="inside", textinfo="percent+label")
                            fig_conv.update_layout(
                                height=520,
                                margin=dict(l=12, r=12, t=64, b=12),
                                showlegend=False,
                                paper_bgcolor="rgba(0,0,0,0)",
                                plot_bgcolor="rgba(0,0,0,0)",
                                font=dict(color=THEME_PLOT_TEXT),
                            )
                            st.plotly_chart(fig_conv, use_container_width=True)

                        with col_pie3:
                            fig_rev = px.pie(
                                agg_ct,
                                names="campaign_type",
                                values="revenue",
                                title="Доля дохода по типам РК",
                                color_discrete_sequence=pie_palette,
                            )
                            fig_rev.update_traces(textposition="inside", textinfo="percent+label")
                            fig_rev.update_layout(
                                height=520,
                                margin=dict(l=12, r=12, t=64, b=12),
                                showlegend=False,
                                paper_bgcolor="rgba(0,0,0,0)",
                                plot_bgcolor="rgba(0,0,0,0)",
                                font=dict(color=THEME_PLOT_TEXT),
                            )
                            st.plotly_chart(fig_rev, use_container_width=True)

                        if is_diy_preset and "segment" in df_sel.columns:
                            seg_work = df_sel.copy()
                            seg_work["segment_norm"] = (
                                seg_work["segment"].astype(str).str.upper().str.strip()
                            )
                            seg_work = seg_work[seg_work["segment_norm"].isin(["B2B", "B2C"])]

                            if not seg_work.empty:
                                if use_ak_budget_metrics:
                                    if use_vat_budget_metrics:
                                        seg_work["budget_basis"] = seg_work["cost_with_vat_ak"]
                                    else:
                                        seg_work["budget_basis"] = seg_work["cost"] + seg_work["ak_cost_wo_vat"]
                                else:
                                    seg_work["budget_basis"] = (
                                        seg_work["cost_with_vat"] if use_vat_budget_metrics else seg_work["cost"]
                                    )

                                agg_seg = (
                                    seg_work.groupby("segment_norm", as_index=False)
                                    .agg(
                                        budget_basis=("budget_basis", "sum"),
                                        revenue=("revenue", "sum"),
                                        orders=("conversions", "sum"),
                                    )
                                )
                                agg_seg["segment_norm"] = pd.Categorical(
                                    agg_seg["segment_norm"], categories=["B2C", "B2B"], ordered=True
                                )
                                agg_seg = agg_seg.sort_values("segment_norm")

                                ui_section_title("DIY: распределение между B2B и B2C")
                                seg_palette = ["#3D8EF0", "#00CDC5"]
                                seg_pie1, seg_pie2, seg_pie3 = st.columns(3)

                                with seg_pie1:
                                    fig_seg_budget = px.pie(
                                        agg_seg,
                                        names="segment_norm",
                                        values="budget_basis",
                                        title="Доля бюджета между B2B и B2C",
                                        color_discrete_sequence=seg_palette,
                                    )
                                    fig_seg_budget.update_traces(textposition="inside", textinfo="percent+label")
                                    fig_seg_budget.update_layout(
                                        height=520,
                                        margin=dict(l=12, r=12, t=64, b=12),
                                        showlegend=False,
                                        paper_bgcolor="rgba(0,0,0,0)",
                                        plot_bgcolor="rgba(0,0,0,0)",
                                        font=dict(color=THEME_PLOT_TEXT),
                                    )
                                    st.plotly_chart(fig_seg_budget, use_container_width=True)

                                with seg_pie2:
                                    fig_seg_rev = px.pie(
                                        agg_seg,
                                        names="segment_norm",
                                        values="revenue",
                                        title="Доля выручки между B2B и B2C",
                                        color_discrete_sequence=seg_palette,
                                    )
                                    fig_seg_rev.update_traces(textposition="inside", textinfo="percent+label")
                                    fig_seg_rev.update_layout(
                                        height=520,
                                        margin=dict(l=12, r=12, t=64, b=12),
                                        showlegend=False,
                                        paper_bgcolor="rgba(0,0,0,0)",
                                        plot_bgcolor="rgba(0,0,0,0)",
                                        font=dict(color=THEME_PLOT_TEXT),
                                    )
                                    st.plotly_chart(fig_seg_rev, use_container_width=True)

                                with seg_pie3:
                                    fig_seg_orders = px.pie(
                                        agg_seg,
                                        names="segment_norm",
                                        values="orders",
                                        title="Доля заказов между B2B и B2C",
                                        color_discrete_sequence=seg_palette,
                                    )
                                    fig_seg_orders.update_traces(textposition="inside", textinfo="percent+label")
                                    fig_seg_orders.update_layout(
                                        height=520,
                                        margin=dict(l=12, r=12, t=64, b=12),
                                        showlegend=False,
                                        paper_bgcolor="rgba(0,0,0,0)",
                                        plot_bgcolor="rgba(0,0,0,0)",
                                        font=dict(color=THEME_PLOT_TEXT),
                                    )
                                    st.plotly_chart(fig_seg_orders, use_container_width=True)

    render_bottom_tab_switcher("Диаграммы", "charts")

# ====================================================================
#                           ТАБ «EXPORT»
# ====================================================================

with tab_export:
    ui_section_title("Экспорт")
    st.markdown(
        """
        <div class="tab-intro">
            <p>Вкладка подготовки и запуска экспорта медиаплана в Excel.</p>
            <p>Доступны: выгрузка детальных данных, сводки по месяцам и экспорт в шаблон.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown("### Экспорт проекта")
    st.caption("Сохраните текущее состояние проекта в JSON, чтобы передать его другому пользователю.")

    project_payload = export_project_state_payload()
    project_json = json.dumps(project_payload, ensure_ascii=False, indent=2).encode("utf-8")
    project_ts = dt.datetime.now().strftime("%Y%m%d_%H%M")
    st.download_button(
        "Скачать проект (JSON)",
        data=project_json,
        file_name=f"mediaplan_project_{project_ts}.json",
        mime="application/json",
        key="download_project_json",
    )

    st.markdown("---")
    if "df_all" not in locals() or df_all.empty:
        st.info("Нет данных для экспорта. Сначала рассчитайте медиаплан на вкладке «Медиаплан».")
    else:
        month_names_full = {
            1: "Январь",
            2: "Февраль",
            3: "Март",
            4: "Апрель",
            5: "Май",
            6: "юнь",
            7: "юль",
            8: "Август",
            9: "Сентябрь",
            10: "Октябрь",
            11: "Ноябрь",
            12: "Декабрь",
        }

        export_months_available = sorted(df_all["month_num"].dropna().astype(int).unique().tolist())
        export_month_options = [f"{m}. {month_names_full.get(m, str(m))}" for m in export_months_available]
        export_month_labels = st.multiselect(
            "Месяцы для экспорта",
            options=export_month_options,
            default=export_month_options,
            key="export_months_multiselect",
        )

        export_ctypes_all = sorted(df_all["campaign_type"].dropna().astype(str).unique().tolist())
        export_ctypes_selected = st.multiselect(
            "Типы РК для экспорта",
            options=export_ctypes_all,
            default=export_ctypes_all,
            key="export_campaign_types_multiselect",
        )

        if not export_month_labels:
            st.info("Выберите хотя бы один месяц.")
        elif not export_ctypes_selected:
            st.info("Выберите хотя бы один тип РК.")
        else:
            export_month_nums = [int(lbl.split(".")[0]) for lbl in export_month_labels]
            export_mask = df_all["month_num"].isin(export_month_nums) & df_all["campaign_type"].isin(export_ctypes_selected)
            df_export = df_all.loc[export_mask].copy()

            if df_export.empty:
                st.info("Нет данных для выбранных фильтров экспорта.")
            else:
                st.caption(f"Строк к экспорту: {len(df_export)}")

                # ---------- 1) Детальный Excel ----------
                timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M")
                detail_buf = io.BytesIO()
                detail_engine = "xlsxwriter" if HAS_XLSXWRITER else "openpyxl"
                with pd.ExcelWriter(detail_buf, engine=detail_engine) as writer:
                    detail_cols = [
                        "month_num", "month_name", "campaign_type", "system", "format",
                        "impressions", "clicks", "ctr", "cpc", "cost", "cost_with_vat",
                        "ak_rate", "ak_rate_pct", "ak_cost_wo_vat", "cost_with_vat_ak",
                        "cr", "conversions", "aov", "cpm", "cpa", "revenue", "roas", "drr",
                    ]
                    for c in detail_cols:
                        if c not in df_export.columns:
                            df_export[c] = np.nan
                    df_detail = df_export[detail_cols].rename(
                        columns={
                            "month_num": "Месяц №",
                            "month_name": "Месяц",
                            "campaign_type": "Тип РК",
                            "system": "Система",
                            "format": "Формат",
                            "impressions": "Показы",
                            "clicks": "Клики",
                            "ctr": "CTR (доля)",
                            "cpc": "CPC, ₽",
                            "cost": "Бюджет без НДС, ₽",
                            "cost_with_vat": "Бюджет с НДС, ₽",
                            "ak_rate": "АК (доля)",
                            "ak_rate_pct": "АК, %",
                            "ak_cost_wo_vat": "АК без НДС, ₽",
                            "cost_with_vat_ak": "Бюджет с НДС и АК, ₽",
                            "cr": "CR (доля)",
                            "conversions": "Конверсии",
                            "aov": "AOV, ₽",
                            "cpm": "CPM, ₽",
                            "cpa": "CPO, ₽",
                            "revenue": "Выручка, ₽",
                            "roas": "ROAS (доля)",
                            "drr": "ДРР (доля)",
                        }
                    )
                    df_detail.to_excel(writer, sheet_name="Детально", index=False)

                    agg_month = df_export.groupby(["month_num", "month_name"], as_index=False).agg(
                        impressions=("impressions", "sum"),
                        clicks=("clicks", "sum"),
                        conversions=("conversions", "sum"),
                        cost=("cost", "sum"),
                        cost_with_vat=("cost_with_vat", "sum"),
                        ak_cost_wo_vat=("ak_cost_wo_vat", "sum"),
                        cost_with_vat_ak=("cost_with_vat_ak", "sum"),
                        revenue=("revenue", "sum"),
                    ).sort_values("month_num")
                    agg_month["ctr_pct"] = np.where(agg_month["impressions"] > 0, agg_month["clicks"] / agg_month["impressions"] * 100.0, 0.0)
                    agg_month["cpc"] = np.where(agg_month["clicks"] > 0, agg_month["cost"] / agg_month["clicks"], 0.0)
                    agg_month["cr_pct"] = np.where(agg_month["clicks"] > 0, agg_month["conversions"] / agg_month["clicks"] * 100.0, 0.0)
                    export_use_vat = bool(st.session_state.get("use_vat_budget_metrics", True))
                    export_use_ak = bool(st.session_state.get("use_ak_budget_metrics", False))
                    if export_use_ak:
                        budget_series = agg_month["cost_with_vat_ak"] if export_use_vat else (agg_month["cost"] + agg_month["ak_cost_wo_vat"])
                    else:
                        budget_series = agg_month["cost_with_vat"] if export_use_vat else agg_month["cost"]
                    agg_month["cpm"] = np.where(agg_month["impressions"] > 0, budget_series / (agg_month["impressions"] / 1000.0), 0.0)
                    agg_month["cpo"] = np.where(agg_month["conversions"] > 0, budget_series / agg_month["conversions"], 0.0)
                    agg_month["roas_pct"] = np.where(budget_series > 0, agg_month["revenue"] / budget_series * 100.0, 0.0)
                    agg_month["drr_pct"] = np.where(agg_month["revenue"] > 0, budget_series / agg_month["revenue"] * 100.0, 0.0)
                    agg_month = agg_month.rename(
                        columns={
                            "month_num": "Месяц №",
                            "month_name": "Месяц",
                            "impressions": "Показы",
                            "clicks": "Клики",
                            "conversions": "Конверсии",
                            "cost": "Бюджет без НДС, ₽",
                            "cost_with_vat": "Бюджет с НДС, ₽",
                            "ak_cost_wo_vat": "АК без НДС, ₽",
                            "cost_with_vat_ak": "Бюджет с НДС и АК, ₽",
                            "revenue": "Выручка, ₽",
                            "ctr_pct": "CTR, %",
                            "cpc": "CPC, ₽",
                            "cr_pct": "CR, %",
                            "cpm": "CPM, ₽",
                            "cpo": "CPO, ₽",
                            "roas_pct": "ROAS, %",
                            "drr_pct": "ДРР, %",
                        }
                    )
                    agg_month.to_excel(writer, sheet_name="Сводка_по_месяцам", index=False)

                detail_buf.seek(0)
                st.download_button(
                    "Скачать Excel (детально + сводка)",
                    data=detail_buf.getvalue(),
                    file_name=f"mediaplan_export_{timestamp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_export_detail_xlsx",
                    type="primary",
                )

                # ---------- 2) Экспорт в шаблон ----------
                template_campaigns = campaigns[campaigns["campaign_type"].isin(export_ctypes_selected)].copy()
                template_month_nums = sorted(export_month_nums)
                if len(template_month_nums) > 12:
                    st.warning("Для экспорта в шаблон используется максимум 12 месяцев.")
                    template_month_nums = template_month_nums[:12]
                if len(template_campaigns) > 10:
                    st.warning("Для экспорта в шаблон используется максимум 10 типов РК (ограничение шаблона).")
                    template_campaigns = template_campaigns.head(10)

                if not template_campaigns.empty and template_month_nums:
                    try:
                        export_preset_key = str(st.session_state.get("metric_preset_key", "ecom")).strip().lower()
                        if export_preset_key not in METRIC_PRESETS:
                            export_preset_key = "ecom"
                        allow_ecom_template = export_preset_key == "ecom"
                        allow_diy_template = export_preset_key == "diy"
                        resolved_tpl_ecom = resolve_template_path("ecom")
                        if resolved_tpl_ecom:
                            tpl_buf_ecom = build_excel_from_template(
                                df_all=df_export,
                                campaigns=template_campaigns,
                                selected_month_nums=template_month_nums,
                                template_kind="ecom",
                            )
                            st.download_button(
                                "Скачать Excel по шаблону E-com",
                                data=tpl_buf_ecom.getvalue(),
                                file_name=f"mediaplan_template_ecom_{timestamp}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="download_export_template_xlsx_ecom",
                                disabled=not allow_ecom_template,
                            )
                        else:
                            st.info("Шаблон E-com не найден. Ожидаемые пути: " + ", ".join(TEMPLATE_PATHS_ECOM))

                        resolved_tpl_diy = resolve_template_path("diy")
                        if resolved_tpl_diy:
                            tpl_buf_diy = build_excel_from_template(
                                df_all=df_export,
                                campaigns=template_campaigns,
                                selected_month_nums=template_month_nums,
                                template_kind="diy",
                            )
                            st.download_button(
                                "Скачать Excel по шаблону DIY",
                                data=tpl_buf_diy.getvalue(),
                                file_name=f"mediaplan_template_diy_{timestamp}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="download_export_template_xlsx_diy",
                                disabled=not allow_diy_template,
                            )
                        else:
                            st.info("Шаблон DIY не найден. Добавьте файл в один из путей: " + ", ".join(TEMPLATE_PATHS_DIY))

                    except FileNotFoundError:
                        st.error("Файл шаблона для выбранного типа не найден.")
                    except Exception as e:
                        st.error(f"Ошибка экспорта в шаблон: {e}")
    st.markdown("---")
    st.markdown("### Импорт проекта")
    st.caption("Загрузите JSON-файл проекта, чтобы восстановить расчеты и настройки в этом приложении.")
    import_result = st.session_state.pop("_pending_project_import_result", None)
    if isinstance(import_result, dict):
        if import_result.get("ok"):
            st.success(str(import_result.get("msg", "Данные проекта импортированы.")))
        else:
            st.error(str(import_result.get("msg", "Не удалось импортировать проект.")))

    uploaded_project = st.file_uploader(
        "Импорт проекта (JSON)",
        type=["json"],
        key="upload_project_json",
        help="После импорта страница перезагрузится и восстановит наборы коэффициентов, связки и настройки.",
    )
    queue_project_import_from_upload(uploaded_project, "export_tab")

    render_bottom_tab_switcher("Export/Import", "export")

