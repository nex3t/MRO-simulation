"""
MRO Efficiency Dashboard — City of Chicago
==========================================
Streamlit dashboard — Efficiency tab adapted from Facilities v3.
Shows 10-year MRO spend projection with pillar-based optimization.

Run: streamlit run mro_dashboard.py
"""

import math
from pathlib import Path
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import warnings

warnings.filterwarnings("ignore")

st.set_page_config(
    page_title="MRO Efficiency — City of Chicago",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Constants ──────────────────────────────────────────────────────────────
SIM_OUTPUT     = Path(__file__).parent / "mro_10yr_projection.xlsx"
BASE_YEAR      = 2024
HIST_YEARS     = list(range(2020, 2026))   # actuals available (MRO=Y source complete from 2020)
SIM_YEARS      = list(range(2025, 2037))
FIT_START      = 2020
MATURE_START   = 2018
MATURE_MIN_YRS = 4
MIN_SPEND      = 100_000
MAX_CAGR       = 0.30
DECAY_START_YR = 3
DECAY_RATE     = 0.10
CAGR_FLOOR     = 0.05
DECAY_THRESHOLD= 0.15
OPT_START      = 2023

COLORS = {
    "bg":      "#eef2f6",
    "sidebar": "#e5ebf1",
    "card":    "#ffffff",
    "border":  "rgba(100,116,139,0.35)",
    "text":    "#0f172a",
    "muted":   "#64748b",
    "navy":    "#1e3a5f",
    "green":   "#1a7a4a",
    "orange":  "#d97706",
    "cyan":    "#3b82f6",
    "red":     "#dc2626",
    "yellow":  "#d97706",
    "purple":  "#7c3aed",
    "indigo":  "#1e3a8a",
    "lime":    "#c026d3",
}

_AXIS_BASE = dict(
    gridcolor="rgba(100,116,139,0.18)",
    zerolinecolor="rgba(100,116,139,0.26)",
    linecolor="rgba(100,116,139,0.30)",
    tickfont=dict(color="#334155", size=11),
    title_font=dict(color="#1e293b", size=12),
    tickcolor="#475569",
)
_LEGEND_BASE = dict(
    bgcolor="rgba(255,255,255,0.95)", bordercolor="rgba(100,116,139,0.30)",
    borderwidth=1, font=dict(size=10, color="#334155"),
    orientation="h", y=1.08, x=0, xanchor="left", yanchor="bottom",
    itemsizing="constant", tracegroupgap=4,
)

MRO_PILLARS = {
    "Catalog_Compliance": {
        "name": "Work Modernization",
        "r_p": 0.10, "r_min": 0.02, "r_max": 0.20, "k": 1.5, "x0": 3.0,
        "color": "#1a7a4a",
    },
    "Demand_Management": {
        "name": "Demand Management",
        "r_p": 0.07, "r_min": 0.01, "r_max": 0.18, "k": 1.2, "x0": 4.0,
        "color": "#3b82f6",
    },
    "Predictive_Maintenance": {
        "name": "Asset Management",
        "r_p": 0.09, "r_min": 0.02, "r_max": 0.22, "k": 1.0, "x0": 5.0,
        "color": "#7c3aed",
    },
    "Category_Rationalization": {
        "name": "Vendor Management",
        "r_p": 0.05, "r_min": 0.01, "r_max": 0.15, "k": 1.0, "x0": 4.5,
        "color": "#d97706",
    },
    "Supplier_Consolidation": {
        "name": "Early Pay",
        "r_p": 0.06, "r_min": 0.01, "r_max": 0.15, "k": 0.8, "x0": 4.0,
        "color": "#64748b",
    },
    "Early_Pay_Discounts": {
        "name": "Early Pay Management",
        "r_p": 0.025, "r_min": 0.00, "r_max": 0.08, "k": 2.0, "x0": 5.0,
        "color": "#dc2626",
    },
}

OPERATIONAL_PILLARS = frozenset({
    "Catalog_Compliance", "Demand_Management",
    "Predictive_Maintenance", "Category_Rationalization",
})

# ── Oil Price / Energy Shock Model ─────────────────────────────────────────
OIL_PRE_CRISIS = 65.0   # $/bbl pre-crisis baseline (WTI)

OIL_ASSUMPTIONS = {
    "fleet_fuel_share":        0.35,
    "fac_energy_share":        0.25,
    "road_asphalt_share":      0.30,
    "fuel_passthrough":        0.85,
    "nat_gas_correlation":     0.50,
    "electric_correlation":    0.25,
    "mro_share_fleet":         0.20,
    "mro_share_fac":           0.15,
    "mro_share_road":          0.12,
    "mro_petroleum_sensitive": 0.45,
    "mro_passthrough":         0.65,
    "mro_supply_premium":      0.15,
}

# 2026 baseline spend for Fleet / Facilities / Roadway ($M) — from NIGP trend analysis
PROJ_2026_BASE = {"fleet": 155.0, "facilities": 240.0, "roadway": 210.0}

OIL_SCENARIOS = [
    {"label": "Pre-Crisis Baseline", "price": 65},
    {"label": "Current (~$115)",     "price": 115},
    {"label": "Sustained $120",      "price": 120},
    {"label": "Qatar Warning $150",  "price": 150},
]

MRO_OIL_CATEGORIES = [
    {"name": "Lubricants, Oils & Fluids",   "share": 0.22, "sensitivity": "high",   "dept": "Fleet",      "desc": "Motor oil, hydraulic fluid, brake fluid, coolant"},
    {"name": "Tires & Rubber Products",     "share": 0.15, "sensitivity": "high",   "dept": "Fleet",      "desc": "Synthetic rubber requires petroleum feedstock"},
    {"name": "Vehicle Parts & Components",  "share": 0.18, "sensitivity": "medium", "dept": "Fleet",      "desc": "Filters, belts, hoses, brake pads"},
    {"name": "HVAC Filters, Belts & Parts", "share": 0.12, "sensitivity": "medium", "dept": "Facilities", "desc": "Building HVAC consumables"},
    {"name": "Cleaning & Janitorial",       "share": 0.08, "sensitivity": "medium", "dept": "Facilities", "desc": "Solvents, degreasers, plastic bags"},
    {"name": "Safety & PPE",                "share": 0.07, "sensitivity": "medium", "dept": "All",        "desc": "Plastic face shields, synthetic gloves"},
    {"name": "Equipment Maintenance Tools", "share": 0.10, "sensitivity": "low",    "dept": "Roadway",    "desc": "Cutting fluids, welding consumables"},
    {"name": "Electrical & Lighting",       "share": 0.08, "sensitivity": "low",    "dept": "Facilities", "desc": "Replacement bulbs, wiring, ballasts"},
]

CATEGORY_PILLAR_RATES = {
    "MRO Chemicals & Lubricants":       {"Catalog_Compliance":0.11,"Demand_Management":0.12,"Predictive_Maintenance":0.10,"Category_Rationalization":0.07,"Supplier_Consolidation":0.07,"Early_Pay_Discounts":0.030},
    "Building & Construction Materials":{"Catalog_Compliance":0.09,"Demand_Management":0.06,"Predictive_Maintenance":0.05,"Category_Rationalization":0.11,"Supplier_Consolidation":0.08,"Early_Pay_Discounts":0.020},
    "Fleet & Vehicle Maintenance":      {"Catalog_Compliance":0.11,"Demand_Management":0.10,"Predictive_Maintenance":0.15,"Category_Rationalization":0.06,"Supplier_Consolidation":0.06,"Early_Pay_Discounts":0.025},
    "Plumbing & Fluid Handling":        {"Catalog_Compliance":0.09,"Demand_Management":0.07,"Predictive_Maintenance":0.13,"Category_Rationalization":0.07,"Supplier_Consolidation":0.06,"Early_Pay_Discounts":0.025},
    "Safety, PPE & Fire Protection":    {"Catalog_Compliance":0.12,"Demand_Management":0.10,"Predictive_Maintenance":0.05,"Category_Rationalization":0.07,"Supplier_Consolidation":0.08,"Early_Pay_Discounts":0.030},
    "Electrical & Electronic":          {"Catalog_Compliance":0.09,"Demand_Management":0.07,"Predictive_Maintenance":0.13,"Category_Rationalization":0.07,"Supplier_Consolidation":0.06,"Early_Pay_Discounts":0.020},
    "Tools & Shop Supplies":            {"Catalog_Compliance":0.11,"Demand_Management":0.07,"Predictive_Maintenance":0.05,"Category_Rationalization":0.11,"Supplier_Consolidation":0.08,"Early_Pay_Discounts":0.025},
    "Packaging & Labeling":             {"Catalog_Compliance":0.12,"Demand_Management":0.12,"Predictive_Maintenance":0.02,"Category_Rationalization":0.08,"Supplier_Consolidation":0.09,"Early_Pay_Discounts":0.030},
    "Fasteners & Hardware":             {"Catalog_Compliance":0.11,"Demand_Management":0.07,"Predictive_Maintenance":0.04,"Category_Rationalization":0.12,"Supplier_Consolidation":0.10,"Early_Pay_Discounts":0.025},
    "Material Handling & Storage":      {"Catalog_Compliance":0.08,"Demand_Management":0.06,"Predictive_Maintenance":0.14,"Category_Rationalization":0.06,"Supplier_Consolidation":0.05,"Early_Pay_Discounts":0.020},
    "Abrasives & Surface Prep":         {"Catalog_Compliance":0.12,"Demand_Management":0.10,"Predictive_Maintenance":0.03,"Category_Rationalization":0.09,"Supplier_Consolidation":0.07,"Early_Pay_Discounts":0.030},
    "Facility Maintenance & Janitorial":{"Catalog_Compliance":0.13,"Demand_Management":0.13,"Predictive_Maintenance":0.02,"Category_Rationalization":0.10,"Supplier_Consolidation":0.08,"Early_Pay_Discounts":0.030},
    "HVAC & Refrigeration":             {"Catalog_Compliance":0.08,"Demand_Management":0.05,"Predictive_Maintenance":0.16,"Category_Rationalization":0.05,"Supplier_Consolidation":0.04,"Early_Pay_Discounts":0.020},
}

# ── Math helpers ───────────────────────────────────────────────────────────
def calc_oil_impact(oil_price):
    """Compute oil price shock impacts in $M on Fleet, Facilities, Roadway & MRO."""
    pct  = (oil_price - OIL_PRE_CRISIS) / OIL_PRE_CRISIS
    a    = OIL_ASSUMPTIONS
    b    = PROJ_2026_BASE
    fleet_impact = b["fleet"]      * a["fleet_fuel_share"]  * pct * a["fuel_passthrough"]
    fac_b        = b["facilities"] * a["fac_energy_share"]
    fac_impact   = fac_b * (0.6 * pct * a["nat_gas_correlation"] + 0.4 * pct * a["electric_correlation"])
    road_impact  = b["roadway"]    * a["road_asphalt_share"] * pct * a["fuel_passthrough"]
    mro_base     = (b["fleet"]      * a["mro_share_fleet"] +
                    b["facilities"] * a["mro_share_fac"]   +
                    b["roadway"]    * a["mro_share_road"])
    mro_petro    = mro_base * a["mro_petroleum_sensitive"]
    mro_direct   = mro_petro * pct * a["mro_passthrough"]
    mro_supply   = mro_petro * pct * a["mro_supply_premium"]
    mro_total    = mro_direct + mro_supply
    return {
        "pct":          pct,
        "fleet":        fleet_impact,
        "facilities":   fac_impact,
        "roadway":      road_impact,
        "mro_direct":   mro_direct,
        "mro_supply":   mro_supply,
        "mro_total":    mro_total,
        "total_direct": fleet_impact + fac_impact + road_impact,
        "grand_total":  fleet_impact + fac_impact + road_impact + mro_total,
        "fleet_2026":   b["fleet"]      + fleet_impact,
        "fac_2026":     b["facilities"] + fac_impact,
        "road_2026":    b["roadway"]    + road_impact,
    }

def sigmoid(elapsed, k, x0):
    return 1.0 / (1.0 + math.exp(-k * (elapsed - x0)))

def linear_ramp(elapsed, adopt_yrs):
    if adopt_yrs <= 0:
        return 1.0
    return min(max(elapsed / adopt_yrs, 0.0), 1.0)

def effective_cagr(base_cagr, years_ahead, apply_decay):
    if not apply_decay or years_ahead <= DECAY_START_YR:
        return base_cagr
    extra = years_ahead - DECAY_START_YR
    return max(base_cagr * ((1 - DECAY_RATE) ** extra), CAGR_FLOOR)

def calc_backtest_accuracy(baseline_df):
    """
    Backtest: compare model baseline projections (2023-2025) against known actuals.
    Returns MAPE, accuracy, per-year signed errors, and directional bias.
    """
    # Backtest: first year(s) of simulation where we also have actuals
    backtest_yrs = [yr for yr in SIM_YEARS if yr in _HIST_ACTUALS]
    per_year = {}
    for yr in backtest_yrs:
        predicted = baseline_df[baseline_df["year"] == yr]["baseline"].sum()
        actual    = _HIST_ACTUALS[yr]
        err_pct   = (predicted - actual) / actual   # signed: + = model over-predicts
        per_year[yr] = {"predicted": predicted, "actual": actual, "err_pct": err_pct}
    abs_errors = [abs(v["err_pct"]) for v in per_year.values()]
    mape       = sum(abs_errors) / len(abs_errors) if abs_errors else 0.0
    bias       = sum(v["err_pct"] for v in per_year.values()) / len(per_year) if per_year else 0.0
    return {"mape": mape, "accuracy": 1.0 - mape, "bias": bias, "per_year": per_year}

def fmt(n):
    if abs(n) >= 1e9:  return f"${n/1e9:.2f}B"
    if abs(n) >= 1e6:  return f"${n/1e6:.1f}M"
    if abs(n) >= 1e3:  return f"${n/1e3:.1f}K"
    return f"${n:.0f}"


# ── Data layer ─────────────────────────────────────────────────────────────
_HIST_ACTUALS = {
    # Source: MRO PO lines Extract.xlsx — MRO == "Y" filter only
    2020: 21_660_000, 2021: 58_930_000, 2022: 83_250_000,
    2023: 120_480_000, 2024: 149_010_000, 2025: 232_240_000,
}

@st.cache_data(show_spinner="Loading simulation output...")
def load_profiles():
    df = pd.read_excel(SIM_OUTPUT, sheet_name="Category Profiles")
    return df[["category", "base_spend_2022", "ols_cagr", "cagr", "apply_decay"]].copy()

@st.cache_data(show_spinner="Loading baseline projections...")
def load_baseline():
    df = pd.read_excel(SIM_OUTPUT, sheet_name="Category Projections")
    return df[["year", "category", "baseline"]].copy()

@st.cache_data
def get_historical_actuals():
    rows = [{"year": yr, "amount_ordered": amt} for yr, amt in _HIST_ACTUALS.items()]
    return pd.DataFrame(rows)

@st.cache_data(show_spinner="Loading department shares...")
def get_dept_shares():
    df = pd.read_excel(SIM_OUTPUT, sheet_name="Dept x Category")
    base = df[df["year"] == min(SIM_YEARS)].copy()  # earliest projection year used for shares
    cat_totals = base.groupby("category")["baseline"].sum()
    base = base.join(cat_totals.rename("cat_total"), on="category")
    base["share"] = base["baseline"] / base["cat_total"].replace(0, float("nan"))
    return base[["category", "department", "share"]].dropna(subset=["share"]).copy()


def compute_optimization(baseline_df, pillar_rates, ai_rate, ai_adopt_yrs):
    d       = baseline_df.copy()
    elapsed = (d["year"] - OPT_START).astype(float)
    residual_all = pd.Series(1.0, index=d.index)
    residual_op  = pd.Series(1.0, index=d.index)
    for key, pillar in MRO_PILLARS.items():
        r_override = pillar_rates.get(key, pillar["r_p"])
        k_v, x0_v = pillar["k"], pillar["x0"]
        mu   = elapsed.apply(lambda e, k=k_v, x=x0_v: sigmoid(e, k, x))
        # Scale category rate by slider override ratio
        rates = d["category"].map(
            lambda cat, pkey=key, r_ov=r_override, r_def=pillar["r_p"]:
                CATEGORY_PILLAR_RATES.get(cat, {}).get(pkey, r_def) * (r_ov / r_def if r_def > 0 else 1)
        ).clip(0, 0.99)
        s = (rates * mu).clip(0.0, 0.95)
        residual_all *= (1.0 - s)
        if key in OPERATIONAL_PILLARS:
            residual_op *= (1.0 - s)
    l1    = (d["baseline"] * (1.0 - residual_all)).clip(upper=d["baseline"])
    l1_op = (d["baseline"] * (1.0 - residual_op)).clip(upper=d["baseline"])
    ramp  = elapsed.apply(lambda e: linear_ramp(e, ai_adopt_yrs))
    l2    = (l1_op * ai_rate * ramp).round(2)
    d["l1"]            = l1
    d["l2"]            = l2
    d["optimized"]     = (d["baseline"] - l1 - l2).clip(lower=0)
    d["total_savings"] = d["baseline"] - d["optimized"]
    d["savings_pct"]   = d["total_savings"] / d["baseline"]
    return d

def compute_per_pillar(baseline_df, pillar_rates):
    d       = baseline_df[["year","category","baseline"]].copy()
    elapsed = (d["year"] - OPT_START).astype(float)
    indiv   = {}
    residual = pd.Series(1.0, index=d.index)
    for key, pillar in MRO_PILLARS.items():
        r_override = pillar_rates.get(key, pillar["r_p"])
        k_v, x0_v  = pillar["k"], pillar["x0"]
        mu          = elapsed.apply(lambda e, k=k_v, x=x0_v: sigmoid(e, k, x))
        rates       = d["category"].map(
            lambda cat, pkey=key, r_ov=r_override, r_def=pillar["r_p"]:
                CATEGORY_PILLAR_RATES.get(cat, {}).get(pkey, r_def) * (r_ov / r_def if r_def > 0 else 1)
        ).clip(0, 0.99)
        s            = (rates * mu).clip(0.0, 0.95)
        residual    *= (1.0 - s)
        indiv[key]   = d["baseline"] * rates * mu
    l1_multi  = (d["baseline"] * (1.0 - residual)).clip(upper=d["baseline"])
    indiv_sum = sum(indiv.values())
    scale     = (l1_multi / indiv_sum.replace(0, float("nan"))).fillna(0.0)
    rows = []
    for key, i_sav in indiv.items():
        tmp    = (i_sav * scale).rename("l1")
        yr_sum = d[["year"]].join(tmp).groupby("year")["l1"].sum().reset_index()
        yr_sum["pillar"] = key
        rows.append(yr_sum)
    return pd.concat(rows, ignore_index=True) if rows else pd.DataFrame(columns=["year","pillar","l1"])


# ── Charts ─────────────────────────────────────────────────────────────────
def _fig_base(fig, height=420, **kw):
    kw.setdefault("margin", dict(l=52, r=18, t=48, b=44))
    fig.update_layout(
        paper_bgcolor="#ffffff",
        plot_bgcolor="#ffffff",
        font=dict(color="#334155", size=11),
        height=height,
        hoverlabel=dict(bgcolor="#ffffff", bordercolor="rgba(100,116,139,0.35)", font=dict(color="#0f172a", size=12)),
        **kw,
    )
    fig.update_xaxes(**_AXIS_BASE)
    fig.update_yaxes(**_AXIS_BASE)
    return fig

def chart_trajectory(yearly, hist_actuals, pillar_rates, oil_impact=None, oil_price=65):
    """
    Historical actuals (2017-2025) as raw observed values.
    Optimized line extends back to 2017 (applied to actual spend).
    Trend anchored to actual 2025 for forward segment so the jump is calibrated.
    """
    df = yearly.sort_values("year")
    proj_years = df["year"].tolist()          # 2023-2032
    baseline_m = df["baseline"].values / 1e6
    l2_m       = df["l2"].values / 1e6

    # ── Historical actuals ────────────────────────────────────────────────────
    all_actuals = hist_actuals.sort_values("year")
    act_years   = all_actuals["year"].tolist()
    act_vals    = (all_actuals["amount_ordered"] / 1e6).tolist()
    act_dict    = dict(zip(act_years, act_vals))

    # ── Trend: logistic curve anchored at actual 2025 ─────────────────────────
    # Logistic produces a continuously decelerating S-curve for the projection
    # segment, avoiding the flat "constant-rate" look of pure compounding.
    # f(yr) = L / (1 + A * exp(-k * (yr - 2025)))
    # Parameters: L=saturation capacity, k=curvature, A derived from actual_2025.
    actual_2025_m = act_dict.get(2025, baseline_m[proj_years.index(2025)])
    L_SAT = 500.0   # long-run saturation (~Chicago MRO program capacity, $M)
    K_LOG = 0.123   # logistic growth rate (fitted so 2036 ≈ $360M)
    A_LOG = (L_SAT / actual_2025_m) - 1.0  # derived from f(2025) = actual_2025_m

    proj_start_yr = max(act_years) + 1   # 2026
    logistic_map  = {
        yr: L_SAT / (1.0 + A_LOG * math.exp(-K_LOG * (yr - 2025)))
        for yr in range(proj_start_yr, max(proj_years) + 1)
    }

    trend_y = [b if yr < proj_start_yr else logistic_map[yr]
               for yr, b in zip(proj_years, baseline_m)]

    # ── Optimal: starts at 2026 (first visible projection year), branches from Trend ─
    OPTIMAL_START = proj_start_yr
    optimal_x = [yr for yr in proj_years if yr >= OPTIMAL_START]
    optimal_y = []
    for yr in optimal_x:
        tr_val  = trend_y[proj_years.index(yr)]
        elapsed = float(yr - OPTIMAL_START)
        if elapsed == 0:
            optimal_y.append(tr_val)   # sits exactly on Trend at start
        else:
            residual = 1.0
            for key, pillar in MRO_PILLARS.items():
                r_p = pillar_rates.get(key, pillar["r_p"])
                if r_p <= 0:
                    continue
                mu = sigmoid(elapsed, pillar["k"] * 0.55, pillar["x0"] + 2.0)
                residual *= (1.0 - min(r_p * mu, 0.95))
            optimal_y.append(tr_val * residual)

    # ── Optimized: r_max from 2015, extended back to 2017 via actuals ─────────
    # Historical segment (2017-2022): apply sigmoid to actual spend at each year.
    # Projection segment (2023-2032): apply to anchored Trend, add AI after OPT_START.
    OPTIMIZED_ORIGIN = 2015
    max_rates = {k: v["r_max"] for k, v in MRO_PILLARS.items()}

    def _opt_residual(yr):
        elapsed  = float(yr - OPTIMIZED_ORIGIN)
        residual = 1.0
        for key, r_max in max_rates.items():
            if r_max <= 0:
                continue
            mu = sigmoid(elapsed, MRO_PILLARS[key]["k"], MRO_PILLARS[key]["x0"])
            residual *= (1.0 - min(r_max * mu, 0.95))
        return residual

    hist_years_only = [y for y in act_years if y < proj_years[0]]
    opt_hist_x = hist_years_only
    opt_hist_y = [act_dict[y] * _opt_residual(y) for y in hist_years_only]

    opt_proj_y = []
    for yr, tr_val, ai_sav in zip(proj_years, trend_y, l2_m):
        l1_val = tr_val * _opt_residual(yr)
        if yr < proj_start_yr:
            opt_proj_y.append(l1_val)
        else:
            ai_elapsed = float(yr - proj_start_yr)
            ai_mu = sigmoid(ai_elapsed, 0.85, 3.5)
            opt_proj_y.append(max(l1_val - ai_sav * ai_mu, 0.0))

    opt_all_x = opt_hist_x + proj_years
    opt_all_y = opt_hist_y + opt_proj_y

    fig = go.Figure()

    # Shaded area: Trend → Optimized (projection segment only: 2026-2032)
    fwd_years   = [yr for yr in proj_years if yr >= proj_start_yr]
    fwd_trend   = [trend_y[proj_years.index(yr)] for yr in fwd_years]
    fwd_opt     = [opt_proj_y[proj_years.index(yr)] for yr in fwd_years]
    fig.add_trace(go.Scatter(
        x=fwd_years + fwd_years[::-1],
        y=fwd_trend + fwd_opt[::-1],
        fill="toself",
        fillcolor="rgba(26,122,74,0.08)",
        line=dict(color="rgba(0,0,0,0)"),
        showlegend=False,
        hoverinfo="skip",
    ))

    # Single line: actuals (2017-2025) connected into smooth trend (2026-2032)
    combined_x = act_years + fwd_years
    combined_y = act_vals  + fwd_trend
    symbols = ["diamond"] * len(act_years) + ["circle"] * (len(combined_x) - len(act_years))
    sizes   = [8]         * len(act_years) + [4]         * (len(combined_x) - len(act_years))
    fig.add_trace(go.Scatter(
        x=combined_x, y=combined_y,
        name="Trend / Actuals",
        mode="lines+markers",
        line=dict(color=COLORS["red"], width=2.5),
        marker=dict(symbol=symbols, size=sizes, color=COLORS["red"]),
        hovertemplate="%{x}: $%{y:.1f}M<extra></extra>",
    ))

    # Optimal
    fig.add_trace(go.Scatter(
        x=optimal_x, y=optimal_y,
        name="Optimal (L1 Pillars)",
        mode="lines+markers",
        line=dict(color=COLORS["cyan"], width=2, dash="dot"),
        marker=dict(size=5, symbol="circle-open"),
        hovertemplate="Optimal: $%{y:.1f}M<extra></extra>",
    ))

    # Optimized — full range 2017-2032
    fig.add_trace(go.Scatter(
        x=opt_all_x, y=opt_all_y,
        name="Optimized (L1 max + AI)",
        mode="lines+markers",
        line=dict(color=COLORS["green"], width=2.5),
        marker=dict(size=5, symbol="circle-open"),
        hovertemplate="Optimized: $%{y:.1f}M<extra></extra>",
    ))

    # Oil shock scenario — MRO premium on top of trend (projection segment only)
    if oil_impact is not None and oil_impact["mro_total"] > 0.01:
        shock_y = [trend_y[proj_years.index(yr)] + oil_impact["mro_total"] for yr in fwd_years]
        fig.add_trace(go.Scatter(
            x=[act_years[-1]] + fwd_years,
            y=[act_vals[-1]]  + shock_y,
            name=f"Trend + Energy Shock (${oil_price:.0f}/bbl)",
            mode="lines+markers",
            line=dict(color=COLORS["orange"], width=2, dash="dash"),
            marker=dict(size=4, symbol="circle"),
            hovertemplate="Energy Shock: $%{y:.1f}M<extra></extra>",
        ))

    # Reference lines
    fig.add_vline(x=2022.5, line_dash="dash", line_color="rgba(255,255,255,0.15)", line_width=1)
    fig.add_annotation(x=2022.6, y=0.94, yref="paper", text="Projection ->",
                       showarrow=False, font=dict(color="#64748b", size=10), xanchor="left")
    fig.add_vline(x=2025.5, line_dash="dot", line_color="rgba(255,255,255,0.10)", line_width=1)
    fig.add_annotation(x=2025.6, y=0.84, yref="paper", text="Backtest end",
                       showarrow=False, font=dict(color="#64748b", size=9), xanchor="left")

    _legend = {**_LEGEND_BASE, "y": -0.18, "yanchor": "top"}
    _fig_base(
        fig, height=500, hovermode="x unified",
        title=dict(text="MRO Spend Scenarios — City of Chicago ($M)",
                   y=0.97, x=0, xanchor="left",
                   font=dict(size=14, color="#e8edf5")),
        legend=_legend,
    )
    fig.update_layout(margin=dict(l=52, r=18, t=54, b=80))
    return fig

def chart_pillar_lines(pil_yr):
    if pil_yr.empty:
        return go.Figure()
    fig = go.Figure()
    for key, grp in pil_yr.groupby("pillar"):
        color = MRO_PILLARS.get(key, {}).get("color", COLORS["navy"])
        label = MRO_PILLARS.get(key, {}).get("name", key)
        fig.add_trace(go.Scatter(
            x=grp["year"], y=grp["l1"],
            name=label, mode="lines+markers",
            line=dict(color=color, width=2),
            marker=dict(size=5),
            hovertemplate=f"{label}: $%{{y:,.0f}}<extra></extra>",
        ))
    _fig_base(fig, height=380, hovermode="x unified",
              xaxis=dict(**_AXIS_BASE, title="Year", dtick=2),
              yaxis=dict(**_AXIS_BASE, title="Annual L1 Savings ($)"),
              legend=_LEGEND_BASE)
    return fig

def chart_category_savings(opt_df):
    """Horizontal bar: savings % by category in 2032 (full adoption)."""
    y2032 = opt_df[opt_df["year"] == proj_yr].copy()
    y2032["rate"] = y2032["total_savings"] / y2032["baseline"]
    y2032 = y2032.sort_values("rate")
    short = [c[:30] for c in y2032["category"]]
    colors_bar = [
        COLORS["green"] if r >= 0.40 else
        COLORS["cyan"]  if r >= 0.35 else
        COLORS["orange"]
        for r in y2032["rate"]
    ]
    fig = go.Figure(go.Bar(
        x=y2032["rate"] * 100,
        y=short,
        orientation="h",
        marker_color=colors_bar,
        marker_opacity=0.85,
        text=[f"{r:.1%}" for r in y2032["rate"]],
        textposition="outside",
        textfont=dict(color="#94a3b8", size=10),
        hovertemplate="%{y}: %{x:.1f}%<extra></extra>",
    ))
    _fig_base(fig, height=380,
              xaxis=dict(**_AXIS_BASE, title=f"Savings % at Full Adoption ({proj_yr})", ticksuffix="%"),
              yaxis={**_AXIS_BASE, "tickfont": dict(size=10, color="#334155")})
    return fig

def chart_cumulative_savings(opt_df):
    """Stacked area: cumulative savings buildup L1 vs L2 by year."""
    yrly = opt_df.groupby("year")[["l1","l2","baseline"]].sum().reset_index()
    fig  = go.Figure()
    fig.add_trace(go.Scatter(
        x=yrly["year"], y=yrly["l1"] / 1e6,
        name="L1 Pillar Savings", fill="tozeroy",
        mode="lines",
        line=dict(color=COLORS["cyan"], width=2),
        fillcolor="rgba(96,165,250,0.18)",
        hovertemplate="L1: $%{y:.1f}M<extra></extra>",
    ))
    fig.add_trace(go.Scatter(
        x=yrly["year"], y=(yrly["l1"] + yrly["l2"]) / 1e6,
        name="L1 + L2 (AI)", fill="tonexty",
        mode="lines",
        line=dict(color=COLORS["green"], width=2),
        fillcolor="rgba(45,212,191,0.22)",
        hovertemplate="L1+L2: $%{y:.1f}M<extra></extra>",
    ))
    _fig_base(fig, height=200, hovermode="x unified",
              xaxis=dict(**_AXIS_BASE, dtick=2),
              yaxis=dict(**_AXIS_BASE, title="Annual Savings ($M)"),
              legend=_LEGEND_BASE, margin=dict(l=52, r=18, t=28, b=36))
    return fig


def chart_oil_breakdown(oil_impact):
    """Grouped bar: Fleet / Facilities / Roadway / MRO shock breakdown."""
    categories = ["Fleet", "Facilities", "Roadway", "MRO (direct)", "MRO (supply chain)"]
    values     = [
        oil_impact["fleet"],
        oil_impact["facilities"],
        oil_impact["roadway"],
        oil_impact["mro_direct"],
        oil_impact["mro_supply"],
    ]
    bar_colors = [COLORS["cyan"], COLORS["navy"], COLORS["green"], COLORS["purple"], COLORS["lime"]]
    fig = go.Figure(go.Bar(
        x=categories, y=values,
        marker_color=bar_colors,
        marker_opacity=0.85,
        text=[f"${v:.1f}M" for v in values],
        textposition="outside",
        textfont=dict(color="#475569", size=11),
        hovertemplate="%{x}: $%{y:.1f}M overspend<extra></extra>",
    ))
    _fig_base(fig, height=280,
              xaxis=dict(**_AXIS_BASE, title=""),
              yaxis=dict(**_AXIS_BASE, title="Shock Impact ($M)"),
              margin=dict(l=52, r=18, t=36, b=44))
    return fig


# ── CSS ────────────────────────────────────────────────────────────────────
def inject_css():
    st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=Inter:wght@400;600;700;800&display=swap');
html, body, .stApp {{
    background-color:{COLORS['bg']} !important;
    color:{COLORS['text']};
    font-family:'Inter',sans-serif;
}}
.block-container {{ padding-top:1.2rem !important; padding-bottom:2rem !important; }}
header[data-testid="stHeader"] {{ background-color:{COLORS['bg']} !important; box-shadow:none !important; }}
section[data-testid="stSidebar"] {{
    background-color:{COLORS['sidebar']} !important;
    border-right:1px solid {COLORS['border']} !important;
}}
section[data-testid="stSidebar"] * {{ color:{COLORS['text']} !important; }}
section[data-testid="stSidebar"] .stSlider [data-baseweb="slider"] {{ background:rgba(100,116,139,0.20); }}
div[data-testid="stTabs"] button {{ color:{COLORS['muted']} !important; font-size:13px; }}
div[data-testid="stTabs"] button[aria-selected="true"] {{
    color:{COLORS['navy']} !important; border-bottom:2px solid {COLORS['navy']} !important;
}}
div[data-testid="stMetric"] {{ background:{COLORS['card']}; border:1px solid {COLORS['border']};
    border-radius:10px; padding:12px 16px; box-shadow:0 1px 3px rgba(0,0,0,0.06); }}
.kcard {{
    background:{COLORS['card']};border:1px solid {COLORS['border']};
    border-radius:12px;padding:16px 18px 14px;min-height:108px;
    box-shadow:0 1px 4px rgba(0,0,0,0.07);
}}
.kcard .kc-label {{font-size:9px;font-weight:700;letter-spacing:1.8px;
    text-transform:uppercase;color:{COLORS['muted']};margin-bottom:4px;}}
.kcard .kc-value {{font-size:22px;font-weight:800;
    font-family:'DM Mono',monospace;line-height:1.2;}}
.kcard .kc-sub {{font-size:10px;color:{COLORS['muted']};margin-top:4px;}}
.section-title {{
    font-size:10px;font-weight:700;letter-spacing:2px;text-transform:uppercase;
    color:{COLORS['muted']};margin:18px 0 8px;padding-bottom:6px;
    border-bottom:1px solid {COLORS['border']};
}}
</style>""", unsafe_allow_html=True)

def kcard(label, value, sub, color):
    return f"""<div class='kcard'>
        <div class='kc-label'>{label}</div>
        <div class='kc-value' style='color:{color};'>{value}</div>
        <div class='kc-sub'>{sub}</div>
    </div>"""

def section(title):
    st.markdown(f"<div class='section-title'>{title}</div>", unsafe_allow_html=True)


# ── App ────────────────────────────────────────────────────────────────────
inject_css()

# Header
st.markdown(f"""
<div style='border-bottom:1px solid {COLORS["border"]};padding-bottom:14px;margin-bottom:18px;'>
  <span style='font-size:22px;font-weight:800;color:{COLORS["text"]};'>MRO Efficiency Dashboard</span>
  <span style='font-size:12px;color:{COLORS["muted"]};margin-left:14px;'>
    City of Chicago — Department of Procurement Services
  </span>
</div>""", unsafe_allow_html=True)

# Load data
profiles     = load_profiles()
baseline_df  = load_baseline()
hist_actuals = get_historical_actuals()
shares       = get_dept_shares()

# ── Sidebar ────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"<div style='font-size:13px;font-weight:700;color:{COLORS['cyan']};letter-spacing:1px;margin-bottom:14px;'>⚙️ OPTIMIZATION CONTROLS</div>", unsafe_allow_html=True)

    st.markdown(f"<div style='font-size:9px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:{COLORS['muted']};margin-bottom:8px;'>PILLAR RATES</div>", unsafe_allow_html=True)
    pillar_rates = {}
    for key, p in MRO_PILLARS.items():
        layer_tag = "OPS" if key in OPERATIONAL_PILLARS else "FIN"
        label = f"{p['name']} [{layer_tag}]"
        pct_val = st.slider(
            label,
            min_value=int(p["r_min"] * 100),
            max_value=int(p["r_max"] * 100),
            value=int(p["r_p"] * 100),
            step=1,
            format="%d%%",
            key=f"sl_{key}",
        )
        pillar_rates[key] = pct_val / 100.0

    st.markdown("<hr style='border-color:rgba(255,255,255,0.08);margin:14px 0;'>", unsafe_allow_html=True)
    st.markdown(f"<div style='font-size:9px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:{COLORS['muted']};margin-bottom:8px;'>AI LAYER (L2)</div>", unsafe_allow_html=True)
    ai_rate_pct  = st.slider("AI Amplification Rate", 0, 25, 12, 1, format="%d%%", key="ai_rate")
    ai_rate      = ai_rate_pct / 100.0
    ai_adopt_yrs = st.slider("Adoption Timeline (years)", 3, 15, 7, 1, key="ai_yrs")

    st.markdown("<hr style='border-color:rgba(255,255,255,0.08);margin:14px 0;'>", unsafe_allow_html=True)
    st.markdown(f"<div style='font-size:9px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:{COLORS['muted']};margin-bottom:8px;'>VIEW</div>", unsafe_allow_html=True)
    dept_filter = st.multiselect(
        "Filter Departments",
        options=sorted(shares["department"].dropna().unique()),
        default=[],
        key="dept_filter",
        placeholder="All departments",
    )

# ── Compute optimization ───────────────────────────────────────────────────
if dept_filter:
    cats_in_dept = shares[shares["department"].isin(dept_filter)]["category"].unique()
    base_filtered = baseline_df[baseline_df["category"].isin(cats_in_dept)]
else:
    base_filtered = baseline_df

opt_df     = compute_optimization(base_filtered, pillar_rates, ai_rate, ai_adopt_yrs)
pil_yr     = compute_per_pillar(base_filtered, pillar_rates)
yearly     = opt_df.groupby("year")[["baseline","l1","l2","optimized","total_savings"]].sum().reset_index()
yearly["final"] = yearly["optimized"]
oil_price  = int(st.session_state.get("oil_price_main", 115))
oil_impact = calc_oil_impact(oil_price)

# ── KPI Cards ──────────────────────────────────────────────────────────────
proj_yr    = int(yearly["year"].max())   # use last available year in loaded data
last_proj  = yearly[yearly["year"] == proj_yr].iloc[0]
yr5_data   = yearly[yearly["year"].between(min(SIM_YEARS)+1, min(SIM_YEARS)+5)]

total_base  = last_proj["baseline"]
total_opt   = last_proj["optimized"]
sav_5yr_l1  = yr5_data["l1"].sum()
sav_5yr_l2  = yr5_data["l2"].sum()
sav_5yr_tot = sav_5yr_l1 + sav_5yr_l2
cum_sav     = yearly["total_savings"].sum()

# CAGR 2022->2025 from actuals
_cagr_y1, _cagr_y2 = BASE_YEAR - 2, BASE_YEAR
_vy1 = hist_actuals[hist_actuals["year"]==_cagr_y1]["amount_ordered"].values
_vy2 = hist_actuals[hist_actuals["year"]==_cagr_y2]["amount_ordered"].values
_nyrs = _cagr_y2 - _cagr_y1
cagr_str  = f"{((_vy2[0]/_vy1[0])**(1/_nyrs)-1):.1%}" if len(_vy1)>0 and len(_vy2)>0 else "N/A"
_cagr_lbl = f"MRO CAGR {_cagr_y1}–{_cagr_y2}"
bt        = calc_backtest_accuracy(baseline_df)
bt_color  = COLORS["green"] if bt["accuracy"] >= 0.90 else COLORS["orange"] if bt["accuracy"] >= 0.80 else COLORS["red"]
bt_bias   = "over-predicts" if bt["bias"] > 0.01 else ("under-predicts" if bt["bias"] < -0.01 else "unbiased")
_bt_yrs   = sorted(bt["per_year"].keys())
_bt_range = f"{_bt_yrs[0]}–{_bt_yrs[-1]}" if len(_bt_yrs) > 1 else str(_bt_yrs[0]) if _bt_yrs else "n/a"
bt_sub    = f"MAPE {bt['mape']:.1%} · backtest {_bt_range} · model {bt_bias}"

k1,k2,k3,k4,k5,k6,k7 = st.columns(7)
with k1: st.markdown(kcard(_cagr_lbl, cagr_str, "Current trajectory, no intervention", COLORS["red"]), unsafe_allow_html=True)
with k2: st.markdown(kcard(f"{proj_yr} Baseline", fmt(total_base), "No optimization applied", COLORS["muted"]), unsafe_allow_html=True)
with k3: st.markdown(kcard(f"{proj_yr} Optimized", fmt(total_opt), "With L1 pillars + L2 AI", COLORS["cyan"]), unsafe_allow_html=True)
_yr5_lbl = f"{min(SIM_YEARS)+1}–{min(SIM_YEARS)+5}"
with k4: st.markdown(kcard(f"5-Yr L1 Savings ({_yr5_lbl})", fmt(sav_5yr_l1), "Pillar optimization", COLORS["orange"]), unsafe_allow_html=True)
with k5: st.markdown(kcard(f"5-Yr L2 Savings ({_yr5_lbl})", fmt(sav_5yr_l2), "AI amplification", COLORS["purple"]), unsafe_allow_html=True)
with k6: st.markdown(kcard("10-Yr Total Savings", fmt(cum_sav), f"vs {fmt(yearly['baseline'].sum())} baseline", COLORS["green"]), unsafe_allow_html=True)
with k7: st.markdown(kcard("Model Accuracy", f"{bt['accuracy']:.1%}", bt_sub, bt_color), unsafe_allow_html=True)

with st.expander("Backtest detail — 2023–2025 model vs actuals"):
    bt_rows = []
    for yr, d in sorted(bt["per_year"].items()):
        bt_rows.append({
            "Year":      yr,
            "Actual":    fmt(d["actual"]),
            "Predicted": fmt(d["predicted"]),
            "Error %":   f"{d['err_pct']:+.1%}",
            "Abs Error": fmt(abs(d["predicted"] - d["actual"])),
        })
    st.dataframe(pd.DataFrame(bt_rows), use_container_width=True, hide_index=True)
    st.caption(f"MAPE: {bt['mape']:.2%}  ·  Accuracy: {bt['accuracy']:.2%}  ·  Avg bias: {bt['bias']:+.2%} ({'model over-predicts' if bt['bias']>0 else 'model under-predicts'})")

st.markdown("<br>", unsafe_allow_html=True)

# ── Trajectory chart ───────────────────────────────────────────────────────
section("Cost Trajectory — Historical Actuals + Optimization Scenarios")
st.plotly_chart(chart_trajectory(yearly, hist_actuals, pillar_rates, oil_impact, oil_price), use_container_width=True, key="traj")

st.markdown("<br>", unsafe_allow_html=True)

# ── Energy Impact Section ──────────────────────────────────────────────────
_price_color = COLORS["red"] if oil_price > 120 else COLORS["cyan"] if oil_price > 90 else COLORS["green"]
_shock_sign  = f"+{oil_impact['pct']:.0%}" if oil_impact["pct"] >= 0 else f"{oil_impact['pct']:.0%}"
st.markdown(f"""
<div style='background:{COLORS["sidebar"]};border:1px solid {COLORS["border"]};
     border-radius:12px;padding:22px 28px 10px;margin-bottom:20px;'>
  <div style='display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:12px;'>
    <div>
      <div style='font-size:10px;font-weight:700;letter-spacing:2px;text-transform:uppercase;
                  color:{COLORS["muted"]};margin-bottom:4px;'>🛢️ INTERACTIVE COST ESTIMATOR</div>
      <div style='font-size:14px;font-weight:600;color:{COLORS["text"]};margin-bottom:2px;'>
        Energy Price Shock — Fleet, Facilities, Roadway &amp; MRO
      </div>
      <div style='font-size:12px;color:{COLORS["muted"]};'>
        Move the slider — all 2026 energy impacts update in real time
      </div>
    </div>
    <div style='text-align:right;'>
      <div style='font-size:46px;font-weight:800;color:{_price_color};
                  font-family:"DM Mono",monospace;line-height:1;'>${oil_price}</div>
      <div style='font-size:11px;color:{COLORS["muted"]};'>per barrel (WTI) · {_shock_sign} vs pre-crisis</div>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

oil_price = st.slider(
    "Crude Oil Price ($/bbl)",
    min_value=50, max_value=160, value=oil_price, step=1,
    format="$%d", key="oil_price_main", label_visibility="collapsed",
)
oil_impact = calc_oil_impact(oil_price)   # recompute with live slider value

sc1, sc2, sc3, sc4, sc5 = st.columns(5)
for col, lbl, val, clr in zip(
    [sc1, sc2, sc3, sc4, sc5],
    ["$50", "Pre-crisis ~$65", "Current ~$115", "$150 (Qatar warning)", "$160"],
    [50, 65, 115, 150, 160],
    [COLORS["green"], COLORS["green"], COLORS["cyan"], COLORS["red"], COLORS["red"]],
):
    with col:
        st.markdown(
            f"<div style='text-align:center;font-size:10px;color:{clr};font-weight:600;'>{lbl}</div>",
            unsafe_allow_html=True,
        )

st.markdown("<br>", unsafe_allow_html=True)

# KPI cards
e1, e2, e3, e4 = st.columns(4)
with e1:
    st.markdown(kcard("Fleet Overspend", f"${oil_impact['fleet']:.1f}M", f"Fuel at ${oil_price}/bbl", COLORS["cyan"]), unsafe_allow_html=True)
with e2:
    st.markdown(kcard("Facilities Overspend", f"${oil_impact['facilities']:.1f}M", "Energy / natural gas", COLORS["navy"]), unsafe_allow_html=True)
with e3:
    st.markdown(kcard("Roadway Overspend", f"${oil_impact['roadway']:.1f}M", "Bitumen / asphalt", COLORS["green"]), unsafe_allow_html=True)
with e4:
    st.markdown(kcard("MRO Overspend", f"${oil_impact['mro_total']:.1f}M",
                      f"Direct ${oil_impact['mro_direct']:.1f}M + Supply ${oil_impact['mro_supply']:.1f}M",
                      COLORS["orange"]), unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# Breakdown chart + scenario table side by side
ecol_l, ecol_r = st.columns([1, 1])
_muted = COLORS["muted"]
with ecol_l:
    st.markdown(f"<div style='font-size:11px;font-weight:600;color:{_muted};margin-bottom:6px;'>Impact Breakdown by Category</div>", unsafe_allow_html=True)
    st.plotly_chart(chart_oil_breakdown(oil_impact), use_container_width=True, key="oil_bkdn")

with ecol_r:
    st.markdown(f"<div style='font-size:11px;font-weight:600;color:{_muted};margin-bottom:6px;'>2026 Scenario Comparison</div>", unsafe_allow_html=True)
    scenario_rows = []
    for sc in OIL_SCENARIOS:
        imp = calc_oil_impact(sc["price"])
        scenario_rows.append({
            "Scenario":      sc["label"],
            "Oil ($/bbl)":   f"${sc['price']}",
            "Fleet 2026":    f"${imp['fleet_2026']:.0f}M",
            "Facilities 2026": f"${imp['fac_2026']:.0f}M",
            "Roadway 2026":  f"${imp['road_2026']:.0f}M",
            "MRO Add":       "—" if sc["price"] == 65 else f"+${imp['mro_total']:.1f}M",
            "Total Shock":   f"${imp['grand_total']:.0f}M",
        })
    st.dataframe(pd.DataFrame(scenario_rows), use_container_width=True, hide_index=True)

# MRO sub-category detail (expandable)
with st.expander("MRO Sub-Category Petroleum Sensitivity Detail"):
    sens_color = {"high": COLORS["red"], "medium": COLORS["orange"], "low": COLORS["green"]}
    cat_rows = []
    for c in MRO_OIL_CATEGORIES:
        est_impact = oil_impact["mro_total"] * c["share"]
        cat_rows.append({
            "Sub-Category":  c["name"],
            "Dept":          c["dept"],
            "MRO Share":     f"{c['share']:.0%}",
            "Sensitivity":   c["sensitivity"].upper(),
            "Est. Shock":    f"${est_impact:.1f}M",
            "Description":   c["desc"],
        })
    st.dataframe(pd.DataFrame(cat_rows), use_container_width=True, hide_index=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Row 2: Pillar lines + Category savings ─────────────────────────────────
col_l, col_r = st.columns(2)
with col_l:
    section("Annual Savings by Pillar (L1)")
    st.plotly_chart(chart_pillar_lines(pil_yr), use_container_width=True, key="pil_lines")
with col_r:
    section(f"Savings Rate by Category — Full Adoption ({proj_yr})")
    st.plotly_chart(chart_category_savings(opt_df), use_container_width=True, key="cat_sav")

st.markdown("<br>", unsafe_allow_html=True)

# ── Cumulative savings mini chart ──────────────────────────────────────────
section("Annual Savings Build-up — L1 + L2 Layer")
st.plotly_chart(chart_cumulative_savings(opt_df), use_container_width=True, key="cum_sav")

st.markdown("<br>", unsafe_allow_html=True)

# ── Portfolio table ────────────────────────────────────────────────────────
section("Year-by-Year Portfolio Summary")
tbl = yearly.copy()
tbl["savings_pct"] = tbl["total_savings"] / tbl["baseline"]
display = pd.DataFrame({
    "Year":          tbl["year"],
    "Baseline":      tbl["baseline"].apply(fmt),
    "L1 Savings":    tbl["l1"].apply(fmt),
    "L2 (AI)":       tbl["l2"].apply(fmt),
    "Optimized":     tbl["optimized"].apply(fmt),
    "Saved %":       tbl["savings_pct"].apply(lambda x: f"{x:.1%}"),
})
st.dataframe(display, use_container_width=True, hide_index=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Category savings table ─────────────────────────────────────────────────
section(f"Category Savings Detail — {proj_yr} Full Adoption")
cat_2032 = opt_df[opt_df["year"] == proj_yr].copy()
cat_2032["rate"] = cat_2032["total_savings"] / cat_2032["baseline"]
cat_2032 = cat_2032.sort_values("rate", ascending=False)
cat_display = pd.DataFrame({
    "Category":       cat_2032["category"],
    "2022 Base":      cat_2032["category"].map(
        profiles.set_index("category")["base_spend_2022"]
    ).apply(lambda x: fmt(x) if pd.notna(x) else "—"),
    f"{proj_yr} Baseline": cat_2032["baseline"].apply(fmt),
    f"{proj_yr} Optimized": cat_2032["optimized"].apply(fmt),
    "Total Savings":  cat_2032["total_savings"].apply(fmt),
    "Savings %":      cat_2032["rate"].apply(lambda x: f"{x:.1%}"),
})
st.dataframe(cat_display, use_container_width=True, hide_index=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Dept breakdown ─────────────────────────────────────────────────────────
section("Department Breakdown — 2032 Optimized Spend")
dept_df = opt_df.merge(shares, on="category", how="left")
dept_df[["baseline","optimized","total_savings","l1","l2"]] = \
    dept_df[["baseline","optimized","total_savings","l1","l2"]].multiply(dept_df["share"], axis=0)

dept_2032 = (dept_df[dept_df["year"] == proj_yr]
             .groupby("department")[["baseline","optimized","total_savings"]]
             .sum().reset_index())
dept_2032["rate"] = dept_2032["total_savings"] / dept_2032["baseline"]
dept_2032 = dept_2032.sort_values("baseline", ascending=False).head(15)

dept_display = pd.DataFrame({
    "Department":   dept_2032["department"],
    "Baseline":     dept_2032["baseline"].apply(fmt),
    "Optimized":    dept_2032["optimized"].apply(fmt),
    "Savings":      dept_2032["total_savings"].apply(fmt),
    "Savings %":    dept_2032["rate"].apply(lambda x: f"{x:.1%}"),
})
st.dataframe(dept_display, use_container_width=True, hide_index=True)

# Footer
st.markdown(f"""<br>
<div style='font-size:10px;color:{COLORS["muted"]};border-top:1px solid {COLORS["border"]};padding-top:10px;'>
  Data: NIGP Line Item PO data (2017-2025) — 227K MRO lines — $932M total MRO spend<br>
  Model: OLS log-linear (2020-2022) + 10% CAGR decay after yr 3 + Multiplicative pillar optimization
</div>""", unsafe_allow_html=True)
