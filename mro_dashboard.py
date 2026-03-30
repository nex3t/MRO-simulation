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
BASE_YEAR      = 2022
HIST_YEARS     = list(range(2017, 2026))   # actuals available
SIM_YEARS      = list(range(2023, 2037))
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

CATEGORY_PILLAR_RATES = {
    "Adhesives, Sealants & Fasteners":   {"Catalog_Compliance":0.12,"Demand_Management":0.08,"Predictive_Maintenance":0.04,"Category_Rationalization":0.07,"Supplier_Consolidation":0.07,"Early_Pay_Discounts":0.030},
    "Fans, Air Circulators & Blowers":   {"Catalog_Compliance":0.08,"Demand_Management":0.05,"Predictive_Maintenance":0.15,"Category_Rationalization":0.04,"Supplier_Consolidation":0.04,"Early_Pay_Discounts":0.020},
    "Hardware, Fasteners & Bearings":    {"Catalog_Compliance":0.11,"Demand_Management":0.07,"Predictive_Maintenance":0.07,"Category_Rationalization":0.11,"Supplier_Consolidation":0.09,"Early_Pay_Discounts":0.025},
    "Lubricants, Oils & Greases":        {"Catalog_Compliance":0.10,"Demand_Management":0.11,"Predictive_Maintenance":0.12,"Category_Rationalization":0.08,"Supplier_Consolidation":0.06,"Early_Pay_Discounts":0.025},
    "Paint, Coatings & Sealants":        {"Catalog_Compliance":0.10,"Demand_Management":0.08,"Predictive_Maintenance":0.04,"Category_Rationalization":0.08,"Supplier_Consolidation":0.07,"Early_Pay_Discounts":0.025},
    "Electrical Equipment & Supplies":   {"Catalog_Compliance":0.09,"Demand_Management":0.07,"Predictive_Maintenance":0.13,"Category_Rationalization":0.06,"Supplier_Consolidation":0.06,"Early_Pay_Discounts":0.020},
    "Fire Protection Equipment":         {"Catalog_Compliance":0.07,"Demand_Management":0.04,"Predictive_Maintenance":0.14,"Category_Rationalization":0.04,"Supplier_Consolidation":0.05,"Early_Pay_Discounts":0.020},
    "Maintenance & Repair Supplies":     {"Catalog_Compliance":0.11,"Demand_Management":0.10,"Predictive_Maintenance":0.12,"Category_Rationalization":0.09,"Supplier_Consolidation":0.08,"Early_Pay_Discounts":0.025},
    "Tools, Hand & Power":               {"Catalog_Compliance":0.11,"Demand_Management":0.07,"Predictive_Maintenance":0.05,"Category_Rationalization":0.10,"Supplier_Consolidation":0.08,"Early_Pay_Discounts":0.025},
    "Instruments, Gauges & Meters":      {"Catalog_Compliance":0.07,"Demand_Management":0.05,"Predictive_Maintenance":0.16,"Category_Rationalization":0.05,"Supplier_Consolidation":0.04,"Early_Pay_Discounts":0.020},
    "Janitorial Supplies":               {"Catalog_Compliance":0.13,"Demand_Management":0.13,"Predictive_Maintenance":0.02,"Category_Rationalization":0.10,"Supplier_Consolidation":0.08,"Early_Pay_Discounts":0.030},
    "HVAC Equipment & Supplies":         {"Catalog_Compliance":0.08,"Demand_Management":0.06,"Predictive_Maintenance":0.15,"Category_Rationalization":0.05,"Supplier_Consolidation":0.04,"Early_Pay_Discounts":0.020},
    "Hardware":                          {"Catalog_Compliance":0.10,"Demand_Management":0.07,"Predictive_Maintenance":0.05,"Category_Rationalization":0.12,"Supplier_Consolidation":0.09,"Early_Pay_Discounts":0.025},
    "Abrasives":                         {"Catalog_Compliance":0.12,"Demand_Management":0.09,"Predictive_Maintenance":0.03,"Category_Rationalization":0.09,"Supplier_Consolidation":0.07,"Early_Pay_Discounts":0.030},
    "Uniforms & Protective Clothing":    {"Catalog_Compliance":0.10,"Demand_Management":0.08,"Predictive_Maintenance":0.01,"Category_Rationalization":0.07,"Supplier_Consolidation":0.05,"Early_Pay_Discounts":0.025},
    "Pipes, Valves & Fittings":          {"Catalog_Compliance":0.09,"Demand_Management":0.08,"Predictive_Maintenance":0.13,"Category_Rationalization":0.06,"Supplier_Consolidation":0.06,"Early_Pay_Discounts":0.025},
}

# ── Math helpers ───────────────────────────────────────────────────────────
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

def fmt(n):
    if abs(n) >= 1e9:  return f"${n/1e9:.2f}B"
    if abs(n) >= 1e6:  return f"${n/1e6:.1f}M"
    if abs(n) >= 1e3:  return f"${n/1e3:.1f}K"
    return f"${n:.0f}"


# ── Data layer ─────────────────────────────────────────────────────────────
_HIST_ACTUALS = {
    2017: 30_080_370, 2018: 30_231_697, 2019: 40_245_268,
    2020: 57_105_823, 2021: 94_484_900, 2022: 114_627_336,
    2023: 150_111_635, 2024: 190_642_205, 2025: 199_051_136,
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
    base = df[df["year"] == 2023].copy()  # earliest projection year used for shares
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

def chart_trajectory(yearly, hist_actuals, pillar_rates):
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

opt_df   = compute_optimization(base_filtered, pillar_rates, ai_rate, ai_adopt_yrs)
pil_yr   = compute_per_pillar(base_filtered, pillar_rates)
yearly   = opt_df.groupby("year")[["baseline","l1","l2","optimized","total_savings"]].sum().reset_index()
yearly["final"] = yearly["optimized"]

# ── KPI Cards ──────────────────────────────────────────────────────────────
proj_yr    = int(yearly["year"].max())   # use last available year in loaded data
last_proj  = yearly[yearly["year"] == proj_yr].iloc[0]
yr5_data   = yearly[yearly["year"].between(2026, 2030)]

total_base  = last_proj["baseline"]
total_opt   = last_proj["optimized"]
sav_5yr_l1  = yr5_data["l1"].sum()
sav_5yr_l2  = yr5_data["l2"].sum()
sav_5yr_tot = sav_5yr_l1 + sav_5yr_l2
cum_sav     = yearly["total_savings"].sum()

# CAGR 2022->2025 from actuals
_v22 = hist_actuals[hist_actuals["year"]==2022]["amount_ordered"].values
_v25 = hist_actuals[hist_actuals["year"]==2025]["amount_ordered"].values
cagr_str = f"{((_v25[0]/_v22[0])**(1/3)-1):.1%}" if len(_v22)>0 and len(_v25)>0 else "N/A"

k1,k2,k3,k4,k5,k6 = st.columns(6)
with k1: st.markdown(kcard(f"MRO CAGR 2022-2025", cagr_str, "Current trajectory, no intervention", COLORS["red"]), unsafe_allow_html=True)
with k2: st.markdown(kcard(f"{proj_yr} Baseline", fmt(total_base), "No optimization applied", COLORS["muted"]), unsafe_allow_html=True)
with k3: st.markdown(kcard(f"{proj_yr} Optimized", fmt(total_opt), "With L1 pillars + L2 AI", COLORS["cyan"]), unsafe_allow_html=True)
with k4: st.markdown(kcard("5-Yr L1 Savings (2026-2030)", fmt(sav_5yr_l1), "Pillar optimization", COLORS["orange"]), unsafe_allow_html=True)
with k5: st.markdown(kcard("5-Yr L2 Savings (2026-2030)", fmt(sav_5yr_l2), "AI amplification", COLORS["purple"]), unsafe_allow_html=True)
with k6: st.markdown(kcard("10-Yr Total Savings", fmt(cum_sav), f"vs ${fmt(yearly['baseline'].sum())} baseline", COLORS["green"]), unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Trajectory chart ───────────────────────────────────────────────────────
section("Cost Trajectory — Historical Actuals + Optimization Scenarios")
st.plotly_chart(chart_trajectory(yearly, hist_actuals, pillar_rates), use_container_width=True, key="traj")

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
