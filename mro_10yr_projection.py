"""
MRO 10-Year Projection with Pillar Optimization
================================================
1. Fit OLS log-linear trend per category (2020-2022 window)
   - MATURE categories (data since <=2018, 4+ yrs): 2017-2022 window
   - CAGR cap: 30%
2. CAGR decay: after year 3, growth decelerates 10%/yr toward a 5% floor
   - Prevents unrealistic 30% compounding over a 10-year horizon
   - Applies to RAMP categories and MATURE with capped CAGR
3. Pillar optimization (L1+L2) — same math as Facilities Dashboard
   - Rates differentiated by category (each category responds differently)
   - L1: multiplicative sigmoid adoption
   - L2: AI amplification on operational pillars only, linear ramp
4. Dept x category breakdown with category-specific savings rates

CAGR Decay parameters:
  DECAY_START_YR  = 3   (years from BASE_YEAR; decay kicks in at 2026)
  DECAY_RATE      = 0.10 (reduce effective CAGR by 10% each year after decay start)
  CAGR_FLOOR      = 0.05 (minimum 5% growth floor for any category)
  Applies to: categories where capped CAGR > DECAY_THRESHOLD (15%)

Category Pillar Rate logic:
  Catalog_Compliance    — highest for commodity/repetitive (Janitorial, Adhesives, Abrasives)
  Demand_Management     — highest for predictable consumables (Lubricants, Janitorial)
  Predictive_Maintenance— highest for equipment-dependent (HVAC, Fans, Instruments, Fire)
  Category_Rationalization — highest for fragmented SKU categories (Hardware, Tools)
  Supplier_Consolidation— highest for many-supplier categories (Hardware, Adhesives)
  Early_Pay_Discounts   — uniform, slightly higher for high-volume
"""

import math
import pandas as pd
import numpy as np
from scipy import stats
import warnings

warnings.filterwarnings("ignore")

# ── Constants ──────────────────────────────────────────────────────────────
DATA_PATH      = "NIGP_Line_Item_Descriptions.xlsx"
BASE_YEAR      = 2022
SIM_YEARS      = list(range(2023, 2037))
FIT_START      = 2020
MATURE_START   = 2018
MATURE_MIN_YRS = 4
MIN_SPEND      = 100_000
MAX_CAGR       = 0.30

# CAGR decay (post year-3 growth deceleration)
DECAY_START_YR  = 3      # years after BASE_YEAR (decay starts at year 2025->2026 projection)
DECAY_RATE      = 0.40   # 40% reduction per year after decay start (faster convergence to floor)
CAGR_FLOOR      = 0.05   # minimum growth floor (~Chicago CPI + organic demand)
DECAY_THRESHOLD = 0.15   # only apply decay to categories above this CAGR

# AI Layer
AI_RATE      = 0.12
AI_ADOPT_YRS = 7
OPT_START    = 2023


# ── MRO Pillar Definitions (sigmoid shape params) ─────────────────────────
# r_p here is the PORTFOLIO DEFAULT rate — overridden per category below
MRO_PILLARS: dict[str, dict] = {
    "Catalog_Compliance": {
        "name": "Catalog & Contract Compliance",
        "desc": "Drive spend through pre-negotiated contracts, eliminate maverick buying",
        "r_p": 0.10, "k": 1.5, "x0": 3.0,
    },
    "Demand_Management": {
        "name": "Demand Management",
        "desc": "Right-size inventory, VMI, reduce over-ordering and emergency spot buys",
        "r_p": 0.07, "k": 1.2, "x0": 4.0,
    },
    "Predictive_Maintenance": {
        "name": "Predictive Maintenance",
        "desc": "Shift reactive to predictive, cut unplanned/corrective spend",
        "r_p": 0.09, "k": 1.0, "x0": 5.0,
    },
    "Category_Rationalization": {
        "name": "Category Rationalization",
        "desc": "Standardize SKUs, eliminate duplicates across 43 departments",
        "r_p": 0.05, "k": 1.0, "x0": 4.5,
    },
    "Supplier_Consolidation": {
        "name": "Supplier Consolidation",
        "desc": "Reduce 1,142 suppliers to strategic base, leverage volume discounts",
        "r_p": 0.06, "k": 0.8, "x0": 4.0,
    },
    "Early_Pay_Discounts": {
        "name": "Early Pay Discounts",
        "desc": "Early payment terms for pricing benefits across MRO contracts",
        "r_p": 0.025, "k": 2.0, "x0": 5.0,
    },
}

OPERATIONAL_PILLARS = frozenset({
    "Catalog_Compliance",
    "Demand_Management",
    "Predictive_Maintenance",
    "Category_Rationalization",
})

# ── Category-level pillar rates ────────────────────────────────────────────
# Each entry overrides the portfolio default r_p for that category.
# Rationale:
#   CC  = Catalog Compliance    (high for commodity/repetitive purchasing)
#   DM  = Demand Management     (high for consumables with stable/predictable demand)
#   PM  = Predictive Maintenance(high for equipment-driven categories)
#   CR  = Category Rationalization (high where SKU fragmentation is greatest)
#   SC  = Supplier Consolidation(high where supplier count is highest)
#   EP  = Early Pay Discounts   (relatively uniform)
#
#                                              CC     DM     PM     CR     SC     EP
CATEGORY_PILLAR_RATES: dict[str, dict] = {
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


# ── Adoption curves ────────────────────────────────────────────────────────
def sigmoid(elapsed: float, k: float, x0: float) -> float:
    return 1.0 / (1.0 + math.exp(-k * (elapsed - x0)))

def linear_ramp(elapsed: float, adopt_yrs: int) -> float:
    if adopt_yrs <= 0:
        return 1.0
    return min(max(elapsed / adopt_yrs, 0.0), 1.0)


# ── 1. Load Data ───────────────────────────────────────────────────────────
def load_mro(path: str) -> pd.DataFrame:
    df  = pd.read_excel(path, sheet_name="POs")
    mro = df[df["MRO"] == "MRO"].copy()
    mro["year"] = pd.to_datetime(mro["po_approved_date"]).dt.year
    return mro[mro["year"].between(2017, max(SIM_YEARS))]


# ── 2. Fit Category Profiles ───────────────────────────────────────────────
def fit_profiles(mro: pd.DataFrame) -> pd.DataFrame:
    records = []
    for cat, grp in mro[mro["year"].between(2017, BASE_YEAR)].groupby("Commodity Description"):
        annual = grp.groupby("year")["amount_ordered"].sum()
        active = annual[annual >= MIN_SPEND]
        if len(active) == 0:
            continue

        data_start = int(active.index.min())
        n_total    = len(active)
        is_mature  = (data_start <= MATURE_START) and (n_total >= MATURE_MIN_YRS)
        win_start  = 2017 if is_mature else FIT_START

        window = active[active.index >= win_start]
        if len(window) < 2:
            window = active

        years_arr  = window.index.astype(float).values
        log_spends = np.log(window.values)
        slope, _, r2, _, _ = stats.linregress(years_arr, log_spends)

        base = float(annual.get(BASE_YEAR, np.exp(slope * BASE_YEAR + np.mean(log_spends - slope * years_arr))))
        if base < MIN_SPEND:
            base = float(np.exp(slope * BASE_YEAR + np.mean(log_spends - slope * years_arr)))

        ols_cagr   = float(np.exp(slope) - 1)
        cagr_capped = min(max(ols_cagr, -0.20), MAX_CAGR)
        apply_decay = (cagr_capped > DECAY_THRESHOLD)

        records.append({
            "category":        cat,
            "type":            "MATURE" if is_mature else "RAMP",
            "fit_start":       int(min(years_arr)),
            "n_fit_years":     len(window),
            "base_spend_2022": base,
            "ols_cagr":        ols_cagr,
            "cagr":            cagr_capped,
            "apply_decay":     apply_decay,
            "r2":              float(r2**2),
        })

    return pd.DataFrame(records).sort_values("base_spend_2022", ascending=False).reset_index(drop=True)


# ── 3. Effective CAGR with decay ───────────────────────────────────────────
def effective_cagr(base_cagr: float, years_ahead: int, apply_decay: bool) -> float:
    """
    Returns the effective CAGR for a given year into the projection.
    After DECAY_START_YR, growth decelerates by DECAY_RATE per additional year,
    floored at CAGR_FLOOR.
    """
    if not apply_decay or years_ahead <= DECAY_START_YR:
        return base_cagr
    extra   = years_ahead - DECAY_START_YR
    decayed = base_cagr * ((1 - DECAY_RATE) ** extra)
    return max(decayed, CAGR_FLOOR)


# ── 4. Baseline Projection ─────────────────────────────────────────────────
def project_baseline(profiles: pd.DataFrame) -> pd.DataFrame:
    """
    Year-by-year baseline projection applying per-year effective CAGR.
    Decay is applied incrementally, not as a simple power of the base CAGR.
    """
    rows = []
    for _, p in profiles.iterrows():
        spend = p["base_spend_2022"]
        for yr in SIM_YEARS:
            yr_idx = yr - BASE_YEAR          # years since anchor
            g      = effective_cagr(p["cagr"], yr_idx, p["apply_decay"])
            spend  = p["base_spend_2022"]    # reset to base each time (compound from base)
            # Compound year by year with per-year effective CAGR
            s = p["base_spend_2022"]
            for t in range(1, yr_idx + 1):
                s *= (1 + effective_cagr(p["cagr"], t, p["apply_decay"]))
            rows.append({
                "year":     yr,
                "category": p["category"],
                "baseline": s,
                "apply_decay": p["apply_decay"],
            })
    return pd.DataFrame(rows)


# ── 5. Dept shares from 2022 ───────────────────────────────────────────────
def dept_shares(mro: pd.DataFrame) -> pd.DataFrame:
    base = mro[mro["year"] == BASE_YEAR]
    cd   = base.groupby(["Commodity Description", "department_code"])["amount_ordered"].sum()
    tots = cd.groupby(level=0).transform("sum")
    s    = (cd / tots).reset_index()
    s.columns = ["category", "department", "share"]
    return s


# ── 6. Apply Pillar Optimization ───────────────────────────────────────────
def apply_optimization(
    baseline_df: pd.DataFrame,
    ai_rate: float = AI_RATE,
    ai_adopt_yrs: int = AI_ADOPT_YRS,
    opt_start: int  = OPT_START,
) -> pd.DataFrame:
    """
    Per-category pillar rates from CATEGORY_PILLAR_RATES.
    Falls back to MRO_PILLARS defaults for any category not in the table.
    """
    d       = baseline_df.copy()
    elapsed = (d["year"] - opt_start).astype(float)

    residual_all = pd.Series(1.0, index=d.index)
    residual_op  = pd.Series(1.0, index=d.index)

    for key, pillar in MRO_PILLARS.items():
        k_v  = pillar["k"]
        x0_v = pillar["x0"]
        mu   = elapsed.apply(lambda e, k=k_v, x=x0_v: sigmoid(e, k, x))

        # Per-row rate: lookup category-specific rate, fall back to pillar default
        rates = d["category"].map(
            lambda cat, pkey=key: CATEGORY_PILLAR_RATES.get(cat, {}).get(pkey, pillar["r_p"])
        )
        s = (rates * mu).clip(0.0, 0.95)
        residual_all *= (1.0 - s)
        if key in OPERATIONAL_PILLARS:
            residual_op *= (1.0 - s)

    l1    = (d["baseline"] * (1.0 - residual_all)).clip(upper=d["baseline"])
    l1_op = (d["baseline"] * (1.0 - residual_op)).clip(upper=d["baseline"])

    ramp = elapsed.apply(lambda e: linear_ramp(e, ai_adopt_yrs))
    l2   = (l1_op * ai_rate * ramp).round(2)

    d["l1"]            = l1
    d["l2"]            = l2
    d["optimized"]     = (d["baseline"] - l1 - l2).clip(lower=0)
    d["total_savings"] = d["baseline"] - d["optimized"]
    d["savings_pct"]   = d["total_savings"] / d["baseline"]
    return d


def savings_by_pillar(
    baseline_df: pd.DataFrame,
    opt_start: int = OPT_START,
) -> pd.DataFrame:
    d       = baseline_df[["year", "category", "baseline"]].copy()
    elapsed = (d["year"] - opt_start).astype(float)

    indiv:   dict[str, pd.Series] = {}
    residual = pd.Series(1.0, index=d.index)

    for key, pillar in MRO_PILLARS.items():
        k_v  = pillar["k"]
        x0_v = pillar["x0"]
        mu   = elapsed.apply(lambda e, k=k_v, x=x0_v: sigmoid(e, k, x))
        rates = d["category"].map(
            lambda cat, pkey=key: CATEGORY_PILLAR_RATES.get(cat, {}).get(pkey, pillar["r_p"])
        )
        s        = (rates * mu).clip(0.0, 0.95)
        residual *= (1.0 - s)
        indiv[key] = d["baseline"] * rates * mu

    if not indiv:
        return pd.DataFrame(columns=["year", "pillar", "pillar_name", "l1_savings"])

    l1_multi  = (d["baseline"] * (1.0 - residual)).clip(upper=d["baseline"])
    indiv_sum = sum(indiv.values())
    scale     = (l1_multi / indiv_sum.replace(0, float("nan"))).fillna(0.0)

    rows = []
    for key, i_sav in indiv.items():
        tmp    = (i_sav * scale).rename("l1_savings")
        yr_sum = d[["year"]].join(tmp).groupby("year")["l1_savings"].sum().reset_index()
        yr_sum["pillar"]      = key
        yr_sum["pillar_name"] = MRO_PILLARS[key]["name"]
        rows.append(yr_sum)

    return pd.concat(rows, ignore_index=True)[["year", "pillar", "pillar_name", "l1_savings"]]


# ── 7. Dept breakdown ─────────────────────────────────────────────────────
def breakdown_by_dept(opt_df: pd.DataFrame, shares: pd.DataFrame) -> pd.DataFrame:
    merged = opt_df.merge(shares, on="category", how="left")
    merged["share"] = merged["share"].fillna(1.0)
    for col in ["baseline", "l1", "l2", "optimized", "total_savings"]:
        merged[col] = merged[col] * merged["share"]
    return merged[["year", "department", "category", "baseline", "l1", "l2", "optimized", "total_savings"]]


# ── 8. Print ───────────────────────────────────────────────────────────────
def print_results(profiles, portfolio, pillar_df, dept_df):
    div = "=" * 80

    # ── Profiles ──
    print(f"\n{div}")
    print("CATEGORY PROFILES  (* = CAGR capped | D = decay applied after yr 3)")
    print(div)
    print(f"  {'Category':<40} {'Type':>6} {'Yrs':>4} {'Base 2022':>14} {'OLS':>8} {'Cap':>7} {'D':>2}")
    print(f"  {'-'*40} {'-'*6} {'-'*4} {'-'*14} {'-'*8} {'-'*7} {'-'*2}")
    for _, p in profiles.iterrows():
        cap_flag   = "*" if abs(p["ols_cagr"]) > MAX_CAGR else " "
        decay_flag = "D" if p["apply_decay"] else " "
        print(f"  {p['category'][:39]:<40} {p['type']:>6} {p['n_fit_years']:>4}"
              f"  ${p['base_spend_2022']:>12,.0f} {p['ols_cagr']:>7.1%} {p['cagr']:>6.1%} {cap_flag}{decay_flag}")
    print(f"  Total 2022 base: ${profiles['base_spend_2022'].sum():,.0f}")

    # ── Pillars ──
    print(f"\n{div}")
    print("MRO OPTIMIZATION PILLARS  (rates differentiated by category)")
    print(div)
    for key, p in MRO_PILLARS.items():
        layer = "OPS+AI" if key in OPERATIONAL_PILLARS else "FIN   "
        cat_rates = [CATEGORY_PILLAR_RATES.get(c, {}).get(key, p["r_p"]) for c in CATEGORY_PILLAR_RATES]
        print(f"  [{layer}] {p['name']:<35}  default={p['r_p']:.1%}  "
              f"range=[{min(cat_rates):.1%}-{max(cat_rates):.1%}]  k={p['k']}  x0=yr{p['x0']}")
    print(f"  AI Layer: {AI_RATE:.0%} on operational L1, {AI_ADOPT_YRS}-yr linear ramp")

    # ── Portfolio ──
    print(f"\n{div}")
    print("PORTFOLIO PROJECTION: BASELINE vs OPTIMIZED  (2023-2032)")
    print(div)
    yrly = portfolio.groupby("year")[["baseline","l1","l2","optimized","total_savings"]].sum()
    print(f"\n  {'Year':>4}  {'Baseline':>12}  {'L1 Savings':>12}  {'L2 (AI)':>10}  {'Optimized':>12}  {'Saved%':>7}")
    print(f"  {'-'*4}  {'-'*12}  {'-'*12}  {'-'*10}  {'-'*12}  {'-'*7}")
    for yr, row in yrly.iterrows():
        pct = row["total_savings"] / row["baseline"]
        print(f"  {yr:>4}  ${row['baseline']/1e6:>9.1f}M  ${row['l1']/1e6:>9.1f}M"
              f"  ${row['l2']/1e6:>7.1f}M  ${row['optimized']/1e6:>9.1f}M  {pct:>6.1%}")

    cum_base = yrly["baseline"].sum()
    cum_opt  = yrly["optimized"].sum()
    cum_sav  = yrly["total_savings"].sum()
    print(f"\n  Cumulative 2023-2032:")
    print(f"    Baseline:  ${cum_base/1e9:.2f}B")
    print(f"    Optimized: ${cum_opt/1e9:.2f}B")
    print(f"    Saved:     ${cum_sav/1e6:.0f}M  ({cum_sav/cum_base:.1%} of baseline)")

    # ── Pillar attribution by year ──
    print(f"\n{div}")
    print("L1 SAVINGS BY PILLAR  (portfolio, $M/yr)")
    print(div)
    piv = pillar_df.pivot_table(index="year", columns="pillar_name", values="l1_savings", aggfunc="sum") / 1e6
    print(piv.round(1).to_string())

    # ── Category-level savings rate comparison ──
    print(f"\n{div}")
    print("SAVINGS RATE BY CATEGORY  (full-adoption year 2032)")
    print(div)
    y2032 = portfolio[portfolio["year"] == 2032].copy()
    y2032["rate"] = y2032["total_savings"] / y2032["baseline"]
    y2032 = y2032.sort_values("rate", ascending=False)
    for _, r in y2032.iterrows():
        bar = "|" * min(int(r["rate"] * 50), 50)
        print(f"  {r['category'][:42]:<42}  {r['rate']:>6.1%}  {bar}")

    # ── Dept breakdown ──
    print(f"\n{div}")
    print("DEPARTMENT BREAKDOWN  (P50 Baseline vs Optimized, $M)")
    print(div)
    for yr in [2025, 2027, 2030, 2032]:
        yr_dept = (dept_df[dept_df["year"] == yr]
                   .groupby("department")[["baseline","optimized","total_savings"]]
                   .sum().sort_values("baseline", ascending=False).head(8))
        print(f"\n  {yr}:")
        print(f"  {'Department':<52}  {'Baseline':>10}  {'Optimized':>10}  {'Saved':>7}")
        print(f"  {'-'*52}  {'-'*10}  {'-'*10}  {'-'*7}")
        for dept, row in yr_dept.iterrows():
            pct = row["total_savings"] / row["baseline"] if row["baseline"] > 0 else 0
            print(f"  {str(dept)[:51]:<52}  ${row['baseline']/1e6:>7.1f}M  ${row['optimized']/1e6:>7.1f}M  {pct:>6.1%}")

    print(f"\n{div}\n")


# ── Main ───────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("[ 1/5 ] Loading MRO data...")
    mro = load_mro(DATA_PATH)

    print("[ 2/5 ] Fitting category profiles...")
    profiles = fit_profiles(mro)
    print(f"        {len(profiles)} categories  |  2022 base: ${profiles['base_spend_2022'].sum():,.0f}")

    print("[ 3/5 ] Projecting baseline 2023-2032 (with CAGR decay)...")
    baseline_df = project_baseline(profiles)

    print("[ 4/5 ] Applying category-differentiated pillar optimization...")
    shares     = dept_shares(mro)
    opt_df     = apply_optimization(baseline_df)
    pillar_df  = savings_by_pillar(baseline_df)
    dept_df    = breakdown_by_dept(opt_df, shares)

    print("[ 5/5 ] Writing output...")
    print_results(profiles, opt_df, pillar_df, dept_df)

    out = "mro_10yr_projection.xlsx"
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        profiles.to_excel(writer, sheet_name="Category Profiles", index=False)

        yrly = opt_df.groupby("year")[["baseline","l1","l2","optimized","total_savings"]].sum().reset_index()
        yrly["savings_pct"] = yrly["total_savings"] / yrly["baseline"]
        yrly.to_excel(writer, sheet_name="Portfolio by Year", index=False)

        pillar_df.to_excel(writer, sheet_name="Savings by Pillar", index=False)

        cat_yrly = opt_df.groupby(["year","category"])[["baseline","optimized","total_savings","savings_pct"]].sum().reset_index()
        cat_yrly.to_excel(writer, sheet_name="Category Projections", index=False)

        dept_df.to_excel(writer, sheet_name="Dept x Category", index=False)

        dept_sum = dept_df.groupby(["year","department"])[["baseline","optimized","total_savings"]].sum().reset_index()
        dept_sum["savings_pct"] = dept_sum["total_savings"] / dept_sum["baseline"]
        dept_sum.to_excel(writer, sheet_name="Dept Summary", index=False)

        # Category pillar rate reference sheet
        rate_rows = []
        for cat, rates in CATEGORY_PILLAR_RATES.items():
            row = {"category": cat}
            row.update({MRO_PILLARS[k]["name"]: v for k, v in rates.items()})
            rate_rows.append(row)
        pd.DataFrame(rate_rows).to_excel(writer, sheet_name="Category Pillar Rates", index=False)

    print(f"Saved: {out}\n")
