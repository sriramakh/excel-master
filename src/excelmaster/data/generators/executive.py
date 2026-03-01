"""Executive metrics dataset: Board-level KPIs, OKRs, strategic scorecard."""
from __future__ import annotations
import numpy as np
import pandas as pd
from .base import BaseGenerator, rng_choice, rng_uniform, rng_integers, rng_normal

STRATEGIC_PILLARS = ["Revenue Growth", "Customer Excellence", "Operational Efficiency",
                      "People & Culture", "Product Innovation", "Market Expansion"]
BUSINESS_UNITS = ["North America", "Europe", "APAC", "Latin America", "Global Services"]
EXEC_OWNERS = ["CEO", "CFO", "CRO", "CMO", "CPO", "CHRO", "COO", "CTO"]
METRIC_FREQUENCY = ["Monthly", "Quarterly", "Annual"]
TREND_DIRECTIONS = ["Up", "Down", "Flat"]
HEALTH_STATUS = ["On Track", "At Risk", "Off Track", "Exceeded"]


def _strategic_kpis() -> pd.DataFrame:
    """Board-level KPI scorecard."""
    kpis = [
        # Financial
        ("Annual Recurring Revenue", "ARR", 85000000, 82000000, 78500000, "USD", "CFO", "Revenue Growth"),
        ("Total Revenue", "Revenue", 95000000, 91000000, 84000000, "USD", "CFO", "Revenue Growth"),
        ("Gross Margin", "GM%", 71.2, 70.0, 68.5, "%", "CFO", "Operational Efficiency"),
        ("EBITDA", "EBITDA", 18500000, 17000000, 14200000, "USD", "CFO", "Operational Efficiency"),
        ("Net Revenue Retention", "NRR", 118, 115, 112, "%", "CRO", "Customer Excellence"),
        ("Customer Acquisition Cost", "CAC", 8200, 9500, 10200, "USD", "CMO", "Revenue Growth"),
        ("LTV:CAC Ratio", "LTV:CAC", 4.8, 4.2, 3.9, "ratio", "CFO", "Customer Excellence"),
        ("Cash Runway", "Runway", 28, 24, 22, "months", "CFO", "Operational Efficiency"),
        # Customer
        ("Total Customers", "Customers", 3847, 3600, 3250, "count", "CRO", "Customer Excellence"),
        ("Net Promoter Score", "NPS", 68, 62, 55, "score", "CPO", "Customer Excellence"),
        ("Customer Churn Rate", "Churn", 4.2, 5.1, 6.8, "%", "CRO", "Customer Excellence"),
        ("CSAT Score", "CSAT", 4.6, 4.4, 4.2, "1-5", "CPO", "Customer Excellence"),
        # Product
        ("Product Adoption Rate", "Adoption", 78, 72, 65, "%", "CPO", "Product Innovation"),
        ("Monthly Active Users", "MAU", 142000, 125000, 108000, "count", "CPO", "Product Innovation"),
        ("Feature Release Velocity", "Releases", 24, 20, 18, "per year", "CTO", "Product Innovation"),
        ("Bug Resolution Time", "MTTR", 4.2, 5.8, 7.1, "hours", "CTO", "Product Innovation"),
        # People
        ("Employee Headcount", "HC", 892, 850, 780, "employees", "CHRO", "People & Culture"),
        ("Employee NPS", "eNPS", 42, 38, 31, "score", "CHRO", "People & Culture"),
        ("Voluntary Turnover", "Turnover", 11.2, 13.5, 15.8, "%", "CHRO", "People & Culture"),
        ("Time to Fill", "TTF", 38, 45, 52, "days", "CHRO", "People & Culture"),
        ("Revenue per FTE", "Rev/FTE", 106500, 107000, 107700, "USD", "CFO", "Operational Efficiency"),
        # Market
        ("Market Share", "MktShare", 8.4, 7.9, 7.2, "%", "CMO", "Market Expansion"),
        ("Brand Awareness", "Awareness", 34, 31, 27, "%", "CMO", "Market Expansion"),
        ("Pipeline Coverage", "Pipeline", 3.8, 3.5, 3.1, "ratio", "CRO", "Revenue Growth"),
        ("Win Rate", "WinRate", 28.5, 26.2, 24.8, "%", "CRO", "Revenue Growth"),
    ]

    rows = []
    for name, abbr, actual, target, py, unit, owner, pillar in kpis:
        vs_target = actual - target
        vs_target_pct = (vs_target / target * 100) if target != 0 else 0
        vs_py_pct = ((actual - py) / py * 100) if py != 0 else 0
        # Status logic
        better_up = unit not in ["%", "ratio"] or name not in ["Customer Churn Rate", "CAC",
                                                                  "Voluntary Turnover", "Bug Resolution Time",
                                                                  "Time to Fill"]
        on_track = (actual >= target) if better_up else (actual <= target)
        exceeded = (actual >= target * 1.05) if better_up else (actual <= target * 0.95)
        status = "Exceeded" if exceeded else ("On Track" if on_track else
                  "At Risk" if abs(vs_target_pct) < 10 else "Off Track")

        rows.append({
            "kpi_name": name,
            "abbreviation": abbr,
            "pillar": pillar,
            "owner": owner,
            "actual_value": actual,
            "target_value": target,
            "prior_year_value": py,
            "unit": unit,
            "variance_vs_target": round(vs_target, 2),
            "variance_vs_target_pct": round(vs_target_pct, 1),
            "yoy_change_pct": round(vs_py_pct, 1),
            "status": status,
            "trend": "Up" if vs_py_pct > 2 else ("Down" if vs_py_pct < -2 else "Flat"),
            "reporting_frequency": "Quarterly",
            "last_updated": "2024-12-31",
        })
    return pd.DataFrame(rows)


def _okr_tracker() -> pd.DataFrame:
    """OKR tracking for current year."""
    objectives = [
        ("Achieve $95M total revenue", "Revenue Growth", "CFO"),
        ("Grow customer base to 4,000", "Customer Excellence", "CRO"),
        ("Launch 3 major product features", "Product Innovation", "CPO"),
        ("Reduce churn to below 4%", "Customer Excellence", "CRO"),
        ("Scale team to 900 employees", "People & Culture", "CHRO"),
        ("Enter 2 new markets", "Market Expansion", "CMO"),
        ("Achieve NPS of 70", "Customer Excellence", "CPO"),
        ("Improve gross margin to 73%", "Operational Efficiency", "CFO"),
    ]
    rows = []
    for obj, pillar, owner in objectives:
        n_krs = rng_integers(3, 5, 1)[0]
        obj_progress = rng_uniform(45, 105, 1)[0].round(1)
        for kr_idx in range(n_krs):
            progress = rng_uniform(20, 110, 1)[0].round(1)
            rows.append({
                "objective": obj,
                "pillar": pillar,
                "objective_owner": owner,
                "key_result": f"KR{kr_idx+1}: " + rng_choice([
                    "Increase metric by X%", "Achieve Y goal by Q4",
                    "Launch initiative Z", "Reduce X to Y level",
                    "Complete X milestone", "Hire N people",
                ], 1)[0],
                "kr_owner": rng_choice(["Director", "VP", "Manager"], 1)[0] + f" {rng_integers(1, 10, 1)[0]}",
                "start_date": "2024-01-01",
                "end_date": "2024-12-31",
                "progress_pct": min(progress, 100),
                "is_complete": progress >= 100,
                "confidence_level": rng_choice(["High", "Medium", "Low"], 1, p=[0.5, 0.35, 0.15])[0],
                "last_checkin_date": f"2024-{rng_integers(1, 12, 1)[0]:02d}-{rng_integers(1, 28, 1)[0]:02d}",
                "objective_progress_pct": min(obj_progress, 100),
                "quarter": rng_choice(["Q1", "Q2", "Q3", "Q4"], 1)[0],
                "blockers": rng_choice(["None", "Resource constraints", "Dependencies", "Market conditions"], 1)[0],
            })
    return pd.DataFrame(rows)


def _quarterly_trend() -> pd.DataFrame:
    """8-quarter trend of top metrics."""
    quarters = ["Q1'23", "Q2'23", "Q3'23", "Q4'23", "Q1'24", "Q2'24", "Q3'24", "Q4'24"]
    base_arr = 60000000
    base_customers = 2800
    base_nps = 51

    rows = []
    for i, q in enumerate(quarters):
        growth = 1 + (0.04 + rng_normal(0, 0.01, 1)[0])
        arr = base_arr * (growth ** i)
        customers = int(base_customers * (1.04 ** i) + rng_integers(-50, 50, 1)[0])
        rows.append({
            "quarter": q,
            "arr_usd": round(arr, 0),
            "total_revenue_usd": round(arr * 1.12, 0),
            "gross_margin_pct": round(64 + i * 0.9 + rng_normal(0, 0.5, 1)[0], 1),
            "ebitda_pct": round(10 + i * 1.2 + rng_normal(0, 0.8, 1)[0], 1),
            "total_customers": customers,
            "nrr_pct": round(108 + i * 1.3 + rng_normal(0, 1, 1)[0], 1),
            "churn_pct": round(7.5 - i * 0.4 + rng_normal(0, 0.3, 1)[0], 1),
            "nps": round(base_nps + i * 2.2 + rng_integers(-3, 3, 1)[0], 0),
            "employee_count": round(620 + i * 34 + rng_integers(-10, 10, 1)[0], 0),
            "enps": round(28 + i * 1.8 + rng_integers(-4, 4, 1)[0], 0),
            "cac_usd": round(12000 - i * 480 + rng_integers(-200, 200, 1)[0], 0),
            "mau": round(85000 + i * 7200 + rng_integers(-2000, 2000, 1)[0], 0),
        })
    return pd.DataFrame(rows)


class ExecutiveGenerator(BaseGenerator):
    name = "executive"
    industry = "Executive / Board"
    description = "Strategic KPI scorecard, OKR tracker, and 8-quarter trend for C-suite reporting"

    def generate(self) -> dict[str, pd.DataFrame]:
        print("  Generating Strategic KPI Scorecard...")
        kpis = _strategic_kpis()
        print("  Generating OKR Tracker...")
        okrs = _okr_tracker()
        print("  Generating Quarterly Trend...")
        trend = _quarterly_trend()

        # Pillar summary
        pillar_summary = kpis.groupby("pillar").agg(
            total_kpis=("kpi_name", "count"),
            on_track=("status", lambda x: (x == "On Track").sum()),
            exceeded=("status", lambda x: (x == "Exceeded").sum()),
            at_risk=("status", lambda x: (x == "At Risk").sum()),
            off_track=("status", lambda x: (x == "Off Track").sum()),
        ).reset_index()
        pillar_summary["health_pct"] = (
            (pillar_summary["on_track"] + pillar_summary["exceeded"]) /
            pillar_summary["total_kpis"] * 100
        ).round(1)

        return {
            "KPI_Scorecard": kpis,
            "OKR_Tracker": okrs,
            "Quarterly_Trend": trend,
            "Pillar_Summary": pillar_summary,
        }
