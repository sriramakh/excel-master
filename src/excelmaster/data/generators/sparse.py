"""Sparse dataset: Survey results with high null rates (low data availability)."""
from __future__ import annotations
import numpy as np
import pandas as pd
from .base import BaseGenerator, rng_choice, rng_integers, rng_uniform, date_range, make_ids

DEPARTMENTS = ["Engineering", "Sales", "Marketing", "HR", "Operations", "Finance", "Product", "Design"]
AGE_GROUPS = ["18-24", "25-34", "35-44", "45-54", "55-64", "65+"]
SATISFACTION_SCALE = [1, 2, 3, 4, 5]
ROLES = ["Individual Contributor", "Team Lead", "Manager", "Senior Manager", "Director", "VP", "C-Suite"]
EDUCATION = ["High School", "Associate's", "Bachelor's", "Master's", "PhD", "Other"]


class SparseGenerator(BaseGenerator):
    name = "sparse"
    industry = "Organizational Research"
    description = "Employee survey with 35-60% null rates across 18 columns (low data availability)"

    def generate(self) -> dict[str, pd.DataFrame]:
        n = 400

        df = pd.DataFrame({
            "response_id": make_ids("SRV", 1, n),
            "survey_date": pd.to_datetime(date_range("2024-01-01", "2024-12-15", n)).date,
            "employee_id": [f"EMP{rng_integers(1001, 2000, 1)[0]}" for _ in range(n)],
            "department": rng_choice(DEPARTMENTS, n),
            "role_level": rng_choice(ROLES, n),
            "age_group": rng_choice(AGE_GROUPS, n),
            "tenure_years": rng_integers(0, 25, n).astype(float),
            "gender": rng_choice(["Male", "Female", "Non-Binary", "Prefer not to say"], n,
                                   p=[0.44, 0.44, 0.06, 0.06]),
            "education_level": rng_choice(EDUCATION, n),
            "remote_work_days": rng_integers(0, 5, n).astype(float),

            # High-null satisfaction questions (40-60% null)
            "overall_satisfaction": rng_choice(SATISFACTION_SCALE, n).astype(float),
            "manager_effectiveness": rng_choice(SATISFACTION_SCALE, n).astype(float),
            "team_collaboration": rng_choice(SATISFACTION_SCALE, n).astype(float),
            "growth_opportunities": rng_choice(SATISFACTION_SCALE, n).astype(float),
            "work_life_balance": rng_choice(SATISFACTION_SCALE, n).astype(float),
            "compensation_fairness": rng_choice(SATISFACTION_SCALE, n).astype(float),
            "company_direction": rng_choice(SATISFACTION_SCALE, n).astype(float),
            "tools_and_resources": rng_choice(SATISFACTION_SCALE, n).astype(float),

            # eNPS
            "enps_score": rng_integers(0, 10, n).astype(float),

            # Free-text (sparse)
            "top_challenge": rng_choice(
                ["Work overload", "Lack of clarity", "Communication gaps",
                 "Limited growth", "Poor tools", "Team conflict", "Unclear strategy", ""],
                n, p=[0.12, 0.10, 0.10, 0.12, 0.09, 0.07, 0.10, 0.30]
            ),
            "improvement_suggestion": rng_choice(
                ["Better 1:1s", "More training", "Flexible hours",
                 "Clearer OKRs", "Better tools", "More headcount", ""],
                n, p=[0.10, 0.13, 0.12, 0.11, 0.09, 0.10, 0.35]
            ),
        })

        # Introduce structured null patterns
        null_config = {
            "gender": 0.08,
            "age_group": 0.12,
            "tenure_years": 0.15,
            "education_level": 0.20,
            "remote_work_days": 0.18,
            "overall_satisfaction": 0.35,
            "manager_effectiveness": 0.42,
            "team_collaboration": 0.38,
            "growth_opportunities": 0.45,
            "work_life_balance": 0.40,
            "compensation_fairness": 0.52,
            "company_direction": 0.48,
            "tools_and_resources": 0.40,
            "enps_score": 0.30,
            "top_challenge": 0.30,
            "improvement_suggestion": 0.35,
        }
        df = self._introduce_nulls(df, list(null_config.keys()),
                                    null_pct=0.0)  # handled column by column below
        for col, npct in null_config.items():
            mask = np.random.random(n) < npct
            df.loc[mask, col] = np.nan

        # Summary with filled data
        dept_summary = df.groupby("department").agg(
            responses=("response_id", "count"),
            avg_satisfaction=("overall_satisfaction", lambda x: x.mean()),
            avg_enps=("enps_score", lambda x: x.mean()),
            response_rate=("overall_satisfaction", lambda x: x.notna().mean() * 100),
        ).round(2).reset_index()

        return {
            "Survey_Responses": df,
            "Department_Summary": dept_summary,
        }
