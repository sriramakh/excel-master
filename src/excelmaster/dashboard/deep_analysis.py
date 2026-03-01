"""Deep Analysis: pre-compute statistics, build LLM prompt, parse response."""
from __future__ import annotations

import json
from typing import Any

import numpy as np
import pandas as pd

from ..models import (
    DatasetProfile, DashboardConfig, DeepAnalysis,
    CorrelationInsight, OutlierInsight, PerformerEntry, TrendInsight,
)


# ── Pre-computation ──────────────────────────────────────────────────────────

def compute_deep_stats(df: pd.DataFrame, profile: DatasetProfile) -> dict[str, Any]:
    """Compute aggregated statistics from the dataframe for LLM interpretation.

    Returns a dict with sections: data_quality, numeric_stats, categorical_dist,
    correlations, trends, dimensional.
    """
    stats: dict[str, Any] = {}

    # ── Data quality ──────────────────────────────────────────────────────
    try:
        total_cells = df.shape[0] * df.shape[1]
        null_cells = int(df.isnull().sum().sum())
        dup_rows = int(df.duplicated().sum())
        col_types = {}
        for ci in profile.columns:
            col_types[ci.dtype] = col_types.get(ci.dtype, 0) + 1

        col_completeness = {}
        for c in df.columns[:30]:
            null_pct = float(df[c].isnull().mean() * 100)
            col_completeness[c] = round(null_pct, 1)

        stats["data_quality"] = {
            "total_rows": df.shape[0],
            "total_columns": df.shape[1],
            "total_cells": total_cells,
            "null_cells": null_cells,
            "null_pct": round(null_cells / total_cells * 100, 2) if total_cells else 0,
            "duplicate_rows": dup_rows,
            "duplicate_pct": round(dup_rows / df.shape[0] * 100, 2) if df.shape[0] else 0,
            "column_type_distribution": col_types,
            "column_completeness": col_completeness,
        }
    except Exception:
        stats["data_quality"] = {"total_rows": df.shape[0], "total_columns": df.shape[1]}

    # ── Numeric stats ─────────────────────────────────────────────────────
    try:
        num_cols = profile.numeric_columns[:15]
        num_stats = {}
        for c in num_cols:
            if c not in df.columns:
                continue
            s = df[c].dropna()
            if s.empty:
                continue
            q1 = float(s.quantile(0.25))
            q3 = float(s.quantile(0.75))
            iqr = q3 - q1
            lower = q1 - 1.5 * iqr
            upper = q3 + 1.5 * iqr
            outlier_count = int(((s < lower) | (s > upper)).sum())
            num_stats[c] = {
                "mean": round(float(s.mean()), 2),
                "median": round(float(s.median()), 2),
                "std": round(float(s.std()), 2),
                "min": round(float(s.min()), 2),
                "max": round(float(s.max()), 2),
                "skew": round(float(s.skew()), 2),
                "q1": round(q1, 2),
                "q3": round(q3, 2),
                "outlier_count": outlier_count,
                "outlier_pct": round(outlier_count / len(s) * 100, 2),
            }
        stats["numeric_stats"] = num_stats
    except Exception:
        stats["numeric_stats"] = {}

    # ── Categorical distributions ─────────────────────────────────────────
    try:
        cat_cols = profile.categorical_columns[:10]
        cat_dist = {}
        for c in cat_cols:
            if c not in df.columns:
                continue
            vc = df[c].value_counts().head(10)
            cat_dist[c] = {
                "unique_count": int(df[c].nunique()),
                "top_values": {str(k): int(v) for k, v in vc.items()},
            }
        stats["categorical_distributions"] = cat_dist
    except Exception:
        stats["categorical_distributions"] = {}

    # ── Correlations ──────────────────────────────────────────────────────
    try:
        num_df = df[profile.numeric_columns[:15]].dropna()
        if len(num_df.columns) >= 2 and len(num_df) >= 5:
            corr = num_df.corr()
            pairs = []
            seen = set()
            for i, ca in enumerate(corr.columns):
                for j, cb in enumerate(corr.columns):
                    if i >= j:
                        continue
                    r = corr.iloc[i, j]
                    if abs(r) > 0.3 and (ca, cb) not in seen:
                        seen.add((ca, cb))
                        pairs.append({"col_a": ca, "col_b": cb,
                                      "coefficient": round(float(r), 3)})
            pairs.sort(key=lambda x: abs(x["coefficient"]), reverse=True)
            stats["correlations"] = pairs[:15]
        else:
            stats["correlations"] = []
    except Exception:
        stats["correlations"] = []

    # ── Trends (period-over-period) ───────────────────────────────────────
    try:
        date_cols = profile.date_columns[:2]
        trends = []
        for dc in date_cols:
            if dc not in df.columns:
                continue
            ds = pd.to_datetime(df[dc], errors="coerce").dropna()
            if ds.empty:
                continue
            for nc in profile.numeric_columns[:5]:
                if nc not in df.columns:
                    continue
                tmp = df[[dc, nc]].copy()
                tmp[dc] = pd.to_datetime(tmp[dc], errors="coerce")
                tmp = tmp.dropna()
                if len(tmp) < 4:
                    continue
                tmp = tmp.sort_values(dc)
                half = len(tmp) // 2
                first_half = tmp[nc].iloc[:half].mean()
                second_half = tmp[nc].iloc[half:].mean()
                if first_half and abs(first_half) > 0.001:
                    pct = round((second_half - first_half) / abs(first_half) * 100, 1)
                    direction = "up" if pct > 2 else ("down" if pct < -2 else "flat")
                    trends.append({
                        "column": nc, "date_column": dc,
                        "direction": direction, "pct_change": pct,
                        "first_half_avg": round(float(first_half), 2),
                        "second_half_avg": round(float(second_half), 2),
                    })
        stats["trends"] = trends
    except Exception:
        stats["trends"] = []

    # ── Dimensional (top/bottom groups) ───────────────────────────────────
    try:
        dim_col = None
        for c in profile.categorical_columns:
            if c in df.columns and 2 <= df[c].nunique() <= 50:
                dim_col = c
                break
        if dim_col and profile.numeric_columns:
            metric_col = profile.numeric_columns[0]
            if metric_col in df.columns:
                grouped = df.groupby(dim_col)[metric_col].sum().sort_values(ascending=False)
                top5 = [{"dimension_value": str(k), "metric_value": round(float(v), 2),
                          "metric_column": metric_col}
                         for k, v in grouped.head(5).items()]
                bot5 = [{"dimension_value": str(k), "metric_value": round(float(v), 2),
                          "metric_column": metric_col}
                         for k, v in grouped.tail(5).items()]
                stats["dimensional"] = {
                    "dimension_column": dim_col,
                    "metric_column": metric_col,
                    "top_5": top5,
                    "bottom_5": bot5,
                }
            else:
                stats["dimensional"] = {}
        else:
            stats["dimensional"] = {}
    except Exception:
        stats["dimensional"] = {}

    return stats


# ── LLM Prompt ────────────────────────────────────────────────────────────────

def build_analysis_prompt(
    stats: dict[str, Any],
    profile: DatasetProfile,
    config: DashboardConfig,
) -> tuple[str, str]:
    """Build (system_prompt, user_prompt) for the deep analysis LLM call."""

    system_prompt = """You are a senior data analyst producing a Deep Analysis report for an executive dashboard.

Given pre-computed statistics about a dataset, produce a comprehensive analysis in JSON format.

Your output MUST be a valid JSON object with these exact keys:

{
  "executive_summary": "<3-sentence high-level overview of the dataset and its key story>",
  "key_findings": ["<finding 1>", "<finding 2>", ... ],  // 4-6 key findings
  "data_quality_score": <0-100 integer>,
  "data_quality_notes": ["<note 1>", ...],  // 2-4 quality observations
  "distribution_insights": ["<insight 1>", ...],  // 2-4 statistical observations
  "correlation_insights": [
    {"col_a": "<col>", "col_b": "<col>", "coefficient": <float>, "interpretation": "<text>"},
    ...
  ],
  "outlier_insights": [
    {"column": "<col>", "count": <int>, "pct": <float>, "description": "<text>"},
    ...
  ],
  "top_performers": [
    {"dimension_value": "<val>", "metric_value": <float>, "metric_column": "<col>"},
    ...
  ],
  "bottom_performers": [
    {"dimension_value": "<val>", "metric_value": <float>, "metric_column": "<col>"},
    ...
  ],
  "dimension_analysis": "<paragraph analyzing the primary grouping dimension>",
  "trend_insights": [
    {"column": "<col>", "direction": "up|down|flat", "description": "<text>", "pct_change": <float>},
    ...
  ],
  "trend_summary": "<paragraph summarizing temporal patterns>",
  "near_term_outlook": "<1-2 sentences on immediate outlook>",
  "long_term_outlook": "<1-2 sentences on long-term trajectory>",
  "recommendations": ["<actionable rec 1>", ...],  // 4-6 recommendations
  "industry_context": "<paragraph placing findings in industry context>"
}

Rules:
- Use EXACT column names from the statistics provided
- All numeric values should be rounded to 2 decimal places
- Findings and recommendations must be specific and data-driven, referencing actual numbers
- If trend data is missing or empty, set trend_insights to [] and trend_summary to "Insufficient temporal data for trend analysis."
- If correlation data is empty, set correlation_insights to [] and add a note in distribution_insights
- data_quality_score: 90-100 = excellent, 70-89 = good, 50-69 = fair, below 50 = poor
- Output ONLY the JSON object, no markdown fences or extra text"""

    # Build the user prompt with actual stats
    profile_summary = {
        "dataset_name": profile.name,
        "industry": profile.industry or config.title,
        "rows": profile.rows,
        "columns": len(profile.columns),
        "numeric_columns": profile.numeric_columns[:15],
        "categorical_columns": profile.categorical_columns[:10],
        "date_columns": profile.date_columns,
        "dashboard_title": config.title,
    }

    user_prompt = f"""Analyze this dataset and produce the Deep Analysis JSON.

DATASET PROFILE:
{json.dumps(profile_summary, indent=2)}

PRE-COMPUTED STATISTICS:
{json.dumps(stats, indent=2, default=str)}

Produce the analysis JSON now."""

    return system_prompt, user_prompt


# ── Safe Parsing ──────────────────────────────────────────────────────────────

def safe_parse_analysis(raw: dict[str, Any]) -> DeepAnalysis:
    """Parse raw LLM dict into DeepAnalysis model with fallbacks for missing fields."""
    try:
        # Parse correlation insights
        corr_insights = []
        for ci in raw.get("correlation_insights", []):
            if isinstance(ci, dict):
                corr_insights.append(CorrelationInsight(**{
                    k: ci.get(k, "") for k in ["col_a", "col_b", "coefficient", "interpretation"]
                }))

        # Parse outlier insights
        outlier_insights = []
        for oi in raw.get("outlier_insights", []):
            if isinstance(oi, dict):
                outlier_insights.append(OutlierInsight(**{
                    k: oi.get(k, 0 if k in ("count", "pct") else "")
                    for k in ["column", "count", "pct", "description"]
                }))

        # Parse performers
        top_perf = []
        for p in raw.get("top_performers", []):
            if isinstance(p, dict):
                top_perf.append(PerformerEntry(
                    dimension_value=str(p.get("dimension_value", "")),
                    metric_value=float(p.get("metric_value", 0)),
                    metric_column=str(p.get("metric_column", "")),
                ))

        bot_perf = []
        for p in raw.get("bottom_performers", []):
            if isinstance(p, dict):
                bot_perf.append(PerformerEntry(
                    dimension_value=str(p.get("dimension_value", "")),
                    metric_value=float(p.get("metric_value", 0)),
                    metric_column=str(p.get("metric_column", "")),
                ))

        # Parse trend insights
        trends = []
        for ti in raw.get("trend_insights", []):
            if isinstance(ti, dict):
                trends.append(TrendInsight(
                    column=str(ti.get("column", "")),
                    direction=str(ti.get("direction", "flat")),
                    description=str(ti.get("description", "")),
                    pct_change=float(ti.get("pct_change", 0)),
                ))

        return DeepAnalysis(
            executive_summary=str(raw.get("executive_summary", "")),
            key_findings=_str_list(raw.get("key_findings", [])),
            data_quality_score=int(raw.get("data_quality_score", 0)),
            data_quality_notes=_str_list(raw.get("data_quality_notes", [])),
            distribution_insights=_str_list(raw.get("distribution_insights", [])),
            correlation_insights=corr_insights,
            outlier_insights=outlier_insights,
            top_performers=top_perf,
            bottom_performers=bot_perf,
            dimension_analysis=str(raw.get("dimension_analysis", "")),
            trend_insights=trends,
            trend_summary=str(raw.get("trend_summary", "")),
            near_term_outlook=str(raw.get("near_term_outlook", "")),
            long_term_outlook=str(raw.get("long_term_outlook", "")),
            recommendations=_str_list(raw.get("recommendations", [])),
            industry_context=str(raw.get("industry_context", "")),
        )
    except Exception:
        return DeepAnalysis()


def _str_list(val: Any) -> list[str]:
    """Coerce a value to a list of strings."""
    if isinstance(val, list):
        return [str(v) for v in val]
    return []
