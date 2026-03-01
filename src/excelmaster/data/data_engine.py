"""Data engine: orchestrates all dataset generators and builds DatasetProfile."""
from __future__ import annotations
import json
from pathlib import Path
import numpy as np
import pandas as pd

from ..models import DatasetProfile, ColumnInfo
from .generators.extreme_load import ExtremeLoadGenerator
from .generators.moderate import ModerateGenerator
from .generators.feature_rich import FeatureRichGenerator
from .generators.sparse import SparseGenerator
from .generators.finance import FinanceGenerator
from .generators.supply_chain import SupplyChainGenerator
from .generators.executive import ExecutiveGenerator
from .generators.hr_admin import HRAdminGenerator
from .generators.marketing import MarketingGenerator

GENERATORS = {
    "extreme_load": ExtremeLoadGenerator,
    "moderate": ModerateGenerator,
    "feature_rich": FeatureRichGenerator,
    "sparse": SparseGenerator,
    "finance": FinanceGenerator,
    "supply_chain": SupplyChainGenerator,
    "executive": ExecutiveGenerator,
    "hr_admin": HRAdminGenerator,
    "marketing": MarketingGenerator,
}

INDUSTRY_MAP = {
    "extreme_load": "Multi-Industry",
    "moderate": "E-Commerce / Retail",
    "feature_rich": "Investment Analytics",
    "sparse": "Organizational Research",
    "finance": "Corporate Finance",
    "supply_chain": "Supply Chain & Logistics",
    "executive": "Executive / Board",
    "hr_admin": "Human Resources",
    "marketing": "Marketing",
}


def generate_dataset(dataset_type: str, output_dir: str | Path = "data") -> Path:
    """Generate a dataset and return the output path."""
    if dataset_type not in GENERATORS:
        raise ValueError(f"Unknown dataset type: {dataset_type}. Choose from: {list(GENERATORS.keys())}")
    gen_cls = GENERATORS[dataset_type]
    gen = gen_cls(output_dir=output_dir)
    print(f"\nGenerating [{dataset_type}] dataset...")
    path = gen.save()
    print(f"  Saved to: {path}")
    return path


def generate_all(output_dir: str | Path = "data") -> dict[str, Path]:
    """Generate all 9 datasets. Returns {type: path} mapping."""
    results = {}
    for ds_type in GENERATORS:
        try:
            path = generate_dataset(ds_type, output_dir)
            results[ds_type] = path
        except Exception as e:
            print(f"  ERROR generating {ds_type}: {e}")
    return results


def _classify_dtype(col: pd.Series) -> str:
    if pd.api.types.is_datetime64_any_dtype(col):
        return "date"
    if pd.api.types.is_bool_dtype(col):
        return "boolean"
    if pd.api.types.is_numeric_dtype(col):
        # Check if it looks like a date represented as int
        return "numeric"
    # Check if object dtype contains dates
    if col.dtype == object:
        sample = col.dropna().head(10)
        try:
            pd.to_datetime(sample, infer_datetime_format=True)
            if sample.str.match(r"\d{4}-\d{2}-\d{2}").all():
                return "date"
        except Exception:
            pass
        # Low cardinality = categorical
        if col.nunique() <= min(50, len(col) * 0.05):
            return "categorical"
        return "text"
    return "text"


def _is_joinable_col(series: pd.Series) -> bool:
    """Return True if the column looks like a join key (text/ID, not a numeric metric)."""
    if pd.api.types.is_bool_dtype(series):
        return False
    if pd.api.types.is_numeric_dtype(series):
        # Numeric columns are only valid keys if they look like IDs
        # (integer-valued, no decimals, name hints at identity)
        name_lower = series.name.lower() if isinstance(series.name, str) else ""
        id_hints = ("_id", "id_", "code", "number", "num", "key", "sku", "zip")
        if any(h in name_lower for h in id_hints):
            return True
        # Reject pure numeric metrics (floats, percentages, amounts, etc.)
        return False
    # Text / object / categorical — valid join key candidate
    return True


def _is_likely_dimension(df: pd.DataFrame) -> bool:
    """Return True if the sheet looks like a dimension/lookup table (mostly text/categorical)."""
    n_text = len(df.select_dtypes(include=["object", "category", "bool"]).columns)
    n_numeric = len(df.select_dtypes("number").columns)
    return n_text >= n_numeric  # dimension tables have more text cols than numbers


def _find_joinable_dims(sheets: dict[str, pd.DataFrame],
                        fact_name: str) -> int:
    """Count how many dimension sheets each candidate fact table can join to."""
    fact_df = sheets[fact_name]
    fact_text_cols = {c for c in fact_df.columns if _is_joinable_col(fact_df[c])}
    count = 0
    for dim_name, dim_df in sheets.items():
        if dim_name == fact_name:
            continue
        if len(dim_df) < 5:
            continue
        # Skip sheets that are bigger AND look like fact tables (mostly numeric)
        if len(dim_df) > len(fact_df) and not _is_likely_dimension(dim_df):
            continue
        dim_text_cols = {c for c in dim_df.columns if _is_joinable_col(dim_df[c])}
        shared = fact_text_cols & dim_text_cols
        if shared:
            count += 1
    return count


def discover_and_join(file_path: str | Path,
                      verbose: bool = True) -> tuple[pd.DataFrame, str, list[str]]:
    """Discover all sheets in an xlsx file, identify relationships, and
    return a unified DataFrame by joining the main fact table with dimension
    tables.

    Returns:
        (unified_df, primary_sheet_name, join_log)
    """
    file_path = Path(file_path)
    xf = pd.ExcelFile(file_path)
    all_sheets = xf.sheet_names

    if len(all_sheets) <= 1:
        name = all_sheets[0]
        df = pd.read_excel(file_path, sheet_name=name)
        return df, name, [f"Single sheet: {name} ({len(df):,} rows)"]

    # Read all sheets
    sheets: dict[str, pd.DataFrame] = {}
    for name in all_sheets:
        sheets[name] = pd.read_excel(file_path, sheet_name=name)

    # Score each sheet as a fact table candidate
    # Fact table = most rows + most numeric columns + bonus for joinable dimensions
    scores = {}
    for name, df in sheets.items():
        n_rows = len(df)
        n_numeric = len(df.select_dtypes("number").columns)
        n_cols = len(df.columns)
        n_joinable = _find_joinable_dims(sheets, name)
        scores[name] = n_rows * 0.5 + n_numeric * 100 + n_cols * 10 + n_joinable * 500
    fact_name = max(scores, key=scores.get)
    fact_df = sheets[fact_name].copy()

    join_log = [
        f"Sheets found: {all_sheets}",
        f"Primary fact table: {fact_name} "
        f"({len(fact_df):,} rows x {len(fact_df.columns)} cols)",
    ]

    # Try to left-join other sheets as dimension/lookup tables
    for dim_name, dim_df in sheets.items():
        if dim_name == fact_name:
            continue

        # Skip very small summary tables (< 5 rows) — they're KPI snapshots
        if len(dim_df) < 5:
            join_log.append(f"  Skip '{dim_name}': too few rows ({len(dim_df)}) — likely summary")
            continue

        # Skip sheets that are bigger AND look like parallel fact tables (mostly numeric)
        if len(dim_df) > len(fact_df) and not _is_likely_dimension(dim_df):
            join_log.append(
                f"  Skip '{dim_name}': {len(dim_df):,} rows — parallel fact table, not a dimension"
            )
            continue

        # Find shared TEXT/ID columns as candidate join keys (never join on numeric metrics)
        shared_cols = sorted(
            c for c in (set(fact_df.columns) & set(dim_df.columns))
            if _is_joinable_col(fact_df[c]) and _is_joinable_col(dim_df[c])
        )
        if not shared_cols:
            join_log.append(f"  Skip '{dim_name}': no shared text/ID columns with {fact_name}")
            continue

        # Pick the best join key
        best_key = None
        best_score = -1
        for col in shared_cols:
            # A good join key: near-unique in the dimension table
            dim_nunique = dim_df[col].nunique()
            dim_len = len(dim_df)
            uniqueness_ratio = dim_nunique / max(dim_len, 1)
            # Accept keys with >= 30% uniqueness (handles composite-key tables like carrier×mode)
            if uniqueness_ratio < 0.3:
                continue

            # Check value overlap
            fact_vals = set(fact_df[col].dropna().astype(str).unique())
            dim_vals = set(dim_df[col].dropna().astype(str).unique())
            overlap = len(fact_vals & dim_vals)
            if overlap == 0:
                continue

            # Score: prefer high overlap + high uniqueness
            overlap_ratio = overlap / max(len(dim_vals), 1)
            score = overlap_ratio * 0.6 + uniqueness_ratio * 0.4
            if score > best_score:
                best_score = score
                best_key = col

        if best_key is None:
            join_log.append(f"  Skip '{dim_name}': no valid join key among {shared_cols}")
            continue

        # Determine new columns the dimension would add
        new_cols = [c for c in dim_df.columns if c not in fact_df.columns]
        if not new_cols:
            join_log.append(f"  Skip '{dim_name}': no new columns to add")
            continue

        # Cap at 12 new columns to avoid bloat
        new_cols = new_cols[:12]

        # Deduplicate dimension on the key to ensure 1:1 or N:1 join
        dim_subset = dim_df[[best_key] + new_cols].drop_duplicates(subset=[best_key])

        # Perform left join
        original_rows = len(fact_df)
        fact_df = fact_df.merge(dim_subset, on=best_key, how="left")

        # Safety: if rows exploded (many-to-many), revert
        if len(fact_df) > original_rows * 1.05:
            fact_df = fact_df.drop(columns=new_cols, errors="ignore")
            fact_df = fact_df.head(original_rows)
            join_log.append(
                f"  Reverted '{dim_name}': join on '{best_key}' caused row explosion "
                f"({original_rows:,} → {len(fact_df):,})"
            )
            continue

        join_log.append(
            f"  Joined '{dim_name}' via '{best_key}' → "
            f"+{len(new_cols)} cols ({', '.join(new_cols[:5])}{'...' if len(new_cols)>5 else ''})"
        )

    join_log.append(
        f"Unified table: {len(fact_df):,} rows x {len(fact_df.columns)} cols"
    )
    return fact_df, fact_name, join_log


def profile_dataset(file_path: str | Path, sheet_name: str = None,
                    industry: str = "", description: str = "") -> DatasetProfile:
    """Read an Excel or CSV file and build a DatasetProfile for LLM consumption.

    For multi-sheet xlsx files, automatically discovers relationships between
    sheets and joins dimension tables into the primary fact table.
    """
    file_path = Path(file_path)

    if file_path.suffix.lower() == ".csv":
        df = pd.read_csv(file_path)
        sheet_name = "Sheet1"
    else:
        xf = pd.ExcelFile(file_path)
        all_sheets = xf.sheet_names

        if len(all_sheets) > 1 and sheet_name is None:
            # Multi-sheet discovery & join
            df, sheet_name, join_log = discover_and_join(file_path, verbose=True)
            for msg in join_log:
                print(f"    {msg}")
        else:
            # Single sheet or explicit sheet requested
            if sheet_name is None:
                sheet_name = all_sheets[0]
            df = pd.read_excel(file_path, sheet_name=sheet_name)

    columns: list[ColumnInfo] = []
    date_cols, numeric_cols, cat_cols = [], [], []

    for col_name in df.columns:
        series = df[col_name]
        dtype = _classify_dtype(series)
        nunique = int(series.nunique(dropna=True))
        null_pct = float(series.isna().mean())
        sample = [str(v) for v in series.dropna().head(5).tolist()]

        try:
            min_v = None if series.isna().all() else (
                str(series.min()) if dtype in ("date", "text", "categorical") else float(series.min()))
            max_v = None if series.isna().all() else (
                str(series.max()) if dtype in ("date", "text", "categorical") else float(series.max()))
        except Exception:
            min_v = max_v = None

        columns.append(ColumnInfo(
            name=str(col_name),
            dtype=dtype,
            unique_values=nunique,
            null_pct=round(null_pct, 3),
            sample_values=sample,
            min_val=min_v,
            max_val=max_v,
        ))
        if dtype == "date":
            date_cols.append(str(col_name))
        elif dtype == "numeric":
            numeric_cols.append(str(col_name))
        elif dtype in ("categorical", "boolean"):
            cat_cols.append(str(col_name))

    # Infer industry from filename if not provided
    if not industry:
        for key, ind in INDUSTRY_MAP.items():
            if key in file_path.stem.lower():
                industry = ind
                break
        else:
            industry = "General"

    return DatasetProfile(
        name=file_path.stem,
        file_path=str(file_path),
        sheet_name=sheet_name,
        rows=len(df),
        columns=columns,
        industry=industry,
        description=description or f"{len(df):,} rows × {len(columns)} columns",
        date_columns=date_cols,
        numeric_columns=numeric_cols,
        categorical_columns=cat_cols,
    )


def profile_to_prompt_text(profile: DatasetProfile, max_cols: int = 30) -> str:
    """Convert a DatasetProfile to a compact text summary for LLM prompts."""
    lines = [
        f"Dataset: {profile.name}",
        f"Industry: {profile.industry}",
        f"Shape: {profile.rows:,} rows × {len(profile.columns)} columns",
        f"Sheet: {profile.sheet_name}",
        "",
        "Columns:",
    ]
    for col in profile.columns[:max_cols]:
        null_str = f" [null:{col.null_pct:.0%}]" if col.null_pct > 0.05 else ""
        sample_str = f" samples={col.sample_values[:3]}" if col.sample_values else ""
        range_str = ""
        if col.min_val is not None and col.max_val is not None and col.dtype == "numeric":
            range_str = f" range=[{col.min_val:.0f}..{col.max_val:.0f}]"
        lines.append(f"  - {col.name} ({col.dtype}, {col.unique_values} unique{null_str}){range_str}{sample_str}")

    if len(profile.columns) > max_cols:
        lines.append(f"  ... and {len(profile.columns) - max_cols} more columns")

    lines += [
        "",
        f"Date columns: {profile.date_columns}",
        f"Numeric columns: {profile.numeric_columns[:15]}",
        f"Categorical columns: {profile.categorical_columns[:10]}",
    ]
    return "\n".join(lines)
