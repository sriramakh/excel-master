"""Base class for all dataset generators."""
from __future__ import annotations
import numpy as np
import pandas as pd
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Any


RNG = np.random.default_rng(42)


def rng_choice(choices: list, size: int, p: list[float] | None = None) -> np.ndarray:
    return RNG.choice(choices, size=size, p=p)


def rng_uniform(low: float, high: float, size: int) -> np.ndarray:
    return RNG.uniform(low, high, size)


def rng_integers(low: int, high: int, size: int) -> np.ndarray:
    return RNG.integers(low, high, size)


def rng_normal(mean: float, std: float, size: int) -> np.ndarray:
    return RNG.normal(mean, std, size)


def date_range(start: str, end: str, size: int) -> np.ndarray:
    start_ts = pd.Timestamp(start).value // 10**9
    end_ts = pd.Timestamp(end).value // 10**9
    timestamps = rng_integers(int(start_ts), int(end_ts), size)
    return pd.to_datetime(timestamps, unit="s").normalize()


def make_ids(prefix: str, start: int, count: int) -> list[str]:
    return [f"{prefix}{i:05d}" for i in range(start, start + count)]


class BaseGenerator(ABC):
    """Abstract base class for all dataset generators."""

    name: str = "base"
    industry: str = "General"
    description: str = ""

    def __init__(self, output_dir: str | Path = "data"):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

    @abstractmethod
    def generate(self) -> dict[str, pd.DataFrame]:
        """Generate and return dict of {sheet_name: DataFrame}."""
        ...

    def save(self, filename: str | None = None) -> Path:
        """Generate data and save to Excel. Returns output path."""
        sheets = self.generate()
        fname = filename or f"{self.name}.xlsx"
        out_path = self.output_dir / fname

        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            for sheet_name, df in sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        return out_path

    def _add_noise(self, series: pd.Series, pct: float = 0.05) -> pd.Series:
        """Add small random noise to a numeric series."""
        noise = rng_normal(0, series.std() * pct, len(series))
        return series + noise

    def _introduce_nulls(self, df: pd.DataFrame, columns: list[str], null_pct: float) -> pd.DataFrame:
        """Randomly set null_pct of values to NaN in given columns."""
        df = df.copy()
        for col in columns:
            mask = RNG.random(len(df)) < null_pct
            df.loc[mask, col] = np.nan
        return df
