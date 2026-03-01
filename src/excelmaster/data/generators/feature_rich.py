"""Feature-rich dataset: Complex financial analysis with 22+ calculated columns."""
from __future__ import annotations
import numpy as np
import pandas as pd
from .base import BaseGenerator, rng_choice, rng_uniform, rng_integers, rng_normal, date_range

SECTORS = ["Technology", "Healthcare", "Finance", "Energy", "Consumer", "Industrials", "Materials"]
PORTFOLIOS = ["Growth Fund", "Value Fund", "Balanced Fund", "Income Fund", "ESG Fund"]
RISK_RATINGS = ["AAA", "AA+", "AA", "AA-", "A+", "A", "A-", "BBB+", "BBB", "BBB-", "BB+", "BB"]


class FeatureRichGenerator(BaseGenerator):
    name = "feature_rich"
    industry = "Investment Analytics"
    description = "Multi-asset portfolio with 22 calculated metrics per instrument"

    def generate(self) -> dict[str, pd.DataFrame]:
        n = 800
        np.random.seed(42)

        sectors = rng_choice(SECTORS, n)
        portfolios = rng_choice(PORTFOLIOS, n)
        risk = rng_choice(RISK_RATINGS, n)

        # Core price data
        initial_price = rng_uniform(10, 500, n).round(2)
        returns_1d = rng_normal(0.0005, 0.015, n)
        returns_5d = rng_normal(0.002, 0.035, n)
        returns_1m = rng_normal(0.01, 0.07, n)
        returns_3m = rng_normal(0.03, 0.12, n)
        returns_ytd = rng_normal(0.08, 0.18, n)
        returns_1y = rng_normal(0.10, 0.22, n)

        current_price = (initial_price * (1 + returns_1y)).round(2)
        market_cap = rng_uniform(100e6, 500e9, n).round(0)
        shares_outstanding = (market_cap / current_price).round(0)
        volume = rng_integers(100000, 50000000, n)
        avg_volume = (volume * rng_uniform(0.7, 1.5, n)).round(0)

        # Volatility & risk
        volatility = rng_uniform(0.08, 0.65, n).round(4)
        beta = rng_normal(1.0, 0.5, n).clip(0.1, 2.5).round(2)
        sharpe = rng_normal(1.0, 0.8, n).round(2)
        sortino = (sharpe * rng_uniform(1.1, 1.8, n)).round(2)
        max_drawdown = rng_uniform(-0.70, -0.05, n).round(4)
        var_95 = (volatility * 1.65 * initial_price).round(2)

        # Valuation
        pe_ratio = np.where(returns_ytd > 0, rng_uniform(8, 60, n), np.nan).round(2)
        pb_ratio = rng_uniform(0.5, 15, n).round(2)
        ps_ratio = rng_uniform(0.3, 25, n).round(2)
        ev_ebitda = rng_uniform(4, 40, n).round(2)
        dividend_yield = np.where(rng_choice([True, False], n, p=[0.4, 0.6]),
                                   rng_uniform(0.005, 0.08, n).round(4), 0.0)
        payout_ratio = np.where(dividend_yield > 0, rng_uniform(0.15, 0.85, n).round(2), 0.0)

        # ESG
        esg_score = rng_integers(20, 95, n)
        env_score = rng_integers(15, 100, n)
        social_score = rng_integers(20, 98, n)
        governance_score = rng_integers(25, 99, n)

        df = pd.DataFrame({
            "ticker": [f"{rng_choice(list('ABCDEFGHIJKLMNOPQRSTUVWXYZ'), 1)[0]}{rng_choice(list('ABCDEFGHIJKLMNOPQRSTUVWXYZ'), 1)[0]}{rng_choice(list('ABCDEFGHIJKLMNOPQRSTUVWXYZ'), 1)[0]}{rng_choice(list('ABCDEFGHIJKLMNOPQRSTUVWXYZ'), 1)[0]}" for _ in range(n)],
            "company_name": [f"Company {i+1} {rng_choice(['Inc', 'Corp', 'Ltd', 'PLC', 'AG'], 1)[0]}" for i in range(n)],
            "sector": sectors,
            "portfolio": portfolios,
            "risk_rating": risk,
            "as_of_date": date_range("2024-01-01", "2024-12-31", n),
            "initial_price": initial_price,
            "current_price": current_price,
            "52w_high": (current_price * rng_uniform(1.0, 1.5, n)).round(2),
            "52w_low": (current_price * rng_uniform(0.5, 1.0, n)).round(2),
            "market_cap": market_cap,
            "shares_outstanding": shares_outstanding,
            "daily_volume": volume,
            "avg_volume_30d": avg_volume.astype(int),
            "return_1d_pct": (returns_1d * 100).round(2),
            "return_5d_pct": (returns_5d * 100).round(2),
            "return_1m_pct": (returns_1m * 100).round(2),
            "return_3m_pct": (returns_3m * 100).round(2),
            "return_ytd_pct": (returns_ytd * 100).round(2),
            "return_1y_pct": (returns_1y * 100).round(2),
            "volatility_annualized": volatility,
            "beta": beta,
            "sharpe_ratio": sharpe,
            "sortino_ratio": sortino,
            "max_drawdown_pct": (max_drawdown * 100).round(2),
            "var_95_usd": var_95,
            "pe_ratio": pe_ratio,
            "pb_ratio": pb_ratio,
            "ps_ratio": ps_ratio,
            "ev_ebitda": ev_ebitda,
            "dividend_yield_pct": (dividend_yield * 100).round(2),
            "payout_ratio": payout_ratio,
            "esg_score": esg_score,
            "environmental_score": env_score,
            "social_score": social_score,
            "governance_score": governance_score,
            "analyst_rating": rng_choice(["Strong Buy", "Buy", "Hold", "Sell", "Strong Sell"], n,
                                          p=[0.20, 0.30, 0.30, 0.12, 0.08]),
            "target_price": (current_price * rng_uniform(0.8, 1.5, n)).round(2),
            "upside_potential_pct": None,
        })
        df["upside_potential_pct"] = ((df["target_price"] - df["current_price"]) / df["current_price"] * 100).round(2)
        df["as_of_date"] = pd.to_datetime(df["as_of_date"]).dt.date

        # Sector summary
        sector_summary = df.groupby("sector").agg(
            count=("ticker", "count"),
            avg_return_ytd=("return_ytd_pct", "mean"),
            avg_pe=("pe_ratio", "mean"),
            avg_beta=("beta", "mean"),
            avg_esg=("esg_score", "mean"),
            total_market_cap=("market_cap", "sum"),
        ).round(2).reset_index()

        # Portfolio summary
        portfolio_summary = df.groupby("portfolio").agg(
            positions=("ticker", "count"),
            avg_return_1y=("return_1y_pct", "mean"),
            avg_sharpe=("sharpe_ratio", "mean"),
            avg_volatility=("volatility_annualized", "mean"),
            avg_max_drawdown=("max_drawdown_pct", "mean"),
        ).round(2).reset_index()

        return {
            "Holdings": df,
            "Sector_Analysis": sector_summary,
            "Portfolio_Summary": portfolio_summary,
        }
