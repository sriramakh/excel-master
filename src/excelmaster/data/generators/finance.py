"""Finance industry dataset: P&L, budget, cashflow, accounts."""
from __future__ import annotations
import numpy as np
import pandas as pd
from .base import BaseGenerator, rng_choice, rng_uniform, rng_integers, rng_normal, date_range, make_ids

COST_CENTERS = ["Sales", "Marketing", "Engineering", "G&A", "R&D", "Operations", "Customer Success"]
EXPENSE_TYPES = ["Personnel", "Technology", "Facilities", "Travel", "Marketing Spend",
                  "Professional Services", "Depreciation", "Other"]
REVENUE_LINES = ["Product Sales", "Subscription Revenue", "Professional Services",
                  "Support & Maintenance", "Partner Revenue", "License Fees"]
ENTITIES = ["HQ - US", "Europe Division", "APAC Division", "LATAM Division"]
ACCOUNTS = {
    "Revenue": {
        "4000": "Product Revenue",
        "4100": "Service Revenue",
        "4200": "License Revenue",
        "4300": "Subscription ARR",
        "4400": "Partner Revenue",
    },
    "COGS": {
        "5000": "Direct Labor",
        "5100": "Cloud Infrastructure",
        "5200": "Support Delivery",
        "5300": "Third-Party Licenses",
    },
    "OpEx": {
        "6000": "Sales Salaries & Commission",
        "6100": "Marketing Programs",
        "6200": "R&D Personnel",
        "6300": "G&A - Finance",
        "6400": "G&A - Legal",
        "6500": "G&A - HR",
        "6600": "IT Infrastructure",
        "6700": "Facilities & Rent",
        "6800": "Travel & Entertainment",
    },
}


def _pl_actuals(n: int = 3000) -> pd.DataFrame:
    """Monthly P&L actuals with budget and prior year comparison."""
    months = pd.date_range("2022-01-01", "2024-12-01", freq="MS")
    rows = []
    for month in months:
        for entity in ENTITIES:
            for acc_type, accs in ACCOUNTS.items():
                for acc_id, acc_name in accs.items():
                    base = rng_uniform(50000, 2000000, 1)[0]
                    actual = base * (1 + rng_normal(0, 0.1, 1)[0])
                    budget = base * (1 + rng_normal(0, 0.05, 1)[0])
                    prior_yr = base * rng_uniform(0.85, 1.1, 1)[0]
                    rows.append({
                        "period": month.strftime("%Y-%m"),
                        "fiscal_year": month.year,
                        "fiscal_quarter": f"Q{(month.month - 1) // 3 + 1}",
                        "month_num": month.month,
                        "entity": entity,
                        "account_id": acc_id,
                        "account_name": acc_name,
                        "account_type": acc_type,
                        "cost_center": rng_choice(COST_CENTERS, 1)[0],
                        "actual_usd": round(actual, 2),
                        "budget_usd": round(budget, 2),
                        "prior_year_usd": round(prior_yr, 2),
                        "budget_variance": round(actual - budget, 2),
                        "yoy_change_pct": round((actual - prior_yr) / prior_yr * 100, 2),
                        "budget_achievement_pct": round(actual / budget * 100, 2),
                    })
    df = pd.DataFrame(rows[:n])
    return df


def _cashflow(months: int = 36) -> pd.DataFrame:
    """Monthly cash flow statement."""
    base_operating = 800000
    rows = []
    date = pd.date_range("2022-01-01", periods=months, freq="MS")
    cash_balance = 5000000

    for d in date:
        operating = base_operating * rng_uniform(0.8, 1.3, 1)[0]
        investing = -rng_uniform(50000, 500000, 1)[0]
        financing = rng_choice([-200000, 0, 500000, 1000000], 1,
                                p=[0.3, 0.4, 0.2, 0.1])[0]
        net_change = operating + investing + financing
        cash_balance += net_change
        rows.append({
            "period": d.strftime("%Y-%m"),
            "fiscal_year": d.year,
            "month": d.month,
            "quarter": f"Q{(d.month-1)//3+1}",
            "operating_cashflow": round(operating, 2),
            "investing_cashflow": round(investing, 2),
            "financing_cashflow": round(financing, 2),
            "net_cash_change": round(net_change, 2),
            "closing_cash_balance": round(cash_balance, 2),
            "free_cashflow": round(operating + investing, 2),
            "capex": round(abs(investing), 2),
            "cash_conversion_ratio": round(operating / (base_operating * 1.1), 2),
            "days_cash_on_hand": round(cash_balance / (base_operating / 30), 0),
        })
    return pd.DataFrame(rows)


def _accounts_receivable(n: int = 1500) -> pd.DataFrame:
    invoice_dates = date_range("2023-01-01", "2024-12-31", n)
    due_dates = pd.to_datetime(invoice_dates) + pd.to_timedelta(rng_integers(15, 60, n), unit="D")
    payment_dates = due_dates + pd.to_timedelta(rng_integers(-10, 45, n), unit="D")
    invoice_amount = rng_uniform(1000, 500000, n).round(2)
    paid_amount = invoice_amount * rng_choice([0, 0.5, 1.0], n, p=[0.08, 0.05, 0.87])

    df = pd.DataFrame({
        "invoice_id": make_ids("INV", 10001, n),
        "customer_id": [f"CUST{rng_integers(1000, 3000, 1)[0]}" for _ in range(n)],
        "invoice_date": pd.to_datetime(invoice_dates).date,
        "due_date": due_dates.date,
        "payment_date": np.where(paid_amount > 0, payment_dates.date, None),
        "invoice_amount": invoice_amount,
        "paid_amount": paid_amount.round(2),
        "outstanding_amount": (invoice_amount - paid_amount).round(2),
        "payment_terms": rng_choice(["Net15", "Net30", "Net45", "Net60"], n),
        "days_outstanding": rng_integers(0, 120, n),
        "aging_bucket": rng_choice(["Current", "1-30 Days", "31-60 Days", "61-90 Days", "90+ Days"], n,
                                    p=[0.45, 0.25, 0.15, 0.09, 0.06]),
        "entity": rng_choice(ENTITIES, n),
        "revenue_type": rng_choice(list(ACCOUNTS["Revenue"].values()), n),
        "collection_status": rng_choice(["Paid", "Outstanding", "Overdue", "In Dispute", "Written Off"], n,
                                         p=[0.65, 0.18, 0.10, 0.05, 0.02]),
    })
    return df


class FinanceGenerator(BaseGenerator):
    name = "finance"
    industry = "Corporate Finance"
    description = "P&L actuals vs budget, cash flow, and accounts receivable across 4 entities"

    def generate(self) -> dict[str, pd.DataFrame]:
        print("  Generating P&L Actuals (3,000 rows)...")
        pl = _pl_actuals(3000)
        print("  Generating Cash Flow (36 months)...")
        cf = _cashflow(36)
        print("  Generating Accounts Receivable (1,500 rows)...")
        ar = _accounts_receivable(1500)

        # KPI summary
        kpis = pd.DataFrame({
            "metric": ["Total Revenue 2024", "Total OpEx 2024", "EBITDA Margin", "Free Cash Flow",
                        "AR Outstanding", "Budget Achievement", "YoY Revenue Growth", "Burn Rate"],
            "value": [42500000, 28300000, 22.4, 8700000, 5200000, 96.3, 14.7, 1850000],
            "unit": ["USD", "USD", "%", "USD", "USD", "%", "%", "USD/Month"],
            "vs_target": ["+2.1%", "+1.8%", "+3.2pp", "+12%", "-8%", "-3.7pp", "+4.2pp", "-5%"],
            "status": ["On Track", "Watch", "On Track", "On Track", "On Track", "Watch", "On Track", "On Track"],
        })

        return {
            "PL_Actuals": pl,
            "Cash_Flow": cf,
            "Accounts_Receivable": ar,
            "KPI_Summary": kpis,
        }
