"""Extreme load dataset: 5 datasets, 15+ columns, 10,000+ rows each."""
from __future__ import annotations
import numpy as np
import pandas as pd
from .base import (
    BaseGenerator, RNG, rng_choice, rng_uniform, rng_integers,
    rng_normal, date_range, make_ids,
)

REGIONS = ["North America", "Europe", "Asia Pacific", "Latin America", "Middle East & Africa"]
COUNTRIES = {
    "North America": ["USA", "Canada", "Mexico"],
    "Europe": ["UK", "Germany", "France", "Italy", "Spain"],
    "Asia Pacific": ["China", "Japan", "India", "Australia", "Singapore"],
    "Latin America": ["Brazil", "Argentina", "Colombia"],
    "Middle East & Africa": ["UAE", "Saudi Arabia", "South Africa"],
}
PRODUCT_CATEGORIES = ["Electronics", "Software", "Hardware", "Services", "Consulting", "Support"]
PRODUCT_SUBCATEGORIES = {
    "Electronics": ["Laptops", "Tablets", "Phones", "Accessories", "Networking"],
    "Software": ["ERP", "CRM", "Analytics", "Security", "Cloud"],
    "Hardware": ["Servers", "Storage", "Compute", "Peripherals"],
    "Services": ["Implementation", "Training", "Support", "Managed"],
    "Consulting": ["Strategy", "Operations", "Technology", "Finance"],
    "Support": ["Premium", "Standard", "Basic"],
}
CHANNELS = ["Direct Sales", "Partner", "Online", "Reseller", "Distributor"]
PAYMENT_METHODS = ["Wire Transfer", "Credit Card", "ACH", "Check", "Invoice Net30", "Invoice Net60"]
ORDER_STATUSES = ["Delivered", "Shipped", "Processing", "Cancelled", "Returned", "Pending"]
SEGMENTS = ["Enterprise", "Mid-Market", "SMB", "Government", "Education", "Healthcare"]
DEPARTMENTS = ["Engineering", "Sales", "Finance", "HR", "Operations", "Marketing", "Legal", "IT"]
SHIFT_TYPES = ["Morning", "Afternoon", "Night", "Flexible"]
ATTENDANCE_STATUS = ["Present", "Absent", "Half-Day", "Work From Home", "On Leave"]
LEAVE_TYPES = ["Annual", "Sick", "Personal", "Unpaid", "Maternity/Paternity"]


def _sales_transactions(n: int = 12000) -> pd.DataFrame:
    regions = rng_choice(REGIONS, n)
    categories = rng_choice(PRODUCT_CATEGORIES, n)
    subcats = [rng_choice(PRODUCT_SUBCATEGORIES[cat], 1)[0] for cat in categories]
    qty = rng_integers(1, 50, n)
    unit_price = rng_uniform(50, 25000, n).round(2)
    discount = rng_choice([0.0, 0.05, 0.10, 0.15, 0.20, 0.25], n,
                          p=[0.4, 0.2, 0.15, 0.1, 0.1, 0.05])
    gross = qty * unit_price
    discount_amt = gross * discount
    net_sales = gross - discount_amt
    cost_pct = rng_uniform(0.35, 0.72, n)
    cost = (net_sales * cost_pct).round(2)
    profit = (net_sales - cost).round(2)
    countries = [rng_choice(COUNTRIES[r], 1)[0] for r in regions]

    df = pd.DataFrame({
        "transaction_id": make_ids("TXN", 10001, n),
        "date": date_range("2022-01-01", "2024-12-31", n),
        "customer_id": [f"CUST{rng_integers(1000, 9999, 1)[0]}" for _ in range(n)],
        "customer_name": [f"Customer Corp {rng_integers(100, 999, 1)[0]}" for _ in range(n)],
        "region": regions,
        "country": countries,
        "city": [f"City_{rng_integers(1, 200, 1)[0]}" for _ in range(n)],
        "product_id": [f"PRD{rng_integers(1000, 5000, 1)[0]}" for _ in range(n)],
        "product_name": [f"{s} Pro {rng_integers(100, 999, 1)[0]}" for s in subcats],
        "category": categories,
        "subcategory": subcats,
        "quantity": qty,
        "unit_price": unit_price,
        "discount_pct": discount,
        "gross_amount": gross.round(2),
        "discount_amount": discount_amt.round(2),
        "net_sales": net_sales.round(2),
        "cost": cost,
        "gross_profit": profit,
        "profit_margin_pct": (profit / net_sales * 100).round(2),
        "channel": rng_choice(CHANNELS, n),
        "payment_method": rng_choice(PAYMENT_METHODS, n),
        "rep_id": [f"REP{rng_integers(101, 150, 1)[0]}" for _ in range(n)],
        "order_status": rng_choice(ORDER_STATUSES, n, p=[0.55, 0.15, 0.10, 0.08, 0.07, 0.05]),
        "shipping_days": rng_integers(1, 30, n),
        "customer_segment": rng_choice(SEGMENTS, n),
        "year": pd.to_datetime(date_range("2022-01-01", "2024-12-31", n)).year.values,
        "quarter": pd.to_datetime(date_range("2022-01-01", "2024-12-31", n)).quarter.values,
        "month": pd.to_datetime(date_range("2022-01-01", "2024-12-31", n)).month.values,
    })
    df["date"] = pd.to_datetime(df["date"]).dt.date
    return df.sort_values("date").reset_index(drop=True)


def _product_inventory(n: int = 10000) -> pd.DataFrame:
    categories = rng_choice(PRODUCT_CATEGORIES, n)
    subcats = [rng_choice(PRODUCT_SUBCATEGORIES[cat], 1)[0] for cat in categories]
    qty_on_hand = rng_integers(0, 5000, n)
    qty_reserved = rng_integers(0, 500, n)
    reorder_pt = rng_integers(50, 500, n)
    safety_stk = rng_integers(20, 200, n)
    unit_cost = rng_uniform(10, 8000, n).round(2)
    sell_price = (unit_cost * rng_uniform(1.2, 3.0, n)).round(2)
    abc_class = rng_choice(["A", "B", "C"], n, p=[0.2, 0.3, 0.5])

    df = pd.DataFrame({
        "product_id": make_ids("PRD", 1000, n),
        "sku": [f"SKU-{rng_integers(10000, 99999, 1)[0]}" for _ in range(n)],
        "product_name": [f"{s} {rng_choice(['Pro', 'Lite', 'Enterprise', 'Standard'], 1)[0]} {rng_integers(100, 999, 1)[0]}" for s in subcats],
        "category": categories,
        "subcategory": subcats,
        "brand": rng_choice(["TechCorp", "InnoSys", "DataPro", "CloudBase", "NetEdge", "SecureIT"], n),
        "supplier_id": [f"SUP{rng_integers(100, 200, 1)[0]}" for _ in range(n)],
        "supplier_name": [f"Supplier {rng_integers(100, 200, 1)[0]} Ltd" for _ in range(n)],
        "warehouse": rng_choice(["WH-EAST", "WH-WEST", "WH-CENTRAL", "WH-NORTH", "WH-SOUTH"], n),
        "bin_location": [f"R{rng_integers(1, 50, 1)[0]}-B{rng_integers(1, 30, 1)[0]}" for _ in range(n)],
        "quantity_on_hand": qty_on_hand,
        "quantity_reserved": np.minimum(qty_reserved, qty_on_hand),
        "quantity_available": np.maximum(qty_on_hand - np.minimum(qty_reserved, qty_on_hand), 0),
        "reorder_point": reorder_pt,
        "safety_stock": safety_stk,
        "below_reorder": (qty_on_hand < reorder_pt).astype(int),
        "last_received_date": date_range("2023-01-01", "2024-12-31", n),
        "last_sold_date": date_range("2023-06-01", "2024-12-31", n),
        "unit_cost": unit_cost,
        "selling_price": sell_price,
        "margin_pct": ((sell_price - unit_cost) / sell_price * 100).round(2),
        "inventory_value": (qty_on_hand * unit_cost).round(2),
        "weight_kg": rng_uniform(0.1, 100, n).round(2),
        "lead_time_days": rng_integers(1, 90, n),
        "abc_class": abc_class,
        "is_active": rng_choice([True, False], n, p=[0.85, 0.15]),
        "days_since_last_sale": rng_integers(0, 365, n),
    })
    return df


def _customer_master(n: int = 11000) -> pd.DataFrame:
    regions = rng_choice(REGIONS, n)
    segments = rng_choice(SEGMENTS, n, p=[0.25, 0.30, 0.25, 0.08, 0.07, 0.05])
    acq_dates = date_range("2018-01-01", "2024-06-01", n)
    acq_year = pd.to_datetime(acq_dates).year
    annual_rev = rng_uniform(10000, 5000000, n).round(2)
    ltv = (annual_rev * rng_uniform(1.5, 8.0, n)).round(2)

    df = pd.DataFrame({
        "customer_id": make_ids("CUST", 1000, n),
        "company_name": [f"Corp {rng_integers(1000, 9999, 1)[0]} Inc" for _ in range(n)],
        "industry": rng_choice(["Technology", "Finance", "Healthcare", "Retail", "Manufacturing",
                                 "Education", "Government", "Energy", "Telecom", "Media"], n),
        "segment": segments,
        "region": regions,
        "country": [rng_choice(COUNTRIES[r], 1)[0] for r in regions],
        "acquisition_date": acq_dates,
        "acquisition_year": acq_year,
        "acquisition_channel": rng_choice(["Inbound", "Outbound", "Referral", "Partner", "Event", "Digital"], n),
        "annual_revenue_usd": annual_rev,
        "lifetime_value": ltv,
        "total_orders": rng_integers(1, 150, n),
        "avg_order_value": rng_uniform(500, 50000, n).round(2),
        "last_order_date": date_range("2023-01-01", "2024-12-31", n),
        "days_since_last_order": rng_integers(0, 365, n),
        "account_manager": [f"AM{rng_integers(101, 130, 1)[0]}" for _ in range(n)],
        "nps_score": rng_integers(-100, 100, n),
        "health_score": rng_integers(0, 100, n),
        "churn_risk": rng_choice(["Low", "Medium", "High"], n, p=[0.6, 0.28, 0.12]),
        "contract_type": rng_choice(["Annual", "Multi-Year", "Month-to-Month", "Perpetual"], n),
        "support_tier": rng_choice(["Premium", "Standard", "Basic", "None"], n),
        "is_active": rng_choice([True, False], n, p=[0.82, 0.18]),
    })
    df["acquisition_date"] = pd.to_datetime(df["acquisition_date"]).dt.date
    df["last_order_date"] = pd.to_datetime(df["last_order_date"]).dt.date
    return df


def _financial_transactions(n: int = 15000) -> pd.DataFrame:
    gl_accounts = {
        "Revenue": ["Product Revenue", "Service Revenue", "License Revenue", "Subscription Revenue"],
        "COGS": ["Product Cost", "Service Delivery Cost", "Support Cost"],
        "OpEx": ["Salaries & Benefits", "Marketing Spend", "R&D Investment", "G&A Expenses",
                 "Facilities", "IT Infrastructure", "Travel & Entertainment"],
        "CapEx": ["Equipment Purchase", "Software License", "Leasehold Improvements"],
    }
    categories = []
    accounts = []
    for cat, accts in gl_accounts.items():
        c = [cat] * len(accts)
        categories.extend(c)
        accounts.extend(accts)

    account_idx = rng_integers(0, len(accounts), n)
    amounts = rng_normal(0, 50000, n)
    # Revenue positive, costs negative
    cat_labels = np.array([categories[i] for i in account_idx])
    signs = np.where(cat_labels == "Revenue", 1, -1)
    amounts = np.abs(amounts) * signs

    df = pd.DataFrame({
        "transaction_id": make_ids("FIN", 50001, n),
        "posting_date": date_range("2022-01-01", "2024-12-31", n),
        "gl_account_id": [f"GL{4000 + account_idx[i]}" for i in range(n)],
        "account_name": [accounts[account_idx[i]] for i in range(n)],
        "account_category": [categories[account_idx[i]] for i in range(n)],
        "cost_center": rng_choice(["CC-SALES", "CC-MKT", "CC-ENG", "CC-OPS", "CC-FIN", "CC-HR", "CC-IT"], n),
        "department": rng_choice(DEPARTMENTS, n),
        "entity": rng_choice(["Corp-US", "Corp-EU", "Corp-APAC", "Corp-LATAM"], n),
        "currency": rng_choice(["USD", "EUR", "GBP", "JPY", "AUD"], n, p=[0.55, 0.20, 0.10, 0.10, 0.05]),
        "local_amount": amounts.round(2),
        "usd_amount": (amounts * rng_uniform(0.85, 1.15, n)).round(2),
        "budget_amount": (amounts * rng_uniform(0.8, 1.2, n)).round(2),
        "budget_variance": None,
        "period": [f"P{d.month:02d}" for d in pd.to_datetime(date_range("2022-01-01", "2024-12-31", n))],
        "fiscal_year": pd.to_datetime(date_range("2022-01-01", "2024-12-31", n)).year,
        "fiscal_quarter": pd.to_datetime(date_range("2022-01-01", "2024-12-31", n)).quarter,
        "is_recurring": rng_choice([True, False], n, p=[0.6, 0.4]),
        "journal_entry": [f"JE{rng_integers(10000, 99999, 1)[0]}" for _ in range(n)],
        "posted_by": [f"USER{rng_integers(101, 120, 1)[0]}" for _ in range(n)],
        "approved_by": [f"MGR{rng_integers(201, 210, 1)[0]}" for _ in range(n)],
    })
    df["budget_variance"] = (df["usd_amount"] - df["budget_amount"]).round(2)
    df["posting_date"] = pd.to_datetime(df["posting_date"]).dt.date
    return df.sort_values("posting_date").reset_index(drop=True)


def _employee_attendance(n: int = 10000) -> pd.DataFrame:
    emp_ids = [f"EMP{rng_integers(1001, 1200, 1)[0]}" for _ in range(n)]
    depts = rng_choice(DEPARTMENTS, n)
    dates = date_range("2024-01-01", "2024-12-31", n)
    status = rng_choice(ATTENDANCE_STATUS, n, p=[0.72, 0.04, 0.06, 0.12, 0.06])
    planned_hrs = np.where(rng_choice(["WD", "WE"], n, p=[0.7, 0.3]) == "WD", 8, 0)
    actual_hrs = np.where(status == "Present", rng_normal(8, 0.5, n).clip(6, 12),
                 np.where(status == "Half-Day", rng_normal(4, 0.5, n).clip(2, 6),
                 np.where(status == "Work From Home", rng_normal(7.5, 1, n).clip(4, 10), 0)))

    df = pd.DataFrame({
        "record_id": make_ids("ATT", 1, n),
        "employee_id": emp_ids,
        "department": depts,
        "date": dates,
        "day_of_week": pd.to_datetime(dates).day_name(),
        "week_number": pd.to_datetime(dates).isocalendar().week.values,
        "month": pd.to_datetime(dates).month,
        "shift_type": rng_choice(SHIFT_TYPES, n),
        "status": status,
        "planned_hours": planned_hrs.astype(float).round(1),
        "actual_hours": actual_hrs.round(1),
        "overtime_hours": np.maximum(actual_hrs - 8, 0).round(1),
        "late_arrival_mins": np.where(status == "Present",
                                       np.maximum(rng_normal(0, 15, n), 0).round(0), 0),
        "early_departure_mins": np.where(status == "Present",
                                          np.maximum(rng_normal(0, 10, n), 0).round(0), 0),
        "leave_type": np.where(status == "On Leave", rng_choice(LEAVE_TYPES, n), ""),
        "productivity_score": np.where(np.isin(status, ["Present", "Work From Home"]),
                                        rng_integers(50, 100, n), 0),
        "tasks_completed": np.where(np.isin(status, ["Present", "Work From Home"]),
                                     rng_integers(2, 20, n), 0),
    })
    df["date"] = pd.to_datetime(df["date"]).dt.date
    return df.sort_values(["employee_id", "date"]).reset_index(drop=True)


class ExtremeLoadGenerator(BaseGenerator):
    name = "extreme_load"
    industry = "Multi-Industry"
    description = "5 large datasets with 15+ columns and 10,000+ rows each"

    def generate(self) -> dict[str, pd.DataFrame]:
        print("  Generating Sales Transactions (12,000 rows)...")
        sales = _sales_transactions(12000)
        print("  Generating Product Inventory (10,000 rows)...")
        inventory = _product_inventory(10000)
        print("  Generating Customer Master (11,000 rows)...")
        customers = _customer_master(11000)
        print("  Generating Financial Transactions (15,000 rows)...")
        financials = _financial_transactions(15000)
        print("  Generating Employee Attendance (10,000 rows)...")
        attendance = _employee_attendance(10000)
        return {
            "Sales_Transactions": sales,
            "Product_Inventory": inventory,
            "Customer_Master": customers,
            "Financial_Transactions": financials,
            "Employee_Attendance": attendance,
        }
