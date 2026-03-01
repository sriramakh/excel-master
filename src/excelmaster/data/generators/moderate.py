"""Moderate dataset: E-commerce orders, 2,500 rows, 14 columns."""
from __future__ import annotations
import numpy as np
import pandas as pd
from .base import (
    BaseGenerator, rng_choice, rng_uniform, rng_integers, date_range, make_ids,
)

CATEGORIES = ["Clothing", "Electronics", "Home & Garden", "Sports", "Beauty", "Books",
              "Automotive", "Toys", "Food & Grocery", "Health"]
STATUSES = ["Delivered", "Shipped", "Processing", "Cancelled", "Returned"]
PAYMENT = ["Credit Card", "Debit Card", "PayPal", "Apple Pay", "Buy Now Pay Later", "Bank Transfer"]
SOURCES = ["Organic Search", "Paid Search", "Social Media", "Email Campaign",
           "Direct", "Referral", "App"]
CITIES = ["New York", "Los Angeles", "Chicago", "Houston", "Phoenix", "Philadelphia",
          "San Antonio", "San Diego", "Dallas", "Seattle", "Boston", "Miami",
          "Denver", "Atlanta", "Portland", "Las Vegas", "Minneapolis", "Detroit"]
DEVICE = ["Mobile", "Desktop", "Tablet", "App"]


class ModerateGenerator(BaseGenerator):
    name = "moderate"
    industry = "E-Commerce / Retail"
    description = "E-commerce orders dataset: 2,500 rows, 14 columns"

    def generate(self) -> dict[str, pd.DataFrame]:
        n = 2500
        categories = rng_choice(CATEGORIES, n)
        qty = rng_integers(1, 10, n)
        unit_price = rng_uniform(9.99, 499.99, n).round(2)
        discount = rng_choice([0.0, 0.05, 0.10, 0.15, 0.20], n, p=[0.50, 0.20, 0.15, 0.10, 0.05])
        subtotal = (qty * unit_price).round(2)
        discount_amt = (subtotal * discount).round(2)
        shipping = rng_choice([0, 4.99, 7.99, 12.99, 19.99], n, p=[0.3, 0.2, 0.25, 0.15, 0.10])
        total = (subtotal - discount_amt + shipping).round(2)
        dates = date_range("2023-01-01", "2024-12-31", n)
        rating = rng_choice([1, 2, 3, 4, 5], n, p=[0.03, 0.07, 0.15, 0.35, 0.40])

        df = pd.DataFrame({
            "order_id": make_ids("ORD", 100001, n),
            "order_date": pd.to_datetime(dates).date,
            "customer_id": [f"USR{rng_integers(10000, 50000, 1)[0]}" for _ in range(n)],
            "city": rng_choice(CITIES, n),
            "category": categories,
            "product_name": [f"{cat} Item #{rng_integers(100, 999, 1)[0]}" for cat in categories],
            "quantity": qty,
            "unit_price": unit_price,
            "discount_pct": discount,
            "subtotal": subtotal,
            "shipping_cost": shipping,
            "total_amount": total,
            "payment_method": rng_choice(PAYMENT, n),
            "order_status": rng_choice(STATUSES, n, p=[0.60, 0.12, 0.10, 0.10, 0.08]),
            "acquisition_source": rng_choice(SOURCES, n),
            "device_type": rng_choice(DEVICE, n),
            "is_repeat_customer": rng_choice([True, False], n, p=[0.45, 0.55]),
            "customer_rating": rating,
            "delivery_days": rng_integers(1, 15, n),
            "month": pd.to_datetime(dates).month.values,
            "quarter": pd.to_datetime(dates).quarter.values,
        })

        # Monthly summary sheet
        df["order_date_dt"] = pd.to_datetime(df["order_date"])
        monthly = (
            df.groupby(df["order_date_dt"].dt.to_period("M"))
            .agg(
                orders=("order_id", "count"),
                revenue=("total_amount", "sum"),
                avg_order_value=("total_amount", "mean"),
                units_sold=("quantity", "sum"),
                avg_rating=("customer_rating", "mean"),
            )
            .reset_index()
        )
        monthly.columns = ["month", "orders", "revenue", "avg_order_value", "units_sold", "avg_rating"]
        monthly["month"] = monthly["month"].astype(str)
        monthly["revenue"] = monthly["revenue"].round(2)
        monthly["avg_order_value"] = monthly["avg_order_value"].round(2)
        monthly["avg_rating"] = monthly["avg_rating"].round(2)

        # Category summary
        cat_summary = (
            df.groupby("category")
            .agg(
                orders=("order_id", "count"),
                revenue=("total_amount", "sum"),
                units=("quantity", "sum"),
                avg_price=("unit_price", "mean"),
                cancellation_rate=("order_status", lambda x: (x == "Cancelled").mean() * 100),
                avg_rating=("customer_rating", "mean"),
            )
            .round(2)
            .reset_index()
        )

        df = df.drop(columns=["order_date_dt"])
        return {
            "Orders": df,
            "Monthly_Summary": monthly,
            "Category_Summary": cat_summary,
        }
