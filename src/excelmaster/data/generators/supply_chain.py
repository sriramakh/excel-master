"""Supply chain industry dataset: Shipments, inventory, carrier performance."""
from __future__ import annotations
import numpy as np
import pandas as pd
from .base import BaseGenerator, rng_choice, rng_uniform, rng_integers, rng_normal, date_range, make_ids

CARRIERS = ["FedEx", "UPS", "DHL", "USPS", "DB Schenker", "Kuehne+Nagel", "XPO Logistics",
            "C.H. Robinson", "J.B. Hunt", "Amazon Freight"]
SHIPPING_MODES = ["Air", "Ocean", "Road", "Rail", "Express", "Multimodal"]
ORIGIN_PORTS = ["Shanghai", "Shenzhen", "Los Angeles", "Rotterdam", "Singapore",
                 "New York", "Dubai", "Hamburg", "Hong Kong", "Busan"]
DEST_COUNTRIES = ["USA", "Germany", "UK", "France", "Japan", "Australia",
                   "Canada", "Brazil", "India", "UAE"]
PRODUCT_TYPES = ["Raw Materials", "WIP Components", "Finished Goods", "Spare Parts",
                  "Packaging", "Chemicals", "Electronics", "Textiles"]
INCOTERMS = ["EXW", "FOB", "CIF", "DDP", "DAP", "FCA"]
SHIPMENT_STATUS = ["Delivered", "In Transit", "Cleared Customs", "Pending Pickup",
                    "Delayed", "Exception", "Returned"]
RISK_LEVELS = ["Low", "Medium", "High", "Critical"]
WAREHOUSES = ["WH-LA", "WH-NY", "WH-CHI", "WH-HOU", "WH-SEA", "WH-ATL",
               "WH-LONDON", "WH-FRANKFURT", "WH-SINGAPORE", "WH-SYDNEY"]


def _shipments(n: int = 5000) -> pd.DataFrame:
    ship_dates = date_range("2023-01-01", "2024-12-31", n)
    modes = rng_choice(SHIPPING_MODES, n, p=[0.20, 0.35, 0.30, 0.05, 0.07, 0.03])
    weight = rng_uniform(0.5, 20000, n).round(2)
    volume_cbm = (weight / rng_uniform(200, 500, n)).round(2)
    freight_rate = rng_uniform(0.5, 8.0, n)  # USD per kg
    freight_cost = (weight * freight_rate).round(2)
    insurance = (freight_cost * rng_uniform(0.001, 0.003, n)).round(2)
    customs = (freight_cost * rng_uniform(0.02, 0.08, n)).round(2)
    planned_days = np.where(modes == "Air", rng_integers(2, 7, n),
                   np.where(modes == "Ocean", rng_integers(14, 45, n),
                   np.where(modes == "Express", rng_integers(1, 3, n),
                   rng_integers(3, 14, n))))
    delay_days = np.maximum(rng_normal(0, 3, n), 0).astype(int)
    actual_days = planned_days + delay_days
    on_time = (delay_days == 0).astype(int)

    df = pd.DataFrame({
        "shipment_id": make_ids("SHP", 10001, n),
        "ship_date": pd.to_datetime(ship_dates).date,
        "eta": (pd.to_datetime(ship_dates) + pd.to_timedelta(planned_days, unit="D")).date,
        "actual_arrival": (pd.to_datetime(ship_dates) + pd.to_timedelta(actual_days, unit="D")).date,
        "carrier": rng_choice(CARRIERS, n),
        "shipping_mode": modes,
        "origin_port": rng_choice(ORIGIN_PORTS, n),
        "destination_country": rng_choice(DEST_COUNTRIES, n),
        "destination_warehouse": rng_choice(WAREHOUSES, n),
        "incoterms": rng_choice(INCOTERMS, n),
        "product_type": rng_choice(PRODUCT_TYPES, n),
        "product_id": [f"PRD{rng_integers(1000, 5000, 1)[0]}" for _ in range(n)],
        "weight_kg": weight,
        "volume_cbm": volume_cbm,
        "units": rng_integers(1, 10000, n),
        "freight_cost_usd": freight_cost,
        "insurance_usd": insurance,
        "customs_duties_usd": customs,
        "total_landed_cost": (freight_cost + insurance + customs).round(2),
        "cost_per_kg": freight_rate.round(2),
        "planned_transit_days": planned_days,
        "delay_days": delay_days,
        "actual_transit_days": actual_days,
        "on_time_delivery": on_time,
        "shipment_status": rng_choice(SHIPMENT_STATUS, n,
                                       p=[0.60, 0.15, 0.08, 0.05, 0.06, 0.04, 0.02]),
        "customer_id": [f"CUST{rng_integers(1000, 3000, 1)[0]}" for _ in range(n)],
        "po_number": [f"PO{rng_integers(100000, 999999, 1)[0]}" for _ in range(n)],
        "risk_level": rng_choice(RISK_LEVELS, n, p=[0.50, 0.30, 0.15, 0.05]),
        "requires_temperature_control": rng_choice([True, False], n, p=[0.15, 0.85]),
        "fiscal_quarter": pd.to_datetime(ship_dates).quarter,
        "month": pd.to_datetime(ship_dates).month,
        "year": pd.to_datetime(ship_dates).year,
    })
    return df.sort_values("ship_date").reset_index(drop=True)


def _carrier_scorecard() -> pd.DataFrame:
    rows = []
    for carrier in CARRIERS:
        for mode in ["Air", "Ocean", "Road"]:
            rows.append({
                "carrier": carrier,
                "mode": mode,
                "shipments_ytd": rng_integers(50, 2000, 1)[0],
                "on_time_rate_pct": rng_uniform(72, 99, 1)[0].round(1),
                "avg_transit_days": rng_uniform(1, 35, 1)[0].round(1),
                "avg_cost_per_kg": rng_uniform(0.5, 9.0, 1)[0].round(2),
                "damage_rate_pct": rng_uniform(0.01, 2.5, 1)[0].round(2),
                "claims_count": rng_integers(0, 30, 1)[0],
                "claims_value_usd": rng_uniform(0, 50000, 1)[0].round(2),
                "customer_satisfaction": rng_uniform(3.0, 5.0, 1)[0].round(1),
                "capacity_utilization_pct": rng_uniform(60, 100, 1)[0].round(1),
                "contract_rate": rng_choice(["Fixed", "Spot", "Hybrid"], 1)[0],
                "overall_score": rng_uniform(50, 100, 1)[0].round(1),
                "ranking": rng_integers(1, 10, 1)[0],
            })
    return pd.DataFrame(rows)


def _warehouse_inventory(n: int = 2000) -> pd.DataFrame:
    warehouses = rng_choice(WAREHOUSES, n)
    capacity = np.array([rng_choice([5000, 10000, 15000, 20000], 1)[0] for _ in range(n)])
    current_stock = (capacity * rng_uniform(0.4, 0.95, n)).astype(int)

    df = pd.DataFrame({
        "record_id": make_ids("INV", 1, n),
        "warehouse": warehouses,
        "product_id": [f"PRD{rng_integers(1000, 3000, 1)[0]}" for _ in range(n)],
        "product_type": rng_choice(PRODUCT_TYPES, n),
        "current_stock_units": current_stock,
        "capacity_units": capacity,
        "utilization_pct": (current_stock / capacity * 100).round(1),
        "reorder_point": (capacity * rng_uniform(0.1, 0.25, n)).astype(int),
        "safety_stock": (capacity * rng_uniform(0.05, 0.15, n)).astype(int),
        "avg_daily_demand": rng_integers(10, 500, n),
        "days_of_stock": rng_integers(5, 120, n),
        "stock_value_usd": rng_uniform(5000, 5000000, n).round(2),
        "last_replenishment_date": pd.to_datetime(date_range("2024-01-01", "2024-12-01", n)).date,
        "supplier_lead_time_days": rng_integers(3, 60, n),
        "shrinkage_pct": rng_uniform(0, 3, n).round(2),
        "age_days": rng_integers(1, 365, n),
        "slow_moving": (rng_integers(0, 100, n) < 15).astype(int),
        "obsolete": (rng_integers(0, 100, n) < 5).astype(int),
    })
    return df


class SupplyChainGenerator(BaseGenerator):
    name = "supply_chain"
    industry = "Supply Chain & Logistics"
    description = "Global shipments, carrier performance, and warehouse inventory"

    def generate(self) -> dict[str, pd.DataFrame]:
        print("  Generating Shipments (5,000 rows)...")
        shipments = _shipments(5000)
        print("  Generating Carrier Scorecard...")
        carriers = _carrier_scorecard()
        print("  Generating Warehouse Inventory (2,000 rows)...")
        inventory = _warehouse_inventory(2000)

        # Monthly KPI trend
        shipments_dt = shipments.copy()
        shipments_dt["ship_date"] = pd.to_datetime(shipments_dt["ship_date"])
        monthly_kpi = shipments_dt.groupby(shipments_dt["ship_date"].dt.to_period("M")).agg(
            shipment_count=("shipment_id", "count"),
            total_freight_cost=("freight_cost_usd", "sum"),
            avg_transit_days=("actual_transit_days", "mean"),
            on_time_rate=("on_time_delivery", "mean"),
            total_weight_kg=("weight_kg", "sum"),
        ).round(2).reset_index()
        monthly_kpi.columns = ["period", "shipment_count", "total_freight_cost",
                                 "avg_transit_days", "on_time_rate_pct", "total_weight_kg"]
        monthly_kpi["period"] = monthly_kpi["period"].astype(str)
        monthly_kpi["on_time_rate_pct"] = (monthly_kpi["on_time_rate_pct"] * 100).round(1)

        return {
            "Shipments": shipments,
            "Carrier_Scorecard": carriers,
            "Warehouse_Inventory": inventory,
            "Monthly_KPIs": monthly_kpi,
        }
