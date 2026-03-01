"""HR & Admin dataset: Employee master, payroll, attendance, recruitment."""
from __future__ import annotations
import numpy as np
import pandas as pd
from .base import BaseGenerator, rng_choice, rng_uniform, rng_integers, rng_normal, date_range, make_ids

DEPARTMENTS = ["Engineering", "Sales", "Marketing", "Finance", "Operations",
               "HR", "Product", "Design", "Legal", "Customer Success", "IT", "Executive"]
JOB_LEVELS = ["L1 - Associate", "L2 - Analyst", "L3 - Senior Analyst", "L4 - Manager",
               "L5 - Senior Manager", "L6 - Director", "L7 - VP", "L8 - C-Suite"]
EMPLOYMENT_TYPES = ["Full-Time", "Part-Time", "Contract", "Intern"]
LOCATIONS = ["New York", "San Francisco", "Austin", "Chicago", "Boston",
             "London", "Berlin", "Singapore", "Sydney", "Toronto", "Remote"]
GENDERS = ["Male", "Female", "Non-Binary"]
ETHNICITIES = ["White", "Asian", "Hispanic", "Black", "Two or More", "Other", "Prefer Not to Say"]
PERFORMANCE_RATINGS = ["Exceptional", "Exceeds Expectations", "Meets Expectations",
                        "Needs Improvement", "Unsatisfactory"]
TERMINATION_REASONS = ["Voluntary - Better Opportunity", "Voluntary - Personal",
                        "Involuntary - Performance", "Involuntary - Restructuring",
                        "Contract End", "Retirement"]
BENEFITS_PLANS = ["Premium Plus", "Premium", "Standard", "Basic", "Waived"]
RECRUITMENT_SOURCES = ["LinkedIn", "Indeed", "Referral", "Agency", "Company Website",
                        "Job Fair", "University", "Glassdoor"]


def _employee_master(n: int = 1500) -> pd.DataFrame:
    depts = rng_choice(DEPARTMENTS, n)
    levels = rng_choice(JOB_LEVELS, n,
                         p=[0.10, 0.18, 0.22, 0.20, 0.14, 0.08, 0.05, 0.03])
    # Salary by level
    level_base = {l: s for l, s in zip(JOB_LEVELS,
        [45000, 62000, 82000, 105000, 135000, 175000, 225000, 380000])}
    base_salaries = np.array([level_base[l] for l in levels])
    salaries = (base_salaries * rng_uniform(0.85, 1.20, n)).round(0)
    bonus_pcts = np.array([{"L1": 0.05, "L2": 0.08, "L3": 0.10, "L4": 0.15, "L5": 0.20,
                             "L6": 0.25, "L7": 0.35, "L8": 0.50}.get(l.split()[0], 0.10)
                           for l in levels])
    bonus = (salaries * bonus_pcts * rng_uniform(0.7, 1.3, n)).round(0)
    hire_dates = date_range("2015-01-01", "2024-06-01", n)
    hire_year = pd.to_datetime(hire_dates).year
    tenure_yrs = (2024 - hire_year + rng_uniform(0, 1, n)).round(1)
    is_active = rng_choice([True, False], n, p=[0.88, 0.12])
    term_dates = np.where(~is_active, date_range("2020-01-01", "2024-12-01", n), None)

    df = pd.DataFrame({
        "employee_id": make_ids("EMP", 1001, n),
        "first_name": [f"First{i}" for i in range(n)],
        "last_name": [f"Last{i}" for i in range(n)],
        "department": depts,
        "sub_department": [f"{d} - {rng_choice(['A', 'B', 'C'], 1)[0]}" for d in depts],
        "job_title": [f"{l.split(' - ')[1]} - {rng_choice(['Specialist', 'Lead', 'Analyst'], 1)[0]}" for l in levels],
        "job_level": levels,
        "manager_id": [f"EMP{rng_integers(1001, 1100, 1)[0]:05d}" for _ in range(n)],
        "location": rng_choice(LOCATIONS, n),
        "employment_type": rng_choice(EMPLOYMENT_TYPES, n, p=[0.80, 0.05, 0.10, 0.05]),
        "hire_date": pd.to_datetime(hire_dates).date,
        "hire_year": hire_year,
        "tenure_years": tenure_yrs,
        "termination_date": term_dates,
        "is_active": is_active,
        "termination_reason": np.where(~is_active,
                                        rng_choice(TERMINATION_REASONS, n), ""),
        "gender": rng_choice(GENDERS, n, p=[0.45, 0.48, 0.07]),
        "ethnicity": rng_choice(ETHNICITIES, n),
        "age_group": rng_choice(["<25", "25-34", "35-44", "45-54", "55+"], n,
                                  p=[0.08, 0.32, 0.30, 0.20, 0.10]),
        "base_salary": salaries,
        "bonus_amount": bonus,
        "total_compensation": (salaries + bonus).round(0),
        "equity_value": (salaries * rng_choice([0, 0.2, 0.4, 0.8, 1.5], n,
                                                p=[0.30, 0.25, 0.20, 0.15, 0.10])).round(0),
        "benefits_plan": rng_choice(BENEFITS_PLANS, n),
        "performance_rating": rng_choice(PERFORMANCE_RATINGS, n,
                                          p=[0.10, 0.25, 0.45, 0.15, 0.05]),
        "last_review_date": pd.to_datetime(date_range("2023-11-01", "2024-04-01", n)).date,
        "promotions_count": rng_integers(0, 5, n),
        "training_hours_ytd": rng_integers(0, 120, n),
        "certifications_count": rng_integers(0, 8, n),
        "remote_work_pct": rng_choice([0, 20, 40, 60, 80, 100], n,
                                       p=[0.15, 0.05, 0.10, 0.20, 0.30, 0.20]),
        "recruitment_source": rng_choice(RECRUITMENT_SOURCES, n),
    })
    return df


def _payroll_monthly(months: int = 24) -> pd.DataFrame:
    """Monthly payroll summary by department."""
    dates = pd.date_range("2023-01-01", periods=months, freq="MS")
    rows = []
    base_headcount = {"Engineering": 280, "Sales": 220, "Marketing": 95, "Finance": 75,
                       "Operations": 110, "HR": 45, "Product": 85, "Design": 40,
                       "Legal": 25, "Customer Success": 100, "IT": 60, "Executive": 12}
    for d in dates:
        for dept, base_hc in base_headcount.items():
            hc_change = rng_integers(-3, 5, 1)[0]
            hc = max(base_hc + hc_change, 5)
            avg_salary = {"Engineering": 145000, "Sales": 118000, "Marketing": 98000,
                          "Finance": 112000, "Operations": 88000, "HR": 85000,
                          "Product": 135000, "Design": 108000, "Legal": 145000,
                          "Customer Success": 92000, "IT": 105000, "Executive": 350000}.get(dept, 100000)
            monthly_sal = (hc * avg_salary / 12 * rng_uniform(0.97, 1.03, 1)[0]).round(2)
            rows.append({
                "period": d.strftime("%Y-%m"),
                "year": d.year,
                "month": d.month,
                "quarter": f"Q{(d.month-1)//3+1}",
                "department": dept,
                "headcount": hc,
                "monthly_salary": monthly_sal,
                "monthly_bonus": round(monthly_sal * rng_uniform(0.02, 0.08, 1)[0], 2),
                "benefits_cost": round(monthly_sal * rng_uniform(0.18, 0.25, 1)[0], 2),
                "overtime_cost": round(monthly_sal * rng_uniform(0, 0.05, 1)[0], 2),
                "total_payroll_cost": None,
                "new_hires": rng_integers(0, 8, 1)[0],
                "terminations": rng_integers(0, 4, 1)[0],
                "avg_salary_per_employee": round(monthly_sal / hc, 2),
            })
    df = pd.DataFrame(rows)
    df["total_payroll_cost"] = (df["monthly_salary"] + df["monthly_bonus"] +
                                 df["benefits_cost"] + df["overtime_cost"]).round(2)
    return df


def _recruitment_pipeline(n: int = 500) -> pd.DataFrame:
    open_dates = date_range("2023-01-01", "2024-12-01", n)
    df = pd.DataFrame({
        "req_id": make_ids("REQ", 1001, n),
        "job_title": [f"{rng_choice(['Senior', 'Lead', '', 'Associate'], 1)[0]} {rng_choice(['Engineer', 'Analyst', 'Manager', 'Designer', 'Specialist'], 1)[0]}".strip() for _ in range(n)],
        "department": rng_choice(DEPARTMENTS, n),
        "level": rng_choice(JOB_LEVELS[:6], n),
        "location": rng_choice(LOCATIONS, n),
        "open_date": pd.to_datetime(open_dates).date,
        "applications": rng_integers(10, 350, n),
        "phone_screens": rng_integers(5, 80, n),
        "technical_interviews": rng_integers(2, 30, n),
        "final_rounds": rng_integers(1, 10, n),
        "offers_extended": rng_integers(0, 5, n),
        "offers_accepted": rng_integers(0, 3, n),
        "time_to_fill_days": rng_integers(15, 120, n),
        "cost_per_hire": rng_uniform(1500, 25000, n).round(2),
        "offer_acceptance_rate": rng_uniform(0.4, 0.95, n).round(2),
        "top_source": rng_choice(RECRUITMENT_SOURCES, n),
        "status": rng_choice(["Filled", "Open", "On Hold", "Cancelled"], n,
                              p=[0.55, 0.30, 0.08, 0.07]),
        "diversity_hire": rng_choice([True, False], n, p=[0.38, 0.62]),
    })
    return df


class HRAdminGenerator(BaseGenerator):
    name = "hr_admin"
    industry = "Human Resources"
    description = "Employee master, monthly payroll, and recruitment pipeline data"

    def generate(self) -> dict[str, pd.DataFrame]:
        print("  Generating Employee Master (1,500 rows)...")
        employees = _employee_master(1500)
        print("  Generating Monthly Payroll (24 months)...")
        payroll = _payroll_monthly(24)
        print("  Generating Recruitment Pipeline (500 rows)...")
        recruitment = _recruitment_pipeline(500)

        # Turnover summary
        active = employees[employees["is_active"]]
        left = employees[~employees["is_active"]]
        turnover_by_dept = employees.groupby("department").agg(
            total=("employee_id", "count"),
            active=("is_active", "sum"),
            departed=("is_active", lambda x: (~x).sum()),
        ).reset_index()
        turnover_by_dept["turnover_rate_pct"] = (
            turnover_by_dept["departed"] / turnover_by_dept["total"] * 100
        ).round(1)
        turnover_by_dept["avg_tenure_years"] = employees.groupby("department")["tenure_years"].mean().round(1).values[:len(turnover_by_dept)]

        return {
            "Employee_Master": employees,
            "Payroll_Monthly": payroll,
            "Recruitment_Pipeline": recruitment,
            "Turnover_Analysis": turnover_by_dept,
        }
