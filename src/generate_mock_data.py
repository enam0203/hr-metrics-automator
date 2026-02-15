from __future__ import annotations

from pathlib import Path
import numpy as np
import pandas as pd


def bounded(value: float, low: float, high: float) -> float:
    return max(low, min(high, value))


def main() -> None:
    np.random.seed(42)

    months = pd.date_range("2025-02-01", "2026-01-01", freq="MS")
    departments = ["Engineering", "Sales", "Operations", "HR", "Finance", "Customer Success"]

    baseline = {
        "Engineering": 95,
        "Sales": 62,
        "Operations": 54,
        "HR": 18,
        "Finance": 21,
        "Customer Success": 37,
    }

    trends = {
        "Engineering": 1.6,
        "Sales": 1.0,
        "Operations": 0.7,
        "HR": 0.2,
        "Finance": 0.15,
        "Customer Success": 0.8,
    }

    rows = []
    for month_idx, month in enumerate(months):
        seasonal_hiring_boost = 1.2 if month.month in [3, 4, 9, 10] else 0.8

        for dept in departments:
            base_hc = baseline[dept] + trends[dept] * month_idx + np.random.normal(0, 1.2)
            headcount = int(round(max(8, base_hc)))

            hire_rate = 0.035 * seasonal_hiring_boost
            new_hires = int(round(max(0, headcount * hire_rate + np.random.normal(0, 1.0))))

            turnover_base = {
                "Engineering": 0.011,
                "Sales": 0.019,
                "Operations": 0.015,
                "HR": 0.010,
                "Finance": 0.009,
                "Customer Success": 0.017,
            }[dept]

            turnover_rate = bounded(
                turnover_base + np.random.normal(0, 0.0035) + (0.002 if month.month in [7, 8] else 0),
                0.004,
                0.05,
            )

            terminations = int(round(max(0, headcount * turnover_rate + np.random.normal(0, 0.6))))
            open_positions = int(round(max(0, new_hires * np.random.uniform(1.1, 2.2) + np.random.normal(0, 1.2))))

            ttf_base = {
                "Engineering": 47,
                "Sales": 34,
                "Operations": 31,
                "HR": 29,
                "Finance": 36,
                "Customer Success": 26,
            }[dept]
            time_to_fill = int(round(max(15, ttf_base + np.random.normal(0, 4) - (1 if month.month in [10, 11] else 0))))

            offer_acceptance_rate = bounded(
                {
                    "Engineering": 0.82,
                    "Sales": 0.79,
                    "Operations": 0.84,
                    "HR": 0.86,
                    "Finance": 0.81,
                    "Customer Success": 0.83,
                }[dept]
                + np.random.normal(0, 0.03)
                - (0.015 if month.month in [6, 7] else 0),
                0.62,
                0.95,
            )

            rows.append(
                {
                    "month": month.strftime("%Y-%m"),
                    "department": dept,
                    "headcount": headcount,
                    "new_hires": new_hires,
                    "terminations": terminations,
                    "open_positions": open_positions,
                    "time_to_fill_days": time_to_fill,
                    "offer_acceptance_rate": round(offer_acceptance_rate * 100, 1),
                    "turnover_rate": round(turnover_rate * 100, 2),
                }
            )

    df = pd.DataFrame(rows)
    out = Path(__file__).resolve().parents[1] / "data" / "hr_metrics_monthly.csv"
    out.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(out, index=False)

    print(f"Generated mock dataset: {out}")
    print(f"Rows: {len(df)} | Months: {df['month'].nunique()} | Departments: {df['department'].nunique()}")


if __name__ == "__main__":
    main()
