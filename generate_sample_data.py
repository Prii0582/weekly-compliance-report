"""Generate two weeks of synthetic operational data for a fictional plant ("Plant A").

All names, areas, suppliers and numbers are invented. A fixed random seed makes
the output identical on every run, so the sample report is reproducible.

Usage:
    python generate_sample_data.py
"""
from __future__ import annotations

from datetime import date, datetime, time, timedelta
from pathlib import Path

import numpy as np
import pandas as pd

SEED = 42
WEEK_ENDING = date(2026, 11, 8)                 # Sunday of the reporting week
START = WEEK_ENDING - timedelta(days=13)        # previous week + reporting week
DAYS = [START + timedelta(days=i) for i in range(14)]
OUT = Path(__file__).resolve().parent / "sample_data"

rng = np.random.default_rng(SEED)


def is_current_week(d: date) -> bool:
    return d > WEEK_ENDING - timedelta(days=7)


# ---------------------------------------------------------------- 1. devices
def devices() -> pd.DataFrame:
    areas = ["Melting", "Pouring Bay", "Moulding", "Fettling", "Lab", "Stores"]
    devs = [(f"DEV-{i:02d}", areas[i % len(areas)]) for i in range(1, 25)]
    chronic = {"DEV-07", "DEV-14"}              # two flaky devices
    rows = []
    for d in DAYS:
        p = 0.07 if is_current_week(d) else 0.04
        for dev, area in devs:
            hrs = 0.0
            if dev in chronic and rng.random() < 0.55:
                hrs = round(rng.uniform(2.5, 9), 1)
            elif rng.random() < p:
                hrs = round(rng.uniform(0.3, 6), 1)
            rows.append({"date": d, "device_id": dev, "area": area, "offline_hours": hrs})
    return pd.DataFrame(rows)


# ------------------------------------------------------------ 2. temperature
def temperature() -> pd.DataFrame:
    eq = {"Furnace F1": 1520, "Furnace F2": 1520, "Furnace F3": 1520, "Holding Furnace H1": 1460}
    rows = []
    for d in DAYS:
        for hour in range(24):
            ts = datetime.combine(d, time(hour))
            for name, mid in eq.items():
                sd = 16
                if name == "Furnace F2" and is_current_week(d):
                    sd = 28                       # F2 drifts this week
                if name == "Holding Furnace H1" and not is_current_week(d):
                    sd = 24
                rows.append({"timestamp": ts, "equipment": name,
                             "temp_c": round(float(rng.normal(mid, sd)), 1)})
    return pd.DataFrame(rows)


# -------------------------------------------------------------------- 3. WIP
def wip() -> pd.DataFrame:
    stages = ["Moulding", "Pouring", "Knockout", "Fettling", "Heat Treatment", "Final Inspection"]
    rows = []
    for i in range(1, 121):
        started = START - timedelta(days=int(rng.integers(0, 10))) + timedelta(days=int(rng.integers(0, 14)))
        duration = int(rng.choice([2, 3, 4, 5, 6, 8, 11, 15], p=[.18, .2, .18, .14, .1, .08, .07, .05]))
        completed = started + timedelta(days=duration)
        rows.append({
            "job_id": f"JOB-{2400 + i}",
            "part_no": f"P-{int(rng.integers(1001, 1060))}",
            "stage": str(rng.choice(stages)),
            "started_on": started,
            "completed_on": completed if completed <= WEEK_ENDING else None,
        })
    return pd.DataFrame(rows)


# ----------------------------------------------------------------- 4. safety
def safety() -> pd.DataFrame:
    areas = ["Pouring Bay", "Melting", "Moulding", "Fettling", "Stores"]
    weights = [0.34, 0.26, 0.18, 0.14, 0.08]
    types = ["No helmet", "No safety shoes", "Restricted zone entry", "No gloves"]
    rows = []
    for d in DAYS:
        n = int(rng.poisson(3.2 if is_current_week(d) else 4.4))
        for _ in range(n):
            area = str(rng.choice(areas, p=weights))
            rows.append({
                "timestamp": datetime.combine(d, time(int(rng.integers(6, 23)), int(rng.integers(0, 60)))),
                "area": area,
                "camera": f"CAM-{areas.index(area) + 1}{int(rng.integers(1, 4))}",
                "event_type": str(rng.choice(types, p=[.45, .2, .15, .2])),
            })
    return pd.DataFrame(rows).sort_values("timestamp")


# ----------------------------------------------------------------- 5. inward
def inward() -> pd.DataFrame:
    suppliers = {"Supplier A": 0.010, "Supplier B": 0.015, "Supplier C": 0.035,
                 "Supplier D": 0.012, "Supplier E": 0.020}
    materials = ["Pig Iron", "Steel Scrap", "Ferro Silicon", "Resin Sand", "Ferro Manganese"]
    reasons = ["Size out of spec", "Moisture", "Contamination", "Chemistry out of spec", "Short supply"]
    rows, lot = [], 5000
    for d in DAYS:
        for _ in range(int(rng.integers(3, 7))):
            lot += 1
            sup = str(rng.choice(list(suppliers)))
            qty = int(rng.integers(5, 40)) * 500
            mean = suppliers[sup] * (1.3 if is_current_week(d) and sup == "Supplier C" else 1)
            rej_pct = max(0.0, float(rng.normal(mean, 0.012)))
            rej = int(round(qty * rej_pct / 10) * 10)
            rows.append({"date": d, "lot_no": f"LOT-{lot}", "supplier": sup,
                         "material": str(rng.choice(materials)), "qty_received_kg": qty,
                         "qty_rejected_kg": rej, "reason": str(rng.choice(reasons)) if rej else ""})
    return pd.DataFrame(rows)


# ------------------------------------------------------------------ 6. yield
def yields() -> pd.DataFrame:
    base = {"Furnace F1": 64.5, "Furnace F2": 63.5, "Furnace F3": 61.0}
    rows, heat = [], 8000
    for d in DAYS:
        for furnace, mu in base.items():
            for _ in range(int(rng.integers(6, 10))):
                heat += 1
                charge = int(rng.integers(18, 26)) * 100
                y = float(rng.normal(mu - (1.2 if furnace == "Furnace F3" and is_current_week(d) else 0), 2.2))
                rows.append({"date": d, "heat_no": f"H-{heat}", "furnace": furnace,
                             "charge_kg": charge, "good_castings_kg": round(charge * y / 100)})
    return pd.DataFrame(rows)


# ---------------------------------------------------------------- 7. reports
def reports() -> pd.DataFrame:
    catalogue = [
        ("Daily Production Report", "Production Planning", time(10, 0), 0.10),
        ("Melting Log", "Melt Shop Supervisor", time(9, 0), 0.22),
        ("Shift Handover Note", "Shift In-charge", time(8, 30), 0.12),
        ("Quality Inspection Summary", "Stores & Quality Control", time(11, 0), 0.08),
        ("Energy Consumption Log", "Maintenance - Electrical", time(12, 0), 0.18),
        ("Dispatch Summary", "Dispatch Desk", time(18, 0), 0.05),
    ]
    rows = []
    for d in DAYS:
        for name, owner, due, p_late in catalogue:
            due_dt = datetime.combine(d, due)
            r = rng.random()
            if r < p_late * 0.2:
                submitted = None                                   # missed entirely
            elif r < p_late:
                submitted = due_dt + timedelta(minutes=int(rng.integers(35, 300)))
            else:
                submitted = due_dt + timedelta(minutes=int(rng.integers(-90, 25)))
            rows.append({"date": d, "report_name": name, "owner": owner,
                         "due_at": due_dt, "submitted_at": submitted})
    return pd.DataFrame(rows)


def main() -> None:
    OUT.mkdir(exist_ok=True)
    files = {
        "device_status.csv": devices(),
        "temperature_readings.csv": temperature(),
        "wip_jobs.csv": wip(),
        "safety_events.csv": safety(),
        "inward_inspection.csv": inward(),
        "heat_yield.csv": yields(),
        "report_submissions.csv": reports(),
    }
    for name, df in files.items():
        df.to_csv(OUT / name, index=False)
        print(f"  wrote sample_data/{name:<26} {len(df):>5} rows")


if __name__ == "__main__":
    main()
