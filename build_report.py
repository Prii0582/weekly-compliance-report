"""Weekly Compliance & Violation Report.

Reads one week (plus the previous week, for comparison) of operational data
from sample_data/, checks it against the thresholds in config/thresholds.yaml
and produces:

  * output/weekly_compliance_report.html   - the rendered report
  * output/sample_weekly_compliance_report.pdf  - management-ready PDF
  * output/violation_log.xlsx              - discrepancy / violation log with closure tracking
  * docs/report_preview.png                - picture of page 1 (for the README)

Usage:
    python build_report.py                  # uses config/thresholds.yaml
    python build_report.py --no-pdf         # skip the Chrome step (HTML + Excel only)
"""
from __future__ import annotations

import argparse
import hashlib
import random
import shutil
import subprocess
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
import yaml
from jinja2 import Environment, FileSystemLoader, select_autoescape
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "sample_data"
OUT = ROOT / "output"
DOCS = ROOT / "docs"
CHROME_CANDIDATES = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    "google-chrome", "chromium", "chromium-browser",
]
CLOSE_WITHIN_DAYS = 3          # target days to close a logged violation


# --------------------------------------------------------------------------- helpers
@dataclass
class Week:
    start: date
    end: date

    def days(self) -> list[date]:
        return [self.start + timedelta(days=i) for i in range(7)]

    def mask(self, s: pd.Series) -> pd.Series:
        d = pd.to_datetime(s).dt.date
        return (d >= self.start) & (d <= self.end)


@dataclass
class Section:
    key: str
    number: int
    title: str
    unit: str                     # what one "count" means, shown under the number
    rule: str                     # plain-English rule, built from the config
    rule_short: str = ""          # compact version for the summary table
    count: int = 0
    prev_count: int = 0
    daily: list[int] = field(default_factory=list)
    columns: list[str] = field(default_factory=list)
    rows: list[list] = field(default_factory=list)
    top_item: str = ""
    status: str = "green"
    action: str = ""
    chart: str = ""               # inline SVG of daily counts

    @property
    def change(self) -> int:
        return self.count - self.prev_count

    @property
    def change_pct(self) -> float | None:
        return None if self.prev_count == 0 else 100 * self.change / self.prev_count

    @property
    def arrow(self) -> str:
        return "up" if self.change > 0 else "down" if self.change < 0 else "flat"


def status_for(count: int, cfg: dict) -> str:
    if count <= cfg["green_max"]:
        return "green"
    if count <= cfg["amber_max"]:
        return "amber"
    return "red"


def read(name: str, parse: list[str]) -> pd.DataFrame:
    return pd.read_csv(DATA / name, parse_dates=parse)


# --------------------------------------------------------------------------- checks
# Each check returns a DataFrame of violations with at least: date, item, detail.
def chk_devices(df, cfg, wk: Week):
    v = df[wk.mask(df["date"]) & (df["offline_hours"] > cfg["max_offline_hours_per_day"])].copy()
    v["item"] = v["device_id"] + " (" + v["area"] + ")"
    v["detail"] = v.apply(lambda r: f"{r.device_id} in {r.area} offline {r.offline_hours:.1f} h "
                                    f"(limit {cfg['max_offline_hours_per_day']} h)", axis=1)
    return v


def top_devices(v, n):
    g = (v.groupby(["device_id", "area"])
           .agg(days=("date", "count"), hours=("offline_hours", "sum"))
           .reset_index().sort_values(["days", "hours"], ascending=False).head(n))
    rows = [[r.device_id, r.area, r.days, f"{r.hours:.1f}"] for r in g.itertuples()]
    return ["Device", "Area", "Days offline", "Hours offline"], rows, (g.iloc[0].device_id if len(g) else "")


def chk_temperature(df, cfg, wk: Week):
    v = df[wk.mask(df["timestamp"])].copy()
    rng = cfg["ranges_c"]
    v["min_c"] = v["equipment"].map(lambda e: rng.get(e, {}).get("min"))
    v["max_c"] = v["equipment"].map(lambda e: rng.get(e, {}).get("max"))
    v = v[v["min_c"].notna() & ((v["temp_c"] < v["min_c"]) | (v["temp_c"] > v["max_c"]))].copy()
    v["date"] = v["timestamp"].dt.date
    v["item"] = v["equipment"]
    v["detail"] = v.apply(lambda r: f"{r.equipment} read {r.temp_c:.0f} °C at {r.timestamp:%H:%M} "
                                    f"(allowed {r.min_c:.0f}-{r.max_c:.0f} °C)", axis=1)
    return v


def top_temperature(v, n):
    g = (v.groupby("equipment")
           .agg(count=("temp_c", "size"), lo=("temp_c", "min"), hi=("temp_c", "max"),
                mn=("min_c", "first"), mx=("max_c", "first"))
           .reset_index().sort_values("count", ascending=False).head(n))
    rows = [[r.equipment, r.count, f"{r.mn:.0f}-{r.mx:.0f}", f"{r.lo:.0f} / {r.hi:.0f}"] for r in g.itertuples()]
    return ["Equipment", "Excursions", "Allowed (°C)", "Lowest / highest"], rows, (g.iloc[0].equipment if len(g) else "")


def wip_open_on(df, day: date, max_age: int) -> pd.DataFrame:
    started = df["started_on"].dt.date
    done = df["completed_on"].dt.date
    open_ = (started <= day) & (df["completed_on"].isna() | (done > day))
    j = df[open_].copy()
    j["age_days"] = j["started_on"].map(lambda s: (day - s.date()).days)
    return j[j["age_days"] > max_age]


def chk_wip(df, cfg, wk: Week):
    v = wip_open_on(df, wk.end, cfg["max_age_days"])
    v["date"] = v["started_on"].dt.date + timedelta(days=cfg["max_age_days"] + 1)   # day it became overdue
    v["date"] = v["date"].map(lambda d: max(d, wk.start))
    v["item"] = v["job_id"]
    v["detail"] = v.apply(lambda r: f"{r.job_id} ({r.part_no}) at {r.stage} for {r.age_days} days "
                                    f"(limit {cfg['max_age_days']})", axis=1)
    return v


def daily_wip(df, cfg, wk: Week):
    return [len(wip_open_on(df, d, cfg["max_age_days"])) for d in wk.days()]


def top_wip(v, n):
    g = v.sort_values("age_days", ascending=False).head(n)
    rows = [[r.job_id, r.part_no, r.stage, r.age_days] for r in g.itertuples()]
    return ["Job", "Part", "Stuck at stage", "Age (days)"], rows, (g.iloc[0].job_id if len(g) else "")


def chk_safety(df, cfg, wk: Week):
    v = df[wk.mask(df["timestamp"]) & df["event_type"].isin(cfg["event_types_counted"])].copy()
    v["date"] = v["timestamp"].dt.date
    v["item"] = v["area"]
    v["detail"] = v.apply(lambda r: f"{r.event_type} - {r.area}, {r.camera} at {r.timestamp:%H:%M}", axis=1)
    return v


def top_safety(v, n):
    piv = v.pivot_table(index="area", columns="event_type", values="camera", aggfunc="count", fill_value=0)
    piv["total"] = piv.sum(axis=1)
    piv = piv.sort_values("total", ascending=False).head(n)
    rows = []
    for area, r in piv.iterrows():
        parts = r.drop("total")
        rows.append([area, int(r.get("No helmet", 0)), int(r["total"]), parts.idxmax()])
    return ["Area", "Missing helmet", "All events", "Most common"], rows, (piv.index[0] if len(piv) else "")


def chk_inward(df, cfg, wk: Week):
    v = df[wk.mask(df["date"])].copy()
    v["rej_pct"] = 100 * v["qty_rejected_kg"] / v["qty_received_kg"]
    v = v[v["rej_pct"] > cfg["max_rejection_pct"]].copy()
    v["date"] = v["date"].dt.date
    v["item"] = v["supplier"]
    v["detail"] = v.apply(lambda r: f"{r.lot_no} {r.material} from {r.supplier}: {r.rej_pct:.1f}% rejected "
                                    f"({r.reason})", axis=1)
    return v


def top_inward(v, n):
    g = (v.groupby("supplier")
           .agg(lots=("lot_no", "count"), rec=("qty_received_kg", "sum"), rej=("qty_rejected_kg", "sum"),
                reason=("reason", lambda s: s.mode().iat[0]))
           .reset_index().sort_values(["lots", "rej"], ascending=False).head(n))
    rows = [[r.supplier, r.lots, f"{100 * r.rej / r.rec:.1f}%", r.reason] for r in g.itertuples()]
    return ["Supplier", "Lots over limit", "Rejection %", "Main reason"], rows, (g.iloc[0].supplier if len(g) else "")


def chk_yield(df, cfg, wk: Week):
    v = df[wk.mask(df["date"])].copy()
    v["yield_pct"] = 100 * v["good_castings_kg"] / v["charge_kg"]
    v["target"] = v["furnace"].map(cfg["target_yield_pct"])
    v = v[v["target"].notna() & (v["yield_pct"] < v["target"])].copy()
    v["date"] = v["date"].dt.date
    v["item"] = v["furnace"]
    v["detail"] = v.apply(lambda r: f"Heat {r.heat_no} on {r.furnace}: yield {r.yield_pct:.1f}% "
                                    f"(target {r.target:.0f}%)", axis=1)
    return v


def top_yield(v, n):
    g = (v.groupby("furnace")
           .agg(heats=("heat_no", "count"), avg=("yield_pct", "mean"), low=("yield_pct", "min"),
                target=("target", "first"))
           .reset_index().sort_values("heats", ascending=False).head(n))
    rows = [[r.furnace, r.heats, f"{r.target:.0f}%", f"{r.avg:.1f}% / {r.low:.1f}%"] for r in g.itertuples()]
    return ["Furnace", "Heats below", "Target", "Avg / lowest yield"], rows, (g.iloc[0].furnace if len(g) else "")


def chk_reports(df, cfg, wk: Week):
    v = df[wk.mask(df["date"])].copy()
    limit = v["due_at"] + pd.Timedelta(minutes=cfg["grace_minutes"])
    v = v[v["submitted_at"].isna() | (v["submitted_at"] > limit)].copy()
    v["delay_min"] = ((v["submitted_at"] - v["due_at"]).dt.total_seconds() / 60)
    v["date"] = v["date"].dt.date
    v["item"] = v["report_name"]

    def det(r):
        if pd.isna(r.submitted_at):
            return f"{r.report_name} for {r.date:%d %b} not submitted (due {r.due_at:%H:%M})"
        return f"{r.report_name} for {r.date:%d %b} submitted {r.delay_min / 60:.1f} h late"
    v["detail"] = v.apply(det, axis=1)
    return v


def top_reports(v, n):
    g = (v.groupby(["report_name", "owner"])
           .agg(late=("date", "count"), missed=("submitted_at", lambda s: int(s.isna().sum())),
                delay=("delay_min", "mean"))
           .reset_index().sort_values(["late", "missed"], ascending=False).head(n))
    rows = [[r.report_name, r.owner, f"{r.late} ({r.missed} missed)",
             "-" if pd.isna(r.delay) else f"{r.delay / 60:.1f} h"] for r in g.itertuples()]
    return ["Report", "Owner", "Late", "Avg delay"], rows, (g.iloc[0].report_name if len(g) else "")


# --------------------------------------------------------------------------- build
def build_sections(cfg: dict, this: Week, prev: Week):
    data = {
        "devices": read("device_status.csv", ["date"]),
        "temperature": read("temperature_readings.csv", ["timestamp"]),
        "wip": read("wip_jobs.csv", ["started_on", "completed_on"]),
        "safety": read("safety_events.csv", ["timestamp"]),
        "inward": read("inward_inspection.csv", ["date"]),
        "yield": read("heat_yield.csv", ["date"]),
        "reports": read("report_submissions.csv", ["date", "due_at", "submitted_at"]),
    }
    c = cfg
    spec = [
        ("devices_offline", "devices", chk_devices, top_devices, "device-days offline",
         f"A device offline for more than {c['devices_offline']['max_offline_hours_per_day']} hours in a day"),
        ("temperature", "temperature", chk_temperature, top_temperature, "readings out of range",
         "Hourly reading outside the allowed band for that equipment"),
        ("wip_ageing", "wip", chk_wip, top_wip, "jobs open too long",
         f"Job still in progress more than {c['wip_ageing']['max_age_days']} days after start"),
        ("safety", "safety", chk_safety, top_safety, "safety events",
         "Camera events: " + ", ".join(c["safety"]["event_types_counted"]).lower()),
        ("inward_rejection", "inward", chk_inward, top_inward, "lots over rejection limit",
         f"Inward lot with more than {c['inward_rejection']['max_rejection_pct']:g}% quantity rejected"),
        ("yield", "yield", chk_yield, top_yield, "heats below target",
         "Heat whose good-casting yield is below the furnace target"),
        ("late_reports", "reports", chk_reports, top_reports, "late / missed reports",
         f"Report submitted more than {c['late_reports']['grace_minutes']} min after due time, or not at all"),
    ]
    n = cfg["report"]["top_rows_per_table"]
    short = [
        f"Offline > {c['devices_offline']['max_offline_hours_per_day']} h in a day",
        "Outside allowed °C band",
        f"Open > {c['wip_ageing']['max_age_days']} days",
        f"{len(c['safety']['event_types_counted'])} camera event types",
        f"Lot rejection > {c['inward_rejection']['max_rejection_pct']:g}%",
        "Yield < furnace target",
        f"> {c['late_reports']['grace_minutes']} min late or missing",
    ]
    sections, log_rows = [], []
    for i, (key, dkey, chk, top, unit, rule) in enumerate(spec, start=1):
        scfg, df = cfg[key], data[dkey]
        cur, old = chk(df, scfg, this), chk(df, scfg, prev)
        s = Section(key=key, number=i, title=scfg["title"], unit=unit, rule=rule,
                    count=len(cur), prev_count=len(old), rule_short=short[i - 1])
        if key == "wip_ageing":
            s.daily = daily_wip(df, scfg, this)
        else:
            per_day = pd.Series(pd.to_datetime(cur["date"]).dt.date).value_counts() if len(cur) else pd.Series(dtype=int)
            s.daily = [int(per_day.get(d, 0)) for d in this.days()]
        s.columns, s.rows, s.top_item = top(cur, n) if len(cur) else (["Item"], [], "")
        s.status = status_for(s.count, scfg)
        s.action = scfg["action"].format(item=s.top_item or "-")
        sections.append(s)
        for frame in (old, cur):
            for r in frame.itertuples():
                log_rows.append({"date": pd.Timestamp(r.date).date(),
                                 "category": scfg["title"], "detail": r.detail,
                                 "owner": getattr(r, "owner", None) if key == "late_reports" else scfg["owner"]})
    return sections, log_rows


def top_actions(sections: list[Section], k: int = 3) -> list[Section]:
    rank = {"red": 2, "amber": 1, "green": 0}
    ordered = sorted(sections, key=lambda s: (rank[s.status], s.change, s.count), reverse=True)
    return [s for s in ordered if s.status != "green"][:k] or ordered[:k]


# --------------------------------------------------------------------------- violation log
def build_log(log_rows: list[dict], run_date: date) -> pd.DataFrame:
    log = pd.DataFrame(log_rows).sort_values(["date", "category"]).reset_index(drop=True)
    ids, status, closed, target, check, days_open = [], [], [], [], [], []
    for i, r in log.iterrows():
        vid = f"VL-{r.date:%y%m}-{i + 1:04d}"
        # Deterministic simulated resolution: older items are more likely to be closed.
        rnd = random.Random(int(hashlib.md5(vid.encode()).hexdigest(), 16))
        age = (run_date - r.date).days
        p_closed = 0.92 if age > 7 else 0.55 if age > 2 else 0.2
        tgt = r.date + timedelta(days=CLOSE_WITHIN_DAYS)
        c_on = None
        if rnd.random() < p_closed:
            c_on = r.date + timedelta(days=rnd.choice([0, 1, 1, 2, 2, 3, 3, 4, 5, 6]))
            if c_on >= run_date:
                c_on = run_date - timedelta(days=1)
        ids.append(vid)
        status.append("Closed" if c_on else "Open")
        closed.append(c_on)
        target.append(tgt)
        if c_on:
            check.append("Closed on time" if c_on <= tgt else "Closed late")
            days_open.append((c_on - r.date).days)
        else:
            check.append("Open - overdue" if run_date > tgt else "Open - within target")
            days_open.append(age)
    log.insert(0, "ID", ids)
    log["status"] = status
    log["closed_on"] = closed
    log["target_close_by"] = target
    log["days_open"] = days_open
    log["closure_check"] = check
    return log


def write_log_xlsx(log: pd.DataFrame, path: Path, plant: str, run_date: date, order: list[str]) -> None:
    navy, teal = "1B2A4A", "0F8B8D"
    thin = Side(style="thin", color="D5DBE3")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    head_font = Font(bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", fgColor=navy)

    wb = Workbook()
    ws = wb.active
    ws.title = "Violation Log"
    headers = ["ID", "Date", "Category", "Detail", "Owner", "Status", "Closed On",
               "Target Close By", "Days Open", "Closure Check"]
    ws.append(headers)
    for r in log.itertuples(index=False):
        ws.append([r.ID, r.date, r.category, r.detail, r.owner, r.status, r.closed_on,
                   r.target_close_by, r.days_open, r.closure_check])
    widths = [14, 12, 34, 70, 26, 10, 12, 15, 10, 20]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for cell in ws[1]:
        cell.font, cell.fill, cell.border = head_font, head_fill, border
        cell.alignment = Alignment(horizontal="center", vertical="center")
    last = ws.max_row
    for row in ws.iter_rows(min_row=2, max_row=last):
        for cell in row:
            cell.border = border
        for idx in (1, 6, 7):
            row[idx].number_format = "DD-MMM-YYYY"
            row[idx].alignment = Alignment(horizontal="center")
        row[5].alignment = Alignment(horizontal="center")
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:J{last}"
    rng = f"A2:J{last}"
    ws.conditional_formatting.add(rng, FormulaRule(formula=['$J2="Open - overdue"'],
                                  fill=PatternFill("solid", fgColor="FBE3E1")))
    ws.conditional_formatting.add(f"F2:F{last}", FormulaRule(formula=['$F2="Closed"'],
                                  fill=PatternFill("solid", fgColor="DFF3EC"), font=Font(color="11734B", bold=True)))
    ws.conditional_formatting.add(f"F2:F{last}", FormulaRule(formula=['$F2="Open"'],
                                  fill=PatternFill("solid", fgColor="FFF1D6"), font=Font(color="9A5B00", bold=True)))
    dv = DataValidation(type="list", formula1='"Open,Closed"', allow_blank=False)
    dv.add(f"F2:F{last + 500}")
    ws.add_data_validation(dv)

    # Summary sheet - live COUNTIFS formulas so it updates when someone closes an item.
    sm = wb.create_sheet("Closure Summary", 0)
    sm["A1"] = f"Violation Log - Closure Summary ({plant})"
    sm["A1"].font = Font(bold=True, size=14, color=navy)
    sm["A2"] = (f"Status as of {run_date:%d %b %Y}. Target: close within {CLOSE_WITHIN_DAYS} days. "
                "Change Status / Closed On in the 'Violation Log' sheet and this summary updates.")
    sm["A2"].font = Font(italic=True, color="5B6577")
    cols = ["Category", "Logged", "Closed", "Open", "Open - overdue", "Closed on time", "On-time closure %"]
    sm.append([])
    sm.append(cols)
    for cell in sm[4]:
        cell.font, cell.fill, cell.border = head_font, PatternFill("solid", fgColor=teal), border
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    L = "'Violation Log'!"
    cats = [c for c in order if c in set(log["category"])]
    for i, cat in enumerate(cats, start=5):
        sm.append([cat,
                   f'=COUNTIF({L}$C:$C,$A{i})',
                   f'=COUNTIFS({L}$C:$C,$A{i},{L}$F:$F,"Closed")',
                   f'=COUNTIFS({L}$C:$C,$A{i},{L}$F:$F,"Open")',
                   f'=COUNTIFS({L}$C:$C,$A{i},{L}$J:$J,"Open - overdue")',
                   f'=COUNTIFS({L}$C:$C,$A{i},{L}$J:$J,"Closed on time")',
                   f'=IF(C{i}=0,"-",F{i}/C{i})'])
    t = 5 + len(cats)
    sm.append(["Total"] + [f"=SUM({c}5:{c}{t - 1})" for c in "BCDEF"] + [f'=IF(C{t}=0,"-",F{t}/C{t})'])
    for row in sm.iter_rows(min_row=5, max_row=t):
        for cell in row:
            cell.border = border
        row[6].number_format = "0%"
    for cell in sm[t]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="EEF2F7")
    for col, w in zip("ABCDEFG", [38, 10, 10, 10, 15, 15, 17]):
        sm.column_dimensions[col].width = w
    sm.row_dimensions[4].height = 30
    wb.save(path)


def log_summary(log: pd.DataFrame, order: list[str]) -> list[dict]:
    out = []
    for cat in order:
        g = log[log["category"] == cat]
        if g.empty:
            continue
        closed = (g["status"] == "Closed").sum()
        on_time = (g["closure_check"] == "Closed on time").sum()
        out.append({"category": cat, "logged": len(g), "closed": int(closed), "open": int(len(g) - closed),
                    "overdue": int((g["closure_check"] == "Open - overdue").sum()),
                    "on_time_pct": f"{100 * on_time / closed:.0f}%" if closed else "-"})
    return out


# --------------------------------------------------------------------------- render
def sparkbars(values: list[int], colour: str) -> str:
    """Tiny inline SVG bar chart (Mon..Sun) - no charting library needed."""
    w, h, gap = 168, 46, 6
    bw = (w - gap * 6) / 7
    top = max(values) or 1
    parts = [f'<svg viewBox="0 0 {w} {h + 14}" width="{w}" height="{h + 14}" role="img">']
    for i, v in enumerate(values):
        bh = max(2, (h - 10) * v / top) if v else 2
        x = i * (bw + gap)
        fill = colour if v else "#D5DBE3"
        parts.append(f'<rect x="{x:.1f}" y="{h - bh:.1f}" width="{bw:.1f}" height="{bh:.1f}" rx="2" fill="{fill}"/>')
        if v:
            parts.append(f'<text x="{x + bw / 2:.1f}" y="{h - bh - 3:.1f}" font-size="8.5" text-anchor="middle" fill="#5B6577">{v}</text>')
        parts.append(f'<text x="{x + bw / 2:.1f}" y="{h + 11}" font-size="8.5" text-anchor="middle" fill="#8A93A3">{"MTWTFSS"[i]}</text>')
    parts.append("</svg>")
    return "".join(parts)


def find_chrome() -> str | None:
    for c in CHROME_CANDIDATES:
        if Path(c).exists() or shutil.which(c):
            return c
    return None


def chrome(args: list[str]) -> None:
    exe = find_chrome()
    if not exe:
        raise SystemExit("Chrome / Chromium not found - run with --no-pdf, or install Chrome.")
    subprocess.run([exe, "--headless=new", "--disable-gpu", "--hide-scrollbars",
                    "--no-first-run", "--disable-extensions", *args],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=120)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--config", default=str(ROOT / "config" / "thresholds.yaml"))
    ap.add_argument("--no-pdf", action="store_true", help="skip PDF / PNG export")
    args = ap.parse_args()

    cfg = yaml.safe_load(Path(args.config).read_text(encoding="utf-8"))
    end = date.fromisoformat(str(cfg["report"]["week_ending"]))
    this = Week(end - timedelta(days=6), end)
    prev = Week(end - timedelta(days=13), end - timedelta(days=7))
    run_date = end + timedelta(days=1)

    sections, log_rows = build_sections(cfg, this, prev)
    log = build_log(log_rows, run_date)
    OUT.mkdir(exist_ok=True)
    DOCS.mkdir(exist_ok=True)
    write_log_xlsx(log, OUT / "violation_log.xlsx", cfg["report"]["plant_name"], run_date,
                   [s.title for s in sections])

    palette = {"green": "#1E9E6A", "amber": "#E0A100", "red": "#D64545"}
    for s in sections:
        s.chart = sparkbars(s.daily, palette[s.status])
    counts = {k: sum(s.status == k for s in sections) for k in palette}
    open_items = int((log["status"] == "Open").sum())
    overdue = int((log["closure_check"] == "Open - overdue").sum())

    env = Environment(loader=FileSystemLoader(ROOT / "templates"), autoescape=select_autoescape(["html", "j2"]),
                      extensions=["jinja2.ext.do"])
    html = env.get_template("report.html.j2").render(
        cfg=cfg, plant=cfg["report"]["plant_name"], this=this, prev=prev, run_date=run_date,
        sections=sections, pages=[sections[0:3], sections[3:6], sections[6:7]],
        actions=top_actions(sections), counts=counts, palette=palette,
        log_summary=log_summary(log, [s.title for s in sections]), log_total=len(log), open_items=open_items, overdue=overdue,
        close_days=CLOSE_WITHIN_DAYS,
        css=(ROOT / "templates" / "report.css").read_text(encoding="utf-8"),
        generated=datetime.now().strftime("%d %b %Y %H:%M"),
    )
    html_path = OUT / "weekly_compliance_report.html"
    html_path.write_text(html, encoding="utf-8")
    print(f"HTML  -> {html_path.relative_to(ROOT)}")
    print(f"Excel -> output/violation_log.xlsx  ({len(log)} entries, {open_items} open, {overdue} overdue)")

    if not args.no_pdf:
        pdf = OUT / "sample_weekly_compliance_report.pdf"
        png = DOCS / "report_preview.png"
        url = html_path.as_uri()
        chrome(["--no-pdf-header-footer", f"--print-to-pdf={pdf}", url])
        # A4 at 96 dpi is 794 x 1123 px; scale 1.5625 gives a crisp 1240 x 1754 picture of page 1.
        chrome(["--window-size=794,1122", "--force-device-scale-factor=1.5625", f"--screenshot={png}", url])
        print(f"PDF   -> {pdf.relative_to(ROOT)}")
        print(f"PNG   -> {png.relative_to(ROOT)}")

    print("\nThis week at a glance:")
    for s in sections:
        sign = "+" if s.change > 0 else ""
        print(f"  {s.number}. {s.title:<40} {s.count:>4}  (last week {s.prev_count:>3}, {sign}{s.change})  {s.status.upper()}")


if __name__ == "__main__":
    main()
