# -*- coding: utf-8 -*-
"""Generate an HTML dashboard from the GEU AI attendance-error ledger.

Reads the per-country detail sheets (法国 / 荷兰 / 意大利) of the ledger,
aggregating error quantities by (周期, 差错类型), then overlays populated
values from the summary sheet when present. Emits a self-contained HTML file
with one interactive chart per country:
  - x-axis: 周期 (period), y-axis: 差错率 (error rate, %)
  - one colored series per 差错类型
  - percentage labels on every point; hover shows 差错数量 and 考勤总数
  - summary text box at the top-right of each chart
"""
import argparse
import glob
import os
import re

import pandas as pd

BASE_DIR = r"D:\Users\Downloads"
OUT_DIR = os.path.dirname(os.path.abspath(__file__))

COUNTRY_ALIASES = {
    "法国": "法国",
    "荷兰": "荷兰",
    "意大利": "意大利",
}
COUNTRY_SUFFIXES = ("汇总", "-汇总", "汇总数据")


def resolve_country(nat):
    """Map a summary-sheet 国家 cell (incl. 'XX汇总' rows) to a country key."""
    if nat in COUNTRY_ALIASES:
        return COUNTRY_ALIASES[nat]
    for suf in COUNTRY_SUFFIXES:
        if nat.endswith(suf) and nat[: -len(suf)] in COUNTRY_ALIASES:
            return COUNTRY_ALIASES[nat[: -len(suf)]]
    return None
# Only these error types are charted. "签到表上传错漏" and "其他问题" are
# intentionally excluded from the dashboard.
ERROR_TYPES = ["签到表模板/填写问题", "工具识别错误"]
ERROR_COLORS = {
    "签到表上传错漏": "#2563eb",
    "签到表模板/填写问题": "#dc2626",  # red
    "工具识别错误": "#2563eb",          # blue
    "其他问题": "#059669",
}
WH_COLORS = ["#2563eb", "#d97706", "#059669", "#7c3aed", "#db2777", "#0891b2", "#65a30d"]

SUMMARY_TOTAL_ROWS = ["法国汇总", "荷兰汇总", "意大利-汇总"]

# Periods to leave out of the dashboard, keyed by date range (see period_key).
# e.g. {("1005", "1011")} excludes week41(1005-1011).
EXCLUDED_PERIOD_KEYS = {("1005", "1011")}

# Only the most recent N periods are charted and summarised.
MAX_PERIODS = 8

PERIOD_DATE_RE = re.compile(r"week\d+\((\d{4})-(\d{4})\)")
PERIOD_WEEK_RE = re.compile(r"week(\d+)")


def period_key(period):
    """Normalize a period label like 'week27(0713-0719)' to its date range
    (month, day) tuple so overlapping weeks from different files match up.
    Falls back to the raw label if it does not match the expected pattern.
    """
    m = PERIOD_DATE_RE.match(period)
    if m:
        return m.group(1), m.group(2)
    return period


def week_no(period):
    """Extract the week number from a label like 'week27(0713-0719)' -> 27."""
    m = PERIOD_WEEK_RE.match(period)
    return int(m.group(1)) if m else None


def find_files(base_dir=BASE_DIR):
    """Return (consolidated_ledger, [other_ledgers]) workbooks.

    The newest "各国" workbook is the authoritative consolidated ledger. "Newest"
    is decided by file modification time (users often re-save an older-numbered
    copy), with the "(N)" version number as tie-breaker. All older "各国" copies
    and any week27 files are returned as supplements; overlapping periods are
    skipped by the caller so only genuinely missing data is added.
    """
    cons = glob.glob(os.path.join(base_dir, "*各国*.xlsx"))
    others = glob.glob(os.path.join(base_dir, "*week27*.xlsx"))
    results = []
    for p in cons + others:
        try:
            xl = pd.ExcelFile(p)
        except Exception:
            continue
        if any("汇总" in str(s) for s in xl.sheet_names):
            results.append((p, xl))
    if not results:
        raise SystemExit("No matching workbook found in %s" % base_dir)

    # Version key: all parenthesised numbers in order, e.g. (3)(1) -> (3, 1),
    # base name -> (). Comparisons fall back to the revision tuple length so
    # (3)(1) sorts as newer than (3).
    def version_key(p):
        nums = [int(x) for x in re.findall(r"\((\d+)\)", os.path.basename(p))]
        return (tuple(nums), len(nums))

    is_ge = [item for item in results if "各国" in os.path.basename(item[0])]
    # Newest "各国" workbook is the primary source. Modification time wins
    # (users often edit and re-save an older "(N)" copy, so the highest "(N)"
    # is not reliably the newest); the version key only breaks ties.
    def recency_key(item):
        try:
            mtime = os.path.getmtime(item[0])
        except OSError:
            mtime = 0.0
        return (mtime, version_key(item[0]))

    is_ge.sort(key=recency_key, reverse=True)
    primary, *rest = is_ge
    others_xls = [xl for _, xl in rest]
    for p, xl in results:
        if not ("各国" in os.path.basename(p)):
            others_xls.append(xl)
    return primary[1], others_xls


def parse_detail_sheets(xl):
    """Read the per-country detail sheets into (country, warehouse, period,
    date, type, qty, total) rows. Warehouse names are normalized to lowercase
    so spelling-case differences (e.g. 'PADOVA' vs 'padova') collapse.

    Period labels are also normalized: when a row's week number disagrees with
    the date range it is labeled under (e.g. 'week38(0831-0906)' where
    0831-0906 is really week36), the row is remapped to that week number's own
    date range. Without this the tie-dated labels would merge into the wrong
    period.
    """
    records = []
    for sheet_name, cname in (("法国", "法国"), ("荷兰", "荷兰"), ("意大利", "意大利")):
        if sheet_name not in xl.sheet_names:
            continue
        df = xl.parse(sheet_name)
        for _, r in df.iterrows():
            if not isinstance(r["周期"], str):
                continue
            period = r["周期"].strip()
            wh = (str(r["仓库名称"]).strip() if isinstance(r["仓库名称"], str) else "")
            records.append(
                {
                    "country": cname,
                    "wh": wh.lower(),
                    "wh_disp": wh,
                    "period": period,
                    "pkey": period_key(period),
                    "week_no": week_no(period),
                    "date": pd.Timestamp(r["考勤日期"]) if pd.notna(r.get("考勤日期")) else None,
                    "etype": r["差错类型"] if isinstance(r["差错类型"], str) else None,
                    "qty": float(r["差错数量"]) if isinstance(r["差错数量"], (int, float)) else 0.0,
                    "total": float(r["考勤总数"]) if isinstance(r["考勤总数"], (int, float)) else 0.0,
                }
            )

    # Normalize mislabeled periods per country: a week number's canonical date
    # range is the range where that week number dominates; rows carrying the
    # same week number under a different range are remapped to the canonical one.
    by_country = {}
    for rec in records:
        by_country.setdefault(rec["country"], []).append(rec)
    for recs in by_country.values():
        range_weeks = {}   # range -> {(week_no): count}
        range_label = {}   # range -> most common full label
        label_count = {}
        for rec in recs:
            if rec["week_no"] is None:
                continue
            range_weeks.setdefault(rec["pkey"], {})
            range_weeks[rec["pkey"]][rec["week_no"]] = \
                range_weeks[rec["pkey"]].get(rec["week_no"], 0) + 1
            lc = label_count.setdefault(rec["pkey"], {})
            lc[rec["period"]] = lc.get(rec["period"], 0) + 1
        for rng, lc in label_count.items():
            range_label[rng] = max(lc, key=lc.get)
        # For each week number, pick the range where it is the strongest claimant.
        week_canon = {}
        for rng, weeks in range_weeks.items():
            total = sum(weeks.values())
            for wk, n in weeks.items():
                share = n / total
                if wk not in week_canon or share > week_canon[wk][1]:
                    week_canon[wk] = (rng, share)
        for rec in recs:
            wk = rec["week_no"]
            if wk in week_canon:
                canon = week_canon[wk][0]
                if rec["pkey"] != canon:
                    rec["pkey"] = canon
                    rec["period"] = range_label.get(canon, rec["period"])
    return records


def warehouse_roster(xls):
    """Extract the per-country warehouse roster from the summary sheet(s).

    The 仓库 column lists every warehouse a country tracks (regardless of
    whether that week had data), so the warehouse chart legend can show all of
    them, not just the ones with recorded errors. Returns {country: {lower_name:
    display_name}} so case variants collapse to one canonical spelling.
    """
    roster = {}
    for xl in xls:
        if "GEU 汇总数据" not in xl.sheet_names:
            continue
        df = xl.parse("GEU 汇总数据")
        cur = None
        for _, r in df.iterrows():
            nat = r["国家"]
            if isinstance(nat, str):
                nat = nat.strip()
                res = resolve_country(nat)
                if res is not None:
                    cur = res
            if cur is None:
                continue
            wh = r["仓库"]
            if isinstance(wh, str):
                wh = wh.strip()
                if wh and wh != "/" and not wh.endswith("汇总"):
                    m = roster.setdefault(cur, {})
                    # Prefer the original summary spelling (e.g. 'padova').
                    if wh.lower() not in m:
                        m[wh.lower()] = wh
    return roster


def build_country_data(all_records, wh_roster=None):
    """Aggregate country-level type data and warehouse-level data.

    Each record is (country, warehouse, period, date, type, qty, total).
    Rows from overlapping periods (same date range, possibly different week
    labels) are merged; warehouse attendance totals are summed across distinct
    warehouses per period. Returns a dict keyed by country with:
      periods  : sorted list of period labels (canonical label per date range)
      rows     : {period: {type: {qty, total, rate}}}   (country-level types)
      wh_rows  : {period: {warehouse: {qty, total, rate}}}  (warehouse-level)
      wh_all   : sorted union of roster + data warehouses (for the legend)
      total_sum: optional aggregate from the summary sheet
    """
    # First pass: dedup by (country, date-range) keeping the earliest record's
    # period label, and aggregate type-level and warehouse-level quantities.
    countries = {}
    seen_period = {}  # (country, pkey) -> canonical period label

    for rec in all_records:
        if rec["pkey"] in EXCLUDED_PERIOD_KEYS:
            continue
        c = rec["country"]
        cnt = countries.setdefault(c, {"periods": [], "rows": {}, "wh_rows": {}, "wh_totals": {}, "total_sum": None})
        pkey = rec["pkey"]
        if (c, pkey) not in seen_period:
            seen_period[(c, pkey)] = rec["period"]
        period = seen_period[(c, pkey)]
        if period not in cnt["rows"]:
            cnt["periods"].append(period)
            cnt["rows"][period] = {}
            cnt["wh_rows"][period] = {}
            cnt["wh_totals"][period] = {}
            # Seed every roster warehouse so the legend shows all of them,
            # even when a warehouse has no errors/attendance for this period.
            if wh_roster:
                for wn in sorted(wh_roster.get(c, ())):
                    cnt["wh_rows"][period].setdefault(wn, {"qty": 0.0, "total": 0.0, "rate": 0.0})
        # Warehouse-level attendance total (per warehouse, max across rows).
        wh = rec["wh"]
        if wh and rec["total"] > 0:
            wht = cnt["wh_totals"][period]
            wht[wh] = max(wht.get(wh, 0), rec["total"])
        etype = rec["etype"]
        if not etype or etype == "/" or rec["qty"] < 0:
            continue
        # Type-level quantity.
        row = cnt["rows"][period].get(etype)
        if row is None:
            cnt["rows"][period][etype] = {"qty": rec["qty"], "total": 0.0, "rate": 0.0}
        else:
            row["qty"] += rec["qty"]
        # Warehouse-level quantity.
        if wh:
            whrow = cnt["wh_rows"][period].get(wh)
            if whrow is None:
                cnt["wh_rows"][period][wh] = {"qty": rec["qty"], "total": 0.0, "rate": 0.0}
            else:
                whrow["qty"] += rec["qty"]
    # Second pass: compute country and warehouse totals/rates.
    for cname, cnt in countries.items():
        for p in cnt["rows"]:
            att = sum(cnt["wh_totals"][p].values())
            for r_ in cnt["rows"][p].values():
                r_["total"] = att
                r_["rate"] = (r_["qty"] / att) if att else 0.0
            for wh in cnt["wh_rows"][p]:
                wt = cnt["wh_totals"][p].get(wh, 0)
                cnt["wh_rows"][p][wh]["total"] = wt
                cnt["wh_rows"][p][wh]["rate"] = (cnt["wh_rows"][p][wh]["qty"] / wt) if wt else 0.0
        # Legend roster: roster ∪ warehouses seen in any period.
        seen_wh = {w for p in cnt["wh_rows"] for w in cnt["wh_rows"][p]}
        if wh_roster:
            seen_wh |= set(wh_roster.get(cname, ()))
        cnt["wh_all"] = sorted(seen_wh)
        # Display names: prefer the roster spelling, else the data spelling.
        disp = {}
        if wh_roster:
            disp = dict(wh_roster.get(cname, {}))
        for w in seen_wh:
            disp.setdefault(w, w)
        cnt["wh_disp"] = {k: disp[k] for k in cnt["wh_all"]}
        cnt["periods"] = sorted(set(cnt["periods"]))
    return countries


def overlay_summary(xl, countries):
    """Overlay populated values from the summary sheet (authoritative when set).

    The consolidated ledger's summary sheet is often a blank template; when a
    cell does hold a number it overrides the detail-derived value, and the
    "XX汇总" row supplies the summary box when present.
    """
    if "GEU 汇总数据" not in xl.sheet_names:
        return
    summary = xl.parse("GEU 汇总数据")
    cur = None
    for _, r in summary.iterrows():
        nat = r["国家"]
        is_agg = False
        if isinstance(nat, str):
            nat = nat.strip()
            for suf in COUNTRY_SUFFIXES:
                if nat.endswith(suf):
                    is_agg = True
                    break
            res = resolve_country(nat)
            if res is not None:
                cur = res
        if cur is None:
            continue
        name = cur
        period = str(r["周期"]) if isinstance(r["周期"], str) else None
        etype = r["差错类型"]
        qty = r["差错数量"]
        total = r["考勤总数"]
        rate = r["差错率"]
        cnt = countries.get(name)
        if cnt is None:
            continue
        if is_agg and pd.notna(qty) and pd.notna(total):
            cnt["total_sum"] = {
                "total_qty": float(qty),
                "total_att": float(total),
                "rate": float(rate) if pd.notna(rate) else (qty / total if total else 0),
            }
            continue
        if period is None or not isinstance(etype, str) or etype == "/":
            continue
        if not isinstance(qty, (int, float)) or not isinstance(total, (int, float)):
            continue
        if pd.isna(qty) or pd.isna(total):
            continue
        if period not in cnt["rows"]:
            cnt["periods"].append(period)
            cnt["rows"][period] = {}
        cnt["rows"][period][etype] = {
            "qty": float(qty),
            "total": float(total),
            "rate": float(rate) if pd.notna(rate) else (qty / total if total else 0),
        }


def fmt_rate(rate):
    """Render a 0..1 rate as a readable percentage string."""
    if rate is None:
        return ""
    pct = rate * 100
    if pct == 0:
        return "0%"
    if pct < 0.1:
        return "%.3g%%" % pct
    return "%.1f%%" % pct


def short_week(period):
    """Shorten 'week28(0706-0712)' to 'week28' for summary labels."""
    m = re.match(r"(week\d+)", period)
    return m.group(1) if m else period


def build_html(countries, src="week27"):

    def chart_section(name, data, window):
        # The shared recent-N window keeps every country on the same x-axis.
        periods = list(window)
        types = list(ERROR_TYPES)  # always show all 4 error types
        period_totals = {
            p: max((row["total"] for row in data["rows"].get(p, {}).values()), default=0)
            for p in periods
        }
        series = []
        for t in types:
            pts = []
            for p in periods:
                rows_this_period = data["rows"].get(p, {})
                if not rows_this_period:
                    # Country has no data for this period → leave a gap.
                    pts.append(None)
                    continue
                row = rows_this_period.get(t)
                if row is None:
                    # Period has data but this type had no errors → 0% point.
                    pts.append(
                        {"period": p, "rate": 0.0, "qty": 0.0, "total": period_totals.get(p, 0)}
                    )
                    continue
                pts.append(
                    {
                        "period": p,
                        "rate": row["rate"],
                        "qty": row["qty"],
                        "total": row["total"],
                    }
                )
            series.append({"name": t, "points": pts, "color": ERROR_COLORS[t]})
        # Summary text over the displayed window (same 口径 as the chart).
        latest = periods[-1] if periods else None
        summ = None
        if periods:
            tot_q = 0.0
            tot_a = 0.0
            for p in periods:
                rows = data["rows"].get(p, {})
                if not rows:
                    continue
                tot_q += sum(r["qty"] for r in rows.values())
                tot_a += max(r["total"] for r in rows.values())
            summ = {
                "total_qty": tot_q,
                "total_att": tot_a,
                "rate": (tot_q / tot_a) if tot_a else 0,
                "first": short_week(periods[0]),
                "last": short_week(latest),
                "n": len(periods),
            }
        return {
            "name": name,
            "periods": periods,
            "series": series,
            "summ": summ,
            "latest": latest,
        }

    # Shared recent-N window across all countries (keeps x-axes aligned).
    all_periods = sorted({p for c in countries.values() for p in c["periods"]})[-MAX_PERIODS:]
    charts = [chart_section(name, data, all_periods) for name, data in countries.items()]
    # Combined overview across all countries: grouped bars by country (差错总量)
    # plus per-country 差错率 lines over the same recent-N window.
    combined = {
        "periods": all_periods,
        "countries": [c["name"] for c in charts],
        "by_country": {},
        "rate_by_country": {},
        "rate": [],
        "qty": [],
        "att": [],
        "latest": all_periods[-1] if all_periods else "",
    }
    for cname, data in countries.items():
        combined["by_country"][cname] = [
            sum(r["qty"] for r in data["rows"].get(p, {}).values()) for p in all_periods
        ]
        combined["rate_by_country"][cname] = []
        for p in all_periods:
            rows = data["rows"].get(p, {})
            if not rows:
                combined["rate_by_country"][cname].append(None)
                continue
            q = sum(r["qty"] for r in rows.values())
            a = max(r["total"] for r in rows.values())
            combined["rate_by_country"][cname].append(q / a if a else 0)
    for p in all_periods:
        q = sum(combined["by_country"][cn][all_periods.index(p)] for cn in combined["countries"])
        a = 0.0
        for data in countries.values():
            rows = data["rows"].get(p, {})
            if rows:
                a += max(r["total"] for r in rows.values())
        combined["qty"].append(q)
        combined["att"].append(a)
        combined["rate"].append(q / a if a else 0)
    # Combined summary box: aggregate the whole displayed window (same 口径).
    recent = all_periods
    cq = ca = 0.0
    for p in recent:
        for data in countries.values():
            rows = data["rows"].get(p, {})
            if not rows:
                continue
            cq += sum(r["qty"] for r in rows.values())
            ca += max(r["total"] for r in rows.values())
    combined["latest_qty"] = cq
    combined["latest_att"] = ca
    combined["latest_rate"] = cq / ca if ca else 0
    combined["first"] = short_week(recent[0]) if recent else ""
    combined["last"] = short_week(all_periods[-1]) if all_periods else ""
    combined["n"] = len(recent)
    charts_json = "var CHARTS = " + json_dumps(charts) + ";\n"
    combined_json = "var COMBINED = " + json_dumps(combined) + ";\n"
    return (
        HTML_TEMPLATE.replace("@@CHARTS@@", charts_json)
        .replace("@@COMBINED@@", combined_json)
        .replace("@@SOURCE@@", src)
        .replace("@@ERROR_TYPES@@", " / ".join(ERROR_TYPES))
    )


def json_dumps(obj):
    import json

    return json.dumps(obj, ensure_ascii=False)


HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>GEU 考勤差错率趋势看板 — @@SOURCE@@</title>
<style>
  :root {
    --bg:#f4f6fa; --card:#ffffff; --ink:#1f2937; --muted:#6b7280;
    --line:#e5e7eb; --accent:#2563eb;
  }
  * { box-sizing: border-box; margin:0; padding:0; }
  body {
    font-family: "Segoe UI","Microsoft YaHei",Arial,sans-serif;
    background:var(--bg); color:var(--ink); padding:20px;
  }
  header { max-width:1200px; margin:0 auto 8px; }
  h1 { font-size:27px; font-weight:600; }
  .sub { color:var(--muted); font-size:15px; margin-top:5px; }
  .legend-hint { color:var(--muted); font-size:14px; margin-top:7px; }
  .grid { max-width:1200px; margin:0 auto; display:grid; grid-template-columns:1fr; gap:18px; margin-top:14px; }
  .card {
    background:var(--card); border:1px solid var(--line); border-radius:12px;
    padding:10px 8px 2px; box-shadow:0 1px 3px rgba(16,24,40,.06);
  }
  .card h2 {
    font-size:19px; font-weight:600; padding:6px 12px 0;
    display:flex; align-items:baseline; gap:10px;
  }
  .card h2 .cn { color:var(--accent); }
  .card h2 .en { color:var(--muted); font-size:14px; font-weight:400; }
  .card .head { display:flex; align-items:flex-start; justify-content:space-between; padding:6px 12px 0; }
  .card .head .titles { display:flex; align-items:baseline; gap:10px; }
  .card .head .cn { color:var(--accent); font-size:19px; font-weight:600; }
  .card .head .en { color:var(--muted); font-size:14px; }
  .card .summ {
    font-size:14px; line-height:1.6; text-align:right;
    border:1px solid var(--line); border-radius:8px;
    padding:7px 12px; background:#f8fafc; color:var(--ink);
    white-space:nowrap;
  }
  .card .summ b { color:var(--accent); }
  .card .summ .rate { color:var(--accent); font-weight:600; }
  .ch-row { display:grid; grid-template-columns:1fr 1fr; gap:12px; padding:4px 8px 2px; }
  .ch-box { min-width:0; }
  .ch-title {
    font-size:14px; color:var(--muted); font-weight:600;
    padding:2px 4px 0; letter-spacing:.2px;
  }
  .ch-title.t-type { color:var(--accent); }
  .ch-title.t-wh { color:#d97706; }
  .chart { width:100%; height:460px; }
  .chart-top { height:480px; margin-top:4px; }
  footer {
    max-width:1200px; margin:18px auto 0; color:var(--muted);
    font-size:14px; line-height:1.8;
  }
  footer b { color:var(--ink); }
</style>
</head>
<body>
<header>
  <h1>GEU 考勤差错率趋势看板 <span style="color:var(--muted);font-weight:400;font-size:16px;">(周期 @@SOURCE@@)</span></h1>
  <div class="sub">数据来源：GEU AI考勤工具考勤纠错台账（法国 / 荷兰 / 意大利）｜按“国家”分别展示</div>
  <div class="legend-hint">仅展示最近 8 周 · 横轴 = 周期 · 纵轴 = 差错率(%) · 按差错类型着色 · 悬停数据点可查看差错数量与考勤总数</div>
</header>
<div class="grid" id="grid">
  <div class="card" id="card-combined">
    <div class="head">
      <div class="titles"><span class="cn">GEU各区域差错总量与差错率总览</span><span class="en">GEU Overview</span></div>
      <div class="summ" id="summ-combined"></div>
    </div>
    <div class="chart chart-top" id="chart-combined"></div>
  </div>
</div>
<footer>
  <b>口径说明：</b><br/>
  1. <b>展示窗口</b>：三国统一取最近 8 个有数据周期（不足 8 周按实际），各图与汇总均以此窗口为准；某国某周无数据时折线断开、不计入汇总。<br/>
  2. <b>单点差错率</b> = 该周期该差错类型差错数量 ÷ 该周期考勤总数 × 100%，数值标注于数据点上方。<br/>
  3. <b>汇总框（近N周汇总）</b>：差错总量 = 窗口内各周差错数量之和；考勤总数 = 窗口内各周考勤总数之和；综合差错率 = 差错总量 ÷ 考勤总数 × 100%。顶部总览为三国合计，各卡片为该国口径。<br/>
  4. <b>数据来源</b>：GEU 各国 AI 考勤工具考勤纠错台账（取版本最新副本为主源，旧版仅补充缺失周期，按日期范围去重）。<br/>
  5. <b>排除周期</b>：week41(1005-1011) 按需求排除，不参与展示与汇总。
</footer>
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js" charset="utf-8"></script>
<script>
@@CHARTS@@
@@COMBINED@@
function esc(s){ return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function fmtPct(x) {
  if (x == null) return '';
  var p = x*100;
  if (p === 0) return '0%';
  if (p < 0.1) return p.toPrecision(3) + '%';
  return p.toFixed(1) + '%';
}
function build(c, idx) {
  var s = c.summ;
  var el = document.createElement('div');
  el.className = 'card';
  el.innerHTML =
    '<div class="head"><div class="titles">' +
      '<span class="cn">'+esc(c.name)+'</span><span class="en">Attendance Error Rate Trend</span>' +
    '</div>' +
    '<div class="summ"><b>近'+s.n+'周汇总 · '+esc(s.first||'')+'~'+esc(s.last||'')+'</b><br>差错总量: '+s.total_qty+' 笔 &nbsp;|&nbsp; 考勤总数: '+s.total_att+' 笔<br>综合差错率: <span class="rate">'+fmtPct(s.rate)+'</span></div>' +
    '</div>' +
    '<div class="ch-box"><div class="ch-title t-type">按差错类型</div><div class="chart" id="chart-'+idx+'a"></div></div>';
  document.getElementById('grid').appendChild(el);

  function mkTrace(s, series) {
    var xs = [], ys = [], txt = [], cd = [];
    c.periods.forEach(function(p) {
      var pt = null;
      for (var i = 0; i < series.points.length; i++) if (series.points[i] && series.points[i].period === p) { pt = series.points[i]; break; }
      xs.push(p);
      ys.push(pt ? +(pt.rate*100).toFixed(3) : null);
      txt.push(pt ? fmtPct(pt.rate) : '');
      cd.push(pt ? [pt.qty, pt.total, fmtPct(pt.rate)] : null);
    });
    return {
      type:'scatter', mode:'lines+markers+text', name: esc(series.name),
      x:xs, y:ys, text:txt, textposition:'top center', textfont:{size:12},
      line:{width:2.5, shape:'spline', smoothing:0.8, color:series.color},
      marker:{size:8, color:series.color, line:{width:1,color:'#fff'}},
      customdata: cd,
      hovertemplate:'<b>'+esc(series.name)+'</b><br>周期: %{x}<br>差错数量: %{customdata[0]}<br>考勤总数: %{customdata[1]}<br>差错率: %{customdata[2]}<extra></extra>'
    };
  }
  var layout = {
    margin:{l:56, r:16, t:16, b:48},
    paper_bgcolor:'#fff', plot_bgcolor:'#fff',
    font:{family:'Segoe UI, Microsoft YaHei, Arial', color:'#374151', size:13},
    xaxis:{title:{text:'周期', font:{size:13}}, gridcolor:'#f3f4f6', tickfont:{size:12}, automargin:true},
    yaxis:{title:{text:'差错率(%)', font:{size:13}}, gridcolor:'#f3f4f6', ticksuffix:'%',
           tickfont:{size:12}, tickformat:'.2~g', rangemode:'tozero', exponentformat:'none'},
    legend:{orientation:'h', y:-0.30, x:0.5, xanchor:'center', font:{size:12}},
    showlegend:true,
    hovermode:'closest'
  };
  Plotly.newPlot('chart-'+idx+'a', c.series.map(function(s){return mkTrace(c,s);}), layout, {responsive:true, displaylogo:false});
}

function buildCombined() {
  var c = COMBINED;
  var cns = c.countries;
  var col = {法国:'#2563eb', 荷兰:'#d97706', 意大利:'#059669'};
  // Summary box: whole displayed window across all countries.
  var s = document.getElementById('summ-combined');
  s.innerHTML = '<b>近'+c.n+'周汇总 · '+esc(c.first||'')+'~'+esc(c.last||'')+'</b><br>差错总量: '+c.latest_qty+' 笔 &nbsp;|&nbsp; 考勤总数: '+c.latest_att+' 笔<br>综合差错率: <span class="rate">'+fmtPct(c.latest_rate)+'</span>';
  var traces = [];
  // Stacked bars: 差错总量 per country per period.
  cns.forEach(function(n, i) {
    var ys = c.by_country[n];
    traces.push({
      type:'bar', name: esc(n),
      x:c.periods, y:ys,
      marker:{color:col[n]||WH_COLORS[i%WH_COLORS.length]},
      text:ys.map(function(v){ return v>0 ? String(v) : ''; }),
      textposition:'inside', textfont:{size:11, color:'#fff'},
      hovertemplate:'<b>'+esc(n)+'</b><br>周期: %{x}<br>差错总量: %{y} 笔<extra></extra>'
    });
  });
  // Per-country 差错率 lines on the secondary axis (new color set, solid spline).
  var lineCol = ['#7c3aed', '#db2777', '#0891b2'];
  cns.forEach(function(n, i) {
    var ys = c.rate_by_country[n];
    var lc = lineCol[i % lineCol.length];
    traces.push({
      type:'scatter', mode:'lines+markers', name: esc(n)+'差错率',
      x:c.periods, y:ys.map(function(r){ return r==null ? null : +(r*100).toFixed(3); }),
      yaxis:'y2',
      line:{width:3, shape:'spline', smoothing:0.8, color:lc},
      marker:{size:7, color:lc, line:{width:1,color:'#fff'}},
      text:ys.map(function(r){ return r==null ? '' : fmtPct(r); }),
      textposition:'top center', textfont:{size:11, color:'#6b7280'},
      customdata:ys.map(function(r,i2){
        return r==null ? null : [fmtPct(r), c.by_country[n][i2], c.att[i2]];
      }),
      hovertemplate:'<b>'+esc(n)+' 差错率</b><br>周期: %{x}<br>差错率: %{customdata[0]}<br>差错总量: %{customdata[1]} 笔<br>考勤总数: %{customdata[2]} 笔<extra></extra>'
    });
  });
  var layout = {
    margin:{l:56, r:56, t:16, b:48},
    paper_bgcolor:'#fff', plot_bgcolor:'#fff',
    font:{family:'Segoe UI, Microsoft YaHei, Arial', color:'#374151', size:13},
    xaxis:{title:{text:'周期', font:{size:13}}, gridcolor:'#f3f4f6', tickfont:{size:12}, automargin:true},
    yaxis:{title:{text:'差错总量(笔)', font:{size:13}}, gridcolor:'#f3f4f6', rangemode:'tozero', tickfont:{size:12}},
    yaxis2:{title:{text:'差错率(%)', font:{size:13}}, overlaying:'y', side:'right', ticksuffix:'%',
            tickformat:'.2~g', rangemode:'tozero', tickfont:{size:12}, gridcolor:'rgba(0,0,0,0)'},
    barmode:'stack',
    bargap:0.55,
    legend:{orientation:'h', y:-0.30, x:0.5, xanchor:'center', font:{size:12}},
    hovermode:'closest'
  };
  Plotly.newPlot('chart-combined', traces, layout, {responsive:true, displaylogo:false});
}
buildCombined();
CHARTS.forEach(build);
</script>
</body>
</html>
"""


def main():
    parser = argparse.ArgumentParser(description="Build GEU attendance-error dashboard.")
    parser.add_argument("--data-dir", default=os.environ.get("GEU_DATA_DIR", BASE_DIR),
                        help="Folder holding the ledger xlsx files (default: %(default)s)")
    parser.add_argument("--out", default=None,
                        help="Output HTML path (default: <script dir>/attendance_error_dashboard.html)")
    args = parser.parse_args()

    cons, others = find_files(args.data_dir)
    all_xls = [cons] + others
    records = parse_detail_sheets(cons)
    # Other (week) files may duplicate date ranges already covered by the
    # consolidated ledger; keep only records for date ranges it lacks.
    covered = {(rec["country"], rec["pkey"]) for rec in records}
    for xl in others:
        for rec in parse_detail_sheets(xl):
            if (rec["country"], rec["pkey"]) in covered:
                continue
            records.append(rec)
    roster = warehouse_roster(all_xls)
    countries = build_country_data(records, wh_roster=roster)
    overlay_summary(cons, countries)
    html = build_html(countries, src="multi-week")
    out = args.out or os.path.join(OUT_DIR, "attendance_error_dashboard.html")
    with open(out, "w", encoding="utf-8") as f:
        f.write(html)
    print("Wrote", out)
    print("Roster:", {k: sorted(v) for k, v in roster.items()})
    for name, c in countries.items():
        print(name, c["periods"], c["total_sum"], "wh_all:", c.get("wh_all"))
        for p in c["periods"]:
            print("   ", p, "types:", {k: round(v["rate"] * 100, 2) for k, v in c["rows"][p].items()})
            print("   ", p, "wh:", {k: round(v["rate"] * 100, 2) for k, v in c["wh_rows"][p].items()})
    print("charts:", len(countries))


if __name__ == "__main__":
    main()
