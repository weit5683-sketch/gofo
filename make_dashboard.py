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
ERROR_TYPES = ["签到表上传错漏", "签到表模板/填写问题", "工具识别错误", "其他问题"]
ERROR_COLORS = {
    "签到表上传错漏": "#2563eb",
    "签到表模板/填写问题": "#d97706",
    "工具识别错误": "#dc2626",
    "其他问题": "#059669",
}
WH_COLORS = ["#2563eb", "#d97706", "#059669", "#7c3aed", "#db2777", "#0891b2", "#65a30d"]

SUMMARY_TOTAL_ROWS = ["法国汇总", "荷兰汇总", "意大利-汇总"]

PERIOD_DATE_RE = re.compile(r"week\d+\((\d{4})-(\d{4})\)")


def period_key(period):
    """Normalize a period label like 'week27(0713-0719)' to its date range
    (month, day) tuple so overlapping weeks from different files match up.
    Falls back to the raw label if it does not match the expected pattern.
    """
    m = PERIOD_DATE_RE.match(period)
    if m:
        return m.group(1), m.group(2)
    return period


def find_files():
    """Return (consolidated_ledger, [other_ledgers]) workbooks.

    The consolidated multi-week ledger (各国) is authoritative; any additional
    week files (week27) may carry earlier periods not yet consolidated.
    """
    cons = pats_glob = glob.glob(os.path.join(BASE_DIR, "*各国*.xlsx"))
    others = glob.glob(os.path.join(BASE_DIR, "*week27*.xlsx"))
    results = []
    for p in cons + others:
        try:
            xl = pd.ExcelFile(p)
        except Exception:
            continue
        if any("汇总" in str(s) for s in xl.sheet_names):
            results.append(xl)
    if not results:
        raise SystemExit("No matching workbook found in %s" % BASE_DIR)
    return results[0], results[1:]


def parse_detail_sheets(xl):
    """Read the per-country detail sheets into (country, warehouse, period,
    date, type, qty, total) rows."""
    records = []
    for sheet_name, cname in (("法国", "法国"), ("荷兰", "荷兰"), ("意大利", "意大利")):
        if sheet_name not in xl.sheet_names:
            continue
        df = xl.parse(sheet_name)
        for _, r in df.iterrows():
            if not isinstance(r["周期"], str):
                continue
            records.append(
                {
                    "country": cname,
                    "wh": (str(r["仓库名称"]).strip() if isinstance(r["仓库名称"], str) else ""),
                    "period": r["周期"].strip(),
                    "pkey": period_key(r["周期"].strip()),
                    "date": pd.Timestamp(r["考勤日期"]) if pd.notna(r.get("考勤日期")) else None,
                    "etype": r["差错类型"] if isinstance(r["差错类型"], str) else None,
                    "qty": float(r["差错数量"]) if isinstance(r["差错数量"], (int, float)) else 0.0,
                    "total": float(r["考勤总数"]) if isinstance(r["考勤总数"], (int, float)) else 0.0,
                }
            )
    return records


def warehouse_roster(xls):
    """Extract the per-country warehouse roster from the summary sheet(s).

    The 仓库 column lists every warehouse a country tracks (regardless of
    whether that week had data), so the warehouse chart legend can show all of
    them, not just the ones with recorded errors.
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
                    roster.setdefault(cur, set()).add(wh)
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


def build_html(countries, src="week27"):

    def chart_section(name, data):
        periods = data["periods"]
        types = list(ERROR_TYPES)  # always show all 4 error types
        period_totals = {
            p: max((row["total"] for row in data["rows"].get(p, {}).values()), default=0)
            for p in periods
        }
        series = []
        for t in types:
            pts = []
            for p in periods:
                row = data["rows"].get(p, {}).get(t)
                if row is None:
                    # Period has no record for this type → 0-error placeholder.
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
        # Warehouse-level series: one per warehouse from the full roster so the
        # legend shows every warehouse even without recorded data for a period.
        wh_names = data.get("wh_all") or sorted(set(
            w for p in periods for w in data["wh_rows"].get(p, {}).keys()
        ))
        wh_series = []
        for i, wh in enumerate(wh_names):
            pts = []
            for p in periods:
                row = data["wh_rows"].get(p, {}).get(wh)
                if row is None:
                    pts.append({"period": p, "rate": 0.0, "qty": 0.0, "total": 0})
                    continue
                pts.append(
                    {
                        "period": p,
                        "rate": row["rate"],
                        "qty": row["qty"],
                        "total": row["total"],
                    }
                )
            wh_series.append({"name": wh, "points": pts, "color": WH_COLORS[i % len(WH_COLORS)]})
        # Summary text: use the "XX汇总" row when present, else the latest
        # period's country aggregate (consistent with the per-week ledger).
        summ = data["total_sum"]
        latest = periods[-1] if periods else None
        if summ is None and latest:
            att = next(
                (r["total"] for r in data["rows"].get(latest, {}).values()),
                0,
            )
            tot_q = sum(r["qty"] for r in data["rows"].get(latest, {}).values())
            summ = {"total_qty": tot_q, "total_att": att, "rate": (tot_q / att) if att else 0}
        return {
            "name": name,
            "periods": periods,
            "series": series,
            "wh_series": wh_series,
            "summ": summ,
            "latest": latest,
        }

    charts = [chart_section(name, data) for name, data in countries.items()]
    # Combined overview across all countries: per-period stacked bars by
    # country (差错总量) plus an overall 差错率 line, and a summary box for the
    # latest period across all countries.
    all_periods = sorted({p for c in countries.values() for p in c["periods"]})
    combined = {
        "periods": all_periods,
        "countries": [c["name"] for c in charts],
        "by_country": {},
        "rate": [],
        "qty": [],
        "att": [],
        "latest": all_periods[-1] if all_periods else "",
    }
    for cname, data in countries.items():
        combined["by_country"][cname] = [
            sum(r["qty"] for r in data["rows"].get(p, {}).values()) for p in all_periods
        ]
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
    # Combined summary box uses the latest period across all countries.
    li = combined["periods"].index(combined["latest"]) if combined["latest"] in combined["periods"] else -1
    if li >= 0:
        combined["latest_qty"] = combined["qty"][li]
        combined["latest_att"] = combined["att"][li]
        combined["latest_rate"] = combined["rate"][li]
    else:
        combined["latest_qty"] = combined["latest_att"] = 0
        combined["latest_rate"] = 0
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
  h1 { font-size:22px; font-weight:600; }
  .sub { color:var(--muted); font-size:13px; margin-top:4px; }
  .legend-hint { color:var(--muted); font-size:12px; margin-top:6px; }
  .grid { max-width:1200px; margin:0 auto; display:grid; grid-template-columns:1fr; gap:18px; margin-top:14px; }
  .card {
    background:var(--card); border:1px solid var(--line); border-radius:12px;
    padding:10px 8px 2px; box-shadow:0 1px 3px rgba(16,24,40,.06);
  }
  .card h2 {
    font-size:16px; font-weight:600; padding:6px 12px 0;
    display:flex; align-items:baseline; gap:10px;
  }
  .card h2 .cn { color:var(--accent); }
  .card h2 .en { color:var(--muted); font-size:12px; font-weight:400; }
  .card .head { display:flex; align-items:flex-start; justify-content:space-between; padding:6px 12px 0; }
  .card .head .titles { display:flex; align-items:baseline; gap:10px; }
  .card .head .cn { color:var(--accent); font-size:16px; font-weight:600; }
  .card .head .en { color:var(--muted); font-size:12px; }
  .card .summ {
    font-size:12px; line-height:1.55; text-align:right;
    border:1px solid var(--line); border-radius:8px;
    padding:6px 10px; background:#f8fafc; color:var(--ink);
    white-space:nowrap;
  }
  .card .summ b { color:var(--accent); }
  .card .summ .rate { color:var(--accent); font-weight:600; }
  .ch-row { display:grid; grid-template-columns:1fr 1fr; gap:12px; padding:4px 8px 2px; }
  .ch-box { min-width:0; }
  .ch-title {
    font-size:12px; color:var(--muted); font-weight:600;
    padding:2px 4px 0; letter-spacing:.2px;
  }
  .ch-title.t-type { color:var(--accent); }
  .ch-title.t-wh { color:#d97706; }
  .chart { width:100%; height:440px; }
  .chart-top { height:360px; margin-top:4px; }
  footer {
    max-width:1200px; margin:18px auto 0; color:var(--muted);
    font-size:12px; line-height:1.7;
  }
  footer b { color:var(--ink); }
</style>
</head>
<body>
<header>
  <h1>GEU 考勤差错率趋势看板 <span style="color:var(--muted);font-weight:400;font-size:14px;">(周期 @@SOURCE@@)</span></h1>
  <div class="sub">数据来源：GEU AI考勤工具考勤纠错台账（法国 / 荷兰 / 意大利）｜按“国家”分别展示</div>
  <div class="legend-hint">每国两图：左图按差错类型着色，右图按仓库着色 · 横轴 = 周期 · 纵轴 = 差错率(%) · 悬停数据点可查看差错数量与考勤总数</div>
</header>
<div class="grid" id="grid">
  <div class="card" id="card-combined">
    <div class="head">
      <div class="titles"><span class="cn">三国差错总量与差错率总览</span><span class="en">France · Netherlands · Italy</span></div>
      <div class="summ" id="summ-combined"></div>
    </div>
    <div class="chart chart-top" id="chart-combined"></div>
  </div>
</div>
<footer>
  <b>说明：</b>“差错率”= 差错数量 ÷ 考勤总数 × 100%，数值已标注于各数据点上方；悬停数据点可查看对应周期与维度的<b>差错数量</b>与<b>考勤总数</b>。<br/>
  左图按<b>差错类型</b>着色（签到表上传错漏 / 签到表模板/填写问题 / 工具识别错误 / 其他问题），右图按<b>仓库</b>着色。卡片右上角汇总框展示该国最近一期的<b>差错总量 / 考勤总数 / 综合差错率</b>。数据来自两份台账（各国综合台账 + week27 台账），按日期范围去重合并。部分周期台账未记录差错数据，图中仅显示有记录的周期。
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
    '<div class="summ"><b>汇总 · '+esc(c.latest||'')+'</b><br>差错总量: '+s.total_qty+' 笔 &nbsp;|&nbsp; 考勤总数: '+s.total_att+' 笔<br>综合差错率: <span class="rate">'+fmtPct(s.rate)+'</span></div>' +
    '</div>' +
    '<div class="ch-row">' +
      '<div class="ch-box"><div class="ch-title t-type">按差错类型</div><div class="chart" id="chart-'+idx+'a"></div></div>' +
      '<div class="ch-box"><div class="ch-title t-wh">按仓库</div><div class="chart" id="chart-'+idx+'b"></div></div>' +
    '</div>';
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
      x:xs, y:ys, text:txt, textposition:'top center', textfont:{size:10},
      line:{width:2.5, color:series.color}, marker:{size:8, color:series.color, line:{width:1,color:'#fff'}},
      customdata: cd,
      hovertemplate:'<b>'+esc(series.name)+'</b><br>周期: %{x}<br>差错数量: %{customdata[0]}<br>考勤总数: %{customdata[1]}<br>差错率: %{customdata[2]}<extra></extra>'
    };
  }
  var layout = {
    margin:{l:48, r:12, t:14, b:40},
    paper_bgcolor:'#fff', plot_bgcolor:'#fff',
    font:{family:'Segoe UI, Microsoft YaHei, Arial', color:'#374151', size:11},
    xaxis:{title:{text:'周期'}, gridcolor:'#f3f4f6', tickfont:{size:10}, automargin:true},
    yaxis:{title:{text:'差错率(%)'}, gridcolor:'#f3f4f6', ticksuffix:'%',
           tickformat:'.2~g', rangemode:'tozero', exponentformat:'none'},
    legend:{orientation:'h', y:-0.28, x:0.5, xanchor:'center', font:{size:10}},
    showlegend:true,
    hovermode:'closest'
  };
  Plotly.newPlot('chart-'+idx+'a', c.series.map(function(s){return mkTrace(c,s);}), layout, {responsive:true, displaylogo:false});
  Plotly.newPlot('chart-'+idx+'b', c.wh_series.map(function(s){return mkTrace(c,s);}), layout, {responsive:true, displaylogo:false});
}

function buildCombined() {
  var c = COMBINED;
  var cns = c.countries;
  var col = {法国:'#2563eb', 荷兰:'#d97706', 意大利:'#059669'};
  // Summary box: latest period across all countries.
  var s = document.getElementById('summ-combined');
  s.innerHTML = '<b>汇总 · '+esc(c.latest)+'</b><br>差错总量: '+c.latest_qty+' 笔 &nbsp;|&nbsp; 考勤总数: '+c.latest_att+' 笔<br>综合差错率: <span class="rate">'+fmtPct(c.latest_rate)+'</span>';
  // Stacked bars: 差错总量 per country per period; line: overall 差错率.
  var traces = cns.map(function(n, i) {
    var ys = c.by_country[n];
    return {
      type:'bar', name: esc(n),
      x:c.periods, y:ys,
      marker:{color:col[n]||WH_COLORS[i%WH_COLORS.length]},
      text:ys.map(function(v){ return v>0 ? String(v) : ''; }),
      textposition:'inside', textfont:{size:10, color:'#fff'},
      hovertemplate:'<b>'+esc(n)+'</b><br>周期: %{x}<br>差错总量: %{y} 笔<extra></extra>'
    };
  });
  traces.push({
    type:'scatter', mode:'lines+markers', name:'综合差错率',
    x:c.periods, y:c.rate.map(function(r){return +(r*100).toFixed(3);}),
    yaxis:'y2',
    line:{width:3, color:'#7c3aed'},
    marker:{size:8, color:'#7c3aed', line:{width:1,color:'#fff'}},
    text:c.rate.map(fmtPct), textposition:'top center', textfont:{size:10, color:'#7c3aed'},
    customdata:c.rate.map(function(r,i){ return [fmtPct(r), c.qty[i], c.att[i]]; }),
    hovertemplate:'<b>综合差错率</b><br>周期: %{x}<br>差错率: %{customdata[0]}<br>差错总量: %{customdata[1]} 笔<br>考勤总数: %{customdata[2]} 笔<extra></extra>'
  });
  var layout = {
    margin:{l:48, r:44, t:14, b:40},
    paper_bgcolor:'#fff', plot_bgcolor:'#fff',
    font:{family:'Segoe UI, Microsoft YaHei, Arial', color:'#374151', size:11},
    xaxis:{title:{text:'周期'}, gridcolor:'#f3f4f6', tickfont:{size:10}, automargin:true},
    yaxis:{title:{text:'差错总量(笔)'}, gridcolor:'#f3f4f6', rangemode:'tozero'},
    yaxis2:{title:{text:'差错率(%)'}, overlaying:'y', side:'right', ticksuffix:'%',
            tickformat:'.2~g', rangemode:'tozero', gridcolor:'rgba(0,0,0,0)'},
    barmode:'stack',
    legend:{orientation:'h', y:-0.28, x:0.5, xanchor:'center', font:{size:10}},
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
    cons, others = find_files()
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
    out = os.path.join(OUT_DIR, "attendance_error_dashboard.html")
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
