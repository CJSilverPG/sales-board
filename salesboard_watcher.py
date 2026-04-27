#!/usr/bin/env python3
"""
Preserve Gold — Sales Board Auto-Updater
=========================================
Watches Sales_Board.xlsx for changes and regenerates sales_intranet.html
with full branding, KPI cards, bar charts, and correct tab layout.

SETUP (one time):
    pip install pandas openpyxl watchdog

RUN:
    python sales_board_watcher.py

Keep this terminal open while working. Press Ctrl+C to stop.
Both files should be in the same folder as this script.
Change XLSX_PATH / OUTPUT_PATH below if needed.
"""

import time, sys, math
from pathlib import Path
from datetime import datetime

XLSX_PATH   = Path(__file__).parent / "Sales_Board.xlsx"
OUTPUT_PATH = Path(__file__).parent / "index.html"

print(f"XLSX  : {XLSX_PATH.resolve()}")
print(f"OUTPUT: {OUTPUT_PATH.resolve()}")

# ── Formatting helpers ─────────────────────────────────────────

def _f(v):
    try: return float(v)
    except: return None

def fmt(v):
    v = _f(v)
    if v is None or math.isnan(v): return "—"
    if abs(v) >= 1e6: return f"${v/1e6:.2f}M"
    if abs(v) >= 1e3: return f"${v:,.0f}"
    return f"${v:.0f}"

def fmtK(v):
    v = _f(v)
    if v is None or math.isnan(v) or v == 0: return "—"
    return f"${round(v/1000)}K"

def fmtN(v, dec=0):
    v = _f(v)
    if v is None or math.isnan(v): return "—"
    return f"{v:,.{dec}f}"

def fmtPct(v):
    v = _f(v)
    if v is None or math.isnan(v): return "—"
    sign = "+" if v > 0 else ""
    return f"{sign}{v:.1f}%"

def safe_str(v):
    if v is None: return ""
    try:
        if math.isnan(float(v)): return ""
    except: pass
    return str(v).strip()

# ── Bar chart HTML ─────────────────────────────────────────────

def bar_chart(rows, max_val):
    """rows = list of (name, value) tuples"""
    html = ""
    for name, val in rows:
        pct = max(2, (val / max_val) * 100) if max_val else 2
        html += f"""<div class="bar-row">
          <div class="bar-name">{name}</div>
          <div class="bar-track"><div class="bar-fill" style="width:{pct:.1f}%"></div></div>
          <div class="bar-val">{fmt(val)}</div>
        </div>"""
    return html

# ── Data readers ───────────────────────────────────────────────

def read_current_month(xl):
    import pandas as pd
    df = xl.get("Sales Data Current Month")
    if df is None: return [], {}

    # Find header row (contains "Rep")
    header_idx = None
    for i, row in df.iterrows():
        if any(str(v).strip() == "Rep" for v in row):
            header_idx = i
            break
    if header_idx is None: return [], {}

    df.columns = [safe_str(c) for c in df.iloc[header_idx]]
    df = df.iloc[header_idx + 1:].reset_index(drop=True)
    df = df[df.iloc[:, 0].notna()]

    reps = []
    totals = {}
    for _, row in df.iterrows():
        rep = safe_str(row.get("Rep", ""))
        if not rep: continue
        if rep.lower() == "totals":
            totals = {
                "sales": _f(row.get("Full Sales")) or 0,
                "pipe":  _f(row.get("Pipeline")) or 0,
                "deals": _f(row.get("Deals")) or 0,
                "calls": _f(row.get("Calls MTD")) or 0,
                "cpd":   _f(row.get("CPD Adjusted")) or 0,
                "tt":    _f(row.get("Talk Time MTD")) or 0,
                "ttpd":  _f(row.get("TTPD Adjusted")) or 0,
            }
            continue
        reps.append({
            "rep":   rep,
            "sales": _f(row.get("Full Sales")) or 0,
            "pipe":  _f(row.get("Pipeline")) or 0,
            "deals": _f(row.get("Deals")) or 0,
            "calls": _f(row.get("Calls MTD")) or 0,
            "cpd":   _f(row.get("CPD Adjusted")) or 0,
            "met":   safe_str(row.get("Call Min Met", "")),
            "tt":    _f(row.get("Talk Time MTD")) or 0,
            "ttpd":  _f(row.get("TTPD Adjusted")) or 0,
        })

    reps.sort(key=lambda r: r["sales"], reverse=True)
    return reps, totals


def read_ytd(xl):
    df = xl.get("YTD 2026")
    if df is None: return [], {}

    # Row 0 is header
    df.columns = [safe_str(c) for c in df.iloc[0]]
    df = df.iloc[1:].reset_index(drop=True)

    reps = []
    company = {}
    for _, row in df.iterrows():
        name = safe_str(row.get("2026", ""))
        if not name: continue
        if name.lower() == "company total":
            company = {
                "jan":  _f(row.get("January")) or 0,
                "feb":  _f(row.get("February")) or 0,
                "ytd":  _f(row.get("YTD Total")) or 0,
                "proj": _f(row.get("Year End Projection")) or 0,
            }
            continue
        reps.append({
            "rep":  name,
            "jan":  _f(row.get("January")) or 0,
            "feb":  _f(row.get("February")) or 0,
            "mar":  _f(row.get("March")) or 0,
            "avg":  _f(row.get("Monthly Average")) or 0,
            "ytd":  _f(row.get("YTD Total")) or 0,
            "proj": _f(row.get("Year End Projection")) or 0,
        })

    reps.sort(key=lambda r: r["ytd"], reverse=True)
    return reps, company


def read_mom2025(xl):
    df = xl.get("MOM 2025")
    if df is None: return [], {}

    df.columns = [safe_str(c) for c in df.iloc[0]]
    df = df.iloc[1:].reset_index(drop=True)

    months = ["January","February","March","April","May","June",
              "July","August","September","October","November","December"]

    reps = []
    total_row = {}
    for _, row in df.iterrows():
        name = safe_str(row.get("2025", ""))
        if not name: continue
        if name.lower() == "total":
            total_row = {m: _f(row.get(m)) or 0 for m in months}
            continue
        vals = {m: _f(row.get(m)) or 0 for m in months}
        tot  = _f(row.get("Monthly Average"))  # actually YTD total column varies
        # sum from months as fallback
        ytd_total = _f(row.get("NaN")) or sum(vals.values())
        reps.append({"rep": name, **vals, "tot": sum(vals.values())})

    reps.sort(key=lambda r: r["tot"], reverse=True)
    return reps, total_row


# ── HTML sections ──────────────────────────────────────────────

def section_current(reps, totals):
    total_sales = totals.get("sales", sum(r["sales"] for r in reps))
    total_pipe  = totals.get("pipe",  sum(r["pipe"]  for r in reps))
    total_deals = totals.get("deals", sum(r["deals"] for r in reps))
    total_calls = totals.get("calls", sum(r["calls"] for r in reps))
    met_count   = sum(1 for r in reps if "✅" in r.get("met",""))
    total_reps  = len(reps)

    # KPI cards
    kpis = f"""<div class="kpi-row">
      <div class="kpi"><div class="kpi-label">Total Full Sales</div><div class="kpi-value">{fmt(total_sales)}</div><div class="kpi-sub">{total_reps} reps active</div></div>
      <div class="kpi navy"><div class="kpi-label">Total Pipeline</div><div class="kpi-value">{fmt(total_pipe)}</div><div class="kpi-sub">Across all reps</div></div>
      <div class="kpi soft"><div class="kpi-label">Total Deals</div><div class="kpi-value">{fmtN(total_deals,2)}</div><div class="kpi-sub">Combined deal count</div></div>
      <div class="kpi"><div class="kpi-label">Total Calls MTD</div><div class="kpi-value">{fmtN(total_calls)}</div><div class="kpi-sub">Across all reps</div></div>
      <div class="kpi navy"><div class="kpi-label">Call Min Met</div><div class="kpi-value">{met_count} / {total_reps}</div><div class="kpi-sub">Reps hitting target</div></div>
    </div>"""

    # Bar charts
    top_sales = [(r["rep"], r["sales"]) for r in reps if r["sales"] > 0][:10]
    top_pipe  = sorted([(r["rep"], r["pipe"]) for r in reps if r["pipe"] > 0], key=lambda x: -x[1])[:10]
    max_s = top_sales[0][1] if top_sales else 1
    max_p = top_pipe[0][1]  if top_pipe  else 1

    charts = f"""<div class="two-col">
      <div class="card"><div class="card-title">Top Producers <span class="badge">Full Sales</span></div>
        <div class="bar-chart">{bar_chart(top_sales, max_s)}</div></div>
      <div class="card"><div class="card-title">Pipeline Leaders <span class="badge">Open Pipeline</span></div>
        <div class="bar-chart">{bar_chart(top_pipe, max_p)}</div></div>
    </div>"""

    # Table rows
    rows = ""
    for r in reps:
        mc = "check" if "✅" in r.get("met","") else "cross"
        met_icon = r.get("met","—")
        rows += f"""<tr>
          <td class="name-cell">{r['rep']}</td>
          <td class="num">{fmt(r['sales'])}</td>
          <td class="num">{fmt(r['pipe'])}</td>
          <td class="num">{fmtN(r['deals'],2)}</td>
          <td class="num">{fmtN(r['calls'])}</td>
          <td class="num">{fmtN(r['cpd'],1)}</td>
          <td class="num {mc}">{met_icon}</td>
          <td class="num">{fmtN(r['tt'])}</td>
          <td class="num">{fmtN(r['ttpd'],1)}</td>
        </tr>"""

    rows += f"""<tr class="totals-row">
      <td class="name-cell">Totals</td>
      <td class="num">{fmt(total_sales)}</td>
      <td class="num">{fmt(total_pipe)}</td>
      <td class="num">{fmtN(total_deals,2)}</td>
      <td class="num">{fmtN(total_calls)}</td>
      <td class="num">—</td><td class="num">—</td>
      <td class="num">{fmt(totals.get('tt',0))}</td>
      <td class="num">{fmtN(totals.get('ttpd',0),1)}</td>
    </tr>"""

    table = f"""<div class="card">
      <div class="card-title">All Reps — Current Period <span class="badge">Full Detail</span></div>
      <div class="table-wrap"><table>
        <thead><tr>
          <th>Rep</th><th class="num">Full Sales</th><th class="num">Pipeline</th>
          <th class="num">Deals</th><th class="num">Calls MTD</th><th class="num">CPD Adj.</th>
          <th class="num">Call Min</th><th class="num">Talk Time</th><th class="num">TTPD Adj.</th>
        </tr></thead>
        <tbody>{rows}</tbody>
      </table></div>
    </div>"""

    return kpis + charts + table


def section_ytd(reps, company):
    top12 = reps[:12]
    max_ytd = top12[0]["ytd"] if top12 else 1

    kpis = f"""<div class="kpi-row">
      <div class="kpi"><div class="kpi-label">Company YTD Total</div><div class="kpi-value">{fmt(company.get('ytd',0))}</div><div class="kpi-sub">Year to date</div></div>
      <div class="kpi navy"><div class="kpi-label">Year-End Projection</div><div class="kpi-value">{fmt(company.get('proj',0))}</div><div class="kpi-sub">At current pace</div></div>
      <div class="kpi soft"><div class="kpi-label">Top Rep YTD</div><div class="kpi-value">{fmt(reps[0]['ytd']) if reps else '—'}</div><div class="kpi-sub">{reps[0]['rep'] if reps else '—'}</div></div>
    </div>"""

    chart = f"""<div class="card">
      <div class="card-title">YTD Rankings <span class="badge">2026</span></div>
      <div class="bar-chart">{bar_chart([(r['rep'],r['ytd']) for r in top12], max_ytd)}</div>
    </div>"""

    rows = ""
    for r in reps:
        rows += f"""<tr>
          <td class="name-cell">{r['rep']}</td>
          <td class="num">{fmt(r['jan'])}</td>
          <td class="num">{fmt(r['feb'])}</td>
          <td class="num" style="color:var(--muted)">{"—" if r['mar']==0 else fmt(r['mar'])}</td>
          <td class="num">{fmt(r['avg'])}</td>
          <td class="num">{fmt(r['ytd'])}</td>
          <td class="num">{fmt(r['proj'])}</td>
        </tr>"""

    rows += f"""<tr class="totals-row">
      <td class="name-cell">Company Total</td>
      <td class="num">{fmt(company.get('jan',0))}</td>
      <td class="num">{fmt(company.get('feb',0))}</td>
      <td class="num">—</td>
      <td class="num">{fmt(company.get('ytd',0))}</td>
      <td class="num">{fmt(company.get('ytd',0))}</td>
      <td class="num">{fmt(company.get('proj',0))}</td>
    </tr>"""

    table = f"""<div class="card">
      <div class="card-title">Full YTD Table <span class="badge">All Reps</span></div>
      <div class="table-wrap"><table>
        <thead><tr>
          <th>Rep</th><th class="num">January</th><th class="num">February</th>
          <th class="num">March</th><th class="num">Monthly Avg</th>
          <th class="num">YTD Total</th><th class="num">Year-End Proj.</th>
        </tr></thead>
        <tbody>{rows}</tbody>
      </table></div>
    </div>"""

    return kpis + chart + table


def section_mom2025(reps, total_row):
    months = ["January","February","March","April","May","June",
              "July","August","September","October","November","December"]
    abbr   = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

    top12   = reps[:12]
    max_tot = top12[0]["tot"] if top12 else 1

    company_total = sum(total_row.get(m, 0) for m in months)

    kpis = f"""<div class="kpi-row">
      <div class="kpi"><div class="kpi-label">2025 Company Total</div><div class="kpi-value">{fmt(company_total)}</div><div class="kpi-sub">Full year</div></div>
      <div class="kpi navy"><div class="kpi-label">Monthly Average</div><div class="kpi-value">{fmt(company_total/12)}</div><div class="kpi-sub">Per month avg</div></div>
      <div class="kpi soft"><div class="kpi-label">Top Rep 2025</div><div class="kpi-value">{fmt(reps[0]['tot']) if reps else '—'}</div><div class="kpi-sub">{reps[0]['rep'] if reps else '—'}</div></div>
    </div>"""

    chart = f"""<div class="card">
      <div class="card-title">2025 Full Year Rankings <span class="badge">Total</span></div>
      <div class="bar-chart">{bar_chart([(r['rep'],r['tot']) for r in top12], max_tot)}</div>
    </div>"""

    header_cells = "".join(f'<th class="num">{a}</th>' for a in abbr)
    rows = ""
    for r in reps:
        cells = "".join(f'<td class="num">{fmtK(r[m])}</td>' for m in months)
        rows += f'<tr><td class="name-cell">{r["rep"]}</td>{cells}<td class="num">{fmt(r["tot"])}</td></tr>'

    total_cells = "".join(f'<td class="num">{fmtK(total_row.get(m,0))}</td>' for m in months)
    rows += f'<tr class="totals-row"><td class="name-cell">Company Total</td>{total_cells}<td class="num">{fmt(company_total)}</td></tr>'

    table = f"""<div class="card">
      <div class="card-title">Monthly Breakdown — All Reps <span class="badge">2025</span></div>
      <div class="table-wrap"><table>
        <thead><tr><th>Rep</th>{header_cells}<th class="num">Total</th></tr></thead>
        <tbody>{rows}</tbody>
      </table></div>
    </div>"""

    return kpis + chart + table


# ── Master build ───────────────────────────────────────────────

def build_html(xlsx_path: Path) -> str:
    import pandas as pd
    xl = pd.read_excel(xlsx_path, sheet_name=None, header=None)

    now = datetime.now().strftime("%B %d, %Y  %I:%M %p")

    cur_reps,  cur_totals  = read_current_month(xl)
    ytd_reps,  ytd_company = read_ytd(xl)
    m25_reps,  m25_totals  = read_mom2025(xl)

    page_current = section_current(cur_reps, cur_totals)
    page_ytd     = section_ytd(ytd_reps, ytd_company)
    page_mom2025 = section_mom2025(m25_reps, m25_totals)

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Preserve Gold — Sales Board</title>
<link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@600;700;800&family=Source+Sans+3:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
  :root{{--navy:#0d1b3e;--navy-mid:#152350;--navy-light:#1e3068;--gold:#c9a84c;--gold-light:#e2c97e;--gold-pale:#f5e9c8;--bg:#f0f2f6;--surface:#fff;--surface2:#f7f8fb;--border:#dde2ed;--text:#0d1b3e;--text-mid:#3a4a6b;--muted:#7a869e;--display:'Playfair Display',Georgia,serif;--sans:'Source Sans 3',sans-serif;--mono:'DM Mono',monospace}}
  *{{box-sizing:border-box;margin:0;padding:0}}
  body{{background:var(--bg);color:var(--text);font-family:var(--sans);font-size:14px;line-height:1.6;min-height:100vh}}
  .top-bar{{background:linear-gradient(90deg,var(--navy) 0%,var(--gold) 50%,var(--navy) 100%);height:3px}}
  header{{background:var(--navy);padding:0 40px;display:flex;align-items:center;justify-content:space-between;height:68px;position:sticky;top:0;z-index:100;box-shadow:0 2px 20px rgba(0,0,0,.3)}}
  .logo{{display:flex;align-items:center;gap:13px}}
  .logo-shield{{width:40px;height:40px;background:linear-gradient(145deg,var(--gold) 0%,var(--gold-light) 100%);border-radius:6px;display:flex;align-items:center;justify-content:center;font-size:20px;box-shadow:0 2px 10px rgba(201,168,76,.5)}}
  .logo-text{{display:flex;flex-direction:column;line-height:1.15}}
  .logo-name{{font-family:var(--display);font-weight:700;font-size:20px;color:#fff}}
  .logo-sub{{font-family:var(--mono);font-size:9px;letter-spacing:2.5px;text-transform:uppercase;color:var(--gold-light)}}
  .header-right{{display:flex;align-items:center;gap:16px}}
  .header-badge{{background:rgba(201,168,76,.15);border:1px solid rgba(201,168,76,.35);color:var(--gold-light);font-family:var(--mono);font-size:10px;letter-spacing:1px;text-transform:uppercase;padding:5px 12px;border-radius:4px}}
  .header-date{{font-family:var(--mono);font-size:11px;color:rgba(255,255,255,.35)}}
  .update-bar{{background:var(--navy-mid);padding:9px 40px;display:flex;align-items:center;gap:10px;font-family:var(--mono);font-size:11px;color:var(--gold-light);letter-spacing:.5px;border-bottom:1px solid rgba(201,168,76,.2)}}
  .pulse-dot{{width:7px;height:7px;border-radius:50%;background:var(--gold);box-shadow:0 0 8px var(--gold);animation:pulse 2.5s ease-in-out infinite;flex-shrink:0}}
  @keyframes pulse{{0%,100%{{opacity:1;transform:scale(1)}}50%{{opacity:.35;transform:scale(.75)}}}}
  nav{{background:var(--surface);border-bottom:1px solid var(--border);padding:0 40px;display:flex;overflow-x:auto;box-shadow:0 1px 6px rgba(0,0,0,.06)}}
  nav button{{background:none;border:none;color:var(--muted);font-family:var(--sans);font-weight:600;font-size:13px;padding:16px 22px;cursor:pointer;transition:all .15s;border-bottom:3px solid transparent;margin-bottom:-1px;white-space:nowrap}}
  nav button:hover{{color:var(--navy)}}
  nav button.active{{color:var(--navy);border-bottom-color:var(--gold)}}
  main{{padding:36px 40px;max-width:1500px;margin:0 auto}}
  .page{{display:none}}.page.active{{display:block}}
  .page-header{{margin-bottom:32px;padding-bottom:22px;border-bottom:1px solid var(--border)}}
  .page-header h1{{font-family:var(--display);font-weight:700;font-size:30px;color:var(--navy);margin-bottom:5px}}
  .page-header p{{color:var(--muted);font-size:12px;font-family:var(--mono);text-transform:uppercase;letter-spacing:.5px}}
  .gold-rule{{width:36px;height:3px;background:linear-gradient(90deg,var(--gold),var(--gold-light));border-radius:2px;margin-top:12px}}
  .kpi-row{{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:14px;margin-bottom:28px}}
  .kpi{{background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:20px 22px;position:relative;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.05)}}
  .kpi::before{{content:'';position:absolute;top:0;left:0;right:0;height:3px;background:var(--gold)}}
  .kpi.navy::before{{background:var(--navy)}}.kpi.soft::before{{background:var(--gold-light)}}
  .kpi-label{{font-family:var(--mono);font-size:10px;text-transform:uppercase;letter-spacing:1.2px;color:var(--muted);margin-bottom:10px}}
  .kpi-value{{font-family:var(--display);font-weight:700;font-size:26px;letter-spacing:-.5px;color:var(--navy);line-height:1}}
  .kpi-sub{{font-size:12px;color:var(--muted);margin-top:6px}}
  .card{{background:var(--surface);border:1px solid var(--border);border-radius:8px;overflow:hidden;margin-bottom:24px;box-shadow:0 1px 5px rgba(0,0,0,.05)}}
  .card-title{{padding:14px 20px;border-bottom:1px solid var(--border);font-family:var(--display);font-weight:600;font-size:15px;color:var(--navy);display:flex;align-items:center;gap:10px;background:var(--surface2)}}
  .badge{{font-family:var(--mono);font-size:10px;background:var(--navy);color:var(--gold-light);padding:2px 9px;border-radius:3px;letter-spacing:.5px}}
  .table-wrap{{overflow-x:auto}}
  table{{width:100%;border-collapse:collapse}}
  thead th{{background:var(--navy);color:rgba(255,255,255,.65);padding:10px 16px;text-align:left;font-family:var(--mono);font-size:10px;text-transform:uppercase;letter-spacing:.8px;white-space:nowrap;font-weight:500}}
  thead th.num{{text-align:right}}
  thead th:first-child{{color:var(--gold-light)}}
  tbody tr{{border-bottom:1px solid var(--border);transition:background .1s}}
  tbody tr:last-child{{border-bottom:none}}
  tbody tr:hover{{background:#f2f5fc}}
  tbody td{{padding:10px 16px;font-size:13px;white-space:nowrap}}
  tbody td.num{{text-align:right;font-family:var(--mono);font-size:12px;color:var(--text-mid)}}
  .name-cell{{font-weight:600;color:var(--navy)}}
  .check{{color:#2e7d32}}.cross{{color:#b71c1c}}
  .totals-row td{{font-weight:700!important;background:var(--navy)!important;color:var(--gold-light)!important;border-top:2px solid var(--gold)!important;font-family:var(--mono);font-size:12px}}
  .bar-chart{{padding:18px 22px;display:flex;flex-direction:column;gap:10px}}
  .bar-row{{display:grid;grid-template-columns:155px 1fr 105px;align-items:center;gap:12px}}
  .bar-name{{font-size:12px;font-weight:600;color:var(--text-mid);text-align:right;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}
  .bar-track{{background:var(--bg);border-radius:3px;height:20px;overflow:hidden;border:1px solid var(--border)}}
  .bar-fill{{height:100%;border-radius:3px;background:linear-gradient(90deg,var(--navy) 0%,var(--navy-light) 55%,var(--gold) 100%);transition:width .9s cubic-bezier(.4,0,.2,1)}}
  .bar-val{{font-family:var(--mono);font-size:11px;color:var(--muted);text-align:right}}
  .two-col{{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:24px}}
  ::-webkit-scrollbar{{width:5px;height:5px}}::-webkit-scrollbar-track{{background:var(--bg)}}::-webkit-scrollbar-thumb{{background:var(--border);border-radius:3px}}
  @media(max-width:900px){{.two-col{{grid-template-columns:1fr}}main,header,nav,.update-bar{{padding-left:16px;padding-right:16px}}}}
</style>
</head>
<body>
<div class="top-bar"></div>
<header>
  <div class="logo">
    <div class="logo-shield">🦅</div>
    <div class="logo-text">
      <span class="logo-name">Preserve Gold</span>
      <span class="logo-sub">Sales Intelligence</span>
    </div>
  </div>
  <div class="header-right">
    <span class="header-badge">Internal Only</span>
    <span class="header-date">{now}</span>
  </div>
</header>
<div class="update-bar">
  <div class="pulse-dot"></div>
  Last generated from workbook: {now}
</div>
<nav>
  <button class="active" onclick="showPage('current',this)">Current Month</button>
  <button onclick="showPage('ytd',this)">YTD 2026</button>
  <button onclick="showPage('mom2025',this)">MOM 2025</button>
</nav>
<main>

<div class="page active" id="current">
  <div class="page-header"><h1>Current Month</h1><p>Sales Data — Active Period</p><div class="gold-rule"></div></div>
  {page_current}
</div>

<div class="page" id="ytd">
  <div class="page-header"><h1>YTD 2026</h1><p>Year-to-date performance by representative</p><div class="gold-rule"></div></div>
  {page_ytd}
</div>

<div class="page" id="mom2025">
  <div class="page-header"><h1>Month-over-Month 2025</h1><p>Full year monthly breakdown by representative</p><div class="gold-rule"></div></div>
  {page_mom2025}
</div>

</main>
<script>
function showPage(id,btn){{
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('nav button').forEach(b=>b.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  btn.classList.add('active');
}}
</script>
</body>
</html>"""


# ── File watcher ───────────────────────────────────────────────

def git_push(output_path: Path):
    import subprocess
    repo_dir = str(output_path.parent)
    try:
        subprocess.run(["git", "add", output_path.name], cwd=repo_dir, check=True, capture_output=True)
        subprocess.run(
            ["git", "commit", "-m", f"update: {datetime.now().strftime('%Y-%m-%d %H:%M')}"],
            cwd=repo_dir, check=True, capture_output=True
        )
        subprocess.run(["git", "push"], cwd=repo_dir, check=True, capture_output=True)
        print(f"  ↑  Pushed to GitHub ✓")
    except subprocess.CalledProcessError as e:
        stderr = e.stderr.decode(errors="replace").strip()
        if "nothing to commit" in stderr or "nothing added" in stderr:
            return  # no-op, file unchanged
        print(f"  ✗  Git error: {stderr or e}")
    except FileNotFoundError:
        print("  ✗  git not found — install Git from git-scm.com and re-run")


def regenerate(xlsx_path: Path, output_path: Path):
    try:
        print(f"  ↻  Detected change — regenerating {output_path.name} ...", end=" ", flush=True)
        html = build_html(xlsx_path)
        output_path.write_text(html, encoding="utf-8")
        print(f"done ✓  ({datetime.now().strftime('%H:%M:%S')})")
        git_push(output_path)
    except Exception as e:
        import traceback
        print(f"\n  ✗  Error: {e}")
        traceback.print_exc()


def main():
    try:
        import pandas
        from watchdog.observers import Observer
        from watchdog.events import FileSystemEventHandler
    except ImportError:
        print("Missing dependencies. Run:\n\n    pip install pandas openpyxl watchdog\n")
        sys.exit(1)

    from watchdog.observers import Observer
    from watchdog.events import FileSystemEventHandler

    if not XLSX_PATH.exists():
        print(f"✗  Cannot find workbook at: {XLSX_PATH}")
        print("   Update XLSX_PATH in this script to point to your file.")
        sys.exit(1)

    print(f"\n Preserve Gold — Sales Board Auto-Updater")
    print(f" ─────────────────────────────────────────")
    print(f" Watching : {XLSX_PATH}")
    print(f" Output   : {OUTPUT_PATH}")
    print(f" Press Ctrl+C to stop\n")
    regenerate(XLSX_PATH, OUTPUT_PATH)

    class XlsxHandler(FileSystemEventHandler):
        def __init__(self): self._last = 0
        def on_modified(self, event):
            if Path(event.src_path).resolve() != XLSX_PATH.resolve(): return
            now = time.time()
            if now - self._last < 2: return
            self._last = now
            regenerate(XLSX_PATH, OUTPUT_PATH)
        on_created = on_modified

    observer = Observer()
    observer.schedule(XlsxHandler(), str(XLSX_PATH.parent), recursive=False)
    observer.start()
    try:
        while True: time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        print("\n Watcher stopped.")
    observer.join()


if __name__ == "__main__":
    main()