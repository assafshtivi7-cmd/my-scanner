"""
פונקציית generate_html מעודכנת — עיצוב Bloomberg Terminal
מחליפה את הפונקציה הקיימת ב-stock_scanner.py
"""

def generate_html(results, market_summary):
    from datetime import datetime, timedelta, timezone

    il_tz       = timezone(timedelta(hours=3))
    israel_time = datetime.now(il_tz).strftime("%d/%m/%Y %H:%M")

    # ── כרטיסי מדדים ──────────────────────────────────────────────────────────
    m_cards = "".join([f'''<div class="idx">
      <div class="idx-bar {'up-bar' if m['color']=='up' else 'dn-bar'}"></div>
      <div class="idx-label">{m['name']}</div>
      <div class="idx-val">{m['price']}</div>
      <div class="idx-chg {'up' if m['color']=='up' else 'dn'}">{m['change']}</div>
    </div>''' for m in market_summary])

    # ── שורות טבלה ────────────────────────────────────────────────────────────
    def score_cls(s):
        return {4:"s4", 3:"s3", 2:"s2"}.get(s, "s1")

    def trend_cls(t):
        if "↑" in t: return "trend-up"
        if "↓" in t: return "trend-dn"
        return "trend-flat"

    rows = "".join([f'''<tr onclick="openChart('{s['Ticker']}')">
      <td><span class="ticker">{s['Ticker']}</span></td>
      <td class="price">${s['Price']:,.2f}</td>
      <td class="{'chg-up' if s['Day_Chg_%']>0 else 'chg-dn'}">{s['Day_Chg_%']:+.2f}%</td>
      <td><span class="score-circle {score_cls(s['SCORE'])}">{s['SCORE']}</span></td>
      <td class="rank">{s['Power_Rank']}</td>
      <td class="adx">{s['ADX']}</td>
      <td class="rsi-col">{s['RSI']}</td>
      <td class="breakout">${s['Breakout']:,.2f}</td>
      <td class="stop">${s['Stop_Loss']:,.2f}</td>
      <td class="{trend_cls(s['TREND'])}">{s['TREND']}</td>
    </tr>''' for s in results])

    html = f'''<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>SHTIVI | COMMAND CENTER</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600;700&display=swap" rel="stylesheet">
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
<link href="https://cdn.datatables.net/1.13.6/css/dataTables.bootstrap5.min.css" rel="stylesheet">
<style>
/* ── Reset ── */
*, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}

/* ── Base ── */
:root {{
  --bg:       #080c10;
  --surface:  #0a0e14;
  --panel:    #0d1420;
  --border:   #1a2332;
  --gold:     #c9aa71;
  --gold-dim: #7a6540;
  --text:     #cbd5e1;
  --muted:    #3a4a5c;
  --up:       #3fb950;
  --dn:       #f85149;
  --mono:     'IBM Plex Mono', 'SF Mono', 'Fira Code', monospace;
}}
html, body {{
  background: var(--bg);
  color: var(--text);
  font-family: var(--mono);
  font-size: 12px;
  min-height: 100vh;
}}

/* ── Topbar ── */
.topbar {{
  background: var(--surface);
  border-bottom: 1px solid var(--border);
  padding: 0 28px;
  height: 48px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  position: sticky;
  top: 0;
  z-index: 100;
}}
.brand {{
  display: flex;
  align-items: center;
  gap: 12px;
  font-size: 11px;
  font-weight: 700;
  letter-spacing: 0.14em;
  text-transform: uppercase;
  color: #e2e8f0;
}}
.brand-accent {{ color: var(--gold); }}
.live-dot {{
  width: 6px;
  height: 6px;
  border-radius: 50%;
  background: var(--up);
  box-shadow: 0 0 6px var(--up);
  animation: blink 2.4s ease-in-out infinite;
  flex-shrink: 0;
}}
@keyframes blink {{
  0%, 100% {{ opacity: 1; }}
  50%       {{ opacity: 0.3; }}
}}
.timestamp {{
  font-size: 10px;
  color: var(--muted);
  border: 1px solid var(--border);
  padding: 3px 11px;
  border-radius: 4px;
  letter-spacing: 0.07em;
}}

/* ── Indices Bar ── */
.indices {{
  display: grid;
  grid-template-columns: repeat(4, 1fr);
  gap: 0;
  border-bottom: 1px solid var(--border);
  background: var(--border);
}}
.idx {{
  background: var(--surface);
  padding: 16px 22px;
  position: relative;
}}
.idx-bar {{
  position: absolute;
  top: 0; left: 0; bottom: 0;
  width: 2px;
}}
.up-bar {{ background: var(--up); opacity: 0.7; }}
.dn-bar {{ background: var(--dn); opacity: 0.7; }}
.idx-label {{
  font-size: 9px;
  letter-spacing: 0.18em;
  color: var(--muted);
  text-transform: uppercase;
  font-weight: 600;
  margin-bottom: 5px;
}}
.idx-val {{
  font-size: 20px;
  font-weight: 700;
  color: #e2e8f0;
  letter-spacing: -0.02em;
  font-variant-numeric: tabular-nums;
  line-height: 1.1;
}}
.idx-chg {{
  font-size: 11px;
  font-weight: 600;
  margin-top: 3px;
  letter-spacing: 0.03em;
}}
.up {{ color: var(--up); }}
.dn {{ color: var(--dn); }}

/* ── Content ── */
.content {{ padding: 20px 28px 48px; }}

/* ── Table header row ── */
.tbl-header {{
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 14px;
}}
.section-label {{
  font-size: 9px;
  letter-spacing: 0.2em;
  color: var(--gold);
  text-transform: uppercase;
  font-weight: 700;
  display: flex;
  align-items: center;
  gap: 9px;
}}
.section-label::before {{
  content: '';
  display: block;
  width: 3px;
  height: 12px;
  background: var(--gold);
  border-radius: 1px;
  flex-shrink: 0;
}}
.count-pill {{
  font-size: 9px;
  color: var(--muted);
  border: 1px solid var(--border);
  padding: 2px 9px;
  border-radius: 20px;
  letter-spacing: 0.07em;
}}

/* ── DataTables Search/Length overrides ── */
.dataTables_wrapper {{ color: var(--text); }}
.dataTables_wrapper .dataTables_filter,
.dataTables_wrapper .dataTables_length {{
  margin-bottom: 12px;
}}
.dataTables_wrapper .dataTables_filter label,
.dataTables_wrapper .dataTables_length label {{
  font-size: 10px;
  color: var(--muted);
  letter-spacing: 0.06em;
  display: flex;
  align-items: center;
  gap: 8px;
}}
.dataTables_wrapper .dataTables_filter input,
.dataTables_wrapper .dataTables_length select {{
  background: var(--panel);
  color: var(--text);
  border: 1px solid var(--border);
  border-radius: 5px;
  padding: 5px 10px;
  font-family: var(--mono);
  font-size: 11px;
  outline: none;
  transition: border-color 0.15s;
}}
.dataTables_wrapper .dataTables_filter input:focus {{
  border-color: var(--gold-dim);
}}
.dataTables_wrapper .dataTables_info {{
  font-size: 10px;
  color: var(--muted);
  letter-spacing: 0.05em;
  padding-top: 10px;
}}
.dataTables_wrapper .dataTables_paginate {{
  padding-top: 10px;
}}
.dataTables_wrapper .paginate_button {{
  font-family: var(--mono) !important;
  font-size: 10px !important;
  color: var(--muted) !important;
  border: 1px solid var(--border) !important;
  border-radius: 4px !important;
  padding: 3px 8px !important;
  margin: 0 2px !important;
  background: transparent !important;
}}
.dataTables_wrapper .paginate_button:hover {{
  background: var(--panel) !important;
  color: var(--gold) !important;
  border-color: var(--gold-dim) !important;
}}
.dataTables_wrapper .paginate_button.current,
.dataTables_wrapper .paginate_button.current:hover {{
  background: var(--panel) !important;
  color: var(--gold) !important;
  border-color: var(--gold) !important;
}}

/* ── Table ── */
.table-wrap {{
  background: var(--panel);
  border: 1px solid var(--border);
  border-radius: 10px;
  overflow: hidden;
}}
#stockTable {{
  width: 100% !important;
  border-collapse: collapse;
  margin: 0 !important;
}}
#stockTable thead th {{
  background: var(--surface) !important;
  color: var(--muted) !important;
  font-family: var(--mono);
  font-size: 9px;
  font-weight: 600;
  letter-spacing: 0.14em;
  text-transform: uppercase;
  border: none !important;
  border-bottom: 1px solid var(--border) !important;
  padding: 11px 12px !important;
  white-space: nowrap;
  cursor: pointer;
  user-select: none;
}}
#stockTable thead th:hover {{ color: var(--gold) !important; }}
#stockTable.dataTable thead th.sorting_asc::after,
#stockTable.dataTable thead th.sorting_desc::after {{
  color: var(--gold) !important;
}}
#stockTable tbody tr {{
  border-bottom: 1px solid rgba(26,35,50,0.6);
  cursor: pointer;
  transition: background 0.1s;
}}
#stockTable tbody tr:hover {{
  background: rgba(201,170,113,0.04) !important;
}}
#stockTable tbody td {{
  border: none !important;
  padding: 9px 12px !important;
  vertical-align: middle;
  text-align: right;
  color: #6b7a8d;
  font-variant-numeric: tabular-nums;
  font-size: 12px;
}}
#stockTable tbody td:first-child {{ text-align: left; }}

/* ── Cell types ── */
.ticker {{
  display: inline-flex;
  align-items: center;
  justify-content: center;
  background: var(--gold);
  color: #000 !important;
  font-family: var(--mono);
  font-size: 10px;
  font-weight: 800;
  letter-spacing: 0.07em;
  padding: 3px 9px;
  border-radius: 4px;
  min-width: 60px;
  pointer-events: none;
  user-select: none;
}}
tr:hover .ticker,
tr:focus .ticker,
tr:active .ticker {{
  background: var(--gold) !important;
  color: #000 !important;
}}
.price  {{ color: #c8d4e0 !important; font-weight: 600; }}
.rank   {{ color: var(--gold) !important; font-weight: 700; }}
.adx    {{ color: #7aa6c8 !important; }}
.rsi-col {{ color: #8892a4 !important; }}
.breakout {{ color: #7ab8c8 !important; }}
.stop   {{ color: #c87171 !important; }}
.chg-up {{ color: var(--up) !important; font-weight: 600; }}
.chg-dn {{ color: var(--dn) !important; font-weight: 600; }}
.trend-up   {{ color: var(--up) !important; font-weight: 600; font-size: 11px; }}
.trend-dn   {{ color: var(--dn) !important; font-weight: 600; font-size: 11px; }}
.trend-flat {{ color: var(--border) !important; font-size: 13px; }}

/* ── Score Circles ── */
.score-circle {{
  width: 26px;
  height: 26px;
  border-radius: 50%;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  font-size: 11px;
  font-weight: 700;
  font-family: var(--mono);
}}
.s4 {{ background: rgba(63,185,80,0.1);  color: var(--up); border: 1px solid rgba(63,185,80,0.3); }}
.s3 {{ background: rgba(201,170,113,0.1); color: var(--gold); border: 1px solid rgba(201,170,113,0.25); }}
.s2 {{ background: rgba(26,35,50,0.8);    color: #4a6080;   border: 1px solid var(--border); }}
.s1 {{ background: rgba(248,81,73,0.08);  color: var(--dn); border: 1px solid rgba(248,81,73,0.2); }}

/* ── Modal ── */
.modal-content {{
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 12px;
  overflow: hidden;
}}
.modal-header {{
  background: var(--panel);
  border-bottom: 1px solid var(--border);
  padding: 13px 20px;
  display: flex;
  align-items: center;
  justify-content: space-between;
}}
.modal-title {{
  font-family: var(--mono);
  font-size: 11px;
  font-weight: 700;
  letter-spacing: 0.14em;
  text-transform: uppercase;
  color: var(--gold);
}}
.modal-close {{
  background: none;
  border: 1px solid var(--border);
  color: var(--muted);
  font-size: 13px;
  cursor: pointer;
  border-radius: 4px;
  padding: 2px 8px;
  font-family: var(--mono);
  transition: all 0.15s;
}}
.modal-close:hover {{ border-color: var(--muted); color: var(--text); }}
.modal-body {{ padding: 0; height: 580px; }}

/* ── Scrollbar ── */
::-webkit-scrollbar {{ width: 5px; height: 5px; }}
::-webkit-scrollbar-track {{ background: var(--bg); }}
::-webkit-scrollbar-thumb {{ background: var(--border); border-radius: 3px; }}
::-webkit-scrollbar-thumb:hover {{ background: var(--muted); }}
</style>
</head>
<body>

<!-- ── Topbar ── -->
<div class="topbar">
  <div class="brand">
    <div class="live-dot"></div>
    Assaf Shtivi &nbsp;<span class="brand-accent">Command Center</span>
  </div>
  <div class="timestamp">IDT {israel_time}</div>
</div>

<!-- ── Indices ── -->
<div class="indices">{m_cards}</div>

<!-- ── Main Table ── -->
<div class="content">
  <div class="tbl-header">
    <div class="section-label">Watchlist Scanner</div>
    <div class="count-pill">{len(results)} stocks</div>
  </div>
  <div class="table-wrap">
    <table id="stockTable" class="table table-hover text-center align-middle mb-0">
      <thead>
        <tr>
          <th style="text-align:left">Ticker</th>
          <th>Price</th>
          <th>Day %</th>
          <th>Score</th>
          <th>Rank</th>
          <th>ADX</th>
          <th>RSI</th>
          <th>Breakout</th>
          <th>Stop Loss</th>
          <th>Trend</th>
        </tr>
      </thead>
      <tbody>{rows}</tbody>
    </table>
  </div>
</div>

<!-- ── TradingView Modal ── -->
<div class="modal fade" id="chartModal" tabindex="-1" aria-hidden="true">
  <div class="modal-dialog modal-xl modal-dialog-centered">
    <div class="modal-content">
      <div class="modal-header">
        <span class="modal-title" id="chartModalLabel">Live Chart</span>
        <button class="modal-close" data-bs-dismiss="modal">✕</button>
      </div>
      <div class="modal-body">
        <div id="tv_chart_container" style="height:100%;"></div>
      </div>
    </div>
  </div>
</div>

<script src="https://code.jquery.com/jquery-3.7.0.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>
<script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
<script src="https://cdn.datatables.net/1.13.6/js/dataTables.bootstrap5.min.js"></script>
<script src="https://s3.tradingview.com/tv.js"></script>
<script>
$(document).ready(function() {{
  $('#stockTable').DataTable({{
    order: [[4, 'desc']],
    pageLength: 50,
    lengthMenu: [25, 50, 100],
    language: {{ search: '', searchPlaceholder: 'חפש טיקר...' }},
    columnDefs: [
      {{ targets: [3], orderable: true }},
      {{ targets: '_all', className: '' }}
    ]
  }});
}});

var chartModal = new bootstrap.Modal(document.getElementById('chartModal'));

function openChart(ticker) {{
  document.getElementById('chartModalLabel').textContent = ticker + ' — Live Analysis';
  document.getElementById('tv_chart_container').innerHTML = '';
  chartModal.show();
  setTimeout(function() {{
    new TradingView.widget({{
      autosize:          true,
      symbol:            ticker,
      interval:          'D',
      timezone:          'Asia/Jerusalem',
      theme:             'dark',
      style:             '1',
      locale:            'en',
      toolbar_bg:        '#0a0e14',
      hide_side_toolbar: false,
      allow_symbol_change: true,
      container_id:      'tv_chart_container',
    }});
  }}, 180);
}}

document.getElementById('chartModal').addEventListener('hidden.bs.modal', function() {{
  document.getElementById('tv_chart_container').innerHTML = '';
}});
</script>
</body>
</html>'''

    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("✅ index.html נוצר")
