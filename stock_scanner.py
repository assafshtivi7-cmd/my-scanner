"""
SHTIVI COMMAND CENTER — stock_scanner.py
שדרוג מלא: UI וול-סטריט, שעון ישראל, תיקוני באגים, Threading
"""

import yfinance as yf
import pandas as pd
import numpy as np
import json, os, smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from datetime import datetime, timedelta, timezone
from concurrent.futures import ThreadPoolExecutor, as_completed

# ─── הגדרות ──────────────────────────────────────────────────────────────────
MY_EMAIL   = "assafshtivi7@gmail.com"
DB_FILE    = "last_run.json"
MAX_WORKERS = 10

WATCHLIST = [
    'BMNR','MSTR','HOOD','SEDG','SOFI','BBAI','SMR','NOW','BTG','HL',
    'ASTS','FISV','ON','QBTS','NVDA','PLTR','BULL','WDC','MU','FUBO',
    'AMPY','CRML','ZS','CTKB','META','RGTI','AMZN','OKLO','NNE','TSLA',
    'GOOG','NFLX','AAPL','AMD','MNDY','ORCL','ALLY','MP','MSFT','VRT',
    'CRM','IREN','UUUU','OPEN','FIG','INTC','LLY','AVGO','SHOP','RIVN',
    'CVNA','AFRM','EQIX','TCMD','DDOG','CRWD','ASML','APP','IONQ','BX',
    'NVTS','CMG','CAT','CNC','NKE','NEM','MRNA','CLSK','BEN','OSS',
    'HUN','SNDK','WULF','RDDT','ONDS','PANW','INTU','CRCL'
]

MARKET_INDICES = {
    "S&P 500":    "^GSPC",
    "NASDAQ":     "^IXIC",
    "BITCOIN":    "BTC-USD",
    "VIX (FEAR)": "^VIX",
}


# ─── לוגיקה טכנית ────────────────────────────────────────────────────────────

def calc_rsi(close, period=14):
    try:
        delta = close.diff()
        gain  = delta.clip(lower=0).ewm(alpha=1/period, adjust=False).mean()
        loss  = (-delta.clip(upper=0)).ewm(alpha=1/period, adjust=False).mean()
        rs    = gain / loss.replace(0, np.nan)
        return round(float((100 - (100 / (1 + rs))).iloc[-1]), 1)
    except:
        return 50

def calc_adx(df, period=14):
    try:
        h, l, c = df['High'], df['Low'], df['Close'].squeeze()
        tr  = pd.concat([h-l, (h-c.shift()).abs(), (l-c.shift()).abs()], axis=1).max(axis=1)
        pdm = (h.diff()).clip(lower=0)
        mdm = (-l.diff()).clip(lower=0)
        atr = tr.ewm(alpha=1/period, adjust=False).mean()
        pdi = 100 * pdm.ewm(alpha=1/period, adjust=False).mean() / atr.replace(0, np.nan)
        mdi = 100 * mdm.ewm(alpha=1/period, adjust=False).mean() / atr.replace(0, np.nan)
        dx  = 100 * (pdi - mdi).abs() / (pdi + mdi).replace(0, np.nan)
        return round(float(dx.ewm(alpha=1/period, adjust=False).mean().iloc[-1]), 1)
    except:
        return 20

def analyze_ticker(ticker, spy_ret, prev_scores):
    try:
        df = yf.Ticker(ticker).history(period="1y")
        if df.empty or len(df) < 22:
            return None
        close   = df['Close'].squeeze()
        curr_p  = float(close.iloc[-1])
        ema9    = float(close.ewm(span=9, adjust=False).mean().iloc[-1])
        ema21   = float(close.ewm(span=21, adjust=False).mean().iloc[-1])
        sma200  = float(close.rolling(200).mean().iloc[-1])
        rvol    = float(df['Volume'].iloc[-1] / df['Volume'].rolling(20).mean().iloc[-1])
        rsi_val = calc_rsi(close)
        adx_val = calc_adx(df)

        score = int(sum([
            ema9 > ema21,
            rvol > 1.1,
            curr_p > sma200,
            40 < rsi_val < 75,
        ]))

        rs_vs_spy   = round(((curr_p - float(close.iloc[-22])) / float(close.iloc[-22]) - spy_ret) * 100, 2)
        overext_pct = round(((curr_p / ema9) - 1) * 100, 1)

        # Power Rank עם קנס גרדואלי על מתיחת יתר
        rank = (score * 25) + (adx_val / 2) + min(rs_vs_spy / 4, 12)
        if overext_pct > 35: rank -= 50
        elif overext_pct > 20: rank -= 30
        elif overext_pct > 10: rank -= 15

        atr_val   = float((df['High'] - df['Low']).rolling(14).mean().iloc[-1])
        stop_loss = round(curr_p - (2 * atr_val), 2)
        day_chg   = round(((curr_p - float(close.iloc[-2])) / float(close.iloc[-2])) * 100, 2)
        prev      = prev_scores.get(ticker)
        trend     = (f"↑ ({prev})" if prev and score > prev else
                     f"↓ ({prev})" if prev and score < prev else "-")

        return {
            'Ticker':     ticker,
            'Price':      round(curr_p, 2),
            'SCORE':      score,
            'Power_Rank': round(rank, 1),
            'ADX':        adx_val,
            'RSI':        rsi_val,
            'RVOL':       round(rvol, 2),
            'RS_vs_SPY':  rs_vs_spy,
            'Overext_%':  overext_pct,
            'Day_Chg_%':  day_chg,
            'Breakout':   round(float(df['High'].rolling(20).max().iloc[-1]), 2),
            'Stop_Loss':  stop_loss,
            'TREND':      trend,
        }
    except:
        return None


# ─── שליפת מדדים (עם fallback ל-7 ימים) ─────────────────────────────────────

def fetch_market_summary():
    summary = []
    for name, sym in MARKET_INDICES.items():
        try:
            h = yf.Ticker(sym).history(period="7d")
            if h.empty or len(h) < 2:
                continue
            p = float(h['Close'].iloc[-1])
            c = ((h['Close'].iloc[-1] - h['Close'].iloc[-2]) / h['Close'].iloc[-2]) * 100
            fmt = f"{p:,.0f}" if sym in ("^GSPC", "^IXIC", "BTC-USD") else f"{p:.2f}"
            summary.append({
                "name":   name,
                "price":  fmt,
                "change": f"{c:+.2f}%",
                "color":  "up" if c >= 0 else "down",
            })
        except:
            continue
    return summary


# ─── HTML ─────────────────────────────────────────────────────────────────────

def generate_html(results, market_summary):
    # שעון ישראל UTC+3
    il_tz       = timezone(timedelta(hours=3))
    israel_time = datetime.now(il_tz).strftime("%d/%m/%Y %H:%M")

    # כרטיסי מדדים
    m_cards = "".join([f'''
        <div class="idx-card">
            <div class="idx-name">{m["name"]}</div>
            <div class="idx-price">{m["price"]}</div>
            <div class="idx-chg {'chg-up' if m['color']=='up' else 'chg-dn'}">{m["change"]}</div>
        </div>''' for m in market_summary])

    # שורות טבלה
    def score_class(s):
        return {4: "s4", 3: "s3", 2: "s2", 1: "s1", 0: "s0"}.get(s, "s0")

    rows = "".join([f'''
        <tr onclick="openChart('{s['Ticker']}')">
            <td><span class="ticker-badge">{s['Ticker']}</span></td>
            <td class="num">${s['Price']:,.2f}</td>
            <td class="num {'chg-up' if s['Day_Chg_%'] > 0 else 'chg-dn'}">{s['Day_Chg_%']:+.2f}%</td>
            <td><span class="score-pill {score_class(s['SCORE'])}">{s['SCORE']}</span></td>
            <td class="num gold">{s['Power_Rank']}</td>
            <td class="num cyan">{s['ADX']}</td>
            <td class="num">{s['RSI']}</td>
            <td class="num cyan-lt">${s['Breakout']:,.2f}</td>
            <td class="num warn">${s['Stop_Loss']:,.2f}</td>
            <td class="trend-cell {'trend-up' if '↑' in s['TREND'] else 'trend-dn' if '↓' in s['TREND'] else ''}">{s['TREND']}</td>
        </tr>''' for s in results])

    html = f'''<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>SHTIVI | COMMAND CENTER</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600;700&family=Chakra+Petch:wght@400;600;700&display=swap" rel="stylesheet">
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
<link href="https://cdn.datatables.net/1.13.6/css/dataTables.bootstrap5.min.css" rel="stylesheet">
<style>
/* ── CSS Variables ── */
:root {{
    --bg:      #060a0f;
    --surface: #0d1117;
    --panel:   #111820;
    --border:  #1e2d3d;
    --accent:  #fdbb2d;
    --accent2: #00d4ff;
    --text:    #c9d1d9;
    --muted:   #586069;
    --up:      #3fb950;
    --dn:      #f85149;
    --font-mono: 'IBM Plex Mono', monospace;
    --font-ui:   'Chakra Petch', sans-serif;
}}

/* ── Reset & Base ── */
*, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
html {{ scroll-behavior: smooth; }}
body {{
    background: var(--bg);
    color: var(--text);
    font-family: var(--font-mono);
    font-size: 13px;
    min-height: 100vh;
    background-image:
        radial-gradient(ellipse at 20% 0%, rgba(253,187,45,0.04) 0%, transparent 60%),
        radial-gradient(ellipse at 80% 100%, rgba(0,212,255,0.04) 0%, transparent 60%);
}}

/* ── Navbar ── */
.navbar {{
    background: var(--surface);
    border-bottom: 1px solid var(--border);
    padding: 0 32px;
    height: 56px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    position: sticky;
    top: 0;
    z-index: 1000;
    box-shadow: 0 1px 24px rgba(0,0,0,0.6);
}}
.brand {{
    font-family: var(--font-ui);
    font-weight: 700;
    font-size: 16px;
    letter-spacing: 0.08em;
    color: #fff;
    display: flex;
    align-items: center;
    gap: 10px;
}}
.brand-accent {{ color: var(--accent); }}
.brand-dot {{
    width: 8px; height: 8px;
    border-radius: 50%;
    background: var(--up);
    box-shadow: 0 0 8px var(--up);
    animation: pulse 2s infinite;
}}
@keyframes pulse {{
    0%, 100% {{ opacity: 1; transform: scale(1); }}
    50%       {{ opacity: 0.5; transform: scale(0.8); }}
}}
.timestamp {{
    font-family: var(--font-mono);
    font-size: 11px;
    color: var(--muted);
    border: 1px solid var(--border);
    padding: 4px 12px;
    border-radius: 4px;
    background: var(--bg);
    letter-spacing: 0.05em;
}}

/* ── Index Cards ── */
.indices-bar {{
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 12px;
    padding: 24px 32px 16px;
}}
.idx-card {{
    background: var(--panel);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 16px 20px;
    position: relative;
    overflow: hidden;
    transition: border-color 0.2s;
}}
.idx-card::before {{
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 2px;
    background: linear-gradient(90deg, var(--accent), transparent);
}}
.idx-card:hover {{ border-color: var(--accent); }}
.idx-name  {{ font-size: 10px; letter-spacing: 0.12em; color: var(--muted); text-transform: uppercase; margin-bottom: 6px; }}
.idx-price {{ font-family: var(--font-ui); font-size: 20px; font-weight: 700; color: #fff; }}
.idx-chg   {{ font-size: 12px; font-weight: 600; margin-top: 4px; }}
.chg-up    {{ color: var(--up); }}
.chg-dn    {{ color: var(--dn); }}

/* ── Main Container ── */
.main-wrap {{ padding: 0 32px 48px; }}

/* ── Panel Header ── */
.panel-header {{
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 16px;
    padding-bottom: 12px;
    border-bottom: 1px solid var(--border);
}}
.panel-title {{
    font-family: var(--font-ui);
    font-size: 13px;
    font-weight: 600;
    color: var(--accent);
    letter-spacing: 0.1em;
    text-transform: uppercase;
}}
.stock-count {{
    font-size: 11px;
    color: var(--muted);
    border: 1px solid var(--border);
    padding: 3px 10px;
    border-radius: 20px;
}}

/* ── Table Wrapper ── */
.table-wrap {{
    background: var(--panel);
    border: 1px solid var(--border);
    border-radius: 12px;
    overflow: hidden;
}}
#stockTable {{
    width: 100% !important;
    border-collapse: collapse;
}}
#stockTable thead th {{
    background: var(--surface) !important;
    color: var(--muted) !important;
    font-family: var(--font-ui);
    font-size: 11px;
    font-weight: 600;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    border: none !important;
    border-bottom: 1px solid var(--border) !important;
    padding: 12px 14px !important;
    white-space: nowrap;
}}
#stockTable tbody tr {{
    border-bottom: 1px solid rgba(30,45,61,0.5);
    cursor: pointer;
    transition: background 0.12s;
}}
#stockTable tbody tr:hover {{
    background: rgba(253,187,45,0.04) !important;
}}
#stockTable tbody td {{
    border: none !important;
    padding: 10px 14px !important;
    vertical-align: middle;
}}
.num {{
    font-family: var(--font-mono);
    font-variant-numeric: tabular-nums;
    text-align: right;
}}
.gold     {{ color: var(--accent) !important; font-weight: 600; }}
.cyan     {{ color: var(--accent2) !important; }}
.cyan-lt  {{ color: #7ee8fa !important; }}
.warn     {{ color: #f9826c !important; }}

/* ── Ticker Badge — ניגודיות קבועה שחור-על-זהב ── */
.ticker-badge {{
    display: inline-block;
    background-color: #fdbb2d !important;
    color: #000000 !important;
    font-family: var(--font-ui);
    font-weight: 700;
    font-size: 11px;
    letter-spacing: 0.06em;
    padding: 4px 10px;
    border-radius: 5px;
    border: none;
    min-width: 72px;
    text-align: center;
    /* עוצר כל שינוי בהובר/פוקוס */
    pointer-events: none;
    user-select: none;
    -webkit-font-smoothing: auto;
}}
/* מונע ירושת צבע מ-hover/focus של TR */
tr:hover .ticker-badge,
tr:focus .ticker-badge,
tr:active .ticker-badge {{
    background-color: #fdbb2d !important;
    color: #000000 !important;
}}

/* ── Score Pills ── */
.score-pill {{
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 28px; height: 28px;
    border-radius: 50%;
    font-family: var(--font-ui);
    font-weight: 700;
    font-size: 13px;
}}
.s4 {{ background: var(--up);    color: #000; box-shadow: 0 0 10px rgba(63,185,80,0.4); }}
.s3 {{ background: var(--accent); color: #000; }}
.s2 {{ background: #2d333b;       color: var(--text); }}
.s1, .s0 {{ background: rgba(248,81,73,0.15); color: var(--dn); border: 1px solid rgba(248,81,73,0.3); }}

/* ── Trend ── */
.trend-up {{ color: var(--up)  !important; font-weight: 600; }}
.trend-dn {{ color: var(--dn)  !important; font-weight: 600; }}

/* ── DataTables overrides ── */
.dataTables_wrapper .dataTables_filter input {{
    background: var(--surface);
    color: var(--text);
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 5px 10px;
    margin-right: 8px;
    font-family: var(--font-mono);
    font-size: 12px;
    outline: none;
}}
.dataTables_wrapper .dataTables_filter input:focus {{
    border-color: var(--accent);
    box-shadow: 0 0 0 2px rgba(253,187,45,0.15);
}}
.dataTables_wrapper .dataTables_length select {{
    background: var(--surface);
    color: var(--text);
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 4px 8px;
    font-family: var(--font-mono);
}}
.dataTables_wrapper .dataTables_info,
.dataTables_wrapper .dataTables_paginate {{ color: var(--muted); font-size: 11px; }}
.dataTables_wrapper .paginate_button.current {{
    background: var(--accent) !important;
    color: #000 !important;
    border: none !important;
    border-radius: 4px !important;
}}
.dataTables_wrapper .paginate_button:hover {{
    background: rgba(253,187,45,0.15) !important;
    color: var(--accent) !important;
    border: none !important;
}}
.sorting, .sorting_asc, .sorting_desc {{
    background-color: transparent !important;
    color: inherit !important;
}}

/* ── Modal ── */
.modal-content {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 14px;
    overflow: hidden;
}}
.modal-header {{
    background: var(--panel);
    border-bottom: 1px solid var(--border);
    padding: 14px 20px;
}}
.modal-title {{
    font-family: var(--font-ui);
    color: var(--accent);
    font-size: 14px;
    letter-spacing: 0.1em;
}}
.btn-close-gold {{
    background: none;
    border: none;
    color: var(--muted);
    font-size: 18px;
    cursor: pointer;
    line-height: 1;
    padding: 0 4px;
    transition: color 0.2s;
}}
.btn-close-gold:hover {{ color: #fff; }}
.modal-body {{ padding: 0; height: 600px; }}

/* ── Scrollbar ── */
::-webkit-scrollbar {{ width: 6px; }}
::-webkit-scrollbar-track {{ background: var(--bg); }}
::-webkit-scrollbar-thumb {{ background: var(--border); border-radius: 3px; }}
::-webkit-scrollbar-thumb:hover {{ background: var(--muted); }}
</style>
</head>
<body>

<!-- Navbar -->
<nav class="navbar">
    <div class="brand">
        <div class="brand-dot"></div>
        ASSAF SHTIVI &nbsp;<span class="brand-accent">COMMAND CENTER</span>
    </div>
    <div class="timestamp">🕐 IDT {israel_time}</div>
</nav>

<!-- Market Indices -->
<div class="indices-bar">{m_cards}</div>

<!-- Main Table -->
<div class="main-wrap">
    <div class="panel-header">
        <div class="panel-title">⚡ WATCHLIST SCANNER</div>
        <div class="stock-count">{len(results)} STOCKS</div>
    </div>
    <div class="table-wrap">
        <table id="stockTable" class="table table-hover text-center align-middle mb-0">
            <thead>
                <tr>
                    <th>Ticker</th>
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

<!-- TradingView Modal -->
<div class="modal fade" id="chartModal" tabindex="-1" aria-hidden="true">
    <div class="modal-dialog modal-xl modal-dialog-centered">
        <div class="modal-content">
            <div class="modal-header">
                <span class="modal-title" id="chartModalLabel">LIVE CHART</span>
                <button type="button" class="btn-close-gold" data-bs-dismiss="modal">✕</button>
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
        language: {{ search: '🔍  ', searchPlaceholder: 'חפש טיקר...' }},
        columnDefs: [{{ targets: '_all', className: '' }}]
    }});
}});

var tvWidget = null;
var chartModal = new bootstrap.Modal(document.getElementById('chartModal'));

function openChart(ticker) {{
    document.getElementById('chartModalLabel').textContent = '⚡ ' + ticker + ' — LIVE ANALYSIS';
    document.getElementById('tv_chart_container').innerHTML = '';
    chartModal.show();
    setTimeout(function() {{
        tvWidget = new TradingView.widget({{
            autosize:     true,
            symbol:       ticker,
            interval:     'D',
            timezone:     'Asia/Jerusalem',
            theme:        'dark',
            style:        '1',
            locale:       'en',
            toolbar_bg:   '#0d1117',
            hide_side_toolbar: false,
            allow_symbol_change: true,
            container_id: 'tv_chart_container',
        }});
    }}, 200);
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


# ─── Excel ────────────────────────────────────────────────────────────────────

def create_styled_excel(df, file_name):
    cols = ['Ticker','Price','SCORE','Power_Rank','ADX','RSI','RVOL',
            'RS_vs_SPY','Overext_%','Day_Chg_%','Breakout','Stop_Loss','TREND']
    df = df[[c for c in cols if c in df.columns]]

    writer    = pd.ExcelWriter(file_name, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Scanner')
    wb, ws    = writer.book, writer.sheets['Scanner']
    last_row  = len(df)

    hdr_fmt   = wb.add_format({'bold': True, 'bg_color': '#FFFF00', 'border': 1,
                                'align': 'center', 'font_name': 'Calibri'})
    green_fmt = wb.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100',
                                'align': 'center', 'border': 1})
    red_fmt   = wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006',
                                'align': 'center', 'border': 1})
    orange_fmt= wb.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C5700',
                                'align': 'center', 'border': 1})
    base_fmt  = wb.add_format({'align': 'center', 'border': 1, 'font_name': 'Calibri'})

    for i, col in enumerate(df.columns):
        ws.write(0, i, col, hdr_fmt)
        ws.set_column(i, i, 12, base_fmt)

    ws.freeze_panes(1, 0)
    ws.set_row(0, 18)

    col_idx = {c: i for i, c in enumerate(df.columns)}

    # SCORE: ירוק 4+, אדום 0-1
    if 'SCORE' in col_idx:
        c = col_idx['SCORE']
        ws.conditional_format(1, c, last_row, c, {'type': 'cell', 'criteria': '>=', 'value': 4, 'format': green_fmt})
        ws.conditional_format(1, c, last_row, c, {'type': 'cell', 'criteria': '<=', 'value': 1, 'format': red_fmt})

    # Power_Rank: color scale
    if 'Power_Rank' in col_idx:
        c = col_idx['Power_Rank']
        ws.conditional_format(1, c, last_row, c, {
            'type': '3_color_scale', 'min_color': '#FFC7CE', 'mid_color': '#FFEB9C', 'max_color': '#C6EFCE'})

    # RS_vs_SPY
    if 'RS_vs_SPY' in col_idx:
        c = col_idx['RS_vs_SPY']
        ws.conditional_format(1, c, last_row, c, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': green_fmt})
        ws.conditional_format(1, c, last_row, c, {'type': 'cell', 'criteria': '<', 'value': 0, 'format': red_fmt})

    # Overext_%: כתום 15+, אדום 30+
    if 'Overext_%' in col_idx:
        c = col_idx['Overext_%']
        ws.conditional_format(1, c, last_row, c, {'type': 'cell', 'criteria': '>=', 'value': 30, 'format': red_fmt})
        ws.conditional_format(1, c, last_row, c, {'type': 'cell', 'criteria': '>=', 'value': 15, 'format': orange_fmt})

    # Day_Chg_%
    if 'Day_Chg_%' in col_idx:
        c = col_idx['Day_Chg_%']
        ws.conditional_format(1, c, last_row, c, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': green_fmt})
        ws.conditional_format(1, c, last_row, c, {'type': 'cell', 'criteria': '<', 'value': 0, 'format': red_fmt})

    # TREND: ↑ ירוק, ↓ אדום
    if 'TREND' in col_idx:
        c = col_idx['TREND']
        ws.conditional_format(1, c, last_row, c, {'type': 'text', 'criteria': 'containing', 'value': '↑', 'format': green_fmt})
        ws.conditional_format(1, c, last_row, c, {'type': 'text', 'criteria': 'containing', 'value': '↓', 'format': red_fmt})

    writer.close()
    print(f"✅ Excel נוצר: {file_name}")


# ─── שליחת מייל ──────────────────────────────────────────────────────────────

def send_email(file_name):
    try:
        pwd = os.getenv("APP_PASSWORD", "").replace(" ", "")
        if not pwd:
            print("⚠️  APP_PASSWORD לא הוגדר — מדלג על שליחת מייל")
            return
        msg            = MIMEMultipart()
        msg['Subject'] = f"🚀 COMMAND CENTER — {datetime.now(timezone(timedelta(hours=3))).strftime('%d/%m/%Y %H:%M')} IDT"
        msg['From']    = MY_EMAIL
        msg['To']      = MY_EMAIL
        msg.attach(MIMEText("הדוח היומי מצורף. האתר עודכן בהתאם.", "plain", "utf-8"))
        with open(file_name, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={file_name}")
        msg.attach(part)
        with smtplib.SMTP_SSL("smtp.gmail.com", 465,
                               context=ssl.create_default_context()) as server:
            server.login(MY_EMAIL, pwd)
            server.send_message(msg)
        print("✅ מייל נשלח בהצלחה!")
    except Exception as e:
        print(f"❌ שגיאה בשליחת מייל: {e}")


# ─── Main ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("📡 שולף SPY...")
    try:
        spy_df  = yf.download('SPY', period='1mo', progress=False)
        spy_ret = float((spy_df['Close'].iloc[-1] - spy_df['Close'].iloc[0]) / spy_df['Close'].iloc[0])
    except:
        spy_ret = 0.0
        print("⚠️  SPY נכשל, ממשיך עם 0")

    prev_scores = json.load(open(DB_FILE)) if os.path.exists(DB_FILE) else {}

    print(f"🔍 סורק {len(WATCHLIST)} מניות...")
    results = []
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(analyze_ticker, t, spy_ret, prev_scores): t for t in WATCHLIST}
        for future in as_completed(futures):
            res = future.result()
            if res:
                results.append(res)

    if not results:
        print("❌ אין תוצאות.")
        raise SystemExit(1)

    results.sort(key=lambda x: (x['SCORE'], x['Power_Rank']), reverse=True)
    print(f"✅ {len(results)} מניות עברו סינון")

    print("📡 שולף מדדי שוק...")
    market_summary = fetch_market_summary()

    il_date   = datetime.now(timezone(timedelta(hours=3))).strftime('%Y-%m-%d')
    file_name = f"Master_Scanner_{il_date}.xlsx"

    create_styled_excel(pd.DataFrame(results), file_name)
    generate_html(results, market_summary)

    json.dump({r['Ticker']: int(r['SCORE']) for r in results}, open(DB_FILE, "w"))
    print("✅ last_run.json עודכן")

    send_email(file_name)
