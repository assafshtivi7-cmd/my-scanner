"""
SHTIVI COMMAND CENTER — stock_scanner.py
תוקן: כרטיסי מדדים גדולים יותר + Excel conditional formatting מתוקן
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

# ─── הגדרות ───────────────────────────────────────────────────────────────────
MY_EMAIL    = "assafshtivi7@gmail.com"
DB_FILE     = "last_run.json"
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


# ─── לוגיקה טכנית ─────────────────────────────────────────────────────────────

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
        ema9    = float(close.ewm(span=9,  adjust=False).mean().iloc[-1])
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

        rank = (score * 25) + (adx_val / 2) + min(rs_vs_spy / 4, 12)
        if   overext_pct > 35: rank -= 50
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


# ─── מדדי שוק ─────────────────────────────────────────────────────────────────

def fetch_market_summary():
    summary = []
    for name, sym in MARKET_INDICES.items():
        try:
            h = yf.Ticker(sym).history(period="7d")
            if h.empty or len(h) < 2:
                continue
            p   = float(h['Close'].iloc[-1])
            c   = ((h['Close'].iloc[-1] - h['Close'].iloc[-2]) / h['Close'].iloc[-2]) * 100
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


# ─── HTML ──────────────────────────────────────────────────────────────────────

def generate_html(results, market_summary):
    il_tz       = timezone(timedelta(hours=3))
    israel_time = datetime.now(il_tz).strftime("%d/%m/%Y %H:%M")

    m_cards = "".join([f'''<div class="idx {'up-card' if m['color']=='up' else 'dn-card'}">
      <div class="idx-bar {'up-bar' if m['color']=='up' else 'dn-bar'}"></div>
      <div class="idx-label">{m['name']}</div>
      <div class="idx-val">{m['price']}</div>
      <div class="idx-chg {'up' if m['color']=='up' else 'dn'}">{m['change']}</div>
    </div>''' for m in market_summary])

    def score_cls(s):
        return {4: "s4", 3: "s3", 2: "s2"}.get(s, "s1")

    def trend_cls(t):
        if "↑" in t: return "trend-up"
        if "↓" in t: return "trend-dn"
        return "trend-flat"

    rows = "".join([f'''<tr onclick="openChart('{s['Ticker']}')">
      <td><span class="ticker">{s['Ticker']}</span></td>
      <td class="price">${s['Price']:,.2f}</td>
      <td class="{'chg-up' if s['Day_Chg_%'] > 0 else 'chg-dn'}">{s['Day_Chg_%']:+.2f}%</td>
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
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
:root{{
  --bg:#080c10;--surface:#0a0e14;--panel:#0d1420;--border:#1a2332;
  --gold:#c9aa71;--gold-dim:#7a6540;--text:#cbd5e1;--muted:#3a4a5c;
  --up:#3fb950;--dn:#f85149;
  --mono:'IBM Plex Mono','SF Mono','Fira Code',monospace;
}}
html,body{{background:var(--bg);color:var(--text);font-family:var(--mono);font-size:12px;min-height:100vh}}

/* ── Topbar — גבוה יותר, טקסט גדול יותר ── */
.topbar{{background:var(--surface);border-bottom:1px solid var(--border);padding:0 32px;height:62px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100;box-shadow:0 2px 24px rgba(0,0,0,0.5)}}
.brand{{display:flex;align-items:center;gap:14px;font-size:16px;font-weight:700;letter-spacing:0.12em;text-transform:uppercase;color:#e2e8f0}}
.brand-accent{{color:var(--gold)}}
.live-dot{{width:9px;height:9px;border-radius:50%;background:var(--up);box-shadow:0 0 10px var(--up),0 0 20px rgba(63,185,80,0.3);animation:blink 2.4s ease-in-out infinite;flex-shrink:0}}
@keyframes blink{{0%,100%{{opacity:1}}50%{{opacity:0.3}}}}
.timestamp{{font-size:12px;color:var(--muted);border:1px solid var(--border);padding:5px 15px;border-radius:5px;letter-spacing:0.07em;background:rgba(255,255,255,0.02)}}

/* ── Index Cards — גדולים ויפים יותר ── */
.indices{{display:grid;grid-template-columns:repeat(4,1fr);gap:1px;border-bottom:1px solid var(--border);background:var(--border)}}
.idx{{
  background:var(--surface);
  padding:24px 30px;
  position:relative;
  overflow:hidden;
  transition:background 0.2s;
}}
.idx:hover{{background:#0e1622}}
/* פס צבע שמאלי ← gradient */
.idx-bar{{position:absolute;top:0;left:0;bottom:0;width:3px}}
.up-bar{{background:linear-gradient(180deg,var(--up) 0%,rgba(63,185,80,0.15) 100%)}}
.dn-bar{{background:linear-gradient(180deg,var(--dn) 0%,rgba(248,81,73,0.15) 100%)}}
/* glow פינה עליונה-שמאלית */
.idx.up-card::after{{content:'';position:absolute;top:-20px;left:-20px;width:140px;height:140px;background:radial-gradient(circle,rgba(63,185,80,0.08) 0%,transparent 70%);pointer-events:none}}
.idx.dn-card::after{{content:'';position:absolute;top:-20px;left:-20px;width:140px;height:140px;background:radial-gradient(circle,rgba(248,81,73,0.08) 0%,transparent 70%);pointer-events:none}}
.idx-label{{font-size:10px;letter-spacing:0.22em;color:var(--muted);text-transform:uppercase;font-weight:700;margin-bottom:10px}}
.idx-val{{font-size:28px;font-weight:700;color:#eaf0f6;letter-spacing:-0.02em;font-variant-numeric:tabular-nums;line-height:1.1}}
.idx-chg{{font-size:14px;font-weight:600;margin-top:6px;letter-spacing:0.02em}}
.up{{color:var(--up)}}.dn{{color:var(--dn)}}

/* ── Content ── */
.content{{padding:20px 28px 48px}}
.tbl-header{{display:flex;align-items:center;justify-content:space-between;margin-bottom:14px}}
.section-label{{font-size:9px;letter-spacing:0.2em;color:var(--gold);text-transform:uppercase;font-weight:700;display:flex;align-items:center;gap:9px}}
.section-label::before{{content:'';display:block;width:3px;height:12px;background:var(--gold);border-radius:1px;flex-shrink:0}}
.count-pill{{font-size:9px;color:var(--muted);border:1px solid var(--border);padding:2px 9px;border-radius:20px;letter-spacing:0.07em}}

/* ── DataTables ── */
.dataTables_wrapper{{color:var(--text)}}
.dataTables_wrapper .dataTables_filter,.dataTables_wrapper .dataTables_length{{margin-bottom:12px}}
.dataTables_wrapper .dataTables_filter label,.dataTables_wrapper .dataTables_length label{{font-size:10px;color:var(--muted);letter-spacing:0.06em;display:flex;align-items:center;gap:8px}}
.dataTables_wrapper .dataTables_filter input,.dataTables_wrapper .dataTables_length select{{background:var(--panel);color:var(--text);border:1px solid var(--border);border-radius:5px;padding:5px 10px;font-family:var(--mono);font-size:11px;outline:none}}
.dataTables_wrapper .dataTables_filter input:focus{{border-color:var(--gold-dim)}}
.dataTables_wrapper .dataTables_info{{font-size:10px;color:var(--muted);padding-top:10px}}
.dataTables_wrapper .dataTables_paginate{{padding-top:10px}}
.dataTables_wrapper .paginate_button{{font-family:var(--mono)!important;font-size:10px!important;color:var(--muted)!important;border:1px solid var(--border)!important;border-radius:4px!important;padding:3px 8px!important;margin:0 2px!important;background:transparent!important}}
.dataTables_wrapper .paginate_button:hover{{background:var(--panel)!important;color:var(--gold)!important;border-color:var(--gold-dim)!important}}
.dataTables_wrapper .paginate_button.current,.dataTables_wrapper .paginate_button.current:hover{{background:var(--panel)!important;color:var(--gold)!important;border-color:var(--gold)!important}}

/* ── Table ── */
.table-wrap{{background:var(--panel);border:1px solid var(--border);border-radius:10px;overflow:hidden}}
#stockTable{{width:100%!important;border-collapse:collapse;margin:0!important}}
#stockTable thead th{{background:var(--surface)!important;color:var(--muted)!important;font-family:var(--mono);font-size:9px;font-weight:600;letter-spacing:0.14em;text-transform:uppercase;border:none!important;border-bottom:1px solid var(--border)!important;padding:11px 12px!important;white-space:nowrap;cursor:pointer}}
#stockTable thead th:hover{{color:var(--gold)!important}}
#stockTable tbody tr{{border-bottom:1px solid rgba(26,35,50,0.6);cursor:pointer;transition:background 0.1s}}
#stockTable tbody tr:hover{{background:rgba(201,170,113,0.04)!important}}
#stockTable tbody td{{border:none!important;padding:9px 12px!important;vertical-align:middle;text-align:right;color:#6b7a8d;font-variant-numeric:tabular-nums;font-size:12px}}
#stockTable tbody td:first-child{{text-align:left}}

/* ── Cells ── */
.ticker{{display:inline-flex;align-items:center;justify-content:center;background:var(--gold);color:#000!important;font-family:var(--mono);font-size:10px;font-weight:800;letter-spacing:0.07em;padding:3px 9px;border-radius:4px;min-width:60px;pointer-events:none;user-select:none}}
tr:hover .ticker,tr:focus .ticker,tr:active .ticker{{background:var(--gold)!important;color:#000!important}}
.price{{color:#c8d4e0!important;font-weight:600}}.rank{{color:var(--gold)!important;font-weight:700}}
.adx{{color:#7aa6c8!important}}.rsi-col{{color:#8892a4!important}}
.breakout{{color:#7ab8c8!important}}.stop{{color:#c87171!important}}
.chg-up{{color:var(--up)!important;font-weight:600}}.chg-dn{{color:var(--dn)!important;font-weight:600}}
.trend-up{{color:var(--up)!important;font-weight:600;font-size:11px}}
.trend-dn{{color:var(--dn)!important;font-weight:600;font-size:11px}}
.trend-flat{{color:var(--border)!important;font-size:13px}}
.score-circle{{width:26px;height:26px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;font-family:var(--mono)}}
.s4{{background:rgba(63,185,80,0.1);color:var(--up);border:1px solid rgba(63,185,80,0.3)}}
.s3{{background:rgba(201,170,113,0.1);color:var(--gold);border:1px solid rgba(201,170,113,0.25)}}
.s2{{background:rgba(26,35,50,0.8);color:#4a6080;border:1px solid var(--border)}}
.s1{{background:rgba(248,81,73,0.08);color:var(--dn);border:1px solid rgba(248,81,73,0.2)}}

/* ── Modal ── */
.modal-content{{background:var(--surface);border:1px solid var(--border);border-radius:12px;overflow:hidden}}
.modal-header{{background:var(--panel);border-bottom:1px solid var(--border);padding:13px 20px;display:flex;align-items:center;justify-content:space-between}}
.modal-title{{font-family:var(--mono);font-size:11px;font-weight:700;letter-spacing:0.14em;text-transform:uppercase;color:var(--gold)}}
.modal-close{{background:none;border:1px solid var(--border);color:var(--muted);font-size:13px;cursor:pointer;border-radius:4px;padding:2px 8px;font-family:var(--mono);transition:all 0.15s}}
.modal-close:hover{{border-color:var(--muted);color:var(--text)}}
.modal-body{{padding:0;height:580px}}
::-webkit-scrollbar{{width:5px;height:5px}}
::-webkit-scrollbar-track{{background:var(--bg)}}
::-webkit-scrollbar-thumb{{background:var(--border);border-radius:3px}}
</style>
</head>
<body>

<div class="topbar">
  <div class="brand"><div class="live-dot"></div>Assaf Shtivi &nbsp;<span class="brand-accent">Command Center</span></div>
  <div class="timestamp">IDT {israel_time}</div>
</div>

<div class="indices">{m_cards}</div>

<div class="content">
  <div class="tbl-header">
    <div class="section-label">Watchlist Scanner</div>
    <div class="count-pill">{len(results)} stocks</div>
  </div>
  <div class="table-wrap">
    <table id="stockTable" class="table table-hover text-center align-middle mb-0">
      <thead><tr>
        <th style="text-align:left">Ticker</th>
        <th>Price</th><th>Day %</th><th>Score</th><th>Rank</th>
        <th>ADX</th><th>RSI</th><th>Breakout</th><th>Stop Loss</th><th>Trend</th>
      </tr></thead>
      <tbody>{rows}</tbody>
    </table>
  </div>
</div>

<div class="modal fade" id="chartModal" tabindex="-1" aria-hidden="true">
  <div class="modal-dialog modal-xl modal-dialog-centered">
    <div class="modal-content">
      <div class="modal-header">
        <span class="modal-title" id="chartModalLabel">Live Chart</span>
        <button class="modal-close" data-bs-dismiss="modal">✕</button>
      </div>
      <div class="modal-body"><div id="tv_chart_container" style="height:100%;"></div></div>
    </div>
  </div>
</div>

<script src="https://code.jquery.com/jquery-3.7.0.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>
<script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
<script src="https://cdn.datatables.net/1.13.6/js/dataTables.bootstrap5.min.js"></script>
<script src="https://s3.tradingview.com/tv.js"></script>
<script>
$(document).ready(function(){{
  $('#stockTable').DataTable({{
    order:[[4,'desc']],pageLength:50,lengthMenu:[25,50,100],
    language:{{search:'',searchPlaceholder:'חפש טיקר...'}},
    columnDefs:[{{targets:'_all',className:''}}]
  }});
}});
var chartModal=new bootstrap.Modal(document.getElementById('chartModal'));
function openChart(ticker){{
  document.getElementById('chartModalLabel').textContent=ticker+' — Live Analysis';
  document.getElementById('tv_chart_container').innerHTML='';
  chartModal.show();
  setTimeout(function(){{
    new TradingView.widget({{
      autosize:true,symbol:ticker,interval:'D',timezone:'Asia/Jerusalem',
      theme:'dark',style:'1',locale:'en',toolbar_bg:'#0a0e14',
      hide_side_toolbar:false,allow_symbol_change:true,
      container_id:'tv_chart_container'
    }});
  }},180);
}}
document.getElementById('chartModal').addEventListener('hidden.bs.modal',function(){{
  document.getElementById('tv_chart_container').innerHTML='';
}});
</script>
</body></html>'''

    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("✅ index.html נוצר")


# ─── Excel — תוקן: conditional formatting עובד נכון ────────────────────────────
#
#  הבאג הקודם: ws.set_column(i, i, 12, base_fmt) העביר fmt לכל עמודה,
#  מה שגרם ל-xlsxwriter לרשום cell format שדורס את ה-conditional_format.
#  הפתרון: set_column ללא format argument (רק רוחב),
#           ולכתוב base_fmt ידנית רק לשורת הכותרת.
#
def create_styled_excel(df, file_name):
    COLS = ['Ticker','Price','SCORE','Power_Rank','ADX','RSI','RVOL',
            'RS_vs_SPY','Overext_%','Day_Chg_%','Breakout','Stop_Loss','TREND']
    df = df[[c for c in COLS if c in df.columns]].copy()

    writer   = pd.ExcelWriter(file_name, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Scanner')
    wb, ws   = writer.book, writer.sheets['Scanner']
    last_row = len(df)   # שורה אחרונה של נתונים (0-based header → data rows 1..last_row)

    # ── פורמטים ───────────────────────────────────────────────────────────────
    hdr_fmt    = wb.add_format({
        'bold': True, 'bg_color': '#1a1a2e', 'font_color': '#c9aa71',
        'border': 1, 'border_color': '#2d3561',
        'align': 'center', 'valign': 'vcenter',
        'font_name': 'Calibri', 'font_size': 10
    })
    base_fmt   = wb.add_format({
        'align': 'center', 'valign': 'vcenter',
        'font_name': 'Calibri', 'font_size': 10,
        'border': 1, 'border_color': '#2a2a3e'
    })
    green_fmt  = wb.add_format({
        'bg_color': '#C6EFCE', 'font_color': '#006100',
        'align': 'center', 'valign': 'vcenter',
        'font_name': 'Calibri', 'font_size': 10,
        'border': 1, 'border_color': '#006100'
    })
    red_fmt    = wb.add_format({
        'bg_color': '#FFC7CE', 'font_color': '#9C0006',
        'align': 'center', 'valign': 'vcenter',
        'font_name': 'Calibri', 'font_size': 10,
        'border': 1, 'border_color': '#9C0006'
    })
    orange_fmt = wb.add_format({
        'bg_color': '#FFEB9C', 'font_color': '#9C5700',
        'align': 'center', 'valign': 'vcenter',
        'font_name': 'Calibri', 'font_size': 10,
        'border': 1, 'border_color': '#9C5700'
    })

    # ── כותרות ────────────────────────────────────────────────────────────────
    col_widths = {
        'Ticker': 10, 'Price': 12, 'SCORE': 8, 'Power_Rank': 12,
        'ADX': 8, 'RSI': 8, 'RVOL': 8, 'RS_vs_SPY': 12,
        'Overext_%': 11, 'Day_Chg_%': 11, 'Breakout': 12,
        'Stop_Loss': 12, 'TREND': 12
    }
    for i, col in enumerate(df.columns):
        ws.write(0, i, col, hdr_fmt)
        # ← set_column ללא format! רק רוחב. זה מה שמאפשר ל-conditional_format לעבוד.
        ws.set_column(i, i, col_widths.get(col, 11))

    ws.freeze_panes(1, 0)
    ws.set_row(0, 20)

    # ── כתיבת base_fmt לכל תאי הנתונים (בלי לדרוס conditional) ───────────────
    # xlsxwriter מיישם conditional_format *על גבי* ה-cell format,
    # אז נכתוב את base_fmt ישירות לכל שורה — זה מאפשר לשניהם לחיות יחד.
    for row_num in range(1, last_row + 1):
        for col_num in range(len(df.columns)):
            val = df.iloc[row_num - 1, col_num]
            ws.write(row_num, col_num, val, base_fmt)

    col_idx = {c: i for i, c in enumerate(df.columns)}

    # ── Conditional Formatting ─────────────────────────────────────────────────

    # SCORE: ירוק ≥ 4 | אדום ≤ 1
    if 'SCORE' in col_idx:
        c = col_idx['SCORE']
        ws.conditional_format(1, c, last_row, c,
            {'type': 'cell', 'criteria': '>=', 'value': 4, 'format': green_fmt})
        ws.conditional_format(1, c, last_row, c,
            {'type': 'cell', 'criteria': '<=', 'value': 1, 'format': red_fmt})

    # Power_Rank: color scale אדום→צהוב→ירוק
    if 'Power_Rank' in col_idx:
        c = col_idx['Power_Rank']
        ws.conditional_format(1, c, last_row, c, {
            'type': '3_color_scale',
            'min_color': '#FFC7CE', 'mid_color': '#FFEB9C', 'max_color': '#C6EFCE'
        })

    # ADX: ירוק ≥ 25 (מגמה חזקה) | אדום < 15 (שוק צדדי)
    if 'ADX' in col_idx:
        c = col_idx['ADX']
        ws.conditional_format(1, c, last_row, c,
            {'type': 'cell', 'criteria': '>=', 'value': 25, 'format': green_fmt})
        ws.conditional_format(1, c, last_row, c,
            {'type': 'cell', 'criteria': '<',  'value': 15, 'format': red_fmt})

    # RSI: ירוק 50–70 (מגמה בריאה) | אדום > 75 (קנייתר-יתר) | כתום < 35 (מכירת-יתר)
    if 'RSI' in col_idx:
        c = col_idx['RSI']
        ws.conditional_format(1, c, last_row, c,
            {'type': 'cell', 'criteria': '>=', 'value': 75, 'format': red_fmt})
        ws.conditional_format(1, c, last_row, c,
            {'type': 'cell', 'criteria': '<=', 'value': 35, 'format': orange_fmt})
        ws.conditional_format(1, c, last_row, c,
            {'type': 'cell', 'criteria': 'between',
             'minimum': 50, 'maximum': 70, 'format': green_fmt})

    # RVOL: ירוק ≥ 1.5 (ווליום חריג חיובי) | אדום < 0.7 (ווליום נמוך)
    if 'RVOL' in col_idx:
        c = col_idx['RVOL']
        ws.conditional_format(1, c, last_row, c,
            {'type': 'cell', 'criteria': '>=', 'value': 1.5, 'format': green_fmt})
        ws.conditional_format(1, c, last_row, c,
            {'type': 'cell', 'criteria': '<',  'value': 0.7, 'format': red_fmt})

    # RS_vs_SPY: ירוק > 0 | אדום < 0
    if 'RS_vs_SPY' in col_idx:
        c = col_idx['RS_vs_SPY']
        ws.conditional_format(1, c, last_row, c,
            {'type': 'cell', 'criteria': '>', 'value': 0, 'format': green_fmt})
        ws.conditional_format(1, c, last_row, c,
            {'type': 'cell', 'criteria': '<', 'value': 0, 'format': red_fmt})

    # Overext_%: אדום ≥ 30 | כתום ≥ 15
    if 'Overext_%' in col_idx:
        c = col_idx['Overext_%']
        ws.conditional_format(1, c, last_row, c,
            {'type': 'cell', 'criteria': '>=', 'value': 30, 'format': red_fmt})
        ws.conditional_format(1, c, last_row, c,
            {'type': 'cell', 'criteria': '>=', 'value': 15, 'format': orange_fmt})

    # Day_Chg_%: ירוק > 0 | אדום < 0
    if 'Day_Chg_%' in col_idx:
        c = col_idx['Day_Chg_%']
        ws.conditional_format(1, c, last_row, c,
            {'type': 'cell', 'criteria': '>', 'value': 0, 'format': green_fmt})
        ws.conditional_format(1, c, last_row, c,
            {'type': 'cell', 'criteria': '<', 'value': 0, 'format': red_fmt})

    # TREND: ↑ ירוק | ↓ אדום
    if 'TREND' in col_idx:
        c = col_idx['TREND']
        ws.conditional_format(1, c, last_row, c,
            {'type': 'text', 'criteria': 'containing', 'value': '↑', 'format': green_fmt})
        ws.conditional_format(1, c, last_row, c,
            {'type': 'text', 'criteria': 'containing', 'value': '↓', 'format': red_fmt})

    writer.close()
    print(f"✅ Excel נוצר: {file_name}")


# ─── שליחת מייל ───────────────────────────────────────────────────────────────

def send_email(file_name):
    print("📧 מתחיל שליחת מייל...")

    pwd = os.getenv("APP_PASSWORD", "").replace(" ", "")
    if not pwd:
        raise EnvironmentError(
            "❌ APP_PASSWORD לא הוגדר! "
            "ודא שה-Secret קיים: GitHub → Settings → Secrets → APP_PASSWORD"
        )
    print(f"✅ סיסמה נמצאה ({len(pwd)} תווים)")

    abs_path = os.path.abspath(file_name)
    print(f"📎 מחפש: {abs_path}")
    if not os.path.exists(abs_path):
        raise FileNotFoundError(
            f"❌ קובץ לא נמצא: {abs_path}\n"
            f"   תיקייה: {os.getcwd()}\n"
            f"   קבצים: {os.listdir('.')}"
        )
    print(f"✅ קובץ נמצא — {os.path.getsize(abs_path):,} bytes")

    il_time        = datetime.now(timezone(timedelta(hours=3))).strftime('%d/%m/%Y %H:%M')
    msg            = MIMEMultipart()
    msg['Subject'] = f"🚀 COMMAND CENTER — {il_time} IDT"
    msg['From']    = MY_EMAIL
    msg['To']      = MY_EMAIL
    msg.attach(MIMEText(
        f"שלום אסף,\n\nהסריקה היומית הושלמה.\nשעה: {il_time} (שעון ישראל)\n\n"
        f"הדוח מצורף כקובץ אקסל.\n"
        f"האתר: https://assafshtivi7-cmd.github.io/my-scanner/\n",
        "plain", "utf-8"
    ))

    with open(abs_path, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition",
                    f'attachment; filename="{os.path.basename(abs_path)}"')
    msg.attach(part)
    print("✅ אקסל צורף למייל")

    print("📡 מתחבר ל-Gmail...")
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ssl.create_default_context()) as server:
        server.login(MY_EMAIL, pwd)
        print("✅ התחברות הצליחה")
        server.send_message(msg)
    print("✅✅ מייל נשלח בהצלחה!")


# ─── Main ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("=" * 50)
    print("🚀 SHTIVI COMMAND CENTER — מתחיל")
    print(f"📁 תיקייה: {os.getcwd()}")
    print("=" * 50)

    print("\n📡 שולף SPY...")
    try:
        spy_df  = yf.download('SPY', period='1mo', progress=False)
        spy_ret = float((spy_df['Close'].iloc[-1] - spy_df['Close'].iloc[0]) / spy_df['Close'].iloc[0])
        print(f"✅ SPY return: {spy_ret:.3f}")
    except Exception as e:
        spy_ret = 0.0
        print(f"⚠️  SPY נכשל ({e}), ממשיך עם 0")

    prev_scores = json.load(open(DB_FILE)) if os.path.exists(DB_FILE) else {}
    print(f"📂 ציונים קודמים: {len(prev_scores)} מניות")

    print(f"\n🔍 סורק {len(WATCHLIST)} מניות ({MAX_WORKERS} במקביל)...")
    results = []
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(analyze_ticker, t, spy_ret, prev_scores): t for t in WATCHLIST}
        for future in as_completed(futures):
            res = future.result()
            if res:
                results.append(res)

    if not results:
        print("❌ אין תוצאות")
        raise SystemExit(1)

    results.sort(key=lambda x: (x['SCORE'], x['Power_Rank']), reverse=True)
    print(f"✅ {len(results)} מניות | מוביל: {results[0]['Ticker']} ({results[0]['Power_Rank']})")

    print("\n📡 שולף מדדים...")
    market_summary = fetch_market_summary()
    print(f"✅ {len(market_summary)} מדדים")

    il_date   = datetime.now(timezone(timedelta(hours=3))).strftime('%Y-%m-%d')
    file_name = f"Master_Scanner_{il_date}.xlsx"

    print(f"\n📊 יוצר {file_name}...")
    create_styled_excel(pd.DataFrame(results), file_name)

    print("\n🌐 יוצר index.html...")
    generate_html(results, market_summary)

    json.dump({r['Ticker']: int(r['SCORE']) for r in results}, open(DB_FILE, "w"))
    print("✅ last_run.json עודכן")

    print("\n" + "=" * 50)
    send_email(file_name)
    print("=" * 50)
    print("\n🏁 הריצה הושלמה בהצלחה!")
