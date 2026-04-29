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
    from datetime import datetime, timedelta, timezone as tz
    il_tz       = tz(timedelta(hours=3))
    israel_time = datetime.now(il_tz).strftime("%d/%m/%Y %H:%M")

    # ── כרטיסי מדדים ──────────────────────────────────────────────────────────
    def make_card(m):
        acc  = "up-acc"      if m["color"] == "up" else "dn-acc"
        glow = "idx-glow-up" if m["color"] == "up" else "idx-glow-dn"
        chg  = "green"       if m["color"] == "up" else "red"
        return (
            '<div class="idx">'
            f'<div class="idx-acc {acc}"></div>'
            f'<div class="{glow}"></div>'
            f'<div class="idx-lbl">{m["name"]}</div>'
            f'<div class="idx-val">{m["price"]}</div>'
            f'<div class="idx-chg {chg}">{m["change"]}</div>'
            '</div>'
        )
    m_cards = "".join(make_card(m) for m in market_summary)

    # ── שורות טבלה ────────────────────────────────────────────────────────────
    def score_cls(sc):
        return {4: "s4", 3: "s3", 2: "s2"}.get(sc, "s1")

    def trend_cls(t):
        if "\u2191" in t: return "tr-up"
        if "\u2193" in t: return "tr-dn"
        return "tr-flat"

    def make_row(s):
        tkr     = s["Ticker"]
        price   = f"{s['Price']:,.2f}"
        chg_v   = s["Day_Chg_%"]
        chg_cls = "up-pct" if chg_v > 0 else "dn-pct"
        chg_str = f"{chg_v:+.2f}%"
        sc_cls  = score_cls(s["SCORE"])
        bk      = f"{s['Breakout']:,.2f}"
        sl      = f"{s['Stop_Loss']:,.2f}"
        tr_cls  = trend_cls(s["TREND"])
        return (
            f'<tr onclick="openChart(\'{tkr}\')">'
            f'<td><span class="tkr">{tkr}</span></td>'
            f'<td class="px">${price}</td>'
            f'<td class="{chg_cls}">{chg_str}</td>'
            f'<td><span class="{sc_cls}">{s["SCORE"]}</span></td>'
            f'<td class="rank-val">{s["Power_Rank"]}</td>'
            f'<td class="adx-val">{s["ADX"]}</td>'
            f'<td class="rsi-val">{s["RSI"]}</td>'
            f'<td class="bk-val">${bk}</td>'
            f'<td class="sl-val">${sl}</td>'
            f'<td class="{tr_cls}">{s["TREND"]}</td>'
            '</tr>'
        )
    rows = "".join(make_row(s) for s in results)

    # ── HTML מלא ──────────────────────────────────────────────────────────────
    css = """
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
html,body{background:#07090d;color:#c8d4e0;font-family:'IBM Plex Mono','SF Mono',monospace;font-size:12px;min-height:100vh}
.topbar{background:#080b10;border-bottom:1px solid #141d2a;height:58px;display:flex;align-items:center;justify-content:space-between;padding:0 28px;position:sticky;top:0;z-index:100;box-shadow:0 1px 30px rgba(0,0,0,0.6)}
.brand{display:flex;align-items:center;gap:12px}
.brand-name{font-size:15px;font-weight:700;letter-spacing:0.16em;color:#dde3ec;text-transform:uppercase}
.brand-sub{font-size:15px;font-weight:700;letter-spacing:0.16em;color:#e8b84b;text-transform:uppercase}
.live-dot{width:8px;height:8px;border-radius:50%;background:#34d058;box-shadow:0 0 0 3px rgba(52,208,88,0.15);animation:pulse 2.4s ease-in-out infinite;flex-shrink:0}
@keyframes pulse{0%,100%{opacity:1;box-shadow:0 0 0 3px rgba(52,208,88,0.15)}50%{opacity:0.4;box-shadow:0 0 0 6px rgba(52,208,88,0.05)}}
.ts-pill{font-size:11px;color:#3d5068;border:1px solid #141d2a;padding:5px 14px;border-radius:6px;letter-spacing:0.06em;background:rgba(255,255,255,0.015)}
.indices{display:grid;grid-template-columns:repeat(4,1fr);border-bottom:1px solid #141d2a;background:#141d2a;gap:1px}
.idx{background:#080b10;padding:22px 26px;position:relative;overflow:hidden;transition:background 0.18s}
.idx:hover{background:#0b0f18}
.idx-acc{position:absolute;top:0;left:0;bottom:0;width:3px}
.up-acc{background:linear-gradient(to bottom,#34d058 0%,rgba(52,208,88,0.06) 100%)}
.dn-acc{background:linear-gradient(to bottom,#ff4d4d 0%,rgba(255,77,77,0.06) 100%)}
.idx-glow-up{position:absolute;top:0;left:0;width:200px;height:200px;background:radial-gradient(circle at 5% 0%,rgba(52,208,88,0.09) 0%,transparent 65%);pointer-events:none}
.idx-glow-dn{position:absolute;top:0;left:0;width:200px;height:200px;background:radial-gradient(circle at 5% 0%,rgba(255,77,77,0.09) 0%,transparent 65%);pointer-events:none}
.idx-lbl{font-size:9px;letter-spacing:0.24em;color:#2e4258;text-transform:uppercase;font-weight:700;margin-bottom:10px;position:relative}
.idx-val{font-size:28px;font-weight:700;color:#eaf0f8;letter-spacing:-0.025em;font-variant-numeric:tabular-nums;position:relative;line-height:1}
.idx-chg{font-size:13px;font-weight:600;margin-top:8px;position:relative;letter-spacing:0.02em}
.green{color:#34d058}.red{color:#ff4d4d}
.body{padding:20px 26px 48px}
.scanner-bar{display:flex;align-items:center;justify-content:space-between;margin-bottom:14px}
.scanner-lbl{display:flex;align-items:center;gap:9px;font-size:9px;letter-spacing:0.24em;color:#e8b84b;text-transform:uppercase;font-weight:700}
.scanner-lbl::before{content:'';width:3px;height:14px;background:linear-gradient(to bottom,#e8b84b,#7a5e18);border-radius:2px;flex-shrink:0}
.count-badge{font-size:9px;color:#2e4258;border:1px solid #141d2a;padding:2px 10px;border-radius:20px;letter-spacing:0.07em}
.dataTables_wrapper{color:#c8d4e0}
.dataTables_wrapper .dataTables_filter,.dataTables_wrapper .dataTables_length{margin-bottom:12px}
.dataTables_wrapper .dataTables_filter label,.dataTables_wrapper .dataTables_length label{font-size:10px;color:#3d5068;display:flex;align-items:center;gap:8px;letter-spacing:0.05em}
.dataTables_wrapper .dataTables_filter input,.dataTables_wrapper .dataTables_length select{background:#0d1420;color:#c8d4e0;border:1px solid #141d2a;border-radius:5px;padding:5px 10px;font-family:inherit;font-size:11px;outline:none;transition:border-color 0.15s}
.dataTables_wrapper .dataTables_filter input:focus{border-color:#4a3a10}
.dataTables_wrapper .dataTables_info{font-size:10px;color:#2e4258;padding-top:10px}
.dataTables_wrapper .dataTables_paginate{padding-top:10px}
.dataTables_wrapper .paginate_button{font-family:inherit!important;font-size:10px!important;color:#3d5068!important;border:1px solid #141d2a!important;border-radius:4px!important;padding:3px 8px!important;margin:0 2px!important;background:transparent!important}
.dataTables_wrapper .paginate_button:hover{background:#0d1420!important;color:#e8b84b!important;border-color:#4a3a10!important}
.dataTables_wrapper .paginate_button.current,.dataTables_wrapper .paginate_button.current:hover{background:#0d1420!important;color:#e8b84b!important;border-color:#e8b84b!important}
.tbl-wrap{border:1px solid #141d2a;border-radius:10px;overflow:hidden;background:#080b10}
#stockTable{width:100%!important;border-collapse:collapse;margin:0!important}
#stockTable thead th{background:#060810!important;color:#1e3048!important;font-size:8px!important;font-weight:700!important;letter-spacing:0.2em!important;text-transform:uppercase!important;border:none!important;border-bottom:1px solid #141d2a!important;padding:10px 12px!important;white-space:nowrap!important;cursor:pointer!important}
#stockTable thead th:hover{color:#e8b84b!important}
#stockTable tbody tr{border-bottom:1px solid rgba(20,29,42,0.8)!important;cursor:pointer;transition:background 0.1s}
#stockTable tbody tr:hover{background:rgba(232,184,75,0.035)!important}
#stockTable tbody td{border:none!important;padding:10px 12px!important;vertical-align:middle!important;text-align:right!important;font-variant-numeric:tabular-nums!important;font-size:11px!important}
#stockTable tbody td:first-child{text-align:left!important}
.tkr{display:inline-flex;align-items:center;justify-content:center;background:#e8b84b;color:#0a0700!important;font-size:10px;font-weight:800;letter-spacing:0.08em;padding:4px 10px;border-radius:5px;min-width:60px;pointer-events:none;user-select:none}
tr:hover .tkr,tr:active .tkr{background:#e8b84b!important;color:#0a0700!important}
.px{color:#d4dde8!important;font-weight:600}
.rank-val{color:#e8b84b!important;font-weight:700;font-size:12px!important}
.adx-val{color:#5b9bd5!important}
.rsi-val{color:#6a7a8a!important}
.bk-val{color:#4aacbf!important}
.sl-val{color:#cc5555!important}
.up-pct{color:#34d058!important;font-weight:600}
.dn-pct{color:#ff4d4d!important;font-weight:600}
.tr-up{color:#34d058!important;font-weight:600}
.tr-dn{color:#ff4d4d!important;font-weight:600}
.tr-flat{color:#1a2a3a!important;font-size:13px!important}
.s4{width:26px;height:26px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;background:rgba(52,208,88,0.12);color:#34d058;border:1px solid rgba(52,208,88,0.28)}
.s3{width:26px;height:26px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;background:rgba(232,184,75,0.12);color:#e8b84b;border:1px solid rgba(232,184,75,0.28)}
.s2{width:26px;height:26px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;background:rgba(20,29,42,0.9);color:#3d5068;border:1px solid #141d2a}
.s1{width:26px;height:26px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;background:rgba(255,77,77,0.08);color:#ff4d4d;border:1px solid rgba(255,77,77,0.2)}
.modal-content{background:#080b10;border:1px solid #141d2a;border-radius:12px;overflow:hidden}
.modal-header{background:#060810;border-bottom:1px solid #141d2a;padding:13px 20px;display:flex;align-items:center;justify-content:space-between}
.modal-title{font-size:11px;font-weight:700;letter-spacing:0.14em;text-transform:uppercase;color:#e8b84b}
.modal-close{background:none;border:1px solid #141d2a;color:#3d5068;font-size:13px;cursor:pointer;border-radius:4px;padding:2px 8px;font-family:inherit;transition:all 0.15s}
.modal-close:hover{border-color:#3d5068;color:#c8d4e0}
.modal-body{padding:0;height:580px}
::-webkit-scrollbar{width:4px;height:4px}
::-webkit-scrollbar-track{background:#07090d}
::-webkit-scrollbar-thumb{background:#141d2a;border-radius:2px}
::-webkit-scrollbar-thumb:hover{background:#2e4258}
"""

    js = """
$(document).ready(function(){
  $('#stockTable').DataTable({
    order:[[4,'desc']],pageLength:50,lengthMenu:[25,50,100],
    language:{search:'',searchPlaceholder:'\\u05d7\\u05e4\\u05e9 \\u05d8\\u05d9\\u05e7\\u05e8...'},
    columnDefs:[{targets:'_all',className:''}]
  });
});
var chartModal=new bootstrap.Modal(document.getElementById('chartModal'));
function openChart(ticker){
  document.getElementById('chartModalLabel').textContent=ticker+' \u2014 Live Analysis';
  document.getElementById('tv_chart_container').innerHTML='';
  chartModal.show();
  setTimeout(function(){
    new TradingView.widget({
      autosize:true,symbol:ticker,interval:'D',timezone:'Asia/Jerusalem',
      theme:'dark',style:'1',locale:'en',toolbar_bg:'#080b10',
      hide_side_toolbar:false,allow_symbol_change:true,
      container_id:'tv_chart_container'
    });
  },180);
}
document.getElementById('chartModal').addEventListener('hidden.bs.modal',function(){
  document.getElementById('tv_chart_container').innerHTML='';
});
"""

    html = f"""<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>SHTIVI | COMMAND CENTER</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600;700&display=swap" rel="stylesheet">
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
<link href="https://cdn.datatables.net/1.13.6/css/dataTables.bootstrap5.min.css" rel="stylesheet">
<style>{css}</style>
</head>
<body>
<div class="topbar">
  <div class="brand">
    <div class="live-dot"></div>
    <span class="brand-name">Assaf Shtivi</span>
    <span class="brand-sub">Command Center</span>
  </div>
  <div class="ts-pill">IDT {israel_time}</div>
</div>
<div class="indices">{m_cards}</div>
<div class="body">
  <div class="scanner-bar">
    <div class="scanner-lbl">Watchlist Scanner</div>
    <div class="count-badge">{len(results)} stocks</div>
  </div>
  <div class="tbl-wrap">
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
        <button class="modal-close" data-bs-dismiss="modal">&#x2715;</button>
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
<script>{js}</script>
</body></html>"""

    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("\u2705 index.html \u05e0\u05d5\u05e6\u05e8")



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
