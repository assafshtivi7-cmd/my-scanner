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

    def make_card(m):
        acc  = "up-acc" if m["color"] == "up" else "dn-acc"
        glow = "glow-up" if m["color"] == "up" else "glow-dn"
        chg  = "green"  if m["color"] == "up" else "red"
        return (
            f'<div class="idx-card">'
            f'<div class="idx-side {acc}"></div>'
            f'<div class="{glow}"></div>'
            f'<div class="idx-lbl">{m["name"]}</div>'
            f'<div class="idx-num">{m["price"]}</div>'
            f'<div class="idx-chg {chg}">{m["change"]}</div>'
            f'</div>'
        )
    m_cards = "".join(make_card(m) for m in market_summary)

    def score_cls(sc):
        return {4:"s4", 3:"s3", 2:"s2"}.get(sc, "s1")

    def trend_cls(t):
        if "\u2191" in t: return "tr-up"
        if "\u2193" in t: return "tr-dn"
        return "tr-flat"

    def make_row(s):
        chg_v   = s["Day_Chg_%"]
        chg_cls = "up-pct" if chg_v > 0 else "dn-pct"
        return (
            f'<tr onclick="openChart(\'{s["Ticker"]}\')">'
            f'<td class="td-ticker"><span class="tkr">{s["Ticker"]}</span></td>'
            f'<td class="td-r px">${s["Price"]:,.2f}</td>'
            f'<td class="td-r {chg_cls}">{chg_v:+.2f}%</td>'
            f'<td class="td-c"><span class="{score_cls(s["SCORE"])}">{s["SCORE"]}</span></td>'
            f'<td class="td-r rank">{s["Power_Rank"]}</td>'
            f'<td class="td-r adxc">{s["ADX"]}</td>'
            f'<td class="td-r rsic">{s["RSI"]}</td>'
            f'<td class="td-r bkc">${s["Breakout"]:,.2f}</td>'
            f'<td class="td-r slc">${s["Stop_Loss"]:,.2f}</td>'
            f'<td class="td-r {trend_cls(s["TREND"])}">{s["TREND"]}</td>'
            f'</tr>'
        )
    rows = "".join(make_row(s) for s in results)

    html = """<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>SHTIVI | COMMAND CENTER</title>
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600;700&display=swap" rel="stylesheet">
<link href="https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css" rel="stylesheet">
<style>
/* ── Reset ─────────────────────────────────────────────────── */
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
html{font-size:13px}
body{
  background:#07090d;
  color:#8fa3b8;
  font-family:'IBM Plex Mono',monospace;
  min-height:100vh;
}

/* ── Topbar ─────────────────────────────────────────────────── */
.topbar{
  background:#090d13;
  border-bottom:1px solid #0f1923;
  height:56px;
  display:flex;
  align-items:center;
  justify-content:space-between;
  padding:0 28px;
  position:sticky;
  top:0;
  z-index:200;
}
.brand{display:flex;align-items:center;gap:14px}
.brand-name{
  font-size:14px;font-weight:600;
  letter-spacing:0.18em;
  color:#cdd9e5;
  text-transform:uppercase;
}
.brand-cmd{
  font-size:14px;font-weight:700;
  letter-spacing:0.18em;
  color:#d4a843;
  text-transform:uppercase;
}
.live-dot{
  width:7px;height:7px;
  border-radius:50%;
  background:#34c759;
  box-shadow:0 0 0 2px rgba(52,199,89,0.18);
  animation:livepulse 2.6s ease-in-out infinite;
  flex-shrink:0;
}
@keyframes livepulse{
  0%,100%{opacity:1;box-shadow:0 0 0 2px rgba(52,199,89,0.18)}
  50%{opacity:0.35;box-shadow:0 0 0 5px rgba(52,199,89,0.05)}
}
.ts-pill{
  font-size:10px;color:#3a5068;
  border:1px solid #0f1923;
  padding:4px 14px;border-radius:5px;
  letter-spacing:0.07em;
  background:#06080c;
}

/* ── Index Cards ─────────────────────────────────────────────── */
.indices{
  display:grid;
  grid-template-columns:repeat(4,1fr);
  border-bottom:1px solid #0f1923;
  gap:1px;
  background:#0f1923;
}
.idx-card{
  background:#09090f;
  padding:20px 24px;
  position:relative;
  overflow:hidden;
  transition:background .15s;
  cursor:default;
}
.idx-card:hover{background:#0c1018}
.idx-side{
  position:absolute;top:0;left:0;bottom:0;width:3px;
}
.up-acc{background:linear-gradient(180deg,#34c759 0%,rgba(52,199,89,0.05) 100%)}
.dn-acc{background:linear-gradient(180deg,#ff3b30 0%,rgba(255,59,48,0.05) 100%)}
.glow-up{
  position:absolute;top:0;left:0;
  width:160px;height:160px;
  background:radial-gradient(circle at 0% 0%,rgba(52,199,89,0.1) 0%,transparent 70%);
  pointer-events:none;
}
.glow-dn{
  position:absolute;top:0;left:0;
  width:160px;height:160px;
  background:radial-gradient(circle at 0% 0%,rgba(255,59,48,0.1) 0%,transparent 70%);
  pointer-events:none;
}
.idx-lbl{
  font-size:9px;letter-spacing:0.26em;
  color:#2a4058;text-transform:uppercase;
  font-weight:700;margin-bottom:10px;
  position:relative;
}
.idx-num{
  font-size:28px;font-weight:700;
  color:#dce8f0;
  letter-spacing:-0.025em;
  font-variant-numeric:tabular-nums;
  line-height:1;position:relative;
}
.idx-chg{
  font-size:13px;font-weight:600;
  margin-top:8px;position:relative;
}
.green{color:#34c759}
.red{color:#ff3b30}

/* ── Body ───────────────────────────────────────────────────── */
.wrap{padding:22px 28px 60px}
.scan-bar{
  display:flex;align-items:center;
  justify-content:space-between;
  margin-bottom:18px;
}
.scan-lbl{
  display:flex;align-items:center;gap:10px;
  font-size:9px;letter-spacing:0.26em;
  color:#d4a843;text-transform:uppercase;font-weight:700;
}
.scan-lbl::before{
  content:'';width:3px;height:14px;
  background:linear-gradient(180deg,#d4a843,#6a5010);
  border-radius:2px;flex-shrink:0;
}
.cnt-badge{
  font-size:9px;color:#253545;
  border:1px solid #0f1923;
  padding:2px 10px;border-radius:20px;
  letter-spacing:0.07em;
}

/* ── DataTables search/length bar ───────────────────────────── */
.dataTables_wrapper{color:#8fa3b8}
.dataTables_length,.dataTables_filter{margin-bottom:14px}
.dataTables_length label,
.dataTables_filter label{
  display:flex;align-items:center;gap:8px;
  font-size:10px;color:#3a5068;letter-spacing:0.06em;
}
.dataTables_filter input,
.dataTables_length select{
  background:#0d1520;
  color:#a0b4c8;
  border:1px solid #0f1923;
  border-radius:5px;
  padding:5px 10px;
  font-family:'IBM Plex Mono',monospace;
  font-size:11px;
  outline:none;
}
.dataTables_filter input:focus{border-color:#6a5010}
.dataTables_info{font-size:10px;color:#253545;padding-top:10px}
.dataTables_paginate{padding-top:10px}
.paginate_button{
  font-family:'IBM Plex Mono',monospace !important;
  font-size:10px !important;
  color:#3a5068 !important;
  border:1px solid #0f1923 !important;
  border-radius:4px !important;
  padding:3px 8px !important;
  margin:0 2px !important;
  background:transparent !important;
  cursor:pointer;
}
.paginate_button:hover{
  background:#0d1520 !important;
  color:#d4a843 !important;
  border-color:#6a5010 !important;
}
.paginate_button.current,
.paginate_button.current:hover{
  background:#0d1520 !important;
  color:#d4a843 !important;
  border-color:#d4a843 !important;
}

/* ── Table ──────────────────────────────────────────────────── */
.tbl-wrap{
  border:1px solid #0f1923;
  border-radius:10px;
  overflow:hidden;
}
#stockTable{
  width:100% !important;
  border-collapse:collapse;
  background:#09090f;
}
#stockTable thead th{
  background:#060810 !important;
  color:#1e3248 !important;
  font-family:'IBM Plex Mono',monospace !important;
  font-size:8px !important;
  font-weight:700 !important;
  letter-spacing:0.22em !important;
  text-transform:uppercase !important;
  border:none !important;
  border-bottom:1px solid #0f1923 !important;
  padding:10px 14px !important;
  white-space:nowrap !important;
  cursor:pointer !important;
}
#stockTable thead th:hover{color:#d4a843 !important}
#stockTable.dataTable thead th.sorting_asc::after,
#stockTable.dataTable thead th.sorting_desc::after{
  color:#d4a843;
}
#stockTable tbody tr{
  background:#09090f !important;
  border-bottom:1px solid #0d1520 !important;
  cursor:pointer;
  transition:background .1s;
}
#stockTable tbody tr:hover{
  background:#0d1826 !important;
}
#stockTable tbody td{
  background:transparent !important;
  border:none !important;
  padding:11px 14px !important;
  font-variant-numeric:tabular-nums;
  color:#4a6278;
  font-size:12px;
}
.td-ticker{text-align:left !important}
.td-r{text-align:right !important}
.td-c{text-align:center !important}

/* ── Cell styles ─────────────────────────────────────────────── */
.tkr{
  display:inline-flex;align-items:center;justify-content:center;
  background:#d4a843;
  color:#1a0d00 !important;
  font-size:10px;font-weight:800;
  letter-spacing:0.08em;
  padding:4px 11px;
  border-radius:5px;
  min-width:60px;
  pointer-events:none;
  user-select:none;
}
tr:hover .tkr{background:#d4a843 !important;color:#1a0d00 !important}
.px{color:#c8d8e8 !important;font-weight:600}
.rank{color:#d4a843 !important;font-weight:700;font-size:13px !important}
.adxc{color:#4a90c8 !important}
.rsic{color:#5a7888 !important}
.bkc{color:#3aacbf !important}
.slc{color:#c05050 !important}
.up-pct{color:#34c759 !important;font-weight:600}
.dn-pct{color:#ff3b30 !important;font-weight:600}
.tr-up{color:#34c759 !important;font-weight:600}
.tr-dn{color:#ff3b30 !important;font-weight:600}
.tr-flat{color:#1a2a3a !important;font-size:14px !important}

/* Score circles */
.s4{
  width:28px;height:28px;border-radius:50%;
  display:inline-flex;align-items:center;justify-content:center;
  font-size:12px;font-weight:700;
  background:#0d2a16;color:#34c759;
  border:1.5px solid rgba(52,199,89,0.35);
}
.s3{
  width:28px;height:28px;border-radius:50%;
  display:inline-flex;align-items:center;justify-content:center;
  font-size:12px;font-weight:700;
  background:#2a1e00;color:#d4a843;
  border:1.5px solid rgba(212,168,67,0.35);
}
.s2{
  width:28px;height:28px;border-radius:50%;
  display:inline-flex;align-items:center;justify-content:center;
  font-size:12px;font-weight:700;
  background:#0d1520;color:#3a5068;
  border:1.5px solid #0f1923;
}
.s1{
  width:28px;height:28px;border-radius:50%;
  display:inline-flex;align-items:center;justify-content:center;
  font-size:12px;font-weight:700;
  background:#2a0a0a;color:#ff3b30;
  border:1.5px solid rgba(255,59,48,0.3);
}

/* ── Modal ──────────────────────────────────────────────────── */
.modal-overlay{
  display:none;position:fixed;
  inset:0;z-index:500;
  background:rgba(0,0,0,0.75);
  align-items:center;justify-content:center;
}
.modal-overlay.open{display:flex}
.modal-box{
  width:90%;max-width:1100px;
  background:#09090f;
  border:1px solid #0f1923;
  border-radius:12px;
  overflow:hidden;
}
.modal-head{
  background:#060810;
  border-bottom:1px solid #0f1923;
  padding:13px 20px;
  display:flex;align-items:center;justify-content:space-between;
}
.modal-lbl{
  font-size:11px;font-weight:700;
  letter-spacing:0.16em;text-transform:uppercase;
  color:#d4a843;
}
.modal-x{
  background:none;border:1px solid #0f1923;
  color:#3a5068;font-size:14px;
  cursor:pointer;border-radius:4px;
  padding:2px 9px;font-family:inherit;
  transition:all .15s;
}
.modal-x:hover{border-color:#3a5068;color:#cdd9e5}
#tv_container{height:580px}

/* ── Scrollbar ──────────────────────────────────────────────── */
::-webkit-scrollbar{width:4px;height:4px}
::-webkit-scrollbar-track{background:#07090d}
::-webkit-scrollbar-thumb{background:#0f1923;border-radius:2px}
::-webkit-scrollbar-thumb:hover{background:#1a2a3a}
</style>
</head>
<body>

<div class="topbar">
  <div class="brand">
    <div class="live-dot"></div>
    <span class="brand-name">Assaf Shtivi</span>
    <span class="brand-cmd">Command Center</span>
  </div>
  <div class="ts-pill">IDT """ + israel_time + """</div>
</div>

<div class="indices">""" + m_cards + """</div>

<div class="wrap">
  <div class="scan-bar">
    <div class="scan-lbl">Watchlist Scanner</div>
    <div class="cnt-badge">""" + str(len(results)) + """ stocks</div>
  </div>
  <div class="tbl-wrap">
    <table id="stockTable">
      <thead><tr>
        <th>Ticker</th>
        <th>Price</th><th>Day %</th><th>Score</th><th>Rank</th>
        <th>ADX</th><th>RSI</th><th>Breakout</th><th>Stop Loss</th><th>Trend</th>
      </tr></thead>
      <tbody>""" + rows + """</tbody>
    </table>
  </div>
</div>

<div class="modal-overlay" id="chartModal">
  <div class="modal-box">
    <div class="modal-head">
      <span class="modal-lbl" id="modalTicker">Live Chart</span>
      <button class="modal-x" onclick="closeModal()">&#x2715;</button>
    </div>
    <div id="tv_container"></div>
  </div>
</div>

<script src="https://code.jquery.com/jquery-3.7.0.min.js"></script>
<script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
<script src="https://s3.tradingview.com/tv.js"></script>
<script>
$(document).ready(function(){
  $('#stockTable').DataTable({
    order:[[4,'desc']],
    pageLength:50,
    lengthMenu:[25,50,100],
    language:{search:'',searchPlaceholder:'\u05d7\u05e4\u05e9 \u05d8\u05d9\u05e7\u05e8...'},
    columnDefs:[{targets:'_all',className:''}]
  });
});
function openChart(t){
  document.getElementById('modalTicker').textContent=t+' \u2014 Live Analysis';
  document.getElementById('tv_container').innerHTML='';
  document.getElementById('chartModal').classList.add('open');
  setTimeout(function(){
    new TradingView.widget({
      autosize:true,symbol:t,interval:'D',timezone:'Asia/Jerusalem',
      theme:'dark',style:'1',locale:'en',toolbar_bg:'#09090f',
      hide_side_toolbar:false,allow_symbol_change:true,
      container_id:'tv_container'
    });
  },200);
}
function closeModal(){
  document.getElementById('chartModal').classList.remove('open');
  document.getElementById('tv_container').innerHTML='';
}
document.getElementById('chartModal').addEventListener('click',function(e){
  if(e.target===this) closeModal();
});
</script>
</body></html>"""

    with open("index.html","w",encoding="utf-8") as f:
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
