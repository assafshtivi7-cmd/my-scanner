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
    "NASDAQ":      "^IXIC",
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
            'Ticker':      ticker,
            'Price':       round(curr_p, 2),
            'SCORE':       score,
            'Power_Rank': round(rank, 1),
            'ADX':         adx_val,
            'RSI':         rsi_val,
            'RVOL':        round(rvol, 2),
            'RS_vs_SPY':  rs_vs_spy,
            'Overext_%':  overext_pct,
            'Day_Chg_%':  day_chg,
            'Breakout':    round(float(df['High'].rolling(20).max().iloc[-1]), 2),
            'Stop_Loss':  stop_loss,
            'TREND':       trend,
        }
    except:
        return None

# ─── HTML ──────────────────────────────────────────────────────────────────────

def generate_html(results, market_summary):
    il_tz       = timezone(timedelta(hours=3))
    israel_time = datetime.now(il_tz).strftime("%d/%m/%Y %H:%M")

    m_cards = "".join([f'''<div class="idx">
      <div class="idx-bar {'up-bar' if m['color']=='up' else 'dn-bar'}"></div>
      <div class="idx-label">{m['name']}</div>
      <div class="idx-val">{m['price']}</div>
      <div class="idx-chg {'up' if m['color']=='up' else 'dn'}">{m['change']}</div>
    </div>''' for m in market_summary])

    def score_cls(s):
        return {4: "s4", 3: "s3", 2: "s2"}.get(s, "s1")

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
      <td class="trend-cell {'trend-up' if '↑' in s['TREND'] else 'trend-dn' if '↓' in s['TREND'] else ''}">{s['TREND']}</td>
    </tr>''' for s in results])

    html = f'''<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>SHTIVI | COMMAND CENTER</title>
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600;700&display=swap" rel="stylesheet">
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
<link href="https://cdn.datatables.net/1.13.6/css/dataTables.bootstrap5.min.css" rel="stylesheet">
<style>
:root{{
  --bg:#080c10;--surface:#0a0e14;--panel:#0d1420;--border:#1a2332;
  --gold:#fdbb2d;--text:#cbd5e1;--muted:#3a4a5c;--up:#3fb950;--dn:#f85149;
}}
body{{background:var(--bg);color:var(--text);font-family:'IBM Plex Mono',monospace;font-size:12px;margin:0}}
.topbar{{background:var(--surface);border-bottom:1px solid var(--border);padding:0 28px;height:48px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100}}
.ticker{{background:var(--gold)!important;color:#000!important;font-weight:800;padding:4px 10px;border-radius:4px;min-width:70px;display:inline-block;text-align:center;border:1px solid #000}}
.idx-bar{{position:absolute;top:0;left:0;bottom:0;width:2px}}.up-bar{{background:var(--up)}}.dn-bar{{background:var(--dn)}}
.indices{{display:grid;grid-template-columns:repeat(4,1fr);background:var(--border);border-bottom:1px solid var(--border)}}
.idx{{background:var(--surface);padding:16px 22px;position:relative}}
.table-wrap{{background:var(--panel);border:1px solid var(--border);border-radius:10px;margin:20px 28px;overflow:hidden}}
.score-circle{{width:26px;height:26px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-weight:700}}
.s4{{background:var(--up);color:#000}} .s3{{background:var(--gold);color:#000}} .s2{{background:#444}} .s1{{background:var(--dn)}}
</style>
</head>
<body>
<div class="topbar"><div>⚡ SHTIVI <b>COMMAND CENTER</b></div><div class="timestamp">IDT {israel_time}</div></div>
<div class="indices">{m_cards}</div>
<div class="table-wrap">
  <table id="stockTable" class="table table-dark table-hover mb-0">
    <thead><tr><th style="text-align:left">Ticker</th><th>Price</th><th>Day %</th><th>Score</th><th>Rank</th><th>ADX</th><th>RSI</th><th>Breakout</th><th>Stop</th><th>Trend</th></tr></thead>
    <tbody>{rows}</tbody>
  </table>
</div>
<div class="modal fade" id="chartModal" tabindex="-1"><div class="modal-dialog modal-xl modal-dialog-centered"><div class="modal-content" style="background:#111"><div class="modal-header"><h5>Live Analysis</h5><button class="btn-close btn-close-white" data-bs-dismiss="modal"></button></div><div class="modal-body"><div id="tv_chart" style="height:600px"></div></div></div></div></div>
<script src="https://code.jquery.com/jquery-3.7.0.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>
<script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
<script src="https://cdn.datatables.net/1.13.6/js/dataTables.bootstrap5.min.js"></script>
<script src="https://s3.tradingview.com/tv.js"></script>
<script>
$(document).ready(function(){{ $('#stockTable').DataTable({{ order: [[4, 'desc']], pageLength: 50 }}); }});
function openChart(ticker){{
  new bootstrap.Modal(document.getElementById('chartModal')).show();
  new TradingView.widget({{ autosize:true, symbol:ticker, interval:'D', timezone:'Asia/Jerusalem', theme:'dark', container_id:'tv_chart' }});
}}
</script>
</body></html>'''

    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("✅ index.html נוצר")

# ─── Excel & Email ────────────────────────────────────────────────────────────

def create_styled_excel(df, file_name):
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Scanner')
    wb, ws = writer.book, writer.sheets['Scanner']
    hdr = wb.add_format({'bold':True, 'bg_color':'#FFFF00', 'border':1})
    for i, col in enumerate(df.columns): ws.write(0, i, col, hdr)
    writer.close()

def send_email(file_name):
    pwd = os.getenv("APP_PASSWORD", "").replace(" ", "")
    if not pwd:
        print("⚠️ APP_PASSWORD לא הוגדר ב-GitHub Secrets")
        return
    msg = MIMEMultipart()
    msg['Subject'] = f"🚀 COMMAND CENTER — {datetime.now(timezone(timedelta(hours=3))).strftime('%d/%m/%Y %H:%M')}"
    msg['From'], msg['To'] = MY_EMAIL, MY_EMAIL
    msg.attach(MIMEText("הדוח היומי מוכן. האתר עודכן.", "plain", "utf-8"))
    with open(file_name, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read()); encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{file_name}"')
        msg.attach(part)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ssl.create_default_context()) as server:
        server.login(MY_EMAIL, pwd); server.send_message(msg)
    print("✅ מייל נשלח בהצלחה!")

# ─── Main ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    try:
        spy_df = yf.download('SPY', period='1mo', progress=False)
        spy_ret = float((spy_df['Close'].iloc[-1] - spy_df['Close'].iloc[0]) / spy_df['Close'].iloc[0])
    except: spy_ret = 0.0

    prev_scores = json.load(open(DB_FILE)) if os.path.exists(DB_FILE) else {}
    results = []
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(analyze_ticker, t, spy_ret, prev_scores): t for t in WATCHLIST}
        for future in as_completed(futures):
            res = future.result()
            if res: results.append(res)

    results.sort(key=lambda x: (x['SCORE'], x['Power_Rank']), reverse=True)
    m_summary = []
    for n, t in MARKET_INDICES.items():
        try:
            h = yf.Ticker(t).history(period="7d")
            p, c = h['Close'].iloc[-1], ((h['Close'].iloc[-1]-h['Close'].iloc[-2])/h['Close'].iloc[-2])*100
            m_summary.append({"name":n,"price":f"{p:,.2f}","change":f"{c:+.2f}%","color":"up" if c>=0 else "down"})
        except: continue

    il_date = datetime.now(timezone(timedelta(hours=3))).strftime('%Y-%m-%d')
    f_name = f"Master_Scanner_{il_date}.xlsx"
    create_styled_excel(pd.DataFrame(results), f_name)
    generate_html(results, m_summary)
    json.dump({r['Ticker']: int(r['SCORE']) for r in results}, open(DB_FILE, "w"))
    send_email(f_name)
