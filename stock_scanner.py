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


# ─── לוגיקה טכנית (ללא שינוי) ────────────────────────────────────────────────────────────

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


# ─── HTML (שדרוג נראות: רקע אפור + מדדים גדולים) ──────────────────────────────────

def generate_html(results, market_summary):
    from datetime import datetime, timedelta, timezone as tz
    il_tz       = tz(timedelta(hours=3))
    israel_time = datetime.now(il_tz).strftime("%d/%m/%Y %H:%M")

    def make_card(m):
        acc   = "up-acc" if m["color"] == "up" else "dn-acc"
        glow = "glow-up" if m["color"] == "up" else "glow-dn"
        chg   = "green"  if m["color"] == "up" else "red"
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
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
html{font-size:13px}
body{
  background:#1c1f26; /* רקע אפור כהה מקצועי */
  color:#cdd9e5;
  font-family:'IBM Plex Mono',monospace;
  min-height:100vh;
}

.topbar{
  background:#242933;
  border-bottom:1px solid #2d333b;
  height:60px;
  display:flex;
  align-items:center;
  justify-content:space-between;
  padding:0 30px;
  position:sticky;
  top:0;
  z-index:200;
}
.brand-name{font-size:14px;font-weight:600;letter-spacing:0.1em;color:#adbac7;text-transform:uppercase}
.brand-cmd{font-size:14px;font-weight:700;letter-spacing:0.1em;color:#fdbb2d;text-transform:uppercase}
.ts-pill{font-size:11px;color:#768390;border:1px solid #444c56;padding:4px 12px;border-radius:6px;background:#22272e}

.indices{
  display:grid;
  grid-template-columns:repeat(4,1fr);
  border-bottom:1px solid #2d333b;
  gap:1px;
  background:#2d333b;
}
.idx-card{
  background:#22272e;
  padding:30px 25px; /* הגדלת כרטיסים */
  position:relative;
  transition:background .2s;
}
.idx-side{position:absolute;top:0;left:0;bottom:0;width:4px}
.up-acc{background:#34c759}.dn-acc{background:#ff3b30}
.idx-lbl{
  font-size:13px; /* כותרת גדולה יותר */
  letter-spacing:0.15em;
  color:#768390;
  text-transform:uppercase;
  font-weight:700;
  margin-bottom:10px;
}
.idx-num{
  font-size:44px; /* מספרים גדולים וברורים */
  font-weight:700;
  color:#ffffff;
  letter-spacing:-0.03em;
  font-variant-numeric:tabular-nums;
  line-height:1;
}
.idx-chg{font-size:16px;font-weight:600;margin-top:8px}
.green{color:#34c759} .red{color:#ff3b30}

.wrap{padding:25px 30px}
.scan-lbl{font-size:11px;letter-spacing:0.2em;color:#fdbb2d;text-transform:uppercase;font-weight:700;display:flex;align-items:center;gap:10px}
.scan-lbl::before{content:'';width:4px;height:16px;background:#fdbb2d;border-radius:2px}

.tbl-wrap{border:1px solid #2d333b;border-radius:12px;overflow:hidden;margin-top:15px}
#stockTable{width:100%!important;background:#22272e;border-collapse:collapse}
#stockTable thead th{background:#2d333b!important;color:#adbac7!important;font-size:10px!important;padding:15px!important;text-transform:uppercase!important;border-bottom:1px solid #444c56!important}
#stockTable tbody tr{border-bottom:1px solid #2d333b!important;transition:background .1s}
#stockTable tbody tr:hover{background:#2d333b!important}
#stockTable tbody td{padding:14px!important;color:#adbac7;font-size:13px}

.tkr{background:#fdbb2d;color:#000!important;font-weight:800;padding:5px 12px;border-radius:6px;min-width:70px;display:inline-flex;justify-content:center}
.px{color:#ffffff!important;font-weight:600}
.up-pct{color:#34c759!important;font-weight:600}
.dn-pct{color:#ff3b30!important;font-weight:600}

/* Score circles */
.s4{width:32px;height:32px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;background:#1d3323;color:#34c759;border:2px solid #34c759;font-weight:700}
.s3{width:32px;height:32px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;background:#332b1d;color:#fdbb2d;border:2px solid #fdbb2d;font-weight:700}
.s2{width:32px;height:32px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;background:#2d333b;color:#768390;border:1px solid #444c56}

.modal-overlay{display:none;position:fixed;inset:0;z-index:500;background:rgba(0,0,0,0.85);align-items:center;justify-content:center}
.modal-overlay.open{display:flex}
.modal-box{width:95%;max-width:1200px;background:#1c1f26;border:1px solid #444c56;border-radius:15px;overflow:hidden}
.modal-head{background:#22272e;padding:15px 25px;display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid #2d333b}
#tv_container{height:650px}
</style>
</head>
<body>
<div class="topbar">
  <div class="brand">
    <div style="width:8px;height:8px;background:#34c759;border-radius:50%"></div>
    <span class="brand-name">Assaf Shtivi</span>
    <span class="brand-cmd">Command Center</span>
  </div>
  <div class="ts-pill">IDT """ + israel_time + """</div>
</div>
<div class="indices">""" + m_cards + """</div>
<div class="wrap">
  <div class="scan-lbl">Watchlist Scanner <span style="font-size:9px;color:#768390;margin-right:15px">""" + str(len(results)) + """ stocks</span></div>
  <div class="tbl-wrap">
    <table id="stockTable">
      <thead><tr><th>Ticker</th><th>Price</th><th>Day %</th><th>Score</th><th>Rank</th><th>ADX</th><th>RSI</th><th>Breakout</th><th>Stop Loss</th><th>Trend</th></tr></thead>
      <tbody>""" + rows + """</tbody>
    </table>
  </div>
</div>
<div class="modal-overlay" id="chartModal">
  <div class="modal-box">
    <div class="modal-head"><span id="modalTicker" style="color:#fdbb2d;font-weight:700">Live Chart</span><button onclick="closeModal()" style="background:none;border:none;color:#768390;font-size:20px;cursor:pointer">✕</button></div>
    <div id="tv_container"></div>
  </div>
</div>
<script src="https://code.jquery.com/jquery-3.7.0.min.js"></script>
<script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
<script src="https://s3.tradingview.com/tv.js"></script>
<script>
$(document).ready(function(){ $('#stockTable').DataTable({order:[[4,'desc']],pageLength:50,language:{search:'',searchPlaceholder:'חיפוש טיקר...'}}); });
function openChart(t){
  document.getElementById('modalTicker').textContent=t+' — Analysis';
  document.getElementById('tv_container').innerHTML='';
  document.getElementById('chartModal').classList.add('open');
  setTimeout(function(){
    new TradingView.widget({autosize:true,symbol:t,interval:'D',timezone:'Asia/Jerusalem',theme:'dark',style:'1',container_id:'tv_container'});
  },200);
}
function closeModal(){document.getElementById('chartModal').classList.remove('open');document.getElementById('tv_container').innerHTML='';}
</script>
</body></html>"""

    with open("index.html","w",encoding="utf-8") as f:
        f.write(html)
    print("✅ index.html נוצר")


# ─── Excel (ללא שינוי לוגי) ────────────────────────────────────────────────────────────

def create_styled_excel(df, file_name):
    COLS = ['Ticker','Price','SCORE','Power_Rank','ADX','RSI','RVOL',
            'RS_vs_SPY','Overext_%','Day_Chg_%','Breakout','Stop_Loss','TREND']
    df = df[[c for c in COLS if c in df.columns]].copy()

    writer   = pd.ExcelWriter(file_name, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Scanner')
    wb, ws   = writer.book, writer.sheets['Scanner']
    last_row = len(df)

    hdr_fmt = wb.add_format({'bold': True, 'bg_color': '#1a1a2e', 'font_color': '#c9aa71', 'border': 1, 'align': 'center'})
    base_fmt = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
    green_fmt = wb.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'align': 'center', 'border': 1})
    red_fmt = wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'align': 'center', 'border': 1})

    for i, col in enumerate(df.columns):
        ws.write(0, i, col, hdr_fmt)
        ws.set_column(i, i, 12)

    for row_num in range(1, last_row + 1):
        for col_num in range(len(df.columns)):
            ws.write(row_num, col_num, df.iloc[row_num - 1, col_num], base_fmt)

    col_idx = {c: i for i, c in enumerate(df.columns)}
    if 'SCORE' in col_idx:
        c = col_idx['SCORE']
        ws.conditional_format(1, c, last_row, c, {'type': 'cell', 'criteria': '>=', 'value': 4, 'format': green_fmt})
    
    writer.close()
    print(f"✅ Excel נוצר: {file_name}")


# ─── שליחת מייל (ללא שינוי) ───────────────────────────────────────────────────────────────

def send_email(file_name):
    pwd = os.getenv("APP_PASSWORD", "").replace(" ", "")
    if not pwd: return
    il_time = datetime.now(timezone(timedelta(hours=3))).strftime('%d/%m/%Y %H:%M')
    msg = MIMEMultipart()
    msg['Subject'] = f"🚀 COMMAND CENTER — {il_time} IDT"
    msg['From'], msg['To'] = MY_EMAIL, MY_EMAIL
    msg.attach(MIMEText(f"הסריקה הושלמה: {il_time}\nהאתר עודכן.", "plain", "utf-8"))
    with open(file_name, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{os.path.basename(file_name)}"')
        msg.attach(part)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ssl.create_default_context()) as server:
        server.login(MY_EMAIL, pwd)
        server.send_message(msg)

if __name__ == "__main__":
    spy_df = yf.download('SPY', period='1mo', progress=False)
    spy_ret = float((spy_df['Close'].iloc[-1] - spy_df['Close'].iloc[0]) / spy_df['Close'].iloc[0])
    prev_scores = json.load(open(DB_FILE)) if os.path.exists(DB_FILE) else {}
    
    results = []
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(analyze_ticker, t, spy_ret, prev_scores): t for t in WATCHLIST}
        for future in as_completed(futures):
            res = future.result()
            if res: results.append(res)

    results.sort(key=lambda x: (x['SCORE'], x['Power_Rank']), reverse=True)
    market_summary = fetch_market_summary()
    il_date = datetime.now(timezone(timedelta(hours=3))).strftime('%Y-%m-%d')
    file_name = f"Master_Scanner_{il_date}.xlsx"
    
    create_styled_excel(pd.DataFrame(results), file_name)
    generate_html(results, market_summary)
    json.dump({r['Ticker']: int(r['SCORE']) for r in results}, open(DB_FILE, "w"))
    send_email(file_name)
