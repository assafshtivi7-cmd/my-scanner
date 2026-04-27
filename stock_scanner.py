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
    "NASDAQ":      "^IXIC",
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

# ─── שליפת מדדים ─────────────────────────────────────────────────────────────

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

# ─── HTML (העיצוב החדש והמשופר) ───────────────────────────────────────────────

def generate_html(results, market_summary):
    il_tz       = timezone(timedelta(hours=3))
    israel_time = datetime.now(il_tz).strftime("%d/%m/%Y %H:%M")

    m_cards = "".join([f'''
        <div class="idx-card">
            <div class="idx-name">{m["name"]}</div>
            <div class="idx-price">{m["price"]}</div>
            <div class="idx-chg {{'chg-up' if m['color']=='up' else 'chg-dn'}}">{m["change"]}</div>
        </div>''' for m in market_summary])

    def score_class(s):
        return {4: "s4", 3: "s3", 2: "s2", 1: "s1", 0: "s0"}.get(s, "s0")

    rows = "".join([f'''
        <tr onclick="openChart('{s['Ticker']}')">
            <td><span class="ticker-badge">{s['Ticker']}</span></td>
            <td class="num">${{s['Price']:,.2f}}</td>
            <td class="num {{'chg-up' if s['Day_Chg_%'] > 0 else 'chg-dn'}}">{{s['Day_Chg_%']:+.2f}}%</td>
            <td><span class="score-pill {{score_class(s['SCORE'])}}">{{s['SCORE']}}</span></td>
            <td class="num gold">{{s['Power_Rank']}}</td>
            <td class="num cyan">{{s['ADX']}}</td>
            <td class="num">{{s['RSI']}}</td>
            <td class="num cyan-lt">${{s['Breakout']:,.2f}}</td>
            <td class="num warn">${{s['Stop_Loss']:,.2f}}</td>
            <td class="trend-cell {{'trend-up' if '↑' in s['TREND'] else 'trend-dn' if '↓' in s['TREND'] else ''}}">{{s['TREND']}}</td>
        </tr>''' for s in results])

    html = f'''
    <!DOCTYPE html>
    <html lang="he" dir="rtl">
    <head>
        <meta charset="UTF-8">
        <title>SHTIVI | COMMAND CENTER</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
        <style>
            :root {{
                --bg: #0a0a0a;
                --card-bg: #161616;
                --text: #e0e0e0;
                --gold: #fdbb2d;
                --cyan: #00bcd4;
                --up: #00ff88;
                --dn: #ff4d4d;
            }}
            body {{ background: var(--bg); color: var(--text); font-family: 'Segoe UI', sans-serif; margin: 20px; }}
            .dashboard-header {{ display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #333; padding-bottom: 10px; margin-bottom: 20px; }}
            
            .ticker-badge {{ 
                background-color: #fdbb2d !important; 
                color: #000000 !important; 
                padding: 4px 12px; 
                border-radius: 4px; 
                font-weight: 900; 
                font-size: 0.9em;
                display: inline-block;
                min-width: 80px;
                text-align: center;
            }}
            
            table {{ width: 100%; border-collapse: collapse; background: var(--card-bg); border-radius: 8px; overflow: hidden; }}
            th {{ text-align: center; padding: 12px; color: #888; font-size: 0.85em; text-transform: uppercase; background: #222; }}
            td {{ padding: 12px; border-bottom: 1px solid #222; text-align: center; }}
            tr:hover {{ background: #1e1e1e; cursor: pointer; }}
            
            .num {{ font-family: 'Courier New', monospace; font-weight: bold; }}
            .chg-up {{ color: var(--up); }}
            .chg-dn {{ color: var(--dn); }}
            .gold {{ color: var(--gold); }}
            .cyan {{ color: var(--cyan); }}
            .cyan-lt {{ color: #7ee8fa; }}
            .warn {{ color: #ff9800; }}
            
            .idx-container {{ display: flex; gap: 15px; margin-bottom: 20px; flex-wrap: wrap; }}
            .idx-card {{ background: var(--card-bg); padding: 15px; border-radius: 8px; border: 1px solid #333; min-width: 160px; text-align: center; }}
            .idx-name {{ font-size: 0.7em; color: #888; text-transform: uppercase; }}
            .idx-price {{ font-size: 1.4em; font-weight: bold; margin: 5px 0; }}
            
            .score-pill {{ width: 28px; height: 28px; display: inline-flex; align-items: center; justify-content: center; border-radius: 50%; font-weight: bold; }}
            .s4 {{ background: var(--up); color: #000; }}
            .s3 {{ background: var(--gold); color: #000; }}
            .s2 {{ background: #444; }}
            .s1, .s0 {{ background: var(--dn); color: #fff; }}

            /* Modal Style */
            #chartModal {{ display:none; position:fixed; top:5%; left:5%; width:90%; height:90%; background:#161616; border:2px solid #333; z-index:10000; padding:10px; border-radius: 12px; box-shadow: 0 0 50px rgba(0,0,0,0.8); }}
        </style>
    </head>
    <body>
        <div class="dashboard-header">
            <h1 style="font-weight: 800; letter-spacing: -1px;">SHTIVI <span style="color:var(--gold)">COMMAND CENTER</span></h1>
            <div style="background:#222; padding: 5px 15px; border-radius: 20px; font-weight: bold;">🇮🇱 IDT {israel_time}</div>
        </div>
        
        <div class="idx-container">{m_cards}</div>
        
        <table>
            <thead>
                <tr><th>Ticker</th><th>Price</th><th>Change</th><th>Score</th><th>Rank</th><th>ADX</th><th>RSI</th><th>Breakout</th><th>Stop</th><th>Trend</th></tr>
            </thead>
            <tbody>{rows}</tbody>
        </table>

        <div id="chartModal">
            <button onclick="document.getElementById('chartModal').style.display='none'" style="float:left; background:var(--gold); border:none; border-radius: 5px; cursor:pointer; font-weight:bold; padding:8px 15px; margin-bottom: 10px;">סגור X</button>
            <div id="tv_container" style="height:92%;"></div>
        </div>

        <script src="https://s3.tradingview.com/tv.js"></script>
        <script>
            function openChart(ticker) {{
                document.getElementById('chartModal').style.display = 'block';
                new TradingView.widget({{
                    "autosize": true, "symbol": ticker, "interval": "D", "timezone": "Asia/Jerusalem",
                    "theme": "dark", "style": "1", "locale": "en", "container_id": "tv_container",
                    "allow_symbol_change": true
                }});
            }}
        </script>
    </body>
    </html>
    '''

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

    hdr_fmt   = wb.add_format({'bold': True, 'bg_color': '#FFFF00', 'border': 1, 'align': 'center'})
    green_fmt = wb.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'align': 'center', 'border': 1})
    red_fmt   = wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'align': 'center', 'border': 1})
    base_fmt  = wb.add_format({'align': 'center', 'border': 1})

    for i, col in enumerate(df.columns):
        ws.write(0, i, col, hdr_fmt)
        ws.set_column(i, i, 12, base_fmt)

    ws.freeze_panes(1, 0)
    col_idx = {c: i for i, c in enumerate(df.columns)}

    if 'SCORE' in col_idx:
        c = col_idx['SCORE']
        ws.conditional_format(1, c, last_row, c, {'type': 'cell', 'criteria': '>=', 'value': 4, 'format': green_fmt})

    if 'Power_Rank' in col_idx:
        c = col_idx['Power_Rank']
        ws.conditional_format(1, c, last_row, c, {'type': '3_color_scale'})

    if 'TREND' in col_idx:
        c = col_idx['TREND']
        ws.conditional_format(1, c, last_row, c, {'type': 'text', 'criteria': 'containing', 'value': '↑', 'format': green_fmt})

    writer.close()

# ─── שליחת מייל ──────────────────────────────────────────────────────────────

def send_email(file_name):
    try:
        pwd = os.getenv("APP_PASSWORD", "").replace(" ", "")
        if not pwd:
            print("⚠️ APP_PASSWORD לא הוגדר")
            return
        msg = MIMEMultipart()
        msg['Subject'] = f"🚀 COMMAND CENTER — {datetime.now(timezone(timedelta(hours=3))).strftime('%d/%m/%Y %H:%M')}"
        msg['From'], msg['To'] = MY_EMAIL, MY_EMAIL
        msg.attach(MIMEText("הדוח היומי מוכן. האתר עודכן.", "plain", "utf-8"))
        with open(file_name, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read()); encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={file_name}")
            msg.attach(part)
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ssl.create_default_context()) as server:
            server.login(MY_EMAIL, pwd); server.send_message(msg)
        print("✅ מייל נשלח")
    except Exception as e: print(f"❌ שגיאה במייל: {e}")

# ─── Main ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    try:
        spy_df  = yf.download('SPY', period='1mo', progress=False)
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
    m_summary = fetch_market_summary()
    
    il_date = datetime.now(timezone(timedelta(hours=3))).strftime('%Y-%m-%d')
    f_name = f"Master_Scanner_{il_date}.xlsx"

    create_styled_excel(pd.DataFrame(results), f_name)
    generate_html(results, m_summary)
    json.dump({r['Ticker']: int(r['SCORE']) for r in results}, open(DB_FILE, "w"))
    send_email(f_name)
