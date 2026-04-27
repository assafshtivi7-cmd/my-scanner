import yfinance as yf
import pandas as pd
import numpy as np
import json
import os
import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

# --- הגדרות ---
MY_EMAIL = "assafshtivi7@gmail.com"
DB_FILE = "last_run.json"
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

# הוספנו את ה-VIX
MARKET_INDICES = {
    "S&P 500": "^GSPC", 
    "NASDAQ": "^IXIC", 
    "Bitcoin": "BTC-USD",
    "VIX (Volatility)": "^VIX"
}

# --- לוגיקה טכנית ---
def calc_rsi(close, period=14):
    delta = close.diff()
    gain = delta.clip(lower=0).ewm(alpha=1/period, adjust=False).mean()
    loss = (-delta.clip(upper=0)).ewm(alpha=1/period, adjust=False).mean()
    rs = gain / loss.replace(0, np.nan)
    return round(float((100 - (100 / (1 + rs))).iloc[-1]), 1)

def calc_adx(df, period=14):
    h, l, c = df['High'], df['Low'], df['Close'].squeeze()
    tr = pd.concat([h-l, (h-c.shift()).abs(), (l-c.shift()).abs()], axis=1).max(axis=1)
    pdm = (h.diff()).clip(lower=0)
    mdm = (-l.diff()).clip(lower=0)
    atr = tr.ewm(alpha=1/period, adjust=False).mean()
    pdi = 100 * pdm.ewm(alpha=1/period, adjust=False).mean() / atr.replace(0, np.nan)
    mdi = 100 * mdm.ewm(alpha=1/period, adjust=False).mean() / atr.replace(0, np.nan)
    dx = 100 * (pdi - mdi).abs() / (pdi + mdi).replace(0, np.nan)
    return round(float(dx.ewm(alpha=1/period, adjust=False).mean().iloc[-1]), 1)

def analyze_ticker(ticker, spy_ret, prev_scores):
    try:
        df = yf.Ticker(ticker).history(period="1y")
        if len(df) < 50: return None
        close = df['Close'].squeeze()
        curr_p = float(close.iloc[-1])
        ema9, ema21 = close.ewm(span=9).mean().iloc[-1], close.ewm(span=21).mean().iloc[-1]
        rvol = float(df['Volume'].iloc[-1] / df['Volume'].rolling(20).mean().iloc[-1])
        rsi_val = calc_rsi(close)
        
        is_rev = (close.diff().where(close.diff()<0,0).rolling(14).mean().abs() > close.diff().where(close.diff()>0,0).rolling(14).mean()).iloc[-15:].any() and curr_p > ema9
        score = sum([ema9 > ema21, rvol > 1.2, curr_p > close.rolling(200).mean().iloc[-1], 40 < rsi_val < 70])

        adx_val = calc_adx(df)
        rs_vs_spy = round(((curr_p - float(close.iloc[-22])) / float(close.iloc[-22]) - spy_ret) * 100, 2)
        rank = (score * 25) + (adx_val / 2) + min(rs_vs_spy / 4, 12)

        prev = prev_scores.get(ticker)
        trend = "↑" if prev and score > prev else "↓" if prev and score < prev else "-"

        return {
            'Ticker': ticker, 'Price': round(curr_p, 2),
            'SCORE': score, 'Power_Rank': round(rank, 1), 'ADX': adx_val, 'RSI': rsi_val,
            'Day_Chg_%': round(((curr_p - float(close.iloc[-2])) / float(close.iloc[-2])) * 100, 2),
            'Breakout': round(float(df['High'].rolling(20).max().iloc[-1]), 2),
            'Stop_Loss': round(curr_p - (2 * float((df['High']-df['Low']).rolling(14).mean().iloc[-1])), 2),
            'TREND': trend
        }
    except: return None

# --- יצירת ה-HTML המפואר ---
def generate_html(results, market_summary):
    now = datetime.now().strftime("%d/%m/%Y %H:%M")
    
    # כרטיסי מדדים (כולל VIX)
    m_cards = "".join([f'''
        <div class="col-md-3 mb-3">
            <div class="card text-center shadow-sm border-{m["color"]}">
                <div class="card-body">
                    <small class="text-uppercase text-muted">{m["name"]}</small>
                    <h3 class="fw-bold">{m["price"]}</h3>
                    <span class="badge bg-{m["color"]}">{m["change"]}</span>
                </div>
            </div>
        </div>''' for m in market_summary])
    
    # טבלת מניות
    rows = "".join([f'''
        <tr onclick="showChart('{s['Ticker']}')" style="cursor: pointer;">
            <td><span class="fw-bold text-primary">{s['Ticker']}</span></td>
            <td>${s['Price']}</td>
            <td class="{'text-success' if s['Day_Chg_%'] > 0 else 'text-danger'}">{s['Day_Chg_%']}%</td>
            <td><span class="badge bg-dark">{s['SCORE']}</span></td>
            <td>{s['Power_Rank']}</td>
            <td class="text-info fw-bold">{s['Breakout']}</td>
            <td class="text-warning">{s['Stop_Loss']}</td>
            <td>{s['TREND']}</td>
        </tr>''' for s in results])

    html = f'''
    <!DOCTYPE html>
    <html lang="he" dir="rtl">
    <head>
        <meta charset="UTF-8">
        <title>ASSAF SHTIVI | COMMAND CENTER</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
        <style>
            body {{ background-color: #f4f7f6; font-family: 'Segoe UI', Tahoma; }}
            .navbar {{ background: #1a2a6c; background: linear-gradient(to right, #b21f1f, #fdbb2d, #1a2a6c); }}
            .table-hover tbody tr:hover {{ background-color: #e9ecef; transition: 0.3s; }}
            #chartModal .modal-dialog {{ max-width: 90%; }}
            .logo-box {{ background: white; padding: 5px 15px; border-radius: 8px; display: inline-block; }}
        </style>
    </head>
    <body>
        <nav class="navbar navbar-dark shadow mb-4">
            <div class="container text-center py-2">
                <div class="logo-box">
                    <h4 class="mb-0 fw-bold text-dark">📊 ASSAF SHTIVI | <span class="text-primary">COMMAND CENTER</span></h4>
                </div>
            </div>
        </nav>

        <div class="container">
            <div class="row mb-4">{m_cards}</div>
            
            <div class="card shadow-lg">
                <div class="card-header bg-white py-3">
                    <h5 class="mb-0 fw-bold">🔍 סורק מניות חכם - {now}</h5>
                    <small class="text-muted">לחץ על מניה כדי לפתוח גרף חי</small>
                </div>
                <div class="table-responsive">
                    <table class="table table-hover align-middle mb-0 text-center">
                        <thead class="table-light">
                            <tr>
                                <th>Ticker</th><th>Price</th><th>שינוי</th><th>Score</th><th>Rank</th><th>פריצה</th><th>Stop Loss</th><th>מגמה</th>
                            </tr>
                        </thead>
                        <tbody>{rows}</tbody>
                    </table>
                </div>
            </div>
        </div>

        <div class="modal fade" id="chartModal" tabindex="-1" aria-hidden="true">
            <div class="modal-dialog modal-dialog-centered">
                <div class="modal-content">
                    <div class="modal-header">
                        <h5 class="modal-title" id="modalTitle">Stock View</h5>
                        <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                    </div>
                    <div class="modal-body" style="height: 600px;">
                        <div id="tradingview_widget" style="height: 100%;"></div>
                    </div>
                </div>
            </div>
        </div>

        <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
        <script type="text/javascript" src="https://s3.tradingview.com/tv.js"></script>
        <script>
            function showChart(ticker) {{
                document.getElementById('modalTitle').innerText = ticker + " - Live Chart";
                var myModal = new bootstrap.Modal(document.getElementById('chartModal'));
                myModal.show();
                new TradingView.widget({{
                    "autosize": true, "symbol": ticker, "interval": "D", "timezone": "Etc/UTC",
                    "theme": "light", "style": "1", "locale": "en", "toolbar_bg": "#f1f3f6",
                    "enable_publishing": false, "allow_symbol_change": true, "container_id": "tradingview_widget"
                }});
            }}
        </script>
    </body>
    </html>'''
    with open("index.html", "w", encoding="utf-8") as f: f.write(html)

def send_full_email(file_path):
    pwd = os.getenv("APP_PASSWORD").replace(" ", "")
    msg = MIMEMultipart()
    msg['Subject'] = f"📊 COMMAND CENTER REPORT - {datetime.now().strftime('%d/%m/%Y')}"
    msg['From'], msg['To'] = MY_EMAIL, MY_EMAIL
    msg.attach(MIMEText("היי אסף, הדוח המלא מצורף. האתר שלך עודכן בעיצוב החדש כולל גרפים חיים.", "plain"))
    with open(file_path, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(file_path)}")
        msg.attach(part)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(MY_EMAIL, pwd)
        server.send_message(msg)

# --- הרצה ---
if __name__ == "__main__":
    spy = yf.download('SPY', period='1mo', progress=False)['Close'].squeeze()
    spy_ret = float((spy.iloc[-1] - spy.iloc[0]) / spy.iloc[0])
    prev_scores = json.load(open(DB_FILE)) if os.path.exists(DB_FILE) else {}
    results = []
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(analyze_ticker, t, spy_ret, prev_scores): t for t in WATCHLIST}
        for f in as_completed(futures):
            res = f.result()
            if res: results.append(res)
    
    m_summary = []
    for n, t in MARKET_INDICES.items():
        h = yf.Ticker(t).history(period="2d")
        p, c = h['Close'].iloc[-1], ((h['Close'].iloc[-1]-h['Close'].iloc[-2])/h['Close'].iloc[-2])*100
        m_summary.append({"name": n, "price": f"{p:,.2f}", "change": f"{c:+.2f}%", "color": "success" if c>=0 else "danger"})

    results.sort(key=lambda x: (x['SCORE'], x['Power_Rank']), reverse=True)
    f_name = f"Master_Scanner_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
    pd.DataFrame(results).to_excel(f_name, index=False)
    generate_html(results, m_summary)
    json.dump({r['Ticker']: r['SCORE'] for r in results}, open(DB_FILE, "w"))
    try: send_full_email(f_name); print("Success!")
    except Exception as e: print(f"Error: {e}")
