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

MARKET_INDICES = {
    "S&P 500": "^GSPC", 
    "NASDAQ": "^IXIC", 
    "BITCOIN": "BTC-USD",
    "VIX (VOLATILITY)": "^VIX"
}

# --- לוגיקה טכנית ---
def calc_rsi(close, period=14):
    try:
        delta = close.diff()
        gain = delta.clip(lower=0).ewm(alpha=1/period, adjust=False).mean()
        loss = (-delta.clip(upper=0)).ewm(alpha=1/period, adjust=False).mean()
        rs = gain / loss.replace(0, np.nan)
        return round(float((100 - (100 / (1 + rs))).iloc[-1]), 1)
    except: return 50

def calc_adx(df, period=14):
    try:
        h, l, c = df['High'], df['Low'], df['Close'].squeeze()
        tr = pd.concat([h-l, (h-c.shift()).abs(), (l-c.shift()).abs()], axis=1).max(axis=1)
        pdm = (h.diff()).clip(lower=0)
        mdm = (-l.diff()).clip(lower=0)
        atr = tr.ewm(alpha=1/period, adjust=False).mean()
        pdi = 100 * pdm.ewm(alpha=1/period, adjust=False).mean() / atr.replace(0, np.nan)
        mdi = 100 * mdm.ewm(alpha=1/period, adjust=False).mean() / atr.replace(0, np.nan)
        dx = 100 * (pdi - mdi).abs() / (pdi + mdi).replace(0, np.nan)
        return round(float(dx.ewm(alpha=1/period, adjust=False).mean().iloc[-1]), 1)
    except: return 20

def analyze_ticker(ticker, spy_ret, prev_scores):
    try:
        df = yf.Ticker(ticker).history(period="1y")
        if df.empty or len(df) < 22: return None
        close = df['Close'].squeeze()
        curr_p = float(close.iloc[-1])
        ema9, ema21 = close.ewm(span=9).mean().iloc[-1], close.ewm(span=21).mean().iloc[-1]
        rvol = float(df['Volume'].iloc[-1] / df['Volume'].rolling(20).mean().iloc[-1])
        rsi_val = calc_rsi(close)
        
        score = int(sum([ema9 > ema21, rvol > 1.1, curr_p > close.rolling(200).mean().iloc[-1], 40 < rsi_val < 75]))
        adx_val = calc_adx(df)
        rs_vs_spy = round(((curr_p - float(close.iloc[-22])) / float(close.iloc[-22]) - spy_ret) * 100, 2)
        rank = (score * 25) + (adx_val / 2) + min(rs_vs_spy / 4, 12)

        prev = prev_scores.get(ticker)
        trend = "↑" if prev and score > prev else "↓" if prev and score < prev else "-"

        return {
            'Ticker': ticker, 'Price': round(curr_p, 2), 'SCORE': score, 'Power_Rank': round(rank, 1),
            'ADX': adx_val, 'RSI': rsi_val, 'Day_Chg_%': round(((curr_p - float(close.iloc[-2])) / float(close.iloc[-2])) * 100, 2),
            'Breakout': round(float(df['High'].rolling(20).max().iloc[-1]), 2),
            'Stop_Loss': round(curr_p - (2 * float((df['High']-df['Low']).rolling(14).mean().iloc[-1])), 2),
            'TREND': trend
        }
    except: return None

def generate_html(results, market_summary):
    now = datetime.now().strftime("%d/%m/%Y %H:%M")
    m_cards = "".join([f'''
        <div class="col-md-3 mb-3">
            <div class="index-card text-center p-3">
                <small class="text-muted d-block mb-1">{m["name"]}</small>
                <div class="h3 mb-0 fw-bold">{m["price"]}</div>
                <div class="text-{m["color"]} small fw-bold">{m["change"]}</div>
            </div>
        </div>''' for m in market_summary])
    
    rows = "".join([f'''
        <tr onclick="showChart('{s['Ticker']}')">
            <td><div class="d-flex align-items-center justify-content-center"><span class="ticker-badge">{s['Ticker']}</span></div></td>
            <td>${s['Price']}</td>
            <td class="fw-bold text-{'success' if s['Day_Chg_%'] > 0 else 'danger'}">{s['Day_Chg_%']}%</td>
            <td><span class="score-dot score-{s['SCORE']}">{s['SCORE']}</span></td>
            <td class="fw-bold">{s['Power_Rank']}</td>
            <td class="text-info">${s['Breakout']}</td>
            <td class="text-warning">${s['Stop_Loss']}</td>
            <td>{s['TREND']}</td>
        </tr>''' for s in results])

    html = f'''
    <!DOCTYPE html>
    <html lang="he" dir="rtl">
    <head>
        <meta charset="UTF-8">
        <title>SHTIVI | PRO COMMAND CENTER</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
        <link rel="stylesheet" href="https://cdn.datatables.net/1.13.4/css/dataTables.bootstrap5.min.css">
        <style>
            :root {{ --bg-color: #0b0e11; --card-bg: #1e2329; --text-main: #eaecef; --accent: #f0b90b; }}
            body {{ background-color: var(--bg-color); color: var(--text-main); font-family: 'Inter', sans-serif; }}
            .navbar {{ background-color: var(--card-bg); border-bottom: 1px solid #333; }}
            .index-card {{ background: var(--card-bg); border-radius: 12px; border: 1px solid #333; }}
            .ticker-badge {{ background: #2b3139; padding: 4px 12px; border-radius: 6px; color: var(--accent); font-weight: bold; }}
            .table {{ background-color: var(--card-bg); color: var(--text-main); border-color: #333; border-radius: 12px; overflow: hidden; }}
            .table thead {{ background-color: #2b3139; }}
            .score-dot {{ width: 28px; height: 28px; display: inline-flex; align-items: center; justify-content: center; border-radius: 50%; font-weight: bold; }}
            .score-4 {{ background: #0ecb81; color: #fff; }} .score-3 {{ background: #f0b90b; color: #000; }}
            .dataTables_wrapper .dataTables_length, .dataTables_wrapper .dataTables_filter {{ color: var(--text-main) !important; margin-bottom: 15px; }}
            tr:hover {{ background-color: #2b3139 !important; cursor: pointer; transition: 0.2s; }}
            .modal-content {{ background-color: var(--card-bg); color: var(--text-main); }}
        </style>
    </head>
    <body>
        <nav class="navbar mb-4 py-3 shadow-sm"><div class="container d-flex justify-content-between align-items-center">
            <span class="h4 mb-0 fw-bold text-uppercase" style="letter-spacing: 2px;">⚡ ASSAF SHTIVI <span style="color: var(--accent);">PRO SCANNER</span></span>
            <span class="text-muted small">{now}</span>
        </div></nav>
        <div class="container">
            <div class="row mb-4">{m_cards}</div>
            <div class="card p-4 shadow" style="background: var(--card-bg); border: none; border-radius: 16px;">
                <table id="stockTable" class="table table-hover text-center align-middle">
                    <thead><tr><th>Ticker</th><th>Price</th><th>Daily %</th><th>Score</th><th>Rank</th><th>Breakout</th><th>Stop Loss</th><th>Trend</th></tr></thead>
                    <tbody>{rows}</tbody>
                </table>
            </div>
        </div>
        <div class="modal fade" id="chartModal" tabindex="-1"><div class="modal-dialog modal-xl modal-dialog-centered">
            <div class="modal-content"><div class="modal-header border-secondary"><h5 class="modal-title" id="modalTitle">Chart</h5><button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button></div>
            <div class="modal-body" style="height: 700px;"><div id="tv_widget" style="height: 100%;"></div></div></div>
        </div></div>
        <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
        <script src="https://cdn.datatables.net/1.13.4/js/jquery.dataTables.min.js"></script>
        <script src="https://cdn.datatables.net/1.13.4/js/dataTables.bootstrap5.min.js"></script>
        <script type="text/javascript" src="https://s3.tradingview.com/tv.js"></script>
        <script>
            $(document).ready(function() {{ $('#stockTable').DataTable({{ "order": [[4, "desc"]], "pageLength": 25, "language": {{ "search": "חיפוש מניה:" }} }}); }});
            function showChart(ticker) {{
                document.getElementById('modalTitle').innerText = ticker + " - Real-Time Analysis";
                new bootstrap.Modal(document.getElementById('chartModal')).show();
                new TradingView.widget({{ "autosize": true, "symbol": ticker, "interval": "D", "timezone": "Etc/UTC", "theme": "dark", "style": "1", "locale": "en", "container_id": "tv_widget" }});
            }}
        </script>
    </body></html>'''
    with open("index.html", "w", encoding="utf-8") as f: f.write(html)

def send_full_email(file_path):
    pwd = os.getenv("APP_PASSWORD")
    if not pwd: return
    msg = MIMEMultipart()
    msg['Subject'] = f"🚀 PRO COMMAND CENTER REPORT - {datetime.now().strftime('%d/%m/%Y')}"
    msg['From'], msg['To'] = MY_EMAIL, MY_EMAIL
    msg.attach(MIMEText("אסף, האתר שודרג לגרסת ה-PRO. מצורף קובץ האקסל המלא.", "plain"))
    with open(file_path, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(file_path)}")
        msg.attach(part)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(MY_EMAIL, pwd.replace(" ", ""))
        server.send_message(msg)

if __name__ == "__main__":
    try:
        spy_df = yf.download('SPY', period='1mo', progress=False)
        spy_ret = float((spy_df['Close'].iloc[-1] - spy_df['Close'].iloc[0]) / spy_df['Close'].iloc[0])
    except: spy_ret = 0
    
    prev_scores = json.load(open(DB_FILE)) if os.path.exists(DB_FILE) else {}
    results = []
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(analyze_ticker, t, spy_ret, prev_scores): t for t in WATCHLIST}
        for f in as_completed(futures):
            res = f.result()
            if res: results.append(res)
    
    m_summary = []
    for n, t in MARKET_INDICES.items():
        try:
            h = yf.Ticker(t).history(period="5d")
            p, c = h['Close'].iloc[-1], ((h['Close'].iloc[-1]-h['Close'].iloc[-2])/h['Close'].iloc[-2])*100
            m_summary.append({"name": n, "price": f"{p:,.2f}", "change": f"{c:+.2f}%", "color": "success" if c>=0 else "danger"})
        except: continue

    results.sort(key=lambda x: (x['SCORE'], x['Power_Rank']), reverse=True)
    f_name = f"Master_Scanner_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
    pd.DataFrame(results).to_excel(f_name, index=False)
    generate_html(results, m_summary)
    json.dump({r['Ticker']: int(r['SCORE']) for r in results}, open(DB_FILE, "w"))
    try: send_full_email(f_name); print("Success!")
    except Exception as e: print(f"Error: {e}")
