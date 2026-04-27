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

# --- הגדרות אישיות ---
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
    "S&P 500": "^GSPC", "NASDAQ": "^IXIC", 
    "BITCOIN": "BTC-USD", "VIX (FEAR)": "^VIX"
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
        trend = f"↑ ({prev})" if prev and score > prev else f"↓ ({prev})" if prev and score < prev else "-"

        return {
            'Ticker': ticker, 'Price': round(curr_p, 2), 'SCORE': score, 'Power_Rank': round(rank, 1),
            'ADX': adx_val, 'RSI': rsi_val, 'Day_Chg_%': round(((curr_p - float(close.iloc[-2])) / float(close.iloc[-2])) * 100, 2),
            'Breakout': round(float(df['High'].rolling(20).max().iloc[-1]), 2),
            'Stop_Loss': round(curr_p - (2 * float((df['High']-df['Low']).rolling(14).mean().iloc[-1])), 2),
            'TREND': trend
        }
    except: return None

# --- תצוגה ---
def generate_html(results, market_summary):
    now = datetime.now().strftime("%d/%m/%Y %H:%M")
    m_cards = "".join([f'''
        <div class="col-md-3 mb-3">
            <div class="index-card text-center p-3">
                <small class="text-muted d-block mb-1 text-uppercase">{m["name"]}</small>
                <div class="h3 mb-0 fw-bold" style="color: #fff;">{m["price"]}</div>
                <div class="text-{m["color"]} small fw-bold">{m["change"]}</div>
            </div>
        </div>''' for m in market_summary])
    
    rows = "".join([f'''
        <tr onclick="showChart('{s['Ticker']}')">
            <td><div class="ticker-badge">{s['Ticker']}</div></td>
            <td class="fw-bold text-white">${s['Price']}</td>
            <td class="fw-bold text-{'success' if s['Day_Chg_%'] > 0 else 'danger'}">{s['Day_Chg_%']}%</td>
            <td><span class="score-dot score-{s['SCORE']}">{s['SCORE']}</span></td>
            <td class="fw-bold" style="color: #f0b90b;">{s['Power_Rank']}</td>
            <td style="color: #00d2ff;">${s['Breakout']}</td>
            <td style="color: #ff9f43;">${s['Stop_Loss']}</td>
            <td class="text-white">{s['TREND']}</td>
        </tr>''' for s in results])

    html = f'''
    <!DOCTYPE html>
    <html lang="he" dir="rtl">
    <head>
        <meta charset="UTF-8">
        <title>SHTIVI | COMMAND CENTER</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
        <link rel="stylesheet" href="https://cdn.datatables.net/1.13.4/css/dataTables.bootstrap5.min.css">
        <style>
            :root {{ --bg: #0b0e11; --card: #1e2329; --text: #eaecef; --accent: #f0b90b; }}
            body {{ background-color: var(--bg); color: var(--text); font-family: 'Inter', sans-serif; }}
            .navbar {{ background-color: var(--card); border-bottom: 2px solid var(--accent); }}
            .index-card {{ background: var(--card); border-radius: 12px; border: 1px solid #333; }}
            .ticker-badge {{ background: var(--accent); padding: 5px 12px; border-radius: 6px; color: #000; font-weight: 800; display: inline-block; min-width: 80px; }}
            .table {{ background-color: var(--card); color: var(--text); border-radius: 12px; overflow: hidden; }}
            .score-dot {{ width: 30px; height: 30px; display: inline-flex; align-items: center; justify-content: center; border-radius: 50%; font-weight: bold; }}
            .score-4 {{ background: #0ecb81; color: #fff; }} .score-3 {{ background: #f0b90b; color: #000; }}
            tr:hover {{ background-color: #2b3139 !important; cursor: pointer; }}
            .dataTables_filter input {{ background: #2b3139; color: white; border: 1px solid #444; }}
        </style>
    </head>
    <body>
        <nav class="navbar mb-4 py-3 shadow"><div class="container d-flex justify-content-between align-items-center">
            <span class="h4 mb-0 fw-bold">⚡ ASSAF SHTIVI <span style="color: var(--accent);">COMMAND CENTER</span></span>
            <span class="badge bg-dark">{now}</span>
        </div></nav>
        <div class="container">
            <div class="row mb-4">{m_cards}</div>
            <div class="card p-4 shadow-lg" style="background: var(--card); border: none; border-radius: 16px;">
                <table id="stockTable" class="table table-hover text-center align-middle">
                    <thead><tr><th>Ticker</th><th>Price</th><th>Daily %</th><th>Score</th><th>Rank</th><th>Breakout</th><th>Stop Loss</th><th>Trend</th></tr></thead>
                    <tbody>{rows}</tbody>
                </table>
            </div>
        </div>
        <div class="modal fade" id="chartModal" tabindex="-1"><div class="modal-dialog modal-xl modal-dialog-centered">
            <div class="modal-content" style="background: #161a1e;"><div class="modal-header border-secondary text-white"><h5>Live Analysis</h5><button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button></div>
            <div class="modal-body" style="height: 750px;"><div id="tv_widget" style="height: 100%;"></div></div></div>
        </div></div>
        <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
        <script src="https://cdn.datatables.net/1.13.4/js/jquery.dataTables.min.js"></script>
        <script src="https://cdn.datatables.net/1.13.4/js/dataTables.bootstrap5.min.js"></script>
        <script type="text/javascript" src="https://s3.tradingview.com/tv.js"></script>
        <script>
            $(document).ready(function() {{ $('#stockTable').DataTable({{ "order": [[4, "desc"]], "pageLength": 50, "language": {{ "search": "חיפוש:" }} }}); }});
            function showChart(ticker) {{
                new bootstrap.Modal(document.getElementById('chartModal')).show();
                new TradingView.widget({{ "autosize": true, "symbol": ticker, "interval": "D", "timezone": "Etc/UTC", "theme": "dark", "style": "1", "locale": "en", "container_id": "tv_widget" }});
            }}
        </script>
    </body></html>'''
    with open("index.html", "w", encoding="utf-8") as f: f.write(html)

def create_styled_excel(df, file_name):
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Scanner')
    workbook, worksheet = writer.book, writer.sheets['Scanner']
    
    # הגדרת פורמטים
    header_f = workbook.add_format({'bold': True, 'bg_color': '#FFFF00', 'border': 1, 'align': 'center'})
    green_f = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'align': 'center'})
    
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_f)
    
    # החלת צביעה מותנית
    last_row = len(df)
    # SCORE >= 4
    worksheet.conditional_format(1, 2, last_row, 2, {'type': 'cell', 'criteria': '>=', 'value': 4, 'format': green_f})
    # Power_Rank - סקאלה
    worksheet.conditional_format(1, 3, last_row, 3, {'type': '3_color_scale'})
    # TREND - חץ למעלה נצבע בירוק
    worksheet.conditional_format(1, 9, last_row, 9, {'type': 'text', 'criteria': 'containing', 'value': '↑', 'format': green_f})
    
    writer.close()

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
            res = f.result(); 
            if res: results.append(res)

    m_summary = []
    for n, t in MARKET_INDICES.items():
        try:
            h = yf.Ticker(t).history(period="7d")
            p, c = h['Close'].iloc[-1], ((h['Close'].iloc[-1]-h['Close'].iloc[-2])/h['Close'].iloc[-2])*100
            m_summary.append({"name": n, "price": f"{p:,.2f}", "change": f"{c:+.2f}%", "color": "success" if c>=0 else "danger"})
        except: continue

    results.sort(key=lambda x: (x['SCORE'], x['Power_Rank']), reverse=True)
    f_name = f"Master_Scanner_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
    create_styled_excel(pd.DataFrame(results), f_name)
    generate_html(results, m_summary)
    json.dump({r['Ticker']: int(r['SCORE']) for r in results}, open(DB_FILE, "w"))
    
    try:
        pwd = os.getenv("APP_PASSWORD").replace(" ", "")
        msg = MIMEMultipart()
        msg['Subject'] = f"🚀 COMMAND CENTER REPORT - {datetime.now().strftime('%d/%m/%Y')}"
        msg['From'], msg['To'] = MY_EMAIL, MY_EMAIL
        msg.attach(MIMEText("אסף, הדוח והאתר מוכנים בגרסת ה-Ultimate. הטיקרים והצבעים תוקנו.", "plain"))
        with open(f_name, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read()); encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={f_name}")
            msg.attach(part)
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(MY_EMAIL, pwd); server.send_message(msg)
        print("Success!")
    except Exception as e: print(f"Error: {e}")
