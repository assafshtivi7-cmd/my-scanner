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

# --- הגדרות אישיות (מעודכן לפרטים הנכונים שלך) ---
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

MARKET_INDICES = {"S&P 500": "^GSPC", "NASDAQ": "^IXIC", "Bitcoin": "BTC-USD"}

# --- פונקציות חישוב (החזרתי את הלוגיקה הישנה) ---

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
        ema9 = float(close.ewm(span=9, adjust=False).mean().iloc[-1])
        ema21 = float(close.ewm(span=21, adjust=False).mean().iloc[-1])
        rvol = float(df['Volume'].iloc[-1] / df['Volume'].rolling(20).mean().iloc[-1])
        rsi_val = calc_rsi(close)
        
        is_rev = (close.diff().where(close.diff()<0,0).rolling(14).mean().abs() > close.diff().where(close.diff()>0,0).rolling(14).mean()).iloc[-15:].any() and curr_p > ema9
        is_sqz = float((df['High']-df['Low']).rolling(5).mean().iloc[-1]) < float((df['High']-df['Low']).rolling(20).mean().iloc[-1]) * 0.85

        score = 0
        if ema9 > ema21: score += 1
        if rvol > 1.2: score += 1
        if curr_p > close.rolling(200).mean().iloc[-1]: score += 1
        if -7 < ((curr_p / float(df['High'].max())) - 1) * 100 <= 0: score += 1
        if 40 < rsi_val < 70: score += 1

        adx_val = calc_adx(df)
        rs_vs_spy = round(((curr_p - float(close.iloc[-22])) / float(close.iloc[-22]) - spy_ret) * 100, 2)
        overext = ((curr_p / ema9) - 1) * 100

        rank = (score * 25) + (adx_val / 2) + min(rs_vs_spy / 4, 12)
        if is_sqz: rank += 20
        if is_rev: rank += 25
        if overext > 15: rank -= 30

        prev = prev_scores.get(ticker)
        trend = f"↑ ({prev})" if prev and score > prev else f"↓ ({prev})" if prev and score < prev else "-"

        return {
            'Ticker': ticker, 'Price': round(curr_p, 2),
            'Status': "⚠️ REV" if is_rev else "⚡ SQZ" if is_sqz else "-",
            'SCORE': score, 'Power_Rank': round(rank, 1), 'ADX': adx_val, 'RSI': rsi_val,
            'RS_vs_SPY': rs_vs_spy, 'Day_Chg_%': round(((curr_p - float(close.iloc[-2])) / float(close.iloc[-2])) * 100, 2),
            'Breakout': round(float(df['High'].rolling(20).max().iloc[-1]), 2),
            'Stop_Loss': round(curr_p - (2 * float((df['High']-df['Low']).rolling(14).mean().iloc[-1])), 2),
            'TREND': trend
        }
    except: return None

# --- פונקציות פלט (אתר, אקסל ומייל) ---

def generate_html(results, market_summary):
    now = datetime.now().strftime("%d/%m/%Y %H:%M")
    m_cards = "".join([f'<div class="col-md-4 mb-3"><div class="card text-center border-{m["color"]}"><div class="card-body"><h6>{m["name"]}</h6><h3>{m["price"]}</h3><span class="badge bg-{m["color"]}">{m["change"]}</span></div></div></div>' for m in market_summary])
    
    top_3 = sorted(results, key=lambda x: x['Power_Rank'], reverse=True)[:3]
    top_html = "".join([f'<div class="col-md-4"><div class="card bg-success text-white mb-3"><div class="card-body">🚀 {s["Ticker"]}: Rank {s["Power_Rank"]}</div></div></div>' for s in top_3])

    rows = "".join([f"<tr><td><strong>{s['Ticker']}</strong></td><td>{s['Price']}</td><td class='{'text-success' if s['Day_Chg_%'] > 0 else 'text-danger'}'>{s['Day_Chg_%']}%</td><td>{s['SCORE']}</td><td>{s['Power_Rank']}</td><td>{s['Status']}</td></tr>" for s in results[:15]])
    
    html = f'''<!DOCTYPE html><html lang="he" dir="rtl"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>דשבורד אסף</title><link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet"></head><body class="bg-light"><div class="container py-4 text-center"><h1>דוח סורק מניות</h1><p>{now}</p><div class="row">{m_cards}</div><h4 class="mt-4">🔥 מניות מובילות (Power Rank)</h4><div class="row">{top_html}</div><div class="card mt-4"><div class="card-header bg-dark text-white"><h5>Top 15 Stocks</h5></div><table class="table mb-0"><thead><tr><th>מניה</th><th>מחיר</th><th>שינוי</th><th>Score</th><th>Rank</th><th>סטטוס</th></tr></thead><tbody>{rows}</tbody></table></div></div></body></html>'''
    with open("index.html", "w", encoding="utf-8") as f: f.write(html)

def send_full_email(file_path):
    pwd = os.getenv("APP_PASSWORD").replace(" ", "")
    msg = MIMEMultipart()
    msg['Subject'] = f"📊 דוח סורק מלא - {datetime.now().strftime('%d/%m/%Y')}"
    msg['From'], msg['To'] = MY_EMAIL, MY_EMAIL
    msg.attach(MIMEText("היי אסף, מצורף דוח האקסל המלא עם כל הפרמטרים (RSI, ADX, Stop Loss).", "plain"))

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

    # נתוני מדדים לאתר
    m_summary = []
    for n, t in MARKET_INDICES.items():
        h = yf.Ticker(t).history(period="2d")
        p, c = h['Close'].iloc[-1], ((h['Close'].iloc[-1]-h['Close'].iloc[-2])/h['Close'].iloc[-2])*100
        m_summary.append({"name": n, "price": f"{p:,.2f}", "change": f"{c:+.2f}%", "color": "success" if c>=0 else "danger"})

    results.sort(key=lambda x: (x['SCORE'], x['Power_Rank']), reverse=True)
    
    # 1. ייצור אקסל
    f_name = f"Master_Scanner_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
    pd.DataFrame(results).to_excel(f_name, index=False)
    
    # 2. ייצור אתר
    generate_html(results, m_summary)
    
    # 3. שמירת היסטוריה
    json.dump({r['Ticker']: r['SCORE'] for r in results}, open(DB_FILE, "w"))
    
    # 4. שליחת מייל עם הקובץ
    try: send_full_email(f_name); print("Email Sent!")
    except Exception as e: print(f"Email Failed: {e}")
