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

# ─── הגדרות אישיות ───────────────────────────────────────────────────────────

SENDER_EMAIL = "assafshtivi7@gmail.com"
RECEIVER_EMAIL = "assafshtivi7@gmail.com"
APP_PASSWORD = "tvyx ixfd xwam adfh"

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

DB_FILE = "last_run.json"
MAX_WORKERS = 10 

# ─── פונקציות עזר ─────────────────────────────────────────────────────────────

def send_email(file_path):
    subject = f"📊 דוח סורק מניות יומי - {datetime.now().strftime('%d/%m/%Y')}"
    body = "היי אסף, מצורף הדוח המעודכן מהסורק. בהצלחה!"
    message = MIMEMultipart()
    message["From"] = SENDER_EMAIL
    message["To"] = RECEIVER_EMAIL
    message["Subject"] = subject
    message.attach(MIMEText(body, "plain"))

    with open(file_path, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
    
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename= {os.path.basename(file_path)}")
    message.attach(part)

    context = ssl.create_default_context()
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
            server.login(SENDER_EMAIL, APP_PASSWORD)
            server.sendmail(SENDER_EMAIL, RECEIVER_EMAIL, message.as_string())
        print("✅ המייל נשלח בהצלחה!")
    except Exception as e:
        print(f"❌ שגיאה בשליחת המייל: {e}")

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

        if score < 1 and not is_rev: return None

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
            'Status': "⚠️ REVERSAL" if is_rev else "⚡ SQUEEZE" if is_sqz else "-",
            'SCORE': score, 'Power_Rank': round(rank, 1), 'ADX': adx_val, 'RSI': rsi_val,
            'RS_vs_SPY': rs_vs_spy, 'Overext_%': round(overext, 1),
            'Day_Chg_%': round(((curr_p - float(close.iloc[-2])) / float(close.iloc[-2])) * 100, 2),
            'Breakout': round(float(df['High'].rolling(20).max().iloc[-1]), 2),
            'Stop_Loss': round(curr_p - (2 * float((df['High']-df['Low']).rolling(14).mean().iloc[-1])), 2),
            'TREND': trend
        }
    except: return None

# ─── הרצה ראשית ───────────────────────────────────────────────────────────────

def main():
    print("📡 שולף נתונים ומריץ סורק V18.0...")
    spy = yf.download('SPY', period='1mo', progress=False)['Close'].squeeze()
    spy_ret = float((spy.iloc[-1] - spy.iloc[0]) / spy.iloc[0])
    
    prev_scores = json.load(open(DB_FILE)) if os.path.exists(DB_FILE) else {}
    results = []
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(analyze_ticker, t, spy_ret, prev_scores): t for t in WATCHLIST}
        for f in as_completed(futures):
            res = f.result()
            if res: results.append(res)

    if not results:
        print("❌ אין תוצאות."); return

    df = pd.DataFrame(results).sort_values(by=['SCORE', 'Power_Rank'], ascending=[False, False])
    file_name = f"Master_Scanner_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
    
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Scanner')
    wb, ws = writer.book, writer.sheets['Scanner']
    
    header_f = wb.add_format({'bold': True, 'bg_color': '#FFFF00', 'border': 1, 'align': 'center'})
    green_f = wb.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'align': 'center'})
    red_f = wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'align': 'center'})

    for col_num, value in enumerate(df.columns.values):
        ws.write(0, col_num, value, header_f)
    ws.freeze_panes(1, 0)
    
    # צביעה מותנית
    last_row = len(df)
    ws.conditional_format(1, 3, last_row, 3, {'type': 'cell', 'criteria': '>=', 'value': 4, 'format': green_f})
    ws.conditional_format(1, 4, last_row, 4, {'type': '3_color_scale'})
    ws.conditional_format(1, 8, last_row, 8, {'type': 'cell', 'criteria': '>', 'value': 15, 'format': red_f})
    ws.conditional_format(1, 12, last_row, 12, {'type': 'text', 'criteria': 'containing', 'value': '↑', 'format': green_f})

    writer.close()
    json.dump({r['Ticker']: r['SCORE'] for r in results}, open(DB_FILE, "w"))
    
    print(f"✅ דוח מוכן: {file_name}")
    send_email(file_name)

if __name__ == "__main__":
    main()