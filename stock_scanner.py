import yfinance as yf
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
import os

WATCHLIST = {
    "AI & Tech": ["PLTR", "SOFI", "AMD", "NVDA", "MSFT"],
    "Crypto & Growth": ["MSTR", "MARA", "COIN", "TSLA"],
    "Portfolio": ["BMNR", "SEDG", "BTG"]
}
MARKET_INDICES = {"S&P 500": "^GSPC", "NASDAQ": "^IXIC", "Bitcoin": "BTC-USD"}

def get_data():
    results, market_summary = {}, []
    email_body = "Stock Scan Report\n"
    
    for name, ticker in MARKET_INDICES.items():
        try:
            data = yf.Ticker(ticker).history(period="2d")
            price = data['Close'].iloc[-1]
            change = ((price - data['Close'].iloc[-2]) / data['Close'].iloc[-2]) * 100
            market_summary.append({"name": name, "price": f"{price:,.2f}", "change": f"{change:+.2f}%", "color": "success" if change >= 0 else "danger"})
        except: continue

    for cat, tickers in WATCHLIST.items():
        results[cat] = []
        email_body += f"\n-- {cat} --\n"
        for t in tickers:
            try:
                s = yf.Ticker(t).history(period="2d")
                curr = s['Close'].iloc[-1]
                chg = ((curr - s['Close'].iloc[-2]) / s['Close'].iloc[-2]) * 100
                results[cat].append({"ticker": t, "price": f"{curr:.2f}", "change": f"{chg:+.2f}%"})
                email_body += f"{t}: {curr:.2f}$ ({chg:+.2f}%)\n"
            except: continue
    return results, market_summary, email_body

def send_email(content):
    pwd = os.getenv("APP_PASSWORD")
    if not pwd: return
    try:
        msg = MIMEMultipart()
        msg['Subject'] = f"Stock Report {datetime.now().strftime('%d/%m/%Y')}"
        msg['From'] = "ashtivi@gmail.com"
        msg['To'] = "ashtivi@gmail.com"
        msg.attach(MIMEText(content, 'plain'))
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login("ashtivi@gmail.com", pwd)
            server.send_message(msg)
    except Exception as e:
        print(f"Email failed (but site will update): {e}")

def generate_html(data, market_summary):
    now = datetime.now().strftime("%d/%m/%Y %H:%M")
    market_cards = "".join([f'<div class="col-md-4 mb-3"><div class="card text-center border-{m["color"]}"><div class="card-body"><h5 class="text-muted">{m["name"]}</h5><h3>{m["price"]}</h3><span class="badge bg-{m["color"]}">{m["change"]}</span></div></div></div>' for m in market_summary])
    
    sections = ""
    for cat, stocks in data.items():
        rows = "".join([f"<tr><td><strong>{s['ticker']}</strong></td><td>{s['price']}$</td><td class='text-{'success' if '+' in s['change'] else 'danger'}'>{s['change']}</td></tr>" for s in stocks])
        sections += f'<div class="card mb-4"><div class="card-header bg-dark text-white"><h5>{cat}</h5></div><table class="table mb-0"><thead><tr><th>Ticker</th><th>Price</th><th>Daily</th></tr></thead><tbody>{rows}</tbody></table></div>'

    html = f'''<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Assaf Dashboard</title><link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet"></head><body class="bg-light"><div class="container py-4 text-center"><h1>Market Dashboard</h1><p>{now}</p><div class="row">{market_cards}</div>{sections}</div></body></html>'''
    with open("index.html", "w", encoding="utf-8") as f: f.write(html)

# הרצה ראשית
res, mkt, txt = get_data()
generate_html(res, mkt) # האתר נוצר קודם כל
send_email(txt)        # המייל נשלח בסוף
