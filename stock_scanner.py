import yfinance as yf
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
import os

# הגדרות רשימת מעקב
WATCHLIST = {
    "AI & Tech": ["PLTR", "SOFI", "AMD", "NVDA", "MSFT"],
    "Crypto & Growth": ["MSTR", "MARA", "COIN", "TSLA"],
    "Portfolio": ["BMNR", "SEDG", "BTG"]
}

MARKET_INDICES = {"S&P 500": "^GSPC", "NASDAQ": "^IXIC", "Bitcoin": "BTC-USD"}

def get_data():
    results = {}
    market_summary = []
    
    # נתוני מדדים
    for name, ticker in MARKET_INDICES.items():
        ticker_obj = yf.Ticker(ticker)
        hist = ticker_obj.history(period="2d")
        if len(hist) >= 2:
            price = hist['Close'].iloc[-1]
            change = ((price - hist['Close'].iloc[-2]) / hist['Close'].iloc[-2]) * 100
            market_summary.append({"name": name, "price": f"{price:,.2f}", "change": f"{change:+.2f}%", "color": "success" if change >= 0 else "danger"})

    # נתוני מניות
    all_stocks_for_email = ""
    for category, tickers in WATCHLIST.items():
        results[category] = []
        all_stocks_for_email += f"\n-- {category} --\n"
        for t in tickers:
            s = yf.Ticker(t).history(period="2d")
            if len(s) >= 2:
                curr = s['Close'].iloc[-1]
                chg = ((curr - s['Close'].iloc[-2]) / s['Close'].iloc[-2]) * 100
                score = "High" if chg > 2 else "Neutral" # לוגיקת ציון פשוטה
                results[category].append({"ticker": t, "price": f"{curr:.2f}", "change": f"{chg:+.2f}%", "score": score})
                all_stocks_for_email += f"{t}: {curr:.2f}$ ({chg:+.2f}%)\n"
    
    return results, market_summary, all_stocks_for_email

def send_email(content):
    password = os.getenv("APP_PASSWORD")
    if not password: return
    
    msg = MIMEMultipart()
    msg['From'] = "Stock Scanner"
    msg['To'] = "ashtivi@gmail.com" # המייל שלך
    msg['Subject'] = f"Stock Report {datetime.now().strftime('%d/%m/%Y')}"
    msg.attach(MIMEText(content, 'plain'))
    
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(msg['To'], password)
        server.send_message(msg)

def generate_html(stocks_data, market_summary):
    now = datetime.now().strftime("%d/%m/%Y %H:%M")
    market_cards = "".join([f'<div class="col-md-4 mb-3"><div class="card text-center border-{m["color"]}"><div class="card-body"><h5 class="text-muted">{m["name"]}</h5><h3>{m["price"]}</h3><span class="badge bg-{m["color"]}">{m["change"]}</span></div></div></div>' for m in market_summary])
    
    sections = ""
    for cat, stocks in stocks_data.items():
        rows = "".join([f"<tr><td><strong>{s['ticker']}</strong></td><td>{s['price']}$</td><td class='text-{'success' if '+' in s['change'] else 'danger'}'>{s['change']}</td><td><span class='badge bg-primary'>{s['score']}</span></td></tr>" for s in stocks])
        sections += f'<div class="card mb-4"><div class="card-header bg-dark text-white"><h5>{cat}</h5></div><table class="table table-hover mb-0"><thead><tr><th>Ticker</th><th>Price</th><th>Daily</th><th>Score</th></tr></thead><tbody>{rows}</tbody></table></div>'

    html = f'''<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Dashboard</title><link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet"></head><body class="bg-light"><div class="container py-4 text-center"><h1>Market Dashboard</h1><p>{now}</p><div class="row">{market_cards}</div>{sections}</div></body></html>'''
    with open("index.html", "w", encoding="utf-8") as f: f.write(html)

# הרצה
data, market, email_text = get_data()
generate_html(data, market)
send_email(email_text)
print("Done! Email sent and HTML generated.")
