import yfinance as yf
import pandas as pd
from datetime import datetime
import os

# רשימת המניות לפי הקטגוריות שלך
WATCHLIST = {
    "AI & Tech": ["PLTR", "SOFI", "AMD", "NVDA", "MSFT"],
    "Crypto & Growth": ["MSTR", "MARA", "COIN", "TSLA"],
    "Portfolio": ["BMNR", "SEDG", "BTG"]
}

# מדדים וקריפטו לראש הדף
MARKET_INDICES = {
    "S&P 500": "^GSPC",
    "NASDAQ": "^IXIC",
    "Bitcoin": "BTC-USD"
}

def get_market_data():
    summary = []
    for name, ticker in MARKET_INDICES.items():
        data = yf.Ticker(ticker).history(period="2d")
        if len(data) >= 2:
            current = data['Close'].iloc[-1]
            prev = data['Close'].iloc[-2]
            change = ((current - prev) / prev) * 100
            summary.append({"name": name, "price": f"{current:,.2f}", "change": f"{change:+.2f}%", "color": "success" if change >= 0 else "danger"})
    return summary

def generate_html(stocks_data, market_summary):
    now = datetime.now().strftime("%d/%m/%Y %H:%M")
    
    market_html = "".join([f'''
        <div class="col-md-4 mb-3">
            <div class="card text-center border-{m['color']}">
                <div class="card-body">
                    <h5 class="card-title text-muted">{m['name']}</h5>
                    <h3 class="card-text text-{m['color']}">{m['price']}</h3>
                    <small class="badge bg-{m['color']}">{m['change']}</small>
                </div>
            </div>
        </div>
    ''' for m in market_summary])

    sections_html = ""
    for category, stocks in stocks_data.items():
        rows = "".join([f'''
            <tr>
                <td><strong>{s['ticker']}</strong></td>
                <td>{s['price']}$</td>
                <td class="text-{'success' if float(s['change'].strip('%')) >= 0 else 'danger'}">{s['change']}</td>
                <td><span class="badge bg-primary">{s['score']}</span></td>
            </tr>
        ''' for s in stocks])
        
        sections_html += f'''
            <div class="card mb-4 shadow-sm">
                <div class="card-header bg-dark text-white"><h5>{category}</h5></div>
                <div class="table-responsive">
                    <table class="table table-hover mb-0">
                        <thead class="table-light"><tr><th>סימול</th><th>מחיר</th><th>שינוי יומי</th><th>ציון</th></tr></thead>
                        <tbody>{rows}</tbody>
                    </table>
                </div>
            </div>
        '''

    html_template = f'''
    <!DOCTYPE html>
    <html lang="he" dir="rtl">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>סורק המניות של אסף</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.rtl.min.css" rel="stylesheet">
        <style>
            body {{ background-color: #f8f9fa; font-family: system-ui; }}
            .card {{ border-radius: 12px; }}
        </style>
    </head>
    <body>
        <div class="container py-4">
            <header class="text-center mb-5">
                <h1 class="display-5 fw-bold text-dark">דשבורד מניות וקריפטו</h1>
                <p class="text-muted">עדכון אחרון: {now}</p>
            </header>
            
            <div class="row mb-4">{market_html}</div>
            {sections_html}
        </div>
    </body>
    </html>
    '''
    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html_template)

# כאן יבוא הלוגיקה של הסורק שלך שמחשבת ציונים (נניח שזה מוכן)
# בסוף הריצה נקרא ל:
# market_data = get_market_data()
# generate_html(final_results, market_data)

print("HTML dashboard generated successfully!")
