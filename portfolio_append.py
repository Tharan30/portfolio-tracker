import yfinance as yf
import pandas as pd
from openpyxl import load_workbook, Workbook
import os
from datetime import datetime

def append_to_excel(tickers, filename):
    today = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    prices = {}

    for ticker in tickers:
        stock = yf.Ticker(ticker)
        info = stock.history(period="1d")
        if not info.empty:
            last_price = info["Close"].iloc[-1]
            prices[ticker] = last_price

    # Create one row with Date + tickers
    df = pd.DataFrame([{ "Date": today, **prices }])

    if os.path.exists(filename):
        book = load_workbook(filename)
        sheet = book.active
        start_row = sheet.max_row + 1

        for r in df.itertuples(index=False, name=None):
            sheet.append(r)   # ✅ appends row at bottom
        book.save(filename)
    else:
        # Create new workbook
        with pd.ExcelWriter(filename, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Sheet1", index=False)

if __name__ == "__main__":
    tickers = ["SUZLON.NS", "TATAMOTORS.NS", "ETERNAL.NS"]
    filename = "portfolio-changes.xlsx"
    append_to_excel(tickers, filename)

