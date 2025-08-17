import yfinance as yf
import pandas as pd
import os
from datetime import datetime
from openpyxl import load_workbook

def append_to_excel(tickers, filename):
    today = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    prices = {}

    for ticker in tickers:
        stock = yf.Ticker(ticker)
        info = stock.history(period="1d")
        if not info.empty:
            last_price = info["Close"].iloc[-1]
            prices[ticker] = last_price

    # Create row as DataFrame
    df = pd.DataFrame([{ "Date": today, **prices }])

    if os.path.exists(filename):
        book = load_workbook(filename)
        sheet = book.active
        start_row = sheet.max_row + 1   # append after last row

        with pd.ExcelWriter(filename, engine="openpyxl", mode="a") as writer:
            writer.book = book
            writer.sheets = {ws.title: ws for ws in book.worksheets}
            df.to_excel(writer, sheet_name="Sheet1", startrow=start_row-1, header=False, index=False)
    else:
        with pd.ExcelWriter(filename, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Sheet1", index=False)

if __name__ == "__main__":
    tickers = ["SUZLON.NS", "TATAMOTORS.NS", "ETERNAL.NS"]
    filename = "portfolio-update.xlsx"
    append_to_excel(tickers, filename)
