import yfinance as yf
import pandas as pd
import os
from datetime import datetime
from openpyxl import load_workbook

def append_to_excel(tickers, filename):
    # Get latest prices
    today = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    prices = {}

    for ticker in tickers:
        stock = yf.Ticker(ticker)
        info = stock.history(period="1d")
        if not info.empty:
            last_price = info["Close"].iloc[-1]
            prices[ticker] = last_price

    # Create DataFrame: one row per day
    df = pd.DataFrame([{ "Date": today, **prices }])

    if os.path.exists(filename):
        # Append new row at bottom
        with pd.ExcelWriter(filename, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
            book = load_workbook(filename)
            sheet = book.active
            start_row = sheet.max_row
            df.to_excel(writer, sheet_name="Sheet1", startrow=start_row, header=False, index=False)
    else:
        # Create new file
        with pd.ExcelWriter(filename, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Sheet1", index=False)

if __name__ == "__main__":
    tickers = ["SUZLON.NS", "TATAMOTORS.NS", "ETERNAL.NS"]
    filename = "portfolio-update.xlsx"
    append_to_excel(tickers, filename)
