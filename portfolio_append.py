import yfinance as yf
import pandas as pd
from openpyxl import load_workbook
import os

def append_to_excel(tickers, filename):
    # Get latest stock data
    data = []
    for ticker in tickers:
        stock = yf.Ticker(ticker)
        info = stock.history(period="1d")
        if not info.empty:
            last_price = info["Close"].iloc[-1]
            data.append({"Ticker": ticker, "Price": last_price})
    df = pd.DataFrame(data)

    if os.path.exists(filename):
        # Load workbook
        book = load_workbook(filename)
        writer = pd.ExcelWriter(filename, engine="openpyxl", mode="a", if_sheet_exists="overlay")
        writer.book = book
        if "Sheet1" in book.sheetnames:
            sheet = book["Sheet1"]
            start_row = sheet.max_row
        else:
            start_row = 0
        df.to_excel(writer, sheet_name="Sheet1", startrow=start_row, header=False, index=False)
        writer.close()
    else:
        # Create new file
        with pd.ExcelWriter(filename, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Sheet1", index=False)

if __name__ == "__main__":
    tickers = ["SUZLON.NS", "TATAMOTORS.NS", "ETERNAL.NS"]  # NSE tickers
    filename = "portfolio-update.xlsx"
    append_to_excel(tickers, filename)
