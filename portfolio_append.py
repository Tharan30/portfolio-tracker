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
        # Open the existing workbook
        with pd.ExcelWriter(filename, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
            start_row = writer.sheets["Sheet1"].max_row  # find last row
            df.to_excel(writer, sheet_name="Sheet1", startrow=start_row, header=False, index=False)
    else:
        # Create a new file
        with pd.ExcelWriter(filename, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Sheet1", index=False)

if __name__ == "__main__":
    tickers = ["SUZLON.NS", "TATAMOTORS.NS", "ETERNAL.NS"  # 🔧 your stock list
    filename = "portfolio-update.xlsx"
    append_to_excel(tickers, filename)

