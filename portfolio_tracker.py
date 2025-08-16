import yfinance as yf
import pandas as pd
from datetime import datetime
import os
from openpyxl import load_workbook

def get_price(ticker):
    data = yf.download(ticker, period="1d", interval="1d", progress=False, auto_adjust=True)['Close']
    if not data.empty:
        return float(data.iloc[0].item())  # Fix warning by using .item()
    return None

def get_prices_and_append(tickers, filename):
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    # Fetch prices in a dict with ticker as key
    prices = {}
    for ticker in tickers:
        price = get_price(ticker)
        prices[ticker] = price

    # Convert dict to DataFrame with timestamp as index
    df_new = pd.DataFrame(prices, index=[now])

    if os.path.exists(filename):
        try:
            # Load existing file into DataFrame
            df_existing = pd.read_excel(filename, index_col=0)
            # Append new row
            df_combined = pd.concat([df_existing, df_new])
        except Exception:
            # If file corrupted or unreadable, overwrite
            df_combined = df_new
    else:
        # If file does not exist, create new
        df_combined = df_new

    # Save combined DataFrame to Excel with timestamp as index column
    df_combined.index.name = "Date"
    df_combined.to_excel(filename)

    print(f"Prices updated and saved to {filename}")

if __name__ == "__main__":
    tickers = ['SUZLON.NS', 'ETERNAL.NS', 'TATAMOTORS.NS']  # Your portfolio tickers
    filename = "portfolio.xlsx"
    get_prices_and_append(tickers, filename)

