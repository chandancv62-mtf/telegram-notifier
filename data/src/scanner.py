import os
import pandas as pd
import yfinance as yf
from tqdm import tqdm
import config

def run_scanner():
    script_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    input_path = os.path.join(script_dir, config.STOCKS_CSV_PATH)
    output_path = os.path.join(script_dir, config.FILTERED_STOCKS_PATH)

    try:
        df_stocks = pd.read_csv(input_path)
        symbols = df_stocks[df_stocks.columns[0]].dropna().tolist()
        print(f"Loaded {len(symbols)} stocks from '{input_path}'.\n")
    except Exception as e:
        print(f"Error reading {input_path}: {e}")
        return

    filtered_results = []

    for symbol in tqdm(symbols, desc="Scanning Stocks", unit="stock"):
        try:
            ticker_symbol = str(symbol).strip()
            if not (ticker_symbol.endswith('.NS') or ticker_symbol.endswith('.BO')):
                ticker_symbol += '.NS'

            ticker = yf.Ticker(ticker_symbol)
            df_daily = ticker.history(period="1mo", interval="1d")

            if df_daily.empty or len(df_daily) < 10:
                continue

            df_weekly = df_daily.resample('W-FRI').agg({
                'Open': 'first', 'High': 'max', 'Low': 'min', 'Close': 'last'
            }).dropna()

            if len(df_weekly) < 2:
                continue

            prev_week_low = df_weekly.iloc[-2]['Low']
            latest_daily = df_daily.iloc[-1]
            latest_date = df_daily.index[-1].strftime('%Y-%m-%d')
            daily_low = latest_daily['Low']
            daily_close = latest_daily['Close']

            if daily_low < prev_week_low and daily_close > prev_week_low:
                filtered_results.append({
                    'Ticker': symbol,
                    'Date': latest_date,
                    'Daily_Low': round(daily_low, 2),
                    'Daily_Close': round(daily_close, 2),
                    'Prev_Week_Low': round(prev_week_low, 2)
                })

        except Exception as e:
            print(f"Error scanning {symbol}: {e}")

    df_results = pd.DataFrame(filtered_results)
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    df_results.to_csv(output_path, index=False)
    print(f"\nScan completed! {len(df_results)} stock(s) matched.")

if __name__ == "__main__":
    run_scanner()
