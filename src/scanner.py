import os
import pandas as pd
import yfinance as yf
from tqdm import tqdm

def run_scanner(input_file='data/stocks.csv', output_file='data/filtered_stocks.csv'):
    # Get current script directory (src/)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Get project root directory (one level up from src/)
    project_root = os.path.dirname(script_dir)
    
    # Construct full absolute paths relative to project root
    input_path = os.path.join(project_root, input_file)
    output_path = os.path.join(project_root, output_file)

    # 1. Read stocks list from CSV
    try:
        df_stocks = pd.read_csv(input_path)
        col_name = df_stocks.columns[0]
        symbols = df_stocks[col_name].dropna().tolist()
        print(f"Loaded {len(symbols)} stocks from '{input_path}'.\n")
    except Exception as e:
        print(f"Error reading {input_path}: {e}")
        return

    filtered_results = []

    # 2. Loop through each stock with progress bar (tqdm)
    for symbol in tqdm(symbols, desc="Scanning Stocks", unit="stock"):
        try:
            ticker_symbol = str(symbol).strip()
            if not (ticker_symbol.endswith('.NS') or ticker_symbol.endswith('.BO')):
                ticker_symbol += '.NS'

            ticker = yf.Ticker(ticker_symbol)
            df_daily = ticker.history(period="1mo", interval="1d")

            if df_daily.empty or len(df_daily) < 10:
                continue

            # Resample daily data to weekly
            df_weekly = df_daily.resample('W-FRI').agg({
                'Open': 'first',
                'High': 'max',
                'Low': 'min',
                'Close': 'last'
            }).dropna()

            if len(df_weekly) < 2:
                continue

            # Previous week low
            prev_week_low = df_weekly.iloc[-2]['Low']

            # Latest daily candle
            latest_daily = df_daily.iloc[-1]
            latest_date = df_daily.index[-1].strftime('%Y-%m-%d')
            daily_low = latest_daily['Low']
            daily_close = latest_daily['Close']

            # Condition: Low broke Previous Week's Low, but Close is ABOVE it
            if daily_low < prev_week_low and daily_close > prev_week_low:
                filtered_results.append({
                    'Ticker': symbol,
                    'Date': latest_date,
                    'Daily_Low': round(daily_low, 2),
                    'Daily_Close': round(daily_close, 2),
                    'Prev_Week_Low': round(prev_week_low, 2)
                })
                tqdm.write(f" -> [MATCH] {symbol} | Low: {round(daily_low, 2)} < PWL: {round(prev_week_low, 2)} | Close: {round(daily_close, 2)} > PWL")

        except Exception as e:
            tqdm.write(f"Error scanning {symbol}: {e}")

    # 3. Save results to CSV file with append, deduplication, and weekly cleanup
    if filtered_results:
        df_new = pd.DataFrame(filtered_results)

        # File exist karti hai toh load karein
        if os.path.exists(output_path):
            try:
                df_existing = pd.read_csv(output_path)
                df_combined = pd.concat([df_existing, df_new], ignore_index=True)
            except Exception as e:
                tqdm.write(f"Error reading existing file, creating new one: {e}")
                df_combined = df_new
        else:
            df_combined = df_new

        # Duplicate entries remove karein
        df_combined = df_combined.drop_duplicates()

        # Retain data only for the current week (Monday to Sunday)
        df_combined['Date_dt'] = pd.to_datetime(df_combined['Date'])
        
        today = pd.Timestamp.now().normalize()
        start_of_week = today - pd.Timedelta(days=today.weekday()) # Current Monday
        end_of_week = start_of_week + pd.Timedelta(days=6)          # Current Sunday

        df_weekly_only = df_combined[
            (df_combined['Date_dt'] >= start_of_week) & 
            (df_combined['Date_dt'] <= end_of_week)
        ].copy()

        df_weekly_only = df_weekly_only.drop(columns=['Date_dt'])

        # Updated CSV save karein
        df_weekly_only.to_csv(output_path, index=False)
        print(f"\nScan completed! {len(df_weekly_only)} stock(s) from current week saved in '{output_path}'.")
    else:
        print("\nScan completed! No new stocks met the condition today.")

if __name__ == "__main__":
    run_scanner()
