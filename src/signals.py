import os
import pandas as pd
import yfinance as yf
from tqdm import tqdm
import config

def find_entry_sl():
    script_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    input_path = os.path.join(script_dir, config.FILTERED_STOCKS_PATH)
    output_path = os.path.join(script_dir, config.ENTRY_SL_SIGNALS_PATH)

    try:
        df_filtered = pd.read_csv(input_path)
        if df_filtered.empty:
            print("Filtered stocks CSV is empty.")
            return
        symbols = df_filtered['Ticker'].dropna().tolist()
    except Exception as e:
        print(f"Error reading {input_path}: {e}")
        return

    results = []

    for symbol in tqdm(symbols, desc="Processing Signals", unit="stock"):
        try:
            ticker_symbol = str(symbol).strip()
            if not (ticker_symbol.endswith('.NS') or ticker_symbol.endswith('.BO')):
                ticker_symbol += '.NS'

            ticker = yf.Ticker(ticker_symbol)
            df_daily = ticker.history(period="1mo", interval="1d")
            df_weekly = df_daily.resample('W-FRI').agg({
                'Open': 'first', 'High': 'max', 'Low': 'min', 'Close': 'last'
            }).dropna()

            if len(df_weekly) < 2:
                continue

            pwl = df_weekly.iloc[-2]['Low']
            df_15m = ticker.history(period="5d", interval="15m")
            if df_15m.empty:
                continue

            if df_15m.index.tz is not None:
                df_15m.index = df_15m.index.tz_localize(None)

            df_15m = df_15m.between_time(config.MARKET_START_TIME, config.MARKET_END_TIME)

            bd_candles = df_15m[df_15m['Low'] < pwl]
            if bd_candles.empty:
                continue

            bd_candle = bd_candles.iloc[0]
            bu_candidates = df_15m[
                (df_15m['Low'] < pwl) & 
                (df_15m['Close'] > pwl) & 
                (df_15m.index >= bd_candle.name)
            ]

            if bu_candidates.empty:
                continue

            bu_candle = bu_candidates.iloc[-1]
            sub_df = df_15m.loc[bd_candle.name:bu_candle.name]
            sl_price = sub_df['Low'].min()
            entry_price = bu_candle['Close']

            results.append({
                'Ticker': symbol,
                'Date': bu_candle.name.strftime('%Y-%m-%d'),
                'PWL': round(pwl, 2),
                'BD_Time': bd_candle.name.strftime('%Y-%m-%d %H:%M'),
                'BD_Low': round(bd_candle['Low'], 2),
                'BU_Time': bu_candle.name.strftime('%Y-%m-%d %H:%M'),
                'BU_Close': round(bu_candle['Close'], 2),
                'Entry_Price': round(entry_price, 2),
                'Stop_Loss': round(sl_price, 2)
            })

        except Exception as e:
            print(f"Error analyzing {symbol}: {e}")

    df_results = pd.DataFrame(results)
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    if not df_results.empty:
        df_results.to_excel(output_path, index=False)
        print(f"\nSaved {len(df_results)} signals to '{output_path}'.")

if __name__ == "__main__":
    find_entry_sl()
