import os
import pandas as pd
import yfinance as yf
from tqdm import tqdm
import config

def track_trade_status():
    script_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    input_path = os.path.join(script_dir, config.WEEKLY_SIGNALS_PATH)
    output_path = os.path.join(script_dir, config.TRADE_TRACKER_PATH)

    try:
        df_valid = pd.read_excel(input_path, sheet_name='Valid_Setups_Only')
        if df_valid.empty:
            print("Valid setups sheet empty hai.")
            return
    except Exception as e:
        print(f"Error reading file '{input_path}': {e}")
        return

    tracked_results = []
    print(f"\nTracking Trade Status for {len(df_valid)} Valid Stocks...\n")

    for idx, row in tqdm(df_valid.iterrows(), total=len(df_valid), desc="Tracking Trades", unit="stock"):
        stock_name = str(row['Ticker']).strip()
        ticker_symbol = stock_name if (stock_name.endswith('.NS') or stock_name.endswith('.BO')) else stock_name + '.NS'

        sl_price = float(row['SL'])
        entry_price = float(row['ENTRY'])
        bu_time_str = str(row['BU_Time'])

        risk = round(entry_price - sl_price, 2)
        sl_pct = round((risk / entry_price) * 100, 2)
        sl_pct_with_points = f"{sl_pct}% ({risk})"

        target_1_1 = float(row['TARGET_1:1']) if 'TARGET_1:1' in row else round(entry_price + (1 * risk), 2)
        target_1_2 = float(row['TARGET_1:2']) if 'TARGET_1:2' in row else round(entry_price + (2 * risk), 2)

        trade_status = "WAIT (No Entry)"
        entry_triggered_time = "N/A"
        exit_time = "N/A"
        exit_price = None
        pnl_pct = 0.0

        try:
            ticker = yf.Ticker(ticker_symbol)
            df_15m = ticker.history(period="5d", interval="15m")

            if not df_15m.empty:
                if df_15m.index.tz is not None:
                    df_15m.index = df_15m.index.tz_localize(None)

                df_15m = df_15m.between_time(config.MARKET_START_TIME, config.MARKET_END_TIME)
                bu_dt = pd.to_datetime(bu_time_str)
                week_start = bu_dt - pd.Timedelta(days=bu_dt.weekday())
                current_week_end = week_start + pd.Timedelta(days=4, hours=15, minutes=30)

                df_after_bu = df_15m[(df_15m.index > bu_dt) & (df_15m.index <= current_week_end)]
                entry_candles = df_after_bu[df_after_bu['High'] >= entry_price]

                if not entry_candles.empty:
                    first_entry_candle = entry_candles.iloc[0]
                    entry_triggered_time = first_entry_candle.name.strftime('%Y-%m-%d %H:%M')
                    trade_status = "ACTIVE"

                    df_after_entry = df_after_bu[df_after_bu.index >= first_entry_candle.name]
                    target_1_1_hit = False
                    last_candle_dt = None
                    last_candle_close = None

                    for dt, c_row in df_after_entry.iterrows():
                        candle_high = float(c_row['High'])
                        candle_low = float(c_row['Low'])
                        last_candle_dt = dt
                        last_candle_close = float(c_row['Close'])

                        if candle_low <= sl_price:
                            trade_status = "TARGET 1:1 HIT (Then SL)" if target_1_1_hit else "SL HIT"
                            exit_time = dt.strftime('%Y-%m-%d %H:%M')
                            exit_price = round(sl_price, 2)
                            break

                        if candle_high >= target_1_1 and not target_1_1_hit:
                            target_1_1_hit = True
                            trade_status = "TARGET 1:1 HIT (Active for 1:2)"
                            exit_time = dt.strftime('%Y-%m-%d %H:%M')
                            exit_price = round(target_1_1, 2)

                        if candle_high >= target_1_2:
                            trade_status = "TARGET 1:2 HIT"
                            exit_time = dt.strftime('%Y-%m-%d %H:%M')
                            exit_price = round(target_1_2, 2)
                            break

                    if exit_price is None and last_candle_dt is not None:
                        exit_time = last_candle_dt.strftime('%Y-%m-%d %H:%M') + " (Week End)"
                        exit_price = round(last_candle_close, 2)

                    if exit_price is not None:
                        pnl_pct = round(((exit_price - entry_price) / entry_price) * 100, 2)

                else:
                    trade_status = "ENTRY EXPIRED (Week End)"

        except Exception as e:
            trade_status = f"ERROR: {str(e)}"

        res = {
            'Ticker': stock_name,
            'PWL': row['PWL'],
            'SL': sl_price,
            'ENTRY': entry_price,
            'SL_%(Point)': sl_pct_with_points,
            'TARGET_1:1': target_1_1,
            'TARGET_1:2': target_1_2,
            'Entry_Triggered_Time': entry_triggered_time,
            'Trade_Status': trade_status,
            'Exit_Time': exit_time,
            'Exit_Price': exit_price,
            'PnL_%': pnl_pct
        }
        tracked_results.append(res)

    df_current = pd.DataFrame(tracked_results)
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_current.to_excel(writer, sheet_name='current_week', index=False)

    print(f"\nSingle-Execution Tracking Done saved to '{output_path}'.")

if __name__ == "__main__":
    track_trade_status()
