import os
import pandas as pd
import yfinance as yf
from tqdm import tqdm
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
import config

def process_exact_bd_bu_lowest_low():
    script_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    input_path = os.path.join(script_dir, config.ENTRY_SL_SIGNALS_PATH)
    output_path = os.path.join(script_dir, config.WEEKLY_SIGNALS_PATH)

    try:
        df_signals = pd.read_excel(input_path)
        if df_signals.empty:
            print("❌ Input file empty hai.")
            return
        symbols = df_signals['Ticker'].dropna().unique().tolist()
    except Exception as e:
        print(f"❌ Error reading file: {e}")
        return

    all_master_results = []
    valid_results = []
    invalid_results = []

    for symbol in tqdm(symbols, desc="Processing Weekly Setups", unit="stock"):
        try:
            stock_name = str(symbol).strip()
            ticker_symbol = stock_name if (stock_name.endswith('.NS') or stock_name.endswith('.BO')) else stock_name + '.NS'

            symbol_row = df_signals[df_signals['Ticker'] == symbol].iloc[-1]
            pwl = float(symbol_row['PWL'])

            ticker = yf.Ticker(ticker_symbol)
            df_15m = ticker.history(period="5d", interval="15m")

            if df_15m.empty:
                res = {'Ticker': stock_name, 'PWL': pwl, 'Status': 'INVALID', 'Reason': 'No Intraday Data'}
                invalid_results.append(res)
                all_master_results.append(res)
                continue

            if df_15m.index.tz is not None:
                df_15m.index = df_15m.index.tz_localize(None)

            df_15m = df_15m.between_time(config.MARKET_START_TIME, config.MARKET_END_TIME)

            bd_candles = df_15m[df_15m['Low'] < pwl]
            if bd_candles.empty:
                res = {'Ticker': stock_name, 'PWL': pwl, 'Status': 'INVALID', 'Reason': 'No BD candle below PWL'}
                invalid_results.append(res)
                all_master_results.append(res)
                continue

            first_bd_candle = bd_candles.iloc[0]

            bu_candidates = df_15m[
                (df_15m['Low'] < pwl) & 
                (df_15m['Close'] > pwl) & 
                (df_15m.index >= first_bd_candle.name)
            ]

            if bu_candidates.empty:
                res = {'Ticker': stock_name, 'PWL': pwl, 'Status': 'INVALID', 'Reason': 'BD Candle found, but no BU candle close > PWL'}
                invalid_results.append(res)
                all_master_results.append(res)
                continue

            last_bu_candle = bu_candidates.iloc[-1]
            bd_bu_range = df_15m.loc[first_bd_candle.name:last_bu_candle.name]
            sl_candle = bd_bu_range.loc[bd_bu_range['Low'].idxmin()]
            lowest_sl_point = round(float(sl_candle['Low']), 2)
            lowest_sl_time = sl_candle.name.strftime('%Y-%m-%d %H:%M')

            sl_idx = df_15m.index.get_loc(sl_candle.name)
            swing_high_found = False
            entry_point = None
            entry_point_time = None

            for i in range(sl_idx - 1, 0, -1):
                if i - 1 < 0 or i + 1 >= len(df_15m):
                    continue
                left_high = df_15m.iloc[i - 1]['High']
                middle_high = df_15m.iloc[i]['High']
                right_high = df_15m.iloc[i + 1]['High']

                if (middle_high > pwl) and (middle_high > left_high) and (middle_high > right_high):
                    swing_high_found = True
                    entry_candle = df_15m.iloc[i]
                    entry_point = round(float(entry_candle['High']), 2)
                    entry_point_time = entry_candle.name.strftime('%Y-%m-%d %H:%M')
                    break

            if not swing_high_found:
                df_before_sl = df_15m.iloc[:sl_idx]
                above_pwl = df_before_sl[df_before_sl['High'] > pwl]
                if not above_pwl.empty:
                    last_above_pwl = above_pwl.iloc[-1]
                    entry_point = round(float(last_above_pwl['High']), 2)
                    entry_point_time = last_above_pwl.name.strftime('%Y-%m-%d %H:%M')
                else:
                    entry_point = round(float(last_bu_candle['High']), 2)
                    entry_point_time = last_bu_candle.name.strftime('%Y-%m-%d %H:%M')

            risk = round(entry_point - lowest_sl_point, 2)
            sl_pct = round((risk / entry_point) * 100, 2)
            sl_pct_with_points = f"{sl_pct}% ({risk})"

            target_1_1 = round(entry_point + (config.RISK_REWARD_RATIOS[0] * risk), 2)
            target_1_2 = round(entry_point + (config.RISK_REWARD_RATIOS[1] * risk), 2)

            res = {
                'Ticker': stock_name,
                'Date': sl_candle.name.strftime('%Y-%m-%d'),
                'PWL': round(pwl, 2),
                'BD_Time': first_bd_candle.name.strftime('%Y-%m-%d %H:%M'),
                'BU_Time': last_bu_candle.name.strftime('%Y-%m-%d %H:%M'),
                'SL': lowest_sl_point,
                'SL_Point_Time': lowest_sl_time,
                'ENTRY': entry_point,
                'Entry_Point_Time': entry_point_time,
                'SL_%(Point)': sl_pct_with_points,
                'TARGET_1:1': target_1_1,
                'TARGET_1:2': target_1_2,
                'Status': 'VALID',
                'Reason': 'Setup Matched'
            }
            valid_results.append(res)
            all_master_results.append(res)

        except Exception as e:
            res = {'Ticker': stock_name, 'PWL': pwl if 'pwl' in locals() else None, 'Status': 'ERROR', 'Reason': str(e)}
            invalid_results.append(res)
            all_master_results.append(res)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        pd.DataFrame(all_master_results).to_excel(writer, sheet_name='All_Stocks_Master', index=False)
        pd.DataFrame(valid_results).to_excel(writer, sheet_name='Valid_Setups_Only', index=False)
        pd.DataFrame(invalid_results).to_excel(writer, sheet_name='Invalid_Setups_Only', index=False)

    print(f"\n✅ Created Excel with Targets & SL saved to '{output_path}'.")

if __name__ == "__main__":
    process_exact_bd_bu_lowest_low()
