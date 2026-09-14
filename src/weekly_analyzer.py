import os
import datetime
import pandas as pd
import yfinance as yf
from tqdm import tqdm
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

def process_exact_bd_bu_lowest_low(input_file='entry_sl_signals.xlsx', output_file='weekly_final_trading_signals.xlsx'):
    script_dir = os.path.dirname(os.path.abspath(__file__))
    input_path = os.path.join(script_dir, input_file)
    output_path = os.path.join(script_dir, output_file)

    today = datetime.datetime.now()
    monday_start = (today - datetime.timedelta(days=today.weekday())).replace(hour=0, minute=0, second=0, microsecond=0)
    sunday_end = monday_start + datetime.timedelta(days=6, hours=23, minutes=59)

    # RULE: नया हफ़्ता (Monday) शुरू होते ही पुरानी फ़ाइल को पूरी तरह हटा दें
    if today.weekday() == 0 and os.path.exists(output_path):
        try:
            os.remove(output_path)
            print("🧹 New Week Started (Monday): Previous week file deleted successfully.")
        except Exception as e:
            print(f"⚠️ Failed to remove old file: {e}")

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

    print(f"\n🚀 Scanning Fresh Setups for Current Week ({monday_start.strftime('%Y-%m-%d')} to {sunday_end.strftime('%Y-%m-%d')})...\n")

    for symbol in tqdm(symbols, desc="Processing Stocks", unit="stock"):
        try:
            stock_name = str(symbol).strip()
            ticker_symbol = stock_name if (stock_name.endswith('.NS') or stock_name.endswith('.BO')) else stock_name + '.NS'

            symbol_row = df_signals[df_signals['Ticker'] == symbol].iloc[-1]
            pwl = float(symbol_row['PWL'])

            ticker = yf.Ticker(ticker_symbol)
            # Fetch 5 days intraday data
            df_15m = ticker.history(period="5d", interval="15m")

            if df_15m.empty:
                continue

            if df_15m.index.tz is not None:
                df_15m.index = df_15m.index.tz_localize(None)

            # Strictly Filter CURRENT WEEK Intraday Data Only (Mon-Sun)
            df_15m = df_15m[(df_15m.index >= monday_start) & (df_15m.index <= sunday_end)]
            df_15m = df_15m.between_time('09:15', '15:15')

            if df_15m.empty:
                continue

            # Step 1: BD Candle
            bd_candles = df_15m[df_15m['Low'] < pwl]
            if bd_candles.empty:
                continue

            first_bd_candle = bd_candles.iloc[0]

            # Step 2: BU Candle
            bu_candidates = df_15m[
                (df_15m['Low'] < pwl) & 
                (df_15m['Close'] > pwl) & 
                (df_15m.index >= first_bd_candle.name)
            ]

            if bu_candidates.empty:
                continue

            last_bu_candle = bu_candidates.iloc[-1]

            # Step 3: SL Calculation
            bd_bu_range = df_15m.loc[first_bd_candle.name:last_bu_candle.name]
            sl_candle = bd_bu_range.loc[bd_bu_range['Low'].idxmin()]
            lowest_sl_point = round(float(sl_candle['Low']), 2)
            lowest_sl_time = sl_candle.name.strftime('%Y-%m-%d %H:%M')

            # Step 4: Entry Calculation
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

            target_1_1 = round(entry_point + (1 * risk), 2)
            target_1_2 = round(entry_point + (2 * risk), 2)

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
            continue

    with pd.ExcelWriter(output_path, engine='openpyxl', mode='w') as writer:
        pd.DataFrame(all_master_results).to_excel(writer, sheet_name='All_Stocks_Master', index=False)
        pd.DataFrame(valid_results).to_excel(writer, sheet_name='Valid_Setups_Only', index=False)
        pd.DataFrame(invalid_results).to_excel(writer, sheet_name='Invalid_Setups_Only', index=False)

    print(f"\n✅ Fresh Current Week File Generated: '{output_file}'")

if __name__ == "__main__":
    process_exact_bd_bu_lowest_low()
