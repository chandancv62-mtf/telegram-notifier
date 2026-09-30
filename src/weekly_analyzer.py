import os
import datetime
import pandas as pd
import yfinance as yf
from tqdm import tqdm
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

def apply_excel_styling(output_path):
    wb = openpyxl.load_workbook(output_path)
    
    # Typography & Styles
    font_family = "Segoe UI"
    
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name=font_family, size=10, bold=True, color="FFFFFF")
    
    row_fill_even = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    row_fill_odd = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
    
    # Custom Color Fills & Fonts for SL (Red) and Entry (Green)
    sl_fill = PatternFill(start_color="FCE8E6", end_color="FCE8E6", fill_type="solid")
    sl_font = Font(name=font_family, size=9, bold=True, color="C5221F")
    
    entry_fill = PatternFill(start_color="E6F4EA", end_color="E6F4EA", fill_type="solid")
    entry_font = Font(name=font_family, size=9, bold=True, color="137333")
    
    valid_status_fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
    valid_status_font = Font(name=font_family, size=9, bold=True, color="274E13")
    
    body_font = Font(name=font_family, size=9, color="000000")
    
    thin_border = Border(
        left=Side(style='thin', color='E0E0E0'),
        right=Side(style='thin', color='E0E0E0'),
        top=Side(style='thin', color='E0E0E0'),
        bottom=Side(style='thin', color='E0E0E0')
    )
    
    numeric_columns = ['PWL', 'SL', 'ENTRY', 'TARGET_1:1', 'TARGET_1:2']
    center_columns = ['Date', 'BD_Time', 'BU_Time', 'SL_Point_Time', 'Entry_Point_Time', 'Status']

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        ws.views.sheetView[0].showGridLines = True
        
        if ws.max_row <= 1 and ws.max_column <= 1 and ws.cell(row=1, column=1).value is None:
            continue
            
        header_map = {}
        
        # Style Header Row
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = thin_border
            header_map[col_idx] = str(cell.value)
            
        ws.row_dimensions[1].height = 24

        # Style Data Rows
        for row_idx in range(2, ws.max_row + 1):
            ws.row_dimensions[row_idx].height = 18
            row_fill = row_fill_odd if row_idx % 2 == 0 else row_fill_even
            
            for col_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                col_name = header_map.get(col_idx, '')
                
                cell.font = body_font
                cell.fill = row_fill
                cell.border = thin_border
                
                # Alignments & Formatting
                if col_name in numeric_columns:
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                    cell.number_format = '#,##0.00'
                elif col_name in center_columns:
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                else:
                    cell.alignment = Alignment(horizontal='left', vertical='center')
                
                # SL Columns -> RED Color Style
                if col_name in ['SL', 'SL_Point_Time']:
                    cell.fill = sl_fill
                    cell.font = sl_font

                # ENTRY Columns -> GREEN Color Style
                elif col_name in ['ENTRY', 'Entry_Point_Time']:
                    cell.fill = entry_fill
                    cell.font = entry_font

                # Status Highlight
                elif col_name == 'Status' and str(cell.value).upper() == 'VALID':
                    cell.fill = valid_status_fill
                    cell.font = valid_status_font

        # Compact Column Widths
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val = str(cell.value or '')
                max_len = max(max_len, len(val))
            ws.column_dimensions[col_letter].width = max(max_len + 2, 8)

    wb.save(output_path)


def process_exact_bd_bu_lowest_low(input_file='filtered_stocks.csv', output_file='weekly_final_trading_signals.xlsx'):
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(script_dir)
    
    # Path resolution for input file (check root, data/ dir, and script_dir)
    input_path = os.path.join(project_root, 'data', input_file)
    if not os.path.exists(input_path):
        input_path = os.path.join(project_root, input_file)
    if not os.path.exists(input_path):
        input_path = os.path.join(project_root, 'data', 'entry_sl_signals.xlsx')

    # Save output strictly inside data/ directory
    data_dir = os.path.join(project_root, 'data')
    os.makedirs(data_dir, exist_ok=True)
    output_path = os.path.join(data_dir, os.path.basename(output_file))

    today = datetime.datetime.now()
    monday_start = (today - datetime.timedelta(days=today.weekday())).replace(hour=0, minute=0, second=0, microsecond=0)
    sunday_end = monday_start + datetime.timedelta(days=6, hours=23, minutes=59)

    # Delete old file on Monday
    if today.weekday() == 0 and os.path.exists(output_path):
        try:
            os.remove(output_path)
            print("🧹 New Week Started (Monday): Previous week file deleted successfully.")
        except Exception as e:
            print(f"⚠️ Failed to remove old file: {e}")

    try:
        if input_path.endswith('.csv'):
            df_signals = pd.read_csv(input_path)
        else:
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
            pwl = float(symbol_row['Prev_Week_Low']) if 'Prev_Week_Low' in symbol_row else float(symbol_row['PWL'])

            ticker = yf.Ticker(ticker_symbol)
            df_15m = ticker.history(period="5d", interval="15m")

            if df_15m.empty:
                continue

            if df_15m.index.tz is not None:
                df_15m.index = df_15m.index.tz_localize(None)

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

    # Format Excel output
    apply_excel_styling(output_path)

    print(f"\n✅ Fresh Compact Week File Generated in Data Folder with Color Coding: '{output_path}'")

if __name__ == "__main__":
    process_exact_bd_bu_lowest_low()
