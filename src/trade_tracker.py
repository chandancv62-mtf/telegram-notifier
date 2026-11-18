import os
import datetime
import pandas as pd
import yfinance as yf
from tqdm import tqdm
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

def get_week_range_str(date_obj):
    """Calculate Monday to Sunday week range string (dd-mm-yyyy to dd-mm-yyyy)"""
    if pd.isna(date_obj) or str(date_obj).strip().upper() in ['N/A', 'NONE', '']:
        today = datetime.datetime.now()
        monday = today - datetime.timedelta(days=today.weekday())
        sunday = monday + datetime.timedelta(days=6)
        return f"({monday.strftime('%d-%m-%Y')} to {sunday.strftime('%d-%m-%Y')})"
    
    if isinstance(date_obj, str):
        try:
            date_obj = pd.to_datetime(date_obj.split(' ')[0])
        except Exception:
            today = datetime.datetime.now()
            monday = today - datetime.timedelta(days=today.weekday())
            sunday = monday + datetime.timedelta(days=6)
            return f"({monday.strftime('%d-%m-%Y')} to {sunday.strftime('%d-%m-%Y')})"
            
    monday = date_obj - datetime.timedelta(days=date_obj.weekday())
    sunday = monday + datetime.timedelta(days=6)
    return f"({monday.strftime('%d-%m-%Y')} to {sunday.strftime('%d-%m-%Y')})"


def calculate_swing_taxes(buy_value, sell_value):
    """Calculate standard Indian stock exchange taxes & charges for swing trade (Delivery)"""
    brokerage = 40.0  # ₹20 Buy + ₹20 Sell
    stt = (buy_value * 0.001) + (sell_value * 0.001)  # 0.1% on Buy & Sell
    exchange_charges = (buy_value * 0.0000325) + (sell_value * 0.0000325)  # NSE txn charges
    gst = 0.18 * (brokerage + exchange_charges)  # 18% GST on brokerage + exchange charges
    sebi_charges = (buy_value + sell_value) * 0.000001  # ₹10 per crore
    stamp_duty = buy_value * 0.00015  # 0.015% on buy side only
    
    total_taxes = brokerage + stt + exchange_charges + gst + sebi_charges + stamp_duty
    return round(total_taxes, 2)


def apply_current_week_styling(ws):
    center_align = Alignment(horizontal='center', vertical='center')

    font_header = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
    fill_header = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')

    font_red = Font(name='Calibri', size=11, color='FF0000', bold=True)
    font_green = Font(name='Calibri', size=11, color='008000', bold=True)
    font_default = Font(name='Calibri', size=11, color='000000')

    fill_target = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    fill_sl = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    fill_active = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
    fill_expired = PatternFill(start_color='D9D9D9', end_color='D9D9D9', fill_type='solid')

    header_map = {str(ws.cell(row=1, column=c).value): c for c in range(1, ws.max_column + 1)}

    for row in range(1, ws.max_row + 1):
        is_total_row = (ws.cell(row=row, column=1).value == "TOTAL / OVERALL")
        ws.row_dimensions[row].height = 24 if is_total_row else (22 if row == 1 else 18)

        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=row, column=col)
            cell.alignment = center_align

            if row == 1:
                cell.font = font_header
                cell.fill = fill_header
            elif is_total_row:
                cell.font = Font(name='Calibri', size=11, bold=True, color='000000')
                cell.fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
            else:
                col_name_list = [k for k, v in header_map.items() if v == col]
                if col_name_list:
                    col_name = col_name_list[0]
                    if col_name in ['SL', 'SL_%(Point)']:
                        cell.font = font_red
                    elif col_name in ['ENTRY', 'TARGET_1:1', 'TARGET_1:2', 'LTP', 'T1_EXIT', 'T2_EXIT', 'TOTAL_EXIT_VALUE', 'Invest', 'all_taxes']:
                        cell.font = font_green
                    elif col_name in ['PnL_%', 'PNL_T1', 'PNL_T2']:
                        val_str = str(cell.value or '').strip()
                        try:
                            cleaned_val = val_str.split('%')[0].replace('+', '').strip()
                            val_num = float(cleaned_val)
                            if val_num < 0:
                                cell.font = font_red
                            else:
                                cell.font = font_green
                        except ValueError:
                            if '-' in val_str:
                                cell.font = font_red
                            else:
                                cell.font = font_green
                    else:
                        cell.font = font_default

                    if col_name == 'Trade_Status':
                        val = str(cell.value)
                        if 'TARGET' in val:
                            cell.fill = fill_target
                            cell.font = font_green
                        elif 'SL HIT' in val:
                            cell.fill = fill_sl
                            cell.font = font_red
                        elif 'ACTIVE' in val:
                            cell.fill = fill_active
                        elif 'EXPIRED' in val:
                            cell.fill = fill_expired
                            cell.font = Font(name='Calibri', size=11, color='595959', italic=True)

    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        col_name = str(ws.cell(row=1, column=col[0].column).value)
        if col_name == 'Ticker':
            ws.column_dimensions[col_letter].width = 14
        else:
            max_len = max(len(str(cell.value or '')) for cell in col)
            ws.column_dimensions[col_letter].width = max(max_len + 2, 8)


def track_trade_status(input_file='weekly_final_trading_signals.xlsx', output_file='trade_tracker_results.xlsx'):
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(script_dir)
    
    input_path = os.path.join(project_root, 'data', os.path.basename(input_file))
    if not os.path.exists(input_path):
        input_path = os.path.join(project_root, os.path.basename(input_file))

    data_dir = os.path.join(project_root, 'data')
    os.makedirs(data_dir, exist_ok=True)
    output_path = os.path.join(data_dir, os.path.basename(output_file))

    df_valid = pd.DataFrame()

    try:
        if os.path.exists(input_path):
            df_valid = pd.read_excel(input_path, sheet_name='Valid_Setups_Only')
    except Exception as e:
        print(f"⚠️ Warning reading '{input_file}': {e}")

    today = datetime.datetime.now()
    monday_start = (today - datetime.timedelta(days=today.weekday())).replace(hour=0, minute=0, second=0, microsecond=0)
    sunday_end = monday_start + datetime.timedelta(days=6, hours=23, minutes=59)

    current_week_str = f"({monday_start.strftime('%d-%m-%Y')} to {sunday_end.strftime('%d-%m-%Y')})"

    tracked_results = []

    if df_valid.empty:
        print("ℹ️ Valid setups sheet is empty for current week. Clearing 'current_week' sheet...")
    else:
        print(f"\n🚀 Tracking Trade Outcomes for Current Week ({len(df_valid)} Valid Stocks)...\n")

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
            t1_exit = 0.0
            t2_exit = 0.0
            latest_exit_price = 0.0
            pnl_pct = 0.0

            amount = 10000.0
            no_stock = int(amount // entry_price) if entry_price > 0 else 0
            if no_stock == 0:
                no_stock = 1  
            
            invest = round(no_stock * entry_price, 2)

            try:
                ticker = yf.Ticker(ticker_symbol)
                df_15m = ticker.history(period="5d", interval="15m")

                if not df_15m.empty:
                    if df_15m.index.tz is not None:
                        df_15m.index = df_15m.index.tz_localize(None)

                    bu_dt = pd.to_datetime(bu_time_str)
                    next_day_start = (bu_dt + datetime.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
                    
                    df_after_bu = df_15m[(df_15m.index >= next_day_start) & (df_15m.index >= monday_start) & (df_15m.index <= sunday_end)]
                    df_after_bu = df_15m.between_time('09:15', '15:15')

                    entry_candles = df_after_bu[df_after_bu['High'] >= entry_price]

                    if not entry_candles.empty:
                        first_entry_candle = entry_candles.iloc[0]
                        entry_triggered_time = first_entry_candle.name.strftime('%Y-%m-%d %H:%M')
                        trade_status = "ACTIVE"

                        df_after_entry = df_after_bu[df_after_bu.index >= first_entry_candle.name]

                        target_1_1_hit = False
                        t2_hit = False
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
                                if not target_1_1_hit:
                                    t1_exit = round(sl_price, 2)
                                else:
                                    t1_exit = round(target_1_1, 2)
                                t2_exit = 0.0
                                latest_exit_price = round(sl_price, 2)
                                break

                            if candle_high >= target_1_1 and not target_1_1_hit:
                                target_1_1_hit = True
                                trade_status = "TARGET 1:1 HIT (Active for 1:2)"
                                exit_time = dt.strftime('%Y-%m-%d %H:%M')
                                t1_exit = round(target_1_1, 2)
                                t2_exit = 0.0
                                latest_exit_price = round(target_1_1, 2)

                            if candle_high >= target_1_2:
                                t2_hit = True
                                trade_status = "TARGET 1:2 HIT"
                                exit_time = dt.strftime('%Y-%m-%d %H:%M')
                                t1_exit = round(target_1_1, 2)
                                t2_exit = round(target_1_2, 2)
                                latest_exit_price = round(target_1_2, 2)
                                break

                        if latest_exit_price == 0.0 and last_candle_dt is not None:
                            exit_time = last_candle_dt.strftime('%Y-%m-%d %H:%M') + " (Active)"
                            latest_exit_price = round(last_candle_close, 2)
                            if target_1_1_hit:
                                t1_exit = round(target_1_1, 2)
                                t2_exit = 0.0
                            else:
                                t1_exit = latest_exit_price
                                t2_exit = 0.0

                    else:
                        trade_status = "WAIT (No Entry)"
                        t1_exit = entry_price
                        t2_exit = 0.0
                        latest_exit_price = entry_price  

            except Exception as e:
                trade_status = f"ERROR: {str(e)}"
                t1_exit = entry_price
                t2_exit = 0.0
                latest_exit_price = entry_price

            if latest_exit_price == 0.0:
                latest_exit_price = entry_price
                t1_exit = entry_price
                t2_exit = 0.0

            t1_value_total = round(t1_exit * no_stock, 2)
            t2_value_total = round(t2_exit * no_stock, 2) if t2_exit > 0 else 0.0
            
            # TOTAL_EXIT_VALUE based on max(T1_EXIT, T2_EXIT) * no_stock
            max_exit_price = max(t1_exit, t2_exit)
            total_exit_val = round(max_exit_price * no_stock, 2)

            # PnL_% formula: ((TOTAL_EXIT_VALUE - INVEST) * 100) / INVEST
            pnl_pct = round(((total_exit_val - invest) * 100) / invest, 2) if invest > 0 else 0.0

            taxes_t1 = calculate_swing_taxes(invest, t1_value_total)
            taxes_t2 = calculate_swing_taxes(invest, t2_value_total) if t2_exit > 0 else 0.0
            all_taxes = calculate_swing_taxes(invest, total_exit_val)

            net_pnl_t1 = round((t1_value_total - invest) - taxes_t1, 2)
            net_pnl_t2 = round((total_exit_val - invest) - all_taxes, 2)

            pnl_t1_pct = round((net_pnl_t1 / amount) * 100, 2)
            pnl_t2_pct = round((net_pnl_t2 / amount) * 100, 2)

            sign_t1 = "+" if net_pnl_t1 >= 0 else ""
            sign_t2 = "+" if net_pnl_t2 >= 0 else ""
            sign_pnl = "+" if pnl_pct >= 0 else ""

            pnl_t1_str = f"{sign_t1}{pnl_t1_pct}% ({sign_t1}₹{net_pnl_t1})"
            pnl_t2_str = f"{sign_t2}{pnl_t2_pct}% ({sign_t2}₹{net_pnl_t2})"
            pnl_pct_str = f"{sign_pnl}{pnl_pct}%"

            wk_range = get_week_range_str(entry_triggered_time if entry_triggered_time != "N/A" else current_week_str)

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
                'T1_EXIT': t1_exit,
                'T2_EXIT': t2_exit,
                'PnL_%': pnl_pct_str,
                'Amount': amount,
                'no_stock': no_stock,
                'Invest': invest,
                'TOTAL_EXIT_VALUE': total_exit_val,
                'all_taxes': all_taxes,
                'PNL_T1': pnl_t1_str,
                'PNL_T2': pnl_t2_str,
                'Week_Range': wk_range
            }
            tracked_results.append(res)

    cols_order = [
        'Ticker', 'PWL', 'SL', 'ENTRY', 'SL_%(Point)', 'TARGET_1:1', 'TARGET_1:2', 
        'Entry_Triggered_Time', 'Trade_Status', 'Exit_Time', 'T1_EXIT', 'T2_EXIT', 'PnL_%',
        'Amount', 'no_stock', 'Invest', 'TOTAL_EXIT_VALUE', 'all_taxes', 'PNL_T1', 'PNL_T2', 'Week_Range'
    ]

    if tracked_results:
        df_current = pd.DataFrame(tracked_results)
    else:
        df_current = pd.DataFrame(columns=cols_order)

    # Compute Totals for current week table
    sum_pnl_pct = 0.0
    if not df_current.empty:
        for val in df_current['PnL_%']:
            try:
                cleaned = str(val).split('%')[0].replace('+', '').strip()
                sum_pnl_pct += float(cleaned)
            except:
                pass
    sum_pnl_pct = round(sum_pnl_pct, 2)

    sum_pnl_t1_rs = 0.0
    sum_pnl_t2_rs = 0.0

    if not df_current.empty:
        for val in df_current['PNL_T1']:
            try:
                rs_part = str(val).split('(')[1].replace(')', '').replace('₹', '').replace('+', '').strip()
                sum_pnl_t1_rs += float(rs_part)
            except:
                pass
        for val in df_current['PNL_T2']:
            try:
                rs_part = str(val).split('(')[1].replace(')', '').replace('₹', '').replace('+', '').strip()
                sum_pnl_t2_rs += float(rs_part)
            except:
                pass

    sum_pnl_t1_rs = round(sum_pnl_t1_rs, 2)
    sum_pnl_t2_rs = round(sum_pnl_t2_rs, 2)

    total_amount_base = len(df_current) * 10000.0 if not df_current.empty else 10000.0
    sum_pnl_t1_pct = round((sum_pnl_t1_rs / total_amount_base) * 100, 2) if total_amount_base > 0 else 0.0
    sum_pnl_t2_pct = round((sum_pnl_t2_rs / total_amount_base) * 100, 2) if total_amount_base > 0 else 0.0

    sign_t1_tot = "+" if sum_pnl_t1_rs >= 0 else ""
    sign_t2_tot = "+" if sum_pnl_t2_rs >= 0 else ""
    sign_pnl_pct = "+" if sum_pnl_pct >= 0 else ""

    total_pnl_t1_str = f"{sign_t1_tot}{sum_pnl_t1_pct}% ({sign_t1_tot}₹{sum_pnl_t1_rs})"
    total_pnl_t2_str = f"{sign_t2_tot}{sum_pnl_t2_pct}% ({sign_t2_tot}₹{sum_pnl_t2_rs})"
    total_pnl_pct_str = f"{sign_pnl_pct}{sum_pnl_pct}%"

    df_history = pd.DataFrame()

    if not df_current.empty:
        df_triggered_only = df_current[
            (df_current['Entry_Triggered_Time'].astype(str) != "N/A") & 
            (~df_current['Trade_Status'].astype(str).str.contains("EXPIRED", case=False, na=False))
        ].copy()
    else:
        df_triggered_only = pd.DataFrame()

    if os.path.exists(output_path):
        try:
            df_existing_history = pd.read_excel(output_path, sheet_name='history_data')
            
            if not df_existing_history.empty and 'Ticker' in df_existing_history.columns:
                df_existing_history = df_existing_history[
                    ~df_existing_history['Ticker'].astype(str).str.startswith("Week Range:") & 
                    ~df_existing_history['Ticker'].astype(str).str.startswith("📊 WEEKLY SUMMARY") &
                    ~df_existing_history['Ticker'].astype(str).str.startswith("🏆 GRAND OVERALL")
                ].copy()

            if not df_existing_history.empty and 'Entry_Triggered_Time' in df_existing_history.columns:
                df_existing_history['Week_Range'] = df_existing_history['Entry_Triggered_Time'].apply(get_week_range_str)

            df_history = pd.concat([df_existing_history, df_triggered_only], ignore_index=True)
            df_history.drop_duplicates(subset=['Ticker', 'Entry_Triggered_Time'], keep='last', inplace=True)
        except Exception:
            df_history = df_triggered_only.copy()
    else:
        df_history = df_triggered_only.copy()

    if not df_history.empty and 'Entry_Triggered_Time' in df_history.columns:
        df_history.sort_values(by=['Entry_Triggered_Time'], ascending=[True], inplace=True)

    if not df_history.empty:
        def update_past_week_active(row):
            w_range = str(row.get('Week_Range', ''))
            if w_range != current_week_str:
                status_str = str(row.get('Trade_Status', ''))
                exit_str = str(row.get('Exit_Time', ''))
                
                if 'ACTIVE' in status_str.upper() and 'WEEK END' not in status_str.upper():
                    row['Trade_Status'] = 'ACTIVE (Week End)'
                if '(Active)' in exit_str:
                    row['Exit_Time'] = exit_str.replace('(Active)', '(Week End)')
            return row

        df_history = df_history.apply(update_past_week_active, axis=1)

    if not df_history.empty:
        df_history = df_history[[c for c in cols_order if c in df_history.columns]]

    # Write output excel with precise structure
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_current_out = df_current.drop(columns=['Week_Range'], errors='ignore')
        df_current_out.to_excel(writer, sheet_name='current_week', index=False)
        
        ws_current = writer.sheets['current_week']
        total_row_data = ['' for _ in range(len(df_current_out.columns))]
        total_row_data[0] = "TOTAL / OVERALL"
        
        col_map = {col_name: idx + 1 for idx, col_name in enumerate(df_current_out.columns)}
        if 'PnL_%' in col_map:
            total_row_data[col_map['PnL_%'] - 1] = total_pnl_pct_str
        if 'PNL_T1' in col_map:
            total_row_data[col_map['PNL_T1'] - 1] = total_pnl_t1_str
        if 'PNL_T2' in col_map:
            total_row_data[col_map['PNL_T2'] - 1] = total_pnl_t2_str
            
        ws_current.append(total_row_data)
        apply_current_week_styling(ws_current)
        
        ws_history = writer.book.create_sheet('history_data')
        
        active_cols_history = [c for c in cols_order if c != 'Week_Range']
        ws_history.append(active_cols_history)
        
        center_align = Alignment(horizontal='center', vertical='center')
        font_header = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
        fill_header = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')

        font_week_range = Font(name='Calibri', size=10, bold=True, color='FFFFFF')
        fill_week_range = PatternFill(start_color='41719C', end_color='41719C', fill_type='solid')

        font_summary = Font(name='Calibri', size=10, bold=True, color='000000')
        fill_summary = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')

        font_grand = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
        fill_grand = PatternFill(start_color='203764', end_color='203764', fill_type='solid')

        font_red = Font(name='Calibri', size=11, color='FF0000', bold=True)
        font_green = Font(name='Calibri', size=11, color='008000', bold=True)
        font_default = Font(name='Calibri', size=11, color='000000')

        fill_target = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
        fill_sl = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
        fill_active = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
        fill_expired = PatternFill(start_color='D9D9D9', end_color='D9D9D9', fill_type='solid')

        ws_history.row_dimensions[1].height = 22
        for col_idx in range(1, len(active_cols_history) + 1):
            cell = ws_history.cell(row=1, column=col_idx)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = center_align

        grand_total_trades = 0
        grand_profit_count = 0
        grand_sl_count = 0
        grand_total_pnl = 0.0

        if not df_history.empty and 'Week_Range' in df_history.columns:
            grouped = df_history.groupby('Week_Range', sort=True)
            current_row = 2
            max_col = len(active_cols_history)
            
            for week_range_val, group_df in grouped:
                ws_history.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=max_col)
                for c in range(1, max_col + 1):
                    cell = ws_history.cell(row=current_row, column=c)
                    cell.fill = fill_week_range
                    cell.font = font_week_range
                    cell.alignment = center_align
                ws_history.cell(row=current_row, column=1).value = f"Week Range: {week_range_val}"
                ws_history.row_dimensions[current_row].height = 22
                current_row += 1

                sl_count = 0
                profit_count = 0
                active_count = 0
                total_pnl = 0.0
                trade_count = len(group_df)
                
                active_minus_count = 0
                active_plus_count = 0
                active_pnl_sum = 0.0
                closed_pnl_sum = 0.0

                for _, r_data in group_df.iterrows():
                    st = str(r_data.get('Trade_Status', '')).upper()
                    pnl_val = 0.0
                    try:
                        v_s = str(r_data.get('PnL_%', 0)).split('%')[0].replace('+', '').strip()
                        pnl_val = float(v_s)
                    except:
                        pass

                    if 'SL' in st and 'TARGET' not in st:
                        sl_count += 1
                        closed_pnl_sum += pnl_val
                    elif 'TARGET' in st:
                        profit_count += 1
                        closed_pnl_sum += pnl_val
                    elif 'ACTIVE' in st:
                        active_count += 1
                        if pnl_val < 0:
                            active_minus_count += 1
                        else:
                            active_plus_count += 1
                        active_pnl_sum += pnl_val
                    else:
                        closed_pnl_sum += pnl_val

                    total_pnl += pnl_val
                
                grand_total_trades += trade_count
                grand_profit_count += profit_count
                grand_sl_count += sl_count
                grand_total_pnl += total_pnl

                net_pnl_rounded = round(closed_pnl_sum, 2)
                week_pnl_total = round(closed_pnl_sum + active_pnl_sum, 2)

                summary_text = (
                    f"📊 WEEKLY SUMMARY -> Total Trades: {trade_count} | Targets Hit: {profit_count} | "
                    f"SL Hit: {sl_count} | Net PnL: {net_pnl_rounded}% | "
                    f"Active -> Minus: {active_minus_count}, Plus: {active_plus_count} | "
                    f"WEEK PNL = {week_pnl_total}%"
                )
                
                ws_history.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=max_col)
                for c in range(1, max_col + 1):
                    cell = ws_history.cell(row=current_row, column=c)
                    cell.fill = fill_summary
                    cell.font = font_summary
                    cell.alignment = center_align
                ws_history.cell(row=current_row, column=1).value = summary_text
                ws_history.row_dimensions[current_row].height = 22
                current_row += 1

                for _, r_data in group_df.iterrows():
                    row_vals = [r_data[c] for c in active_cols_history]
                    ws_history.append(row_vals)
                    ws_history.row_dimensions[current_row].height = 18
                    
                    for c_idx in range(1, max_col + 1):
                        cell = ws_history.cell(row=current_row, column=c_idx)
                        cell.alignment = center_align
                        col_name = active_cols_history[c_idx - 1]
                        
                        if col_name in ['SL', 'SL_%(Point)']:
                            cell.font = font_red
                        elif col_name in ['ENTRY', 'TARGET_1:1', 'TARGET_1:2', 'T1_EXIT', 'T2_EXIT', 'TOTAL_EXIT_VALUE', 'Invest', 'all_taxes']:
                            cell.font = font_green
                        elif col_name in ['PnL_%', 'PNL_T1', 'PNL_T2']:
                            try:
                                v_str = str(cell.value).split('%')[0].replace('+', '').strip()
                                v_num = float(v_str)
                                cell.font = font_red if v_num < 0 else font_green
                            except:
                                cell.font = font_green
                        elif col_name == 'Trade_Status':
                            val = str(cell.value)
                            if 'TARGET' in val:
                                cell.fill = fill_target
                                cell.font = font_green
                            elif 'SL HIT' in val:
                                cell.fill = fill_sl
                                cell.font = font_red
                            elif 'ACTIVE' in val:
                                cell.fill = fill_active
                            elif 'EXPIRED' in val:
                                cell.fill = fill_expired
                                cell.font = Font(name='Calibri', size=11, color='595959', italic=True)
                        else:
                            cell.font = font_default
                    current_row += 1

            grand_text = f"🏆 GRAND OVERALL SUMMARY -> Total Trades: {grand_total_trades} | Total Targets Hit: {grand_profit_count} | Total SL Hit: {grand_sl_count} | Net Cumulative PnL: {round(grand_total_pnl, 2)}%"
            ws_history.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=max_col)
            for c in range(1, max_col + 1):
                cell = ws_history.cell(row=current_row, column=c)
                cell.fill = fill_grand
                cell.font = font_grand
                cell.alignment = center_align
            ws_history.cell(row=current_row, column=1).value = grand_text
            ws_history.row_dimensions[current_row].height = 24

        for col in ws_history.columns:
            col_letter = get_column_letter(col[0].column)
            col_name = active_cols_history[col[0].column - 1] if col[0].column <= len(active_cols_history) else ''
            
            if col_name == 'Ticker':
                ws_history.column_dimensions[col_letter].width = 14
            else:
                max_len = max(len(str(cell.value or '')) for cell in col if not str(cell.value or '').startswith("Week Range:") and not str(cell.value or '').startswith("📊 WEEKLY SUMMARY") and not str(cell.value or '').startswith("🏆 GRAND OVERALL"))
                ws_history.column_dimensions[col_letter].width = max(max_len + 2, 8)

    print(f"\n✅ Tracking Complete & Excel Updated Successfully with Correct PnL% Formula and Totals!")

if __name__ == "__main__":
    track_trade_status()
