import os
import requests
import pandas as pd
from datetime import datetime

def send_telegram_message(bot_token, chat_id, text):
    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    payload = {"chat_id": chat_id, "text": text, "parse_mode": "Markdown"}
    res = requests.post(url, data=payload)
    if res.status_code != 200:
        print(f"Failed to send Telegram Alert: {res.text}")

def get_valid_filepath(filename):
    """Check data/ directory first, then root directory"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(script_dir)
    
    path_data = os.path.join(project_root, "data", os.path.basename(filename))
    path_root = os.path.join(project_root, os.path.basename(filename))
    path_script_data = os.path.join(script_dir, "data", os.path.basename(filename))
    
    if os.path.exists(path_data):
        return path_data
    elif os.path.exists(path_root):
        return path_root
    elif os.path.exists(path_script_data):
        return path_script_data
    return path_data

def send_telegram_summary(tracker_file='trade_tracker_results.xlsx', next_day_file='weekly_final_trading_signals.xlsx'):
    tracker_path = get_valid_filepath(tracker_file)
    next_day_path = get_valid_filepath(next_day_file)
    
    bot_token = os.environ.get("TELEGRAM_TOKEN")
    chat_id = os.environ.get("TELEGRAM_TO")
    
    if not bot_token or not chat_id:
        print("❌ Telegram Credentials missing in environment variables.")
        return

    # MESSAGE 1: TRADE EXECUTION SUMMARY (Minimalist & Clean Look with Daily & Total PnL)
    try:
        if os.path.exists(tracker_path):
            df_curr = pd.read_excel(tracker_path, sheet_name='current_week').dropna(how='all')
            
            if not df_curr.empty:
                # Exclude 'WAIT (No Entry)' from execution summary dashboard
                df_executed = df_curr[df_curr['Trade_Status'].astype(str) != 'WAIT (No Entry)'].copy()
                
                targets_df = df_executed[df_executed['Trade_Status'].astype(str).str.contains('TARGET', case=False, na=False)]
                sl_df = df_executed[df_executed['Trade_Status'].astype(str).str.contains('SL HIT', case=False, na=False)]
                active_df = df_executed[df_executed['Trade_Status'].astype(str) == 'ACTIVE']
                other_df = df_executed[~df_executed.index.isin(targets_df.index.union(sl_df.index).union(active_df.index))]

                total_monitored = len(df_executed)
                cnt_target = len(targets_df)
                cnt_sl = len(sl_df)
                cnt_active = len(active_df)

                # Calculate PnL (Daily/Today vs Total Week)
                df_executed['PnL_%'] = pd.to_numeric(df_executed['PnL_%'], errors='coerce').fillna(0)
                total_week_pnl = df_executed['PnL_%'].sum()
                
                # For demonstration, today's PnL (checking exit date or current system date, e.g., today)
                today_str = datetime.now().strftime('%Y-%m-%d')
                df_today = df_executed[df_executed['Exit_Time'].astype(str).str.contains(today_str, na=False)]
                daily_pnl = df_today['PnL_%'].sum() if not df_today.empty else 0.0

                daily_pnl_str = f"+{daily_pnl:.2f}%" if daily_pnl >= 0 else f"{daily_pnl:.2f}%"
                week_pnl_str = f"+{total_week_pnl:.2f}%" if total_week_pnl >= 0 else f"{total_week_pnl:.2f}%"

                # Header Banner
                msg = "📊 *TRADING PERFORMANCE SUMMARY*\n"
                msg += "──────────────────────────────\n"
                msg += f"📈 *Total Monitored:* `{total_monitored}`\n"
                msg += f"🎯 *Target Hit:* `{cnt_target}`  |  🔴 *SL Hit:* `{cnt_sl}`\n"
                msg += f"🟢 *Active:* `{cnt_active}`\n"
                msg += f"📅 *Today's PnL:* `{daily_pnl_str}` ⚡\n"
                msg += f"💰 *Total Week PnL:* `{week_pnl_str}` 🚀\n"
                msg += "──────────────────────────────\n\n"

                def build_category_msg(df_subset, category_title, icon):
                    if df_subset.empty:
                        return ""
                    section_text = f"{icon} *{category_title}*\n"
                    for _, row in df_subset.iterrows():
                        ticker = row['Ticker']
                        status = str(row['Trade_Status'])
                        entry = row['ENTRY']
                        sl = row['SL']
                        t2 = row['TARGET_1:2']
                        trig_time = str(row['Entry_Triggered_Time'])
                        exit_time = str(row['Exit_Time'])
                        pnl = float(row['PnL_%'] or 0)
                        
                        pnl_display = f"🟢 `+{pnl}%`" if pnl > 0 else (f"🔴 `{pnl}%`" if pnl < 0 else f"`{pnl}%`")
                        
                        # Highlighting target/SL status distinctly
                        status_highlight = f"🌟 *{status}* 🌟" if "TARGET" in status else (f"⚠️ *{status}* ⚠️" if "SL" in status else f"`{status}`")

                        section_text += f"▪️ *{ticker}* ⟾ {status_highlight}\n"
                        section_text += f"   ├ *Entry:* ₹{entry} │ *SL:* ₹{sl}\n"
                        section_text += f"   ├ *Target (1:2):* ₹{t2} │ *PnL:* {pnl_display}\n"
                        if trig_time not in ["nan", "N/A", "None", ""]:
                            section_text += f"   ├ *Entry Date/Time:* `{trig_time}`\n"
                        if exit_time not in ["nan", "N/A", "None", ""]:
                            section_text += f"   └ *Exit Date/Time:* `{exit_time}`\n"
                        else:
                            section_text += f"   └ *Exit Date/Time:* `Running / Active`\n"
                        section_text += "\n"
                    return section_text

                sections = [
                    build_category_msg(targets_df, "TARGET HIT STOCKS:", "🎯"),
                    build_category_msg(sl_df, "SL HIT STOCKS:", "🔴"),
                    build_category_msg(active_df, "ACTIVE TRADES:", "🟢"),
                    build_category_msg(other_df, "OTHER STATUS:", "📌")
                ]

                for sec in sections:
                    if sec:
                        if len(msg + sec) > 3800:
                            send_telegram_message(bot_token, chat_id, msg)
                            msg = "📌 *TRADING SUMMARY (Contd.)*\n\n"
                        msg += sec

                send_telegram_message(bot_token, chat_id, msg)
            else:
                send_telegram_message(bot_token, chat_id, "📌 *TRADING SUMMARY*\n──────────────────────────────\n\nAaj ke liye koi executed trade nahi hai.")
        else:
            send_telegram_message(bot_token, chat_id, "📌 *TRADING SUMMARY*\n──────────────────────────────\n\nTrade tracker file mil nahi rahi hai.")
    except Exception as e:
        send_telegram_message(bot_token, chat_id, f"⚠️ Trade tracker read error: {e}")

    # MESSAGE 2: NEXT DAY WATCHLIST (Pending Signals Only)
    try:
        executed_tickers = set()
        if os.path.exists(tracker_path):
            try:
                df_curr_chk = pd.read_excel(tracker_path, sheet_name='current_week').dropna(how='all')
                if not df_curr_chk.empty:
                    triggered_df = df_curr_chk[df_curr_chk['Trade_Status'].astype(str) != 'WAIT (No Entry)']
                    executed_tickers = set(triggered_df['Ticker'].tolist())
            except Exception as e:
                print(f"Warning reading tracker for filtering next day list: {e}")

        if os.path.exists(next_day_path):
            df_next = pd.read_excel(next_day_path, sheet_name='Valid_Setups_Only').dropna(how='all')
            
            if not df_next.empty and executed_tickers:
                df_next = df_next[~df_next['Ticker'].isin(executed_tickers)]

            msg_next = "📋 *NEXT DAY WATCHLIST (Pending Signals)*\n"
            msg_next += "──────────────────────────────\n\n"
            
            if df_next.empty:
                msg_next += "Agle din ke liye koi naya pending setup nahi mila."
                send_telegram_message(bot_token, chat_id, msg_next)
            else:
                for _, row in df_next.iterrows():
                    item_msg = f"🔹 *{row['Ticker']}*\n"
                    item_msg += f"   • Entry Trigger Above: ₹{row['ENTRY']}\n"
                    item_msg += f"   • Stop Loss: ₹{row['SL']}\n"
                    item_msg += f"   • Targets: T1 ₹{row['TARGET_1:1']} | T2 ₹{row['TARGET_1:2']}\n\n"

                    if len(msg_next + item_msg) > 3800:
                        send_telegram_message(bot_token, chat_id, msg_next)
                        msg_next = "📋 *NEXT DAY WATCHLIST (Contd.)*\n\n"

                    msg_next += item_msg

                send_telegram_message(bot_token, chat_id, msg_next)
        else:
            send_telegram_message(bot_token, chat_id, "📋 *NEXT DAY WATCHLIST*\n──────────────────────────────\n\nAaj kal ke liye koi naya setup nahi mila.")
    except Exception as e:
        send_telegram_message(bot_token, chat_id, f"⚠️ Next day signals read error: {e}")

if __name__ == "__main__":
    send_telegram_summary()
