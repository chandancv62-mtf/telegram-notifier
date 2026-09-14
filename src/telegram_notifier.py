import os
import requests
import pandas as pd

def send_telegram_message(bot_token, chat_id, text):
    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    payload = {"chat_id": chat_id, "text": text, "parse_mode": "Markdown"}
    res = requests.post(url, data=payload)
    if res.status_code != 200:
        print(f"Failed to send Telegram Alert: {res.text}")

def send_telegram_summary(tracker_file='trade_tracker_results.xlsx', next_day_file='weekly_final_trading_signals.xlsx'):
    script_dir = os.path.dirname(os.path.abspath(__file__))
    tracker_path = os.path.join(script_dir, tracker_file)
    next_day_path = os.path.join(script_dir, next_day_file)
    
    bot_token = os.environ.get("TELEGRAM_TOKEN")
    chat_id = os.environ.get("TELEGRAM_TO")
    
    if not bot_token or not chat_id:
        print("❌ Telegram Credentials missing in environment variables.")
        return

    # MESSAGE 1: TODAY'S TRADE EXECUTION DETAILS
    try:
        if os.path.exists(tracker_path):
            df_curr = pd.read_excel(tracker_path, sheet_name='current_week')
            
            if not df_curr.empty:
                msg = "📌 *TODAY'S EXECUTED TRADES*\n"
                msg += "=============================\n\n"
                
                for _, row in df_curr.iterrows():
                    ticker = row['Ticker']
                    status = str(row['Trade_Status'])
                    entry = row['ENTRY']
                    sl = row['SL']
                    t1 = row['TARGET_1:1']
                    t2 = row['TARGET_1:2']
                    trig_time = row['Entry_Triggered_Time']
                    exit_time = row['Exit_Time']
                    pnl = row['PnL_%']
                    pnl_str = f"+{pnl}%" if pnl > 0 else f"{pnl}%"

                    if "TARGET" in status:
                        icon = "🎯"
                    elif "SL HIT" in status:
                        icon = "🔴"
                    elif status == "ACTIVE":
                        icon = "🟢"
                    else:
                        icon = "⏳"

                    item_msg = f"{icon} *{ticker}* — `{status}`\n"
                    item_msg += f"   • Entry: ₹{entry} | SL: ₹{sl}\n"
                    item_msg += f"   • Target 1:1: ₹{t1} | Target 1:2: ₹{t2}\n"
                    if str(trig_time) != "N/A":
                        item_msg += f"   • Trigger Time: `{trig_time}`\n"
                    if str(exit_time) != "N/A":
                        item_msg += f"   • Exit Time: `{exit_time}`\n"
                    if status != "WAIT (No Entry)":
                        item_msg += f"   • Result PnL: `{pnl_str}`\n"
                    item_msg += "\n"

                    # If message reaches limit, send it and start a new message
                    if len(msg + item_msg) > 3800:
                        send_telegram_message(bot_token, chat_id, msg)
                        msg = "📌 *TODAY'S EXECUTED TRADES (Contd.)*\n\n"
                    
                    msg += item_msg

                send_telegram_message(bot_token, chat_id, msg)
            else:
                send_telegram_message(bot_token, chat_id, "📌 *TODAY'S EXECUTED TRADES*\n=============================\n\nआज के लिए कोई एक्टिव ट्रेड नहीं है।")
    except Exception as e:
        send_telegram_message(bot_token, chat_id, f"⚠️ Trade tracker read error: {e}")

    # MESSAGE 2: NEXT DAY WATCHLIST (NEW SETUPS)
    try:
        if os.path.exists(next_day_path):
            df_next = pd.read_excel(next_day_path, sheet_name='Valid_Setups_Only')
            msg_next = "📋 *NEXT DAY WATCHLIST (New Signals)*\n"
            msg_next += "=============================\n\n"
            
            if df_next.empty:
                msg_next += "आज कल के लिए कोई नया setup नहीं मिला।"
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
    except Exception as e:
        send_telegram_message(bot_token, chat_id, f"⚠️ Next day signals read error: {e}")

if __name__ == "__main__":
    send_telegram_summary()
