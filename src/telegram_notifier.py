import os
import pandas as pd
import requests
import config

def send_telegram_summary():
    script_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    tracker_path = os.path.join(script_dir, config.TRADE_TRACKER_PATH)
    next_day_path = os.path.join(script_dir, config.WEEKLY_SIGNALS_PATH)
    
    bot_token = os.environ.get("TELEGRAM_TOKEN")
    chat_id = os.environ.get("TELEGRAM_TO")
    
    if not bot_token or not chat_id:
        print("Telegram Credentials missing in environment variables.")
        return

    msg = "🚀 *DAILY TRADING SYSTEM REPORT*\n"
    msg += "=============================\n\n"

    # SECTION 1: TODAY'S TRADE EXECUTION DETAILS
    try:
        df_curr = pd.read_excel(tracker_path, sheet_name='current_week')
        if not df_curr.empty:
            msg += "📌 *TODAY'S EXECUTED TRADES*\n\n"
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

                msg += f"{icon} *{ticker}* — `{status}`\n"
                msg += f"   • Entry: ₹{entry} | SL: ₹{sl}\n"
                msg += f"   • Target 1:1: ₹{t1} | Target 1:2: ₹{t2}\n"
                if trig_time != "N/A":
                    msg += f"   • Trigger Time: `{trig_time}`\n"
                if exit_time != "N/A":
                    msg += f"   • Exit Time: `{exit_time}`\n"
                if status != "WAIT (No Entry)":
                    msg += f"   • Result PnL: `{pnl_str}`\n"
                msg += "\n"
    except Exception as e:
        msg += f"⚠️ Trade tracker read error: {e}\n\n"

    # SECTION 2: NEXT DAY WATCHLIST (NEW SETUPS)
    try:
        df_next = pd.read_excel(next_day_path, sheet_name='Valid_Setups_Only')
        msg += "-----------------------------\n"
        msg += "📋 *NEXT DAY WATCHLIST (New Signals)*\n\n"
        if df_next.empty:
            msg += "आज कल के लिए कोई नया setup नहीं मिला।\n"
        else:
            for _, row in df_next.iterrows():
                msg += f"🔹 *{row['Ticker']}*\n"
                msg += f"   • Entry Trigger Above: ₹{row['ENTRY']}\n"
                msg += f"   • Stop Loss: ₹{row['SL']}\n"
                msg += f"   • Targets: T1 ₹{row['TARGET_1:1']} | T2 ₹{row['TARGET_1:2']}\n\n"
    except Exception as e:
        msg += f"⚠️ Next day signals read error: {e}\n\n"

    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    payload = {"chat_id": chat_id, "text": msg, "parse_mode": "Markdown"}
    
    res = requests.post(url, data=payload)
    if res.status_code == 200:
        print("Telegram Alert Sent Successfully!")
    else:
        print(f"Failed to send Telegram Alert: {res.text}")

if __name__ == "__main__":
    send_telegram_summary()
