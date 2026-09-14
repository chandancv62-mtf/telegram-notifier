import os
import pandas as pd
import requests
import config

def send_telegram_summary():
    script_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    file_path = os.path.join(script_dir, config.TRADE_TRACKER_PATH)
    
    bot_token = os.environ.get("TELEGRAM_TOKEN")
    chat_id = os.environ.get("TELEGRAM_TO")
    
    if not bot_token or not chat_id:
        print("⚠️ Telegram Credentials missing in environment variables.")
        return

    try:
        df_curr = pd.read_excel(file_path, sheet_name='current_week')
    except Exception as e:
        print(f"❌ Error loading tracker file: {e}")
        return

    total_stocks = len(df_curr)
    target_hits = len(df_curr[df_curr['Trade_Status'].str.contains('TARGET', na=False)])
    sl_hits = len(df_curr[df_curr['Trade_Status'].str.contains('SL HIT', na=False)])
    active_trades = len(df_curr[df_curr['Trade_Status'] == 'ACTIVE'])
    avg_pnl = df_curr['PnL_%'].mean()

    msg = f"📊 *DAILY TRADING SUMMARY REPORT*\n"
    msg += f"-----------------------------------\n"
    msg += f"🎯 *Target Hits:* {target_hits}\n"
    msg += f"⏳ *Active Trades:* {active_trades}\n"
    msg += f"🛑 *SL Hits:* {sl_hits}\n"
    msg += f"📈 *Avg PnL Today:* {avg_pnl:.2f}%\n"
    msg += f"-----------------------------------\n\n"
    msg += f"*Executed Trades Details:*\n"

    for _, row in df_curr.iterrows():
        if row['Trade_Status'] != 'ENTRY EXPIRED (Week End)':
            pnl_str = f"+{row['PnL_%']}%" if row['PnL_%'] > 0 else f"{row['PnL_%']}%"
            msg += f"• *{row['Ticker']}*: `{row['Trade_Status']}` | PnL: `{pnl_str}`\n"

    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    payload = {"chat_id": chat_id, "text": msg, "parse_mode": "Markdown"}
    
    res = requests.post(url, data=payload)
    if res.status_code == 200:
        print("✅ Telegram Summary Alert Sent Successfully!")
    else:
        print(f"❌ Failed to send Telegram Alert: {res.text}")

if __name__ == "__main__":
    send_telegram_summary()
