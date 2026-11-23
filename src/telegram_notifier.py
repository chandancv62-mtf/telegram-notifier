import os
import re
import requests
import pandas as pd
from datetime import datetime


def send_telegram_message(bot_token, chat_id, text):
    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    payload = {
        "chat_id": chat_id,
        "text": text,
        "parse_mode": "Markdown",
        "disable_web_page_preview": True,
    }
    res = requests.post(url, data=payload)
    if res.status_code != 200:
        print(f"Failed to send Telegram Alert: {res.text}")


def get_valid_filepath(filename):
    """Check data/ directory first, then root directory."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(script_dir)

    candidates = [
        os.path.join(project_root, "data", os.path.basename(filename)),
        os.path.join(project_root, os.path.basename(filename)),
        os.path.join(script_dir, "data", os.path.basename(filename)),
        os.path.join(script_dir, os.path.basename(filename)),
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    return candidates[0]


def parse_pct(val):
    """Safely extract numeric percentage from strings like '+1.52%', '-7.89%', '+0.18% (+₹17.61)'."""
    if pd.isna(val):
        return 0.0
    s = str(val).strip()
    if s in ["", "nan", "NaN", "None", "N/A"]:
        return 0.0
    m = re.search(r"([+-]?\d+(?:\.\d+)?)\s*%", s)
    if m:
        return float(m.group(1))
    cleaned = s.replace("%", "").replace("+", "").strip()
    try:
        return float(cleaned)
    except Exception:
        return 0.0


def format_target_label(status_str):
    st = str(status_str).strip()
    st_upper = st.upper()
    if "1:2" in st_upper:
        return "Target T1, T2 hit"
    if "1:1" in st_upper:
        return "Target T1 hit"
    return st


def send_telegram_summary(
    tracker_file="trade_tracker_results.xlsx",
    next_day_file="weekly_final_trading_signals.xlsx",
):
    tracker_path = get_valid_filepath(tracker_file)
    next_day_path = get_valid_filepath(next_day_file)

    bot_token = os.environ.get("TELEGRAM_TOKEN")
    chat_id = os.environ.get("TELEGRAM_TO")

    if not bot_token or not chat_id:
        print("❌ Telegram Credentials missing in environment variables.")
        return

    now_dt = datetime.now()
    today_str = now_dt.strftime("%Y-%m-%d")

    try:
        if not os.path.exists(tracker_path):
            send_telegram_message(bot_token, chat_id, "Trade tracker file nahi mili.")
            return

        df_curr = pd.read_excel(tracker_path, sheet_name="current_week").dropna(how="all")

        ticker_str = df_curr["Ticker"].astype(str).str.strip()
        is_summary_row = (
            ticker_str.str.upper().str.startswith("TOTAL")
            | ticker_str.str.contains(
                "WEEKLY SUMMARY|GRAND OVERALL SUMMARY|Week Range",
                case=False,
                na=False,
            )
        )
        df_monitored = df_curr[~is_summary_row].copy()

        if df_monitored.empty:
            send_telegram_message(bot_token, chat_id, "Aaj ke liye koi monitored trade nahi hai.")
            return

        df_executed = df_monitored[
            df_monitored["Trade_Status"].astype(str).str.strip() != "WAIT (No Entry)"
        ].copy()

        df_executed["pnl_num"] = df_executed["PnL_%"].apply(parse_pct)
        df_executed["t1_pct"] = (
            df_executed["PNL_T1"].apply(parse_pct)
            if "PNL_T1" in df_executed.columns
            else 0.0
        )
        df_executed["t2_pct"] = (
            df_executed["PNL_T2"].apply(parse_pct)
            if "PNL_T2" in df_executed.columns
            else 0.0
        )

        status_s = df_executed["Trade_Status"].astype(str).str.strip()
        t1_hit_df = df_executed[status_s.str.contains("1:1", case=False, na=False)]
        t2_hit_df = df_executed[status_s.str.contains("1:2", case=False, na=False)]
        sl_hit_df = df_executed[status_s.str.upper() == "SL HIT"]
        active_df = df_executed[status_s.str.upper() == "ACTIVE"]

        total_monitored = len(df_monitored)
        cnt_t1 = len(t1_hit_df)
        cnt_t2 = len(t2_hit_df)
        cnt_sl = len(sl_hit_df)
        cnt_active = len(active_df)

        today_mask = (
            df_executed["Exit_Time"].astype(str).str.contains(today_str, na=False)
            | df_executed["Entry_Triggered_Time"].astype(str).str.contains(today_str, na=False)
        )
        today_df = df_executed[today_mask]

        todays_pnl_t1 = float(today_df["t1_pct"].sum()) if not today_df.empty else 0.0
        todays_pnl_t2 = float(today_df["t2_pct"].sum()) if not today_df.empty else 0.0
        overall_day_pnl = float(today_df["pnl_num"].sum()) if not today_df.empty else 0.0
        weekly_pnl = float(df_executed["pnl_num"].sum())

        msg = (
            f"*Total Monitored:* `{total_monitored}`\n"
            f"*Target Hit T1:* `{cnt_t1}` | *T2:* `{cnt_t2}`\n"
            f"*SL Hit:* `{cnt_sl}` | *Active:* `{cnt_active}`\n"
            f"*Today's PnL T1:* `{todays_pnl_t1:+.2f}%` | *T2:* `{todays_pnl_t2:+.2f}%`\n"
            f"*Overall Day PnL:* `{overall_day_pnl:+.2f}%`\n"
            f"*Weekly PnL:* `{weekly_pnl:+.2f}%`\n"
            f"──────────────────────────\n"
            f"*TARGET TODAY STOCKS*\n"
        )

        target_today_df = df_executed[
            status_s.str.contains("TARGET", case=False, na=False)
            & df_executed["Exit_Time"].astype(str).str.contains(today_str, na=False)
        ]

        if target_today_df.empty:
            msg += "_Aaj koi Target hit nahi hua._\n"
        else:
            for _, r in target_today_df.iterrows():
                target_lbl = format_target_label(r["Trade_Status"])
                pnl_val = float(r["pnl_num"])
                t1_val = r.get("TARGET_1:1", "-")
                t2_val = r.get("TARGET_1:2", "-")
                entry_dt = str(r.get("Entry_Triggered_Time", "-")).strip()
                exit_dt = str(r.get("Exit_Time", "-")).strip()
                msg += (
                    f"\n*{r['Ticker']}* — `{target_lbl}`\n"
                    f"*Entry:* `₹{r['ENTRY']}` | *SL:* `₹{r['SL']}` | *PnL:* `{pnl_val:+.2f}%`\n"
                    f"*Target T1(`₹{t1_val}`) | T2(`₹{t2_val}`)*\n"
                    f"*Entry Date/Time:* `{entry_dt}`\n"
                    f"*Exit Date/Time:* `{exit_dt}`\n"
                )

        msg += "\n*ACTIVE STOCKS*\n"
        if active_df.empty:
            msg += "_Koi active stock nahi hai._\n"
        else:
            for _, r in active_df.iterrows():
                pnl_val = float(r["pnl_num"])
                t1_val = r.get("TARGET_1:1", "-")
                t2_val = r.get("TARGET_1:2", "-")
                entry_dt = str(r.get("Entry_Triggered_Time", "-")).strip()
                status_dt = str(r.get("Exit_Time", "")).strip() or "Active"
                msg += (
                    f"\n*{r['Ticker']}* — `Active`\n"
                    f"*Entry:* `₹{r['ENTRY']}` | *SL:* `₹{r['SL']}` | *PnL:* `{pnl_val:+.2f}%`\n"
                    f"*Target T1(`₹{t1_val}`) | T2(`₹{t2_val}`)*\n"
                    f"*Entry Date/Time:* `{entry_dt}`\n"
                    f"*Status/Time:* `{status_dt}`\n"
                )

        msg += "\n──────────────────────────\n*NEXT DAY WATCHLIST*\n"
        executed_tickers = set(df_executed["Ticker"].dropna().astype(str).str.strip().tolist())
        if os.path.exists(next_day_path):
            try:
                df_next = pd.read_excel(next_day_path, sheet_name="Valid_Setups_Only").dropna(how="all")
                if not df_next.empty and executed_tickers:
                    df_next = df_next[~df_next["Ticker"].astype(str).str.strip().isin(executed_tickers)]
                if df_next.empty:
                    msg += "_Agle din ke liye koi naya pending setup nahi mila._\n"
                else:
                    for _, row in df_next.iterrows():
                        t1_val = row.get("TARGET_1:1", "-")
                        t2_val = row.get("TARGET_1:2", "-")
                        msg += (
                            f"\n• *{row['Ticker']}*\n"
                            f"  *Entry:* `₹{row['ENTRY']}` | *SL:* `₹{row['SL']}`\n"
                            f"  *Target T1(`₹{t1_val}`) | T2(`₹{t2_val}`)*\n"
                        )
            except Exception:
                msg += "_Watchlist file read karne mein dikkat aayi._\n"
        else:
            msg += "_Agle din ke liye koi naya pending setup nahi mila._\n"

        if len(msg) > 3800:
            send_telegram_message(bot_token, chat_id, msg[:3800])
            send_telegram_message(bot_token, chat_id, msg[3800:])
        else:
            send_telegram_message(bot_token, chat_id, msg.strip())

    except Exception as e:
        send_telegram_message(bot_token, chat_id, f"⚠️ Trade tracker read error: {e}")


if __name__ == "__main__":
    send_telegram_summary()
