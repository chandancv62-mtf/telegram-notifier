# System Configurations & Parameters

# Market Hours & Timing Settings
MARKET_START_TIME = "09:15"
MARKET_END_TIME = "15:15"
SCHEDULE_CRON_TIME = "15 10 * * 1-5"  # 4:15 PM IST (Mon-Fri)

# Target & Risk Parameters
RISK_REWARD_RATIOS = [1.0, 2.0]  # 1:1 and 1:2 Targets

# File Paths
STOCKS_CSV_PATH = "data/stocks.csv"
FILTERED_STOCKS_PATH = "data/filtered_stocks.csv"
ENTRY_SL_SIGNALS_PATH = "data/entry_sl_signals.xlsx"
WEEKLY_SIGNALS_PATH = "data/weekly_final_trading_signals.xlsx"
TRADE_TRACKER_PATH = "data/trade_tracker_results.xlsx"
