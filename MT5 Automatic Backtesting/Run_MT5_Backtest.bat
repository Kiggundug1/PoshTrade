@echo off
set MT5_PATH="C:\Program Files\MetaTrader 5 EXNESS\terminal64.exe"
set CONFIG_FILE="D:\KGDFRY\Forex Trading\GitHub\PoshTrade\MT5 Automatic Backtesting\MT5_Backtest_Config.ini"

%MT5_PATH% /config:%CONFIG_FILE%
pause