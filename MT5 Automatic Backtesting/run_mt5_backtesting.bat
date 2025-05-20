@echo off
set MT5_PATH="C:\Program Files\MetaTrader 5 EXNESS\terminal64.exe"
set INI_PATH=D:\FOREX\Coding\Git_&_Github\GitHub\PoshTrade\MT5 Automatic Backtesting\test_config.ini

start "" %MT5_PATH% /config:%INI_PATH%
