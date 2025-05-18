@echo off
echo Starting MT5 Backtest Automation...

set MT5_PATH="C:\Program Files\MetaTrader 5 EXNESS\terminal64.exe"
set CONFIG_PATH=".\MT5_Backtest_Config.ini"
set LOG_PATH=".\Reports\logs"

if not exist "%LOG_PATH%" mkdir "%LOG_PATH%"

echo %date% %time% - Starting MT5 with config file >> "%LOG_PATH%\backtest_launch.log"
start "" %MT5_PATH% /config:%CONFIG_PATH%

echo MT5 launched. Power Automate will now handle the backtest process.