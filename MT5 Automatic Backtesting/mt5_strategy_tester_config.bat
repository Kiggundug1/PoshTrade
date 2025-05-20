@echo off
echo Opening MetaTrader 5 with Strategy Tester configuration...

:: Path to MetaTrader 5 executable - update this path to match your installation
set MT5_PATH="C:\Program Files\MetaTrader 5 EXNESS\terminal64.exe"
set CONFIG_PATH=".\MT5_Backtest_Config.ini"
set LOG_PATH=".\logs"

:: Create logs directory if it doesn't exist
if not exist "%LOG_PATH%" mkdir "%LOG_PATH%"

:: Check if MT5 exists at the specified path
if not exist %MT5_PATH% (
    echo MetaTrader 5 not found at %MT5_PATH%
    echo Please update the path in this script to match your installation
    echo %date% %time% - ERROR: MetaTrader 5 not found at %MT5_PATH% >> "%LOG_PATH%\backtest_launch.log"
    pause
    exit /b 1
)

:: Launch MT5 with config file
echo %date% %time% - Starting MT5 with config file >> "%LOG_PATH%\backtest_launch.log"
start "" %MT5_PATH% /config:%CONFIG_PATH%

echo MT5 launched with Strategy Tester configuration.
echo Check the MT5 window to ensure the Strategy Tester has opened correctly.

pause