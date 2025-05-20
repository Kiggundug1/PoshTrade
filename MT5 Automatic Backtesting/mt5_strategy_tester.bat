@echo off
echo Opening MetaTrader 5 and configuring Strategy Tester...

:: Path to MetaTrader 5 executable - update this path to match your installation
set MT5_PATH="C:\Program Files\MetaTrader 5 EXNESS\terminal64.exe"
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

:: Launch MT5 directly with parameters
echo %date% %time% - Starting MT5 >> "%LOG_PATH%\backtest_launch.log"
start "" %MT5_PATH% /portable

:: Wait for MT5 to initialize
timeout /t 5 /nobreak

:: Send keystrokes to open Strategy Tester (Ctrl+R)
echo Opening Strategy Tester...
powershell -command "$wshell = New-Object -ComObject wscript.shell; $wshell.AppActivate('MetaTrader 5'); Start-Sleep -Milliseconds 1000; $wshell.SendKeys('^r')"

:: Wait for Strategy Tester to open
timeout /t 3 /nobreak

:: Configure Strategy Tester parameters using keyboard navigation and input
echo Configuring Strategy Tester...
powershell -command "$wshell = New-Object -ComObject wscript.shell; $wshell.AppActivate('Strategy Tester'); Start-Sleep -Milliseconds 1000; $wshell.SendKeys('%%s'); Start-Sleep -Milliseconds 500; $wshell.SendKeys('EURUSD'); Start-Sleep -Milliseconds 500; $wshell.SendKeys('{TAB}'); Start-Sleep -Milliseconds 500; $wshell.SendKeys('{TAB}'); Start-Sleep -Milliseconds 500; $wshell.SendKeys('2019.01.01'); Start-Sleep -Milliseconds 500; $wshell.SendKeys('{TAB}'); Start-Sleep -Milliseconds 500; $wshell.SendKeys('2024.12.31'); Start-Sleep -Milliseconds 500; $wshell.SendKeys('{TAB}'); Start-Sleep -Milliseconds 500; $wshell.SendKeys('10000'); Start-Sleep -Milliseconds 500; $wshell.SendKeys('{TAB}'); Start-Sleep -Milliseconds 500; $wshell.SendKeys('2000'); Start-Sleep -Milliseconds 500; $wshell.SendKeys('{TAB}'); Start-Sleep -Milliseconds 500; $wshell.SendKeys(' '); Start-Sleep -Milliseconds 500;"

echo %date% %time% - Configuration complete >> "%LOG_PATH%\backtest_launch.log"
echo Configuration complete!
echo Note: You may need to adjust the script if your MT5 interface differs from the standard layout.

pause