# Start-MT5Backtest.ps1
try {
    Import-Module .\MT5BacktestAutomation.psd1 -Force -ErrorAction Stop
    
    # Check if MT5 is running
    $mt5Process = Get-Process -Name "terminal64" -ErrorAction SilentlyContinue
    if (-not $mt5Process) {
        Write-Host "MetaTrader 5 is not running. Please launch MT5 first or use the batch file." -ForegroundColor Yellow
        Write-Host "Press any key to exit..."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit
    }
    
    # Give MT5 some time to fully initialize if it was just started
    Write-Host "Waiting for MetaTrader 5 to fully initialize..." -ForegroundColor Cyan
    Start-Sleep -Seconds 10
    
    # First, let's read the available symbols and timeframes from the INI file
    function Get-AvailableSymbolsAndTimeframes {
        $iniPath = ".\Modified_MT5_Backtest_Config.ini"
        if (-not (Test-Path $iniPath)) {
            if ($script:config -and $script:config.configIniPath) {
                $iniPath = $script:config.configIniPath
            }
            
            if (-not (Test-Path $iniPath)) {
                Write-Error "INI file not found: $iniPath"
                return @(), @()
            }
        }
        
        $iniContent = Get-Content -Path $iniPath -Raw
        
        # Try to extract symbols list
        $symbols = @()
        if ($iniContent -match "Symbols=([^\r\n]+)") {
            $symbolsString = $matches[1]
            $symbols = $symbolsString -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
        }
        
        # If no symbols found in INI, use the default list
        if ($symbols.Count -eq 0) {
            if ($script:constants -and $script:constants.VALID_SYMBOLS) {
                $symbols = $script:constants.VALID_SYMBOLS
            } else {
                $symbols = @("EURUSD", "GBPUSD", "USDJPY", "AUDUSD")
            }
        }
        
        # Try to extract timeframes list
        $timeframes = @()
        if ($iniContent -match "Timeframes=([^\r\n]+)") {
            $timeframesString = $matches[1]
            $timeframes = $timeframesString -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
        }
        
        # If no timeframes found in INI, use the default list
        if ($timeframes.Count -eq 0) {
            if ($script:constants -and $script:constants.VALID_TIMEFRAMES) {
                $timeframes = $script:constants.VALID_TIMEFRAMES
            } else {
                $timeframes = @("M1", "M5", "M15", "H1", "H4", "D1")
            }
        }
        
        return $symbols, $timeframes
    }

    # Get available symbols and timeframes
    $symbols, $timeframes = Get-AvailableSymbolsAndTimeframes

    # Display what we're going to test
    Write-Host "Found $($symbols.Count) symbols and $($timeframes.Count) timeframes in the INI file."
    Write-Host "Symbols: $($symbols -join ', ')"
    Write-Host "Timeframes: $($timeframes -join ', ')"
    Write-Host "Total backtests to run: $($symbols.Count * $timeframes.Count)"

    # Confirm before proceeding with a large number of tests
    $totalTests = $symbols.Count * $timeframes.Count
    if ($totalTests -gt 10) {
        $confirmation = Read-Host "This will run $totalTests backtests which may take a long time. Continue? (Y/N)"
        if ($confirmation -ne 'Y') {
            Write-Host "Operation cancelled by user."
            exit
        }
    }

    # Run backtests for all symbols and timeframes found in the INI file
    Write-Host "Starting backtest automation..." -ForegroundColor Green
    Start-MultiSymbolBacktestAutomation -Symbols $symbols -Timeframes $timeframes
} catch {
    Write-Host "An error occurred: $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    Write-Host "Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}