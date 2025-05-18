# Start-MT5Backtest.ps1
Import-Module .\MT5BacktestAutomation.psd1 -Force

# First, let's read the available symbols and timeframes from the INI file
function Get-AvailableSymbolsAndTimeframes {
    $iniPath = $script:config.configIniPath
    if (-not (Test-Path $iniPath)) {
        Write-Error "INI file not found: $iniPath"
        return @{}, @{}
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
        $symbols = $script:constants.VALID_SYMBOLS
    }
    
    # Try to extract timeframes list
    $timeframes = @()
    if ($iniContent -match "Timeframes=([^\r\n]+)") {
        $timeframesString = $matches[1]
        $timeframes = $timeframesString -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
    }
    
    # If no timeframes found in INI, use the default list
    if ($timeframes.Count -eq 0) {
        $timeframes = $script:constants.VALID_TIMEFRAMES
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
if ($totalTests -gt 10000) {
    $confirmation = Read-Host "This will run $totalTests backtests which may take a long time. Continue? (Y/N)"
    if ($confirmation -ne 'Y') {
        Write-Host "Operation cancelled by user."
        exit
    }
}

# Run backtests for all symbols and timeframes found in the INI file
Start-MultiSymbolBacktestAutomation -Symbols $symbols -Timeframes $timeframes