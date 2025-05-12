# Script to automate MT5 Strategy Tester
# Note: This script uses UI automation which may be fragile and require adjustments

# Path to MetaTrader 5
$mt5Path = "C:\Program Files\MetaTrader 5 EXNESS\terminal64.exe"
$eaPath = "C:\Users\GodfreyKiggundu\AppData\Roaming\MetaQuotes\Terminal\53785E099C927DB68A545C249CDBCE06\MQL5\Experts\Custom EAs\Moving Average EURUSD M15.ex5"

# Check if MT5 is already running
$mt5Process = Get-Process -Name "terminal64" -ErrorAction SilentlyContinue

# Launch MT5 if not already running
if ($null -eq $mt5Process) {
    Write-Host "Starting MetaTrader 5..."
    Start-Process -FilePath $mt5Path
    # Wait for MT5 to initialize
    Start-Sleep -Seconds 10
} else {
    Write-Host "MetaTrader 5 is already running."
}

# Load Windows Forms for sending keystrokes
Add-Type -AssemblyName System.Windows.Forms

# Function to send keystrokes with delay
function Send-KeysWithDelay {
    param (
        [string]$keys,
        [int]$delayMs = 500
    )
    [System.Windows.Forms.SendKeys]::SendWait($keys)
    Start-Sleep -Milliseconds $delayMs
}

# Activate MT5 window
$mt5Process = Get-Process -Name "terminal64" -ErrorAction SilentlyContinue
if ($null -ne $mt5Process) {
    # Bring MT5 to foreground
    $wshell = New-Object -ComObject wscript.shell
    $wshell.AppActivate($mt5Process.MainWindowTitle)
    Start-Sleep -Seconds 2
    
    # Open Strategy Tester (Ctrl+R)
    Write-Host "Opening Strategy Tester..."
    Send-KeysWithDelay "^r" 2000
    
    # Tab to Expert Advisor selection and open the dropdown
    Write-Host "Configuring Strategy Tester..."
    Send-KeysWithDelay "{TAB}" 500
    Send-KeysWithDelay " " 500
    
    # This part is tricky - we need to navigate to the specific EA
    # Since we can't directly input the path, we'll try to search for it
    $eaName = "Moving Average EURUSD M15"
    foreach ($char in $eaName.ToCharArray()) {
        Send-KeysWithDelay $char 100
    }
    Send-KeysWithDelay "{ENTER}" 1000
    
    # Tab to Symbol field and set to EURUSDm
    Send-KeysWithDelay "{TAB}" 500
    Send-KeysWithDelay " " 500
    Send-KeysWithDelay "EURUSDm" 500
    Send-KeysWithDelay "{ENTER}" 1000
    
    # Tab to Period field and set to M15
    Send-KeysWithDelay "{TAB}" 500
    Send-KeysWithDelay " " 500
    Send-KeysWithDelay "M15" 500
    Send-KeysWithDelay "{ENTER}" 1000
    
    # Tab to Model field (usually already set)
    Send-KeysWithDelay "{TAB}" 500
    
    # Tab to Date range fields
    Send-KeysWithDelay "{TAB}" 500
    # Set From date
    Send-KeysWithDelay "2019.01.01" 500
    Send-KeysWithDelay "{TAB}" 500
    # Set To date
    Send-KeysWithDelay "2024.12.31" 500
    
    # Tab to Deposit field
    Send-KeysWithDelay "{TAB}" 500
    Send-KeysWithDelay "10000" 500
    
    # Tab to Currency field
    Send-KeysWithDelay "{TAB}" 500
    Send-KeysWithDelay "USD" 500
    
    # Tab to Leverage field
    Send-KeysWithDelay "{TAB}" 500
    Send-KeysWithDelay "1:2000" 500
    
    # Enable Visual Mode (usually a checkbox)
    # Tab to Visual Mode checkbox
    Send-KeysWithDelay "{TAB}{TAB}{TAB}" 500
    # Space to check it
    Send-KeysWithDelay " " 500
    
    # Start the test (usually a Start button)
    # Tab to Start button
    Send-KeysWithDelay "{TAB}{TAB}" 500
    Send-KeysWithDelay " " 500
    
    Write-Host "Backtest started. Waiting for completion..."
    # This is the most challenging part - we need to wait for the backtest to complete
    # There's no reliable way to detect this via UI automation
    # We'll wait for a reasonable time based on the backtest period
    
    # Approximate wait time based on date range (adjust as needed)
    $waitTimeMinutes = 10
    Write-Host "Waiting $waitTimeMinutes minutes for backtest to complete..."
    Start-Sleep -Seconds ($waitTimeMinutes * 60)
    
    # After waiting, try to save the report
    # Right-click in the report area to bring up context menu
    Write-Host "Attempting to save report..."
    Send-KeysWithDelay "+{F10}" 1000  # Shift+F10 for context menu
    
    # Navigate to "Save Report" or similar option
    # This will vary based on MT5 version and language
    Send-KeysWithDelay "s" 500  # First letter of "Save" option
    Send-KeysWithDelay "{ENTER}" 1000
    
    # In the save dialog, navigate to save as XML
    # Tab to file type dropdown
    Send-KeysWithDelay "{TAB}{TAB}{TAB}{TAB}" 500
    Send-KeysWithDelay " " 500
    # Select XML format
    Send-KeysWithDelay "x" 500  # First letter of "XML"
    Send-KeysWithDelay "{ENTER}" 500
    
    # Save the file
    Send-KeysWithDelay "{ENTER}" 1000
    
    Write-Host "Backtest completed and report saved (if successful)."
    Write-Host "Note: UI automation is fragile - please verify the results manually."
} else {
    Write-Host "Failed to find MetaTrader 5 process."
}