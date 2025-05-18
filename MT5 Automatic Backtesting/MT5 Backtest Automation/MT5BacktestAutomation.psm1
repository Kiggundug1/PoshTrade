# Main module file that imports all component files
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path

# Import component scripts in the correct order
. "$ScriptPath\Private\Configuration.ps1"
. "$ScriptPath\Private\Logging.ps1"
. "$ScriptPath\Private\UIAutomation.ps1"
. "$ScriptPath\Private\SystemMonitoring.ps1"
. "$ScriptPath\Public\BacktestFunctions.ps1"

# Export public functions
Export-ModuleMember -Function Start-BacktestAutomation, Start-MultiSymbolBacktestAutomation