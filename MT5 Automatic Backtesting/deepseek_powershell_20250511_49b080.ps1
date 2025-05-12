# Create a scheduled task that runs at startup and restarts on failure
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-File `"D:\Scripts\MT5Backtest.ps1`""
$Trigger = New-ScheduledTaskTrigger -AtStartup
$Settings = New-ScheduledTaskSettingsSet -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 5) -DontStopOnIdleEnd
Register-ScheduledTask -TaskName "MT5 Backtest Automation" -Action $Action -Trigger $Trigger -Settings $Settings -RunLevel Highest