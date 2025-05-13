$taskName = "MT5BacktestAutomation"
$scriptPath = "C:\Path\To\Your\Script.bat"  # Replace with your actual script or shortcut path

$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -WindowStyle Hidden -File `"$scriptPath`""
$trigger = New-ScheduledTaskTrigger -AtStartup
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

Register-ScheduledTask -Action $action -Trigger $trigger -Settings $settings -TaskName $taskName -Description "Launch MT5 Automation on startup" -User "SYSTEM" -RunLevel Highest