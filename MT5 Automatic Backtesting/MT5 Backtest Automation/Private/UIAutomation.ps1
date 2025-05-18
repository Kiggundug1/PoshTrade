# UI Automation functions

function Test-WindowExists {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$WindowTitle,
        
        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 100)]
        [int]$MaxAttempts = 1
    )
    
    $attempts = 0
    $windowFound = $false
    
    while ($attempts -lt $MaxAttempts -and -not $windowFound) {
        $windows = [System.Diagnostics.Process]::GetProcesses() | 
            Where-Object { $_.MainWindowTitle -like "*$WindowTitle*" -and $_.MainWindowHandle -ne 0 }
        
        $windowFound = ($windows -ne $null -and $windows.Count -gt 0)
        
        if (-not $windowFound -and $attempts -lt $MaxAttempts - 1) {
            Start-Sleep -Milliseconds 500
        }
        
        $attempts++
    }
    
    return $windowFound
}

function Set-WindowFocus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$WindowTitle,
        
        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 10)]
        [int]$MaxAttempts = 3
    )
    
    $attempts = 0
    $focusSet = $false
    
    while ($attempts -lt $MaxAttempts -and -not $focusSet) {
        $windows = [System.Diagnostics.Process]::GetProcesses() | 
            Where-Object { $_.MainWindowTitle -like "*$WindowTitle*" -and $_.MainWindowHandle -ne 0 }
        
        if ($windows -ne $null -and $windows.Count -gt 0) {
            # Get the first matching window
            $window = $windows[0]
            
            try {
                # Set focus to the window
                [void][Win32]::SetForegroundWindow($window.MainWindowHandle)
                
                # Ensure window is visible
                [void][Win32]::ShowWindow($window.MainWindowHandle, [Win32]::SW_RESTORE)
                
                # Verify window is visible and has focus
                if ([Win32]::IsWindowVisible($window.MainWindowHandle)) {
                    $focusSet = $true
                    
                    if ($script:config.verboseLogging) {
                        Write-Debug "Set focus to window: $WindowTitle"
                        Write-Log -Level "DEBUG" -Message "Set focus to window: $WindowTitle" -Details @{}
                    }
                    
                    break
                }
            }
            catch {
                Write-Debug "Error setting focus to window: $($_.Exception.Message)"
                Write-Log -Level "DEBUG" -Message "Error setting focus to window: $($_.Exception.Message)" -Details @{}
            }
        }
        
        if (-not $focusSet -and $attempts -lt $MaxAttempts - 1) {
            Start-Sleep -Milliseconds 500
        }
        
        $attempts++
    }
    
    return $focusSet
}

function Wait-ForWindow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$WindowTitle,
        
        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 300)]
        [int]$MaxWaitTime = 60
    )
    
    $waitTime = 0
    $windowAppeared = $false
    $checkInterval = 1
    
    while ($waitTime -lt $MaxWaitTime -and -not $windowAppeared) {
        if (Test-WindowExists -WindowTitle $WindowTitle) {
            $windowAppeared = $true
            Write-Information "Window '$WindowTitle' appeared successfully after $waitTime seconds"
            Write-Log -Level "INFO" -Message "Window '$WindowTitle' appeared successfully" -Details @{
                "waitTime" = $waitTime
            }
            break
        }
        
        Start-Sleep -Seconds $checkInterval
        $waitTime += $checkInterval
        
        # Increase check interval for longer waits to reduce CPU usage
        if ($waitTime -gt 10 -and $checkInterval -eq 1) {
            $checkInterval = 2
        }
    }
    
    if (-not $windowAppeared) {
        Write-Warning "Timed out waiting for window: $WindowTitle after $MaxWaitTime seconds"
        Write-Log -Level "WARN" -Message "Timed out waiting for window: $WindowTitle" -Details @{
            "maxWaitTime" = $MaxWaitTime
        }
    }
    
    return $windowAppeared
}

function Start-MT5WithBatchFile {
    [CmdletBinding()]
    param()
    
    try {
        Write-Information "Launching MT5 using batch file: $($script:config.batchFilePath)"
        Write-Log -Level "INFO" -Message "Launching MT5 using batch file: $($script:config.batchFilePath)" -Details @{}
        
        # Run the batch file
        $process = Start-Process -FilePath $script:config.batchFilePath -PassThru
        
        if (-not $process) {
            throw "Failed to start batch file process"
        }
        
        # Wait for MT5 to launch
        $mt5Launched = Wait-ForWindow -WindowTitle $script:constants.WINDOW_MT5 -MaxWaitTime $script:constants.MAX_LAUNCH_WAIT
        
        if (-not $mt5Launched) {
            throw "MetaTrader 5 did not launch within the expected time"
        }
        
        # Wait additional time for MT5 to fully initialize
        Invoke-AdaptiveWait -WaitTime $script:config.initialLoadTime
        
        return $true
    }
    catch {
        Write-Error "Failed to launch MT5 using batch file: $($_.Exception.Message)"
        Write-Log -Level "ERROR" -Message "Failed to launch MT5 using batch file: $($_.Exception.Message)" -Details @{
            "stackTrace" = $_.ScriptStackTrace
        }
        Capture-ErrorState -ErrorContext "LaunchMT5"
        return $false
    }
}

function Wait-ForStrategyTester {
    [CmdletBinding()]
    param()
    
    try {
        Write-Information "Waiting for Strategy Tester to open..."
        Write-Log -Level "INFO" -Message "Waiting for Strategy Tester to open..." -Details @{}
        
        # Wait for Strategy Tester window to appear
        $testerOpened = $false
        
        # Try both possible window names
        if (Wait-ForWindow -WindowTitle $script:constants.WINDOW_STRATEGY_TESTER -MaxWaitTime $script:constants.MAX_TESTER_WAIT) {
            $testerOpened = $true
        }
        elseif (Wait-ForWindow -WindowTitle $script:constants.WINDOW_TESTER -MaxWaitTime $script:constants.MAX_TESTER_WAIT) {
            $testerOpened = $true
        }
        
        if (-not $testerOpened) {
            throw "Strategy Tester did not open within the expected time"
        }
        
        # Give it a moment to fully initialize
        Invoke-AdaptiveWait -WaitTime $script:constants.WAIT_MEDIUM
        
        return $true
    }
    catch {
        Write-Error "Failed to wait for Strategy Tester: $($_.Exception.Message)"
        Write-Log -Level "ERROR" -Message "Failed to wait for Strategy Tester: $($_.Exception.Message)"
# Continuing from where we left off

        Write-Log -Level "ERROR" -Message "Failed to wait for Strategy Tester: $($_.Exception.Message)" -Details @{
            "stackTrace" = $_.ScriptStackTrace
        }
        Capture-ErrorState -ErrorContext "WaitForTester"
        return $false
    }
}

function Set-StrategyTesterFocus {
    [CmdletBinding()]
    param()
    
    $maxAttempts = 3
    $focusSet = $false
    
    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        if (Test-WindowExists -WindowTitle $script:constants.WINDOW_STRATEGY_TESTER) {
            $focusSet = Set-WindowFocus -WindowTitle $script:constants.WINDOW_STRATEGY_TESTER
            if ($focusSet) {
                break
            }
        }
        elseif (Test-WindowExists -WindowTitle $script:constants.WINDOW_TESTER) {
            $focusSet = Set-WindowFocus -WindowTitle $script:constants.WINDOW_TESTER
            if ($focusSet) {
                break
            }
        }
        
        if (-not $focusSet -and $attempt -lt $maxAttempts) {
            Invoke-AdaptiveWait -WaitTime $script:constants.WAIT_SHORT
        }
    }
    
    return $focusSet
}

function Start-Backtest {
    [CmdletBinding()]
    param()
    
    try {
        Write-Information "Starting backtest..."
        Write-Log -Level "INFO" -Message "Starting backtest..." -Details @{}
        
        # Ensure Strategy Tester window is active
        if (-not (Set-StrategyTesterFocus)) {
            throw "Strategy Tester window not found or could not set focus"
        }
        
        Invoke-AdaptiveWait -WaitTime $script:constants.WAIT_SHORT
        
        # Try different methods to start the test
        $testStarted = $false
        $methods = @(
            @{ "Name" = "F9 Key"; "Function" = { Start-BacktestWithKeyboard } },
            @{ "Name" = "Alt+S"; "Function" = { Start-BacktestWithAltKey } }
        )
        
        foreach ($method in $methods) {
            Write-Debug "Attempting to start backtest using method: $($method.Name)"
            Write-Log -Level "DEBUG" -Message "Attempting to start backtest using method: $($method.Name)" -Details @{}
            
            $testStarted = & $method.Function
            
            if ($testStarted) {
                Write-Information "Backtest started successfully using method: $($method.Name)"
                Write-Log -Level "INFO" -Message "Backtest started successfully using method: $($method.Name)" -Details @{}
                break
            }
            
            Invoke-AdaptiveWait -WaitTime $script:constants.WAIT_SHORT
        }
        
        if (-not $testStarted) {
            throw "Failed to start backtest using all available methods"
        }
        
        return $true
    }
    catch {
        Write-Error "Failed to start backtest: $($_.Exception.Message)"
        Write-Log -Level "ERROR" -Message "Failed to start backtest: $($_.Exception.Message)" -Details @{
            "stackTrace" = $_.ScriptStackTrace
        }
        Capture-ErrorState -ErrorContext "StartBacktest"
        return $false
    }
}

function Start-BacktestWithKeyboard {
    [CmdletBinding()]
    param()
    
    try {
        # Try using keyboard shortcut first (F9)
        Set-StrategyTesterFocus
        [System.Windows.Forms.SendKeys]::SendWait("{F9}")
        Invoke-AdaptiveWait -WaitTime $script:constants.WAIT_MEDIUM
        
        # Check if test started by looking for progress indicators
        $testStarted = Confirm-BacktestStarted
        
        return $testStarted
    }
    catch {
        Write-Debug "Error starting backtest with F9: $($_.Exception.Message)"
        Write-Log -Level "DEBUG" -Message "Error starting backtest with F9: $($_.Exception.Message)" -Details @{}
        return $false
    }
}

function Start-BacktestWithAltKey {
    [CmdletBinding()]
    param()
    
    try {
        # Last resort - try Alt+S for Start
        Set-StrategyTesterFocus
        [System.Windows.Forms.SendKeys]::SendWait("%s")
        Invoke-AdaptiveWait -WaitTime $script:constants.WAIT_MEDIUM
        
        # Check if test started
        $testStarted = Confirm-BacktestStarted
        
        return $testStarted
    }
    catch {
        Write-Debug "Error starting backtest with Alt+S: $($_.Exception.Message)"
        Write-Log -Level "DEBUG" -Message "Error starting backtest with Alt+S: $($_.Exception.Message)" -Details @{}
        return $false
    }
}

function Confirm-BacktestStarted {
    [CmdletBinding()]
    param()
    
    try {
        # In a real implementation, you would use UI Automation to check for:
        # 1. Start button being disabled
        # 2. Progress bar appearing
        # 3. Status text changing
        
        # For this simplified version, we'll wait and assume it started
        # A more robust implementation would use UI Automation framework
        
        # Wait a moment to see if the test starts
        Invoke-AdaptiveWait -WaitTime $script:constants.WAIT_MEDIUM
        
        # For now, we'll just return true and rely on the monitoring function
        # to detect if the test actually started
        return $true
    }
    catch {
        Write-Debug "Error confirming backtest started: $($_.Exception.Message)"
        Write-Log -Level "DEBUG" -Message "Error confirming backtest started: $($_.Exception.Message)" -Details @{}
        return $false
    }
}

function Monitor-BacktestProgress {
    [CmdletBinding()]
    param()
    
    try {
        Write-Information "Monitoring backtest progress..."
        Write-Log -Level "INFO" -Message "Monitoring backtest progress..." -Details @{}
        
        # Initialize monitoring variables
        $script:runtime.testStartTime = [int](Get-Date).ToFileTime()
        $testCompleted = $false
        $testWaitTime = 0
        $lastProgressValue = "0"
        $noProgressCounter = 0
        $mtFrozenCounter = 0
        $script:runtime.previousLoggedProgress = "0"
        $script:runtime.lastProgressLogTime = 0
        
        # Main monitoring loop
        while (-not $testCompleted) {
            # Check for test completion using multiple methods
            $testCompleted = Test-BacktestCompletion
            
            if ($testCompleted) {
                break
            }
            
            # Check for progress changes to detect if test is still running
            Update-TestProgress -LastProgressValue ([ref]$lastProgressValue) -NoProgressCounter ([ref]$noProgressCounter)
            
            # Check system resources periodically during the test
            if ($testWaitTime % $script:constants.SYSTEM_CHECK_INTERVAL -eq 0) {
                Invoke-DetailedSystemCheck
            }
            
            # Check if MT5 is responsive
            if ($testWaitTime % 60 -eq 0 -and $testWaitTime -gt 0) {
                Test-MT5Responsiveness -MtFrozenCounter ([ref]$mtFrozenCounter) -TestCompleted ([ref]$testCompleted)
            }
            
            # Check if test is stuck with no progress
            if ($noProgressCounter -ge $script:constants.MAX_NO_PROGRESS_INTERVALS) {
                Write-Error "Backtest appears to be stuck at $lastProgressValue%. Attempting recovery..."
                Write-Log -Level "ERROR" -Message "Backtest appears to be stuck at $lastProgressValue%. Attempting recovery..." -Details @{
                    "noProgressIntervals" = $noProgressCounter
                    "maxAllowedIntervals" = $script:constants.MAX_NO_PROGRESS_INTERVALS
                }
                Capture-ErrorState -ErrorContext "BacktestStuck"
                $testCompleted = $true
                $script:runtime.consecutiveFailures++
                break
            }
            
            # Periodic heartbeat log
            if ($testWaitTime % 300 -eq 0 -and $testWaitTime -gt 0) {
                Write-Information "Backtest still running after $testWaitTime seconds. Current progress: $lastProgressValue%"
                Write-Log -Level "INFO" -Message "Backtest still running after $testWaitTime seconds. Current progress: $lastProgressValue%" -Details @{
                    "elapsedTime" = $testWaitTime
                    "progress" = $lastProgressValue
                }
            }
            
            Invoke-AdaptiveWait -WaitTime $script:constants.PROGRESS_CHECK_INTERVAL
            $testWaitTime += $script:constants.PROGRESS_CHECK_INTERVAL
            
            # Safety timeout - don't wait forever
            if ($testWaitTime -gt $script:config.maxWaitTimeForTest * 10) {
                Write-Error "Maximum wait time exceeded. Forcing test completion."
                Write-Log -Level "ERROR" -Message "Maximum wait time exceeded. Forcing test completion." -Details @{
                    "maxWaitTime" = $script:config.maxWaitTimeForTest * 10
                    "elapsedTime" = $testWaitTime
                }
                Capture-ErrorState -ErrorContext "TimeoutExceeded"
                $testCompleted = $true
                $script:runtime.consecutiveFailures++
                break
            }
        }
        
        # Record actual test duration for future estimates
        $actualTestDuration = [int](Get-Date).ToFileTime() - $script:runtime.testStartTime
        Update-PerformanceHistory -Currency $script:runtime.currency -Timeframe $script:runtime.timeframe -EaName $script:config.eaName -ActualDuration $actualTestDuration
        
        # Return success if test completed normally
        if ($script:runtime.consecutiveFailures -eq 0) {
            return $true
        }
        else {
            return $false
        }
    }
    catch {
        Write-Error "Error monitoring backtest progress: $($_.Exception.Message)"
        Write-Log -Level "ERROR" -Message "Error monitoring backtest progress: $($_.Exception.Message)" -Details @{
            "stackTrace" = $_.ScriptStackTrace
        }
        Capture-ErrorState -ErrorContext "MonitorProgress"
        return $false
    }
}

# Additional UI Automation functions would continue here...
