# System monitoring functions

function Get-SystemResources {
    [CmdletBinding()]
    param()
    
    # Get available memory
    $memoryInfo = Get-CimInstance -ClassName Win32_OperatingSystem
    $script:runtime.availableMemory = [math]::Round($memoryInfo.FreePhysicalMemory / 1024)
    
    # Only get CPU usage if memory is concerning (reduces overhead)
    if ($script:runtime.availableMemory -lt $script:config.lowMemoryThreshold * 2) {
        try {
            $cpuLoad = Get-CimInstance -ClassName Win32_Processor | Measure-Object -Property LoadPercentage -Average
            $script:runtime.cpuUsage = $cpuLoad.Average
        }
        catch {
            $script:runtime.cpuUsage = 50  # Default value if can't get actual CPU usage
            Write-Debug "Could not get CPU usage: $($_.Exception.Message)"
        }
    }
    else {
        # Assume moderate CPU usage if memory is plentiful
        $script:runtime.cpuUsage = 50
    }
}

function Invoke-DetailedSystemCheck {
    [CmdletBinding()]
    param()
    
    # Get current time
    $currentTime = [int](Get-Date).ToFileTime()
    
    # Only check periodically to avoid overhead
    if ($currentTime - $script:runtime.lastSystemLoadCheck -ge $script:config.detailedSystemCheckInterval) {
        $script:runtime.lastSystemLoadCheck = $currentTime
        
        # Get available memory and CPU usage
        Get-SystemResources
        
        # Create system metrics object
        $systemMetrics = @{
            "cpuUsage" = $script:runtime.cpuUsage
            "availableMemory" = $script:runtime.availableMemory
            "adaptiveMultiplier" = $script:runtime.currentAdaptiveMultiplier
        }
        
        # Only log if verbose logging is enabled or if system is under stress
        if ($script:config.verboseLogging -or 
            $script:runtime.availableMemory -lt $script:config.lowMemoryThreshold * 2 -or 
            $script:runtime.cpuUsage -gt 80) {
            Write-Information "System check: CPU $($script:runtime.cpuUsage)%, Memory $($script:runtime.availableMemory) MB"
            Write-Log -Level "INFO" -Message "System check" -Details $systemMetrics
        }
        
        # Adjust wait multiplier based on system metrics
        Update-AdaptiveMultiplier
    }
}

function Update-AdaptiveMultiplier {
    [CmdletBinding()]
    param()
    
    if ($script:runtime.cpuUsage -gt 90 -or $script:runtime.availableMemory -lt $script:config.lowMemoryThreshold) {
        # Critical system load - maximum wait times
        $script:runtime.currentAdaptiveMultiplier = $script:config.maxAdaptiveWaitMultiplier
        Write-Warning "Critical system load detected. Increasing wait times to maximum."
        Write-Log -Level "WARN" -Message "Critical system load detected. Increasing wait times to maximum." -Details @{
            "cpuUsage" = $script:runtime.cpuUsage
            "availableMemory" = $script:runtime.availableMemory
        }
        
        # Check if we need to perform memory cleanup
        Invoke-MemoryCleanup
    }
    elseif ($script:runtime.cpuUsage -gt 70 -or $script:runtime.availableMemory -lt $script:config.lowMemoryThreshold * 2) {
        # High system load - increase wait times
        $script:runtime.currentAdaptiveMultiplier = [Math]::Min(
            $script:config.maxAdaptiveWaitMultiplier, 
            $script:runtime.currentAdaptiveMultiplier * 1.5
        )
        
        # Only log if verbose logging is enabled
        if ($script:config.verboseLogging) {
            Write-Debug "High system load detected. Increasing wait multiplier to $($script:runtime.currentAdaptiveMultiplier)"
            Write-Log -Level "DEBUG" -Message "High system load detected. Increasing wait multiplier to $($script:runtime.currentAdaptiveMultiplier)" -Details @{
                "cpuUsage" = $script:runtime.cpuUsage
                "availableMemory" = $script:runtime.availableMemory
            }
        }
    }
    elseif ($script:runtime.cpuUsage -lt 40 -and $script:runtime.availableMemory -gt $script:config.lowMemoryThreshold * 3) {
        # Low system load - decrease wait times
        $script:runtime.currentAdaptiveMultiplier = [Math]::Max(
            1.0, 
            $script:runtime.currentAdaptiveMultiplier * 0.8
        )
        
        # Only log if verbose logging is enabled
        if ($script:config.verboseLogging) {
            Write-Debug "Low system load detected. Decreasing wait multiplier to $($script:runtime.currentAdaptiveMultiplier)"
            Write-Log -Level "DEBUG" -Message "Low system load detected. Decreasing wait multiplier to $($script:runtime.currentAdaptiveMultiplier)" -Details @{
                "cpuUsage" = $script:runtime.cpuUsage
                "availableMemory" = $script:runtime.availableMemory
            }
        }
    }
    else {
        # Moderate system load - gradually normalize wait times
        $script:runtime.currentAdaptiveMultiplier = [Math]::Max(
            1.0, 
            $script:runtime.currentAdaptiveMultiplier * 0.95
        )
    }
}

# Continuing from where we left off

function Invoke-MemoryCleanup {
    [CmdletBinding()]
    param()
    
    # Only perform cleanup when memory is critically low
    if ($script:runtime.availableMemory -lt ($script:config.lowMemoryThreshold / 2)) {
        Write-Warning "Performing memory cleanup due to low memory ($($script:runtime.availableMemory) MB)"
        Write-Log -Level "WARN" -Message "Performing memory cleanup due to low memory ($($script:runtime.availableMemory) MB)" -Details @{}
        
        try {
            # Attempt to free memory by restarting Explorer (lightweight cleanup)
            $explorerProcess = Get-Process -Name $script:constants.PROCESS_EXPLORER -ErrorAction SilentlyContinue
            if ($explorerProcess) {
                Stop-Process -Name $script:constants.PROCESS_EXPLORER -Force
                Invoke-AdaptiveWait -WaitTime $script:constants.WAIT_MEDIUM
                Start-Process $script:constants.PROCESS_EXPLORER
                Invoke-AdaptiveWait -WaitTime $script:constants.WAIT_LONG
                
                # Check if cleanup helped
                Get-SystemResources
                
                Write-Information "Memory after cleanup: $($script:runtime.availableMemory) MB"
                Write-Log -Level "INFO" -Message "Memory after cleanup: $($script:runtime.availableMemory) MB" -Details @{
                    "beforeCleanup" = $script:runtime.availableMemory
                    "afterCleanup" = $script:runtime.availableMemory
                }
            }
        }
        catch {
            Write-Error "Memory cleanup attempt failed: $($_.Exception.Message)"
            Write-Log -Level "ERROR" -Message "Memory cleanup attempt failed: $($_.Exception.Message)" -Details @{}
        }
        finally {
            # Force garbage collection
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    }
}

function Invoke-AdaptiveWait {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateRange(0, [int]::MaxValue)]
        [int]$WaitTime,
        
        [Parameter(Mandatory = $false)]
        [bool]$IsRetry = $false,
        
        [Parameter(Mandatory = $false)]
        [int]$RetryCount = 0
    )
    
    # Validate parameters
    if ($WaitTime -le 0) {
        $WaitTime = 1
    }
    
    # Calculate wait time based on parameters
    if ($IsRetry) {
        # Use exponential backoff for retries
        $backoffFactor = [Math]::Min($script:config.maxRetryWaitTime / $WaitTime, [Math]::Pow($script:config.retryBackoffMultiplier, $RetryCount))
        $adjustedWaitTime = $WaitTime * $backoffFactor
        
        # Cap at maximum retry wait time
        $adjustedWaitTime = [Math]::Min($adjustedWaitTime, $script:config.maxRetryWaitTime)
    }
    elseif ($script:config.adaptiveWaitEnabled) {
        # Use adaptive wait for normal operations
        $adjustedWaitTime = $WaitTime * $script:runtime.currentAdaptiveMultiplier
    }
    else {
        $adjustedWaitTime = $WaitTime
    }
    
    # Log wait time if it's significantly adjusted
    if ($adjustedWaitTime -gt $WaitTime * 1.5 -and $script:config.verboseLogging) {
        Write-Debug "Adjusted wait time from $WaitTime to $adjustedWaitTime seconds"
        Write-Log -Level "DEBUG" -Message "Adjusted wait time from $WaitTime to $adjustedWaitTime seconds" -Details @{}
    }
    
    # Perform the wait
    Start-Sleep -Seconds $adjustedWaitTime
}

function Initialize-PerformanceHistory {
    [CmdletBinding()]
    param()
    
    if (Test-Path -Path $script:config.performanceHistoryFile) {
        try {
            $historyData = Get-Content -Path $script:config.performanceHistoryFile -Raw
            $script:runtime.performanceHistory = $historyData | ConvertFrom-Json -AsHashtable
            Write-Information "Loaded performance history with $($script:runtime.performanceHistory.Count) entries"
            Write-Log -Level "INFO" -Message "Loaded performance history" -Details @{ "entries" = ($script:runtime.performanceHistory.Count) }
        }
        catch {
            Write-Warning "Error reading performance history file. Initializing new history."
            Write-Log -Level "WARN" -Message "Error reading performance history file. Initializing new history." -Details @{ "error" = $_.Exception.Message }
            $script:runtime.performanceHistory = @{}
        }
    }
    else {
        $script:runtime.performanceHistory = @{}
    }
}

function Update-PerformanceHistory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Currency,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Timeframe,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$EaName,
        
        [Parameter(Mandatory = $true)]
        [ValidateRange(1, [int]::MaxValue)]
        [int]$ActualDuration
    )
    
    # Validate parameters
    if ([string]::IsNullOrEmpty($Currency) -or 
        [string]::IsNullOrEmpty($Timeframe) -or 
        [string]::IsNullOrEmpty($EaName) -or 
        $ActualDuration -le 0) {
        
        Write-Warning "Invalid parameters for performance history update"
        Write-Log -Level "WARN" -Message "Invalid parameters for performance history update" -Details @{
            "currency" = $Currency
            "timeframe" = $Timeframe
            "eaName" = $EaName
            "duration" = $ActualDuration
        }
        return
    }
    
    # Create key for this combination
    $historyKey = "${EaName}_${Currency}_${Timeframe}"
    
    # Update or add the entry
    if ($script:runtime.performanceHistory.ContainsKey($historyKey)) {
        # Calculate weighted average (70% history, 30% new data)
        $historicalDuration = $script:runtime.performanceHistory[$historyKey]
        $newDuration = ($historicalDuration * 0.7) + ($ActualDuration * 0.3)
    }
    else {
        # First entry for this combination
        $newDuration = $ActualDuration
    }
    
    # Update the dictionary
    $script:runtime.performanceHistory[$historyKey] = $newDuration
    
    # Save to file
    try {
        $script:runtime.performanceHistory | ConvertTo-Json | Set-Content -Path $script:config.performanceHistoryFile
        Write-Debug "Updated performance history for $historyKey: $newDuration"
        Write-Log -Level "DEBUG" -Message "Updated performance history" -Details @{
            "combination" = $historyKey
            "duration" = $newDuration
        }
    }
    catch {
        Write-Error "Failed to save performance history: $($_.Exception.Message)"
        Write-Log -Level "ERROR" -Message "Failed to save performance history: $($_.Exception.Message)" -Details @{
            "historyKey" = $historyKey
        }
    }
}
