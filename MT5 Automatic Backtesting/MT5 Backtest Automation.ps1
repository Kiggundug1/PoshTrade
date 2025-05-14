# MT5 Backtest Automation PowerShell Script
# Based on Backtest flow script V1.1.5

# Add required assemblies for UI automation
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

# Initialize variables with default configuration
$defaultConfig = @{
    "mt5Path" = "C:\Program Files\MetaTrader 5 EXNESS\terminal64.exe"
    "eaPath" = "C:\Users\kigundu\AppData\Roaming\MetaQuotes\Terminal\53785E099C927DB68A545C249CDBCE06\MQL5\Experts\Custom EAs\Moving Average"
    "reportPath" = "D:\FOREX\FOREX DOCUMENTS\MT5 STRATEGY TESTER REPORTS\Reports"
    "startDate" = "2019.01.01"
    "endDate" = "2024.12.31"
    "reportCounter" = 1
    "maxWaitTimeForTest" = 180
    "initialLoadTime" = 15
    "maxRetries" = 3
    "skipOnError" = $true
    "autoRestartOnFailure" = $true
    "maxConsecutiveFailures" = 5
    "adaptiveWaitEnabled" = $true
    "baseWaitMultiplier" = 1.0
    "maxAdaptiveWaitMultiplier" = 5
    "systemLoadCheckInterval" = 300
    "lowMemoryThreshold" = 200
    "verboseLogging" = $false
    "logProgressInterval" = 10
    "detailedSystemCheckInterval" = 600
    "logFilePath" = "D:\FOREX\FOREX DOCUMENTS\MT5 STRATEGY TESTER REPORTS\automation_log.json"
    "errorScreenshotsPath" = "D:\FOREX\FOREX DOCUMENTS\MT5 STRATEGY TESTER REPORTS\Reports\errors"
    "checkpointFile" = "D:\FOREX\FOREX DOCUMENTS\MT5 STRATEGY TESTER REPORTS\Reports\checkpoint.json"
    "configFilePath" = "D:\FOREX\FOREX DOCUMENTS\MT5 STRATEGY TESTER REPORTS\Reports\backtest_config.json"
    "performanceHistoryFile" = "D:\FOREX\FOREX DOCUMENTS\MT5 STRATEGY TESTER REPORTS\Reports\performance_history.json"
    "retryBackoffMultiplier" = 1.5
    "maxRetryWaitTime" = 60
}

# Initialize variables with default values
$mt5Path = $defaultConfig.mt5Path
$eaPath = $defaultConfig.eaPath
$reportPath = $defaultConfig.reportPath
$startDate = $defaultConfig.startDate
$endDate = $defaultConfig.endDate
$reportCounter = $defaultConfig.reportCounter
$maxWaitTimeForTest = $defaultConfig.maxWaitTimeForTest
$initialLoadTime = $defaultConfig.initialLoadTime
$maxRetries = $defaultConfig.maxRetries
$skipOnError = $defaultConfig.skipOnError
$autoRestartOnFailure = $defaultConfig.autoRestartOnFailure
$maxConsecutiveFailures = $defaultConfig.maxConsecutiveFailures
$consecutiveFailures = 0
$adaptiveWaitEnabled = $defaultConfig.adaptiveWaitEnabled
$baseWaitMultiplier = $defaultConfig.baseWaitMultiplier
$maxAdaptiveWaitMultiplier = $defaultConfig.maxAdaptiveWaitMultiplier
$currentAdaptiveMultiplier = 1.0
$systemLoadCheckInterval = $defaultConfig.systemLoadCheckInterval
$lastSystemLoadCheck = 0
$lowMemoryThreshold = $defaultConfig.lowMemoryThreshold
$availableMemory = 1000
$verboseLogging = $defaultConfig.verboseLogging
$logProgressInterval = $defaultConfig.logProgressInterval
$detailedSystemCheckInterval = $defaultConfig.detailedSystemCheckInterval
$logFilePath = $defaultConfig.logFilePath
$errorScreenshotsPath = $defaultConfig.errorScreenshotsPath
$checkpointFile = $defaultConfig.checkpointFile
$eaIndex = 0
$currencyIndex = 0
$timeframeIndex = 0
$resumeFromCheckpoint = $false
$configFilePath = $defaultConfig.configFilePath
$performanceHistoryFile = $defaultConfig.performanceHistoryFile
$retryBackoffMultiplier = $defaultConfig.retryBackoffMultiplier
$maxRetryWaitTime = $defaultConfig.maxRetryWaitTime

# Check if paths exist
if (-not (Test-Path $mt5Path)) {
    Write-Error "MetaTrader 5 path does not exist: $mt5Path"
    if (-not $skipOnError) { exit }
}

if (-not (Test-Path $eaPath)) {
    Write-Error "EA path does not exist: $eaPath"
    if (-not $skipOnError) { exit }
}

if (-not (Test-Path $reportPath)) {
    # Create reports folder if it doesn't exist
    New-Item -Path $reportPath -ItemType Directory -Force
}

if (-not (Test-Path $errorScreenshotsPath)) {
    New-Item -Path $errorScreenshotsPath -ItemType Directory -Force
}

# Initialize performance history data structure
$performanceHistory = @{}
if (Test-Path $performanceHistoryFile) {
    try {
        $historyData = Get-Content -Path $performanceHistoryFile -Raw
        $performanceHistory = $historyData | ConvertFrom-Json -AsHashtable
    }
    catch {
        # If file exists but can't be parsed, initialize empty
        Add-Content -Path $logFilePath -Value "Error reading performance history file. Initializing new history."
    }
}

# Log start of execution with structured format
$logEntry = @{
    "timestamp" = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    "level" = "INFO"
    "message" = "Started execution"
    "details" = @{
        "mt5Path" = $mt5Path
        "eaPath" = $eaPath
        "reportPath" = $reportPath
    }
}
Add-Content -Path $logFilePath -Value ($logEntry | ConvertTo-Json -Compress)

# Function to log messages in structured format
function LogMessage {
    param (
        [string]$level,
        [string]$message,
        $details = $null
    )
    
    $logEntry = @{
        "timestamp" = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        "level" = $level
        "message" = $message
    }
    
    # Add details if provided
    if ($details -ne $null) {
        $logEntry.details = $details
    }
    
    # Only log if verbose mode is on or if it's an important message
    if ($verboseLogging -or $level -eq "ERROR" -or $level -eq "WARN" -or $level -eq "INFO") {
        Add-Content -Path $logFilePath -Value ($logEntry | ConvertTo-Json -Compress)
    }
}

# Function to verify and set a value only if needed
function VerifyAndSetValue {
    param (
        [string]$fieldName,
        [string]$currentValue,
        [string]$targetValue
    )
    
    if ($currentValue -ne $targetValue) {
        LogMessage -level "DEBUG" -message "Changing $fieldName from '$currentValue' to '$targetValue'"
        return $false  # Value needs to be set
    }
    else {
        LogMessage -level "DEBUG" -message "$fieldName already set to '$targetValue', skipping"
        return $true  # Value already correct
    }
}

# Function to load configuration
function LoadConfiguration {
    if (Test-Path $configFilePath) {
        try {
            LogMessage -level "INFO" -message "Loading configuration from $configFilePath"
            $configData = Get-Content -Path $configFilePath -Raw
            
            # Try to parse as JSON
            try {
                $config = $configData | ConvertFrom-Json
                
                # Apply configuration from JSON object
                if ($config.PSObject.Properties.Name -contains "mt5Path") { $script:mt5Path = $config.mt5Path }
                if ($config.PSObject.Properties.Name -contains "eaPath") { $script:eaPath = $config.eaPath }
                if ($config.PSObject.Properties.Name -contains "reportPath") { $script:reportPath = $config.reportPath }
                if ($config.PSObject.Properties.Name -contains "startDate") { $script:startDate = $config.startDate }
                if ($config.PSObject.Properties.Name -contains "endDate") { $script:endDate = $config.endDate }
                if ($config.PSObject.Properties.Name -contains "maxWaitTimeForTest") { $script:maxWaitTimeForTest = $config.maxWaitTimeForTest }
                if ($config.PSObject.Properties.Name -contains "initialLoadTime") { $script:initialLoadTime = $config.initialLoadTime }
                if ($config.PSObject.Properties.Name -contains "maxRetries") { $script:maxRetries = $config.maxRetries }
                if ($config.PSObject.Properties.Name -contains "skipOnError") { $script:skipOnError = $config.skipOnError }
                if ($config.PSObject.Properties.Name -contains "autoRestartOnFailure") { $script:autoRestartOnFailure = $config.autoRestartOnFailure }
                if ($config.PSObject.Properties.Name -contains "maxConsecutiveFailures") { $script:maxConsecutiveFailures = $config.maxConsecutiveFailures }
                if ($config.PSObject.Properties.Name -contains "adaptiveWaitEnabled") { $script:adaptiveWaitEnabled = $config.adaptiveWaitEnabled }
                if ($config.PSObject.Properties.Name -contains "baseWaitMultiplier") { $script:baseWaitMultiplier = $config.baseWaitMultiplier }
                if ($config.PSObject.Properties.Name -contains "maxAdaptiveWaitMultiplier") { $script:maxAdaptiveWaitMultiplier = $config.maxAdaptiveWaitMultiplier }
                if ($config.PSObject.Properties.Name -contains "systemLoadCheckInterval") { $script:systemLoadCheckInterval = $config.systemLoadCheckInterval }
                if ($config.PSObject.Properties.Name -contains "lowMemoryThreshold") { $script:lowMemoryThreshold = $config.lowMemoryThreshold }
                if ($config.PSObject.Properties.Name -contains "verboseLogging") { $script:verboseLogging = $config.verboseLogging }
                if ($config.PSObject.Properties.Name -contains "logProgressInterval") { $script:logProgressInterval = $config.logProgressInterval }
                if ($config.PSObject.Properties.Name -contains "detailedSystemCheckInterval") { $script:detailedSystemCheckInterval = $config.detailedSystemCheckInterval }
                if ($config.PSObject.Properties.Name -contains "retryBackoffMultiplier") { $script:retryBackoffMultiplier = $config.retryBackoffMultiplier }
                if ($config.PSObject.Properties.Name -contains "maxRetryWaitTime") { $script:maxRetryWaitTime = $config.maxRetryWaitTime }
                
                LogMessage -level "INFO" -message "Configuration loaded successfully from JSON"
            }
            catch {
                # Fallback to legacy text format parsing
                LogMessage -level "WARN" -message "Failed to parse JSON config, falling back to text format"
                
                # Parse config data line by line
                $configLines = $configData -split "`r`n"
                
                foreach ($configLine in $configLines) {
                    # Skip empty lines and comments
                    if ([string]::IsNullOrWhiteSpace($configLine) -or $configLine.StartsWith("#")) {
                        continue
                    }
                    
                    # Extract key and value
                    $keyValue = $configLine -split "=", 2
                    if ($keyValue.Length -eq 2) {
                        $configKey = $keyValue[0].Trim()
                        $configValue = $keyValue[1].Trim()
                        
                        # Apply configuration based on key
                        switch ($configKey) {
                            "mt5Path" { $script:mt5Path = $configValue }
                            "eaPath" { $script:eaPath = $configValue }
                            "reportPath" { $script:reportPath = $configValue }
                            "startDate" { $script:startDate = $configValue }
                            "endDate" { $script:endDate = $configValue }
                            "maxWaitTimeForTest" { $script:maxWaitTimeForTest = [int]$configValue }
                            "initialLoadTime" { $script:initialLoadTime = [int]$configValue }
                            "maxRetries" { $script:maxRetries = [int]$configValue }
                            "skipOnError" { $script:skipOnError = [bool]::Parse($configValue) }
                            "autoRestartOnFailure" { $script:autoRestartOnFailure = [bool]::Parse($configValue) }
                            "maxConsecutiveFailures" { $script:maxConsecutiveFailures = [int]$configValue }
                            "adaptiveWaitEnabled" { $script:adaptiveWaitEnabled = [bool]::Parse($configValue) }
                            "baseWaitMultiplier" { $script:baseWaitMultiplier = [double]$configValue }
                            "maxAdaptiveWaitMultiplier" { $script:maxAdaptiveWaitMultiplier = [double]$configValue }
                            "systemLoadCheckInterval" { $script:systemLoadCheckInterval = [int]$configValue }
                            "lowMemoryThreshold" { $script:lowMemoryThreshold = [int]$configValue }
                            "verboseLogging" { $script:verboseLogging = [bool]::Parse($configValue) }
                            "logProgressInterval" { $script:logProgressInterval = [int]$configValue }
                            "detailedSystemCheckInterval" { $script:detailedSystemCheckInterval = [int]$configValue }
                            "retryBackoffMultiplier" { $script:retryBackoffMultiplier = [double]$configValue }
                            "maxRetryWaitTime" { $script:maxRetryWaitTime = [int]$configValue }
                        }
                        
                        LogMessage -level "DEBUG" -message "Config: $configKey = $configValue"
                    }
                }
                
                LogMessage -level "INFO" -message "Configuration loaded successfully from text format"
            }
        }
        catch {
            LogMessage -level "ERROR" -message "Error loading configuration: $($_.Exception.Message). Using default settings."
        }
    }
    else {
        LogMessage -level "INFO" -message "No configuration file found at $configFilePath. Using default settings."
    }
}

# Define adaptive wait function with exponential backoff for retries
function AdaptiveWait {
    param (
        [double]$waitTime,
        [bool]$isRetry = $false,
        [int]$retryCount = 0
    )
    
    # Calculate wait time based on parameters
    if ($isRetry) {
                # Use exponential backoff for retries
        $backoffFactor = [Math]::Min($maxRetryWaitTime / $waitTime, [Math]::Pow($retryBackoffMultiplier, $retryCount))
        $adjustedWaitTime = $waitTime * $backoffFactor
        
        # Cap at maximum retry wait time
        $adjustedWaitTime = [Math]::Min($adjustedWaitTime, $maxRetryWaitTime)
    }
    elseif ($adaptiveWaitEnabled) {
        # Use adaptive wait for normal operations
        $adjustedWaitTime = $waitTime * $currentAdaptiveMultiplier
    }
    else {
        $adjustedWaitTime = $waitTime
    }
    
    # Log wait time if it's significantly adjusted
    if ($adjustedWaitTime -gt $waitTime * 1.5 -and $verboseLogging) {
        LogMessage -level "DEBUG" -message "Adjusted wait time from $waitTime to $adjustedWaitTime seconds"
    }
    
    Start-Sleep -Seconds $adjustedWaitTime
}

# Simplified version for backward compatibility
function LegacyAdaptiveWait {
    param (
        [double]$waitTime
    )
    
    AdaptiveWait -waitTime $waitTime -isRetry $false -retryCount 0
}

# Function to capture error state with screenshots
function CaptureErrorState {
    param (
        [string]$errorContext
    )
    
    try {
        # Format timestamp for filename
        $timestamp = (Get-Date -Format "yyyy-MM-dd_HH-mm-ss")
        
        # Take screenshot of error state
        $screenshotPath = Join-Path $errorScreenshotsPath "error_${errorContext}_${timestamp}.png"
        
        # Create a bitmap of the screen
        $screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
        $bitmap = New-Object System.Drawing.Bitmap $screen.Width, $screen.Height
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $graphics.CopyFromScreen($screen.X, $screen.Y, 0, 0, $screen.Size)
        $bitmap.Save($screenshotPath, [System.Drawing.Imaging.ImageFormat]::Png)
        $graphics.Dispose()
        $bitmap.Dispose()
        
        $screenshotDetails = @{
            "filename" = "error_${errorContext}_${timestamp}.png"
            "context" = $errorContext
            "timestamp" = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        }
        
        LogMessage -level "INFO" -message "Error screenshot saved" -details $screenshotDetails
        
        # Try to save any partial results if in Strategy Tester
        $strategyTesterWindow = FindWindow -windowTitle "Strategy Tester"
        if ($strategyTesterWindow -ne $null) {
            # Send Ctrl+S to save
            SendKeysToWindow -window $strategyTesterWindow -keys "^s"
            LegacyAdaptiveWait -waitTime 2
            
            # Set partial results filename
            $partialFileName = "partial_${errorContext}_${timestamp}"
            $saveAsWindow = FindWindow -windowTitle "Save As"
            if ($saveAsWindow -ne $null) {
                SendKeysToWindow -window $saveAsWindow -keys "$reportPath\$partialFileName"
                LegacyAdaptiveWait -waitTime 1
                SendKeysToWindow -window $saveAsWindow -keys "{ENTER}"
                LegacyAdaptiveWait -waitTime 2
                
                $partialResultDetails = @{
                    "filename" = $partialFileName
                    "context" = $errorContext
                    "timestamp" = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                }
                
                LogMessage -level "INFO" -message "Partial results saved" -details $partialResultDetails
            }
        }
    }
    catch {
        LogMessage -level "ERROR" -message "Failed to capture error state: $($_.Exception.Message)"
    }
}

# Function to save checkpoint with improved JSON structure
function SaveCheckpoint {
    try {
        # Create checkpoint data as a proper JSON object
        $checkpointData = @{
            "eaIndex" = $eaIndex
            "currencyIndex" = $currencyIndex
            "timeframeIndex" = $timeframeIndex
            "eaName" = $eaName
            "currency" = $currency
            "timeframe" = $timeframe
            "reportCounter" = $reportCounter
            "timestamp" = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        }
        
        # Write checkpoint data to file
        $checkpointData | ConvertTo-Json | Set-Content -Path $checkpointFile
        
        $checkpointDetails = @{
            "eaName" = $eaName
            "currency" = $currency
            "timeframe" = $timeframe
        }
        
        LogMessage -level "INFO" -message "Checkpoint saved" -details $checkpointDetails
    }
    catch {
        LogMessage -level "ERROR" -message "Failed to save checkpoint: $($_.Exception.Message)"
    }
}

# Function to perform memory cleanup
function PerformMemoryCleanup {
    # Only perform cleanup when memory is critically low
    if ($availableMemory -lt ($lowMemoryThreshold / 2)) {
        LogMessage -level "WARN" -message "Performing memory cleanup due to low memory ($availableMemory MB)"
        
        # Attempt to free memory by restarting Explorer (lightweight cleanup)
        try {
            Stop-Process -Name "explorer" -Force
            LegacyAdaptiveWait -waitTime 3
            Start-Process "explorer.exe"
            LegacyAdaptiveWait -waitTime 5
            
            # Check if cleanup helped
            $newAvailableMemory = (Get-Counter '\Memory\Available MBytes').CounterSamples.CookedValue
            
            $cleanupDetails = @{
                "beforeCleanup" = $availableMemory
                "afterCleanup" = $newAvailableMemory
                "improvement" = $newAvailableMemory - $availableMemory
            }
            
            LogMessage -level "INFO" -message "Memory after cleanup: $newAvailableMemory MB (was $availableMemory MB)" -details $cleanupDetails
            $script:availableMemory = $newAvailableMemory
        }
        catch {
            LogMessage -level "ERROR" -message "Memory cleanup attempt failed: $($_.Exception.Message)"
        }
    }
}

# Enhanced system resource check function - optimized for long runs
function DetailedSystemCheck {
    # Get current time
    $currentTime = [int](Get-Date -UFormat %s)
    
    # Only check periodically to avoid overhead
    if ($currentTime - $lastSystemLoadCheck -ge $detailedSystemCheckInterval) {
        $script:lastSystemLoadCheck = $currentTime
        
        # Get available memory
        $script:availableMemory = (Get-Counter '\Memory\Available MBytes').CounterSamples.CookedValue
        
        # Only get CPU usage if memory is concerning (reduces overhead)
        if ($availableMemory -lt $lowMemoryThreshold * 2) {
            try {
                $cpuUsage = (Get-Counter '\Processor(_Total)\% Processor Time').CounterSamples.CookedValue
            }
            catch {
                $cpuUsage = 50  # Default value if can't get actual CPU usage
            }
        }
        else {
            # Assume moderate CPU usage if memory is plentiful
            $cpuUsage = 50
        }
        
        # Create system metrics object
        $systemMetrics = @{
            "cpuUsage" = $cpuUsage
            "availableMemory" = $availableMemory
            "adaptiveMultiplier" = $currentAdaptiveMultiplier
        }
        
        # Only log if verbose logging is enabled or if system is under stress
        if ($verboseLogging -or $availableMemory -lt $lowMemoryThreshold * 2 -or $cpuUsage -gt 80) {
            LogMessage -level "INFO" -message "System check" -details $systemMetrics
        }
        
        # Adjust wait multiplier based on system metrics
        if ($cpuUsage -gt 90 -or $availableMemory -lt $lowMemoryThreshold) {
            # Critical system load - maximum wait times
            $script:currentAdaptiveMultiplier = $maxAdaptiveWaitMultiplier
            LogMessage -level "WARN" -message "Critical system load detected. Increasing wait times to maximum." -details $systemMetrics
            
            # Check if we need to perform memory cleanup
            PerformMemoryCleanup
        }
        elseif ($cpuUsage -gt 70 -or $availableMemory -lt $lowMemoryThreshold * 2) {
            # High system load - increase wait times
            $script:currentAdaptiveMultiplier = [Math]::Min($maxAdaptiveWaitMultiplier, $currentAdaptiveMultiplier * 1.5)
            
            # Only log if verbose logging is enabled
            if ($verboseLogging) {
                LogMessage -level "DEBUG" -message "High system load detected. Increasing wait multiplier to $currentAdaptiveMultiplier" -details $systemMetrics
            }
        }
        elseif ($cpuUsage -lt 40 -and $availableMemory -gt $lowMemoryThreshold * 3) {
            # Low system load - decrease wait times
            $script:currentAdaptiveMultiplier = [Math]::Max(1.0, $currentAdaptiveMultiplier * 0.8)
            
            # Only log if verbose logging is enabled
            if ($verboseLogging) {
                LogMessage -level "DEBUG" -message "Low system load detected. Decreasing wait multiplier to $currentAdaptiveMultiplier" -details $systemMetrics
            }
        }
        else {
            # Moderate system load - gradually normalize wait times
            $script:currentAdaptiveMultiplier = [Math]::Max(1.0, $currentAdaptiveMultiplier * 0.95)
        }
    }
}

# Function to estimate test duration with historical data
function EstimateTestDuration {
    param (
        [string]$currency,
        [string]$timeframe,
        [string]$eaName
    )
    
    # Check if we have historical data for this combination
    $historyKey = "${eaName}_${currency}_${timeframe}"
    
    if ($performanceHistory.ContainsKey($historyKey)) {
        # Use historical data with some adjustment for current system conditions
        $historicalDuration = $performanceHistory[$historyKey]
        $estimatedDuration = $historicalDuration * $currentAdaptiveMultiplier
        
        LogMessage -level "DEBUG" -message "Using historical duration data for estimation" -details @{
            "historicalDuration" = $historicalDuration
            "estimatedDuration" = $estimatedDuration
            "multiplier" = $currentAdaptiveMultiplier
        }
    }
    else {
        # Base estimates on timeframe (in seconds)
        switch ($timeframe) {
            "M1" { $baseDuration = 300 }  # 5 minutes
            "M5" { $baseDuration = 240 }  # 4 minutes
            "M15" { $baseDuration = 180 } # 3 minutes
            "M30" { $baseDuration = 150 } # 2.5 minutes
            "H1" { $baseDuration = 120 }  # 2 minutes
            "H4" { $baseDuration = 90 }   # 1.5 minutes
            "D1" { $baseDuration = 60 }   # 1 minute
            default { $baseDuration = 180 } # 3 minutes default
        }
        
        # Adjust for currency pair complexity (some pairs take longer)
        if ($currency -eq "EURUSD" -or $currency -eq "GBPUSD" -or $currency -eq "USDJPY") {
            # Major pairs typically have more data and take longer
            $currencyMultiplier = 1.2
        }
        elseif ($currency -match "JPY" -or $currency -match "CHF") {
            # Cross pairs with JPY or CHF often take longer
            $currencyMultiplier = 1.1
        }
        else {
            $currencyMultiplier = 1.0
        }
        
        # Adjust for date range
        # Calculate approximate number of years in test
        $startYear = [int]($startDate -replace "^(\d{4}).*", '$1')
        $endYear = [int]($endDate -replace "^(\d{4}).*", '$1')
        $yearDifference = $endYear - $startYear + 1
        $dateRangeMultiplier = [Math]::Max(1.0, $yearDifference / 3)  # Normalize to 3 years as baseline
        
        # Calculate final estimate
        $estimatedDuration = $baseDuration * $currencyMultiplier * $dateRangeMultiplier
        
        # Apply system load factor
        $estimatedDuration = $estimatedDuration * $currentAdaptiveMultiplier
        
        # Log estimation factors
        LogMessage -level "DEBUG" -message "Estimated test duration based on parameters" -details @{
            "baseDuration" = $baseDuration
            "currencyMultiplier" = $currencyMultiplier
            "dateRangeMultiplier" = $dateRangeMultiplier
            "systemLoadMultiplier" = $currentAdaptiveMultiplier
            "estimatedDuration" = $estimatedDuration
        }
    }
    
    # Round to nearest 10 seconds
    $estimatedDuration = [Math]::Round($estimatedDuration / 10) * 10
    
    return $estimatedDuration
}

# Function to update performance history
function UpdatePerformanceHistory {
    param (
        [string]$currency,
        [string]$timeframe,
        [string]$eaName,
        [double]$actualDuration
    )
    
    # Create key for this combination
    $historyKey = "${eaName}_${currency}_${timeframe}"
    
    # Update or add the entry
    if ($performanceHistory.ContainsKey($historyKey)) {
        # Calculate weighted average (70% history, 30% new data)
        $historicalDuration = $performanceHistory[$historyKey]
        $newDuration = ($historicalDuration * 0.7) + ($actualDuration * 0.3)
    }
    else {
        # First entry for this combination
        $newDuration = $actualDuration
    }
    
    # Update the dictionary
    $performanceHistory[$historyKey] = $newDuration
    
    # Save to file
    try {
        $performanceHistory | ConvertTo-Json | Set-Content -Path $performanceHistoryFile
        LogMessage -level "DEBUG" -message "Updated performance history" -details @{
            "combination" = $historyKey
            "duration" = $newDuration
        }
    }
    catch {
        LogMessage -level "ERROR" -message "Failed to save performance history: $($_.Exception.Message)"
    }
}

# UI Automation Helper Functions
function FindWindow {
    param (
        [string]$windowTitle
    )
    
        $automation = [System.Windows.Automation.AutomationElement]::RootElement
    $condition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, $windowTitle)
    return $automation.FindFirst([System.Windows.Automation.TreeScope]::Children, $condition)
}

function SendKeysToWindow {
    param (
        $window,
        [string]$keys
    )
    
    if ($window -eq $null) {
        throw "Window not found"
    }
    
    # Get the window pattern
    $windowPattern = $window.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern)
    
    # Ensure window is active
    if ($windowPattern -ne $null) {
        $windowPattern.SetWindowVisualState([System.Windows.Automation.WindowVisualState]::Normal)
    }
    
    # Set focus to the window
    $window.SetFocus()
    
    # Send keys
    [System.Windows.Forms.SendKeys]::SendWait($keys)
    Start-Sleep -Milliseconds 500
}

function ClickElement {
    param (
        $element
    )
    
    if ($element -eq $null) {
        throw "Element not found"
    }
    
    $invokePattern = $element.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
    if ($invokePattern -ne $null) {
        $invokePattern.Invoke()
    }
    else {
        # Try to click using mouse if invoke pattern is not available
        $rect = $element.Current.BoundingRectangle
        $x = $rect.Left + $rect.Width / 2
        $y = $rect.Top + $rect.Height / 2
        
        [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x, $y)
        Start-Sleep -Milliseconds 100
        
        # Simulate mouse click
        $mouseDown = New-Object -TypeName System.Windows.Forms.MouseEventArgs([System.Windows.Forms.MouseButtons]::Left, 1, $x, $y, 0)
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 100
        
        $mouseUp = New-Object -TypeName System.Windows.Forms.MouseEventArgs([System.Windows.Forms.MouseButtons]::Left, 1, $x, $y, 0)
        [System.Windows.Forms.Application]::DoEvents()
    }
}

function FindElementByAutomationId {
    param (
        $parent,
        [string]$automationId
    )
    
    $condition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::AutomationIdProperty, $automationId)
    return $parent.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $condition)
}

function FindElementByName {
    param (
        $parent,
        [string]$name
    )
    
    $condition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, $name)
    return $parent.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $condition)
}

function NavigateToField {
    param (
        [int]$fieldPosition,
        [string]$windowName
    )
    
    try {
        $window = FindWindow -windowTitle $windowName
        if ($window -eq $null) {
            throw "Window '$windowName' not found"
        }
        
        # Start from home position
        SendKeysToWindow -window $window -keys "{HOME}"
        LegacyAdaptiveWait -waitTime 0.5
        
        # Tab to the desired field
        for ($i = 1; $i -le $fieldPosition; $i++) {
            SendKeysToWindow -window $window -keys "{TAB}"
            LegacyAdaptiveWait -waitTime 0.2
        }
        
        LogMessage -level "DEBUG" -message "Navigated to field position $fieldPosition in $windowName"
    }
    catch {
        LogMessage -level "ERROR" -message "Failed to navigate to field position $fieldPosition in $windowName: $($_.Exception.Message)"
        throw "Navigation failed"
    }
}

function EnterValue {
    param (
        [string]$value,
        [string]$windowName
    )
    
    try {
        $window = FindWindow -windowTitle $windowName
        if ($window -eq $null) {
            throw "Window '$windowName' not found"
        }
        
        # Select all existing text
        SendKeysToWindow -window $window -keys "^a"
        LegacyAdaptiveWait -waitTime 0.5
        
        # Enter the new value
        SendKeysToWindow -window $window -keys $value
        LegacyAdaptiveWait -waitTime 0.5
        
        LogMessage -level "DEBUG" -message "Entered value '$value' in $windowName"
    }
    catch {
        LogMessage -level "ERROR" -message "Failed to enter value '$value' in $windowName: $($_.Exception.Message)"
        throw "Value entry failed"
    }
}

function SelectDropdownItem {
    param (
        [string]$item,
        [string]$windowName
    )
    
    try {
        $window = FindWindow -windowTitle $windowName
        if ($window -eq $null) {
            throw "Window '$windowName' not found"
        }
        
        # Open dropdown
        SendKeysToWindow -window $window -keys "%{DOWN}"
        LegacyAdaptiveWait -waitTime 0.5
        
        # Type to find the item
        SendKeysToWindow -window $window -keys $item
        LegacyAdaptiveWait -waitTime 0.5
        
        # Select the item
        SendKeysToWindow -window $window -keys "{ENTER}"
        LegacyAdaptiveWait -waitTime 0.5
        
        LogMessage -level "DEBUG" -message "Selected item '$item' from dropdown in $windowName"
    }
    catch {
        LogMessage -level "ERROR" -message "Failed to select item '$item' from dropdown in $windowName: $($_.Exception.Message)"
        throw "Dropdown selection failed"
    }
}

function ConfigureField {
    param (
        [int]$fieldPosition,
        [string]$value,
        [string]$windowName
    )
    
    $retryCount = 0
    while ($retryCount -lt $maxRetries) {
        try {
            NavigateToField -fieldPosition $fieldPosition -windowName $windowName
            EnterValue -value $value -windowName $windowName
            break
        }
        catch {
            $retryCount++
            
            if ($retryCount -ge $maxRetries) {
                LogMessage -level "ERROR" -message "Failed to configure field at position $fieldPosition after $maxRetries attempts"
                throw "Field configuration failed"
            }
            else {
                LogMessage -level "WARN" -message "Retrying field configuration (attempt $retryCount)"
                AdaptiveWait -waitTime 2 -isRetry $true -retryCount $retryCount
            }
        }
    }
}

function ConfigureDropdownField {
    param (
        [int]$fieldPosition,
        [string]$item,
        [string]$windowName
    )
    
    $retryCount = 0
    while ($retryCount -lt $maxRetries) {
        try {
            NavigateToField -fieldPosition $fieldPosition -windowName $windowName
            SelectDropdownItem -item $item -windowName $windowName
            break
        }
        catch {
            $retryCount++
            
            if ($retryCount -ge $maxRetries) {
                LogMessage -level "ERROR" -message "Failed to configure dropdown field at position $fieldPosition after $maxRetries attempts"
                throw "Dropdown field configuration failed"
            }
            else {
                LogMessage -level "WARN" -message "Retrying dropdown field configuration (attempt $retryCount)"
                AdaptiveWait -waitTime 2 -isRetry $true -retryCount $retryCount
            }
        }
    }
}

# Define timeframes - updated to include all 21 MT5 timeframes
$timeframes = @(
    # Minute timeframes
    "M1", "M2", "M3", "M4", "M5", "M6", "M10", "M12", "M15", "M20", "M30",
    # Hour timeframes
    "H1", "H2", "H3", "H4", "H6", "H8", "H12",
    # Day and above timeframes
    "D1", "W1", "MN1"
)

# Define currency pairs - full set for comprehensive testing
$currencies = @(
    # Majors
    "EURUSD", "GBPUSD", "USDJPY", "USDCHF", "AUDUSD", "USDCAD", "NZDUSD",
    # Cross pairs
    "EURGBP", "EURJPY", "GBPJPY", "AUDJPY", "CADJPY", "CHFJPY",
    # Exotics
    "EURTRY", "USDZAR", "USDMXN", "USDSEK", "USDNOK"
)

# Get list of all EAs in the folder
$eaList = @()
try {
    # Get all EA files in the directory
    $eaList = Get-ChildItem -Path $eaPath -Filter "*.ex5" -File | ForEach-Object { $_.Name -replace "\.ex5$", "" }
    
    # If no EAs found, add a default one to prevent errors
    if ($eaList.Count -eq 0) {
        $eaList = @("Moving Average")
        LogMessage -level "WARN" -message "No EA files found in $eaPath, using default 'Moving Average'"
    }
    else {
        LogMessage -level "INFO" -message "Found $($eaList.Count) EA files in $eaPath"
        
        # Log all found EAs
        foreach ($eaName in $eaList) {
            LogMessage -level "DEBUG" -message "Found EA: $eaName"
        }
    }
}
catch {
    LogMessage -level "ERROR" -message "Error getting EA list: $($_.Exception.Message). Using default EA."
    # Add default EA to prevent errors
    $eaList = @("Moving Average")
}

# Load configuration if available
LoadConfiguration

# Check for checkpoint file to resume from previous run
if (Test-Path $checkpointFile) {
    try {
        $checkpointData = Get-Content -Path $checkpointFile -Raw
        LogMessage -level "INFO" -message "Found checkpoint file. Attempting to resume from last position..."
        
        # Try to parse as JSON
        try {
            $checkpoint = $checkpointData | ConvertFrom-Json
            $script:eaIndex = $checkpoint.eaIndex
            $script:currencyIndex = $checkpoint.currencyIndex
            $script:timeframeIndex = $checkpoint.timeframeIndex
            $script:eaName = $checkpoint.eaName
            $script:currency = $checkpoint.currency
            $script:timeframe = $checkpoint.timeframe
            $script:reportCounter = $checkpoint.reportCounter
        }
        catch {
            # Fallback to regex extraction if JSON parsing fails
            $script:eaIndex = [int]([regex]::Match($checkpointData, '"eaIndex":\s*(\d+)').Groups[1].Value)
            $script:currencyIndex = [int]([regex]::Match($checkpointData, '"currencyIndex":\s*(\d+)').Groups[1].Value)
            $script:timeframeIndex = [int]([regex]::Match($checkpointData, '"timeframeIndex":\s*(\d+)').Groups[1].Value)
            $script:eaName = [regex]::Match($checkpointData, '"eaName":\s*"([^"]+)"').Groups[1].Value
            $script:currency = [regex]::Match($checkpointData, '"currency":\s*"([^"]+)"').Groups[1].Value
            $script:timeframe = [regex]::Match($checkpointData, '"timeframe":\s*"([^"]+)"').Groups[1].Value
            $script:reportCounter = [int]([regex]::Match($checkpointData, '"reportCounter":\s*(\d+)').Groups[1].Value)
        }
        
        $script:resumeFromCheckpoint = $true
        
        $checkpointDetails = @{
            "eaName" = $eaName
            "eaIndex" = $eaIndex
            "currency" = $currency
            "currencyIndex" = $currencyIndex
            "timeframe" = $timeframe
            "timeframeIndex" = $timeframeIndex
        }
        
        LogMessage -level "INFO" -message "Resuming from checkpoint" -details $checkpointDetails
    }
    catch {
        LogMessage -level "ERROR" -message "Failed to parse checkpoint file: $($_.Exception.Message). Starting from beginning."
        $script:resumeFromCheckpoint = $false
    }
}

# Initial system resource check
DetailedSystemCheck

# Function to optimize MT5 settings for backtesting
function OptimizeMT5ForBacktesting {
    try {
        LogMessage -level "INFO" -message "Optimizing MT5 settings for backtesting"
        
        # Find MT5 window
        $mt5Window = FindWindow -windowTitle "MetaTrader 5"
        if ($mt5Window -eq $null) {
            throw "MetaTrader 5 window not found"
        }
        
        # Open settings dialog
        SendKeysToWindow -window $mt5Window -keys "%o"
        LegacyAdaptiveWait -waitTime 2
        
        # Find settings window
        $settingsWindow = FindWindow -windowTitle "Settings"
        if ($settingsWindow -eq $null) {
            throw "Settings window not found"
        }
        
        # Navigate to Strategy Tester tab
        NavigateToField -fieldPosition 4 -windowName "Settings"
        SendKeysToWindow -window $settingsWindow -keys "{RIGHT}{RIGHT}{RIGHT}"
        LegacyAdaptiveWait -waitTime 1
        
        # Disable visual mode by default
        NavigateToField -fieldPosition 4 -windowName "Settings"
        SendKeysToWindow -window $settingsWindow -keys " "
        LegacyAdaptiveWait -waitTime 1
        
        # Increase max bars in history for more accurate testing
        NavigateToField -fieldPosition 6 -windowName "Settings"
        EnterValue -value "0" -windowName "Settings"  # 0 means unlimited
        
        # Optimize memory usage settings
        NavigateToField -fieldPosition 9 -windowName "Settings"
        EnterValue -value "8192" -windowName "Settings"  # Increase memory buffer
        
        # Save settings
        SendKeysToWindow -window $settingsWindow -keys "%o"
        LegacyAdaptiveWait -waitTime 2
        
        LogMessage -level "INFO" -message "MT5 settings optimized for backtesting"
    }
    catch {
                LogMessage -level "WARN" -message "Failed to optimize MT5 settings: $($_.Exception.Message)"
        # Continue despite failure - not critical
    }
}

# Function to save current Strategy Tester settings
function SaveCurrentSettings {
    try {
        # Create a unique template name based on EA and currency
        $templateName = "${eaName}_${currency}_settings"
        
        # Find Strategy Tester window
        $testerWindow = FindWindow -windowTitle "Strategy Tester"
        if ($testerWindow -eq $null) {
            throw "Strategy Tester window not found"
        }
        
        # Save settings as template
        SendKeysToWindow -window $testerWindow -keys "%t"  # Alt+T for Template menu
        LegacyAdaptiveWait -waitTime 1
        SendKeysToWindow -window $testerWindow -keys "s"  # S for Save As option
        LegacyAdaptiveWait -waitTime 2
        
        # Enter template name in Save As dialog
        $saveAsWindow = FindWindow -windowTitle "Save As"
        if ($saveAsWindow -eq $null) {
            throw "Save As dialog not found"
        }
        
        SendKeysToWindow -window $saveAsWindow -keys "$reportPath\templates\$templateName"
        LegacyAdaptiveWait -waitTime 1
        SendKeysToWindow -window $saveAsWindow -keys "{ENTER}"
        LegacyAdaptiveWait -waitTime 2
        
        # Handle potential overwrite confirmation
        $confirmWindow = FindWindow -windowTitle "Confirm"
        if ($confirmWindow -ne $null) {
            SendKeysToWindow -window $confirmWindow -keys "y"
            LegacyAdaptiveWait -waitTime 2
        }
        
        LogMessage -level "INFO" -message "Settings saved as template: $templateName"
    }
    catch {
        LogMessage -level "ERROR" -message "Failed to save settings template: $($_.Exception.Message)"
    }
}

# Function to load saved Strategy Tester settings
function LoadSettings {
    try {
        # Use the same template name format as in SaveCurrentSettings
        $templateName = "${eaName}_${currency}_settings"
        
        # Find Strategy Tester window
        $testerWindow = FindWindow -windowTitle "Strategy Tester"
        if ($testerWindow -eq $null) {
            throw "Strategy Tester window not found"
        }
        
        # Load settings from template
        SendKeysToWindow -window $testerWindow -keys "%t"  # Alt+T for Template menu
        LegacyAdaptiveWait -waitTime 1
        SendKeysToWindow -window $testerWindow -keys "l"  # L for Load option
        LegacyAdaptiveWait -waitTime 2
        
        # Navigate to and select the template in Open dialog
        $openWindow = FindWindow -windowTitle "Open"
        if ($openWindow -eq $null) {
            throw "Open dialog not found"
        }
        
        SendKeysToWindow -window $openWindow -keys $templateName
        LegacyAdaptiveWait -waitTime 1
        SendKeysToWindow -window $openWindow -keys "{ENTER}"
        LegacyAdaptiveWait -waitTime 2
        
        LogMessage -level "INFO" -message "Settings loaded from template: $templateName"
    }
    catch {
        LogMessage -level "WARN" -message "Failed to load settings template: $($_.Exception.Message). Using default settings."
    }
}

# Launch MT5 with error handling and retry mechanism
$retryCount = 0
while ($retryCount -lt $maxRetries) {
    try {
        LogMessage -level "INFO" -message "Launching MetaTrader 5 (attempt $($retryCount + 1))"
        Start-Process -FilePath $mt5Path
        
        # Wait for MT5 to fully load with verification
        LegacyAdaptiveWait -waitTime $initialLoadTime
        
        # Check if MT5 is running
        $mt5Window = FindWindow -windowTitle "MetaTrader 5"
        if ($mt5Window -eq $null) {
            throw "MetaTrader 5 did not start properly"
        }
        
        LogMessage -level "INFO" -message "MetaTrader 5 launched successfully"
        
        # After MT5 is launched successfully, optimize it for backtesting
        OptimizeMT5ForBacktesting
        
        # Configure MT5 to use unlimited bars in chart
        try {
            LogMessage -level "INFO" -message "Configuring MT5 to use unlimited bars in chart"
            
            # Find MT5 window
            $mt5Window = FindWindow -windowTitle "MetaTrader 5"
            if ($mt5Window -eq $null) {
                throw "MetaTrader 5 window not found"
            }
            
            # Open MT5 settings dialog
            SendKeysToWindow -window $mt5Window -keys "%o"
            LegacyAdaptiveWait -waitTime 2
            
            # Find settings window
            $settingsWindow = FindWindow -windowTitle "Settings"
            if ($settingsWindow -eq $null) {
                throw "Settings window not found"
            }
            
            # Navigate to Charts tab (typically the 2nd tab)
            NavigateToField -fieldPosition 1 -windowName "Settings"
            SendKeysToWindow -window $settingsWindow -keys "{RIGHT}"
            LegacyAdaptiveWait -waitTime 1
            
            # Navigate to "Max bars in chart" field
            NavigateToField -fieldPosition 8 -windowName "Settings"
            
            # Enter unlimited value (0)
            EnterValue -value "0" -windowName "Settings"
            
            # Navigate to OK button and click it
            NavigateToField -fieldPosition 16 -windowName "Settings"
            SendKeysToWindow -window $settingsWindow -keys "{ENTER}"
            LegacyAdaptiveWait -waitTime 2
            
            LogMessage -level "INFO" -message "Successfully set maximum bars in chart to unlimited"
        }
        catch {
            LogMessage -level "WARN" -message "Could not set maximum bars in chart to unlimited: $($_.Exception.Message)"
            
            # Try alternative method if first method fails
            try {
                # Find MT5 window
                $mt5Window = FindWindow -windowTitle "MetaTrader 5"
                if ($mt5Window -eq $null) {
                    throw "MetaTrader 5 window not found"
                }
                
                # Alternative method using more direct keyboard navigation
                SendKeysToWindow -window $mt5Window -keys "%"
                LegacyAdaptiveWait -waitTime 1
                SendKeysToWindow -window $mt5Window -keys "t"  # Tools menu
                LegacyAdaptiveWait -waitTime 1
                SendKeysToWindow -window $mt5Window -keys "o"  # Options
                LegacyAdaptiveWait -waitTime 2
                
                # Find settings window
                $settingsWindow = FindWindow -windowTitle "Settings"
                if ($settingsWindow -eq $null) {
                    throw "Settings window not found"
                }
                
                # Navigate to Charts tab
                SendKeysToWindow -window $settingsWindow -keys "{RIGHT}"
                LegacyAdaptiveWait -waitTime 1
                
                # Try to find the Max bars field by typing its name
                SendKeysToWindow -window $settingsWindow -keys "Max bars"
                LegacyAdaptiveWait -waitTime 1
                SendKeysToWindow -window $settingsWindow -keys "{TAB}"
                LegacyAdaptiveWait -waitTime 1
                
                # Enter 0 for unlimited
                SendKeysToWindow -window $settingsWindow -keys "0"
                LegacyAdaptiveWait -waitTime 1
                
                # Press OK
                SendKeysToWindow -window $settingsWindow -keys "%o"
                LegacyAdaptiveWait -waitTime 2
                
                LogMessage -level "INFO" -message "Successfully set maximum bars in chart to unlimited using alternative method"
            }
            catch {
                LogMessage -level "WARN" -message "Failed to set maximum bars in chart to unlimited. Backtests may have limited historical data."
                # Continue with the script despite this error
            }
        }
        
        break
    }
    catch {
        $retryCount++
        LogMessage -level "ERROR" -message "Error launching MetaTrader 5 (attempt $retryCount): $($_.Exception.Message)"
        
        # Capture error state
        CaptureErrorState -errorContext "MT5Launch"
        
        if ($retryCount -ge $maxRetries) {
            if ($skipOnError) {
                LogMessage -level "WARN" -message "Maximum retries reached. Continuing with script..."
                break
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("Failed to launch MetaTrader 5 after $maxRetries attempts", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                exit
            }
        }
        else {
            # Wait before retrying with exponential backoff
            AdaptiveWait -waitTime 10 -isRetry $true -retryCount $retryCount
        }
    }
}

# Create templates directory if it doesn't exist
if (-not (Test-Path "$reportPath\templates")) {
    New-Item -Path "$reportPath\templates" -ItemType Directory -Force
}

# Main loop for EA testing
for ($eaIdx = 0; $eaIdx -lt $eaList.Count; $eaIdx++) {
    # Skip to checkpoint position if resuming
    if ($resumeFromCheckpoint -and $eaIdx -lt $eaIndex) {
        continue
    }
    
    $eaName = $eaList[$eaIdx]
    
    LogMessage -level "INFO" -message "Starting tests for EA: $eaName"
    
    # Loop through each currency pair
    for ($currIdx = 0; $currIdx -lt $currencies.Count; $currIdx++) {
        # Skip to checkpoint position if resuming
        if ($resumeFromCheckpoint -and $eaIdx -eq $eaIndex -and $currIdx -lt $currencyIndex) {
            continue
        }
        
        $currency = $currencies[$currIdx]
        
        try {
            LogMessage -level "INFO" -message "Starting tests for EA: $eaName on $currency"
            
            # Check system resources before starting a new test
            DetailedSystemCheck
            
            # Open Strategy Tester using hotkey with retry
            $retryCount = 0
            while ($retryCount -lt $maxRetries) {
                try {
                    # Find MT5 window
                    $mt5Window = FindWindow -windowTitle "MetaTrader 5"
                    if ($mt5Window -eq $null) {
                        throw "MetaTrader 5 window not found"
                    }
                    
                    # Primary method: Alt+T
                    SendKeysToWindow -window $mt5Window -keys "%t"
                    LegacyAdaptiveWait -waitTime 3
                    
                    # Check if Strategy Tester opened
                    $testerWindow = FindWindow -windowTitle "Strategy Tester"
                    $altTesterWindow = FindWindow -windowTitle "Tester"
                    
                    if ($testerWindow -eq $null -and $altTesterWindow -eq $null) {
                        # Secondary method: Ctrl+R
                        SendKeysToWindow -window $mt5Window -keys "^r"
                        LegacyAdaptiveWait -waitTime 3
                        
                        # Check again
                        $testerWindow = FindWindow -windowTitle "Strategy Tester"
                        $altTesterWindow = FindWindow -windowTitle "Tester"
                        
                        if ($testerWindow -eq $null -and $altTesterWindow -eq $null) {
                            # Tertiary method: Use View menu
                            SendKeysToWindow -window $mt5Window -keys "%"
                            LegacyAdaptiveWait -waitTime 1
                            SendKeysToWindow -window $mt5Window -keys "v"
                            LegacyAdaptiveWait -waitTime 1
                            SendKeysToWindow -window $mt5Window -keys "t"
                            LegacyAdaptiveWait -waitTime 3
                            
                            # Check again
                            $testerWindow = FindWindow -windowTitle "Strategy Tester"
                            $altTesterWindow = FindWindow -windowTitle "Tester"
                            
                            if ($testerWindow -eq $null -and $altTesterWindow -eq $null) {
                                throw "Strategy Tester window did not open"
                            }
                        }
                    }
                    
                    break
                }
                catch {
                    $retryCount++
                    LogMessage -level "ERROR" -message "Failed to open Strategy Tester (attempt $retryCount): $($_.Exception.Message)"
                    
                    # Capture error state
                    CaptureErrorState -errorContext "OpenTester"
                    
                    if ($retryCount -ge $maxRetries) {
                        throw "Failed to open Strategy Tester after $maxRetries attempts"
                    }
                    else {
                        # Use exponential backoff for retries
                        AdaptiveWait -waitTime 5 -isRetry $true -retryCount $retryCount
                    }
                }
            }
            
            # Determine which window title to use
            $testerWindowTitle = if (FindWindow -windowTitle "Strategy Tester" -ne $null) { "Strategy Tester" } else { "Tester" }
            
            # Configure Strategy Tester using modular UI interaction functions
            # First, ensure focus is in the Strategy Tester window
            $testerWindow = FindWindow -windowTitle $testerWindowTitle
            SendKeysToWindow -window $testerWindow -keys "%{TAB}"
            LegacyAdaptiveWait -waitTime 1
            
            # Select EA using the ConfigureDropdownField function
            try {
                ConfigureDropdownField -fieldPosition 1 -item $eaName -windowName $testerWindowTitle
            }
            catch {
                LogMessage -level "ERROR" -message "Error selecting EA $eaName: $($_.Exception.Message). Skipping to next EA."
                CaptureErrorState -errorContext "SelectEA"
                continue
            }
            
            # Check if symbol exists and select Symbol
            try {
                ConfigureDropdownField -fieldPosition 2 -item $currency -windowName $testerWindowTitle
            }
            catch {
                LogMessage -level "WARN" -message "Symbol $currency not available with broker for EA $eaName, skipping..."
                # Close Strategy Tester window to prepare for next currency
                $testerWindow = FindWindow -windowTitle $testerWindowTitle
                SendKeysToWindow -window $testerWindow -keys "%{F4}"
                LegacyAdaptiveWait -waitTime 2
                
                # Reset consecutive failures counter since we're skipping by design
                $consecutiveFailures = 0
                
                continue
            }
            
            # Select Model - only once per currency
            try {
                ConfigureDropdownField -fieldPosition 3 -item "Every tick" -windowName $testerWindowTitle
            }
            catch {
                LogMessage -level "WARN" -message "Error selecting model for EA $eaName: $($_.Exception.Message). Using default model."
                                # Continue anyway with default model
            }
            
            # Set Date Range - only once per currency
            try {
                # Navigate to "Use date" checkbox
                NavigateToField -fieldPosition 5 -windowName $testerWindowTitle
                # Check the box if not already checked
                SendKeysToWindow -window (FindWindow -windowTitle $testerWindowTitle) -keys " "
                LegacyAdaptiveWait -waitTime 1
                
                # Configure From date field
                ConfigureField -fieldPosition 6 -value $startDate -windowName $testerWindowTitle
                
                # Configure To date field
                ConfigureField -fieldPosition 7 -value $endDate -windowName $testerWindowTitle
            }
            catch {
                LogMessage -level "WARN" -message "Error setting date range for EA $eaName: $($_.Exception.Message). Using default date range."
                # Continue anyway with default date range
            }
            
            # Save these base settings
            SaveCurrentSettings
            
            # Loop through timeframes
            for ($tfIdx = 0; $tfIdx -lt $timeframes.Count; $tfIdx++) {
                # Skip to checkpoint position if resuming
                if ($resumeFromCheckpoint -and $eaIdx -eq $eaIndex -and $currIdx -eq $currencyIndex -and $tfIdx -lt $timeframeIndex) {
                    continue
                }
                elseif ($resumeFromCheckpoint -and $eaIdx -eq $eaIndex -and $currIdx -eq $currencyIndex -and $tfIdx -eq $timeframeIndex) {
                    # We've reached the exact checkpoint position, disable resuming for subsequent iterations
                    $resumeFromCheckpoint = $false
                }
                
                $timeframe = $timeframes[$tfIdx]
                
                try {
                    LogMessage -level "INFO" -message "Setting timeframe to $timeframe for EA $eaName on $currency"
                    
                    # Check system resources before starting a new timeframe test
                    DetailedSystemCheck
                    
                    # Load the saved settings
                    LoadSettings
                    
                    # Only modify the timeframe using the ConfigureDropdownField function
                    ConfigureDropdownField -fieldPosition 4 -item $timeframe -windowName $testerWindowTitle
                    
                    # Save the modified settings (optional - only if you want to preserve the last state)
                    SaveCurrentSettings
                    
                    # Estimate test duration for progress reporting
                    $estimatedDuration = EstimateTestDuration -currency $currency -timeframe $timeframe -eaName $eaName
                    
                    # Save checkpoint before starting test
                    $script:eaIndex = $eaIdx
                    $script:currencyIndex = $currIdx
                    $script:timeframeIndex = $tfIdx
                    SaveCheckpoint
                    
                    # Start Test with keyboard shortcut
                    $testerWindow = FindWindow -windowTitle $testerWindowTitle
                    SendKeysToWindow -window $testerWindow -keys "{F9}"
                    LegacyAdaptiveWait -waitTime 5
                    
                    $testStartDetails = @{
                        "ea" = $eaName
                        "currency" = $currency
                        "timeframe" = $timeframe
                        "estimatedDuration" = $estimatedDuration
                    }
                    
                    LogMessage -level "INFO" -message "Test started" -details $testStartDetails
                    
                    # Wait for test to complete with optimized progress monitoring
                    $testWaitTime = 0
                    $testCompleted = $false
                    $statusCheckInterval = 10  # Check every 10 seconds
                    $lastProgressValue = "0"
                    $noProgressCounter = 0
                    $maxNoProgressIntervals = 30  # Allow 5 minutes (30 × 10s) without progress before considering frozen
                    $mtFrozenCounter = 0
                    $previousLoggedProgress = "0"
                    $lastProgressLogTime = 0
                    $testStartTime = [int](Get-Date -UFormat %s)
                    
                    LogMessage -level "INFO" -message "Monitoring backtest progress..."
                    
                    while (-not $testCompleted) {
                        # Method 1: Check if Start button is enabled again (test completed)
                        try {
                            $startButton = FindElementByName -parent (FindWindow -windowTitle $testerWindowTitle) -name "Start"
                            if ($startButton -ne $null) {
                                $isEnabled = $startButton.Current.IsEnabled
                                if ($isEnabled) {
                                    LogMessage -level "INFO" -message "Test completion detected: Start button is enabled again"
                                    $testCompleted = $true
                                    break
                                }
                            }
                        }
                        catch {
                            # Continue to other detection methods
                        }
                        
                        # Method 2: Check for report tab appearance
                        try {
                            $reportTab = FindElementByName -parent (FindWindow -windowTitle $testerWindowTitle) -name "Report"
                            if ($reportTab -ne $null) {
                                LogMessage -level "INFO" -message "Test completion detected: Report tab appeared"
                                $testCompleted = $true
                                break
                            }
                        }
                        catch {
                            # Continue to other detection methods
                        }
                        
                        # Method 3: Check status bar text for completion indicators
                        try {
                            $statusBar = FindElementByAutomationId -parent (FindWindow -windowTitle $testerWindowTitle) -automationId "StatusBar"
                            if ($statusBar -ne $null) {
                                $statusText = $statusBar.Current.Name
                                if ($statusText -match "complete|100%|finished") {
                                    LogMessage -level "INFO" -message "Test completion detected: Status bar indicates completion"
                                    $testCompleted = $true
                                    break
                                }
                                
                                # Check for progress changes to detect if test is still running
                                # Extract progress percentage from status text if available
                                if ($statusText -match "(\d+)%") {
                                    # Extract just the percentage value
                                    $currentProgress = $matches[1]
                                    
                                    if ($currentProgress -ne $lastProgressValue) {
                                        # Progress has changed, reset the no-progress counter
                                        $lastProgressValue = $currentProgress
                                        $noProgressCounter = 0
                                        
                                        # Calculate estimated remaining time
                                        if ([int]$lastProgressValue -gt 0) {
                                            $elapsedTime = [int](Get-Date -UFormat %s) - $testStartTime
                                            $progressFraction = [int]$lastProgressValue / 100
                                            $totalEstimatedTime = $elapsedTime / $progressFraction
                                            $remainingTime = $totalEstimatedTime - $elapsedTime
                                            
                                            # Format remaining time
                                            $remainingMinutes = [Math]::Floor($remainingTime / 60)
                                            $remainingSeconds = [Math]::Floor($remainingTime % 60)
                                            $remainingTimeFormatted = "${remainingMinutes}m ${remainingSeconds}s"
                                            
                                            # Log progress less frequently for long runs
                                            $currentTime = [int](Get-Date -UFormat %s)
                                            if (($currentTime - $lastProgressLogTime -gt 300) -or 
                                                ([int]$lastProgressValue % $logProgressInterval -eq 0 -and $lastProgressValue -ne $previousLoggedProgress)) {
                                                $progressDetails = @{
                                                    "progress" = [int]$lastProgressValue
                                                    "elapsedTime" = $elapsedTime
                                                    "estimatedRemaining" = $remainingTimeFormatted
                                                }
                                                
                                                LogMessage -level "INFO" -message "Backtest in progress: $lastProgressValue% complete" -details $progressDetails
                                                $previousLoggedProgress = $lastProgressValue
                                                $lastProgressLogTime = $currentTime
                                            }
                                        }
                                    }
                                    else {
                                        # No change in progress, increment counter
                                        $noProgressCounter++
                                    }
                                }
                            }
                        }
                        catch {
                            # Continue to other detection methods
                        }
                        
                        # Check system resources periodically during the test - less frequently for long runs
                        if ($testWaitTime % 120 -eq 0) {
                            DetailedSystemCheck
                        }
                        
                        # Check if MT5 is responsive - less frequently for long runs
                        if ($testWaitTime % 120 -eq 0 -and $testWaitTime -gt 0) {
                            try {
                                # Send a harmless key to check if window responds
                                $testerWindow = FindWindow -windowTitle $testerWindowTitle
                                if ($testerWindow -ne $null) {
                                    SendKeysToWindow -window $testerWindow -keys "{HOME}"
                                    LegacyAdaptiveWait -waitTime 1
                                    
                                    # Reset frozen counter if MT5 responds
                                    $mtFrozenCounter = 0
                                }
                                else {
                                    # Window not found, might be frozen or closed
                                    $mtFrozenCounter++
                                }
                            }
                            catch {
                                # MT5 didn't respond
                                $mtFrozenCounter++
                                LogMessage -level "WARN" -message "Warning: MT5 may be unresponsive (attempt $mtFrozenCounter)"
                                
                                # Check system resources before declaring frozen
                                DetailedSystemCheck
                                
                                # Only consider MT5 frozen after multiple failed response checks
                                # Be more tolerant if system is under load
                                if ($mtFrozenCounter -ge 3 -and $availableMemory -gt $lowMemoryThreshold) {
                                    LogMessage -level "ERROR" -message "MT5 appears to be frozen. Attempting recovery..."
                                    CaptureErrorState -errorContext "MT5Frozen"
                                    $testCompleted = $true
                                    $consecutiveFailures++
                                    break
                                }
                                elseif ($mtFrozenCounter -ge 5) {
                                    # Even with low memory, don't wait forever
                                    LogMessage -level "ERROR" -message "MT5 appears to be frozen despite system load. Attempting recovery..."
                                    CaptureErrorState -errorContext "MT5Frozen"
                                    $testCompleted = $true
                                    $consecutiveFailures++
                                    break
                                }
                            }
                        }
                        
                        # Periodic heartbeat log to show script is still running - reduced frequency
                        if ($testWaitTime % 1800 -eq 0 -and $testWaitTime -gt 0) {
                            LogMessage -level "INFO" -message "Backtest still running after $testWaitTime seconds. Continuing to wait for completion..."
                        }
                        
                        LegacyAdaptiveWait -waitTime $statusCheckInterval
                        $testWaitTime += $statusCheckInterval
                    }
                    
                    # Record actual test duration for future estimates
                    $actualTestDuration = [int](Get-Date -UFormat %s) - $testStartTime
                    UpdatePerformanceHistory -currency $currency -timeframe $timeframe -eaName $eaName -actualDuration $actualTestDuration
                    
                    # If we detected the test is frozen, try to recover
                    if ($testCompleted -and $consecutiveFailures -gt 0) {
                        # Try to cancel any hanging test
                        try {
                            $testerWindow = FindWindow -windowTitle $testerWindowTitle
                            if ($testerWindow -ne $null) {
                                SendKeysToWindow -window $testerWindow -keys "{ESC}"
                                LegacyAdaptiveWait -waitTime 2
                                SendKeysToWindow -window $testerWindow -keys "{ENTER}"  # In case a confirmation dialog appears
                                LegacyAdaptiveWait -waitTime 1
                            }
                        }
                        catch {
                            # Ignore errors when canceling
                        }
                        
                        # Skip to next timeframe if we had to force completion
                        if ($mtFrozenCounter -ge 3 -or $noProgressCounter -ge $maxNoProgressIntervals * 2) {
                            LogMessage -level "WARN" -message "Skipping to next timeframe due to detected freeze"
                            continue
                        }
                    }
                    
                    # Save Report as Excel format
                    try {
                        # Find the report tab
                        $testerWindow = FindWindow -windowTitle $testerWindowTitle
                        $reportTab = FindElementByName -parent $testerWindow -name "Report"
                        
                        if ($reportTab -eq $null) {
                            throw "Report tab not found"
                        }
                        
                        # Right-click on the report tab to open context menu
                        $rect = $reportTab.Current.BoundingRectangle
                        $x = $rect.Left + $rect.Width / 2
                        $y = $rect.Top + $rect.Height / 2
                        
                        [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x, $y)
                        Start-Sleep -Milliseconds 100
                        
                        # Simulate right-click
                        [System.Windows.Forms.MouseEventArgs]::new([System.Windows.Forms.MouseButtons]::Right, 1, $x, $y, 0)
                        [System.Windows.Forms.Application]::DoEvents()
                        Start-Sleep -Milliseconds 100
                        
                        # Select "Report" from the context menu
                        LegacyAdaptiveWait -waitTime 1
                        SendKeysToWindow -window $testerWindow -keys "r"
                        LegacyAdaptiveWait -waitTime 1
                        
                        # Select "Excel" from the submenu
                        SendKeysToWindow -window $testerWindow -keys "e"
                        LegacyAdaptiveWait -waitTime 2
                        
                        # Now we should have the Save As dialog with the Excel-compatible XML format
                        $saveAsWindow = FindWindow -windowTitle "Save As"
                        if ($saveAsWindow -eq $null) {
                            throw "Save As dialog not found"
                        }
                        
                        # Get the default filename that MT5 suggests
                        SendKeysToWindow -window $saveAsWindow -keys "^a"
                        LegacyAdaptiveWait -waitTime 1
                        SendKeysToWindow -window $saveAsWindow -keys "^c"
                        LegacyAdaptiveWait -waitTime 1
                        $defaultFileName = [System.Windows.Forms.Clipboard]::GetText()
                        
                        # Append our custom naming to MT5's default filename
                        $customSuffix = "_${eaName}_${currency}_${timeframe}_($reportCounter)"
                        
                        # Remove extension to add our suffix before the extension
                        if ($defaultFileName -match "(.*)\.[^.]*$") {
                            $baseFileName = $matches[1]
                        }
                        else {
                            $baseFileName = $defaultFileName
                        }
                        
                        # Get the extension
                        if ($defaultFileName -match ".*(\.[^.]*$)") {
                                                        $fileExtension = $matches[1]
                        }
                        else {
                            $fileExtension = ".xml"
                        }
                        
                        $reportFileName = "${baseFileName}${customSuffix}${fileExtension}"
                        
                        # Set the complete path and filename
                        SendKeysToWindow -window $saveAsWindow -keys "$reportPath\$reportFileName"
                        LegacyAdaptiveWait -waitTime 1
                        SendKeysToWindow -window $saveAsWindow -keys "{ENTER}"
                        LegacyAdaptiveWait -waitTime 3
                        
                        $reportDetails = @{
                            "filename" = $reportFileName
                            "ea" = $eaName
                            "currency" = $currency
                            "timeframe" = $timeframe
                            "duration" = $actualTestDuration
                        }
                        
                        LogMessage -level "INFO" -message "Excel report saved" -details $reportDetails
                        
                        # Increment counter
                        $reportCounter++
                        
                        # Reset consecutive failures counter on success
                        $consecutiveFailures = 0
                        
                        # Close report tab with keyboard shortcut
                        $testerWindow = FindWindow -windowTitle $testerWindowTitle
                        SendKeysToWindow -window $testerWindow -keys "^{F4}"
                        LegacyAdaptiveWait -waitTime 2
                    }
                    catch {
                        # Fallback method using alternative approach
                        try {
                            # Try alternative method to access Excel report
                            # First ensure focus is on Strategy Tester
                            $testerWindow = FindWindow -windowTitle $testerWindowTitle
                            SendKeysToWindow -window $testerWindow -keys "%{TAB}"
                            LegacyAdaptiveWait -waitTime 1
                            
                            # Use Alt key to access menu
                            SendKeysToWindow -window $testerWindow -keys "%"
                            LegacyAdaptiveWait -waitTime 1
                            SendKeysToWindow -window $testerWindow -keys "v"  # View menu
                            LegacyAdaptiveWait -waitTime 1
                            SendKeysToWindow -window $testerWindow -keys "r"  # Report submenu
                            LegacyAdaptiveWait -waitTime 1
                            SendKeysToWindow -window $testerWindow -keys "e"  # Excel option
                            LegacyAdaptiveWait -waitTime 2
                            
                            # Now we should have the Save As dialog with the Excel-compatible XML format
                            $saveAsWindow = FindWindow -windowTitle "Save As"
                            if ($saveAsWindow -eq $null) {
                                throw "Save As dialog not found"
                            }
                            
                            # Get the default filename that MT5 suggests
                            SendKeysToWindow -window $saveAsWindow -keys "^a"
                            LegacyAdaptiveWait -waitTime 1
                            SendKeysToWindow -window $saveAsWindow -keys "^c"
                            LegacyAdaptiveWait -waitTime 1
                            $defaultFileName = [System.Windows.Forms.Clipboard]::GetText()
                            
                            # Append our custom naming to MT5's default filename
                            $customSuffix = "_${eaName}_${currency}_${timeframe}_($reportCounter)"
                            
                            # Remove extension to add our suffix before the extension
                            if ($defaultFileName -match "(.*)\.[^.]*$") {
                                $baseFileName = $matches[1]
                            }
                            else {
                                $baseFileName = $defaultFileName
                            }
                            
                            # Get the extension
                            if ($defaultFileName -match ".*(\.[^.]*$)") {
                                $fileExtension = $matches[1]
                            }
                            else {
                                $fileExtension = ".xml"
                            }
                            
                            $reportFileName = "${baseFileName}${customSuffix}${fileExtension}"
                            
                            # Set the complete path and filename
                            SendKeysToWindow -window $saveAsWindow -keys "$reportPath\$reportFileName"
                            LegacyAdaptiveWait -waitTime 1
                            SendKeysToWindow -window $saveAsWindow -keys "{ENTER}"
                            LegacyAdaptiveWait -waitTime 3
                            
                            $reportDetails = @{
                                "filename" = $reportFileName
                                "ea" = $eaName
                                "currency" = $currency
                                "timeframe" = $timeframe
                                "duration" = $actualTestDuration
                                "method" = "alternative"
                            }
                            
                            LogMessage -level "INFO" -message "Excel report saved using alternative method" -details $reportDetails
                            
                            # Increment counter
                            $reportCounter++
                            
                            # Reset consecutive failures counter on success
                            $consecutiveFailures = 0
                            
                            # Close report tab with keyboard shortcut
                            $testerWindow = FindWindow -windowTitle $testerWindowTitle
                            SendKeysToWindow -window $testerWindow -keys "^{F4}"
                            LegacyAdaptiveWait -waitTime 2
                        }
                        catch {
                            LogMessage -level "ERROR" -message "Failed to save Excel report: $($_.Exception.Message)"
                            CaptureErrorState -errorContext "SaveReport"
                            
                            # Increment consecutive failures counter
                            $consecutiveFailures++
                            
                            # Try to close any open dialogs or tabs
                            try {
                                $saveAsWindow = FindWindow -windowTitle "Save As"
                                if ($saveAsWindow -ne $null) {
                                    SendKeysToWindow -window $saveAsWindow -keys "{ESC}"
                                    LegacyAdaptiveWait -waitTime 1
                                }
                                
                                $testerWindow = FindWindow -windowTitle $testerWindowTitle
                                if ($testerWindow -ne $null) {
                                    SendKeysToWindow -window $testerWindow -keys "^{F4}"
                                    LegacyAdaptiveWait -waitTime 2
                                }
                            }
                            catch {
                                # Ignore errors when closing
                            }
                        }
                    }
                    
                    # Check if we need to restart MT5 due to consecutive failures
                    if ($autoRestartOnFailure -and $consecutiveFailures -ge $maxConsecutiveFailures) {
                        LogMessage -level "WARN" -message "Detected $consecutiveFailures consecutive failures. Restarting MT5..."
                        
                        # Save checkpoint before restart
                        SaveCheckpoint
                        
                        # Close MT5
                        try {
                            $mt5Window = FindWindow -windowTitle "MetaTrader 5"
                            if ($mt5Window -ne $null) {
                                SendKeysToWindow -window $mt5Window -keys "%{F4}"
                                LegacyAdaptiveWait -waitTime 2
                            }
                            
                            # Handle potential "Save changes" dialog
                            $saveWindow = FindWindow -windowTitle "Save"
                            if ($saveWindow -ne $null) {
                                SendKeysToWindow -window $saveWindow -keys "n"  # Don't save changes
                                LegacyAdaptiveWait -waitTime 2
                            }
                            
                            # Make sure MT5 is closed
                            Stop-Process -Name "terminal64" -Force -ErrorAction SilentlyContinue
                            LegacyAdaptiveWait -waitTime 5
                            
                            # Restart MT5
                            Start-Process -FilePath $mt5Path
                            LegacyAdaptiveWait -waitTime $initialLoadTime
                            
                            # Reset consecutive failures counter
                            $consecutiveFailures = 0
                            
                            # Wait for MT5 to fully load
                            LegacyAdaptiveWait -waitTime 10
                            
                            # Re-open Strategy Tester
                            $mt5Window = FindWindow -windowTitle "MetaTrader 5"
                            if ($mt5Window -ne $null) {
                                SendKeysToWindow -window $mt5Window -keys "%t"
                                LegacyAdaptiveWait -waitTime 3
                            }
                            
                            LogMessage -level "INFO" -message "MT5 restarted successfully"
                        }
                        catch {
                            LogMessage -level "ERROR" -message "Failed to restart MT5: $($_.Exception.Message)"
                            CaptureErrorState -errorContext "RestartMT5"
                            
                            # Try to continue anyway
                            $consecutiveFailures = 0  # Reset to avoid infinite restart loop
                        }
                    }
                }
                catch {
                    LogMessage -level "ERROR" -message "Error during test for EA $eaName on $currency with $timeframe: $($_.Exception.Message)"
                    CaptureErrorState -errorContext "TestExecution"
                    
                    # Increment consecutive failures counter
                    $consecutiveFailures++
                    
                    # Try to close Strategy Tester if it's open
                    try {
                        $testerWindow = FindWindow -windowTitle $testerWindowTitle
                        if ($testerWindow -ne $null) {
                            SendKeysToWindow -window $testerWindow -keys "%{F4}"
                            LegacyAdaptiveWait -waitTime 2
                        }
                    }
                    catch {
                        # Ignore errors when closing
                    }
                    
                    # Check if we need to restart MT5 due to consecutive failures
                    if ($autoRestartOnFailure -and $consecutiveFailures -ge $maxConsecutiveFailures) {
                        LogMessage -level "WARN" -message "Detected $consecutiveFailures consecutive failures. Restarting MT5..."
                        
                        # Save checkpoint before restart
                        SaveCheckpoint
                        
                        # Close MT5
                        try {
                            $mt5Window = FindWindow -windowTitle "MetaTrader 5"
                            if ($mt5Window -ne $null) {
                                SendKeysToWindow -window $mt5Window -keys "%{F4}"
                                LegacyAdaptiveWait -waitTime 2
                            }
                            
                            # Handle potential "Save changes" dialog
                            $saveWindow = FindWindow -windowTitle "Save"
                            if ($saveWindow -ne $null) {
                                SendKeysToWindow -window $saveWindow -keys "n"  # Don't save changes
                                LegacyAdaptiveWait -waitTime 2
                            }
                            
                            # Make sure MT5 is closed
                            Stop-Process -Name "terminal64" -Force -ErrorAction SilentlyContinue
                            LegacyAdaptiveWait -waitTime 5
                            
                            # Restart MT5
                            Start-Process -FilePath $mt5Path
                            LegacyAdaptiveWait -waitTime $initialLoadTime
                            
                            # Reset consecutive failures counter
                            $consecutiveFailures = 0
                            
                            # Wait for MT5 to fully load
                            LegacyAdaptiveWait -waitTime 10
                            
                            # Re-open Strategy Tester
                            $mt5Window = FindWindow -windowTitle "MetaTrader 5"
                            if ($mt5Window -ne $null) {
                                SendKeysToWindow -window $mt5Window -keys "%t"
                                LegacyAdaptiveWait -waitTime 3
                            }
                            
                            LogMessage -level "INFO" -message "MT5 restarted successfully"
                        }
                        catch {
                            LogMessage -level "ERROR" -message "Failed to restart MT5: $($_.Exception.Message)"
                            CaptureErrorState -errorContext "RestartMT5"
                            
                            # Try to continue anyway
                            $consecutiveFailures = 0  # Reset to avoid infinite restart loop
                        }
                    }
                }
            }
            
            # Close Strategy Tester window after all timeframes are tested
            try {
                $testerWindow = FindWindow -windowTitle $testerWindowTitle
                if ($testerWindow -ne $null) {
                    SendKeysToWindow -window $testerWindow -keys "%{F4}"
                    LegacyAdaptiveWait -waitTime 2
                }
            }
            catch {
                # Ignore errors when closing
            }
        }
        catch {
            LogMessage -level "ERROR" -message "Error processing currency $currency for EA $eaName: $($_.Exception.Message)"
            CaptureErrorState -errorContext "ProcessCurrency"
            
            # Increment consecutive failures counter
            $consecutiveFailures++
            
            # Try to close Strategy Tester if it's open
            try {
                $testerWindow = FindWindow -windowTitle $testerWindowTitle
                if ($testerWindow -ne $null) {
                    SendKeysToWindow -window $testerWindow -keys "%{F4}"
                    LegacyAdaptiveWait -waitTime 2
                }
            }
            catch {
                # Ignore errors when closing
            }
            
            # Check if we need to restart MT5 due to consecutive failures
            if ($autoRestartOnFailure -and $consecutiveFailures -ge $maxConsecutiveFailures) {
                LogMessage -level "WARN" -message "Detected $consecutiveFailures consecutive failures. Restarting MT5..."
                
                # Save checkpoint before restart
                SaveCheckpoint
                
                # Close MT5
                try {
                    $mt5Window = FindWindow -windowTitle "MetaTrader 5"
                    if ($mt5Window -ne $null) {
                        SendKeysToWindow -window $mt5Window -keys "%{F4}"
                        LegacyAdaptiveWait -waitTime 2
                    }
                    
                    # Handle potential "Save changes" dialog
                    $saveWindow = FindWindow -windowTitle "Save"
                    if ($saveWindow -ne $null) {
                        SendKeysToWindow -window $saveWindow -keys "n"  # Don't save changes
                        LegacyAdaptiveWait -waitTime 2
                    }
                    
                    # Make sure MT5 is closed
                    Stop-Process -Name "terminal64" -Force -ErrorAction SilentlyContinue
                    LegacyAdaptiveWait -waitTime 5
                    
                    # Restart MT5
                    Start-Process -FilePath $mt5Path
                    LegacyAdaptiveWait -waitTime $initialLoadTime
                    
                    # Reset consecutive failures counter
                    $consecutiveFailures = 0
                    
                    # Wait for MT5 to fully load
                    LegacyAdaptiveWait -waitTime 10
                    
                    LogMessage -level "INFO" -message "MT5 restarted successfully"
                }
                catch {
                    LogMessage -level "ERROR" -message "Failed to restart MT5: $($_.Exception.Message)"
                    CaptureErrorState -errorContext "RestartMT5"
                    
                    # Try to continue anyway
                    $consecutiveFailures = 0  # Reset to avoid infinite restart loop
                }
            }
        }
    }
}

# Clean up after all tests are complete
try {
    # Close MT5
    $mt5Window = FindWindow -windowTitle "MetaTrader 5"
    if ($mt5Window -ne $null) {
                SendKeysToWindow -window $mt5Window -keys "%{F4}"
        LegacyAdaptiveWait -waitTime 2
    }
    
    # Handle potential "Save changes" dialog
    $saveWindow = FindWindow -windowTitle "Save"
    if ($saveWindow -ne $null) {
        SendKeysToWindow -window $saveWindow -keys "n"  # Don't save changes
        LegacyAdaptiveWait -waitTime 2
    }
    
    # Remove checkpoint file since we've completed all tests
    if (Test-Path $checkpointFile) {
        Remove-Item -Path $checkpointFile -Force
    }
    
    # Generate summary report
    try {
        $summaryReportPath = "$reportPath\backtest_summary_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').txt"
        
        $summaryContent = "=== BACKTEST AUTOMATION SUMMARY ===`r`n"
        $summaryContent += "`r`nCompleted at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`r`n"
        $summaryContent += "`r`nTotal reports generated: $($reportCounter - 1)`r`n"
        $summaryContent += "`r`nEAs tested: $($eaList.Count)`r`n"
        $summaryContent += "`r`nCurrency pairs tested: $($currencies.Count)`r`n"
        $summaryContent += "`r`nTimeframes tested: $($timeframes.Count)`r`n"
        $summaryContent += "`r`nDate range: $startDate to $endDate`r`n"
        
        Set-Content -Path $summaryReportPath -Value $summaryContent
        
        LogMessage -level "INFO" -message "Summary report generated" -details @{
            "path" = $summaryReportPath
        }
    }
    catch {
        LogMessage -level "ERROR" -message "Failed to generate summary report: $($_.Exception.Message)"
    }
    
    # Log completion
    LogMessage -level "INFO" -message "All backtests completed successfully" -details @{
        "totalReports" = $reportCounter - 1
        "completionTime" = $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    }
    
    # Display completion message
    [System.Windows.Forms.MessageBox]::Show("All backtests completed successfully. Generated $($reportCounter - 1) reports.", "Backtest Automation", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}
catch {
    LogMessage -level "ERROR" -message "Error during cleanup: $($_.Exception.Message)"
    
    # Display completion message with warning
    [System.Windows.Forms.MessageBox]::Show("Backtests completed with some errors. Check log file for details.", "Backtest Automation", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
}

# End of script
LogMessage -level "INFO" -message "Script execution completed"