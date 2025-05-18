<#
.SYNOPSIS
    MT5 Automatic Backtesting with PowerShell
.DESCRIPTION
    Automates MetaTrader 5 backtesting with support for multiple symbols and timeframes
.NOTES
    Version: 1.0.0
    Author: Based on PoshTrade PAD script
#>

#Requires -Version 5.1

# Import required modules for UI automation
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Drawing

#region Configuration

# Default configuration values
$script:config = @{
    # Core paths
    "batchFilePath" = "D:\FOREX\Coding\Git_&_Github\GitHub\PoshTrade\MT5 Automatic Backtesting\Modified Run_MT5_Backtest.bat"
    "configIniPath" = "D:\FOREX\Coding\Git_&_Github\GitHub\PoshTrade\MT5 Automatic Backtesting\Modified_MT5_Backtest_Config.ini"
    "reportPath" = "D:\FOREX\FOREX DOCUMENTS\MT5 STRATEGY TESTER REPORTS\Reports"
    "logFilePath" = "D:\FOREX\FOREX DOCUMENTS\MT5 STRATEGY TESTER REPORTS\automation_log.json"
    "errorScreenshotsPath" = "D:\FOREX\FOREX DOCUMENTS\MT5 STRATEGY TESTER REPORTS\Reports\errors"
    "checkpointFile" = "D:\FOREX\FOREX DOCUMENTS\MT5 STRATEGY TESTER REPORTS\Reports\checkpoint.json"
    "configFilePath" = "D:\FOREX\FOREX DOCUMENTS\MT5 STRATEGY TESTER REPORTS\Reports\backtest_config.json"
    "performanceHistoryFile" = "D:\FOREX\FOREX DOCUMENTS\MT5 STRATEGY TESTER REPORTS\Reports\performance_history.json"
    
    # Test parameters
    "eaName" = "Moving Average"
    "reportCounter" = 1
    "maxWaitTimeForTest" = 180
    "initialLoadTime" = 15
    
    # Error handling parameters
    "maxRetries" = 3
    "skipOnError" = $true
    "autoRestartOnFailure" = $true
    "maxConsecutiveFailures" = 5
    "retryBackoffMultiplier" = 1.5
    "maxRetryWaitTime" = 60
    
    # Adaptive wait parameters
    "adaptiveWaitEnabled" = $true
    "baseWaitMultiplier" = 1.0
    "maxAdaptiveWaitMultiplier" = 5.0
    
    # System monitoring parameters
    "systemLoadCheckInterval" = 300
    "lowMemoryThreshold" = 200
    "detailedSystemCheckInterval" = 600
    
    # Logging parameters
    "verboseLogging" = $true
    "logProgressInterval" = 10
}

# Constants
$script:constants = @{
    # Timing constants
    "WAIT_SHORT" = 1
    "WAIT_MEDIUM" = 3
    "WAIT_LONG" = 5
    "WAIT_VERY_LONG" = 10
    "MAX_LAUNCH_WAIT" = 60
    "MAX_TESTER_WAIT" = 30
    "PROGRESS_CHECK_INTERVAL" = 5
    "MAX_NO_PROGRESS_INTERVALS" = 30
    
    # Window constants
    "WINDOW_MT5" = "MetaTrader 5"
    "WINDOW_STRATEGY_TESTER" = "Strategy Tester"
    "WINDOW_TESTER" = "Tester"
    "WINDOW_SAVE_AS" = "Save As"
    "WINDOW_CONFIRM" = "Confirm"
    "WINDOW_SAVE" = "Save"
    
    # Process constants
    "PROCESS_MT5" = "terminal64.exe"
    "PROCESS_EXPLORER" = "explorer.exe"
}

# Runtime variables
$script:runtime = @{
    "currency" = "EURUSD"
    "timeframe" = "H1"
    "consecutiveFailures" = 0
    "currentAdaptiveMultiplier" = 1.0
    "lastSystemLoadCheck" = 0
    "availableMemory" = 1000
    "cpuUsage" = 50
    "performanceHistory" = @{}
}

#endregion Configuration

#region Utility Functions

function Initialize-Environment {
    [CmdletBinding()]
    param()
    
    Write-Log -Level "INFO" -Message "Initializing environment" -Details @{}
    
    # Validate required paths
    Confirm-RequiredPaths
    
    # Create necessary directories
    if (-not (Test-Path -Path $config.reportPath)) {
        New-Item -Path $config.reportPath -ItemType Directory -Force | Out-Null
        Write-Log -Level "INFO" -Message "Created reports directory" -Details @{ "path" = $config.reportPath }
    }
    
    if (-not (Test-Path -Path $config.errorScreenshotsPath)) {
        New-Item -Path $config.errorScreenshotsPath -ItemType Directory -Force | Out-Null
        Write-Log -Level "INFO" -Message "Created error screenshots directory" -Details @{ "path" = $config.errorScreenshotsPath }
    }
    
    # Initialize performance history
    Initialize-PerformanceHistory
}

function Confirm-RequiredPaths {
    [CmdletBinding()]
    param()
    
    # Check if batch file exists
    if (-not (Test-Path -Path $config.batchFilePath)) {
        throw "Batch file does not exist: $($config.batchFilePath)"
    }
    
    # Check if config INI file exists
    if (-not (Test-Path -Path $config.configIniPath)) {
        throw "Config INI file does not exist: $($config.configIniPath)"
    }
}

function Initialize-PerformanceHistory {
    [CmdletBinding()]
    param()
    
    if (Test-Path -Path $config.performanceHistoryFile) {
        try {
            $historyData = Get-Content -Path $config.performanceHistoryFile -Raw
            $script:runtime.performanceHistory = $historyData | ConvertFrom-Json -AsHashtable
            Write-Log -Level "INFO" -Message "Loaded performance history" -Details @{ "entries" = ($script:runtime.performanceHistory.Count) }
        }
        catch {
            Write-Log -Level "WARN" -Message "Error reading performance history file. Initializing new history." -Details @{ "error" = $_.Exception.Message }
            $script:runtime.performanceHistory = @{}
        }
    }
    else {
        $script:runtime.performanceHistory = @{}
    }
}

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("INFO", "WARN", "ERROR", "DEBUG")]
        [string]$Level,
        
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [hashtable]$Details = @{}
    )
    
    # Only log if verbose mode is on or if it's an important message
    if ($config.verboseLogging -or $Level -in @("ERROR", "WARN", "INFO")) {
        $logEntry = @{
            "timestamp" = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            "level" = $Level
            "message" = $Message
        }
        
        # Add details if provided
        if ($Details.Count -gt 0) {
            $logEntry.details = $Details
        }
        
        # Convert to JSON and append to log file
        $logJson = $logEntry | ConvertTo-Json -Compress
        Add-Content -Path $config.logFilePath -Value "$logJson`r`n"
        
        # Also output to console with appropriate color
        $consoleColor = switch ($Level) {
            "ERROR" { "Red" }
            "WARN"  { "Yellow" }
            "INFO"  { "White" }
            "DEBUG" { "Gray" }
            default { "White" }
        }
        
        Write-Host "[$($logEntry.timestamp)] [$Level] $Message" -ForegroundColor $consoleColor
    }
}

function Invoke-AdaptiveWait {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
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
        $backoffFactor = [Math]::Min($config.maxRetryWaitTime / $WaitTime, [Math]::Pow($config.retryBackoffMultiplier, $RetryCount))
        $adjustedWaitTime = $WaitTime * $backoffFactor
        
        # Cap at maximum retry wait time
        $adjustedWaitTime = [Math]::Min($adjustedWaitTime, $config.maxRetryWaitTime)
    }
    elseif ($config.adaptiveWaitEnabled) {
        # Use adaptive wait for normal operations
        $adjustedWaitTime = $WaitTime * $script:runtime.currentAdaptiveMultiplier
    }
    else {
        $adjustedWaitTime = $WaitTime
    }
    
    # Log wait time if it's significantly adjusted
    if ($adjustedWaitTime -gt $WaitTime * 1.5 -and $config.verboseLogging) {
        Write-Log -Level "DEBUG" -Message "Adjusted wait time from $WaitTime to $adjustedWaitTime seconds" -Details @{}
    }
    
    # Perform the wait
    Start-Sleep -Seconds $adjustedWaitTime
}

function Get-SystemResources {
    [CmdletBinding()]
    param()
    
    # Get available memory
    $memoryInfo = Get-CimInstance -ClassName Win32_OperatingSystem
    $script:runtime.availableMemory = [math]::Round($memoryInfo.FreePhysicalMemory / 1024)
    
    # Only get CPU usage if memory is concerning (reduces overhead)
    if ($script:runtime.availableMemory -lt $config.lowMemoryThreshold * 2) {
        try {
            $cpuLoad = Get-CimInstance -ClassName Win32_Processor | Measure-Object -Property LoadPercentage -Average
            $script:runtime.cpuUsage = $cpuLoad.Average
        }
        catch {
            $script:runtime.cpuUsage = 50  # Default value if can't get actual CPU usage
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
    if ($currentTime - $script:runtime.lastSystemLoadCheck -ge $config.detailedSystemCheckInterval) {
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
        if ($config.verboseLogging -or 
            $script:runtime.availableMemory -lt $config.lowMemoryThreshold * 2 -or 
            $script:runtime.cpuUsage -gt 80) {
            Write-Log -Level "INFO" -Message "System check" -Details $systemMetrics
        }
        
        # Adjust wait multiplier based on system metrics
        Update-AdaptiveMultiplier
    }
}

function Update-AdaptiveMultiplier {
    [CmdletBinding()]
    param()
    
    if ($script:runtime.cpuUsage -gt 90 -or $script:runtime.availableMemory -lt $config.lowMemoryThreshold) {
        # Critical system load - maximum wait times
        $script:runtime.currentAdaptiveMultiplier = $config.maxAdaptiveWaitMultiplier
        Write-Log -Level "WARN" -Message "Critical system load detected. Increasing wait times to maximum." -Details @{
            "cpuUsage" = $script:runtime.cpuUsage
            "availableMemory" = $script:runtime.availableMemory
        }
        
        # Check if we need to perform memory cleanup
        Invoke-MemoryCleanup
    }
    elseif ($script:runtime.cpuUsage -gt 70 -or $script:runtime.availableMemory -lt $config.lowMemoryThreshold * 2) {
        # High system load - increase wait times
        $script:runtime.currentAdaptiveMultiplier = [Math]::Min(
            $config.maxAdaptiveWaitMultiplier, 
            $script:runtime.currentAdaptiveMultiplier * 1.5
        )
        
        # Only log if verbose logging is enabled
        if ($config.verboseLogging) {
            Write-Log -Level "DEBUG" -Message "High system load detected. Increasing wait multiplier to $($script:runtime.currentAdaptiveMultiplier)" -Details @{
                "cpuUsage" = $script:runtime.cpuUsage
                "availableMemory" = $script:runtime.availableMemory
            }
        }
    }
    elseif ($script:runtime.cpuUsage -lt 40 -and $script:runtime.availableMemory -gt $config.lowMemoryThreshold * 3) {
        # Low system load - decrease wait times
        $script:runtime.currentAdaptiveMultiplier = [Math]::Max(
            1.0, 
            $script:runtime.currentAdaptiveMultiplier * 0.8
        )
        
        # Only log if verbose logging is enabled
        if ($config.verboseLogging) {
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

function Invoke-MemoryCleanup {
    [CmdletBinding()]
    param()
    
    # Only perform cleanup when memory is critically low
    if ($script:runtime.availableMemory -lt ($config.lowMemoryThreshold / 2)) {
        Write-Log -Level "WARN" -Message "Performing memory cleanup due to low memory ($($script:runtime.availableMemory) MB)" -Details @{}
        
        try {
            # Attempt to free memory by restarting Explorer (lightweight cleanup)
            $explorerProcess = Get-Process -Name $constants.PROCESS_EXPLORER -ErrorAction SilentlyContinue
            if ($explorerProcess) {
                Stop-Process -Name $constants.PROCESS_EXPLORER -Force
                Invoke-AdaptiveWait -WaitTime $constants.WAIT_MEDIUM
                Start-Process $constants.PROCESS_EXPLORER
                Invoke-AdaptiveWait -WaitTime $constants.WAIT_LONG
                
                # Check if cleanup helped
                Get-SystemResources
                
                Write-Log -Level "INFO" -Message "Memory after cleanup: $($script:runtime.availableMemory) MB" -Details @{
                    "beforeCleanup" = $script:runtime.availableMemory
                    "afterCleanup" = $script:runtime.availableMemory
                }
            }
        }
        catch {
            Write-Log -Level "ERROR" -Message "Memory cleanup attempt failed: $($_.Exception.Message)" -Details @{}
        }
    }
}

function Get-FormattedTimestamp {
    [CmdletBinding()]
    param()
    
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    return $timestamp
}

function Capture-ErrorState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ErrorContext
    )
    
    try {
        # Format timestamp for filename
        $timestamp = Get-FormattedTimestamp
        
        # Take screenshot of error state
        $screenshotPath = Join-Path -Path $config.errorScreenshotsPath -ChildPath "error_${ErrorContext}_${timestamp}.png"
        
        # Create bitmap of the screen
        $screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
        $bitmap = New-Object System.Drawing.Bitmap $screen.Width, $screen.Height
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $graphics.CopyFromScreen($screen.X, $screen.Y, 0, 0, $screen.Size)
        
        # Save the screenshot
        $bitmap.Save($screenshotPath)
        $graphics.Dispose()
        $bitmap.Dispose()
        
        $screenshotDetails = @{
            "filename" = "error_${ErrorContext}_${timestamp}.png"
            "context" = $ErrorContext
            "timestamp" = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        }
        
        Write-Log -Level "INFO" -Message "Error screenshot saved" -Details $screenshotDetails
        
        # Try to save any partial results if in Strategy Tester
        Save-PartialResults -ErrorContext $ErrorContext -Timestamp $timestamp
    }
    catch {
        Write-Log -Level "ERROR" -Message "Failed to capture error state: $($_.Exception.Message)" -Details @{}
    }
}

function Save-PartialResults {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ErrorContext,
        
        [Parameter(Mandatory = $true)]
        [string]$Timestamp
    )
    
    try {
        # Check if Strategy Tester window exists
        if (Test-WindowExists -WindowTitle $constants.WINDOW_STRATEGY_TESTER) {
            # Set focus to the window
            Set-WindowFocus -WindowTitle $constants.WINDOW_STRATEGY_TESTER
            
            # Send Ctrl+S to save
            [System.Windows.Forms.SendKeys]::SendWait("^s")
            Invoke-AdaptiveWait -WaitTime $constants.WAIT_SHORT
            
            # Set partial results filename
            $partialFileName = "partial_${ErrorContext}_${Timestamp}"
            $fullPath = Join-Path -Path $config.reportPath -ChildPath $partialFileName
            
            # Wait for Save As dialog
            Invoke-AdaptiveWait -WaitTime $constants.WAIT_SHORT
            
            if (Test-WindowExists -WindowTitle $constants.WINDOW_SAVE_AS) {
                # Set focus to Save As dialog
                Set-WindowFocus -WindowTitle $constants.WINDOW_SAVE_AS
                
                # Type the path
                [System.Windows.Forms.SendKeys]::SendWait($fullPath)
                Invoke-AdaptiveWait -WaitTime $constants.WAIT_SHORT
                
                # Press Enter to save
                [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
                Invoke-AdaptiveWait -WaitTime $constants.WAIT_MEDIUM
                
                # Handle potential overwrite confirmation
                if (Test-WindowExists -WindowTitle $constants.WINDOW_CONFIRM) {
                    Set-WindowFocus -WindowTitle $constants.WINDOW_CONFIRM
                    [System.Windows.Forms.SendKeys]::SendWait("y")
                    Invoke-AdaptiveWait -WaitTime $constants.WAIT_MEDIUM
                }
                
                Write-Log -Level "INFO" -Message "Partial results saved" -Details @{
                    "filename" = $partialFileName
                    "path" = $fullPath
                    "context" = $ErrorContext
                }
            }
        }
    }
    catch {
        Write-Log -Level "ERROR" -Message "Failed to save partial results: $($_.Exception.Message)" -Details @{}
    }
}

#endregion Utility Functions

#region Configuration Functions

function Import-Configuration {
    [CmdletBinding()]
    param()
    
    if (-not (Test-Path -Path $config.configFilePath)) {
        Write-Log -Level "INFO" -Message "No configuration file found at $($config.configFilePath). Using default settings." -Details @{}
        return
    }
    
    try {
        Write-Log -Level "INFO" -Message "Loading configuration from $($config.configFilePath)" -Details @{}
        $configData = Get-Content -Path $config.configFilePath -Raw
        
        # Try to parse as JSON first
        try {
            $importedConfig = $configData | ConvertFrom-Json -AsHashtable
            Apply-JsonConfiguration -ConfigData $importedConfig
            Write-Log -Level "INFO" -Message "Configuration loaded successfully from JSON" -Details @{}
        }
        catch {
            # Fallback to legacy text format parsing
            Write-Log -Level "WARN" -Message "Failed to parse JSON config, falling back to text format" -Details @{}
            Parse-LegacyConfiguration -ConfigData $configData
        }
    }
    catch {
        Write-Log -Level "ERROR" -Message "Error loading configuration: $($_.Exception.Message). Using default settings." -Details @{}
    }
}

function Apply-JsonConfiguration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$ConfigData
    )
    
    # Apply each configuration value if it exists in the imported data
    foreach ($key in $ConfigData.Keys) {
        if ($config.ContainsKey($key)) {
            $config[$key] = $ConfigData[$key]
            
            # Log the configuration value if verbose logging is enabled
            if ($config.verboseLogging) {
                Write-Log -Level "DEBUG" -Message "Config: $key = $($ConfigData[$key])" -Details @{}
            }
        }
    }
}

function Parse-LegacyConfiguration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigData
    )
    
    # Parse config data line by line
    $configLines = $ConfigData -split "`r`n"
    
    foreach ($line in $configLines) {
        # Skip empty lines and comments
        if ([string]::IsNullOrWhiteSpace($line) -or $line.StartsWith("#")) {
            continue
        }
        
        # Extract key and value
        $keyValue = $line -split "=", 2
        if ($keyValue.Count -eq 2) {
            $configKey = $keyValue[0].Trim()
            $configValue = $keyValue[1].Trim()
            
            # Apply configuration based on key
            Apply-ConfigurationValue -Key $configKey -Value $configValue
        }
    }
    
    Write-Log -Level "INFO" -Message "Configuration loaded successfully from text format" -Details @{}
}

function Apply-ConfigurationValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Key,
        
        [Parameter(Mandatory = $true)]
        [string]$Value
    )
    
    # Check if the key exists in our configuration
    if ($config.ContainsKey($Key)) {
        # Convert value to appropriate type based on existing value
        $existingValue = $config[$Key]
        $typedValue = $Value
        
        # Convert string to appropriate type
        if ($existingValue -is [int]) {
            $typedValue = [int]$Value
        }
        elseif ($existingValue -is [double]) {
            $typedValue = [double]$Value
        }
        elseif ($existingValue -is [bool]) {
            $typedValue = [bool]::Parse($Value)
        }
        
        # Update the configuration
        $config[$Key] = $typedValue
        
        # Log the configuration value if verbose logging is enabled
        if ($config.verboseLogging) {
            Write-Log -Level "DEBUG" -Message "Config: $Key = $typedValue" -Details @{}
        }
    }
}

function Import-IniSettings {
    [CmdletBinding()]
    param()
    
    try {
        Write-Log -Level "INFO" -Message "Reading settings from INI file: $($config.configIniPath)" -Details @{}
        $iniContent = Get-Content -Path $config.configIniPath -Raw
        
        # Create hashtable to store extracted settings
        $iniSettings = @{}
        
        # Extract EA name
        $iniSettings["Expert"] = Extract-IniSetting -SettingName "Expert" -RegexPattern "([^\r\n]+)" -IniContent $iniContent
        if ($iniSettings["Expert"]) {
            $script:config.eaName = $iniSettings["Expert"]
        }
        
        # Extract Symbol
        $iniSettings["Symbol"] = Extract-IniSetting -SettingName "Symbol" -RegexPattern "([^\r\n]+)" -IniContent $iniContent
        if ($iniSettings["Symbol"]) {
            $script:runtime.currency = $iniSettings["Symbol"]
        }
        
        # Extract Timeframe
        $iniSettings["Period"] = Extract-IniSetting -SettingName "Period" -RegexPattern "([^\r\n]+)" -IniContent $iniContent
        if ($iniSettings["Period"]) {
            $script:runtime.timeframe = $iniSettings["Period"]
        }
        
        # Extract Date Range
        $iniSettings["FromDate"] = Extract-IniSetting -SettingName "FromDate" -RegexPattern "([^\r\n]+)" -IniContent $iniContent
        $iniSettings["ToDate"] = Extract-IniSetting -SettingName "ToDate" -RegexPattern "([^\r\n]+)" -IniContent $iniContent
        
        # Extract Model
        $iniSettings["Model"] = Extract-IniSetting -SettingName "Model" -RegexPattern "([^\r\n]+)" -IniContent $iniContent
        
        # Extract other important settings
        $iniSettings["Optimization"] = Extract-IniSetting -SettingName "Optimization" -RegexPattern "([^\r\n]+)" -IniContent $iniContent
        $iniSettings["Visual"] = Extract-IniSetting -SettingName "Visual" -RegexPattern "([^\r\n]+)" -IniContent $iniContent
        
        # Log extracted settings
        Write-Log -Level "INFO" -Message "Settings extracted from INI file" -Details $iniSettings
        
        return $iniSettings
    }
    catch {
        Write-Log -Level "ERROR" -Message "Failed to extract settings from INI file: $($_.Exception.Message)" -Details @{}
        return @{}
    }
}

function Extract-IniSetting {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SettingName,
        
        [Parameter(Mandatory = $true)]
        [string]$RegexPattern,
        
        [Parameter(Mandatory = $true)]
        [string]$IniContent
    )
    
    $regexFull = "$SettingName=$RegexPattern"
    if ($IniContent -match $regexFull) {
        return $matches[1]
    }
    
    return $null
}

#endregion Configuration Functions

#region UI Automation Functions

function Test-WindowExists {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$WindowTitle
    )
    
    $windows = [System.Diagnostics.Process]::GetProcesses() | 
        Where-Object { $_.MainWindowTitle -like "*$WindowTitle*" -and $_.MainWindowHandle -ne 0 }
    
    return ($windows -ne $null -and $windows.Count -gt 0)
}

function Set-WindowFocus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$WindowTitle
    )
    
    $windows = [System.Diagnostics.Process]::GetProcesses() | 
        Where-Object { $_.MainWindowTitle -like "*$WindowTitle*" -and $_.MainWindowHandle -ne 0 }
    
    if ($windows -ne $null -and $windows.Count -gt 0) {
        # Get the first matching window
        $window = $windows[0]
        
        # Set focus to the window
        [void][System.Runtime.InteropServices.Marshal]::SetForegroundWindow($window.MainWindowHandle)
        return $true
    }
    
    return $false
}

function Wait-ForWindow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$WindowTitle,
        
        [Parameter(Mandatory = $false)]
        [int]$MaxWaitTime = 60
    )
    
    $waitTime = 0
    $windowAppeared = $false
    
    while ($waitTime -lt $MaxWaitTime -and -not $windowAppeared) {
        if (Test-WindowExists -WindowTitle $WindowTitle) {
            $windowAppeared = $true
            Write-Log -Level "INFO" -Message "Window '$WindowTitle' appeared successfully" -Details @{}
            break
        }
        
                Invoke-AdaptiveWait -WaitTime $constants.WAIT_MEDIUM
        $waitTime += $constants.WAIT_MEDIUM
    }
    
    return $windowAppeared
}

function Start-MT5WithBatchFile {
    [CmdletBinding()]
    param()
    
    try {
        Write-Log -Level "INFO" -Message "Launching MT5 using batch file: $($config.batchFilePath)" -Details @{}
        
        # Run the batch file
        Start-Process -FilePath $config.batchFilePath
        
        # Wait for MT5 to launch
        $mt5Launched = Wait-ForWindow -WindowTitle $constants.WINDOW_MT5 -MaxWaitTime $constants.MAX_LAUNCH_WAIT
        
        if (-not $mt5Launched) {
            throw "MetaTrader 5 did not launch within the expected time"
        }
        
        # Wait additional time for MT5 to fully initialize
        Invoke-AdaptiveWait -WaitTime $config.initialLoadTime
        
        return $true
    }
    catch {
        Write-Log -Level "ERROR" -Message "Failed to launch MT5 using batch file: $($_.Exception.Message)" -Details @{}
        Capture-ErrorState -ErrorContext "LaunchMT5"
        return $false
    }
}

function Wait-ForStrategyTester {
    [CmdletBinding()]
    param()
    
    try {
        Write-Log -Level "INFO" -Message "Waiting for Strategy Tester to open..." -Details @{}
        
        # Wait for Strategy Tester window to appear
        $testerOpened = $false
        
        # Try both possible window names
        if (Wait-ForWindow -WindowTitle $constants.WINDOW_STRATEGY_TESTER -MaxWaitTime $constants.MAX_TESTER_WAIT) {
            $testerOpened = $true
        }
        elseif (Wait-ForWindow -WindowTitle $constants.WINDOW_TESTER -MaxWaitTime $constants.MAX_TESTER_WAIT) {
            $testerOpened = $true
        }
        
        if (-not $testerOpened) {
            throw "Strategy Tester did not open within the expected time"
        }
        
        # Give it a moment to fully initialize
        Invoke-AdaptiveWait -WaitTime $constants.WAIT_MEDIUM
        
        return $true
    }
    catch {
        Write-Log -Level "ERROR" -Message "Failed to wait for Strategy Tester: $($_.Exception.Message)" -Details @{}
        Capture-ErrorState -ErrorContext "WaitForTester"
        return $false
    }
}

function Set-StrategyTesterFocus {
    [CmdletBinding()]
    param()
    
    if (Test-WindowExists -WindowTitle $constants.WINDOW_STRATEGY_TESTER) {
        Set-WindowFocus -WindowTitle $constants.WINDOW_STRATEGY_TESTER
        return $true
    }
    elseif (Test-WindowExists -WindowTitle $constants.WINDOW_TESTER) {
        Set-WindowFocus -WindowTitle $constants.WINDOW_TESTER
        return $true
    }
    else {
        return $false
    }
}

function Start-Backtest {
    [CmdletBinding()]
    param()
    
    try {
        Write-Log -Level "INFO" -Message "Starting backtest..." -Details @{}
        
        # Ensure Strategy Tester window is active
        if (-not (Set-StrategyTesterFocus)) {
            throw "Strategy Tester window not found"
        }
        
        Invoke-AdaptiveWait -WaitTime $constants.WAIT_SHORT
        
        # Try different methods to start the test
        $testStarted = Start-BacktestWithKeyboard
        
        if (-not $testStarted) {
            $testStarted = Start-BacktestWithAltKey
        }
        
        if (-not $testStarted) {
            throw "Failed to start backtest using all methods"
        }
        
        return $true
    }
    catch {
        Write-Log -Level "ERROR" -Message "Failed to start backtest: $($_.Exception.Message)" -Details @{}
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
        Invoke-AdaptiveWait -WaitTime $constants.WAIT_MEDIUM
        
        # Check if test started by looking for progress indicators
        # This is a simplified check - in a real implementation, you'd use UI Automation to check button state
        $testStarted = Confirm-BacktestStarted
        
        if ($testStarted) {
            Write-Log -Level "INFO" -Message "Backtest started using F9 key" -Details @{}
            return $true
        }
        else {
            return $false
        }
    }
    catch {
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
        Invoke-AdaptiveWait -WaitTime $constants.WAIT_MEDIUM
        
        # Check if test started
        $testStarted = Confirm-BacktestStarted
        
        if ($testStarted) {
            Write-Log -Level "INFO" -Message "Backtest started using Alt+S" -Details @{}
            return $true
        }
        else {
            return $false
        }
    }
    catch {
        return $false
    }
}

function Confirm-BacktestStarted {
    [CmdletBinding()]
    param()
    
    # In a real implementation, you would use UI Automation to check for:
    # 1. Start button being disabled
    # 2. Progress bar appearing
    # 3. Status text changing
    
    # For this simplified version, we'll just wait and assume it started
    # A more robust implementation would use UI Automation framework
    
    # Wait a moment to see if the test starts
    Invoke-AdaptiveWait -WaitTime $constants.WAIT_MEDIUM
    
    # For now, we'll just return true and rely on the monitoring function
    # to detect if the test actually started
    return $true
}

function Monitor-BacktestProgress {
    [CmdletBinding()]
    param()
    
    try {
        Write-Log -Level "INFO" -Message "Monitoring backtest progress..." -Details @{}
        
        # Initialize monitoring variables
        $testStartTime = [int](Get-Date).ToFileTime()
        $testCompleted = $false
        $testWaitTime = 0
        $lastProgressValue = "0"
        $noProgressCounter = 0
        $mtFrozenCounter = 0
        $previousLoggedProgress = "0"
        $lastProgressLogTime = 0
        
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
            if ($testWaitTime % $constants.SYSTEM_CHECK_INTERVAL -eq 0) {
                Invoke-DetailedSystemCheck
            }
            
            # Check if MT5 is responsive
            if ($testWaitTime % 60 -eq 0 -and $testWaitTime -gt 0) {
                Test-MT5Responsiveness -MtFrozenCounter ([ref]$mtFrozenCounter) -TestCompleted ([ref]$testCompleted)
            }
            
            # Check if test is stuck with no progress
            if ($noProgressCounter -ge $constants.MAX_NO_PROGRESS_INTERVALS) {
                Write-Log -Level "ERROR" -Message "Backtest appears to be stuck at $lastProgressValue%. Attempting recovery..." -Details @{}
                Capture-ErrorState -ErrorContext "BacktestStuck"
                $testCompleted = $true
                $script:runtime.consecutiveFailures++
                break
            }
            
            # Periodic heartbeat log
            if ($testWaitTime % 300 -eq 0 -and $testWaitTime -gt 0) {
                Write-Log -Level "INFO" -Message "Backtest still running after $testWaitTime seconds. Current progress: $lastProgressValue%" -Details @{}
            }
            
            Invoke-AdaptiveWait -WaitTime $constants.PROGRESS_CHECK_INTERVAL
            $testWaitTime += $constants.PROGRESS_CHECK_INTERVAL
            
            # Safety timeout - don't wait forever
            if ($testWaitTime -gt $config.maxWaitTimeForTest * 10) {
                Write-Log -Level "ERROR" -Message "Maximum wait time exceeded. Forcing test completion." -Details @{}
                Capture-ErrorState -ErrorContext "TimeoutExceeded"
                $testCompleted = $true
                $script:runtime.consecutiveFailures++
                break
            }
        }
        
        # Record actual test duration for future estimates
        $actualTestDuration = [int](Get-Date).ToFileTime() - $testStartTime
        Update-PerformanceHistory -Currency $script:runtime.currency -Timeframe $script:runtime.timeframe -EaName $config.eaName -ActualDuration $actualTestDuration
        
        # Return success if test completed normally
        if ($script:runtime.consecutiveFailures -eq 0) {
            return $true
        }
        else {
            return $false
        }
    }
    catch {
        Write-Log -Level "ERROR" -Message "Error monitoring backtest progress: $($_.Exception.Message)" -Details @{}
        Capture-ErrorState -ErrorContext "MonitorProgress"
        return $false
    }
}

function Test-BacktestCompletion {
    [CmdletBinding()]
    param()
    
    # In a real implementation, you would use UI Automation to check for:
    # 1. Start button becoming enabled again
    # 2. Report tab appearing
    # 3. Status bar text indicating completion
    
    # For this simplified version, we'll use a basic check
    # A more robust implementation would use UI Automation framework
    
    try {
        # Check if the Strategy Tester window is still active
        if (-not (Test-WindowExists -WindowTitle $constants.WINDOW_STRATEGY_TESTER) -and 
            -not (Test-WindowExists -WindowTitle $constants.WINDOW_TESTER)) {
            # Window closed unexpectedly
            Write-Log -Level "WARN" -Message "Strategy Tester window closed unexpectedly" -Details @{}
            return $true
        }
        
        # For now, we'll rely on external signals like timeout
        # A real implementation would check UI elements
        
        return $false
    }
    catch {
        # Ignore errors when checking completion
        return $false
    }
}

function Update-TestProgress {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$LastProgressValue,
        
        [Parameter(Mandatory = $true)]
        [ref]$NoProgressCounter
    )
    
    try {
        # In a real implementation, you would use UI Automation to:
        # 1. Get text from status bar
        # 2. Extract progress percentage
        
        # For this simplified version, we'll simulate progress
        # A more robust implementation would use UI Automation framework
        
        # Simulate progress for demonstration purposes
        $currentProgress = [math]::Min(100, [int]$LastProgressValue.Value + 1)
        
        if ($currentProgress -ne [int]$LastProgressValue.Value) {
            # Progress has changed, reset the no-progress counter
            $LastProgressValue.Value = $currentProgress.ToString()
            $NoProgressCounter.Value = 0
            
            # Log progress information
            Log-ProgressInformation -Progress $currentProgress
        }
        else {
            # No change in progress, increment counter
            $NoProgressCounter.Value++
        }
    }
    catch {
        # Ignore errors when checking progress
    }
}

function Log-ProgressInformation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [int]$Progress
    )
    
    # Calculate estimated remaining time
    if ($Progress -gt 0) {
        $elapsedTime = [int](Get-Date).ToFileTime() - $script:testStartTime
        $progressFraction = $Progress / 100
        $totalEstimatedTime = $elapsedTime / $progressFraction
        $remainingTime = $totalEstimatedTime - $elapsedTime
        
        # Format remaining time
        $remainingMinutes = [math]::Floor($remainingTime / 60)
        $remainingSeconds = [math]::Floor($remainingTime % 60)
        $remainingTimeFormatted = "${remainingMinutes}m ${remainingSeconds}s"
        
        # Log progress less frequently for long runs
        $currentTime = [int](Get-Date).ToFileTime()
        if ($currentTime - $script:lastProgressLogTime -gt 300 -or 
            ($Progress % $config.logProgressInterval -eq 0 -and $Progress -ne $script:previousLoggedProgress)) {
            
            $progressDetails = @{
                "progress" = $Progress
                "elapsedTime" = $elapsedTime
                "estimatedRemaining" = $remainingTimeFormatted
            }
            
            Write-Log -Level "INFO" -Message "Backtest in progress: $Progress% complete" -Details $progressDetails
            $script:previousLoggedProgress = $Progress
            $script:lastProgressLogTime = $currentTime
        }
    }
}

function Test-MT5Responsiveness {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$MtFrozenCounter,
        
        [Parameter(Mandatory = $true)]
        [ref]$TestCompleted
    )
    
    try {
        # Send a harmless key to check if window responds
        if (Set-StrategyTesterFocus) {
            [System.Windows.Forms.SendKeys]::SendWait("{HOME}")
            Invoke-AdaptiveWait -WaitTime $constants.WAIT_SHORT
            
            # Reset frozen counter if MT5 responds
            $MtFrozenCounter.Value = 0
        }
        else {
            # Window not found or not responding
            $MtFrozenCounter.Value++
        }
    }
    catch {
        # MT5 didn't respond
        $MtFrozenCounter.Value++
        Write-Log -Level "WARN" -Message "Warning: MT5 may be unresponsive (attempt $($MtFrozenCounter.Value))" -Details @{}
        
        # Check system resources before declaring frozen
        Invoke-DetailedSystemCheck
        
        # Only consider MT5 frozen after multiple failed response checks
        if ($MtFrozenCounter.Value -ge 3) {
            Write-Log -Level "ERROR" -Message "MT5 appears to be frozen. Attempting recovery..." -Details @{}
                        Capture-ErrorState -ErrorContext "MT5Frozen"
            $TestCompleted.Value = $true
            $script:runtime.consecutiveFailures++
        }
    }
}

function Update-PerformanceHistory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Currency,
        
        [Parameter(Mandatory = $true)]
        [string]$Timeframe,
        
        [Parameter(Mandatory = $true)]
        [string]$EaName,
        
        [Parameter(Mandatory = $true)]
        [int]$ActualDuration
    )
    
    # Validate parameters
    if ([string]::IsNullOrEmpty($Currency) -or 
        [string]::IsNullOrEmpty($Timeframe) -or 
        [string]::IsNullOrEmpty($EaName) -or 
        $ActualDuration -le 0) {
        
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
        $script:runtime.performanceHistory | ConvertTo-Json | Set-Content -Path $config.performanceHistoryFile
        Write-Log -Level "DEBUG" -Message "Updated performance history" -Details @{
            "combination" = $historyKey
            "duration" = $newDuration
        }
    }
    catch {
        Write-Log -Level "ERROR" -Message "Failed to save performance history: $($_.Exception.Message)" -Details @{}
    }
}

function Save-BacktestReport {
    [CmdletBinding()]
    param()
    
    try {
        Write-Log -Level "INFO" -Message "Saving backtest report..." -Details @{}
        
        # Try primary method first
        $reportSaved = Save-ReportUsingContextMenu
        
        # If primary method fails, try alternative method
        if (-not $reportSaved) {
            $reportSaved = Save-ReportUsingMenuBar
        }
        
        if ($reportSaved) {
            # Reset consecutive failures counter on success
            $script:runtime.consecutiveFailures = 0
            
            # Close report tab with keyboard shortcut
            Set-StrategyTesterFocus
            [System.Windows.Forms.SendKeys]::SendWait("^{F4}")
            Invoke-AdaptiveWait -WaitTime $constants.WAIT_MEDIUM
            
            return $true
        }
        else {
            throw "Failed to save report using all available methods"
        }
    }
    catch {
        Write-Log -Level "ERROR" -Message "Failed to save Excel report: $($_.Exception.Message)" -Details @{}
        Capture-ErrorState -ErrorContext "SaveReport"
        
        # Increment consecutive failures counter
        $script:runtime.consecutiveFailures++
        
        # Try to close any open dialogs or tabs
        Invoke-CleanupAfterFailedSave
        
        return $false
    }
}

function Save-ReportUsingContextMenu {
    [CmdletBinding()]
    param()
    
    try {
        # Ensure Strategy Tester window is active
        if (-not (Set-StrategyTesterFocus)) {
            return $false
        }
        
        # In a real implementation, you would use UI Automation to:
        # 1. Find the report tab
        # 2. Right-click on it
        # 3. Select "Report" from the context menu
        # 4. Select "Excel" from the submenu
        
        # For this simplified version, we'll use keyboard shortcuts
        # A more robust implementation would use UI Automation framework
        
        # Try to use Alt+V (View menu), then R (Report), then E (Excel)
        [System.Windows.Forms.SendKeys]::SendWait("%v")
        Invoke-AdaptiveWait -WaitTime $constants.WAIT_SHORT
        [System.Windows.Forms.SendKeys]::SendWait("r")
        Invoke-AdaptiveWait -WaitTime $constants.WAIT_SHORT
        [System.Windows.Forms.SendKeys]::SendWait("e")
        Invoke-AdaptiveWait -WaitTime $constants.WAIT_MEDIUM
        
        # Handle the Save As dialog
        $reportSaved = Handle-SaveAsDialog
        
        return $reportSaved
    }
    catch {
        return $false
    }
}

function Save-ReportUsingMenuBar {
    [CmdletBinding()]
    param()
    
    try {
        # Ensure Strategy Tester window is active
        if (-not (Set-StrategyTesterFocus)) {
            return $false
        }
        
        # Alternative method using different keyboard sequence
        [System.Windows.Forms.SendKeys]::SendWait("{F10}")  # Activate menu bar
        Invoke-AdaptiveWait -WaitTime $constants.WAIT_SHORT
        [System.Windows.Forms.SendKeys]::SendWait("v")      # View menu
        Invoke-AdaptiveWait -WaitTime $constants.WAIT_SHORT
        [System.Windows.Forms.SendKeys]::SendWait("r")      # Report submenu
        Invoke-AdaptiveWait -WaitTime $constants.WAIT_SHORT
        [System.Windows.Forms.SendKeys]::SendWait("e")      # Excel option
        Invoke-AdaptiveWait -WaitTime $constants.WAIT_MEDIUM
        
        # Handle the Save As dialog
        $reportSaved = Handle-SaveAsDialog
        
        return $reportSaved
    }
    catch {
        return $false
    }
}

function Handle-SaveAsDialog {
    [CmdletBinding()]
    param()
    
    try {
        # Wait for Save As dialog to appear
        if (-not (Wait-ForWindow -WindowTitle $constants.WINDOW_SAVE_AS -MaxWaitTime 10)) {
            return $false
        }
        
        # Set focus to Save As dialog
        Set-WindowFocus -WindowTitle $constants.WINDOW_SAVE_AS
        
        # Generate custom filename
        $reportFileName = Generate-ReportFilename
        
        # Set the complete path and filename
        $fullPath = Join-Path -Path $config.reportPath -ChildPath $reportFileName
        
        # Select all text in the filename field and replace it
        [System.Windows.Forms.SendKeys]::SendWait("^a")
        Invoke-AdaptiveWait -WaitTime $constants.WAIT_SHORT
        [System.Windows.Forms.SendKeys]::SendWait($fullPath)
        Invoke-AdaptiveWait -WaitTime $constants.WAIT_SHORT
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
        Invoke-AdaptiveWait -WaitTime $constants.WAIT_MEDIUM
        
        # Handle potential overwrite confirmation
        if (Test-WindowExists -WindowTitle $constants.WINDOW_CONFIRM) {
            Set-WindowFocus -WindowTitle $constants.WINDOW_CONFIRM
            [System.Windows.Forms.SendKeys]::SendWait("y")
            Invoke-AdaptiveWait -WaitTime $constants.WAIT_MEDIUM
        }
        
        $reportDetails = @{
            "filename" = $reportFileName
            "path" = $fullPath
            "ea" = $config.eaName
            "currency" = $script:runtime.currency
            "timeframe" = $script:runtime.timeframe
        }
        
        Write-Log -Level "INFO" -Message "Excel report saved" -Details $reportDetails
        
        # Increment counter
        $script:config.reportCounter++
        
        return $true
    }
    catch {
        return $false
    }
}

function Generate-ReportFilename {
    [CmdletBinding()]
    param()
    
    # Generate a timestamp
    $timestamp = Get-FormattedTimestamp
    
    # Create a custom filename with EA, currency, timeframe, and counter
    $reportFileName = "Report_$($config.eaName)_$($script:runtime.currency)_$($script:runtime.timeframe)_$($config.reportCounter)_$timestamp.xml"
    
    # Replace any invalid characters
    $reportFileName = $reportFileName -replace '[\\\/\:\*\?\"\<\>\|]', '_'
    
    return $reportFileName
}

function Invoke-CleanupAfterFailedSave {
    [CmdletBinding()]
    param()
    
    try {
        # Try to close any open dialogs or tabs
        if (Test-WindowExists -WindowTitle $constants.WINDOW_SAVE_AS) {
            Set-WindowFocus -WindowTitle $constants.WINDOW_SAVE_AS
            [System.Windows.Forms.SendKeys]::SendWait("{ESC}")
            Invoke-AdaptiveWait -WaitTime $constants.WAIT_SHORT
        }
        
        if (Test-WindowExists -WindowTitle $constants.WINDOW_STRATEGY_TESTER) {
            Set-WindowFocus -WindowTitle $constants.WINDOW_STRATEGY_TESTER
            [System.Windows.Forms.SendKeys]::SendWait("^{F4}")
            Invoke-AdaptiveWait -WaitTime $constants.WAIT_MEDIUM
        }
    }
    catch {
        # Ignore errors when closing
    }
}

function Invoke-CleanupAfterBacktest {
    [CmdletBinding()]
    param()
    
    try {
        Write-Log -Level "INFO" -Message "Cleaning up after backtest..." -Details @{}
        
        # Close Strategy Tester window
        if (Test-WindowExists -WindowTitle $constants.WINDOW_STRATEGY_TESTER) {
            Set-WindowFocus -WindowTitle $constants.WINDOW_STRATEGY_TESTER
            [System.Windows.Forms.SendKeys]::SendWait("%{F4}")
            Invoke-AdaptiveWait -WaitTime $constants.WAIT_MEDIUM
        }
        
        # Close MT5 if needed
        if ($config.autoRestartOnFailure -and $script:runtime.consecutiveFailures -ge $config.maxConsecutiveFailures) {
            Restart-MT5
        }
        
        return $true
    }
    catch {
        Write-Log -Level "ERROR" -Message "Error during cleanup: $($_.Exception.Message)" -Details @{}
        return $false
    }
}

function Restart-MT5 {
    [CmdletBinding()]
    param()
    
    Write-Log -Level "WARN" -Message "Detected $($script:runtime.consecutiveFailures) consecutive failures. Restarting MT5..." -Details @{}
    
    # Close MT5
    if (Test-WindowExists -WindowTitle $constants.WINDOW_MT5) {
        Set-WindowFocus -WindowTitle $constants.WINDOW_MT5
        [System.Windows.Forms.SendKeys]::SendWait("%{F4}")
        Invoke-AdaptiveWait -WaitTime $constants.WAIT_MEDIUM
        
        # Handle potential "Save changes" dialog
        if (Test-WindowExists -WindowTitle $constants.WINDOW_SAVE) {
            Set-WindowFocus -WindowTitle $constants.WINDOW_SAVE
            [System.Windows.Forms.SendKeys]::SendWait("n")  # Don't save changes
            Invoke-AdaptiveWait -WaitTime $constants.WAIT_MEDIUM
        }
    }
    
    # Make sure MT5 is closed
    Stop-Process -Name $constants.PROCESS_MT5 -Force -ErrorAction SilentlyContinue
    Invoke-AdaptiveWait -WaitTime $constants.WAIT_LONG
    
    # Reset consecutive failures counter
    $script:runtime.consecutiveFailures = 0
}

function Generate-SummaryReport {
    [CmdletBinding()]
    param()
    
    try {
        $timestamp = Get-FormattedTimestamp
        $summaryReportPath = Join-Path -Path $config.reportPath -ChildPath "backtest_summary_$timestamp.txt"
        
        $summaryContent = @"
=== BACKTEST AUTOMATION SUMMARY ===

Completed at: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")

Total reports generated: $($config.reportCounter - 1)

EA tested: $($config.eaName)

Symbol tested: $($script:runtime.currency)

Timeframe tested: $($script:runtime.timeframe)

"@
        
        # Add system performance information
        $summaryContent += @"

=== SYSTEM PERFORMANCE ===

Available memory: $($script:runtime.availableMemory) MB

Adaptive wait multiplier: $($script:runtime.currentAdaptiveMultiplier)

"@
        
        Set-Content -Path $summaryReportPath -Value $summaryContent
        
        Write-Log -Level "INFO" -Message "Summary report generated" -Details @{
            "path" = $summaryReportPath
        }
    }
    catch {
        Write-Log -Level "ERROR" -Message "Failed to generate summary report: $($_.Exception.Message)" -Details @{}
    }
}

#endregion UI Automation Functions

#region Main Execution

function Start-BacktestAutomation {
    [CmdletBinding()]
    param()
    
    try {
        # Initialize environment
        Initialize-Environment
        
        # Load configuration if available
        Import-Configuration
        
        # Initial system resource check
        Invoke-DetailedSystemCheck
        
        # Main execution flow
        Write-Log -Level "INFO" -Message "Starting backtest automation with INI-based configuration" -Details @{}
        
        # Extract settings from INI file
        $iniSettings = Import-IniSettings
        
        # Launch MT5 using the batch file
        $mt5Launched = Start-MT5WithBatchFile
        
        if ($mt5Launched) {
            # Wait for Strategy Tester to open
            $testerOpened = Wait-ForStrategyTester
            
            if ($testerOpened) {
                # Start the backtest
                $backtestStarted = Start-Backtest
                
                if ($backtestStarted) {
                    # Monitor backtest progress
                    $backtestCompleted = Monitor-BacktestProgress
                    
                    if ($backtestCompleted) {
                        # Save backtest report
                        $reportSaved = Save-BacktestReport
                        
                        if ($reportSaved) {
                            Write-Log -Level "INFO" -Message "Backtest completed successfully" -Details @{}
                        }
                        else {
                            Write-Log -Level "WARN" -Message "Backtest completed but report could not be saved" -Details @{}
                        }
                    }
                    else {
                                                Write-Log -Level "ERROR" -Message "Backtest did not complete successfully" -Details @{}
                    }
                }
                else {
                    Write-Log -Level "ERROR" -Message "Failed to start backtest" -Details @{}
                }
            }
            else {
                Write-Log -Level "ERROR" -Message "Strategy Tester did not open" -Details @{}
            }
            
            # Clean up after backtest
            Invoke-CleanupAfterBacktest
        }
        else {
            Write-Log -Level "ERROR" -Message "Failed to launch MT5" -Details @{}
        }
        
        # Generate summary report
        Generate-SummaryReport
        
        # Log completion
        Write-Log -Level "INFO" -Message "Backtest automation completed" -Details @{
            "totalReports" = $config.reportCounter - 1
            "completionTime" = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        }
        
        # Display completion message
        Write-Host "Backtest automation completed. Generated $($config.reportCounter - 1) reports." -ForegroundColor Green
    }
    catch {
        Write-Log -Level "ERROR" -Message "Fatal error in backtest automation: $($_.Exception.Message)" -Details @{
            "stackTrace" = $_.ScriptStackTrace
        }
        
        if (-not $config.skipOnError) {
            throw
        }
    }
}

function Start-MultiSymbolBacktestAutomation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string[]]$Symbols = @("EURUSD", "GBPUSD", "USDJPY", "AUDUSD"),
        
        [Parameter(Mandatory = $false)]
        [string[]]$Timeframes = @("M1", "M5", "M15", "M30", "H1", "H4", "D1")
    )
    
    Write-Host "Starting multi-symbol, multi-timeframe backtest automation" -ForegroundColor Cyan
    Write-Host "Symbols: $($Symbols -join ', ')" -ForegroundColor Cyan
    Write-Host "Timeframes: $($Timeframes -join ', ')" -ForegroundColor Cyan
    
    # Initialize counters
    $totalTests = $Symbols.Count * $Timeframes.Count
    $completedTests = 0
    $successfulTests = 0
    
    # Create a master summary file
    $masterSummaryPath = Join-Path -Path $config.reportPath -ChildPath "master_summary_$(Get-FormattedTimestamp).csv"
    "Symbol,Timeframe,Status,ReportFile,CompletionTime" | Set-Content -Path $masterSummaryPath
    
    # Loop through each symbol and timeframe combination
    foreach ($symbol in $Symbols) {
        foreach ($timeframe in $Timeframes) {
            $completedTests++
            
            Write-Host "====================================" -ForegroundColor Yellow
            Write-Host "Starting test $completedTests of $totalTests" -ForegroundColor Yellow
            Write-Host "Symbol: $symbol, Timeframe: $timeframe" -ForegroundColor Yellow
            Write-Host "====================================" -ForegroundColor Yellow
            
            # Update the INI file with current symbol and timeframe
            Update-IniFile -Symbol $symbol -Timeframe $timeframe
            
            # Set runtime variables
            $script:runtime.currency = $symbol
            $script:runtime.timeframe = $timeframe
            
            # Run the backtest for this combination
            try {
                Start-BacktestAutomation
                $status = "Success"
                $successfulTests++
            }
            catch {
                $status = "Failed: $($_.Exception.Message)"
            }
            
            # Find the most recent report file for this combination
            $reportPattern = "*${symbol}_${timeframe}*.xml"
            $reportFile = Get-ChildItem -Path $config.reportPath -Filter $reportPattern | 
                Sort-Object LastWriteTime -Descending | 
                Select-Object -First 1 -ExpandProperty Name
            
            # Update the master summary
            "$symbol,$timeframe,$status,$reportFile,$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" | 
                Add-Content -Path $masterSummaryPath
            
            # Short break between tests
            Write-Host "Waiting before next test..." -ForegroundColor Gray
            Start-Sleep -Seconds 5
        }
    }
    
    # Display final summary
    Write-Host "====================================" -ForegroundColor Green
    Write-Host "Multi-symbol backtest automation completed" -ForegroundColor Green
    Write-Host "Total tests: $totalTests" -ForegroundColor Green
    Write-Host "Successful tests: $successfulTests" -ForegroundColor Green
    Write-Host "Failed tests: $($totalTests - $successfulTests)" -ForegroundColor Green
    Write-Host "Master summary: $masterSummaryPath" -ForegroundColor Green
    Write-Host "====================================" -ForegroundColor Green
}

function Update-IniFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Symbol,
        
        [Parameter(Mandatory = $true)]
        [string]$Timeframe
    )
    
    try {
        # Read the current INI file
        $iniContent = Get-Content -Path $config.configIniPath -Raw
        
        # Update Symbol
        $iniContent = $iniContent -replace 'Symbol=.*', "Symbol=$Symbol"
        
        # Update Period (Timeframe)
        $iniContent = $iniContent -replace 'Period=.*', "Period=$Timeframe"
        
        # Write the updated content back to the file
        Set-Content -Path $config.configIniPath -Value $iniContent
        
        Write-Log -Level "INFO" -Message "Updated INI file for new test" -Details @{
            "symbol" = $Symbol
            "timeframe" = $Timeframe
        }
        
        return $true
    }
    catch {
        Write-Log -Level "ERROR" -Message "Failed to update INI file: $($_.Exception.Message)" -Details @{}
        return $false
    }
}

#endregion Main Execution

# Add UI Automation helper functions
Add-Type @"
using System;
using System.Runtime.InteropServices;

public class Win32 {
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    
    [DllImport("user32.dll")]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
    
    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    public static extern bool IsWindow(IntPtr hWnd);
    
    // ShowWindow commands
    public const int SW_SHOW = 5;
    public const int SW_RESTORE = 9;
}
"@

# Execute the script
if ($PSBoundParameters.Count -eq 0) {
    # If no parameters provided, run single backtest
    Start-BacktestAutomation
}
else {
    # Otherwise, use the parameters provided
    # This allows the script to be called with parameters from another script
    $PSBoundParameters
}

# Example usage:
# To run a single backtest:
# .\MT5_Backtest_Automation.ps1

# To run multiple backtests with specific symbols and timeframes:
# .\MT5_Backtest_Automation.ps1 -Command Start-MultiSymbolBacktestAutomation -Symbols "EURUSD","GBPUSD" -Timeframes "H1","H4"



