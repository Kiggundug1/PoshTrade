# Configuration functions

function Import-Configuration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [ValidateScript({Test-Path $_ -PathType Leaf})]
        [string]$ConfigFilePath = $script:config.configFilePath
    )
    
    if (-not (Test-Path -Path $ConfigFilePath)) {
        Write-Log -Level "INFO" -Message "No configuration file found at $ConfigFilePath. Using default settings." -Details @{}
        return
    }
    
    try {
        Write-Log -Level "INFO" -Message "Loading configuration from $ConfigFilePath" -Details @{}
        $configData = Get-Content -Path $ConfigFilePath -Raw
        
        # Try to parse as JSON first
        try {
            $importedConfig = $configData | ConvertFrom-Json -AsHashtable
            Apply-JsonConfiguration -ConfigData $importedConfig
            Write-Log -Level "INFO" -Message "Configuration loaded successfully from JSON" -Details @{}
        }
        catch {
            # Fallback to legacy text format parsing
            Write-Log -Level "WARN" -Message "Failed to parse JSON config, falling back to text format" -Details @{
                "error" = $_.Exception.Message
            }
            Parse-LegacyConfiguration -ConfigData $configData
        }
        
        # Validate configuration after loading
        Validate-Configuration
    }
    catch {
        Write-Log -Level "ERROR" -Message "Error loading configuration: $($_.Exception.Message). Using default settings." -Details @{
            "stackTrace" = $_.ScriptStackTrace
        }
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
        if ($script:config.ContainsKey($key)) {
            $script:config[$key] = $ConfigData[$key]
            
            # Log the configuration value if verbose logging is enabled
            if ($script:config.verboseLogging) {
                Write-Verbose "Config: $key = $($ConfigData[$key])"
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
    if ($script:config.ContainsKey($Key)) {
        # Convert value to appropriate type based on existing value
        $existingValue = $script:config[$Key]
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
        $script:config[$Key] = $typedValue
        
        # Log the configuration value if verbose logging is enabled
        if ($script:config.verboseLogging) {
            Write-Verbose "Config: $Key = $typedValue"
            Write-Log -Level "DEBUG" -Message "Config: $Key = $typedValue" -Details @{}
        }
    }
}

function Import-IniSettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [ValidateScript({Test-Path $_ -PathType Leaf})]
        [string]$IniPath = $script:config.configIniPath
    )
    
    try {
        Write-Information "Reading settings from INI file: $IniPath"
        Write-Log -Level "INFO" -Message "Reading settings from INI file: $IniPath" -Details @{}
        $iniContent = Get-Content -Path $IniPath -Raw
        
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
        
        # Validate extracted settings
        Validate-IniSettings -IniSettings $iniSettings
        
        # Log extracted settings
        Write-Log -Level "INFO" -Message "Settings extracted from INI file" -Details $iniSettings
        
        return $iniSettings
    }
    catch {
        Write-Error "Failed to extract settings from INI file: $($_.Exception.Message)"
        Write-Log -Level "ERROR" -Message "Failed to extract settings from INI file: $($_.Exception.Message)" -Details @{
            "stackTrace" = $_.ScriptStackTrace
        }
        return @{}
    }
}

function Validate-IniSettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$IniSettings
    )
    
    # Validate Symbol
    if ($IniSettings.ContainsKey("Symbol") -and $IniSettings["Symbol"]) {
        if ($IniSettings["Symbol"] -notin $script:constants.VALID_SYMBOLS) {
            Write-Warning "Symbol '$($IniSettings["Symbol"])' is not in the list of recognized symbols"
            Write-Log -Level "WARN" -Message "Symbol '$($IniSettings["Symbol"])' is not in the list of recognized symbols" -Details @{
                "validSymbols" = $script:constants.VALID_SYMBOLS -join ", "
            }
        }
    }
    
    # Validate Timeframe
    if ($IniSettings.ContainsKey("Period") -and $IniSettings["Period"]) {
        if ($IniSettings["Period"] -notin $script:constants.VALID_TIMEFRAMES) {
            Write-Warning "Timeframe '$($IniSettings["Period"])' is not in the list of recognized timeframes"
            Write-Log -Level "WARN" -Message "Timeframe '$($IniSettings["Period"])' is not in the list of recognized timeframes" -Details @{
                "validTimeframes" = $script:constants.VALID_TIMEFRAMES -join ", "
            }
        }
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

function Update-IniFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Symbol,
        
        [Parameter(Mandatory = $true)]
        [string]$Timeframe,
        
        [Parameter(Mandatory = $false)]
        [ValidateScript({Test-Path $_ -PathType Leaf})]
        [string]$IniPath = $script:config.configIniPath
    )
    
    try {
        # Read the current INI file
        $iniContent = Get-Content -Path $IniPath -Raw
        
        # Update Symbol
        $iniContent = $iniContent -replace 'Symbol=.*', "Symbol=$Symbol"
        
        # Update Period (Timeframe)
        $iniContent = $iniContent -replace 'Period=.*', "Period=$Timeframe"
        
        # Write the updated content back to the file
        Set-Content -Path $IniPath -Value $iniContent
        
        Write-Information "Updated INI file for new test: $Symbol $Timeframe"
        Write-Log -Level "INFO" -Message "Updated INI file for new test" -Details @{
            "symbol" = $Symbol
            "timeframe" = $Timeframe
        }
        
        return $true
    }
    catch {
        Write-Error "Failed to update INI file: $($_.Exception.Message)"
        Write-Log -Level "ERROR" -Message "Failed to update INI file: $($_.Exception.Message)" -Details @{
            "symbol" = $Symbol
            "timeframe" = $Timeframe
            "iniPath" = $IniPath
            "stackTrace" = $_.ScriptStackTrace
        }
        return $false
    }
}

function Validate-Configuration {
    [CmdletBinding()]
    param()
    
    $configErrors = @()
    
    # Check for required paths
    if (-not (Test-Path -Path $script:config.batchFilePath)) {
        $configErrors += "Batch file does not exist: $($script:config.batchFilePath)"
    }
    
    if (-not (Test-Path -Path $script:config.configIniPath)) {
        $configErrors += "Config INI file does not exist: $($script:config.configIniPath)"
    }
    
    # Check for valid numeric values
    if ($script:config.maxWaitTimeForTest -le 0) {
        $configErrors += "maxWaitTimeForTest must be greater than 0"
    }
    
    if ($script:config.maxRetries -lt 0) {
        $configErrors += "maxRetries must be greater than or equal to 0"
    }
    
    # Log any configuration errors
    if ($configErrors.Count -gt 0) {
        foreach ($error in $configErrors) {
            Write-Error "Configuration error: $error"
            Write-Log -Level "ERROR" -Message "Configuration error: $error" -Details @{}
        }
        
        if (-not $script:config.skipOnError) {
            throw "Configuration validation failed with $($configErrors.Count) errors"
        }
    }
}

function Initialize-Environment {
    [CmdletBinding()]
    param()
    
    Write-Information "Initializing environment"
    Write-Log -Level "INFO" -Message "Initializing environment" -Details @{}
    
    # Validate configuration
    Validate-Configuration
    
    # Validate required paths
    Confirm-RequiredPaths
    
    # Create necessary directories
    if (-not (Test-Path -Path $script:config.reportPath)) {
        New-Item -Path $script:config.reportPath -ItemType Directory -Force | Out-Null
        Write-Information "Created reports directory: $($script:config.reportPath)"
        Write-Log -Level "INFO" -Message "Created reports directory" -Details @{ "path" = $script:config.reportPath }
    }
    
    if (-not (Test-Path -Path $script:config.errorScreenshotsPath)) {
        New-Item -Path $script:config.errorScreenshotsPath -ItemType Directory -Force | Out-Null
        Write-Information "Created error screenshots directory: $($script:config.errorScreenshotsPath)"
        Write-Log -Level "INFO" -Message "Created error screenshots directory" -Details @{ "path" = $script:config.errorScreenshotsPath }
    }
    
    # Initialize performance history
    Initialize-PerformanceHistory
}

function Confirm-RequiredPaths {
    [CmdletBinding()]
    param()
    
    # Check if batch file exists
    if (-not (Test-Path -Path $script:config.batchFilePath)) {
        throw "Batch file does not exist: $($script:config.batchFilePath)"
    }
    
    # Check if config INI file exists
    if (-not (Test-Path -Path $script:config.configIniPath)) {
        throw "Config INI file does not exist: $($script:config.configIniPath)"
    }
}

function Export-Configuration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$ConfigFilePath = $script:config.configFilePath
    )
    
    try {
        # Convert configuration to JSON
        $configJson = $script:config | ConvertTo-Json -Depth 5
        
        # Save to file
        Set-Content -Path $ConfigFilePath -Value $configJson -Force
        
        Write-Information "Configuration saved to $ConfigFilePath"
        Write-Log -Level "INFO" -Message "Configuration saved to file" -Details @{
            "path" = $ConfigFilePath
        }
        
        return $true
    }
    catch {
        Write-Error "Failed to save configuration: $($_.Exception.Message)"
        Write-Log -Level "ERROR" -Message "Failed to save configuration" -Details @{
            "error" = $_.Exception.Message
            "stackTrace" = $_.ScriptStackTrace
        }
        
        return $false
    }
}

function Reset-Configuration {
    [CmdletBinding()]
    param()
    
    # Reset to default configuration
    Initialize-DefaultConfiguration
    
    Write-Information "Configuration reset to defaults"
    Write-Log -Level "INFO" -Message "Configuration reset to defaults" -Details @{}
}

function Initialize-DefaultConfiguration {
    [CmdletBinding()]
    param()
    
    # Initialize default configuration
    $script:config = @{
        # Paths
        batchFilePath = ".\Modified_Run_MT5_Backtest.bat"
        configIniPath = ".\Modified_MT5_Backtest_Config.ini"
        configFilePath = ".\config.json"
        reportPath = ".\Reports"
        logPath = ".\Reports\logs"
        errorScreenshotsPath = ".\Reports\errors"
        checkpointFile = ".\Reports\checkpoint.json"
        performanceHistoryFile = ".\Reports\performance_history.json"
        
        # Test settings
        maxWaitTimeForTest = 180
        initialLoadTime = 15
        maxRetries = 3
        skipOnError = $true
        autoRestartOnFailure = $true
        maxConsecutiveFailures = 5
        
        # Adaptive wait settings
        adaptiveWaitEnabled = $true
        baseWaitMultiplier = 1.0
        maxAdaptiveWaitMultiplier = 5
        
        # System monitoring
        systemLoadCheckInterval = 300
        lowMemoryThreshold = 200
        
        # Logging
        verboseLogging = $true
        logProgressInterval = 10
        
        # EA settings
        eaName = ""
    }
    
    Write-Verbose "Default configuration initialized"
}