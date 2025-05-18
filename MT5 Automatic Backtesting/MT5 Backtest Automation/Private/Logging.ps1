# Logging functions

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
    
    # Check log file size and rotate if needed
    if ((Test-Path -Path $script:config.logFilePath) -and 
        ((Get-Item -Path $script:config.logFilePath).Length / 1MB) -gt $script:config.maxLogSizeMB) {
        
        Rotate-LogFiles
    }
    
    # Only log if verbose mode is on or if it's an important message
    if ($script:config.verboseLogging -or $Level -in @("ERROR", "WARN", "INFO")) {
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
        Add-Content -Path $script:config.logFilePath -Value "$logJson`r`n"
        
        # Also output to appropriate PowerShell streams
        switch ($Level) {
            "ERROR" { 
                Write-Error $Message
                # Also write to host with color for visibility
                Write-Host "[$($logEntry.timestamp)] [$Level] $Message" -ForegroundColor Red
            }
            "WARN"  { 
                Write-Warning $Message
                Write-Host "[$($logEntry.timestamp)] [$Level] $Message" -ForegroundColor Yellow
            }
            "INFO"  { 
                Write-Information $Message -InformationAction Continue
                Write-Host "[$($logEntry.timestamp)] [$Level] $Message" -ForegroundColor White
            }
            "DEBUG" { 
                Write-Debug $Message
                if ($script:config.verboseLogging) {
                    Write-Host "[$($logEntry.timestamp)] [$Level] $Message" -ForegroundColor Gray
                }
            }
        }
    }
}

function Rotate-LogFiles {
    [CmdletBinding()]
    param()
    
    try {
        # Shift existing backups
        for ($i = $script:config.maxLogBackups; $i -gt 0; $i--) {
            $currentBackup = "$($script:config.logFilePath).$i"
            if ($i -eq $script:config.maxLogBackups) {
                # Delete the oldest backup if it exists
                if (Test-Path -Path $currentBackup) {
                    Remove-Item -Path $currentBackup -Force
                }
            }
            else {
                # Shift backup to next number
                $nextBackup = "$($script:config.logFilePath).$($i+1)"
                if (Test-Path -Path $currentBackup) {
                    Move-Item -Path $currentBackup -Destination $nextBackup -Force
                }
            }
        }
        
        # Move current log to .1
        if (Test-Path -Path $script:config.logFilePath) {
            Move-Item -Path $script:config.logFilePath -Destination "$($script:config.logFilePath).1" -Force
        }
        
        Write-Warning "Log file rotated due to size limit ($($script:config.maxLogSizeMB) MB)"
    }
    catch {
        Write-Error "Error rotating log files: $($_.Exception.Message)"
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
        $screenshotPath = Join-Path -Path $script:config.errorScreenshotsPath -ChildPath "error_${ErrorContext}_${timestamp}.png"
        
        # Create bitmap of the screen
        $screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
        $bitmap = New-Object System.Drawing.Bitmap $screen.Width, $screen.Height
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        
        try {
            $graphics.CopyFromScreen($screen.X, $screen.Y, 0, 0, $screen.Size)
            
            # Save the screenshot
            $bitmap.Save($screenshotPath)
            
            $screenshotDetails = @{
                "filename" = "error_${ErrorContext}_${timestamp}.png"
                "context" = $ErrorContext
                "timestamp" = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            }
            
            Write-Information "Error screenshot saved: $screenshotPath"
            Write-Log -Level "INFO" -Message "Error screenshot saved" -Details $screenshotDetails
        }
        finally {
            # Ensure resources are properly disposed
            if ($graphics) { $graphics.Dispose() }
            if ($bitmap) { $bitmap.Dispose() }
        }
        
        # Try to save any partial results if in Strategy Tester
        Save-PartialResults -ErrorContext $ErrorContext -Timestamp $timestamp
    }
    catch {
        Write-Log -Level "ERROR" -Message "Failed to capture error state: $($_.Exception.Message)" -Details @{
            "errorContext" = $ErrorContext
            "errorDetails" = $_.Exception.ToString()
        }
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
        if (Test-WindowExists -WindowTitle $script:constants.WINDOW_STRATEGY_TESTER) {
            # Set focus to the window
            Set-WindowFocus -WindowTitle $script:constants.WINDOW_STRATEGY_TESTER
            
            # Send Ctrl+S to save
            [System.Windows.Forms.SendKeys]::SendWait("^s")
            Invoke-AdaptiveWait -WaitTime $script:constants.WAIT_SHORT
            
            # Set partial results filename
            $partialFileName = "partial_${ErrorContext}_${Timestamp}"
            $fullPath = Join-Path -Path $script:config.reportPath -ChildPath $partialFileName
            
            # Wait for Save As dialog
            Invoke-AdaptiveWait -WaitTime $script:constants.WAIT_SHORT
            
            if (Test-WindowExists -WindowTitle $script:constants.WINDOW_SAVE_AS) {
                try {
                    # Set focus to Save As dialog
                    Set-WindowFocus -WindowTitle $script:constants.WINDOW_SAVE_AS
                    
                    # Type the path
                    [System.Windows.Forms.SendKeys]::SendWait($fullPath)
                    Invoke-AdaptiveWait -WaitTime $script:constants.WAIT_SHORT
                    
                    # Press Enter to save
                    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
                    Invoke-AdaptiveWait -WaitTime $script:constants.WAIT_MEDIUM
                    
                    # Handle potential overwrite confirmation
                    if (Test-WindowExists -WindowTitle $script:constants.WINDOW_CONFIRM) {
                        Set-WindowFocus -WindowTitle $script:constants.WINDOW_CONFIRM
                        [System.Windows.Forms.SendKeys]::SendWait("y")
                        Invoke-AdaptiveWait -WaitTime $script:constants.WAIT_MEDIUM
                    }
                    
                    Write-Information "Partial results saved: $fullPath"
                    Write-Log -Level "INFO" -Message "Partial results saved" -Details @{
                        "filename" = $partialFileName
                        "path" = $fullPath
                        "context" = $ErrorContext
                    }
                }
                catch {
                    Write-Log -Level "ERROR" -Message "Error during save dialog interaction: $($_.Exception.Message)" -Details @{}
                }
                finally {
                    # Ensure dialogs are closed even if there's an error
                    if (Test-WindowExists -WindowTitle $script:constants.WINDOW_SAVE_AS) {
                        Set-WindowFocus -WindowTitle $script:constants.WINDOW_SAVE_AS
                        [System.Windows.Forms.SendKeys]::SendWait("{ESC}")
                    }
                    
                    if (Test-WindowExists -WindowTitle $script:constants.WINDOW_CONFIRM) {
                        Set-WindowFocus -WindowTitle $script:constants.WINDOW_CONFIRM
                        [System.Windows.Forms.SendKeys]::SendWait("{ESC}")
                    }
                }
            }
        }
    }
    catch {
        Write-Log -Level "ERROR" -Message "Failed to save partial results: $($_.Exception.Message)" -Details @{
            "errorContext" = $ErrorContext
            "errorDetails" = $_.Exception.ToString()
        }
    }
}