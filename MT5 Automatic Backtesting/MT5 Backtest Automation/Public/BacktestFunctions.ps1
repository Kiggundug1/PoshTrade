# Public functions for backtest automation

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
        Write-Information "Starting backtest automation with INI-based configuration"
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
                            Write-Information "Backtest completed successfully"
                            Write-Log -Level "INFO" -Message "Backtest completed successfully" -Details @{
                                "ea" = $script:config.eaName
                                "symbol" = $script:runtime.currency
                                "timeframe" = $script:runtime.timeframe
                            }
                        }
                        else {
                            Write-Warning "Backtest completed but report could not be saved"
                            Write-Log -Level "WARN" -Message "Backtest completed but report could not be saved" -Details @{}
                        }
                    }
                    else {
                        Write-Error "Backtest did not complete successfully"
                        Write-Log -Level "ERROR" -Message "Backtest did not complete successfully" -Details @{}
                    }
                }
                else {
                    Write-Error "Failed to start backtest"
                    Write-Log -Level "ERROR" -Message "Failed to start backtest" -Details @{}
                }
            }
            else {
                Write-Error "Strategy Tester did not open"
                Write-Log -Level "ERROR" -Message "Strategy Tester did not open" -Details @{}
            }
            
            # Clean up after backtest
            Invoke-CleanupAfterBacktest
        }
        else {
            Write-Error "Failed to launch MT5"
            Write-Log -Level "ERROR" -Message "Failed to launch MT5" -Details @{}
        }
        
        # Generate summary report
        Generate-SummaryReport
        
        # Log completion
        Write-Information "Backtest automation completed. Generated $($script:config.reportCounter - 1) reports."
        Write-Log -Level "INFO" -Message "Backtest automation completed" -Details @{
            "totalReports" = $script:config.reportCounter - 1
            "completionTime" = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        }
        
        # Display completion message
        Write-Host "Backtest automation completed. Generated $($script:config.reportCounter - 1) reports." -ForegroundColor Green
    }
    catch {
        Write-Error "Fatal error in backtest automation: $($_.Exception.Message)"
        Write-Log -Level "ERROR" -Message "Fatal error in backtest automation: $($_.Exception.Message)" -Details @{
            "stackTrace" = $_.ScriptStackTrace
        }
        
        if (-not $script:config.skipOnError) {
            throw
        }
    }
}

function Start-MultiSymbolBacktestAutomation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [ValidateSet("EURUSD", "GBPUSD", "USDJPY", "AUDUSD", "USDCHF", "EURGBP", "EURJPY", "EURCHF", "AUDJPY", "NZDUSD", "USDCAD", "GBPJPY")]
        [string[]]$Symbols = @("EURUSD", "GBPUSD", "USDJPY", "AUDUSD"),
        
        # Continuing from where we left off

        [Parameter(Mandatory = $false)]
        [ValidateSet("M1", "M5", "M15", "M30", "H1", "H4", "D1", "W1", "MN1")]
        [string[]]$Timeframes = @("M1", "M5", "M15", "M30", "H1", "H4", "D1")
    )

    # Ensure configuration is loaded
    if (-not $script:config -or -not $script:config.reportPath) {
        # Initialize environment and configuration
        Initialize-Environment
        Import-Configuration
        
        # If still null, set a default path
        if (-not $script:config.reportPath) {
            $script:config.reportPath = Join-Path -Path $PSScriptRoot -ChildPath "..\Reports"
            # Ensure the directory exists
            if (-not (Test-Path -Path $script:config.reportPath)) {
                New-Item -Path $script:config.reportPath -ItemType Directory -Force | Out-Null
            }
        }
    }
    
    Write-Information "Starting multi-symbol, multi-timeframe backtest automation"
    Write-Host "Starting multi-symbol, multi-timeframe backtest automation" -ForegroundColor Cyan
    Write-Host "Symbols: $($Symbols -join ', ')" -ForegroundColor Cyan
    Write-Host "Timeframes: $($Timeframes -join ', ')" -ForegroundColor Cyan
    
    # Initialize counters
    $totalTests = $Symbols.Count * $Timeframes.Count
    $completedTests = 0
    $successfulTests = 0
    
    # Create a master summary file
    $masterSummaryPath = Join-Path -Path $script:config.reportPath -ChildPath "master_summary_$(Get-FormattedTimestamp).csv"
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
            $iniUpdated = Update-IniFile -Symbol $symbol -Timeframe $timeframe
            
            if (-not $iniUpdated) {
                Write-Error "Failed to update INI file for $symbol $timeframe. Skipping this combination."
                Write-Log -Level "ERROR" -Message "Failed to update INI file for $symbol $timeframe. Skipping this combination." -Details @{}
                "$symbol,$timeframe,Failed: INI update error,,$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" | 
                    Add-Content -Path $masterSummaryPath
                continue
            }
            
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
                Write-Error "Exception in backtest automation: $($_.Exception.Message)"
                Write-Log -Level "ERROR" -Message "Exception in backtest automation: $($_.Exception.Message)" -Details @{
                    "symbol" = $symbol
                    "timeframe" = $timeframe
                    "stackTrace" = $_.ScriptStackTrace
                }
            }
            finally {
                # Ensure we always try to clean up, even if there's an error
                try {
                    Invoke-CleanupAfterBacktest
                }
                catch {
                    Write-Warning "Error during cleanup: $($_.Exception.Message)"
                }
            }
            
            # Find the most recent report file for this combination
            $reportPattern = "*${symbol}_${timeframe}*.xml"
            $reportFile = Get-ChildItem -Path $script:config.reportPath -Filter $reportPattern | 
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
    
    Write-Information "Multi-symbol backtest automation completed. $successfulTests of $totalTests tests successful."
    Write-Log -Level "INFO" -Message "Multi-symbol backtest automation completed" -Details @{
        "totalTests" = $totalTests
        "successfulTests" = $successfulTests
        "failedTests" = $totalTests - $successfulTests
        "masterSummary" = $masterSummaryPath
    }
}

function Save-BacktestReport {
    [CmdletBinding()]
    param()
    
    try {
        Write-Information "Saving backtest report..."
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
            Invoke-AdaptiveWait -WaitTime $script:constants.WAIT_MEDIUM
            
            return $true
        }
        else {
            throw "Failed to save report using all available methods"
        }
    }
    catch {
        Write-Error "Failed to save Excel report: $($_.Exception.Message)"
        Write-Log -Level "ERROR" -Message "Failed to save Excel report: $($_.Exception.Message)" -Details @{
            "stackTrace" = $_.ScriptStackTrace
        }
        Capture-ErrorState -ErrorContext "SaveReport"
        
        # Increment consecutive failures counter
        $script:runtime.consecutiveFailures++
        
        # Try to close any open dialogs or tabs
        Invoke-CleanupAfterFailedSave
        
        return $false
    }
}

function Generate-SummaryReport {
    [CmdletBinding()]
    param()
    
    try {
        $timestamp = Get-FormattedTimestamp
        $summaryReportPath = Join-Path -Path $script:config.reportPath -ChildPath "backtest_summary_$timestamp.txt"
        
        $summaryContent = @"
=== BACKTEST AUTOMATION SUMMARY ===

Completed at: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")

Total reports generated: $($script:config.reportCounter - 1)

EA tested: $($script:config.eaName)

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
        
        Write-Information "Summary report generated: $summaryReportPath"
        Write-Log -Level "INFO" -Message "Summary report generated" -Details @{
            "path" = $summaryReportPath
        }
    }
    catch {
        Write-Error "Failed to generate summary report: $($_.Exception.Message)"
        Write-Log -Level "ERROR" -Message "Failed to generate summary report: $($_.Exception.Message)" -Details @{
            "stackTrace" = $_.ScriptStackTrace
        }
    }
}
