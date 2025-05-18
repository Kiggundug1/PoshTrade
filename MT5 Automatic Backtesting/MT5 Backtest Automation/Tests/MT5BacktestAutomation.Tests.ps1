# Pester tests for MT5 Backtest Automation module

BeforeAll {
    # Import the module
    Import-Module -Name "$PSScriptRoot\..\MT5BacktestAutomation.psd1" -Force
    
    # Mock configuration for testing
    $script:config = @{
        # Core paths
        "batchFilePath" = "TestDrive:\Run_MT5_Backtest.bat"
        "configIniPath" = "TestDrive:\MT5_Backtest_Config.ini"
        "reportPath" = "TestDrive:\Reports"
        "logFilePath" = "TestDrive:\automation_log.json"
        "errorScreenshotsPath" = "TestDrive:\Reports\errors"
        "checkpointFile" = "TestDrive:\Reports\checkpoint.json"
        "configFilePath" = "TestDrive:\Reports\backtest_config.json"
        "performanceHistoryFile" = "TestDrive:\Reports\performance_history.json"
        
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
        "maxLogSizeMB" = 10
        "maxLogBackups" = 3
    }
    
    # Mock runtime variables
    $script:runtime = @{
        "currency" = "EURUSD"
        "timeframe" = "H1"
        "consecutiveFailures" = 0
        "currentAdaptiveMultiplier" = 1.0
        "lastSystemLoadCheck" = 0
        "availableMemory" = 1000
        "cpuUsage" = 50
        "performanceHistory" = @{}
        "testStartTime" = 0
        "previousLoggedProgress" = "0"
        "lastProgressLogTime" = 0
    }
    
    # Mock constants
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
        "SYSTEM_CHECK_INTERVAL" = 60
        
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
        
        # Valid symbols and timeframes
        "VALID_SYMBOLS" = @("EURUSD", "GBPUSD", "USDJPY", "AUDUSD", "USDCHF", "EURGBP", "EURJPY", "EURCHF", "AUDJPY", "NZDUSD", "USDCAD", "GBPJPY")
        "VALID_TIMEFRAMES" = @("M1", "M5", "M15", "M30", "H1", "H4", "D1", "W1", "MN1")
    }
    
    # Continuing from where we left off

    # Create test files
    New-Item -Path "TestDrive:\Run_MT5_Backtest.bat" -ItemType File -Force
    New-Item -Path "TestDrive:\MT5_Backtest_Config.ini" -ItemType File -Force
    New-Item -Path "TestDrive:\Reports" -ItemType Directory -Force
    New-Item -Path "TestDrive:\Reports\errors" -ItemType Directory -Force
    
    # Create sample INI content
    @"
Expert=Moving Average
Symbol=EURUSD
Period=H1
FromDate=2022.01.01
ToDate=2022.12.31
Model=1
Optimization=0
Visual=0
"@ | Set-Content -Path "TestDrive:\MT5_Backtest_Config.ini"
    
    # Mock functions that interact with external systems
    function Test-WindowExists { param($WindowTitle, $MaxAttempts) return $true }
    function Set-WindowFocus { param($WindowTitle, $MaxAttempts) return $true }
    function Wait-ForWindow { param($WindowTitle, $MaxWaitTime) return $true }
    
    # Mock Start-Process to avoid actually starting processes
    Mock Start-Process { 
        return [PSCustomObject]@{
            Id = 12345
            ProcessName = "MockProcess"
        }
    }
    
    # Mock Get-CimInstance for system monitoring
    Mock Get-CimInstance {
        if ($ClassName -eq 'Win32_OperatingSystem') {
            return [PSCustomObject]@{
                FreePhysicalMemory = 4096000  # 4GB in KB
            }
        }
        elseif ($ClassName -eq 'Win32_Processor') {
            return [PSCustomObject]@{
                LoadPercentage = 25
            }
        }
    }
}

Describe "MT5 Backtest Automation Module" {
    Context "Configuration Functions" {
        It "Should validate configuration correctly" {
            # Mock Test-Path to return true for required files
            Mock Test-Path { return $true }
            
            # This should not throw an error
            { Validate-Configuration } | Should -Not -Throw
        }
        
        It "Should detect missing required files" {
            # Mock Test-Path to return false for batch file
            Mock Test-Path { 
                if ($Path -like "*Run_MT5_Backtest.bat") { return $false }
                return $true 
            }
            
            # This should throw an error if skipOnError is false
            $script:config.skipOnError = $false
            { Validate-Configuration } | Should -Throw
            
            # Reset skipOnError
            $script:config.skipOnError = $true
        }
        
        It "Should extract settings from INI file" {
            # Mock Get-Content to return sample INI content
            Mock Get-Content {
                return @"
Expert=Moving Average
Symbol=EURUSD
Period=H1
FromDate=2022.01.01
ToDate=2022.12.31
Model=1
Optimization=0
Visual=0
"@
            }
            
            $iniSettings = Import-IniSettings
            
            $iniSettings["Expert"] | Should -Be "Moving Average"
            $iniSettings["Symbol"] | Should -Be "EURUSD"
            $iniSettings["Period"] | Should -Be "H1"
        }
    }
    
    Context "System Monitoring Functions" {
        It "Should get system resources" {
            Get-SystemResources
            
            $script:runtime.availableMemory | Should -BeGreaterThan 0
            $script:runtime.cpuUsage | Should -BeGreaterOrEqual 0
        }
        
        It "Should update adaptive multiplier based on system load" {
            # Test high load scenario
            $script:runtime.cpuUsage = 95
            $script:runtime.availableMemory = 100
            
            Update-AdaptiveMultiplier
            
            $script:runtime.currentAdaptiveMultiplier | Should -Be $script:config.maxAdaptiveWaitMultiplier
            
            # Test low load scenario
            $script:runtime.cpuUsage = 20
            $script:runtime.availableMemory = 4000
            
            Update-AdaptiveMultiplier
            
            $script:runtime.currentAdaptiveMultiplier | Should -BeLessThan $script:config.maxAdaptiveWaitMultiplier
        }
    }
    
    Context "Utility Functions" {
        It "Should calculate adaptive wait times correctly" {
            $script:runtime.currentAdaptiveMultiplier = 2.0
            $script:config.adaptiveWaitEnabled = $true
            
            # Mock Start-Sleep to avoid actual waiting
            Mock Start-Sleep {}
            
            Invoke-AdaptiveWait -WaitTime 5
            
            # Should have called Start-Sleep with adjusted time
            Should -Invoke Start-Sleep -ParameterFilter { $Seconds -eq 10 }
        }
        
        It "Should use exponential backoff for retries" {
            $script:config.retryBackoffMultiplier = 2.0
            $script:config.maxRetryWaitTime = 30
            
            # Mock Start-Sleep to avoid actual waiting
            Mock Start-Sleep {}
            
            Invoke-AdaptiveWait -WaitTime 5 -IsRetry $true -RetryCount 2
            
            # Should have called Start-Sleep with backoff time (5 * 2^2 = 20)
            Should -Invoke Start-Sleep -ParameterFilter { $Seconds -eq 20 }
        }
    }
}
