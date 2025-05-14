# Add to initialization
SET monitoringInterval TO 900  # 15 minutes
SET lastMonitorTime TO 0
SET maxCpuUsage TO 90
SET maxMemoryUsage TO 90
SET maxDiskUsage TO 90

# Enhanced monitoring function
FUNCTION ComprehensiveMonitor
    # Only run at specified intervals
    IF %CURRENT TIME IN SECONDS% - %lastMonitorTime% < %monitoringInterval%
        RETURN
    END IF
    
    SET lastMonitorTime TO %CURRENT TIME IN SECONDS%
    SET systemHealthy TO true
    
    # Check CPU
    TRY
        GET CPU USAGE PERCENTAGE STORE RESULT IN cpuUsage
        IF %cpuUsage% > %maxCpuUsage%
            APPEND TEXT "Warning: High CPU usage detected (%cpuUsage%%)\r\n" TO FILE "%logFilePath%"
            SET systemHealthy TO false
        END IF
    CATCH
        APPEND TEXT "Error monitoring CPU: %ERROR MESSAGE%\r\n" TO FILE "%logFilePath%"
    END TRY
    
    # Check Memory
    TRY
        GET MEMORY USAGE PERCENTAGE STORE RESULT IN memoryUsage
        IF %memoryUsage% > %maxMemoryUsage%
            APPEND TEXT "Warning: High memory usage detected (%memoryUsage%%)\r\n" TO FILE "%logFilePath%"
            SET systemHealthy TO false
            
            # Perform emergency memory cleanup
            RUN PROGRAM "powershell.exe -Command \"Clear-RecycleBin -Force\"" WAIT FOR COMPLETION No
            KILL PROCESS "chrome.exe" WAIT FOR COMPLETION No  # Example of memory-hungry process
        END IF
    CATCH
        APPEND TEXT "Error monitoring memory: %ERROR MESSAGE%\r\n" TO FILE "%logFilePath%"
    END TRY
    
    # Check Disk
    TRY
        GET DISK USAGE PERCENTAGE STORE RESULT IN diskUsage
        IF %diskUsage% > %maxDiskUsage%
            APPEND TEXT "Warning: High disk usage detected (%diskUsage%%)\r\n" TO FILE "%logFilePath%"
            SET systemHealthy TO false
            END IF
        END IF
    CATCH
        APPEND TEXT "Error monitoring disk: %ERROR MESSAGE%\r\n" TO FILE "%logFilePath%"
    END TRY
    
    # Check MT5 process state
    TRY
        IF PROCESS "terminal64.exe" NOT EXISTS
            APPEND TEXT "Critical: MT5 process not running! Attempting restart...\r\n" TO FILE "%logFilePath%"
            RUN PROGRAM "%mt5Path%" WAIT FOR COMPLETION No
            CALL AdaptiveWait WITH PARAMETERS %initialLoadTime%
            SET systemHealthy TO false
        END IF
    CATCH
        APPEND TEXT "Error checking MT5 process: %ERROR MESSAGE%\r\n" TO FILE "%logFilePath%"
    END TRY
    
    # Send notification if system is unhealthy
    IF NOT %systemHealthy%
        CALL SendNotification WITH PARAMETERS "SystemAlert" "Warning: EC2 instance experiencing performance issues"
    END IF
    
    # Log overall status
    APPEND TEXT "System Monitor: CPU=%cpuUsage%%, Memory=%memoryUsage%%, Disk=%diskUsage%%, MT5=%PROCESS "terminal64.exe" EXISTS% \r\n" TO FILE "%logFilePath%"
END FUNCTION