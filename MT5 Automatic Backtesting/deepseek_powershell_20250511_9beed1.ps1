# Add to initialization
SET maxRuntime TO 604800  # Maximum runtime in seconds (7 days)
SET startTime TO %CURRENT TIME IN SECONDS%

# Add this check in your main loop
IF %CURRENT TIME IN SECONDS% - %startTime% > %maxRuntime%
    APPEND TEXT "Maximum runtime reached. Exiting gracefully.\r\n" TO FILE "%logFilePath%"
    EXIT FLOW
END IF

# Call this periodically in your main loop
CALL KeepSessionAlive
CALL SyncToCloudStorage