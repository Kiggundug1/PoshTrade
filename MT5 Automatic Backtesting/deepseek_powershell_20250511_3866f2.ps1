# Add to initialization
SET cloudSyncPath TO "D:\FOREX\CloudSync\"  # Folder to sync with cloud storage
SET syncInterval TO 3600  # Sync every hour (3600 seconds)
SET lastSyncTime TO 0

# Add this function
FUNCTION SyncToCloudStorage
    # Only run at specified intervals
    IF %CURRENT TIME IN SECONDS% - %lastSyncTime% < %syncInterval%
        RETURN
    END IF
    
    TRY
        # Example for AWS S3 sync (would need AWS CLI installed)
        RUN PROGRAM "aws s3 sync %reportPath% s3://your-bucket-name/reports/" WAIT FOR COMPLETION Yes
        RUN PROGRAM "aws s3 sync %logFilePath% s3://your-bucket-name/logs/" WAIT FOR COMPLETION Yes
        
        # Update last sync time
        SET lastSyncTime TO %CURRENT TIME IN SECONDS%
        APPEND TEXT "Synced reports and logs to cloud storage at %CURRENT TIME%\r\n" TO FILE "%logFilePath%"
    CATCH
        APPEND TEXT "Cloud sync failed: %ERROR MESSAGE%\r\n" TO FILE "%logFilePath%"
    END TRY
END FUNCTION