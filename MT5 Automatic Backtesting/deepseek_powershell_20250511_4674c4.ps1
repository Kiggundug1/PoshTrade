# Add this more robust sync function
FUNCTION RobustCloudSync
    PARAMETERS syncSource, syncTarget
    
    SET retryCount TO 0
    SET maxSyncRetries TO 5
    SET syncSuccess TO false
    
    WHILE %retryCount% < %maxSyncRetries% AND NOT %syncSuccess%
        TRY
            # First attempt with AWS CLI
            RUN PROGRAM "aws s3 sync %syncSource% %syncTarget% --no-progress --only-show-errors" WAIT FOR COMPLETION Yes
            SET syncSuccess TO true
            APPEND TEXT "Successfully synced %syncSource% to %syncTarget%\r\n" TO FILE "%logFilePath%"
            
            # Verify sync completion
            CREATE FILE "%syncSource%\sync_verify.tmp"
            CALL AdaptiveWait WITH PARAMETERS 5
            RUN PROGRAM "aws s3 cp %syncSource%\sync_verify.tmp %syncTarget%sync_verify.tmp" WAIT FOR COMPLETION Yes
            DELETE FILE "%syncSource%\sync_verify.tmp"
            
        CATCH
            SET retryCount TO %retryCount% + 1
            APPEND TEXT "Sync attempt %retryCount% failed: %ERROR MESSAGE%\r\n" TO FILE "%logFilePath%"
            
            # Fallback to alternative methods
            IF %retryCount% = 2
                TRY
                    # Try with PowerShell instead of AWS CLI
                    RUN PROGRAM "powershell.exe -Command \"Write-S3Object -BucketName your-bucket -Folder %syncSource% -KeyPrefix reports/\"" WAIT FOR COMPLETION Yes
                    SET syncSuccess TO true
                CATCH
                    APPEND TEXT "PowerShell sync attempt failed: %ERROR MESSAGE%\r\n" TO FILE "%logFilePath%"
                END TRY
            ELSIF %retryCount% = 3
                TRY
                    # Compress files before syncing
                    RUN PROGRAM "powershell.exe -Command \"Compress-Archive -Path %syncSource%* -DestinationPath %syncSource%temp.zip -Force\"" WAIT FOR COMPLETION Yes
                    RUN PROGRAM "aws s3 cp %syncSource%temp.zip %syncTarget%temp.zip" WAIT FOR COMPLETION Yes
                    DELETE FILE "%syncSource%temp.zip"
                    SET syncSuccess TO true
                CATCH
                    APPEND TEXT "Compressed sync attempt failed: %ERROR MESSAGE%\r\n" TO FILE "%logFilePath%"
                END TRY
            END IF
            
            IF NOT %syncSuccess%
                CALL AdaptiveWait WITH PARAMETERS 30
            END IF
        END TRY
    END WHILE
    
    IF NOT %syncSuccess%
        APPEND TEXT "Warning: All sync attempts failed for %syncSource%\r\n" TO FILE "%logFilePath%"
        CALL CaptureErrorState WITH PARAMETERS "CloudSyncFailed"
    END IF
END FUNCTION