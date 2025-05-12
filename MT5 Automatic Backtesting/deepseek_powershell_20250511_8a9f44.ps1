# Enhanced Power Automate signaling system
FUNCTION NotifyPowerAutomate
    PARAMETERS fileName
    
    # Create flag file for Power Automate
    TRY
        WRITE TEXT "%fileName%" TO FILE "%powerAutomateFlagFile%"
        APPEND TEXT "Created Power Automate flag for file: %fileName%\r\n" TO FILE "%logFilePath%"
        
        # Optional: HTTP trigger for Power Automate
        RUN PROGRAM "powershell.exe -Command \"Invoke-RestMethod -Uri 'https://your-powerautomate-webhook' -Method Post -Body (@{filename='%fileName%'} | ConvertTo-Json) -ContentType 'application/json'\"" WAIT FOR COMPLETION No
    CATCH
        APPEND TEXT "Failed to notify Power Automate: %ERROR MESSAGE%\r\n" TO FILE "%logFilePath%"
        
        # Fallback to simple file creation
        TRY
            CREATE FILE "%powerAutomateFlagFile%"
        CATCH
            APPEND TEXT "Critical: Could not create Power Automate flag file\r\n" TO FILE "%logFilePath%"
        END TRY
    END TRY
    
    # Clean up old flag files
    TRY
        DELETE FILES "%powerAutomateFlagFile%.old"
        RENAME FILE "%powerAutomateFlagFile%" TO "%powerAutomateFlagFile%.old"
    CATCH
        # Ignore cleanup errors
    END TRY
END FUNCTION