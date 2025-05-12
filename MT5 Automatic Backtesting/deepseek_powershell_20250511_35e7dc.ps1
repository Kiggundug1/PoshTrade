# Add to initialization section
SET rdpCheckInterval TO 600  # Check RDP connection every 10 minutes
SET lastInputTime TO %CURRENT TIME IN SECONDS%
SET maxInactivity TO 1800  # 30 minutes maximum inactivity

# Add this function
FUNCTION CheckSessionState
    # Get last input time (simulated since Power Automate doesn't have direct API)
    TRY
        # This is a simulated check - you may need actual RDP session monitoring tools
        GET LAST INPUT TIME STORE RESULT IN lastInput
        SET lastInputTime TO %lastInput%
        
        IF %CURRENT TIME IN SECONDS% - %lastInputTime% > %maxInactivity%
            APPEND TEXT "Warning: Session inactive for over 30 minutes. Taking preventive measures.\r\n" TO FILE "%logFilePath%"
            
            # Try to simulate activity
            CALL KeepSessionAlive
            
            # If still no activity, consider restarting
            IF %CURRENT TIME IN SECONDS% - %lastInputTime% > %maxInactivity% * 2
                APPEND TEXT "Critical: Session inactive for over 1 hour. Initiating restart.\r\n" TO FILE "%logFilePath%"
                RUN PROGRAM "shutdown /r /t 60 /c \"Automated restart due to session inactivity\"" WAIT FOR COMPLETION No
                EXIT FLOW
            END IF
        END IF
    CATCH
        APPEND TEXT "Error checking session state: %ERROR MESSAGE%\r\n" TO FILE "%logFilePath%"
    END TRY
END FUNCTION