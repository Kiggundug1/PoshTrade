# Add these to your initialization section
SET isEC2Instance TO false  # Set to true when running on EC2
SET remoteDesktopTimeout TO 300  # Timeout for RDP session (5 minutes)

# Add this function
FUNCTION KeepSessionAlive
    # Only needed on EC2 to prevent RDP session timeout
    IF %isEC2Instance%
        # Simulate mouse movement to keep session active
        MOVE MOUSE BY 1 1
        CALL AdaptiveWait WITH PARAMETERS 1
        MOVE MOUSE BY -1 -1
    END IF
END FUNCTION