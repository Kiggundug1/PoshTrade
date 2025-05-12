# Add to initialization
SET notificationEmail TO "your-email@example.com"
SET notificationPhone TO "+1234567890"  # For SMS alerts
SET lastNotificationTime TO 0
SET notificationCooldown TO 3600  # 1 hour between notifications

# Notification function
FUNCTION SendNotification
    PARAMETERS subject, message
    
    # Check cooldown period
    IF %CURRENT TIME IN SECONDS% - %lastNotificationTime% < %notificationCooldown%
        RETURN
    END IF
    
    SET lastNotificationTime TO %CURRENT TIME IN SECONDS%
    
    # Email notification
    TRY
        RUN PROGRAM "powershell.exe -Command \"Send-MailMessage -From 'mt5-ec2@example.com' -To '%notificationEmail%' -Subject '%subject%' -Body '%message%' -SmtpServer 'smtp.example.com'\"" WAIT FOR COMPLETION No
        APPEND TEXT "Sent email notification: %subject%\r\n" TO FILE "%logFilePath%"
    CATCH
        APPEND TEXT "Failed to send email notification: %ERROR MESSAGE%\r\n" TO FILE "%logFilePath%"
    END TRY
    
    # SMS notification (via AWS SNS)
    TRY
        RUN PROGRAM "aws sns publish --phone-number %notificationPhone% --message \"%subject%: %message%\"" WAIT FOR COMPLETION No
        APPEND TEXT "Sent SMS notification\r\n" TO FILE "%logFilePath%"
    CATCH
        APPEND TEXT "Failed to send SMS notification: %ERROR MESSAGE%\r\n" TO FILE "%logFilePath%"
    END TRY
    
    # Discord/Teams webhook (alternative)
    TRY
        RUN PROGRAM "powershell.exe -Command \"Invoke-RestMethod -Uri 'https://discord-webhook-url' -Method Post -Body (@{content='%subject%: %message%'} | ConvertTo-Json) -ContentType 'application/json'\"" WAIT FOR COMPLETION No
    CATCH
        # Ignore failures on secondary notification methods
    END TRY
END FUNCTION