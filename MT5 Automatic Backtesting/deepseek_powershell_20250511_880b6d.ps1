# Add these calls to your main testing loop:

# In the timeframe loop, add:
CALL ComprehensiveMonitor
CALL CheckSessionState
CALL KeepSessionAlive

# After saving each report:
CALL NotifyPowerAutomate WITH PARAMETERS "%reportFileName%"
CALL RobustCloudSync WITH PARAMETERS "%reportPath%" "s3://your-bucket/reports/"