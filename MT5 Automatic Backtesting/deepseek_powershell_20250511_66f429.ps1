# Run these commands in an elevated PowerShell session on EC2:

# Install AWS CLI
msiexec.exe /i https://awscli.amazonaws.com/AWSCLIV2.msi /quiet

# Configure auto-login (replace password)
$RegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
Set-ItemProperty $RegPath "AutoAdminLogon" -Value "1" -Type String
Set-ItemProperty $RegPath "DefaultUsername" -Value "Administrator" -Type String
Set-ItemProperty $RegPath "DefaultPassword" -Value "your-password" -Type String

# Disable sleep/hibernation
powercfg /h off
powercfg /change standby-timeout-ac 0
powercfg /change monitor-timeout-ac 0