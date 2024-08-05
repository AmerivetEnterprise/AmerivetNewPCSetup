﻿#Version 5.0 
#08/05/2024

$host.UI.RawUI.WindowTitle = "New User Setup Post Restart 5"

Write-Host 'Script will continue in 5 seconds'
Start-Sleep 5

#Enables RunOnce 
$RunOnceKey = "HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce"
set-itemproperty $RunOnceKey "NextRun" ('C:\Windows\System32\WindowsPowerShell\v1.0\Powershell.exe -executionPolicy Unrestricted -File ' + "C:\IT\New_User_Setup_Post_Restart_5.ps1")

###############################################################################################################################################################################################

$FalconCIDFile = "C:\IT\FalconCID.txt"
if (Test-Path -path $FalconCIDFile)
{ 
    #Do Nothing and Continue
}
else
{
Read-host 'Falcon CID txt file missing, place in C:\IT before continuing' -ForegroundColor Red
}

# Specify the path to the input file
$inputFilePath = "C:\IT\FalconCID.txt"

# Read the content from the input file
$CID = Get-Content -Path $inputFilePath

###############################################################################################################################################################################################

#Installs CrowdStrike Falcon Sensor
Write-Host "Downloading and Installing CrowdStrike Falcon Sensor"
Invoke-WebRequest -Uri 'https://amerivetusa.sharepoint.com/:u:/s/IT/Ea7qg4xOmd5Ppy0QgALsHmwBHqNLhgs4FumgOoQ4DRQmGw?download=1' -OutFile "$env:TEMP/WindowsSensor.MaverickGyr.exe"; Start-Process -FilePath "$env:TEMP/WindowsSensor.MaverickGyr.exe" -ArgumentList "/install /quiet /norestart CID=$CID" -Wait
Write-host ''
Write-Host "CrowdStrike Falcon Sensor Installed Successfully" -ForegroundColor Green
Write-host ''

###############################################################################################################################################################################################

$input = Read-Host "Are you done? [y/n]"

switch($input)

    {

y
{

Remove-itemproperty $RunOnceKey "NextRun"

write-warning "Run once has been disabled"
write-host ''
write-warning "Begin Final Check"
write-host ''

#Starts Disk Cleanup

Reg Import "C:\IT\DiskCleanUpSelections.reg"
Cleanmgr /sagerun:0000

Read-Host 'Disk Clean up initiated, when it disappears its complete'
Write-Host ''
Read-Host 'Please verify all checklist items below are complete, press enter to continue'
Write-Host ''
winver
Read-Host 'Is Windows on the latest version? [y/n]'
Get-BitLockerVolume
Write-Host ''
Read-Host 'Is bitlocker enabled and disk encrypted to 100%? [y/n]'
Write-Host ''
Read-Host 'Serial #, Model, OS Key added to spreadsheet [y/n]'
Write-Host ''
Read-Host 'If HQ user, are printers mapped? [y/n]'
Write-Host ''
Read-Host 'Did you login into UniFi app as user and connect to WiFi? [y/n]'
Write-Host ''
Read-Host 'Did you set GAL as default in Outlook? [y/n]'
Write-Host ''
Read-Host 'Has Edge been launched? [y/n]'
Write-Host ''
Read-Host 'Has --NEW-- Outlook been pinned to task bar? [y/n]'
Write-Host ''
Read-Host 'If Acrobat user, has acrobat been launched and set to default? [y/n]'
Write-Host ''
Read-Host 'Are all windows updates complete? [y/n]'
Write-Host ''
Read-Host 'Are all HP support assistant driver updates complete? [y/n]'
write-Host ''
Read-Host 'HP Support Assistant: Uncheck auto software update and uncheck all settings in more settings? [y/n]'
Write-Host ''
Read-Host 'If this is a computer replacement, you will have to manually put their new device in the Security - Autopatch Group [y/n]'
Write-Host ''
Read-Host 'Did you click Discover Devices in Windows Autopatch https://endpoint.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DevicesMenu/~/windowsAutopatchDevices [y/n]'
Write-Host ''
write-warning 'Disbale AutoLogin'
Start-Process -FilePath "C:\IT\autologon64.exe" -wait
Write-Host ''
Read-Host 'Did you disable AutoLogin? [y/n]'
write-host ''

#Enables RunOnce 
Write-Host "Changing RunOnce script for Post Restart CleanUp PS1." -foregroundcolor "magenta"
write-host ''

$RunOnceKey = "HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce"
set-itemproperty $RunOnceKey "NextRun" ('C:\Windows\System32\WindowsPowerShell\v1.0\Powershell.exe -executionPolicy Unrestricted -File ' + "C:\IT\New_User_Setup_Post_Restart_CleanUp.ps1")

$UPN = whoami /upn

$UPN

Write-Host ''

Read-Host "Press Enter to remove from New Scripts Group and Send Password Reset Email to user"

#Starts New Employee Inventory Update Flow
$PA_URL = "https://prod-164.westus.logic.azure.com:443/workflows/7b9c0ef9ccaf49ac9a0914bd563b5fc8/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=oztyW61WYISqt9xomTy2zcJGIHPDtO6Qt7MERCgPO3w"
$Result = @{

UPN = $UPN

}

#Send Results to PA
Invoke-WebRequest -Uri $PA_URL -Method POST -ContentType 'application/json' -Body ($result | ConvertTo-Json -Compress)

###################################################################
Add-Content -Path C:\IT\Complete.txt -value "`NUSPR_5 - Complete"

Write-Host "Restarting in 30 seconds"
Start-Sleep 30
Restart-Computer


}

#n{exit}
n
{
#Enables RunOnce 
write-warning "Auto Login is still enabled, do not forget to disable it."
write-host ''
start-sleep 3
write-host ''
}

default
{
#Enables RunOnce 
write-warning "Auto Login is still enabled, do not forget to disable it."
write-host ''
start-sleep 3
write-host ''
}

    }