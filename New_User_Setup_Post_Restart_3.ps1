<#﻿
Version 5.1
08/06/2024
#>

$host.UI.RawUI.WindowTitle = "New User Setup Post Restart 3"

Write-Host 'Script will continue in 30 seconds'
Start-Sleep 30

start Outlook.exe

Read-Host "Press enter after Outlook has successfully loaded and you have signed into the Amerivet account"

#########################Update Existing Outlook Registry Keys#########################

Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook\0a0d020000000000c000000000000046\" -Name 00030429 -Value ([byte[]](0x04,0x00,0x00,0x00))

Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook\0a0d020000000000c000000000000046\" -Name 00030455 -Value ([byte[]](0xc5,0x53,0xcd,0x11))

Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook\0a0d020000000000c000000000000046\" -Name 01023d15 -Value ([byte[]](0xa3,0x45,0xf1,0xe1,0xfd,0x57,0xfc,0x40,0xa9,0x8c,0x62,0x65,0xae,0xf8,0xeb,0xbd))

#########################Create New Outlook Registry Keys#########################

New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook\0a0d020000000000c000000000000046\" -Name 00030456 -Value ([byte[]](0x70,0x00,0x00,0x00))

New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook\0a0d020000000000c000000000000046\" -Name 00033d1b -Value ([byte[]](0x01,0x00,0x00,0x00))

New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook\0a0d020000000000c000000000000046\" -Name 000b3d1c -Value ([byte[]](0x00,0x00))

#Logs in to OneDrive and Enables Backup

Write-Host "Enabling One Drive Auto Login." -foregroundcolor "magenta"

$HKLMregistryPath = 'HKLM:\SOFTWARE\Policies\Microsoft\OneDrive'##Path to HKLM keys
$DiskSizeregistryPath = 'HKLM:\SOFTWARE\Policies\Microsoft\OneDrive\DiskSpaceCheckThresholdMB'##Path to max disk size key
$TenantGUID = 'c0456220-7d8e-4458-b009-91d710a877d4'

if(!(Test-Path $HKLMregistryPath)){New-Item -Path $HKLMregistryPath -Force}
if(!(Test-Path $DiskSizeregistryPath)){New-Item -Path $DiskSizeregistryPath -Force}

New-ItemProperty -Path $HKLMregistryPath -Name 'SilentAccountConfig' -Value '1' -PropertyType DWORD -Force | Out-Null ##Enable silent account configuration
New-ItemProperty -Path $DiskSizeregistryPath -Name $TenantGUID -Value '102400' -PropertyType DWORD -Force | Out-Null ##Set max OneDrive threshold before prompting
New-ItemProperty -Path $HKLMregistryPath -Name "KFMSilentOptIn" -Value $TenantGUID -PropertyType string -Force | Out-Null
New-ItemProperty -Path $HKLMregistryPath -Name "KFMSilentOptInWithNotification" -Value '1' -PropertyType DWORD -Force | Out-Null

##############################################################################################################################################

#Enables RunOnce 

Write-Host "Changing RunOnce script for Disabling Admin Auto Login." -foregroundcolor "magenta"
write-host ''
$RunOnceKey = "HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce"
set-itemproperty $RunOnceKey "NextRun" ('C:\Windows\System32\WindowsPowerShell\v1.0\Powershell.exe -executionPolicy Unrestricted -File ' + "C:\IT\New_User_Setup_Post_Restart_4.ps1")
write-host ''

##############################################################################################################################################

Add-Content -Path C:\IT\Complete.txt -value "`NUSPR_3 - Complete"

Restart-Computer
