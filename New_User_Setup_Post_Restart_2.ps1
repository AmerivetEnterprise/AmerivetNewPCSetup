﻿<#﻿
Version 5.1
08/06/2024
#>

$host.UI.RawUI.WindowTitle = "New User Setup Post Restart 2"

################################################################## Aquires PC Name and Serial Number ############################################################################

$PCName = $env:computername

$SN = get-ciminstance win32_bios | select serialnumber
$SN = $SN.SerialNumber

$SKU = get-computerinfo | select CsSystemSKUNumber
$SKU = $SKU.CsSystemSKUNumber

$Model = get-computerinfo | select CsModel
$Model = $Model.CsModel

$OS_Key = Get-WmiObject -query 'select * from SoftwareLicensingService' | Select OA3xOriginalProductKey
$OS_Key = $OS_Key.OA3xOriginalProductKey

################################################################## AUTOPATCH ############################################################################

$UPN = whoami /upn

# Split the email into the username and domain parts
$parts = $UPN.Split('@')
$username = $parts[0]
$domain = $parts[1]

# Split the username on the dot, capitalize each part, then rejoin them
$capitalizedUsername = ($username.Split('.') | ForEach-Object { $_.Substring(0,1).ToUpper() + $_.Substring(1).ToLower() }) -join '.'

# Reassemble the full email
$capitalizedUPN = "$capitalizedUsername@$domain"

Write-Host ''

Invoke-Item "C:\Program Files\Internet Explorer\iexplore.exe"

Read-Host 'Launching IE, click "Use Recommended Settings" then Press enter to proceed, This is so the HTTP Request can function'

Write-Host ''

###########################################################################################
write-host ''
Write-Host 'Is this a New User Setup or a replacement computer for an existing User?'
write-host ''
[console]::ForeGroundColor = "Green"
Write-Host '  1. New User Setup'
write-host ''
[console]::ForeGroundColor = "WHITE"
Write-Host '  2. Replacement computer for an existing User'
write-host ''
#[console]::ForeGroundColor = "Green"

$Option = Read-Host 'Option'

Switch ($Option) {
    0 {
    
        CLS
        CD C:/
    
    }
    1 {
    
#################################

$Identifier = "New"

#################################
    
    }
    2 {
    
#################################

$Identifier = "Replacement"

#################################

    }
    }

####################################################################################################################################


#Confirms Info.csv file is in place
$InfoFile = "C:\IT\Info.csv"
if (Test-Path -path $InfoFile)
{ 
Write-host "Info.csv File in place" -ForegroundColor Green
}
else
{
Write-Host "Location: IT Software\IT Software - Internal Only\New PC Setup Software\Info CSV\Info.csv" -ForegroundColor Yellow
Read-host 'Info.csv is missing, place in C:\IT before continuing '
}

    #Import Info CSV
    $csvPath = "C:\IT\Info.csv"
    $data = Import-Csv -Path $csvPath

    foreach ($row in $data) {
    $NAS_IP = $row.IP
    $NAS_User = $row.Username
    $NAS_Pw = $row.Password
    $HQIP1 = $row.HQIP1
    $HQIP2 = $row.HQIP2
    $PA_URL_PCInfo  = $row.PA_URL_PCInfo 

#############################################################################################

#Starts New Employee Inventory Update Flow via HTTP Request
# https://make.powerautomate.com/environments/Default-c0456220-7d8e-4458-b009-91d710a877d4/flows/shared/d78b6bf4-29f8-4dfd-802a-1550e1ad3d2f/details?v3=false
$PA_URL = "$PA_URL_PCInfo"

$Result = @{

UPN = $capitalizedUPN
SN = $SN
Model = $Model
OS_Key = $OS_Key
SKU = $SKU
Identifier = $Identifier
}

#Send Results to PA
Invoke-WebRequest -Uri $PA_URL -Method POST -ContentType 'application/json' -Body ($result | ConvertTo-Json -Compress)

#########################################################################

#Installs ScreenConnect
Start-Process msiexec.exe -wait -ArgumentList '/I C:\IT\ConnectWiseControl.ClientSetup.Amerivet.msi /quiet /qn'
Write-Host "ScreenConnect Amerivet Installed"

# UID Desktop Software install MSI  
Start-Process msiexec.exe -wait -ArgumentList '/I C:\IT\UI_Desktop.msi'
Write-Host "UI Desktop Installed"

#Set Power Settings 
Powercfg /Change -monitor-timeout-ac 60
Powercfg /Change -standby-timeout-ac 0
Powercfg /Change -hibernate-timeout-ac 0
Write-Host "Power Settings Adjusted"

#Intalls Office 
Start-Process -wait C:\IT\setupo365businessretail.x64.en-us_.exe
Write-Host "Office Installed"
Stop-Process -name "OfficeC2RClient" -Confirm:$false

#############################################################################################
# Determines Location On Prem or Remote Setup

# Get the public IP address using an external service
$publicIP = (Invoke-WebRequest -Uri "http://ipinfo.io/ip" -TimeoutSec 10).Content.Trim()

# Check conditions: 
$HQ_IPs = "$HQIP1", "$HQIP2"
if ($HQ_IPs -contains $publicIP) {
    Write-Host "On Prem Setup detected"
    }    

    #Authenticates to OnPrem NAS for Adobe Download
    Write-Host "Mappting to NAS"
    $NASUser = '$NAS_User'
    $NASPass = '$NASPass'
    $NASCred = New-Object System.Management.Automation.PsCredential($NASUser,(ConvertTo-SecureString $NASPass -AsPlainText -Force))

    New-PSDrive -Name "A" -Root "\\$NAS_IP\AmerivetNewUser" -Persist -PSProvider "FileSystem" -Credential $NAScred

    #############################################################################################

    #Progress Tracker for Large Download
    $waitTimeMilliseconds = 9 * 60 * 1000 #

    # Script to execute in the new PowerShell instance
    $scriptBlock = {
    C:\IT\FileSizeTracker.ps1
    } 

    # Start a new instance of Windows PowerShell and run the script
    $powershellPath = "$env:windir\system32\windowspowershell\v1.0\powershell.exe"
    $process = Start-Process $powershellPath -NoNewWindow -ArgumentList ("-ExecutionPolicy Bypass -noninteractive -noprofile " + $scriptBlock) -PassThru
    #$process.WaitForExit($waitTimeMilliseconds)

    #############################################################################################

    Write-Host "Downloading Adobe from NAS"
    #Adobe Download 
    Copy-Item "\\$NAS_IP\AmerivetNewUser\NewUserSetup\Software\AmerivetAcrobat.zip -Destination C:\IT\AmerivetAcrobat.zip"

    Write-Host "Downloading HP Support Assistant from NAS"
    #HP Support Assistant Download 
    Copy-Item "\\$NAS_IP\AmerivetNewUser\NewUserSetup\Software\HP Support Assistant.exe" -Destination "C:\IT\HP Support Assistant.exe"

} else {
    Write-Host "Remote Setup Detected"
    $publicIP
    Write-Host "Downloading Adobe from OneDrive - This is slow"
    #Location: IT Software\IT Software - Internal Only\New PC Setup Software\Adobe Acrobat\AdobeAcrobat.zip
    Invoke-WebRequest -Uri "https://amerivetusa.sharepoint.com/:u:/s/IT/EfO2m45Pf9FKj3gHBRahY00B4nHQP4WVoOIgjzl17pHhCA?download=1" -OutFile "C:\IT\AmerivetAcrobat.zip"
    
    #Location: IT Software\IT Software - Internal Only\New PC Setup Software\HP Support Assistant\HP Support Assistant.exe
    Write-Host "Downloading HP Support Assistant from OneDrive - This is slow also"
    Invoke-WebRequest -Uri "https://amerivetusa.sharepoint.com/:u:/s/IT/EUayqvTnOIJJlY8idS8mVLwBbMbDpVu-vFJ2cPXZmuE8tw?download=1" -OutFile "C:\IT\HP Support Assistant.exe"

}

#############################################################################################

#unZip Adobe
$unzipAdobe = Expand-Archive -Path "C:\IT\AmerivetAcrobat.zip" -DestinationPath "C:\IT\"

#Adobe Silent Install
$AdobeVersion = "24.1"
Start-Process -FilePath "C:\IT\AmerivetAcrobat\Build\Setup\APRO$AdobeVersion\Adobe Acrobat\setup.exe" -ArgumentList "/sAll", "/msi EULA_ACCEPT=YES /qn" -NoNewWindow -Wait

Write-Host "Adobe Installed"

###############################################################################################

#Enables RunOnce 

Write-Host "Changing RunOnce script for Post Restart 3." -foregroundcolor "magenta"
write-host ''
$RunOnceKey = "HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce"
set-itemproperty $RunOnceKey "NextRun" ('C:\Windows\System32\WindowsPowerShell\v1.0\Powershell.exe -executionPolicy Unrestricted -File ' + "C:\IT\New_User_Setup_Post_Restart_3.ps1")
write-host ''

##############################################################################################################################################

Add-Content -Path C:\IT\Complete.txt -value "`NUSPR_2 - Complete"

Restart-Computer
