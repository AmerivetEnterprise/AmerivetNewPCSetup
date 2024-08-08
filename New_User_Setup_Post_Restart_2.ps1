<#﻿
Version 5.1
08/08/2024
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
    $NASUser = $row.Username
    $NASPass = $row.Password
    $HQIP1 = $row.HQIP1
    $HQIP2 = $row.HQIP2
    $PA_URL_PCInfo = $row.PA_URL_PCInfo
    }

#############################################################################################

#Starts New Employee Inventory Update Flow via HTTP Request
# https://make.powerautomate.com/environments/Default-c0456220-7d8e-4458-b009-91d710a877d4/flows/shared/d78b6bf4-29f8-4dfd-802a-1550e1ad3d2f/details?v3=false

$Result = @{

UPN = $capitalizedUPN
SN = $SN
Model = $Model
OS_Key = $OS_Key
SKU = $SKU
Identifier = $Identifier
}

#Send Results to PA
Invoke-WebRequest -Uri "$PA_URL_PCInfo" -Method POST -ContentType 'application/json' -Body ($result | ConvertTo-Json -Compress)

#########################################################################

#Installs ScreenConnect
Start-Process msiexec.exe -wait -ArgumentList '/I C:\IT\ConnectWiseControl.ClientSetup.Amerivet.msi /quiet /qn'
Write-Host "ScreenConnect Amerivet Installed"

# UID Desktop Software install MSI  
Start-Process msiexec.exe -wait -ArgumentList '/I C:\IT\UI_Desktop.msi'
Write-Host "UI Desktop Installed"

#Installs Datto for Corp Site
#Start-process "C:\IT\DattoAgentInstaller.exe"

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
    
    #Authenticates to OnPrem NAS for Adobe Download
    Write-Host "Mappting to NAS"
    $NASCred = New-Object System.Management.Automation.PsCredential($NASUser,(ConvertTo-SecureString $NASPass -AsPlainText -Force))
    New-PSDrive -Name "A" -Root "\\172.16.0.158\AmerivetNewUser" -Persist -PSProvider "FileSystem" -Credential $NAScred
    
    Write-Host "Downloading Adobe from NAS"
    Write-Host "Downloading HP Support Assistant from NAS"

    ##############################################################################################

    function Copy-FileWithProgress {
    param(
        [string]$sourcePath,
        [string]$destinationPath
    )

    # Create the destination directory if it doesn't exist
    $destinationDir = Split-Path -Path $destinationPath -Parent
    if (-not (Test-Path -Path $destinationDir)) {
        New-Item -ItemType Directory -Path $destinationDir -Force
    }

    # Get the file size
    $fileSize = (Get-Item $sourcePath).Length
    $fileStreamRead = [System.IO.File]::OpenRead($sourcePath)
    $fileStreamWrite = [System.IO.File]::OpenWrite($destinationPath)

    # Set buffer size (e.g., 1MB)
    $bufferSize = 1MB
    $buffer = New-Object byte[] $bufferSize
    $totalBytesCopied = 0

    try {
        do {
            $readBytes = $fileStreamRead.Read($buffer, 0, $bufferSize)
            $fileStreamWrite.Write($buffer, 0, $readBytes)
            $totalBytesCopied += $readBytes

            # Update progress bar
            $percentage = ($totalBytesCopied / $fileSize) * 100
            Write-Progress -Activity "Copying file..." -Status "$percentage% Complete:" -PercentComplete $percentage
        }
        while ($readBytes -ne 0)
    }
    finally {
        $fileStreamRead.Close()
        $fileStreamWrite.Close()
    }
}

# Specify the source and destination file paths
$sourceFileAdobe = "\\172.16.0.158\AmerivetNewUser\NewUserSetup\Software\AmerivetAcrobat.zip"
$destinationFileAdobe = "C:\IT\AmerivetAcrobat.zip"

# File paths for HP Support Assistant
$sourceFileHP = "\\172.16.0.158\AmerivetNewUser\NewUserSetup\Software\HP Support Assistant.exe"
$destinationFileHP = "C:\IT\HP Support Assistant.exe"

# Call the function for Adobe Acrobat
Copy-FileWithProgress -sourcePath $sourceFileAdobe -destinationPath $destinationFileAdobe

# Call the function for HP Support Assistant
Copy-FileWithProgress -sourcePath $sourceFileHP -destinationPath $destinationFileHP

##############################################################################################

    } 

    else {

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

