# Version 6.0
# 08/09/2024

$host.UI.RawUI.WindowTitle = "New User Setup Post Restart 2"

#Enables RunOnce 
Write-Host "Changing RunOnce script for Post Restart CleanUp PS1." -foregroundcolor "magenta"
$RunOnceKey = "HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce"
set-itemproperty $RunOnceKey "NextRun" ('C:\Windows\System32\WindowsPowerShell\v1.0\Powershell.exe -executionPolicy Unrestricted -File ' + "C:\IT\New_User_Setup_Post_Restart_CleanUp.ps1")
Add-Content -Path C:\IT\Complete.txt -value "Set RunOnce Key for Cleanup Script"

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

Add-Content -Path C:\IT\Complete.txt -value "Acquired PC Info for $SN"

##############################################################################################################################################

$UPN = whoami /upn

# Split the email into the username and domain parts
$parts = $UPN.Split('@')
$username = $parts[0]
$domain = $parts[1]

# Split the username on the dot, capitalize each part, then rejoin them
$capitalizedUsername = ($username.Split('.') | ForEach-Object { $_.Substring(0,1).ToUpper() + $_.Substring(1).ToLower() }) -join '.'

# Reassemble the full email
$capitalizedUPN = "$capitalizedUsername@$domain"

Add-Content -Path C:\IT\Complete.txt -value "Acquired UPN and capitalized $capitalizedUPN"

#################################################################################################

#Confirms Info.csv file is in place
$InfoFile = "C:\IT\Info.csv"
if (Test-Path -path $InfoFile)
{ 
Write-host "Info.csv File in place" -ForegroundColor Green
Add-Content -Path C:\IT\Complete.txt -value "Info.csv confirmed"
}
else
{
Write-Host "Location: IT Software\IT Software - Internal Only\New PC Setup Software\Info CSV\Info.csv" -ForegroundColor Yellow
Write-Host "Info.csv is missing, place in C:\IT before continuing" -foregroundcolor red
Add-Content -Path C:\IT\Complete.txt -value "Info.csv MISSING"
Read-host "Press Enter to continue once complete"
Add-Content -Path C:\IT\Complete.txt -value "Info.csv manually placed"
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
    $CID = $row.FalconCID
    $PA_URL_Complete = $row.PA_URL_Complete
    }

#########################################################################

#Installs ScreenConnect
Write-Host 'Installing ScreenConnect' -foregroundcolor yellow
Start-Process msiexec.exe -wait -ArgumentList '/I C:\IT\ConnectWiseControl.ClientSetup.Amerivet.msi /quiet /qn'
Write-Host "ScreenConnect Amerivet Installed" -foregroundcolor "green"
Add-Content -Path C:\IT\Complete.txt -value "ScreenConnect Amerivet Installed"

# UID Desktop Software install MSI
Write-Host 'Installing UI' -foregroundcolor yellow
Start-Process msiexec.exe -wait -ArgumentList '/I C:\IT\UI_Desktop.msi'
Write-Host "UI Desktop Installed" -foregroundcolor "green"
Add-Content -Path C:\IT\Complete.txt -value "UID Desktop Software installed"

#Installs Datto for Corp Site
#Start-process "C:\IT\DattoAgentInstaller.exe"
#Add-Content -Path C:\IT\Complete.txt -value "Datto RMM Installed"

#Set Power Settings 
Write-Host "Updating Power Settings" -foregroundcolor yellow
Powercfg /Change -monitor-timeout-ac 60
Powercfg /Change -standby-timeout-ac 0
Powercfg /Change -hibernate-timeout-ac 0
Write-Host "Power Settings Adjusted" -foregroundcolor green
Add-Content -Path C:\IT\Complete.txt -value "Power settings udpated"

#Intalls Office 
Write-Host 'Office Installing' -foregroundcolor yellow
Start-Process -wait C:\IT\setupo365businessretail.x64.en-us_.exe
Write-Host "Office Installed" -foregroundcolor green
Add-Content -Path C:\IT\Complete.txt -value "Office Installed"
Stop-Process -name "OfficeC2RClient" -Confirm:$false

#############################################################################################
# Determines Location On Prem or Remote Setup

# Get the public IP address using an external service
$publicIP = (Invoke-WebRequest -Uri "http://ipinfo.io/ip" -TimeoutSec 10).Content.Trim()
Add-Content -Path C:\IT\Complete.txt -value "IP $publicIP"

# Check conditions: 
$HQ_IPs = "$HQIP1", "$HQIP2"
if ($HQ_IPs -contains $publicIP) {

    Write-Host "On Prem Setup detected" -foregroundcolor "yellow"
    
    #Authenticates to OnPrem NAS for Adobe Download
    Write-Host "Mapping to NAS" -foregroundcolor "yellow"
    $NASCred = New-Object System.Management.Automation.PsCredential($NASUser,(ConvertTo-SecureString $NASPass -AsPlainText -Force))
    New-PSDrive -Name "A" -Root "\\172.16.0.158\AmerivetNewUser" -Persist -PSProvider "FileSystem" -Credential $NAScred
    
    Write-Host "Downloading Adobe from NAS" -foregroundcolor "yellow"
    Write-Host "Downloading HP Support Assistant from NAS" -foregroundcolor "yellow"

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

# File paths for CrowdStrike
$sourceFileHP = "\\172.16.0.158\AmerivetNewUser\NewUserSetup\Software\WindowsSensor.MaverickGyr_7.16.18608.exe"
$destinationFileHP = "C:\IT\WindowsSensor.MaverickGyr.exe"

# Call the function for Adobe Acrobat
Copy-FileWithProgress -sourcePath $sourceFileAdobe -destinationPath $destinationFileAdobe
Add-Content -Path C:\IT\Complete.txt -value "Adobe Downloaded from NAS"

# Call the function for HP Support Assistant
Copy-FileWithProgress -sourcePath $sourceFileHP -destinationPath $destinationFileHP
Add-Content -Path C:\IT\Complete.txt -value "HP Support Assistant Downloaded from NAS"

# Call the function for CrowdStrike Download
Copy-FileWithProgress -sourcePath $sourceFileCS -destinationPath $destinationFileCS
Add-Content -Path C:\IT\Complete.txt -value "CrowdStrike Downloaded from NAS"

##############################################################################################

    } 

    else {

    Write-Host "Remote Setup Detected" -foregroundcolor "yellow"
    Write-Host ''
    $publicIP
    Write-Host ''
    Write-Host "Downloading Adobe from OneDrive - This is slow" -foregroundcolor "yellow"
    #Location: IT Software\IT Software - Internal Only\New PC Setup Software\Adobe Acrobat\AdobeAcrobat.zip
    Invoke-WebRequest -Uri "https://amerivetusa.sharepoint.com/:u:/s/IT/EfO2m45Pf9FKj3gHBRahY00B4nHQP4WVoOIgjzl17pHhCA?download=1" -OutFile "C:\IT\AmerivetAcrobat.zip"
    Add-Content -Path C:\IT\Complete.txt -value "Adobe Downloaded from OneDrive"

    #Location: IT Software\IT Software - Internal Only\New PC Setup Software\HP Support Assistant\HP Support Assistant.exe
    Write-Host "Downloading HP Support Assistant from OneDrive - This is slow also" -foregroundcolor "yellow"
    Invoke-WebRequest -Uri "https://amerivetusa.sharepoint.com/:u:/s/IT/EUayqvTnOIJJlY8idS8mVLwBbMbDpVu-vFJ2cPXZmuE8tw?download=1" -OutFile "C:\IT\HP Support Assistant.exe"
    Add-Content -Path C:\IT\Complete.txt -value "HP Support Assistant Downloaded from OneDrive"

    #Location: IT Software\IT Software - Internal Only\New PC Setup Software\CrowdStrike Falcon Sensor\
    Write-Host "Downloading CrowdStrike Falcon Sensor - This is also slow" -ForegroundColor yellow
    Invoke-WebRequest -Uri 'https://amerivetusa.sharepoint.com/:u:/s/IT/Ea7qg4xOmd5Ppy0QgALsHmwBHqNLhgs4FumgOoQ4DRQmGw?download=1' -OutFile "C:\IT\WindowsSensor.MaverickGyr.exe"
    Add-Content -Path C:\IT\Complete.txt -value "CrowdStrike Falcon Sensor downloaded"
    }

#############################################################################################

Write-Host ''
Write-Host "Adobe Downloaded from OneDrive Successfully" -foregroundcolor green
Write-Host "HP Support Assistant Downloaded from OneDrive Successfully" -foregroundcolor green
Write-Host "CrowdStrike Falcon Sensor Downloaded from OneDrive Successfully" -ForegroundColor green
Write-Host ''

#unZip Adobe
write-host "Unzipping Adobe" -foregroundcolor "yellow"
$unzipAdobe = Expand-Archive -Path "C:\IT\AmerivetAcrobat.zip" -DestinationPath "C:\IT\"
Add-Content -Path C:\IT\Complete.txt -value "Adobe Unzipped"

#Adobe Silent Install
write-host "Installing Adobe" -foregroundcolor "yellow"
$AdobeVersion = "24.1"
Start-Process -FilePath "C:\IT\AmerivetAcrobat\Build\Setup\APRO$AdobeVersion\Adobe Acrobat\setup.exe" -ArgumentList "/sAll", "/msi EULA_ACCEPT=YES /qn" -NoNewWindow -Wait

Write-Host "Adobe Installed" -foregroundcolor "green"
Add-Content -Path C:\IT\Complete.txt -value "Adobe Installed"

###############################################################################################

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

Add-Content -Path C:\IT\Complete.txt -value "OneDrive Logged in and backup enabled"

##############################################################################################################################################

$OS_Key = Get-WmiObject -query 'select * from SoftwareLicensingService' | Select OA3xOriginalProductKey
$OS_Key = $OS_Key.OA3xOriginalProductKey

#The below command would import OS key, no need for this now, but will leave here just in case.
#slmgr.vbs /ipk $OS_Key

Add-Content -Path C:\IT\Complete.txt -value "OS Key verified $OS_Key"

##############################################################################################################

# Download Teams Installer
$TeamsDLPath = "C:\IT\Teams_windows_x64.msix"
Write-host 'Downloading Teams' -ForegroundColor Yellow
Invoke-WebRequest -Uri "https://go.microsoft.com/fwlink/?linkid=2196106&clcid=0x409&culture=en-us&country=us" -OutFile "$TeamsDLPath"
Write-Host "Teams Download completed." -ForegroundColor Green
Add-Content -Path C:\IT\Complete.txt -value "Downloading Teams"

# Install Teams Silently
Add-Type -AssemblyName System.Windows.Forms

try {
# Install Teams Silently
Write-host 'Installing Teams' -ForegroundColor Yellow
Add-AppxPackage -Path $TeamsDLPath
Add-Content -Path C:\IT\Complete.txt -value "Attempting Teams Install"
}
catch {
    # Error handling if opening the store fails
    Write-Host "Teams Appears to already be installed"
    $result = [System.Windows.Forms.MessageBox]::Show("Teams Appears to already be installed", "Teams already Installed", [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Error)
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        Write-Host "Teams Already Installed"
        Add-Content -Path C:\IT\Complete.txt -value "Teams Already Installed"
    } else {
        Write-Host "Operation cancelled."
        Add-Content -Path C:\IT\Complete.txt -value "Teams Already Installed"
    }
}

Write-Host "Teams Installation completed."-ForegroundColor Green
Add-Content -Path C:\IT\Complete.txt -value "Teams Installation completed"

##################################################################################################################################################################

#Installs CrowdStrike Falcon Sensor
Write-Host "Installing CrowdStrike Falcon Sensor" -ForegroundColor yellow
Start-Process -FilePath "C:\IT\WindowsSensor.MaverickGyr.exe" -ArgumentList "/install /quiet /norestart CID=$CID" -Wait
Write-host ''
Write-Host "CrowdStrike Falcon Sensor Installed Successfully" -ForegroundColor Green
Write-host ''
Add-Content -Path C:\IT\Complete.txt -value "CrowdStrike Falcon Sensor installed"

#################################################################################################################################################

start Outlook.exe
Add-Content -Path C:\IT\Complete.txt -value "Outlook Launched"

Read-Host "Press enter after Outlook has successfully loaded and you have signed into the Amerivet account"
Write-Host ""
Write-Host "TURN NEW OUTLOOK ON" -ForegroundColor Red
Write-Host ""
Read-Host "Press Enter after this is complete"
Add-Content -Path C:\IT\Complete.txt -value "New Outlook enabled... Hopefully"

#########################Update Existing Outlook Registry Keys#########################

Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook\0a0d020000000000c000000000000046\" -Name 00030429 -Value ([byte[]](0x04,0x00,0x00,0x00))

Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook\0a0d020000000000c000000000000046\" -Name 00030455 -Value ([byte[]](0xc5,0x53,0xcd,0x11))

Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook\0a0d020000000000c000000000000046\" -Name 01023d15 -Value ([byte[]](0xa3,0x45,0xf1,0xe1,0xfd,0x57,0xfc,0x40,0xa9,0x8c,0x62,0x65,0xae,0xf8,0xeb,0xbd))

Add-Content -Path C:\IT\Complete.txt -value "Update Existing Outlook Registry Keys"

#########################Create New Outlook Registry Keys#########################

New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook\0a0d020000000000c000000000000046\" -Name 00030456 -Value ([byte[]](0x70,0x00,0x00,0x00))

New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook\0a0d020000000000c000000000000046\" -Name 00033d1b -Value ([byte[]](0x01,0x00,0x00,0x00))

New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook\0a0d020000000000c000000000000046\" -Name 000b3d1c -Value ([byte[]](0x00,0x00))

Add-Content -Path C:\IT\Complete.txt -value "Create New Outlook Registry Keys"

#################################################################################################################################################################

write-host ''
Write-Host 'Is this a New User Setup or a replacement computer for an existing User?'
write-host ''
[console]::ForeGroundColor = "Green"
Write-Host '  1. New User Setup'
write-host ''
[console]::ForeGroundColor = "WHITE"
Write-Host '  2. Replacement computer for an existing User'
write-host ''

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

Add-Content -Path C:\IT\Complete.txt -value "New Employee Inventory Update Flow Initaited"
Add-Content -Path C:\IT\Complete.txt -value "Identifier: $Identifier PC"

####################################################################################################################################

# Busylight Software install 
$input = Read-Host "Install Busylight Software (HQ Employee's only) [y/n]"

switch($input)

    {

y
{
    Start-Process msiexec.exe -wait -ArgumentList '/i C:\IT\Busylight4MS_Teams_Setup64.msi CMDLINE="MODIFY=FALSE REMOVE=FALSE SILENT=TRUE ALLUSERS=TRUE" /quiet'

    Write-Host "Busylight4MS_Teams Installed"
    Add-Content -Path C:\IT\Complete.txt -value "Busylight4MS_Teams Installed"
}

#n {exit}
n
{
write-warning "BusyLight Install skipped"
Add-Content -Path C:\IT\Complete.txt -value "BusyLight Install skipped"
}

default
{
write-warning "BusyLight Install skipped"
Add-Content -Path C:\IT\Complete.txt -value "BusyLight Install skipped"
}

    }

##################################################################################################################################################################

#Intalls HP Support Assistant

$CompletionFile = "C:\Program Files (x86)\HP\HP Support Framework\HP Support Assistant.ico"
if (Test-Path -path $CompletionFile)

    {
    write-host 'HP Support Assistant already installed' -ForegroundColor Green
    } 

else
    {
    Start-Process -wait "C:\IT\HP Support Assistant.exe"
    Write-Host "HP Support Assistant Installed"
    }

###############################################################################################################################################################################################

write-warning "Begin Final Check"
write-host ''

#Starts Disk Cleanup

Reg Import "C:\IT\DiskCleanUpSelections.reg"
Cleanmgr /sagerun:0000

Read-Host 'Disk Clean up initiated, when it disappears its complete'
Write-Host ''
Read-Host 'Please verify all checklist items below are complete, press enter to continue'
Write-Host ''
Read-Host 'If HQ user, are printers mapped? [y/n]'
Write-Host ''
Read-Host 'Did you login into UniFi app as user and connect to WiFi? [y/n]'
Write-Host ''
[console]::ForeGroundColor = "Red"
Read-Host 'CONFIRM NEW OUTLOOK IS ENABLED [y/n]'
[console]::ForeGroundColor = "White"
Write-host ''
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

if($Identifier -eq "Replacement")
{
Read-Host 'If this is a computer replacement, you will have to manually put their new device in the Security - Autopatch Group [y/n]'
Write-Host ''
Read-Host 'Did you click Discover Devices in Windows Autopatch https://endpoint.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DevicesMenu/~/windowsAutopatchDevices [y/n]'
Write-Host ''
}
else {}

Read-Host 'Did you disable AutoLogin? [y/n]'
write-host ''
[console]::ForeGroundColor = "Red"
Read-Host 'IS HP WOLF SECURITY CONSOLE REMOVED? [y/n]'
[console]::ForeGroundColor = "White"
Write-host ''

########################################################################################

#Final confrim for WOLF Security and New outlook

# Load necessary assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create a new form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Confirm Action'
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = 'CenterScreen'
$form.BackColor = [System.Drawing.Color]::Red  # Set the background color to red

# Create a label
$label = New-Object System.Windows.Forms.Label
$label.Text = "You must agree to the conditions to proceed."
$label.Location = New-Object System.Drawing.Point(10,10)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.ForeColor = [System.Drawing.Color]::White  # Set text color to white for visibility

# Create a checkbox
$checkBox = New-Object System.Windows.Forms.CheckBox
$checkBox.Location = New-Object System.Drawing.Point(10,40)
$checkBox.Size = New-Object System.Drawing.Size(280,40)
$checkBox.Text = "HP WOLF SECURITY and releated HP Security UNINSTALLED"
$checkBox.ForeColor = [System.Drawing.Color]::White  # Set text color to white for visibility

$checkBox2 = New-Object System.Windows.Forms.CheckBox
$checkBox2.Location = New-Object System.Drawing.Point(10,80)
$checkBox2.Size = New-Object System.Drawing.Size(280,40)
$checkBox2.Text = "New Outlook enabled"
$checkBox2.ForeColor = [System.Drawing.Color]::White  # Set text color to white for visibility

# Create an OK button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(110,135)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = "OK"
$okButton.Enabled = $false

# Event handler for the checkbox HP WOLF
$checkBox.Add_Click({
    $okButton.Enabled = $checkBox.Checked
})


# Event handler for the OK button click
$okButton.Add_Click({
    $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Close()
})

# Add controls to the form
$form.Controls.Add($checkBox)
$form.Controls.Add($checkBox2)
$form.Controls.Add($label)
$form.Controls.Add($okButton)
$form.AcceptButton = $okButton


# Show the form as a dialog box
$result = $form.ShowDialog()

# Cleanup
$form.Dispose()

##################################################################################################################################################

#Disables Old Outlook
Rename-Item -Path 'C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE' -NewName "OUTLOOK.EXE.BAK" 
Write-Host "Renaming C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE to OUTLOOK.EXE.BAK" -foregroundcolor "yellow"
Write-Host "Rename successful" -foregroundcolor "green"
Add-Content -Path C:\IT\Complete.txt -value "Old Outlook DISABLED"

##################################################################################################################################################

Read-Host "Press Enter to remove from New Scripts Group and Send Password Reset Email to user"

# Starts New Hire PC Setup Complete v3 Flow
# https://make.powerautomate.com/environments/Default-c0456220-7d8e-4458-b009-91d710a877d4/flows/shared/4895a808-04ff-479c-a4ce-758996522c1b/details?v3=false
$PA_URL = "$PA_URL_Complete"
$Result = @{

UPN = $capitalizedUPN

}

#Send Results to PA
Invoke-WebRequest -Uri $PA_URL -Method POST -ContentType 'application/json' -Body ($result | ConvertTo-Json -Compress)

##################################################################################################################################################

Write-Host "New Hire PC Setup Complete v3 Flow Initiated" -ForegroundColor Green
Add-Content -Path C:\IT\Complete.txt -value "New Hire PC Setup Complete v3 Flow Initiated"
Add-Content -Path C:\IT\Complete.txt -value "NUSPR_2 - Complete"

Write-Host 'Restarting in 30 seconds' -ForegroundColor Red
Start-Sleep 20
Write-Host 'Restarting in 10 seconds' -ForegroundColor Red
Start-Sleep 1
Write-Host 'Restarting in 9 seconds' -ForegroundColor Red
Start-Sleep 1
Write-Host 'Restarting in 8 seconds' -ForegroundColor Red
Start-Sleep 1
Write-Host 'Restarting in 7 seconds' -ForegroundColor Red
Start-Sleep 1
Write-Host 'Restarting in 6 seconds' -ForegroundColor Red
Start-Sleep 1
Write-Host 'Restarting in 5 seconds' -ForegroundColor Red
Start-Sleep 1
Write-Host 'Restarting in 4 seconds' -ForegroundColor Red
Start-Sleep 1
Write-Host 'Restarting in 3 seconds' -ForegroundColor Red
Start-Sleep 1
Write-Host 'Restarting in 2 seconds' -ForegroundColor Red
Start-Sleep 1
Write-Host 'Restarting in 1 seconds' -ForegroundColor Red
Start-Sleep 1

$formattedTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Add-Content -Path C:\IT\Complete.txt -value "Restarting: $formattedTimestamp"
Start-Sleep 2

Restart-Computer
