# Version 6.0
# 08/09/2024

#New User Script #1 Run this after restart - Computer Setup Automation

$host.UI.RawUI.WindowTitle = "New User Setup Post Restart 1"

Write-Host 'Script will continue in 30 seconds'
Start-Sleep 30

#############################################################################################

#Start-Process -FilePath "C:\IT\autologon64.exe" -wait

#############################################################################################

Write-Host 'Renaming Computer'

$NewPCName = $env:UserName

#count characters 
$charCount = ($NewPCName.ToCharArray() | Measure-Object).Count

#if more then 15 then limit

if($charCount -le 15) 
{
rename-computer -newname "$NewPCName"
} 

else 
{
$NewPCName = $NewPCName.substring(0,15)
$NewPCName = $NewPCName -replace '[_]',''
rename-computer -newname "$NewPCName"
} 

Write-Host "Computer Renamed: $NewPCName"
Add-Content -Path C:\IT\Complete.txt -value "PC renamed to : $NewPCName"

#############################################################################################

#Enables RunOnce 

Write-Host "Changing RunOnce script for Post Restart 2." -foregroundcolor "magenta"
write-host ''
$RunOnceKey = "HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce"
set-itemproperty $RunOnceKey "NextRun" ('C:\Windows\System32\WindowsPowerShell\v1.0\Powershell.exe -executionPolicy Unrestricted -File ' + "C:\IT\New_User_Setup_Post_Restart_2.ps1")
write-host ''
write-host 'RunOnce Enabled'
Add-Content -Path C:\IT\Complete.txt -value "RunOnce Enabled"

#############################################################################################
Write-Host "Launching IE, click "Use Recommended Settings" then Press enter to proceed, This is so the HTTP Request can function"
# Open Microsoft Edge
#Start-Process "msedge.exe"
Start-Process "iexplore.exe"
Add-Content -Path C:\IT\Complete.txt -value "Edge Launched"

# Wait 10 seconds with countdown display
for ($i = 30; $i -ge 0; $i--) {
    Write-Host "$i seconds remaining..." -ForegroundColor Yellow
    Start-Sleep -Seconds 1
}

# Close Microsoft Edge
Get-Process "msedge" | Stop-Process
Add-Content -Path C:\IT\Complete.txt -value "Edge Terminated"

#############################################################################################

Add-Type -AssemblyName System.Windows.Forms

try {
    # Attempt to open the Microsoft Store
    Start-Process "ms-windows-store://updates"
    Write-Host "Opening Microsoft Store..."
    Add-Content -Path C:\IT\Complete.txt -value "MSFT Store Update Page Opened auto"
}
catch {
    # Error handling if opening the store fails
    Write-Host "Failed to open Microsoft Store automatically."
    $result = [System.Windows.Forms.MessageBox]::Show("Unable to open the Microsoft Store automatically. Please open the Microsoft Store manually from the Start menu and start updates before proceeding.", "Error Opening Store", [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Error)
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        Write-Host "MSFT Store Opened Manually."
        Add-Content -Path C:\IT\Complete.txt -value "MSFT Store Update Page Opened manually"
    } else {
        Write-Host "Operation cancelled."
        Add-Content -Path C:\IT\Complete.txt -value "MSFT Store Operation cancelled"
    }
}

#############################################################################################

# Ensure the Speech assembly is loaded
Add-Type -AssemblyName System.Speech

# Create a new SpeechSynthesizer object
$synthesizer = New-Object System.Speech.Synthesis.SpeechSynthesizer

#############################################################################################

#Removes Bloatware 

Write-Host ''
write-host "Uninstalling Bloatware"
write-host ''

Get-AppxPackage Microsoft.XboxOneSmartGlass | Remove-AppxPackage
write-host 'Uninstalled XboxOneSmartGlass'
Get-AppxPackage Microsoft.XboxSpeechToTextOverlay | Remove-AppxPackage
write-host 'Uninstalled XboxSpeechToTextOverlay'
Get-AppxPackage Microsoft.Xbox.TCUI | Remove-AppxPackage
write-host 'Uninstalled Microsoft.Xbox.TCUI'
Get-AppxPackage Microsoft.XboxGameOverlay | Remove-AppxPackage
write-host 'Uninstalled XboxGameOverlay'
Get-AppxPackage Microsoft.Office.OneNote | Remove-AppxPackage
write-host 'Uninstalled Office.OneNote'
Get-AppxPackage Microsoft.SkypeApp | Remove-AppxPackage
write-host 'Uninstalled SkypeApp'
Get-AppxPackage Microsoft.GetHelp | Remove-AppxPackage
write-host 'Uninstalled GetHelp'
Get-AppxPackage Microsoft.WindowsFeedbackHub | Remove-AppxPackage
write-host 'Uninstalled WindowsFeedbackHub'
Get-AppxPackage Microsoft.ZuneMusic | Remove-AppxPackage
write-host 'Uninstalled ZuneMusic'
Get-AppxPackage Microsoft.ZuneVideo | Remove-AppxPackage
write-host 'Uninstalled ZuneVideo'
Get-AppxPackage Microsoft.XboxApp | Remove-AppxPackage
write-host 'Uninstalled XboxApp'
Get-AppxPackage Microsoft.BingWeather | Remove-AppxPackage
write-host 'Uninstalled BingWeather'
Get-AppxPackage Microsoft.XboxIdentityProvider | Remove-AppxPackage
write-host 'Uninstalled XboxIdentityProvider'
Get-AppxPackage Microsoft.MixedReality.Portal | Remove-AppxPackage
write-host 'Uninstalled MixedReality.Portal'
Get-AppxPackage Microsoft.MicrosoftOfficeHub | Remove-AppxPackage
write-host 'Uninstalled MicrosoftOfficeHub'
Get-AppxPackage Microsoft.XboxGamingOverlay | Remove-AppxPackage
write-host 'Uninstalled XboxGamingOverlay'
Get-AppxPackage Microsoft.Getstarted | Remove-AppxPackage
write-host 'Uninstalled Getstarted'
Get-AppxPackage Microsoft.MicrosoftSolitaireCollection | Remove-AppxPackage
write-host 'Uninstalled MicrosoftSolitaireCollection'
Get-AppxPackage Microsoft.YourPhone | Remove-AppxPackage
write-host 'Uninstalled YourPhone'
Get-AppxPackage 7EE7776C.LinkedInforWindows | Remove-AppxPackage
write-host 'Uninstalled LinkedInforWindows'
Get-AppxPackage microsoft.windowscommunicationsapps | Remove-AppxPackage
write-host 'Uninstalled windowscommunicationsapps'
Get-AppxPackage -allusers Microsoft.549981C3F5F10* | Remove-AppxPackage
write-host 'Uninstalled Microsoft.549981C3F5F10*'
Get-AppxPackage -allusers Microsoft.BingWeather* | Remove-AppxPackage
write-host 'Uninstalled Microsoft.BingWeather*'
Get-AppxPackage -allusers Microsoft.GetHelp* | Remove-AppxPackage
write-host 'Uninstalled Microsoft.GetHelp*'
Get-AppxPackage -allusers Microsoft.Getstarted* | Remove-AppxPackage
write-host 'Uninstalled Microsoft.Getstarted*'
Get-AppxPackage -allusers Microsoft.HEIFImageExtension* | Remove-AppxPackage
write-host 'Uninstalled Microsoft.HEIFImageExtension*'
Get-AppxPackage -allusers Microsoft.Microsoft3DViewer* | Remove-AppxPackage
write-host 'Uninstalled Microsoft.Microsoft3DViewer*'
Get-AppxPackage -allusers Microsoft.MicrosoftOfficeHub* | Remove-AppxPackage
write-host 'Uninstalled Microsoft.MicrosoftOfficeHub*'
Get-AppxPackage -allusers Microsoft.MicrosoftSolitaireCollection* | Remove-AppxPackage
write-host 'Uninstalled Microsoft.MicrosoftSolitaireCollection*'
Get-AppxPackage -allusers Microsoft.MixedReality.Portal* | Remove-AppxPackage
write-host 'Uninstalled Microsoft.MixedReality.Portal*'
Get-AppxPackage -allusers Microsoft.Office.OneNote* | Remove-AppxPackage
write-host 'Uninstalled Microsoft.Office.OneNote*'
Get-AppxPackage -allusers Microsoft.People* | Remove-AppxPackage
write-host 'Uninstalled Microsoft.People*'
Get-AppxPackage -allusers Microsoft.Print3D* | Remove-AppxPackage
write-host 'Uninstalled Microsoft.Print3D*'
Get-AppxPackage -allusers Microsoft.ScreenSketch* | Remove-AppxPackage
write-host 'Uninstalled Microsoft.ScreenSketch*'
Get-AppxPackage -allusers Microsoft.SkypeApp* | Remove-AppxPackage
write-host 'Uninstalled Microsoft.SkypeApp*'
Get-AppxPackage -allusers Microsoft.StorePurchaseApp* | Remove-AppxPackage
write-host 'Uninstalled Microsoft.StorePurchaseApp*'
Get-AppxPackage -allusers Microsoft.Wallet* | Remove-AppxPackage
write-host 'Uninstalled Microsoft.Wallet*'
Get-AppxPackage -allusers microsoft.windowscommunicationsapps* | Remove-AppxPackage
write-host 'Uninstalled microsoft.windowscommunicationsapps*'
Get-AppxPackage -allusers Microsoft.WindowsFeedbackHub* | Remove-AppxPackage
write-host 'Uninstalled Microsoft.WindowsFeedbackHub*'
Get-AppxPackage -allusers Microsoft.WindowsMaps* | Remove-AppxPackage
write-host 'Uninstalled Microsoft.WindowsMaps*'
Get-AppxPackage -allusers Microsoft.Xbox.TCUI* | Remove-AppxPackage
write-host 'Uninstalled Microsoft.Xbox.TCUI*'
Get-AppxPackage -allusers Microsoft.XboxApp* | Remove-AppxPackage
write-host 'Uninstalled Microsoft.XboxApp*'
Get-AppxPackage -allusers Microsoft.XboxGameOverlay* | Remove-AppxPackage
write-host 'Uninstalled Microsoft.XboxGameOverlay*'
Get-AppxPackage -allusers Microsoft.XboxGamingOverlay* | Remove-AppxPackage
write-host 'Uninstalled Microsoft.XboxGamingOverlay*'
Get-AppxPackage -allusers Microsoft.XboxIdentityProvider* | Remove-AppxPackage
write-host 'Uninstalled Microsoft.XboxIdentityProvider*'
Get-AppxPackage -allusers Microsoft.XboxSpeechToTextOverlay* | Remove-AppxPackage
write-host 'Uninstalled Microsoft.XboxSpeechToTextOverlay*'
Get-AppxPackage -allusers Microsoft.YourPhone* | Remove-AppxPackage
write-host 'Uninstalled Microsoft.YourPhone*'
Get-AppxPackage -allusers Microsoft.ZuneMusic* | Remove-AppxPackage
write-host 'Uninstalled Microsoft.ZuneMusic*'
Get-AppxPackage -allusers Microsoft.ZuneVideo* | Remove-AppxPackage
write-host 'Uninstalled Microsoft.ZuneVideo*'
Get-AppxPackage -AllUsers *ActiproSoftwareLLC* | Remove-AppxPackage
write-host 'Uninstalled *ActiproSoftwareLLC*'
Get-AppxPackage -AllUsers *AdobeSystemsIncorporated.AdobePhotoshopExpress* | Remove-AppxPackage
write-host 'Uninstalled *AdobeSystemsIncorporated.AdobePhotoshopExpress*'
Get-AppxPackage -AllUsers Microsoft.BingNews* | Remove-AppxPackage
write-host 'Uninstalled Microsoft.BingNews*'
Get-AppxPackage -AllUsers *CandyCrush* | Remove-AppxPackage
write-host 'Uninstalled *CandyCrush*'
Get-AppxPackage -AllUsers *Duolingo* | Remove-AppxPackage
write-host 'Uninstalled *Duolingo*'
Get-AppxPackage -AllUsers *EclipseManager* | Remove-AppxPackage
write-host 'Uninstalled *EclipseManager*'
Get-AppxPackage -AllUsers *Facebook* | Remove-AppxPackage
write-host 'Uninstalled *Facebook*'
Get-AppxPackage -AllUsers *king.com.FarmHeroesSaga* | Remove-AppxPackage
write-host 'Uninstalled *king.com.FarmHeroesSaga*'
Get-AppxPackage -AllUsers *Flipboard* | Remove-AppxPackage
write-host 'Uninstalled *Flipboard*'
Get-AppxPackage -AllUsers *HiddenCityMysteryofShadows* | Remove-AppxPackage
write-host 'Uninstalled *HiddenCityMysteryofShadows*'
Get-AppxPackage -AllUsers *HuluLLC.HuluPlus* | Remove-AppxPackage
write-host 'Uninstalled *HuluLLC.HuluPlus*'
Get-AppxPackage -AllUsers *Pandora* | Remove-AppxPackage
write-host 'Uninstalled *Pandora*'
Get-AppxPackage -AllUsers *Plex* | Remove-AppxPackage
write-host 'Uninstalled *Plex*'
Get-AppxPackage -AllUsers *ROBLOXCORPORATION.ROBLOX* | Remove-AppxPackage
write-host 'Uninstalled *ROBLOXCORPORATION.ROBLOX*'
Get-AppxPackage -AllUsers *Spotify* | Remove-AppxPackage
write-host 'Uninstalled *Spotify*'
Get-AppxPackage -AllUsers *Netflix* | Remove-AppxPackage
write-host 'Uninstalled *Netflix*'
Get-AppxPackage -AllUsers *Microsoft.SkypeApp* | Remove-AppxPackage
write-host 'Uninstalled *Microsoft.SkypeApp*'
Get-AppxPackage -AllUsers *Twitter* | Remove-AppxPackage
write-host 'Uninstalled *Twitter*'
Get-AppxPackage -AllUsers *Wunderlist* | Remove-AppxPackage
write-host 'Uninstalled *Wunderlist*'

Add-Content -Path C:\IT\Complete.txt -value "Bloatware uninstalled"

############################################################################################

# List of programs to uninstall
$UninstallPrograms = @(
    "HP Sure Click"
    "HP Sure Run"
    "HP Sure Recover"
    "HP Sure Sense"
    "HP Sure Sense Installer"
)

$InstalledPrograms = Get-Package | Where {$UninstallPrograms -contains $_.Name}

# Remove installed programs
$InstalledPrograms | ForEach {

    Write-Host -Object "Attempting to uninstall: [$($_.Name)]..."

    Try {
        $Null = $_ | Uninstall-Package -AllVersions -Force -ErrorAction Stop
        Write-Host -Object "Successfully uninstalled: [$($_.Name)]"
    }
    Catch {Write-Warning -Message "Failed to uninstall: [$($_.Name)]"}
}
############################################################################################

# Remove HP bloatware/crapware
# -- ref: https://community.spiceworks.com/topic/2296941-powershell-script-to-remove-windowsapps-folder?page=1#entry-9032247
# -- note: this script could use improvements. contributions welcome!
# -- todo: Wolf Security improvements ref: https://www.reddit.com/r/SCCM/comments/nru942/hp_wolf_security_how_to_remove_it/

# List of built-in apps to remove
$UninstallPackages = @(
    "AD2F1837.HPJumpStarts"
    "AD2F1837.HPPCHardwareDiagnosticsWindows"
    "AD2F1837.HPPowerManager"
    "AD2F1837.HPPrivacySettings"
    "AD2F1837.HPSureShieldAI"
    "AD2F1837.HPSystemInformation"
    "AD2F1837.HPQuickDrop"
    "AD2F1837.HPWorkWell"
    "AD2F1837.myHP"
    "AD2F1837.HPDesktopSupportUtilities"
)

# List of programs to uninstall
$UninstallPrograms = @(
    "HP Connection Optimizer"
    "HP Documentation"
    "HP MAC Address Manager"
    "HP Notifications"
    "HP Security Update Service"
    "HP System Default Settings"
    "HP Dock Audio"
    "HP Dock Accessory WMI Provider"
    "HP Collaboration Keyboard for Cisco UCC"
    "HP Collaboration Keyboard for Skype for Business"
    "HP Sure Click"
    "HP Sure Run"
    "HP Sure Run Module"
    "HP Sure Recover"
    "HP Sure Sense"
    "HP Sure Sense Installer"
    "HP Wolf Security"
    "HP Wolf Security - Console"
    "HP Wolf Security Application Support for Sure Sense"
    "HP Wolf Security Application Support for Windows"
    "HP Client Security Manager"
    "HP Security Update Service"
)

$HPidentifier = "AD2F1837"

$InstalledPackages = Get-AppxPackage -AllUsers | Where {($UninstallPackages -contains $_.Name)} #-or ($_.Name -match "^$HPidentifier")}

$ProvisionedPackages = Get-AppxProvisionedPackage -Online | Where {($UninstallPackages -contains $_.DisplayName)} #-or ($_.DisplayName -match "^$HPidentifier")}

$InstalledPrograms = Get-Package | Where {$UninstallPrograms -contains $_.Name}

# Remove provisioned packages first
ForEach ($ProvPackage in $ProvisionedPackages) {

    Write-Host -Object "Attempting to remove provisioned package: [$($ProvPackage.DisplayName)]..."

    Try {
        $Null = Remove-AppxProvisionedPackage -PackageName $ProvPackage.PackageName -Online -ErrorAction Stop
        Write-Host -Object "Successfully removed provisioned package: [$($ProvPackage.DisplayName)]"
    }
    Catch {Write-Warning -Message "Failed to remove provisioned package: [$($ProvPackage.DisplayName)]"}
}

# Remove appx packages
ForEach ($AppxPackage in $InstalledPackages) {
                                            
    Write-Host -Object "Attempting to remove Appx package: [$($AppxPackage.Name)]..."

    Try {
        $Null = Remove-AppxPackage -Package $AppxPackage.PackageFullName -AllUsers -ErrorAction Stop
        Write-Host -Object "Successfully removed Appx package: [$($AppxPackage.Name)]"
    }
    Catch {Write-Warning -Message "Failed to remove Appx package: [$($AppxPackage.Name)]"}
}

# Remove installed programs
$InstalledPrograms | ForEach {

    Write-Host -Object "Attempting to uninstall: [$($_.Name)]..."

    Try {
        $Null = $_ | Uninstall-Package -AllVersions -Force -ErrorAction Stop
        Write-Host -Object "Successfully uninstalled: [$($_.Name)]"
    }
    Catch {Write-Warning -Message "Failed to uninstall: [$($_.Name)]"}
}

Add-Content -Path C:\IT\Complete.txt -value "HP Software removed hopefully"

############################################################################################

#Remove HP Connection Optmizer 
Start-Process -FilePath "C:\Program Files (x86)\InstallShield Installation Information\{6468C4A5-E47E-405F-B675-A70A70983EA6}\setup.exe" -ArgumentList "-runfromtemp", "-l0x0409", "-removeonly", "/passive" -NoNewWindow -Wait
Add-Content -Path C:\IT\Complete.txt -value "Connection Optimizer Removed"

# Remove shortcuts
Remove-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\HP Documentation.lnk" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\HP\HP Documentation.lnk" -Force -ErrorAction SilentlyContinue
Add-Content -Path C:\IT\Complete.txt -value "HP Shortcuts removed"

# Remove directory
Remove-Item -Path "C:\Program Files\HP\Documentation" -Recurse -Force -ErrorAction SilentlyContinue
Add-Content -Path C:\IT\Complete.txt -value "HP Directory Shortcuts removed"

#Final Bloatware Check
Start-Process "appwiz.cpl"

#############################################################################################
#Final Confirm before restart

# Speak a text
$synthesizer.Speak('Read Me')

# Load necessary assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create a new form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Confirm Action'
$form.Size = New-Object System.Drawing.Size(300,250)
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
$checkBox.Text = "Confirm Windows Store Apps are Updated"
$checkBox.ForeColor = [System.Drawing.Color]::White  # Set text color to white for visibility

$checkBox2 = New-Object System.Windows.Forms.CheckBox
$checkBox2.Location = New-Object System.Drawing.Point(10,70)
$checkBox2.Size = New-Object System.Drawing.Size(280,40)
$checkBox2.Text = "Remove HP Wolf Security"
$checkBox2.ForeColor = [System.Drawing.Color]::White  # Set text color to white for visibility

$checkBox3 = New-Object System.Windows.Forms.CheckBox
$checkBox3.Location = New-Object System.Drawing.Point(10,90)
$checkBox3.Size = New-Object System.Drawing.Size(280,60)
$checkBox3.Text = "Clicking OK will restart computer"
$checkBox3.ForeColor = [System.Drawing.Color]::White  # Set text color to white for visibility

# Create an OK button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(110,180)
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
$form.Controls.Add($checkBox3)
$form.Controls.Add($label)
$form.Controls.Add($okButton)
$form.AcceptButton = $okButton


# Show the form as a dialog box
$result = $form.ShowDialog()

# Cleanup
$form.Dispose()

#############################################################################################

# Speak a text
$synthesizer.Speak('Initiating Restart')

#############################################################################################

Add-Content -Path C:\IT\Complete.txt -value "NUSPR_1 - Complete"
write-host ''
$formattedTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Add-Content -Path C:\IT\Complete.txt -value "Restarting: $formattedTimestamp"
write-host ''
Start-Sleep 2
Restart-Computer
