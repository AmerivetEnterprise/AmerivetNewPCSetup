#Version 5.0 
#08/01/2024

$host.UI.RawUI.WindowTitle = "New User Setup Post Restart 4"

Write-Host 'Script will continue in 30 seconds'
Start-Sleep 30

##############################################################################################################################################

#Enables RunOnce 
Write-Host "Changing RunOnce script for Post Restart 5 PS1." -foregroundcolor "magenta"
write-host ''
$RunOnceKey = "HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce"
set-itemproperty $RunOnceKey "NextRun" ('C:\Windows\System32\WindowsPowerShell\v1.0\Powershell.exe -executionPolicy Unrestricted -File ' + "C:\IT\New_User_Setup_Post_Restart_5.ps1")
write-host ''

##################################################################################################################################################################

$OS_Key = Get-WmiObject -query 'select * from SoftwareLicensingService' | Select OA3xOriginalProductKey
$OS_Key = $OS_Key.OA3xOriginalProductKey

slmgr.vbs /ipk $OS_Key

##################################################################################################################################################################
# Busylight Software install 

$input = Read-Host "Install Busylight Software (HQ Employee's only) [y/n]"

switch($input)

    {

y
{
    Start-Process msiexec.exe -wait -ArgumentList '/i C:\IT\Busylight4MS_Teams_Setup64.msi CMDLINE="MODIFY=FALSE REMOVE=FALSE SILENT=TRUE ALLUSERS=TRUE" /quiet'

    Write-Host "Busylight4MS_Teams Installed"
}

#n {exit}
n
{
write-warning "BusyLight Install skipped"
}

default
{
write-warning "BusyLight Install skipped"
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

##############################################################################################################################################

#Remove Desktop Shortcuts 
Remove-Item -Path "C:\Users\Public\Desktop\Busylight for MS Teams.lnk" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "C:\Users\Public\Desktop\Adobe Acrobat DC.lnk" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "C:\Users\Public\Desktop\TeamViewer.lnk" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "C:\Users\Public\Desktop\HP Support Assistant.lnk" -Force -ErrorAction SilentlyContinue

Write-Host "Shortcuts Removed"

#####################################################################################

#Delete Adobe Zip File
Remove-Item -Path "C:\IT\AmerivetAcrobat.zip" -recurse -Force -Confirm:$false

##############################################################################################################################################

Add-Content -Path C:\IT\Complete.txt -value "`NUSPR_4 - Complete"

#####################################################################################

