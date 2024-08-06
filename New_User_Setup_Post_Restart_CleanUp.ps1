#Version 5.0 
#08/01/2024

$host.UI.RawUI.WindowTitle = "New User Setup Cleanup"

Write-Host 'Script will continue in 10 seconds'
Start-Sleep 10

Remove-Item -Path "C:\IT\New_User_Setup_Post_Restart_1.ps1" -Force
Write-Host "New_User_Setup_Post_Restart_1 Deleted" -ForegroundColor Green
Remove-Item -Path "C:\IT\New_User_Setup_Post_Restart_2.ps1" -Force 
Write-Host "New_User_Setup_Post_Restart_2 Deleted" -ForegroundColor Green
Remove-Item -Path "C:\IT\New_User_Setup_Post_Restart_3.ps1" -Force
Write-Host "New_User_Setup_Post_Restart_3 Deleted" -ForegroundColor Green
Remove-Item -Path "C:\IT\New_User_Setup_Post_Restart_4.ps1" -Force
Write-Host "New_User_Setup_Post_Restart_4 Deleted" -ForegroundColor Green
Remove-Item -Path "C:\IT\New_User_Setup_Post_Restart_5.ps1" -Force
Write-Host "New_User_Setup_Post_Restart_5 Deleted" -ForegroundColor Green
Remove-Item -Path "C:\IT\Info.csv" -Force
Write-Host "Info.csv Deleted" -ForegroundColor Green

#####################################################################################

#Remove Desktop Shortcuts 
Remove-Item -Path "C:\Users\Public\Desktop\Busylight for MS Teams.lnk" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "C:\Users\Public\Desktop\Adobe Acrobat DC.lnk" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "C:\Users\Public\Desktop\TeamViewer.lnk" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "C:\Users\Public\Desktop\HP Support Assistant.lnk" -Force -ErrorAction SilentlyContinue
Write-Host "Shortcuts Removed" -ForegroundColor Green

#####################################################################################

#Delete Adobe Zip File
Remove-Item -Path "C:\IT\AmerivetAcrobat.zip" -recurse -Force -Confirm:$false
Write-Host "Adobe Install Zip Deleted" -ForegroundColor Green

#####################################################################################

Write-Host "Script Cleanup Complete" -foregroundcolor "Green"

Add-Content -Path C:\IT\Complete.txt -value "`Cleanup - Complete"

Cleanmgr

Read-Host "All Scripts Complete"
