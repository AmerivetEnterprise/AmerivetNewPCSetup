#Version 5.0 
#08/01/2024

$host.UI.RawUI.WindowTitle = "New User Setup Cleanup"

Write-Host 'Script will continue in 10 seconds'
Start-Sleep 10

Remove-Item -Path "C:\IT\New_User_Setup_Post_Restart_1.ps1" -Force
Remove-Item -Path "C:\IT\New_User_Setup_Post_Restart_2.ps1" -Force 
Remove-Item -Path "C:\IT\New_User_Setup_Post_Restart_3.ps1" -Force
Remove-Item -Path "C:\IT\New_User_Setup_Post_Restart_4.ps1" -Force
Remove-Item -Path "C:\IT\New_User_Setup_Post_Restart_5.ps1" -Force
Remove-Item -Path "C:\IT\FalconCID.txt" -Force

Write-Host "Script Cleanup Complete" -foregroundcolor "Green"

Add-Content -Path C:\IT\Complete.txt -value "`Cleanup - Complete"

Cleanmgr

Read-Host "All Scripts Complete"
