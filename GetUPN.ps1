Mkdir "C:\IT" -ErrorAction SilentlyContinue

#######

#Define the folder path
$folderPath = "C:\IT\UPN.txt"

# Check if the file exists
if (Test-Path -Path $folderPath)

{
    #DO NOTHING
} 

else 

{
#Make UPN.txt
New-Item -Path "C:\IT\UPN.txt" -ItemType File -Force

start-sleep 03

$UPN = whoami /upn

Add-Content -Path C:\IT\UPN.txt -value "$UPN"
}