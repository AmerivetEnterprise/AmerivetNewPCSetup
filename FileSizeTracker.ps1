#Version 5.0 
#08/01/2024

$T = 3.22
$P = 1

do {

$file = 'C:\IT\AmerivetAcrobat.zip'

$1 = ((Get-Item $file).length/1GB)

$1 = [math]::Round($1, 2)

#
Write-Progress -Activity "Downloading Amerivet Acrobat" -status "$P% Complete:" -PercentComplete $P -CurrentOperation "Total Downloaded $("$1 GB") / 3.22GB"

#write-host  "$1 GB"

$R = $1 * (100 / $T) 

$P = [math]::Round($R, 2)

start-sleep 2

} until($1 -eq 3.22)

$1 = $null 

#} while ( $1 -ne 2.26 )