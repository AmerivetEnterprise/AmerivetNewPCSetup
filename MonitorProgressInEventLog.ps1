# Define the log source and interval
$logName = "Application"
$eventSource = "NewPCSetup"
$refreshInterval = 10  # Refresh every 10 seconds

Write-Host "Monitoring logs from source '$eventSource' in '$logName' log. Press Ctrl+C to stop manually."

while ($true) {
    # Clear the console
    Clear-Host

    # Display a header
    Write-Host "Fetching logs from $logName for source $eventSource (Updated: $(Get-Date))" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------" -ForegroundColor Green

    # Get logs and display relevant information, but reverse the order
    $logs = Get-WinEvent -LogName $logName -FilterXPath "*[System[Provider[@Name='$eventSource']]]" |
        Select-Object TimeCreated, Id, LevelDisplayName, Message |
        Sort-Object TimeCreated  # This sorts by time in ascending order (oldest first)

    # Output logs
    $logs | Format-Table -AutoSize

    # Check if any message contains "Complete"
    if ($logs | Where-Object { $_.Message -like "*New PC Setup Script Complete*" }) {
        Write-Host "--------------------------------------------------------------" -ForegroundColor Yellow
        Write-Host "New PC Setup Script Complete message detected. Stopping monitoring." -ForegroundColor Yellow
        break
    }

    # Wait for the refresh interval
    Start-Sleep -Seconds $refreshInterval
}
